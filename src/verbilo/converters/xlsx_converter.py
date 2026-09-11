from collections import Counter
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.cell.rich_text import CellRichText, TextBlock
import openpyxl.packaging.manifest as _opxl_manifest
from typing import Any, Callable, Mapping
import logging
import os
import os.path
import shutil
import threading
import tempfile
import unicodedata
import re
import posixpath
from zipfile import ZipFile, ZIP_DEFLATED
from io import BytesIO
from ..utils import CancelledError
from ..semantic import (
    ProtectedText,
    TranslationContext,
    TranslationService,
    TranslationUnit as SemanticTranslationUnit,
)
from lxml import etree

logger = logging.getLogger(__name__)


def _patch_openpyxl_mimetypes() -> None:
    # openpyxl's Manifest._register_mimetypes() KeyErrors when the workbook's zip package contains a file extension it doesn't have a MIME type for
    known_extra_types = {
        ".webp": "image/webp",
        ".wmf": "image/wmf",
        ".mpo": "image/mpo",
        ".avif": "image/avif",
        ".heic": "image/heic",
        ".heif": "image/heif",
    }
    for ext, mime in known_extra_types.items():
        _opxl_manifest.mimetypes.add_type(mime, ext)

    def _safe_register_mimetypes(self, filenames):
        for fn in filenames:
            ext = os.path.splitext(fn)[-1]
            if not ext:
                continue
            try:
                mime = _opxl_manifest.mimetypes.types_map[True][ext]
            except KeyError:
                logger.warning(
                    "Unknown MIME type for embedded file extension '%s' in XLSX package; "
                    "defaulting to application/octet-stream", ext,
                )
                mime = "application/octet-stream"
            self.Default.append(_opxl_manifest.FileExtension(ext[1:], mime))

    _opxl_manifest.Manifest._register_mimetypes = _safe_register_mimetypes


_patch_openpyxl_mimetypes()

# Control-character pattern: matches C0/C1 control chars except tab, newline, carriage return,
# plus invisible Unicode format/zero-width characters that can silently corrupt translation.
_CONTROL_CHAR_RE = re.compile(
    r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f'
    r'\u00ad'          # soft hyphen
    r'\u200b-\u200d'   # zero-width space / non-joiner / joiner
    r'\u2060'          # word joiner
    r'\ufeff]'         # BOM / zero-width no-break space
)

# Separator used to group multiple cells into a single translation unit.
# Chosen to be extremely unlikely in real spreadsheet data.
_CELL_SEP_TOKEN = "\u27EASEP\u27EB"
_CELL_SEP = f"\n{_CELL_SEP_TOKEN}\n"
_CELL_SEP_SPLIT_RE = re.compile(rf"\s*{re.escape(_CELL_SEP_TOKEN)}\s*")

# Maximum characters per grouped row before falling back to per-cell.
_GROUP_MAX_CHARS = 4000

# ── Protected-term glossary ────────────────────────────────────────────────
# Uses a bracket pair distinct from the existing ⟨rN⟩/⟨sN⟩ (DOCX/PDF run tags) and ⟪SEP⟫ (row-grouping) markers so the three schemes never collide.
_GLOSSARY_OPEN = "\u27E6"   # ⟦
_GLOSSARY_CLOSE = "\u27E7"  # ⟧
_GLOSSARY_TOKEN_RE = re.compile(re.escape(_GLOSSARY_OPEN) + r"G\d+" + re.escape(_GLOSSARY_CLOSE))

# ISO 4217 currency codes commonly seen in pricing/trade documents.
_CURRENCY_CODES: frozenset[str] = frozenset({
    "USD", "EUR", "GBP", "JPY", "CNY", "RMB", "HKD", "AUD", "CAD", "CHF",
    "SGD", "NZD", "SEK", "NOK", "DKK", "INR", "KRW", "MXN", "BRL", "ZAR",
    "RUB", "TRY", "AED", "SAR", "THB", "MYR", "IDR", "PHP", "VND", "TWD",
    "PLN", "CZK", "HUF", "ILS", "EGP", "NGN", "PKR", "BDT", "CLP", "COP",
})

# Incoterms 2020 delivery terms.
_INCOTERMS: frozenset[str] = frozenset({
    "EXW", "FCA", "FAS", "FOB", "CFR", "CIF", "CPT", "CIP", "DAP", "DPU", "DDP",
})

# Common trade / shipping / packaging abbreviations found on catalog and packing-list spreadsheets.
_TRADE_ABBREVIATIONS: frozenset[str] = frozenset({
    "CBM", "CTN", "CTNS", "PCS", "SET", "SETS", "KGS", "MT", "GW", "NW",
    "QTY", "MOQ", "SKU", "PO", "ETA", "ETD", "FCL", "LCL", "TEU", "HS",
})

_DEFAULT_GLOSSARY_TERMS: frozenset[str] = _CURRENCY_CODES | _INCOTERMS | _TRADE_ABBREVIATIONS
_glossary_pattern_cache: dict[frozenset, re.Pattern] = {}


def _build_glossary_pattern(extra_terms: frozenset | None) -> re.Pattern | None:
    terms = _DEFAULT_GLOSSARY_TERMS | extra_terms if extra_terms else _DEFAULT_GLOSSARY_TERMS
    if not terms:
        return None
    cached = _glossary_pattern_cache.get(terms)
    if cached is not None:
        return cached
    # Longest-first so multi-word/longer codes aren't shadowed by shorter ones.
    ordered = sorted(terms, key=len, reverse=True)
    pattern = re.compile(r"\b(?:%s)\b" % "|".join(re.escape(t) for t in ordered))
    _glossary_pattern_cache[terms] = pattern
    return pattern


def _protect_glossary_terms(text: str, extra_terms: frozenset | None) -> tuple[str, dict[str, str]]:
    # Replace glossary terms in *text* with opaque placeholder tokens
    pattern = _build_glossary_pattern(extra_terms)
    if pattern is None:
        return text, {}

    token_map: dict[str, str] = {}
    counter = 0

    def _sub(m: re.Match) -> str:
        nonlocal counter
        token = f"{_GLOSSARY_OPEN}G{counter}{_GLOSSARY_CLOSE}"
        token_map[token] = m.group(0)
        counter += 1
        return token

    return pattern.sub(_sub, text), token_map


def _restore_glossary_terms(text: str, token_map: dict[str, str]) -> str:
    if not token_map:
        return text
    for token, original in token_map.items():
        text = text.replace(token, original)
    return text


def _glossary_tokens_match(text: str, token_map: dict[str, str]) -> bool:
    """Return True only when every protected glossary token appears exactly once.

    Checking only for leftover tokens after restoration misses dropped placeholders.
    Exact multiplicity also catches model-generated duplicates.
    """
    if not token_map:
        return True
    expected = Counter(token_map.keys())
    found = Counter(_GLOSSARY_TOKEN_RE.findall(str(text or "")))
    return found == expected

# Symbol-only pattern: matches strings entirely composed of punctuation,
# symbols, geometric shapes (including "►" U+25B6)
_SYMBOL_ONLY_RE = re.compile(
    r'^[\s'
    r'\u0021-\u002F\u003A-\u0040\u005B-\u0060\u007B-\u007E'
    r'\u00A0-\u00BF'
    r'\u2000-\u206F'
    r'\u2190-\u21FF'
    r'\u2300-\u23FF'
    r'\u2460-\u24FF'
    r'\u2500-\u257F'
    r'\u2580-\u259F'
    r'\u25A0-\u25FF'
    r'\u2600-\u26FF'
    r'\u2700-\u27BF'
    r'\u3000-\u303F'
    r'\uFE50-\uFE6F'
    r'\uFF00-\uFF0F\uFF1A-\uFF20\uFF3B-\uFF40\uFF5B-\uFF65'
    r'\uFF66-\uFFEF'
    r']+$'
)


def _is_symbol_char(c: str) -> bool:
    cp = ord(c)
    return (
        0x0021 <= cp <= 0x002F or 0x003A <= cp <= 0x0040 or
        0x005B <= cp <= 0x0060 or 0x007B <= cp <= 0x007E or
        0x00A0 <= cp <= 0x00BF or
        0x2000 <= cp <= 0x206F or 0x2190 <= cp <= 0x21FF or
        0x2300 <= cp <= 0x23FF or 0x2460 <= cp <= 0x24FF or
        0x2500 <= cp <= 0x257F or 0x2580 <= cp <= 0x259F or
        0x25A0 <= cp <= 0x25FF or 0x2600 <= cp <= 0x26FF or
        0x2700 <= cp <= 0x27BF or
        (0x3000 <= cp <= 0x303F) or
        0xFE50 <= cp <= 0xFE6F or
        (0xFF00 <= cp <= 0xFF0F) or
        0xFF1A <= cp <= 0xFF20 or
        0xFF3B <= cp <= 0xFF40 or 0xFF5B <= cp <= 0xFF65 or
        0xFF66 <= cp <= 0xFFEF or
        c in ' \t\n\r'
    )


def _is_symbol_only(text: str) -> bool:
    return bool(_SYMBOL_ONLY_RE.match(text))


def _strip_symbol_frame(text: str) -> tuple[str, str, str]:
    i = 0
    while i < len(text) and _is_symbol_char(text[i]):
        i += 1
    j = len(text)
    while j > i and _is_symbol_char(text[j - 1]):
        j -= 1
    return text[:i], text[i:j], text[j:]


def _strip_symbol_frame_multiline(text: str) -> tuple[str, list[tuple[str, str]]]:
    # Returns (core_text, frames) where core_text is the per-line-stripped text rejoined with "\n", and frames[i] = (prefix_i, suffix_i) is what was stripped from line i (either may be "").
    
    lines = text.split("\n")
    core_lines: list[str] = []
    frames: list[tuple[str, str]] = []
    for line in lines:
        pref, core, suff = _strip_symbol_frame(line)
        core_lines.append(core)
        frames.append((pref, suff))
    return "\n".join(core_lines), frames


def _reattach_symbol_frame_multiline(translated: str, frames: list[tuple[str, str]]) -> str:
    # Inverse of _strip_symbol_frame_multiline. Reattaches each line's own frame after translation, matching lines positionally.
    if not frames:
        return translated
    t_lines = translated.split("\n")
    if len(t_lines) == len(frames):
        return "\n".join(
            pref + line + suff
            for line, (pref, suff) in zip(t_lines, frames)
        )
    logger.debug(
        "XLSX symbol-frame line count mismatch: expected %d lines, got %d; "
        "recovering frame(s) without positional match",
        len(frames), len(t_lines),
    )

    unique_frames = set(frames)
    if len(unique_frames) == 1:
        pref, suff = frames[0]
        return "\n".join(pref + line + suff for line in t_lines)

    non_empty_frames = {f for f in unique_frames if f != ("", "")}
    if len(non_empty_frames) == 1:
        pref, suff = non_empty_frames.pop()
        return "\n".join(
            pref + line + suff if line.strip() else line
            for line in t_lines
        )

    prefixes = [p for p, _ in frames if p]
    suffixes = [s for _, s in frames if s]
    lead = "".join(dict.fromkeys(prefixes))
    trail = "".join(dict.fromkeys(suffixes))
    return lead + translated + trail


def _sanitize_text(text: str) -> str:
    # Normalize Unicode and strip problematic control characters
    text = unicodedata.normalize("NFC", text)
    text = _CONTROL_CHAR_RE.sub("", text)
    return text


def _split_grouped_row_translation(text: str) -> list[str]:
    if _CELL_SEP in text:
        return text.split(_CELL_SEP)
    if _CELL_SEP_TOKEN not in text:
        return [text]
    return _CELL_SEP_SPLIT_RE.split(text)


def _snapshot_formula_values(input_path: str) -> list[dict[str, object]]:
    # Snapshot the *cached results* stored in the original worksheet XML.
    # openpyxl exposes the formula expression as cell.value for data_type='f',
    # so using cell.value here would incorrectly write '=SUM(...)' into <v>.
    import xml.etree.ElementTree as _ET

    _SML_URI = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    _DOC_REL_URI = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    _PKG_REL_URI = 'http://schemas.openxmlformats.org/package/2006/relationships'
    ns = {'s': _SML_URI, 'r': _DOC_REL_URI, 'pr': _PKG_REL_URI}

    caches: list[dict[str, object]] = []
    with ZipFile(input_path, 'r') as zin:
        try:
            wb_root = _ET.fromstring(zin.read('xl/workbook.xml'))
            rel_root = _ET.fromstring(zin.read('xl/_rels/workbook.xml.rels'))
        except Exception:
            logger.warning('Could not snapshot XLSX formula caches', exc_info=True)
            return caches

        rel_targets = {
            rel.get('Id'): rel.get('Target')
            for rel in rel_root
            if rel.get('Id') and rel.get('Target')
        }

        for sheet in wb_root.findall('.//s:sheets/s:sheet', ns):
            rid = sheet.get(f'{{{_DOC_REL_URI}}}id')
            target = rel_targets.get(rid)
            cache: dict[str, object] = {}
            if target:
                target_path = target.lstrip('/')
                sheet_path = (
                    posixpath.normpath(target_path)
                    if target_path.startswith('xl/')
                    else posixpath.normpath(posixpath.join('xl', target_path))
                ).replace('\\', '/')
                try:
                    root = _ET.fromstring(zin.read(sheet_path))
                    sml = '{' + _SML_URI + '}'
                    for cell in root.findall(f'.//{sml}c'):
                        if cell.find(f'{sml}f') is None:
                            continue
                        v = cell.find(f'{sml}v')
                        ref = cell.get('r')
                        if ref and v is not None and v.text is not None:
                            cache[ref] = v.text
                except Exception:
                    logger.debug('Could not read cached formula values from %s', sheet_path, exc_info=True)
            caches.append(cache)

    return caches


def _patch_formula_values(zip_bytes: bytes, formula_cache: list[dict[str, object]]) -> bytes | None:
    # openpyxl strips the <v> element of formula cells during save.  This patches the serialised zip in memory, restoring cached values from the snapshot taken at load time.
    import xml.etree.ElementTree as _ET
    _SML_URI = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    _ET.register_namespace('', _SML_URI)
    _SML_NS = '{' + _SML_URI + '}'

    out_buf = BytesIO()
    changed = False
    with ZipFile(BytesIO(zip_bytes), 'r') as zin:
        with ZipFile(out_buf, 'w', ZIP_DEFLATED) as zout:
            for name in zin.namelist():
                data = zin.read(name)
                cache_entry: dict[str, object] | None = None
                if name.startswith('xl/worksheets/sheet') and name.endswith('.xml'):
                    stem = name[len('xl/worksheets/sheet'):-len('.xml')]
                    try:
                        idx = int(stem) - 1
                        if 0 <= idx < len(formula_cache):
                            cache_entry = formula_cache[idx]
                    except (ValueError, IndexError):
                        pass

                if cache_entry:
                    root = _ET.fromstring(data)
                    patched = False
                    for row in root.findall(f'.//{_SML_NS}row'):
                        for cell in row.findall(f'{_SML_NS}c'):
                            f_elem = cell.find(f'{_SML_NS}f')
                            v_elem = cell.find(f'{_SML_NS}v')
                            ref = cell.get('r')
                            if f_elem is not None and ref in cache_entry:
                                v_ok = v_elem is not None and v_elem.text is not None and v_elem.text.strip() != ''
                                if not v_ok:
                                    if v_elem is None:
                                        v = _ET.SubElement(cell, f'{_SML_NS}v')
                                    else:
                                        v = v_elem
                                    v.text = str(cache_entry[ref])
                                    patched = True
                    if patched:
                        data = _ET.tostring(root, xml_declaration=True, encoding='UTF-8')
                        changed = True
                zout.writestr(name, data)

    if not changed:
        return None
    return out_buf.getvalue()


def _save_preserving_orphan_rels(wb, output_path: str, input_path: str, _formula_cache: list[dict[str, object]] | None = None) -> None:
    import xml.etree.ElementTree as _ET
    _RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'

    buf = BytesIO()
    wb.save(buf)

    if _formula_cache:
        _saved_bytes = buf.getvalue()
        _patched = _patch_formula_values(_saved_bytes, _formula_cache)
        if _patched is not None:
            buf = BytesIO()
            buf.write(_patched)

    # Always compare against the original package.  Relying on wb.vba_archive
    # makes preservation dependent on openpyxl internals / keep_vba behavior and
    # can silently skip restoration for ordinary .xlsx files.
    # ── 1. Read both archives into memory ─────────────────────────────
    saved_bytes = buf.getvalue()
    saved_files: dict[str, bytes] = {}
    with ZipFile(BytesIO(saved_bytes), 'r') as z:
        for name in z.namelist():
            saved_files[name] = z.read(name)

    original_files: dict[str, bytes] = {}
    with ZipFile(input_path, 'r') as original_archive:
        for name in original_archive.namelist():
            original_files[name] = original_archive.read(name)

    # ── 2. Find orphan parts ──────────────────────────────────────────
    # Parts under xl/ that exist in the original but not in the saved
    # archive — openpyxl dropped them because it doesn't understand them.
    orphan_all: dict[str, bytes] = {
        name: data for name, data in original_files.items()
        if name.startswith('xl/') and name not in saved_files
    }

    if not orphan_all:
        with open(output_path, 'wb') as f:
            f.write(buf.getvalue())
        return

    orphan_media = {n: d for n, d in orphan_all.items() if n.startswith('xl/media/')}
    orphan_parts = {n: d for n, d in orphan_all.items() if not n.startswith('xl/media/')}

    # ── 3. Rebuild the zip with orphan parts injected ─────────────────
    # Write these separately to avoid duplicates (merge logic below).
    _WB_RELS = 'xl/_rels/workbook.xml.rels'
    _CT = '[Content_Types].xml'
    _deferred = {_WB_RELS, _CT}

    out_buf = BytesIO()
    with ZipFile(out_buf, 'w', ZIP_DEFLATED) as zout:
        for name, data in saved_files.items():
            if name not in _deferred:
                zout.writestr(name, data)
        for name, data in orphan_media.items():
            zout.writestr(name, data)
        for name, data in orphan_parts.items():
            zout.writestr(name, data)

        # ── 4. Merge workbook.xml.rels ────────────────────────────────
        saved_rels = saved_files.get(_WB_RELS, b'')
        orig_rels = original_files.get(_WB_RELS, b'')
        wb_rels = saved_rels  # fallback
        if saved_rels and orig_rels:
            saved_root = _ET.fromstring(saved_rels)
            saved_targets = {child.get('Target') for child in saved_root}
            orig_root = _ET.fromstring(orig_rels)
            changed = False
            for child in orig_root:
                target = child.get('Target')
                if not target:
                    continue
                norm_target = posixpath.normpath(target).replace('\\', '/')
                orphan_path = f'xl/{norm_target}' if not norm_target.startswith('xl/') else norm_target
                if orphan_path in orphan_all and target not in saved_targets:
                    elem = _ET.SubElement(saved_root, f'{{{_RELS_NS}}}Relationship')
                    elem.set('Id', child.get('Id'))
                    elem.set('Type', child.get('Type'))
                    elem.set('Target', child.get('Target'))
                    saved_targets.add(target)
                    changed = True
            if changed:
                wb_rels = _ET.tostring(saved_root, xml_declaration=True, encoding='UTF-8')
            else:
                wb_rels = saved_rels
        zout.writestr(_WB_RELS, wb_rels)

        # ── 5. Merge [Content_Types].xml ──────────────────────────────
        saved_ct = saved_files.get(_CT, b'')
        orig_ct = original_files.get(_CT, b'')
        ct_xml = saved_ct  # fallback
        if saved_ct and orig_ct:
            _CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
            saved_root = _ET.fromstring(saved_ct)
            orig_root = _ET.fromstring(orig_ct)

            saved_defaults = {
                child.get('Extension')
                for child in saved_root
                if child.tag.endswith('Default') and child.get('Extension')
            }
            saved_overrides = {
                child.get('PartName')
                for child in saved_root
                if child.tag.endswith('Override') and child.get('PartName')
            }
            changed = False

            for name in orphan_media:
                ext = os.path.splitext(name)[-1].lstrip('.')
                if ext and ext not in saved_defaults:
                    for child in orig_root:
                        if child.tag.endswith('Default') and child.get('Extension') == ext:
                            elem = _ET.SubElement(saved_root, f'{{{_CT_NS}}}Default')
                            elem.set('Extension', ext)
                            elem.set('ContentType', child.get('ContentType'))
                            saved_defaults.add(ext)
                            changed = True
                            break

            for name in orphan_parts:
                part_name = '/' + name.replace('\\', '/')
                if part_name not in saved_overrides:
                    for child in orig_root:
                        if child.tag.endswith('Override') and child.get('PartName') == part_name:
                            elem = _ET.SubElement(saved_root, f'{{{_CT_NS}}}Override')
                            elem.set('PartName', part_name)
                            elem.set('ContentType', child.get('ContentType'))
                            saved_overrides.add(part_name)
                            changed = True
                            break

            if changed:
                ct_xml = _ET.tostring(saved_root, xml_declaration=True, encoding='UTF-8')
            else:
                ct_xml = saved_ct
        zout.writestr(_CT, ct_xml)

    with open(output_path, 'wb') as f:
        f.write(out_buf.getvalue())



def _redistribute_text_parts(original_parts: list[str], translated: str) -> list[str]:
    """Split one logical translation back across the original rich-text runs.

    The concatenation of the returned parts is exactly ``translated``.  Empty
    source runs stay empty, and boundaries are nudged toward whitespace or
    punctuation so a font/style change is less likely to land inside a word.
    """
    if not original_parts:
        return []
    if len(original_parts) == 1:
        return [translated]
    if ''.join(original_parts) == translated:
        return list(original_parts)

    visible = [i for i, part in enumerate(original_parts) if part]
    if not visible:
        return ['' for _ in original_parts]
    if len(visible) == 1:
        out = ['' for _ in original_parts]
        out[visible[0]] = translated
        return out

    weights = [max(1, len(original_parts[i].strip()) or len(original_parts[i])) for i in visible]
    total_weight = sum(weights)
    target_len = len(translated)
    safe_chars = set(' \t\r\n,.;:!?，。；：！？、/\\()[]{}<>-–—|')
    boundaries: list[int] = []
    cumulative = 0
    previous = 0

    for weight in weights[:-1]:
        cumulative += weight
        ideal = round(target_len * cumulative / total_weight)
        ideal = max(previous, min(ideal, target_len))
        best = ideal
        best_distance = 10**9
        radius = min(14, max(4, target_len // 20))
        lo = max(previous, ideal - radius)
        hi = min(target_len, ideal + radius)
        for pos in range(lo, hi + 1):
            if pos <= previous:
                continue
            left_safe = pos > 0 and translated[pos - 1] in safe_chars
            right_safe = pos < target_len and translated[pos] in safe_chars
            if left_safe or right_safe:
                distance = abs(pos - ideal)
                if distance < best_distance:
                    best = pos
                    best_distance = distance
        boundaries.append(best)
        previous = best

    starts = [0] + boundaries
    ends = boundaries + [target_len]
    chunks = [translated[a:b] for a, b in zip(starts, ends)]
    out = ['' for _ in original_parts]
    for idx, chunk in zip(visible, chunks):
        out[idx] = chunk
    return out


def _translate_many_with_fallback(
    translator: Any,
    texts: list[str],
    target_lang: str,
    cancel_event: threading.Event | None,
) -> tuple[list[str | None], int]:
    """Compatibility wrapper around the shared format-neutral translation service."""
    if not texts:
        return [], 0
    service = TranslationService(translator)
    result = service.translate_units(
        [
            SemanticTranslationUnit(
                text=text,
                role="table-cell",
                mode="natural",
                metadata={"content_hint": "table-cell", "strategy": "semantic"},
            )
            for text in texts
        ],
        target_lang,
        cancel_event=cancel_event,
    )
    return result.texts, len(result.failed_indices)


def _xlsx_set_text_node(node: etree._Element, text: str) -> None:
    node.text = text
    xml_space = '{http://www.w3.org/XML/1998/namespace}space'
    if text[:1].isspace() or text[-1:].isspace():
        node.set(xml_space, 'preserve')


def _xlsx_collect_text_nodes(container: etree._Element, text_tag: str) -> list[etree._Element]:
    """Collect visible text nodes while excluding phonetic/ruby helpers."""
    nodes: list[etree._Element] = []
    for node in container.iter(text_tag):
        parent = node.getparent()
        skip = False
        while parent is not None and parent is not container:
            if etree.QName(parent).localname in {'rPh', 'phoneticPr'}:
                skip = True
                break
            parent = parent.getparent()
        if not skip:
            nodes.append(node)
    return nodes


def _xlsx_prepare_unit(
    nodes: list[etree._Element],
    extra_glossary_terms: frozenset | None,
) -> dict[str, Any] | None:
    if not nodes:
        return None
    original_parts = [node.text or '' for node in nodes]
    original_text = ''.join(original_parts)
    if not original_text or not original_text.strip():
        return None

    sanitized = _sanitize_text(original_text)
    if not sanitized.strip() or _is_symbol_only(sanitized):
        return None
    core_text, frames = _strip_symbol_frame_multiline(sanitized)
    if not core_text.strip():
        return None
    protected, glossary_map = _protect_glossary_terms(core_text, extra_glossary_terms)
    return {
        'nodes': nodes,
        'original_parts': original_parts,
        'original_text': original_text,
        'raw_text': core_text,
        'source_text': protected,
        'frames': frames,
        'glossary_map': glossary_map,
    }


def _xlsx_parse_xml(data: bytes) -> etree._Element:
    parser = etree.XMLParser(resolve_entities=False, remove_blank_text=False, recover=False)
    return etree.fromstring(data, parser=parser)


def _xlsx_serialize_xml(root: etree._Element, original: bytes) -> bytes:
    encoding = 'UTF-8'
    if original.startswith(b'<?xml'):
        try:
            decl = original[: original.index(b'?>') + 2].decode('ascii', errors='ignore')
            m = re.search(r'encoding=["\\\']([^"\\\']+)', decl, re.I)
            if m:
                encoding = m.group(1)
        except Exception:
            pass
    standalone = b'standalone="yes"' in original[:200] or b"standalone='yes'" in original[:200]
    return etree.tostring(
        root,
        xml_declaration=True,
        encoding=encoding,
        standalone=True if standalone else None,
    )


def _xlsx_rebuild_package(
    input_path: str,
    output_path: str,
    patches: dict[str, bytes],
) -> None:
    """Rebuild the OOXML zip, changing only explicitly patched XML parts."""
    if not patches:
        if os.path.abspath(input_path) != os.path.abspath(output_path):
            shutil.copy2(input_path, output_path)
        return
    out_dir = os.path.dirname(os.path.abspath(output_path)) or '.'
    os.makedirs(out_dir, exist_ok=True)
    fd, tmp_path = tempfile.mkstemp(prefix='.verbilo-xlsx-', suffix='.tmp', dir=out_dir)
    os.close(fd)
    try:
        with ZipFile(input_path, 'r') as zin, ZipFile(tmp_path, 'w') as zout:
            zout.comment = zin.comment
            for info in zin.infolist():
                data = patches.get(info.filename)
                if data is None:
                    data = zin.read(info.filename)
                # Passing the original ZipInfo preserves timestamps, attributes,
                # compression method, comments and extra metadata.
                zout.writestr(info, data)
        os.replace(tmp_path, output_path)
    except Exception:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        raise


_CELL_REF_RE = re.compile(r"^([A-Za-z]+)(\d+)$")


def _xlsx_column_number(cell_ref: str) -> int:
    match = _CELL_REF_RE.match(str(cell_ref or ""))
    if match is None:
        return 0
    value = 0
    for char in match.group(1).upper():
        value = value * 26 + (ord(char) - ord("A") + 1)
    return value


def _xlsx_sheet_names_by_part(zin: ZipFile, names: set[str]) -> dict[str, str]:
    workbook_name = "xl/workbook.xml"
    rels_name = "xl/_rels/workbook.xml.rels"
    if workbook_name not in names or rels_name not in names:
        return {}

    try:
        workbook = _xlsx_parse_xml(zin.read(workbook_name))
        rels = _xlsx_parse_xml(zin.read(rels_name))
    except Exception:
        logger.debug("Could not parse XLSX workbook relationships for semantic context", exc_info=True)
        return {}

    rel_by_id = {
        rel.get("Id"): rel.get("Target")
        for rel in rels
        if rel.get("Id") and rel.get("Target")
    }
    rel_attr = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    mapping: dict[str, str] = {}
    for sheet in workbook.iter():
        if etree.QName(sheet).localname != "sheet":
            continue
        rid = sheet.get(rel_attr)
        target = rel_by_id.get(rid)
        name = sheet.get("name")
        if not target or not name:
            continue
        normalized = target.lstrip("/")
        if not normalized.startswith("xl/"):
            normalized = posixpath.normpath(posixpath.join("xl", normalized))
        normalized = normalized.replace("\\", "/")
        mapping[normalized] = name
    return mapping


def _xlsx_collect_shared_string_contexts(
    zin: ZipFile,
    names: set[str],
    shared_texts: list[str],
) -> tuple[dict[int, TranslationContext], dict[str, str]]:
    """Map shared-string ids to lightweight sheet/header context.

    Shared strings can be referenced from many cells. We only retain context that
    is common or unambiguous enough to help translation: sheet, cell coordinate,
    and nearest string-valued row/column headers. Context is advisory only and is
    never written back into the workbook.
    """
    sheet_names = _xlsx_sheet_names_by_part(zin, names)
    usages: dict[int, list[dict[str, str]]] = {}
    sml = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

    for sheet_part, sheet_name in sheet_names.items():
        if sheet_part not in names:
            continue
        try:
            root = _xlsx_parse_xml(zin.read(sheet_part))
        except Exception:
            logger.debug("Could not parse %s for XLSX semantic context", sheet_part, exc_info=True)
            continue

        cells: list[tuple[str, int, int, int]] = []
        for cell in root.iter(f"{sml}c"):
            if cell.get("t") != "s":
                continue
            ref = cell.get("r") or ""
            match = _CELL_REF_RE.match(ref)
            value = cell.find(f"{sml}v")
            if match is None or value is None or value.text is None:
                continue
            try:
                shared_idx = int(value.text)
            except ValueError:
                continue
            if not (0 <= shared_idx < len(shared_texts)):
                continue
            row = int(match.group(2))
            col = _xlsx_column_number(ref)
            cells.append((ref, row, col, shared_idx))

        by_row: dict[int, list[tuple[str, int, int, int]]] = {}
        by_col: dict[int, list[tuple[str, int, int, int]]] = {}
        for entry in cells:
            by_row.setdefault(entry[1], []).append(entry)
            by_col.setdefault(entry[2], []).append(entry)
        for entries in by_row.values():
            entries.sort(key=lambda item: item[2])
        for entries in by_col.values():
            entries.sort(key=lambda item: item[1])

        for ref, row, col, shared_idx in cells:
            metadata: dict[str, str] = {"sheet": sheet_name, "cell": ref}
            left = [entry for entry in by_row[row] if entry[2] < col]
            above = [entry for entry in by_col[col] if entry[1] < row]
            if left:
                # Prefer the left-most string in the row as a stable row label
                # rather than the immediately previous data cell.
                header_idx = left[0][3]
                header = str(shared_texts[header_idx]).strip()
                if header:
                    metadata["row_header"] = header[:160]
            if above:
                # Prefer the top-most string in the column as a stable column
                # heading rather than the previous row's data value.
                header_idx = above[0][3]
                header = str(shared_texts[header_idx]).strip()
                if header:
                    metadata["column_header"] = header[:160]
            usages.setdefault(shared_idx, []).append(metadata)

    contexts: dict[int, TranslationContext] = {}
    for shared_idx, entries in usages.items():
        sheets = {entry.get("sheet", "") for entry in entries if entry.get("sheet")}
        section = next(iter(sheets)) if len(sheets) == 1 else None
        metadata: dict[str, str] = {}
        if len(entries) == 1:
            metadata.update(entries[0])
        else:
            if section:
                metadata["sheet"] = section
            cells = [f"{entry.get('sheet', '')}!{entry.get('cell', '')}".strip("!") for entry in entries[:6]]
            if cells:
                metadata["cells"] = "; ".join(cells)
            for key in ("row_header", "column_header"):
                values = {entry.get(key, "") for entry in entries if entry.get(key)}
                if len(values) == 1:
                    metadata[key] = next(iter(values))
        contexts[shared_idx] = TranslationContext.from_values(
            section=section,
            metadata=metadata,
        )
    return contexts, sheet_names


def _xlsx_inline_context(
    inline: etree._Element,
    *,
    sheet_name: str | None,
    part_name: str,
) -> TranslationContext:
    node = inline
    cell_ref = None
    while node is not None:
        if etree.QName(node).localname == "c":
            cell_ref = node.get("r")
            break
        node = node.getparent()
    metadata = {"xml_part": part_name}
    if sheet_name:
        metadata["sheet"] = sheet_name
    if cell_ref:
        metadata["cell"] = cell_ref
    return TranslationContext.from_values(section=sheet_name, metadata=metadata)


def translate_xlsx(
    input_path: str,
    output_path: str,
    translator: Any,
    target_lang: str,
    *,
    cancel_event: threading.Event | None = None,
    source_lang: str = "auto",
    progress_callback: 'Callable[[int, int], None] | None' = None,
    protected_terms: list[str] | None = None,
    group_rows: bool = False,
    strict_errors: bool = False,
    terminology: Mapping[str, str] | None = None,
):
    """Translate spreadsheet text by patching OOXML text parts directly.

    Phase 2 deliberately does *not* save the workbook through openpyxl.  Re-saving
    a complex 43 MB catalog can rewrite/drop unsupported drawing/media/package
    features even when only a few cell strings changed.  Instead we modify only
    ``sharedStrings.xml`` (plus true inline strings and DrawingML shape text when
    present) and copy every other zip member untouched at the XML-content level.

    Rich shared strings are one logical translation unit; their translated text is
    redistributed across the original ``<r>`` runs so fonts/colours/bold styling
    remain attached.  Failed units retain their original text.  ``strict_errors``
    restores fail-the-whole-file behaviour for callers that explicitly want it.
    """
    if group_rows:
        logger.info(
            'XLSX phase-2 OOXML mode ignores group_rows: shared/inline strings are translated as logical cells'
        )

    extra_glossary_terms = (
        frozenset(t.strip() for t in protected_terms if t and t.strip())
        if protected_terms else None
    )
    SML_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    SML_T = f'{{{SML_NS}}}t'
    SML_SI = f'{{{SML_NS}}}si'
    SML_IS = f'{{{SML_NS}}}is'
    A_T = f'{{{A_NS}}}t'
    A_P = f'{{{A_NS}}}p'

    roots: dict[str, tuple[etree._Element, bytes]] = {}
    units: list[dict[str, Any]] = []

    with ZipFile(input_path, 'r') as zin:
        names = set(zin.namelist())

        # Shared strings cover the normal cell text path, including rich text.
        shared_name = 'xl/sharedStrings.xml'
        shared_contexts: dict[int, TranslationContext] = {}
        sheet_names: dict[str, str] = _xlsx_sheet_names_by_part(zin, names)
        if shared_name in names:
            raw = zin.read(shared_name)
            root = _xlsx_parse_xml(raw)
            roots[shared_name] = (root, raw)
            shared_entries = list(root.iter(SML_SI))
            shared_texts = [
                ''.join(node.text or '' for node in _xlsx_collect_text_nodes(si, SML_T))
                for si in shared_entries
            ]
            shared_contexts, sheet_names = _xlsx_collect_shared_string_contexts(
                zin, names, shared_texts
            )
            for shared_index, si in enumerate(shared_entries):
                unit = _xlsx_prepare_unit(
                    _xlsx_collect_text_nodes(si, SML_T), extra_glossary_terms,
                )
                if unit is not None:
                    unit['part_name'] = shared_name
                    unit['role'] = 'table-cell'
                    unit['mode'] = 'natural'
                    unit['context'] = shared_contexts.get(shared_index, TranslationContext())
                    units.append(unit)

        # Some producers use inlineStr instead of sharedStrings.  Parse a sheet
        # only when that marker is present, avoiding needless sheet reserialization.
        for name in sorted(n for n in names if n.startswith('xl/worksheets/') and n.endswith('.xml')):
            raw = zin.read(name)
            if b'inlineStr' not in raw:
                continue
            root = _xlsx_parse_xml(raw)
            found_unit = False
            for inline in root.iter(SML_IS):
                unit = _xlsx_prepare_unit(
                    _xlsx_collect_text_nodes(inline, SML_T), extra_glossary_terms,
                )
                if unit is not None:
                    unit['part_name'] = name
                    unit['role'] = 'table-cell'
                    unit['mode'] = 'natural'
                    unit['context'] = _xlsx_inline_context(
                        inline, sheet_name=sheet_names.get(name), part_name=name
                    )
                    units.append(unit)
                    found_unit = True
            if found_unit:
                roots[name] = (root, raw)

        # Text boxes/shapes in worksheet drawings use DrawingML <a:p>/<a:t>.
        # Pictures and anchors contain no <a:t>, so they remain byte-identical.
        drawing_parts = sorted(
            n for n in names
            if (n.startswith('xl/drawings/') or n.startswith('xl/diagrams/'))
            and n.endswith('.xml')
        )
        for name in drawing_parts:
            raw = zin.read(name)
            if b'<a:t' not in raw and b':t>' not in raw:
                continue
            try:
                root = _xlsx_parse_xml(raw)
            except Exception:
                logger.debug('Skipping unparseable drawing text part %s', name, exc_info=True)
                continue
            found_unit = False
            for paragraph in root.iter(A_P):
                text_nodes = list(paragraph.iter(A_T))
                unit = _xlsx_prepare_unit(text_nodes, extra_glossary_terms)
                if unit is not None:
                    unit['part_name'] = name
                    unit['role'] = 'shape'
                    unit['mode'] = 'concise'
                    unit['context'] = TranslationContext.from_values(
                        metadata={"xml_part": name, "part_role": "shape"}
                    )
                    units.append(unit)
                    found_unit = True
            if found_unit:
                roots[name] = (root, raw)

    logger.info("XLSX '%s': collected %d logical OOXML text units", input_path, len(units))
    if not units:
        _xlsx_rebuild_package(input_path, output_path, {})
        return

    translation_service = TranslationService(translator, terminology=terminology)
    semantic_units: list[SemanticTranslationUnit] = []
    for unit in units:
        masked_text = str(unit['source_text'])
        raw_text = str(unit['raw_text'])
        glossary_map = unit['glossary_map']
        protected = (
            ProtectedText(masked_text, glossary_map)
            if glossary_map
            else None
        )
        semantic_units.append(
            SemanticTranslationUnit(
                text=raw_text,
                source_lang=source_lang,
                role=str(unit.get('role', 'table-cell') or 'table-cell'),
                mode=str(unit.get('mode', 'natural') or 'natural'),
                context=(
                    unit.get('context')
                    if isinstance(unit.get('context'), TranslationContext)
                    else TranslationContext()
                ),
                protected=protected,
                allow_unprotected_fallback=bool(protected),
                metadata={"content_hint": "table-cell", "strategy": "semantic"},
            )
        )

    batch = translation_service.translate_units(
        semantic_units, target_lang, cancel_event=cancel_event,
    )
    translated = batch.texts
    errors = len(batch.failed_indices)

    changed_parts: set[str] = set()
    for index, (unit, tr_text) in enumerate(zip(units, translated)):
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError('Translation cancelled')

        raw_text = str(unit['raw_text'])
        if tr_text is None:
            # Per-item fallback already logged/counts this failure.
            continue

        final_text = _reattach_symbol_frame_multiline(tr_text, unit['frames'])
        original_text = str(unit['original_text'])
        if final_text == original_text:
            continue

        chunks = _redistribute_text_parts(unit['original_parts'], final_text)
        for node, chunk in zip(unit['nodes'], chunks):
            _xlsx_set_text_node(node, chunk)
        changed_parts.add(str(unit['part_name']))

        if progress_callback is not None and (index + 1 == len(units) or (index + 1) % 10 == 0):
            progress_callback(index + 1, len(units))

    if progress_callback is not None:
        progress_callback(len(units), len(units))

    patches: dict[str, bytes] = {}
    for name in changed_parts:
        root, original = roots[name]
        patches[name] = _xlsx_serialize_xml(root, original)

    if cancel_event is not None and cancel_event.is_set():
        raise CancelledError('Translation cancelled before saving XLSX')
    _xlsx_rebuild_package(input_path, output_path, patches)

    if errors:
        message = (
            f'Translation completed with {errors} failed cells/text units; '
            'their original text was preserved'
        )
        if strict_errors:
            raise RuntimeError(message)
        logger.warning(message)
