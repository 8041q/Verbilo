from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.cell.rich_text import CellRichText, TextBlock
import openpyxl.packaging.manifest as _opxl_manifest
from typing import Any, Callable
import logging
import os
import os.path
import threading
import unicodedata
import re
import posixpath
from zipfile import ZipFile, ZIP_DEFLATED
from io import BytesIO
from ..utils import CancelledError

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

    original_archive = getattr(wb, 'vba_archive', None)
    if original_archive is None:
        with open(output_path, 'wb') as f:
            f.write(buf.getvalue())
        return

    # ── 1. Read both archives into memory ─────────────────────────────
    saved_bytes = buf.getvalue()
    saved_files: dict[str, bytes] = {}
    with ZipFile(BytesIO(saved_bytes), 'r') as z:
        for name in z.namelist():
            saved_files[name] = z.read(name)

    original_files: dict[str, bytes] = {}
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


def translate_xlsx(input_path: str, output_path: str, translator: Any, target_lang: str, *, cancel_event: threading.Event | None = None, source_lang: str = "auto", progress_callback: 'Callable[[int, int], None] | None' = None):
    # batch-translate XLSX with row-level contextual grouping.
    # When source_lang=="auto" row grouping is skipped so each cell is its own translation
    # unit, letting the API auto-detect the language per cell.
    wb = load_workbook(filename=input_path, rich_text=True, keep_vba=True)

    _formula_cache: list[dict[str, object]] = []
    for _sheet in wb.worksheets:
        _cache: dict[str, object] = {}
        for _row in _sheet.iter_rows():
            for _cell in _row:
                if _cell.data_type == 'f' and _cell.value is not None:
                    _cache[_cell.coordinate] = _cell.value
        _formula_cache.append(_cache)

    # --- collect every string cell that is writable, grouped by row ---
    # Each row group: list of (cell, core_text, symbol_prefix, symbol_suffix)
    # rich_segs: list of (cell, original_CellRichText, segment_index, sanitized_text)
    RowGroup = list[tuple[Any, str, list[tuple[str, str]]]]
    row_groups: list[RowGroup] = []
    current_row: RowGroup = []
    current_row_key: tuple | None = None  # (sheet_title, row_number)
    rich_segs: list[tuple[Any, CellRichText, int, str, list[tuple[str, str]]]] = []

    for sheet in wb.worksheets:
        for row in sheet.iter_rows(values_only=False):
            row_cells: RowGroup = []
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                val = cell.value
                if isinstance(val, str) and not val.startswith("="):
                    sanitized = _sanitize_text(val)
                    if not sanitized.strip():
                        continue
                    # Skip cells that are entirely symbols (no translatable content)
                    if _is_symbol_only(sanitized):
                        continue
                    sym_core, frames = _strip_symbol_frame_multiline(sanitized)
                    if not sym_core.strip():
                        continue
                    row_cells.append((cell, sym_core, frames))
                elif isinstance(val, CellRichText):
                    for seg_idx, seg in enumerate(val):
                        if isinstance(seg, TextBlock):
                            text = seg.text
                        elif isinstance(seg, str):
                            text = seg
                        else:
                            continue
                        if not (text and text.strip()):
                            continue
                        sanitized = _sanitize_text(text)
                        if not sanitized.strip():
                            continue
                        if _is_symbol_only(sanitized):
                            continue
                        sym_core, frames = _strip_symbol_frame_multiline(sanitized)
                        if not sym_core.strip():
                            continue
                        rich_segs.append((cell, val, seg_idx, sym_core, frames))
            if row_cells:
                row_groups.append(row_cells)

    total_cells = sum(len(rg) for rg in row_groups)
    logger.info(
        "XLSX '%s': collected %d plain-text cells in %d rows, %d rich-text segments",
        input_path, total_cells, len(row_groups), len(rich_segs),
    )

    if not row_groups and not rich_segs:
        logger.warning("No translatable text found in XLSX file '%s'", input_path)
        _save_preserving_orphan_rels(wb, output_path, input_path, _formula_cache)
        return

    # --- build translation units: group rows or send individually ---
    # A "unit" is either a joined row (multiple cells with separator) or a single cell.
    # After translation we split on separator to recover per-cell results.
    units: list[str] = []
    # Map: unit_index -> list of (cell, original_text) to write back
    unit_cells: list[list[tuple[Any, str]]] = []

    # In auto-detect mode do NOT group cells — every cell is its own unit so the
    # translation API sees a single-language segment and can auto-detect correctly.
    group_rows = source_lang != "auto"

    for rg in row_groups:
        row_texts = [t for _, t, _ in rg]
        total_chars = sum(len(t) for t in row_texts)

        if group_rows and len(rg) > 1 and total_chars <= _GROUP_MAX_CHARS:
            # Group the row into a single unit with separators
            units.append(_CELL_SEP.join(row_texts))
            unit_cells.append(rg)
        else:
            for cell, core, frames in rg:
                units.append(core)
                unit_cells.append([(cell, core, frames)])

    # --- batch-translate (plain cells + rich-text segments in one call) ---
    plain_count = len(units)
    all_units = units + [t for _, _, _, t, _ in rich_segs]
    total_units = len(all_units)
    try:
        translated_all = translator.translate_batch(all_units, target_lang, cancel_event=cancel_event)
    except CancelledError:
        raise
    except Exception:
        logger.exception("Batch translation failed for XLSX; falling back to per-item")
        translated_all = []
        for t in all_units:
            try:
                r = translator.translate_text(t, target_lang)
                translated_all.append(r if r is not None else t)
            except Exception:
                logger.exception("Per-item fallback also failed")
                translated_all.append(t)

    if progress_callback is not None:
        progress_callback(total_units, total_units)

    plain_translated = translated_all[:plain_count]
    rich_translated = translated_all[plain_count:]

    # --- write results back (plain cells) ---
    errors = 0
    for cells_in_unit, tr_text in zip(unit_cells, plain_translated):
        if tr_text is None:
            # Translation returned None — keep originals
            errors += 1
            continue

        if len(cells_in_unit) == 1:
            cell, orig, frames = cells_in_unit[0]
            cell.value = _reattach_symbol_frame_multiline(tr_text, frames)
        else:
            # Grouped row — split on separator
            parts = _split_grouped_row_translation(tr_text)
            if len(parts) == len(cells_in_unit):
                for (cell, orig, frames), part in zip(cells_in_unit, parts):
                    translated_part = part.strip() if part else orig
                    if translated_part == orig:
                        try:
                            r = translator.translate_text(orig, target_lang)
                            translated_part = r if r is not None else orig
                        except Exception:
                            logger.exception("Per-cell retry failed for unchanged segment")
                    cell.value = _reattach_symbol_frame_multiline(translated_part, frames)
            else:
                # Separator was consumed/mangled by the model — fall back to
                # per-cell translation for this row
                logger.debug(
                    "XLSX row separator mismatch: expected %d, got %d; per-cell fallback",
                    len(cells_in_unit), len(parts),
                )
                for cell, orig, frames in cells_in_unit:
                    try:
                        r = translator.translate_text(orig, target_lang)
                        cell.value = _reattach_symbol_frame_multiline(r if r is not None else orig, frames)
                    except Exception:
                        logger.exception("Per-cell fallback failed")
                        cell.value = _reattach_symbol_frame_multiline(orig, frames)
                        errors += 1

    # --- write results back (rich-text cells) ---
    # Group translated segments by cell so each cell is written exactly once.
    # cell_rich_updates: id(cell) -> (cell, original_CellRichText, {seg_idx: translated_text})
    cell_rich_updates: dict[int, tuple[Any, CellRichText, dict[int, str]]] = {}
    for (cell, rt, seg_idx, _, frames), tr_text in zip(rich_segs, rich_translated):
        cid = id(cell)
        if cid not in cell_rich_updates:
            cell_rich_updates[cid] = (cell, rt, {})
        if tr_text is not None:
            cell_rich_updates[cid][2][seg_idx] = _reattach_symbol_frame_multiline(tr_text, frames)

    for cid, (cell, rt, updates) in cell_rich_updates.items():
        if not updates:
            # Every segment translation failed for this cell
            errors += 1
            continue
        new_segs: list[TextBlock | str] = []
        for i, seg in enumerate(rt):
            if i in updates:
                if isinstance(seg, TextBlock):
                    new_segs.append(TextBlock(seg.font, updates[i]))
                else:
                    new_segs.append(updates[i])
            else:
                new_segs.append(seg)
        cell.value = CellRichText(*new_segs)

    # Check for cancellation before saving
    if cancel_event is not None and cancel_event.is_set():
        raise CancelledError("Translation cancelled before saving XLSX")

    _save_preserving_orphan_rels(wb, output_path, input_path, _formula_cache)
    if errors:
        raise RuntimeError(f"Translation completed with {errors} failed cells")
