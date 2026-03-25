from __future__ import annotations

import copy
import logging
import shutil
import threading
import zipfile
import io
from dataclasses import dataclass, field
from typing import Any, Callable

from lxml import etree
from docx.oxml.ns import qn

from ..utils import CancelledError

logger = logging.getLogger(__name__)


# XML namespace constants

_NS_W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_NS_A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_VML = "urn:schemas-microsoft-com:vml"
_DIAGRAM_DATA_CT = (
    "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"
)

_WT_TAG      = qn('w:t')
_WTAB_TAG    = qn('w:tab')
_WVAL_ATTR   = qn('w:val')

# Tags that should be treated as opaque containers — text inside these is
# handled by the SDT collector or skipped intentionally.
_SDT_TAG     = qn('w:sdt')
_SDT_PR_TAG  = qn('w:sdtPr')
_SDT_CONTENT = qn('w:sdtContent')

# Tracked-change insertion wrapper — we translate text inside <w:ins> normally.
# Deleted text (<w:del> / <w:delText>) is intentionally skipped.
_WINS_TAG    = qn('w:ins')
_WDEL_TAG    = qn('w:del')

# XML parts inside the zip that carry translatable body text
_BODY_PARTS = [
    "word/document.xml",
]
# These parts are discovered dynamically from relationships, but we also check
# well-known paths as a fallback.
_FOOTNOTE_PARTS    = ["word/footnotes.xml"]
_ENDNOTE_PARTS     = ["word/endnotes.xml"]
_HEADER_GLOB       = "word/header"
_FOOTER_GLOB       = "word/footer"


# Shared translation-unit container

@dataclass
class _TranslationUnit:
    # Uniform container for a text segment to be translated
    source_text: str
    is_heading: bool = False
    write_back: Callable[[str], None] = field(default=lambda t: None, repr=False)


# Grouping + batch translate

_PARA_SEP = "\n\u27EASEP\u27EB\n"
_GROUP_MAX_UNITS = 30
_GROUP_MAX_CHARS = 4000


def _group_units(
    units: list[_TranslationUnit],
    *,
    auto_detect: bool,
) -> list[list[int]]:
    if auto_detect:
        return [[i] for i in range(len(units))]

    groups: list[list[int]] = []
    current: list[int] = []
    current_chars = 0

    for i, unit in enumerate(units):
        text_len = len(unit.source_text)

        if unit.is_heading:
            if current:
                groups.append(current)
                current = []
                current_chars = 0
            groups.append([i])
            continue

        if current and (
            len(current) >= _GROUP_MAX_UNITS
            or current_chars + text_len > _GROUP_MAX_CHARS
        ):
            groups.append(current)
            current = []
            current_chars = 0

        current.append(i)
        current_chars += text_len

    if current:
        groups.append(current)

    return groups


def _translate_and_writeback(
    units: list[_TranslationUnit],
    groups: list[list[int]],
    translator: Any,
    target_lang: str,
    cancel_event: threading.Event | None,
) -> int:
    #  Translate each group and call write_back on each unit. Returns error count. Uses ``translate_batch`` when available, falling back to ``translate_text`` for single-item groups
    
    has_batch = callable(getattr(translator, 'translate_batch', None))
    has_text  = callable(getattr(translator, 'translate_text', None))

    if not has_batch and not has_text:
        raise AttributeError(
            f"{type(translator).__name__} exposes neither translate_batch nor "
            "translate_text — cannot translate."
        )

    errors = 0
    for group in groups:
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError("Translation cancelled")

        texts = [units[i].source_text for i in group]

        # ── Batch path (preferred) ────────────────────────────────────────────
        if has_batch and len(group) > 1:
            try:
                translated_parts = translator.translate_batch(
                    texts, target_lang, cancel_event=cancel_event
                )
                if len(translated_parts) != len(texts):
                    raise ValueError("translate_batch returned wrong number of results")
                for idx, i in enumerate(group):
                    units[i].write_back(translated_parts[idx])
                continue
            except CancelledError:
                raise
            except Exception:
                logger.warning(
                    "translate_batch failed for group %s; falling back to translate_text",
                    group, exc_info=True,
                )
                # fall through to per-item path below

        # ── Per-item path (single items, or batch fallback) ───────────────────
        translate_fn = (
            translator.translate_text if has_text
            else lambda t, tl: translator.translate_batch([t], tl)[0]
        )
        for i in group:
            try:
                result = translate_fn(units[i].source_text, target_lang)
                units[i].write_back(result)
            except CancelledError:
                raise
            except Exception:
                logger.warning("Failed to translate unit %d", i, exc_info=True)
                errors += 1

    return errors


# Direct XML collectors — operate on lxml elements in-place

def _is_in_heading(elem: etree._Element) -> bool:
    # Return True if *elem* is inside a heading paragraph (w:pStyle starts with 'Heading')
    # Walk up to find w:p, then check w:pStyle
    node = elem
    while node is not None:
        if node.tag == qn('w:p'):
            ppr = node.find(qn('w:pPr'))
            if ppr is not None:
                pstyle = ppr.find(qn('w:pStyle'))
                if pstyle is not None:
                    val = pstyle.get(qn('w:val'), '')
                    if val.lower().startswith('heading'):
                        return True
            return False
        node = node.getparent()
    return False


def _is_inside_del(elem: etree._Element) -> bool:
    """Return True if *elem* is nested inside a ``<w:del>`` tracked-change block.

    Text inside ``<w:del>`` represents deleted content and must NOT be
    translated — it would corrupt the tracked-changes record.
    """
    node = elem.getparent()
    while node is not None:
        if node.tag == _WDEL_TAG:
            return True
        node = node.getparent()
    return False


import re as _re

_RE_PURE_NUMBER = _re.compile(
    r"""
    ^[\s\u00a0]*           # optional leading whitespace / NBSP
    [+-]?                  # optional sign
    (?:
        \d{1,3}(?:[.,\s]\d{3})*  # thousands-grouped integer e.g. 1,234
        |\d+                      # plain integer
    )
    (?:[.,]\d+)?           # optional decimal part
    [\s\u00a0]*            # optional trailing whitespace
    (?:%|°|㎡|㎞|km|cm|mm|m²|m³|€|\$|£|¥|₹|元|円|₩)?  # optional unit
    [\s\u00a0]*$
    """,
    _re.VERBOSE,
)

# Matches strings that are MOSTLY digits — the text contains digits but also
# CJK/Latin characters that give the translator context to spell out the number.
# We extract and protect the numeric portion instead of skipping the whole unit.
_RE_HAS_DIGIT = _re.compile(r'\d')


def _is_numeric_only(text: str) -> bool:
    """Return True if *text* is a standalone number (possibly with units).

    These are never sent to the translator to prevent silent conversion of
    "42" → "forty-two" or decimal-separator reformatting.
    """
    return bool(_RE_PURE_NUMBER.match(text))


def _strip_protected_tokens(text: str) -> tuple[str, dict[str, str]]:
    """Replace numeric tokens in *text* with placeholders before translation.

    Returns (masked_text, placeholder_map).  The caller should substitute
    placeholders back after translation.

    We protect:
      - standalone integers and decimals (possibly thousands-grouped)
      - years expressed as 4-digit sequences
      - version strings like "3.14", "v2.0"
      - sequences that are >= 50% digits by character count
    """
    placeholders: dict[str, str] = {}
    counter = [0]

    def _replace(m: _re.Match) -> str:
        # Use a bracket format that is XML-safe, unlikely to appear in real text,
        # and opaque enough that translators treat it as a token to preserve.
        key = f'[[N{counter[0]}]]'
        placeholders[key] = m.group(0)
        counter[0] += 1
        return key

    # Protect numeric tokens: integers, decimals, version numbers, years
    masked = _re.sub(
        r'(?<!\w)(\d[\d,.\s]*\d|\d)(?!\w)',
        _replace,
        text,
    )
    return masked, placeholders


def _restore_protected_tokens(text: str, placeholders: dict[str, str]) -> str:
    """Substitute placeholders back into *text* after translation."""
    for key, original in placeholders.items():
        text = text.replace(key, original)
    return text


# ── Layout-protection passes ─────────────────────────────────────────────────
# Run these on each XML root AFTER translation so that text-expansion caused
# by longer translated strings doesn't break the document layout.

_WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


def _fix_table_row_heights(root: etree._Element) -> None:
    """Convert ``exact`` row-height constraints to ``atLeast``.

    When a table row uses ``w:trHeight w:hRule="exact"``, Word will clip any
    text that overflows the fixed height rather than expanding the row.  After
    translation the text is usually longer, so clipping becomes visible.
    Changing the rule to ``atLeast`` lets the row grow to fit its content while
    still respecting the designer's minimum-height intent.
    """
    TR_HEIGHT = qn('w:trHeight')
    H_RULE    = qn('w:hRule')
    for trh in root.iter(TR_HEIGHT):
        if trh.get(H_RULE, '').lower() == 'exact':
            trh.set(H_RULE, 'atLeast')
            logger.debug("Relaxed exact row height → atLeast")


def _fix_textbox_autofit(root: etree._Element) -> None:
    """Enable auto-fit (``spAutoFit``) on text-box body properties where the
    translated text might overflow the original fixed size.

    A ``<a:bodyPr>`` with no resize child — or with ``<a:noAutofit/>`` — will
    clip text that is longer than the box.  We replace ``noAutofit`` with
    ``normAutofit`` so the text box grows vertically to fit its translated
    content.  We leave ``spAutoFit`` (shape auto-fit) alone because it already
    works correctly.
    """
    A_BODY_PR    = f'{{{_NS_A}}}bodyPr'
    A_NO_AUTOFIT = f'{{{_NS_A}}}noAutofit'
    A_NORM_AUTO  = f'{{{_NS_A}}}normAutofit'
    A_SP_AUTO    = f'{{{_NS_A}}}spAutoFit'

    for body_pr in root.iter(A_BODY_PR):
        no_fit = body_pr.find(A_NO_AUTOFIT)
        if no_fit is not None:
            idx = list(body_pr).index(no_fit)
            body_pr.remove(no_fit)
            norm = etree.Element(A_NORM_AUTO)
            body_pr.insert(idx, norm)
            logger.debug("Replaced noAutofit → normAutofit in text-box bodyPr")


def _fix_frame_autosize(root: etree._Element) -> None:
    """Ensure legacy word-processing frames (``<w:framePr>``) can grow.

    Frames created with ``<w:framePr w:h="..." w:hRule="exact">`` clip content
    exactly like table rows.  We relax them to ``atLeast`` for the same reason.
    """
    FRAME_PR = qn('w:framePr')
    H_RULE   = qn('w:hRule')
    for fp in root.iter(FRAME_PR):
        if fp.get(H_RULE, '').lower() == 'exact':
            fp.set(H_RULE, 'atLeast')
            logger.debug("Relaxed framePr exact height → atLeast")


def _is_spacer_paragraph(para: etree._Element) -> bool:
    """Return True if *para* is a layout-spacer paragraph with no visible text."""
    for wt in para.iter(_WT_TAG):
        if wt.text and wt.text.strip():
            return False
    return True


def _para_has_page_break(para: etree._Element) -> bool:
    """Return True if *para* starts a new page via pageBreakBefore or a w:br page break."""
    W_P     = qn('w:p')
    W_PPR   = qn('w:pPr')
    W_PBB   = qn('w:pageBreakBefore')
    W_BR    = qn('w:br')
    W_TYPE  = qn('w:type')

    # Check w:pPr/w:pageBreakBefore
    ppr = para.find(W_PPR)
    if ppr is not None:
        pbb = ppr.find(W_PBB)
        if pbb is not None:
            val = pbb.get(qn('w:val'), '1')
            if val.lower() not in ('0', 'false'):
                return True

    # Check for explicit w:br type="page" inside any run
    for br in para.iter(W_BR):
        if br.get(W_TYPE, '').lower() == 'page':
            return True

    return False


def _collect_page_sections(root: etree._Element) -> list[list[etree._Element]]:
    """Split all body paragraphs into page sections divided by hard page breaks.

    Returns a list of sections, each section being a list of ``<w:p>`` elements
    that belong to the same logical page.  The paragraph that carries the page
    break starts the new section.
    """
    W_BODY = qn('w:body')
    W_P    = qn('w:p')
    W_TBL  = qn('w:tbl')

    body = root.find(W_BODY)
    if body is None:
        body = root  # headers/footers have no w:body wrapper

    sections: list[list[etree._Element]] = [[]]
    for child in body:
        if child.tag == W_P:
            if _para_has_page_break(child) and sections[-1]:
                sections.append([])
            sections[-1].append(child)
        # Tables count as bulk but we don't split inside them
    return [s for s in sections if s]


def _compress_section_spacing(
    section: list[etree._Element],
    expansion_ratio: float,
) -> None:
    """Compress spacing within one page section proportionally to expansion_ratio.

    Strategy
    --------
    1. Spacer paragraphs (no text): compress ``w:before/after`` by
       ``1/expansion_ratio``, clamped to [0.25, 1.0].  This reclaims the
       vertical space that would otherwise push content onto the next page.
    2. Content paragraphs: compress ``w:before/after`` more gently —
       ``max(0.55, 1/expansion_ratio)`` — only when ratio > 1.3.
    3. ``w:lineRule="exact"`` is always relaxed to ``atLeast`` regardless of
       ratio, because Latin descenders clip under exact CJK-sized line heights.
    4. Never reduce a spacing value below 0.
    """
    if expansion_ratio <= 1.0:
        # Relax exact line spacing even when text didn't expand
        for para in section:
            _relax_exact_line_spacing(para)
        return

    spacer_compression  = max(0.25, 1.0 / expansion_ratio)
    content_compression = max(0.55, 1.0 / expansion_ratio) if expansion_ratio > 1.3 else 1.0

    W_PPR     = qn('w:pPr')
    W_SPACING = qn('w:spacing')
    W_BEFORE  = qn('w:before')
    W_AFTER   = qn('w:after')

    for para in section:
        _relax_exact_line_spacing(para)
        ppr = para.find(W_PPR)
        if ppr is None:
            continue
        spacing = ppr.find(W_SPACING)
        if spacing is None:
            continue

        is_spacer = _is_spacer_paragraph(para)
        factor = spacer_compression if is_spacer else content_compression

        for attr in (W_BEFORE, W_AFTER):
            raw = spacing.get(attr)
            if raw is None:
                continue
            try:
                val = int(raw)
            except ValueError:
                continue
            spacing.set(attr, str(max(0, int(val * factor))))


def _relax_exact_line_spacing(para: etree._Element) -> None:
    """Change ``w:lineRule="exact"`` to ``atLeast`` on a single paragraph."""
    W_PPR      = qn('w:pPr')
    W_SPACING  = qn('w:spacing')
    W_LINERULE = qn('w:lineRule')

    ppr = para.find(W_PPR)
    if ppr is None:
        return
    spacing = ppr.find(W_SPACING)
    if spacing is None:
        return
    if spacing.get(W_LINERULE, '').lower() == 'exact':
        spacing.set(W_LINERULE, 'atLeast')


def _compress_spacer_spacing(root: etree._Element, expansion_ratio: float) -> None:
    """Compress spacing section-by-section, respecting hard page boundaries."""
    if expansion_ratio <= 1.0:
        return
    for section in _collect_page_sections(root):
        _compress_section_spacing(section, expansion_ratio)


def _compress_paragraph_spacing(root: etree._Element, expansion_ratio: float) -> None:
    """Relax exact line spacing on all paragraphs (handled inside section pass)."""
    # This is now a no-op — the section-aware pass in _compress_spacer_spacing
    # handles both spacer and content paragraphs together so we don't double-apply.
    pass


def _apply_layout_fixes(root: etree._Element, expansion_ratio: float = 1.0) -> None:
    """Run all layout-protection passes on *root* in one call.

    Parameters
    ----------
    expansion_ratio:
        Ratio of total translated characters to total source characters for
        this XML part.  Values > 1 mean the translation is longer (common for
        CJK→Latin).  Used to scale down spacing on spacer and content
        paragraphs so page layout is preserved.
    """
    _fix_table_row_heights(root)
    _fix_textbox_autofit(root)
    _fix_frame_autosize(root)
    _compress_spacer_spacing(root, expansion_ratio)
    _compress_paragraph_spacing(root, expansion_ratio)


def _collect_wt_units(root: etree._Element) -> list[_TranslationUnit]:
    #  Collect TranslationUnits from all ``<w:t>`` elements under *root*.
    #
    #  Skips:
    #    - empty / whitespace-only text nodes
    #    - text inside ``<w:del>`` tracked-change blocks (deleted text must be
    #      kept verbatim to preserve the revision history)
    #
    #  This is the primary collector for body text, headers, footers,
    #  footnotes and endnotes.  It mutates elements in-place via write_back.

    units: list[_TranslationUnit] = []

    for wt in root.iter(_WT_TAG):
        text = wt.text
        if not text or not text.strip():
            continue

        # Skip deleted text — translating it would corrupt tracked changes.
        if _is_inside_del(wt):
            continue

        # Skip pure numbers — translators convert "42" → "forty-two", etc.
        if _is_numeric_only(text):
            continue

        # For mixed text (e.g. "2023年"), mask numeric tokens so they survive
        # translation unchanged, then restore them in write_back.
        masked_text, placeholders = _strip_protected_tokens(text.strip())
        is_heading = _is_in_heading(wt)

        def _make_wb(elem: etree._Element, orig_text: str, ph: dict[str, str]):
            def wb(translated: str) -> None:
                # Restore any numeric placeholders the translator may have mangled
                result = _restore_protected_tokens(translated, ph)
                # Preserve leading/trailing whitespace from the original
                leading  = orig_text[: len(orig_text) - len(orig_text.lstrip())]
                trailing = orig_text[len(orig_text.rstrip()):]
                elem.text = leading + result.strip() + trailing
                if leading or trailing:
                    elem.set(
                        '{http://www.w3.org/XML/1998/namespace}space', 'preserve'
                    )
            return wb

        units.append(_TranslationUnit(
            source_text=masked_text,
            is_heading=is_heading,
            write_back=_make_wb(wt, text, placeholders),
        ))

    return units


def _collect_drawingml_units(root: etree._Element) -> list[_TranslationUnit]:
    # Collect TranslationUnits from DrawingML ``<a:t>`` elements
    units: list[_TranslationUnit] = []
    A_T_TAG = f'{{{_NS_A}}}t'

    for at_elem in root.iter(A_T_TAG):
        text = at_elem.text
        if not text or not text.strip():
            continue

        if _is_numeric_only(text):
            continue

        masked_text, placeholders = _strip_protected_tokens(text.strip())

        def _make_wb(e: etree._Element, orig_text: str, ph: dict[str, str]):
            def wb(translated: str) -> None:
                result = _restore_protected_tokens(translated, ph)
                leading  = orig_text[: len(orig_text) - len(orig_text.lstrip())]
                trailing = orig_text[len(orig_text.rstrip()):]
                e.text = leading + result.strip() + trailing
                if leading or trailing:
                    e.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            return wb

        units.append(_TranslationUnit(
            source_text=masked_text,
            write_back=_make_wb(at_elem, text, placeholders),
        ))

    return units


def _collect_vml_units(root: etree._Element) -> list[_TranslationUnit]:
    # Collect TranslationUnits from legacy VML ``<v:textpath>`` elements
    units: list[_TranslationUnit] = []
    TEXTPATH_TAG = f'{{{_NS_VML}}}textpath'

    for tp_elem in root.iter(TEXTPATH_TAG):
        text = tp_elem.get('string', '')
        if not text or not text.strip():
            continue

        def _make_wb(e: etree._Element):
            def wb(translated: str) -> None:
                e.set('string', translated)
            return wb

        units.append(_TranslationUnit(
            source_text=text,
            write_back=_make_wb(tp_elem),
        ))

    return units


def _collect_smartart_units_from_blob(
    blob: bytes,
) -> tuple[list[_TranslationUnit], etree._Element]:
    # Parse a SmartArt diagramData blob and return (units, root_element)
    root = etree.fromstring(blob)
    units: list[_TranslationUnit] = []
    A_T_TAG = f'{{{_NS_A}}}t'

    for at_elem in root.iter(A_T_TAG):
        text = at_elem.text
        if not text or not text.strip():
            continue

        def _make_wb(e: etree._Element, orig_text: str):
            def wb(translated: str) -> None:
                leading  = orig_text[: len(orig_text) - len(orig_text.lstrip())]
                trailing = orig_text[len(orig_text.rstrip()):]
                e.text = leading + translated.strip() + trailing
                if leading or trailing:
                    e.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            return wb

        units.append(_TranslationUnit(
            source_text=text.strip(),
            write_back=_make_wb(at_elem, text),
        ))

    return units, root


# Zip-level patching — the key to preserving all formatting

def _parse_xml_part(data: bytes) -> etree._Element:
    return etree.fromstring(data)


def _serialise_xml_part(root: etree._Element, original_data: bytes) -> bytes:
    """Serialise *root* back to bytes, preserving the original XML declaration."""
    # Detect encoding from original declaration (default UTF-8)
    encoding = 'UTF-8'
    if original_data.startswith(b'<?xml'):
        decl_end = original_data.index(b'?>')
        decl = original_data[: decl_end + 2].decode('ascii', errors='replace')
        if 'encoding=' in decl:
            enc_start = decl.index('encoding=') + len('encoding=') + 1
            enc_end   = decl.index(decl[enc_start - 1], enc_start)
            encoding  = decl[enc_start:enc_end]

    return etree.tostring(
        root,
        xml_declaration=True,
        encoding=encoding,
        standalone=True,
    )


def _get_xml_part_names(zip_path: str) -> list[str]:
    # Return all XML part names inside the docx zip
    with zipfile.ZipFile(zip_path, 'r') as zf:
        return [n for n in zf.namelist() if n.endswith('.xml')]


def _patch_docx_in_place(
    docx_path: str,
    patches: dict[str, bytes],
) -> None:
    # Rewrite *patches* (part_name → new_bytes) inside the docx zip in-place.
    #  All other entries in the zip are left completely untouched - this is what preserves images, styles, themes, and relationships.

    # Read the whole zip into memory first (needed because ZipFile can't overwrite individual entries without rewriting the whole archive).
    with zipfile.ZipFile(docx_path, 'r') as zf:
        names    = zf.namelist()
        originals = {n: zf.read(n) for n in names}
        infos    = {info.filename: info for info in zf.infolist()}

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', compression=zipfile.ZIP_DEFLATED) as out:
        for name in names:
            data = patches.get(name, originals[name])
            out.writestr(infos[name], data)

    with open(docx_path, 'wb') as f:
        f.write(buf.getvalue())



# Main entry point

def translate_docx(
    input_path: str,
    output_path: str,
    translator: Any,
    target_lang: str,
    *,
    cancel_event: threading.Event | None = None,
    source_lang: str = "auto",
) -> None:
    auto_detect = source_lang == "auto"
    errors = 0

    if cancel_event is not None and cancel_event.is_set():
        raise CancelledError("Translation cancelled before starting")

    # ── Step 1: copy original to output ──────────────────────────────────────
    shutil.copy2(input_path, output_path)
    logger.debug("Copied %s → %s", input_path, output_path)

    # ── Step 2: discover all XML parts to patch ───────────────────────────────
    all_xml_parts = _get_xml_part_names(output_path)

    # Collect parts: body + headers + footers + footnotes + endnotes
    text_parts = []
    for name in all_xml_parts:
        lower = name.lower()
        if (
            lower == "word/document.xml"
            or lower.startswith("word/header")
            or lower.startswith("word/footer")
            or lower == "word/footnotes.xml"
            or lower == "word/endnotes.xml"
        ):
            text_parts.append(name)

    # SmartArt diagram data parts
    smartart_parts = []
    for name in all_xml_parts:
        # content type sniffing via [Content_Types].xml is more correct, but
        # the path convention is reliable enough for our purposes
        if 'diagrams/data' in name.lower() or 'diagramdata' in name.lower():
            smartart_parts.append(name)

    logger.debug("Text parts: %s", text_parts)
    logger.debug("SmartArt parts: %s", smartart_parts)

    # ── Step 3: parse, collect, translate, serialise ─────────────────────────
    patches: dict[str, bytes] = {}

    with zipfile.ZipFile(output_path, 'r') as zf:
        raw_parts = {name: zf.read(name) for name in text_parts + smartart_parts}

    # --- text parts (w:t, a:t, VML) ---
    for part_name in text_parts:
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError("Translation cancelled during XML collection")

        raw = raw_parts[part_name]
        try:
            root = _parse_xml_part(raw)
        except etree.XMLSyntaxError:
            logger.warning("Skipping malformed XML part: %s", part_name)
            continue

        pool: list[_TranslationUnit] = []
        pool.extend(_collect_wt_units(root))
        pool.extend(_collect_drawingml_units(root))
        pool.extend(_collect_vml_units(root))

        if not pool:
            continue

        logger.debug("%s: %d units collected", part_name, len(pool))
        groups = _group_units(pool, auto_detect=auto_detect)

        # Snapshot source lengths before translation so we can compute the
        # expansion ratio used by the layout-fix passes below.
        source_chars = sum(len(u.source_text) for u in pool)

        part_errors = _translate_and_writeback(
            pool, groups, translator, target_lang, cancel_event,
        )
        errors += part_errors

        # Compute how much longer the translated text is compared to the source.
        # We measure the live elem.text values via write_back, but the simplest
        # proxy is to compare the unit texts before and after translation — the
        # write_back closures already updated the XML, so we re-read the <w:t>
        # nodes.  Instead, we use the translated strings that ended up in the
        # units.  Since write_back mutates the XML in-place we can't read them
        # back easily, so we estimate: collect current <w:t> text sum.
        translated_chars = sum(
            len(wt.text) for wt in root.iter(_WT_TAG)
            if wt.text and wt.text.strip()
        )
        expansion_ratio = (
            translated_chars / source_chars if source_chars > 0 else 1.0
        )
        logger.debug(
            "%s: expansion_ratio=%.2f (%d src → %d tgt chars)",
            part_name, expansion_ratio, source_chars, translated_chars,
        )

        # Relax fixed-size containers and compress spacing so that translated
        # (often longer) text doesn't overflow table rows, text boxes, frames,
        # or push spacer-separated sections out of place.
        _apply_layout_fixes(root, expansion_ratio)

        patches[part_name] = _serialise_xml_part(root, raw)

    # --- SmartArt parts ---
    for part_name in smartart_parts:
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError("Translation cancelled during SmartArt collection")

        raw = raw_parts[part_name]
        try:
            pool, root = _collect_smartart_units_from_blob(raw)
        except Exception:
            logger.warning("Failed to parse SmartArt part %s", part_name, exc_info=True)
            continue

        if not pool:
            continue

        logger.debug("%s: %d SmartArt units", part_name, len(pool))
        groups = _group_units(pool, auto_detect=auto_detect)
        errors += _translate_and_writeback(
            pool, groups, translator, target_lang, cancel_event,
        )
        patches[part_name] = _serialise_xml_part(root, raw)

    # ── Step 4: write only modified parts back into the zip ──────────────────
    if patches:
        logger.debug("Patching %d XML parts in %s", len(patches), output_path)
        _patch_docx_in_place(output_path, patches)

    if errors:
        raise RuntimeError(f"Translation completed with {errors} failed units")