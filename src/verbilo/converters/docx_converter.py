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


def _collect_wt_units(root: etree._Element) -> list[_TranslationUnit]:
    #  Collect TranslationUnits from all ``<w:t>`` elements under *root*. This is the primary collector for body text, headers, footers, footnotes and endnotes.  It mutates elements in-place via write_back.

    units: list[_TranslationUnit] = []

    for wt in root.iter(_WT_TAG):
        text = wt.text
        if not text or not text.strip():
            continue

        is_heading = _is_in_heading(wt)

        def _make_wb(elem: etree._Element, orig_text: str):
            def wb(translated: str) -> None:
                # Preserve leading/trailing whitespace from the original
                leading  = orig_text[: len(orig_text) - len(orig_text.lstrip())]
                trailing = orig_text[len(orig_text.rstrip()):]
                elem.text = leading + translated.strip() + trailing
                # Ensure xml:space="preserve" if there's surrounding whitespace
                if leading or trailing:
                    elem.set(
                        '{http://www.w3.org/XML/1998/namespace}space', 'preserve'
                    )
            return wb

        units.append(_TranslationUnit(
            source_text=text.strip(),
            is_heading=is_heading,
            write_back=_make_wb(wt, text),
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

        def _make_wb(e: etree._Element):
            def wb(translated: str) -> None:
                e.text = translated
            return wb

        units.append(_TranslationUnit(
            source_text=text,
            write_back=_make_wb(at_elem),
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

        def _make_wb(e: etree._Element):
            def wb(translated: str) -> None:
                e.text = translated
            return wb

        units.append(_TranslationUnit(
            source_text=text,
            write_back=_make_wb(at_elem),
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
        part_errors = _translate_and_writeback(
            pool, groups, translator, target_lang, cancel_event,
        )
        errors += part_errors

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