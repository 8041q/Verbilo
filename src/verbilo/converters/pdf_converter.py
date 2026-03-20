# in-place PDF translator using PyMuPDF; skips scanned/OCR-only PDFs
# Uses insert_htmlbox for automatic text fitting, wrapping, font selection,
# and RTL support.  No manual font management needed.

from __future__ import annotations

import logging
import re
import threading
from pathlib import Path
from typing import Any

import fitz  # PyMuPDF
from ..utils import CancelledError

logger = logging.getLogger(__name__)

_MIN_CHARS_PER_PAGE = 20  # min avg chars/page to treat as real text (not scanned)

# Span flag bit-masks as per PDF spec / PyMuPDF docs
_FLAG_ITALIC = 2
_FLAG_BOLD = 16

# Tagged-span markers for mixed-format lines
_STAG_OPEN = "\u27E8"   # ⟨
_STAG_CLOSE = "\u27E9"  # ⟩
_SPAN_TAG_RE = re.compile(r'\u27E8s(\d+)\u27E9(.*?)\u27E8/s\1\u27E9', re.DOTALL)

# Bullet/list item characters that act as paragraph boundaries.
_BULLET_CHARS: frozenset[str] = frozenset("•◦▪▸▹➜➤→·‣⁃")
_BULLET_ASCII: tuple[str, ...] = ("-", "*", "–", "—")

# Minimum coverage fraction of a text rect by an opaque drawing to treat the
# line as intentionally hidden (z-order guard).
_COVERAGE_THRESHOLD = 0.60


def _is_ocr_required(doc: fitz.Document) -> bool:
    # True if the PDF looks scanned (too few extractable chars per page)
    if doc.page_count == 0:
        return True
    total_chars = 0
    for page in doc:
        text = page.get_text("text")
        if isinstance(text, str):
            total_chars += len(text.strip())
    avg = total_chars / doc.page_count
    return avg < _MIN_CHARS_PER_PAGE


# Line-level span grouping helpers

def _span_fmt_key(span: dict) -> tuple:
    """Hashable formatting tuple for a span."""
    return (span.get("size", 11), span.get("color", 0), span.get("flags", 0))


def _line_is_bullet(text: str) -> bool:
    """Return True if the line begins with a common bullet / list marker."""
    stripped = text.lstrip()
    if not stripped:
        return False
    first = stripped[0]
    if first in _BULLET_CHARS:
        return True
    # ASCII markers only when followed by a space (avoid matching "-" in words)
    if any(stripped.startswith(m + " ") for m in _BULLET_ASCII):
        return True
    return False


def _collect_opaque_rects(page: fitz.Page) -> list[fitz.Rect]:
    """Return rects of solid fully-opaque filled drawing elements on *page*.

    These are used to detect text intentionally hidden under white/opaque shapes.
    """
    opaque: list[fitz.Rect] = []
    try:
        for drawing in page.get_drawings():
            fill = drawing.get("fill")
            if fill is None:
                continue
            if drawing.get("fill_opacity", 1.0) < 0.98:
                continue
            r = fitz.Rect(drawing["rect"])
            if not r.is_empty:
                opaque.append(r)
    except Exception:
        pass
    return opaque


def _rect_covered_by_drawing(rect: fitz.Rect, drawing_rects: list[fitz.Rect]) -> bool:
    """Return True if *rect* is significantly covered by an opaque drawing rect."""
    area = rect.width * rect.height
    if area < 1.0:
        return False
    for dr in drawing_rects:
        inter = rect & dr
        if inter.is_empty:
            continue
        if (inter.width * inter.height) / area >= _COVERAGE_THRESHOLD:
            return True
    return False


def _collect_underline_texts(page: fitz.Page) -> set[str]:
    """Return a set of text strings that are underlined on *page*.

    Primary method: underline markup annotations.
    Fallback method: parse get_text("html") for <u>…</u> tags.
    """
    underlined: set[str] = []
    try:
        # Collect quads from underline annotations
        annot_quads: list[fitz.Quad] = []
        for annot in page.annots(types=[fitz.PDF_ANNOT_UNDERLINE]):  # type: ignore[attr-defined]
            for quad in annot.vertices or []:
                annot_quads.append(fitz.Quad(quad))
        # If there are annotation quads we could intersect them with span bboxes.
        # For simplicity we just mark any text that is in the covered area.
        # Collect as text-rect pairs from the page dict for matching.
        if annot_quads:
            html_text = page.get_text("html")
            for m in re.finditer(r'<u>(.*?)</u>', html_text, re.DOTALL):
                raw = re.sub(r'<[^>]+>', '', m.group(1))
                t = raw.strip()
                if t:
                    underlined.append(t)
        else:
            # No annotations — try HTML for underline CSS
            html_text = page.get_text("html")
            for m in re.finditer(r'<u>(.*?)</u>', html_text, re.DOTALL):
                raw = re.sub(r'<[^>]+>', '', m.group(1))
                t = raw.strip()
                if t:
                    underlined.append(t)
    except Exception:
        pass
    return set(underlined)


def _group_spans_by_line(blocks: list[dict], underline_texts: set[str] | None = None) -> list[dict]:
    """Group spans into lines and return per-line metadata.

    Each line carries its ``block_id`` for paragraph grouping.
    Lines with mixed-format spans are tagged for later per-span
    HTML reconstruction.  Whitespace between adjacent spans is preserved.
    """
    lines_out: list[dict] = []
    for block_idx, block in enumerate(blocks):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue
            valid_spans = [s for s in spans if s.get("text", "").strip()]
            if not valid_spans:
                continue

            # Check if spans have mixed formatting
            fmt_set = {_span_fmt_key(s) for s in valid_spans}
            is_tagged = len(valid_spans) > 1 and len(fmt_set) > 1

            if is_tagged:
                # Build tagged text preserving each span's own whitespace.
                # Leading/trailing space on a span is placed OUTSIDE the tag
                # so that the translator sees natural word boundaries.
                parts: list[str] = []
                span_formats: list[tuple] = []
                for si, s in enumerate(valid_spans):
                    raw = s.get("text", "")
                    stripped = raw.strip()
                    leading = raw[: len(raw) - len(raw.lstrip())]
                    trailing = raw[len(raw.rstrip()):]
                    # Include leading space before tag, trailing space after tag
                    parts.append(
                        f"{leading}{_STAG_OPEN}s{si}{_STAG_CLOSE}"
                        f"{stripped}"
                        f"{_STAG_OPEN}/s{si}{_STAG_CLOSE}{trailing}"
                    )
                    span_formats.append(_span_fmt_key(s))
                combined_text = "".join(parts)
            else:
                # Preserve original span text as-is (don't strip inter-span spaces)
                combined_text = "".join(s.get("text", "") for s in valid_spans)
                span_formats = []

            if not combined_text.strip():
                continue

            # Determine predominant size / color / font by character count
            size_counts: dict[float, int] = {}
            color_counts: dict[int, int] = {}
            combined_flags = 0
            for s in valid_spans:
                n = len(s.get("text", ""))
                sz = s.get("size", 11)
                cl = s.get("color", 0)
                size_counts[sz] = size_counts.get(sz, 0) + n
                color_counts[cl] = color_counts.get(cl, 0) + n
                combined_flags |= s.get("flags", 0)

            best_size = max(size_counts, key=lambda k: size_counts[k])
            best_color = max(color_counts, key=lambda k: color_counts[k])

            # Line bounding rect = union of all span bboxes
            rects = [fitz.Rect(s["bbox"]) for s in valid_spans]
            line_rect = rects[0]
            for r in rects[1:]:
                line_rect |= r  # union

            # Underline detection: check if this line's text appears in the
            # set of underlined strings extracted from annotations / HTML.
            underline = False
            if underline_texts:
                plain = combined_text.strip()
                if plain in underline_texts:
                    underline = True

            lines_out.append({
                "rect": line_rect,
                "text": combined_text,
                "size": best_size,
                "color": best_color,
                "flags": combined_flags,
                "spans": valid_spans,
                "block_id": block_idx,
                "is_tagged": is_tagged,
                "span_formats": span_formats,
                "underline": underline,
            })
    return lines_out


def _group_lines_into_paragraphs(line_infos: list[dict]) -> list[list[int]]:
    """Group line indices by block_id for contextual translation.

    Bullet/list lines are always their own group (boundary), regardless of
    block membership, so their list marker is preserved in translation.
    """
    if not line_infos:
        return []
    groups: list[list[int]] = []
    current: list[int] = []
    current_block = None
    for i, info in enumerate(line_infos):
        bid = info.get("block_id")
        text = info.get("text", "")

        # Bullet lines are always isolated — flush current group first
        if _line_is_bullet(text):
            if current:
                groups.append(current)
                current = []
            groups.append([i])
            current_block = bid
            continue

        if bid != current_block and current:
            groups.append(current)
            current = []
        current.append(i)
        current_block = bid
    if current:
        groups.append(current)
    return groups



def _html_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("\n", "<br>")
    )


def _span_style(
    fontsize: float,
    color: int,
    flags: int,
    *,
    underline: bool = False,
) -> str:
    """Build an inline CSS style string for a single span."""
    hex_c = f"#{color:06x}" if isinstance(color, int) else "#000000"
    weight = "bold" if flags & _FLAG_BOLD else "normal"
    style = "italic" if flags & _FLAG_ITALIC else "normal"
    decoration = "underline" if underline else "none"
    return (
        f"font-size:{fontsize:.1f}px;"
        f"color:{hex_c};"
        f"font-weight:{weight};"
        f"font-style:{style};"
        f"text-decoration:{decoration};"
    )


def _infer_text_align(rect: fitz.Rect, page_width: float) -> str:
    """Infer text alignment from the position of the line rect on the page."""
    if page_width <= 0:
        return "left"
    center_x = rect.x0 + rect.width / 2
    # Centre-aligned: midpoint within 8% of page centre
    if abs(center_x - page_width / 2) < page_width * 0.08:
        return "center"
    # Right-aligned: rect starts in the right 40% of the page
    if rect.x0 > page_width * 0.60:
        return "right"
    return "left"


def _build_html(
    text: str,
    fontsize: float,
    color: int,
    flags: int,
    *,
    underline: bool = False,
    text_align: str = "left",
) -> tuple[str, str]:
    """Build an HTML snippet and CSS string for insert_htmlbox."""
    inner_style = _span_style(fontsize, color, flags, underline=underline)
    safe_text = _html_escape(text)
    html = (
        f'<div style="text-align:{text_align};">'
        f'<span style="{inner_style}">{safe_text}</span>'
        f'</div>'
    )
    # Global CSS provides a base font-size fallback for insert_htmlbox layout
    css = f"* {{font-size:{fontsize:.1f}px; font-family: sans-serif;}}"
    return html, css


def _build_multi_span_html(
    translated: str,
    span_formats: list[tuple],
    fallback_size: float,
    fallback_color: int,
    fallback_flags: int,
    *,
    underline: bool = False,
    text_align: str = "left",
) -> tuple[str, str]:
    """Build per-span HTML from tagged translated text.

    If tags parse successfully, each span gets its original formatting
    (including its own font size).  Falls back to single-span HTML if
    tags are mangled.
    """
    matches = list(_SPAN_TAG_RE.finditer(translated))
    if not matches or len(matches) < max(1, len(span_formats) // 2):
        # Tags mangled — fall back to single-span
        return _build_html(
            translated, fallback_size, fallback_color, fallback_flags,
            underline=underline, text_align=text_align,
        )

    html_parts: list[str] = []
    max_size = fallback_size
    for m in matches:
        idx = int(m.group(1))
        text = m.group(2)
        if idx < len(span_formats):
            sz, cl, fl = span_formats[idx]
        else:
            sz, cl, fl = fallback_size, fallback_color, fallback_flags
        max_size = max(max_size, sz)

        inner_style = _span_style(sz, cl, fl, underline=underline)
        safe = _html_escape(text)
        html_parts.append(f'<span style="{inner_style}">{safe}</span>')

    html = (
        f'<div style="text-align:{text_align};">'
        + "".join(html_parts)
        + "</div>"
    )
    css = f"* {{font-size:{max_size:.1f}px; font-family: sans-serif;}}"
    return html, css


# Main entry point

def translate_pdf(
    input_path: str,
    output_path: str,
    translator: Any,
    target_lang: str,
    *,
    cancel_event: threading.Event | None = None,
    source_lang: str = "auto",
) -> str | None:
    """Translate text in a PDF while preserving layout.

    1.  Extract structured text from ALL pages (blocks -> lines -> spans).
    2.  Translate everything in a single batch call to minimise API round trips.
    3.  Per page: redact original text, re-insert translated text via insert_htmlbox.

    When source_lang==\"auto\", paragraph-level line grouping is skipped so each
    line is its own translation unit, allowing per-line language auto-detection.
    """
    src = fitz.open(input_path)

    # --- OCR check: skip scanned PDFs ---
    if _is_ocr_required(src):
        logger.warning(
            "Skipping '%s': appears to be a scanned/image PDF requiring OCR.",
            Path(input_path).name,
        )
        src.close()
        return "skipped-ocr"

    errors = 0

    # --- Phase 1: collect all lines across all pages, grouped by block ---
    # Each entry: (page_num, line_infos, para_groups)
    page_line_data: list[tuple[int, list[dict], list[list[int]]]] = []
    # Translation units: one per paragraph group (lines joined with \n)
    all_units: list[str] = []
    # Map each unit back: (page_data_idx, group_idx)
    unit_map: list[tuple[int, int]] = []

    for page_num in range(src.page_count):
        if cancel_event is not None and cancel_event.is_set():
            src.close()
            raise CancelledError("Translation cancelled")

        page = src[page_num]
        page_width = page.rect.width

        # Collect per-page helpers for formatting preservation
        underline_texts = _collect_underline_texts(page)
        opaque_rects = _collect_opaque_rects(page)

        text_dict: dict[str, Any] = page.get_text(  # type: ignore[assignment]
            "dict", flags=fitz.TEXT_PRESERVE_WHITESPACE
        )
        blocks = text_dict.get("blocks", [])
        line_infos = _group_spans_by_line(blocks, underline_texts=underline_texts)

        # Stamp page_width into each line for alignment inference, and skip
        # lines that are intentionally covered by opaque drawing shapes.
        visible_infos: list[dict] = []
        for li in line_infos:
            if _rect_covered_by_drawing(li["rect"], opaque_rects):
                logger.debug(
                    "Skipping covered line on page %d: %s", page_num + 1, li["text"][:40]
                )
                continue
            li["page_width"] = page_width
            visible_infos.append(li)

        # In auto-detect mode skip block-level grouping so each line is its own
        # translation unit, letting the API auto-detect the language per line.
        if source_lang == "auto":
            para_groups = [[i] for i in range(len(visible_infos))]
        else:
            para_groups = _group_lines_into_paragraphs(visible_infos)

        pdi = len(page_line_data)
        page_line_data.append((page_num, visible_infos, para_groups))

        for gi, group in enumerate(para_groups):
            if len(group) == 1:
                all_units.append(line_infos[group[0]]["text"])
            else:
                all_units.append("\n".join(line_infos[idx]["text"] for idx in group))
            unit_map.append((pdi, gi))

    # --- Phase 2: single batch translation for all paragraph groups ---
    if all_units:
        try:
            all_translated = translator.translate_batch(
                all_units, target_lang, cancel_event=cancel_event
            )
        except CancelledError:
            src.close()
            raise
        except Exception:
            logger.exception(
                "Batch translation failed for PDF; falling back to per-item"
            )
            all_translated = []
            for t in all_units:
                try:
                    r = translator.translate_text(t, target_lang)
                    all_translated.append(r if r is not None else t)
                except Exception:
                    logger.exception("Per-item fallback failed")
                    all_translated.append(t)
                    errors += 1
    else:
        all_translated = []

    # --- Phase 2b: distribute translated units back to per-line results ---
    # Build a flat per-line translated text list per page
    page_translated: dict[int, dict[int, str]] = {}  # pdi -> {line_idx: translated}
    for unit_idx, (pdi, gi) in enumerate(unit_map):
        page_num, line_infos, para_groups = page_line_data[pdi]
        group = para_groups[gi]
        tr_text = all_translated[unit_idx] if unit_idx < len(all_translated) else None

        line_map = page_translated.setdefault(pdi, {})
        if tr_text is None:
            for lidx in group:
                line_map[lidx] = line_infos[lidx]["text"]
            errors += 1
        elif len(group) == 1:
            line_map[group[0]] = tr_text
        else:
            parts = tr_text.split("\n")
            if len(parts) == len(group):
                for lidx, part in zip(group, parts):
                    line_map[lidx] = part
            else:
                # Newline count mismatch — distribute proportionally
                # or just use the full translated text for the first line
                # and empty for the rest (safe fallback).
                for j, lidx in enumerate(group):
                    if j < len(parts):
                        line_map[lidx] = parts[j]
                    else:
                        line_map[lidx] = ""

    # --- Phase 3: apply translations per page ---
    for pdi, (page_num, line_infos, _para_groups) in enumerate(page_line_data):
        if not line_infos:
            continue

        if cancel_event is not None and cancel_event.is_set():
            src.close()
            raise CancelledError("Translation cancelled")

        page = src[page_num]
        line_map = page_translated.get(pdi, {})

        # Redact original text per-line
        for info in line_infos:
            page.add_redact_annot(info["rect"], fill=None)  # type: ignore[arg-type]
        page.apply_redactions(images=0)  # type: ignore[arg-type]

        # Insert translated text using insert_htmlbox
        for li_idx, info in enumerate(line_infos):
            tr_text = line_map.get(li_idx, info["text"])

            orig_rect: fitz.Rect = info["rect"]
            fontsize: float = info["size"]
            underline: bool = info.get("underline", False)
            page_width: float = info.get("page_width", 0.0)
            text_align: str = _infer_text_align(orig_rect, page_width)

            if info.get("is_tagged") and info.get("span_formats"):
                html, css = _build_multi_span_html(
                    tr_text,
                    info["span_formats"],
                    fontsize,
                    info["color"],
                    info["flags"],
                    underline=underline,
                    text_align=text_align,
                )
            else:
                html, css = _build_html(
                    tr_text, fontsize, info["color"], info["flags"],
                    underline=underline,
                    text_align=text_align,
                )

            try:
                result = page.insert_htmlbox(orig_rect, html, css=css)
                if result[0] < 0:
                    logger.debug(
                        "insert_htmlbox could not fit text on page %d at rect %s",
                        page_num + 1,
                        orig_rect,
                    )
            except Exception:
                logger.exception(
                    "Failed to insert translated text on page %d", page_num + 1
                )
                errors += 1

    # Check for cancellation before saving
    if cancel_event is not None and cancel_event.is_set():
        src.close()
        raise CancelledError("Translation cancelled before saving PDF")

    src.save(str(output_path), garbage=4, deflate=True, clean=True)
    src.close()

    if errors:
        raise RuntimeError(f"PDF translation completed with {errors} issues")
    return None
