# PDF translator using PyMuPDF — preserves layout, skips scanned/OCR-only PDFs.
#
# ARCHITECTURE: block-centric
# ─────────────────────────────
# PDF text is organised as: Page → Blocks → Lines → Spans
# A *block* is the natural translation unit: it represents one logical region
# (e.g. a heading + body paragraph, or a standalone label).  The original block
# bounding-box is the "intended container" for that text.
#
# Pipeline:
#   1. Extract all blocks from every page (type=0 text blocks only).
#   2. Build one translation unit per block (all lines joined by \n).
#   3. Batch-translate all units in one API call.
#   4. Per page:
#      a. Redact each original block rect (removes original text cleanly).
#      b. Rebuild HTML for the translated block, preserving the original per-line
#         formatting (size, weight, colour) from the stored span metadata.
#      c. Call insert_htmlbox with scale_low=0 — PyMuPDF auto-scales to fit.
#      d. If the returned scale < SCALE_THRESHOLD, the text was shrunk too much;
#         try expanding the block rect downward into obstacle/sibling-free space.

from __future__ import annotations

import logging
import re
import threading
import time
from pathlib import Path
from typing import Any, Callable, Mapping

import fitz  # PyMuPDF >= 1.24
from ..advisors import NullAdvisor
from ..utils import CancelledError
from ..semantic import TranslationConstraints, TranslationService, TranslationUnit as SemanticTranslationUnit
from ..progress import ProgressReporter, ProgressUpdate

logger = logging.getLogger(__name__)

# ── tuneable constants ──────────────────────────────────────────────────────
_MIN_CHARS_PER_PAGE  = 20    # avg chars/page below which PDF is treated as scanned
_SCALE_THRESHOLD     = 0.78  # if insert_htmlbox scale < this, try rect expansion
_CJK_SCALE_THRESHOLD = 0.86  # CJK-origin blocks need a higher readable floor
_SCALE_EPSILON       = 0.015 # treat near-identical scales as equivalent
_MAX_EXPAND_DOWN     = 60.0  # max points a block rect may grow downward
_EXPAND_STEP         = 6.0   # vertical expansion increment (points)
_OBSTACLE_MIN_AREA   = 200   # filled drawings smaller than this (pt²) are ignored
_LINE_OBSTACLE_THICKNESS = 2.5   # thin stroked lines can be table/cell borders
_LINE_OBSTACLE_MIN_SPAN  = 12.0  # ignore tiny decorative strokes
_CELL_BOUNDARY_PADDING = 3.0     # keep translated text off line-drawn borders
_MAX_EXPAND_LEFT = 48.0          # conservative leftward slack inside a cell
_MAX_EXPAND_UP = 24.0            # conservative upward slack inside a cell
_HEURISTIC_LITERAL_MAX_CHARS = 18
_HEURISTIC_LABEL_MAX_CHARS = 42
_HEURISTIC_ROOMY_FREE_RATIO = 0.72
_HEURISTIC_ROOMY_CJK_FREE_RATIO = 0.33
_LAYOUT_RETRY_RISK_RATIO = 0.92
_LAYOUT_RETRY_TARGET_RATIO = 0.85
# ────────────────────────────────────────────────────────────────────────────


# ── OCR guard ────────────────────────────────────────────────────────────────

def _is_ocr_required(doc: fitz.Document) -> bool:
    """Return True when the PDF has too little extractable text (likely scanned)."""
    if doc.page_count == 0:
        return True
    total = sum(len(p.get_text("text").strip()) for p in doc)
    return (total / doc.page_count) < _MIN_CHARS_PER_PAGE


# ── font / style helpers ─────────────────────────────────────────────────────

# CJK font families whose "-Medium" variant is visually regular weight
_CJK_REGULAR_MEDIUM = {
    "notosanscjk", "notoserifcjk", "sourcehansans", "sourcehanserif",
    "stsong", "stkaiti", "stheiti", "stfangsong", "stzhongsong",
    "microsoftyahei", "microsoftjhenghei", "simsun", "simhei",
    "fangsonggb", "fangsong", "kaiti", "mingliu", "pmingliu",
    "mssong", "msyahei", "nsimsun", "dengxian", "华文",
}


def _css_weight(font_name: str, flags: int) -> int:
    """Map a PDF span font name + flags to a CSS font-weight integer.

    The PDF bold flag (bit 16) is the most reliable indicator.  Font-name
    heuristics are used only as a secondary signal, with extra care for CJK
    fonts where "-Medium" is the normal/regular weight.
    """
    # Primary: trust the PDF bold flag
    if flags & 16:
        return 700

    name_lc = font_name.lower()
    # Strip common prefixes that confuse substring matching
    base = name_lc.replace("-", "").replace("_", "").replace(" ", "")

    # Exact suffix matching only — avoid false positives from substrings
    if base.endswith("bold") or base.endswith("black") or base.endswith("heavy"):
        return 700
    if base.endswith("semibold") or base.endswith("demibold"):
        return 600

    # "Medium" in CJK fonts means regular weight (400), not CSS 500
    if "medium" in name_lc:
        is_cjk = any(fam in base for fam in _CJK_REGULAR_MEDIUM)
        return 400 if is_cjk else 500

    if base.endswith("light") or base.endswith("thin"):
        return 300
    return 400


def _css_style(flags: int) -> str:
    return "italic" if flags & 2 else "normal"


def _css_color(color: int) -> str:
    if isinstance(color, int):
        return f"#{color:06x}"
    return "#000000"


def _css_font_family(font_name: str, flags: int) -> str:
    # Embedded PDF font names are not automatically available to Story / HTML
    # rendering. Preserve the closest generic family instead of forcing every
    # translated line to sans-serif. PyMuPDF text flags: serif=4, mono=8.
    name = (font_name or "").lower()
    if (flags & 8) or any(k in name for k in ("mono", "courier", "console")):
        return "monospace"
    if (flags & 4) or any(k in name for k in ("serif", "times", "mincho", "song")):
        return "serif"
    return "sans-serif"


def _rotation_from_dir(direction: object) -> int:
    # Convert PyMuPDF line direction vectors to insert_htmlbox's 90-degree
    # rotation values. Non-axis-aligned text is left unrotated rather than
    # pretending we can faithfully reproduce arbitrary transforms.
    try:
        x, y = direction  # type: ignore[misc]
        x = float(x)
        y = float(y)
    except Exception:
        return 0
    if abs(x) >= 0.92 and abs(y) <= 0.25:
        return 0 if x >= 0 else 180
    if abs(y) >= 0.92 and abs(x) <= 0.25:
        return 90 if y < 0 else 270
    return 0


def _html_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
    )


def _normalize_lang_code(lang: str) -> str:
    return lang.strip().replace("_", "-").lower()


def _is_cjk_char(ch: str) -> bool:
    cp = ord(ch)
    return (
        0x3400 <= cp <= 0x4DBF or
        0x4E00 <= cp <= 0x9FFF or
        0xF900 <= cp <= 0xFAFF or
        0x3040 <= cp <= 0x30FF or
        0xAC00 <= cp <= 0xD7AF
    )


def _is_cjk_source_block(text: str, source_lang: str = "auto") -> bool:
    lang = _normalize_lang_code(source_lang)
    if lang.startswith(("zh", "ja", "ko")):
        return True

    visible = [ch for ch in text if not ch.isspace()]
    if not visible:
        return False

    cjk_count = sum(1 for ch in visible if _is_cjk_char(ch))
    return cjk_count >= 2 and (cjk_count / len(visible)) >= 0.35


def _preferred_scale_for_block(source_text: str, source_lang: str = "auto") -> float:
    if _is_cjk_source_block(source_text, source_lang):
        return _CJK_SCALE_THRESHOLD
    return _SCALE_THRESHOLD


# ── block extraction ─────────────────────────────────────────────────────────

def _line_record(line: dict) -> dict[str, Any] | None:
    spans = line.get("spans", [])
    valid = [s for s in spans if s.get("text", "").strip()]
    if not valid:
        return None

    parts: list[str] = [valid[0]["text"]]
    for k in range(1, len(valid)):
        prev_bbox = fitz.Rect(valid[k - 1]["bbox"])
        cur_bbox = fitz.Rect(valid[k]["bbox"])
        gap = cur_bbox.x0 - prev_bbox.x1
        avg_cw = prev_bbox.width / max(len(valid[k - 1]["text"]), 1)
        if gap > avg_cw * 0.5:
            parts.append(" ")
        parts.append(valid[k]["text"])
    line_text = "".join(parts).strip()
    if not line_text:
        return None

    weight_chars: dict[int, int] = {}
    for s in valid:
        w = _css_weight(s["font"], s["flags"])
        weight_chars[w] = weight_chars.get(w, 0) + len(s["text"])
    final_weight = min(weight_chars, key=lambda w: (-weight_chars[w], w))
    dominant = max(valid, key=lambda s: len(s["text"]))
    rect = fitz.Rect(line["bbox"])
    style = {
        "size": dominant["size"],
        "weight": final_weight,
        "style": _css_style(dominant["flags"]),
        "color": _css_color(dominant["color"]),
        "font_family": _css_font_family(dominant["font"], dominant["flags"]),
        "rotation": _rotation_from_dir(line.get("dir", (1, 0))),
        "align": "left",
    }
    return {"text": line_text, "style": style, "rect": rect}


def _rect_contains_center(container: fitz.Rect, item: fitz.Rect, tolerance: float = 1.5) -> bool:
    cx = (item.x0 + item.x1) / 2
    cy = (item.y0 + item.y1) / 2
    return (
        container.x0 - tolerance <= cx <= container.x1 + tolerance
        and container.y0 - tolerance <= cy <= container.y1 + tolerance
    )


def _dominant_style(records: list[dict[str, Any]]) -> dict[str, Any]:
    if not records:
        return {
            "size": 11.0, "weight": 400, "style": "normal", "color": "#000000",
            "font_family": "sans-serif", "rotation": 0, "align": "left",
        }
    winner = max(records, key=lambda r: len(r["text"]))
    return dict(winner["style"])


def _infer_alignment_in_container(
    records: list[dict[str, Any]],
    container: fitz.Rect,
) -> str:
    if not records:
        return "left"
    votes = {"left": 0, "center": 0, "right": 0}
    center_x = (container.x0 + container.x1) / 2
    center_tol = max(container.width * 0.12, 5.0)
    edge_tol = max(container.width * 0.08, 4.0)
    for rec in records:
        r = rec["rect"]
        weight = max(len(rec["text"]), 1)
        rcx = (r.x0 + r.x1) / 2
        if abs(rcx - center_x) <= center_tol:
            votes["center"] += weight
        elif container.x1 - r.x1 <= edge_tol:
            votes["right"] += weight
        else:
            votes["left"] += weight
    return max(votes, key=votes.get)


def _join_visual_lines(records: list[dict[str, Any]]) -> str:
    """Join soft-wrapped visual lines into one translation unit.

    PDF line breaks are layout artifacts, not necessarily sentence boundaries.
    Table cells in particular often split a single value across several extracted
    blocks. Preserve explicit bullet/list boundaries but otherwise let the target
    language reflow inside the cell.
    """
    if not records:
        return ""
    ordered = sorted(records, key=lambda r: (r["rect"].y0, r["rect"].x0))
    result = ordered[0]["text"].strip()
    bullet_re = re.compile(r"^[\s]*[●•▪■◆◇▶►✓√×]|^[\s]*\d+[.)、]")
    for rec in ordered[1:]:
        nxt = rec["text"].strip()
        if not nxt:
            continue
        if bullet_re.match(nxt):
            sep = "\n"
        elif result.endswith(("-", "–", "—", "/", "\\")):
            sep = ""
        else:
            prev_last = result[-1:] if result else ""
            next_first = nxt[:1]
            # CJK wraps do not need an inserted Latin space.
            if (prev_last and _is_cjk_char(prev_last)) or (next_first and _is_cjk_char(next_first)):
                sep = ""
            else:
                sep = " "
        result += sep + nxt
    return result


def _safe_inset_rect(rect: fitz.Rect, inset: float = 2.5) -> fitz.Rect:
    if rect.width <= inset * 2 + 1 or rect.height <= inset * 2 + 1:
        return fitz.Rect(rect)
    return fitz.Rect(rect.x0 + inset, rect.y0 + inset, rect.x1 - inset, rect.y1 - inset)


def _extract_table_blocks(
    page: fitz.Page,
    line_records: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], set[int]]:
    """Return table-cell blocks plus line indices consumed by those cells.

    PyMuPDF's normal block extraction frequently merges all three cells of a table
    row into one text block. Re-inserting that block destroys the column layout.
    ``find_tables`` gives us the actual cell rectangles, so each cell becomes its
    own translation / fitting container instead.
    """
    finder = getattr(page, "find_tables", None)
    if not callable(finder):
        return [], set()
    try:
        found = finder()
        tables = list(getattr(found, "tables", []) or [])
    except Exception:
        logger.debug("PDF table detection failed", exc_info=True)
        return [], set()

    page_w = page.rect.width
    page_h = page.rect.height
    consumed: set[int] = set()
    result: list[dict[str, Any]] = []
    seen_cells: set[tuple[float, float, float, float]] = set()

    for table_index, table in enumerate(tables):
        rows = list(getattr(table, "rows", []) or [])
        for row_index, row in enumerate(rows):
            cells = list(getattr(row, "cells", []) or [])
            for col_index, cell_bbox in enumerate(cells):
                if not cell_bbox:
                    continue
                cell = fitz.Rect(cell_bbox)
                key = tuple(round(v, 2) for v in (cell.x0, cell.y0, cell.x1, cell.y1))
                if key in seen_cells:
                    continue
                seen_cells.add(key)

                ids = [
                    i for i, rec in enumerate(line_records)
                    if i not in consumed and _rect_contains_center(cell, rec["rect"])
                ]
                if not ids:
                    continue
                records = [line_records[i] for i in ids]
                for i in ids:
                    consumed.add(i)

                inner = _safe_inset_rect(cell, _CELL_BOUNDARY_PADDING)
                style = _dominant_style(records)
                style["align"] = _infer_alignment_in_container(records, inner)
                content_y0 = min(r["rect"].y0 for r in records)
                # Preserve the original vertical placement when there is room, but
                # allow the layout planner to try a no-padding fallback if English
                # expansion needs the extra height.
                padding_top = max(0.0, min(content_y0 - inner.y0, inner.height * 0.38))
                rotation_chars: dict[int, int] = {}
                for rec in records:
                    rot = int(rec["style"].get("rotation", 0))
                    rotation_chars[rot] = rotation_chars.get(rot, 0) + len(rec["text"])
                rotation = max(rotation_chars, key=rotation_chars.get) if rotation_chars else 0

                text = _join_visual_lines(records)
                if not text.strip():
                    continue
                result.append({
                    "rect": inner,
                    "text": text,
                    "line_styles": [style],
                    "line_rects": [fitz.Rect(r["rect"]) for r in records],
                    "page_width": page_w,
                    "page_height": page_h,
                    "rotation": rotation,
                    "is_table_cell": True,
                    "fixed_container": True,
                    "table_index": table_index,
                    "table_row": row_index,
                    "table_col": col_index,
                    "padding_top_hint": padding_top,
                    "strategy": "semantic",
                    "advisor_reason": "table-cell",
                    "content_hint": "table-cell",
                })

    return result, consumed


def _has_parallel_lines(records: list[dict[str, Any]]) -> bool:
    """True when one MuPDF block actually contains side-by-side regions."""
    for i, left in enumerate(records):
        a = left["rect"]
        for right in records[i + 1:]:
            b = right["rect"]
            overlap = max(0.0, min(a.y1, b.y1) - max(a.y0, b.y0))
            if overlap < min(a.height, b.height) * 0.45:
                continue
            gap = max(b.x0 - a.x1, a.x0 - b.x1)
            if gap > max(10.0, min(a.width, b.width) * 0.12):
                return True
    return False


def _make_single_line_block(
    rec: dict[str, Any],
    page_w: float,
    page_h: float,
) -> dict[str, Any]:
    r = fitz.Rect(rec["rect"])
    xpad = min(30.0, max(6.0, r.width * 0.12))
    ypad = min(3.0, max(1.0, r.height * 0.12))
    container = fitz.Rect(
        max(0.0, r.x0 - xpad), max(0.0, r.y0 - ypad),
        min(page_w, r.x1 + xpad), min(page_h, r.y1 + ypad),
    )
    style = dict(rec["style"])
    if abs(((r.x0 + r.x1) / 2) - page_w / 2) <= page_w * 0.035:
        style["align"] = "center"
    else:
        style["align"] = "left"
    return {
        "rect": container,
        "text": rec["text"],
        "line_styles": [style],
        "line_rects": [r],
        "page_width": page_w,
        "page_height": page_h,
        "rotation": int(style.get("rotation", 0)),
        "parallel_region": True,
    }


def _extract_blocks(page: fitz.Page) -> list[dict]:
    """Extract layout-aware translation regions from a PDF page.

    Tables are cell-centric; ordinary content remains block-centric, except MuPDF
    blocks containing side-by-side lines are split back into independent regions.
    This prevents translated content from collapsing into the left-most column.
    """
    page_w = page.rect.width
    page_h = page.rect.height
    d = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)

    block_records: list[list[dict[str, Any]]] = []
    all_records: list[dict[str, Any]] = []
    record_ids_by_block: list[list[int]] = []
    for block in d.get("blocks", []):
        if block.get("type") != 0:
            block_records.append([])
            record_ids_by_block.append([])
            continue
        records: list[dict[str, Any]] = []
        ids: list[int] = []
        for line in block.get("lines", []):
            rec = _line_record(line)
            if rec is None:
                continue
            ids.append(len(all_records))
            all_records.append(rec)
            records.append(rec)
        block_records.append(records)
        record_ids_by_block.append(ids)

    table_blocks, consumed = _extract_table_blocks(page, all_records)
    out: list[dict[str, Any]] = list(table_blocks)

    for block, records, ids in zip(d.get("blocks", []), block_records, record_ids_by_block):
        if block.get("type") != 0 or not records:
            continue
        residual = [rec for rec, rid in zip(records, ids) if rid not in consumed]
        if not residual:
            continue

        if _has_parallel_lines(residual):
            out.extend(_make_single_line_block(rec, page_w, page_h) for rec in residual)
            continue

        lines_data = [(rec["text"], dict(rec["style"]), fitz.Rect(rec["rect"])) for rec in residual]
        block_rect = fitz.Rect(
            min(r.x0 for _, _, r in lines_data),
            min(r.y0 for _, _, r in lines_data),
            max(r.x1 for _, _, r in lines_data),
            max(r.y1 for _, _, r in lines_data),
        )
        n_lines = len(lines_data)
        if n_lines >= 3:
            x0s = [lr.x0 for _, _, lr in lines_data]
            x1s = [lr.x1 for _, _, lr in lines_data]
            cxs = [(lr.x0 + lr.x1) / 2 for _, _, lr in lines_data]
            x0_range = max(x0s) - min(x0s)
            x1_range = max(x1s) - min(x1s)
            cx_range = max(cxs) - min(cxs)
            avg_w = sum(lr.width for _, _, lr in lines_data) / n_lines
            tolerance = max(avg_w * 0.12, 4.0)
            if cx_range < tolerance and x0_range > tolerance:
                block_align = "center"
            elif x1_range < tolerance and x0_range > tolerance:
                block_align = "right"
            else:
                block_align = "left"
        else:
            bcx = (block_rect.x0 + block_rect.x1) / 2
            block_align = "left"
            if block_rect.width < page_w * 0.50 and abs(bcx - page_w / 2) < page_w * 0.03:
                block_align = "center"
        for _, style, _ in lines_data:
            style["align"] = block_align

        rotation_chars: dict[int, int] = {}
        for line_text, style, _ in lines_data:
            rot = int(style.get("rotation", 0))
            rotation_chars[rot] = rotation_chars.get(rot, 0) + len(line_text)
        block_rotation = max(rotation_chars, key=rotation_chars.get) if rotation_chars else 0

        out.append({
            "rect": block_rect,
            "text": "\n".join(t for t, _, _ in lines_data),
            "line_styles": [s for _, s, _ in lines_data],
            "line_rects": [lr for _, _, lr in lines_data],
            "page_width": page_w,
            "page_height": page_h,
            "rotation": block_rotation,
        })

    # Keep deterministic visual order for translation / placement.
    out.sort(key=lambda b: (round(fitz.Rect(b["rect"]).y0, 2), round(fitz.Rect(b["rect"]).x0, 2)))
    return out


# ── HTML builder ─────────────────────────────────────────────────────────────

def _build_block_html(
    translated: str,
    line_styles: list[dict],
    *,
    line_height: float = 1.15,
    align_override: str | None = None,
    padding_top: float = 0.0,
) -> tuple[str, str]:
    # Build HTML + CSS for *translated* text using the original per-line styles.
    orig_lines = translated.split("\n")
    n = len(line_styles)
    max_size = max((s["size"] for s in line_styles), default=11.0)

    # Use the dominant alignment (most common across lines)
    align_counts: dict[str, int] = {}
    for s in line_styles:
        align_counts[s["align"]] = align_counts.get(s["align"], 0) + 1
    dominant_align = max(align_counts, key=align_counts.get)  # type: ignore[arg-type]
    container_align = align_override or dominant_align
    top_padding = max(padding_top, 0.0)

    parts: list[str] = []
    for i, line_text in enumerate(orig_lines):
        s = line_styles[min(i, n - 1)]
        # Use block-level <div> per line so PyMuPDF's HTML renderer treats each
        # line as its own block — inline <span>+<br/> can collapse onto one line
        # when consecutive spans have different font sizes or the renderer treats
        # the break as optional whitespace.
        inline = (
            f"font-family:{s.get('font_family', 'sans-serif')};"
            f"font-size:{s['size']:.1f}px;"
            f"font-weight:{s['weight']};"
            f"font-style:{s['style']};"
            f"color:{s['color']};"
            f"text-decoration:none;"
            f"margin:0; padding:0;"
            f"overflow-wrap:anywhere; word-break:break-word;"
        )
        safe = _html_escape(line_text)
        parts.append(f'<div style="{inline}">{safe}</div>')

    html = (
        f'<div style="text-align:{container_align}; margin:0; padding:0;'
        f' padding-top:{top_padding:.1f}px; box-sizing:border-box;">'
        + "".join(parts)
        + "</div>"
    )
    css = (
        f"* {{font-family:sans-serif; font-size:{max_size:.1f}px;"
        f" font-weight:normal; font-style:normal; text-decoration:none;"
        f" margin:0; padding:0; line-height:{line_height:.2f};}}"
    )
    return html, css


def _placement_overrides_for_block(
    orig_rect: fitz.Rect,
    fit_rect: fitz.Rect,
    line_styles: list[dict],
) -> dict[str, Any]:
    if not line_styles:
        return {}

    # Detect any horizontal or upward expansion.  The original `len > 2` guard
    # was removed: multi-line blocks need the same padding_top correction as
    # short blocks when the rect was pushed upward into free space — without it,
    # text renders above the original block area and overlaps nearby shapes.
    expanded_x = fit_rect.x0 + 0.5 < orig_rect.x0 and fit_rect.x1 > orig_rect.x1 + 0.5
    # Check for upward expansion regardless of whether y1 also moved.
    expanded_y = fit_rect.y0 + 0.5 < orig_rect.y0
    if not (expanded_x or expanded_y):
        return {}

    overrides: dict[str, Any] = {}

    if expanded_x:
        fit_cx = (fit_rect.x0 + fit_rect.x1) / 2
        orig_cx = (orig_rect.x0 + orig_rect.x1) / 2
        if abs(orig_cx - fit_cx) <= max(fit_rect.width * 0.12, 6.0):
            overrides["align"] = "center"

    if expanded_y:
        top_inset = max(orig_rect.y0 - fit_rect.y0, 0.0)
        bottom_inset = max(fit_rect.y1 - orig_rect.y1, 0.0)
        if top_inset > 1.0:
            if bottom_inset > 1.0:
                # Symmetric expansion: centre check — only pad if the expanded
                # rect is roughly centred on the original (otherwise the rect
                # grew asymmetrically and the scale/flow already handles it).
                fit_cy  = (fit_rect.y0 + fit_rect.y1) / 2
                orig_cy = (orig_rect.y0 + orig_rect.y1) / 2
                if abs(orig_cy - fit_cy) <= max(fit_rect.height * 0.18, 4.0):
                    overrides["padding_top"] = min(
                        top_inset,
                        bottom_inset,
                        max(2.0, fit_rect.height * 0.18),
                    )
            else:
                # Upward-only expansion: unconditionally offset text back down
                # to where the original block started so it doesn't float above
                # shapes that live between fit_rect.y0 and orig_rect.y0.
                overrides["padding_top"] = top_inset

    return overrides


def _layout_variants_for_block(
    source_text: str,
    source_lang: str = "auto",
) -> list[dict[str, Any]]:
    compact_line_height = 1.00 if _is_cjk_source_block(source_text, source_lang) else 1.08
    return [
        {"name": "default", "line_height": 1.15, "priority": 0},
        {"name": "compact", "line_height": compact_line_height, "priority": 1},
    ]


def _collect_obstacles(page: fitz.Page) -> list[fitz.Rect]:
    # Return rects of all image blocks and significant filled drawings
    obs: list[fitz.Rect] = []

    for block in page.get_text("dict")["blocks"]:
        if block.get("type") == 1:
            r = fitz.Rect(block["bbox"])
            if not r.is_empty:
                obs.append(r)

    try:
        for drawing in page.get_drawings():
            r = fitz.Rect(drawing["rect"])
            if r.is_empty:
                continue

            if drawing.get("stroke_opacity", 1.0) >= 0.1:
                thin = min(r.width, r.height)
                span = max(r.width, r.height)
                if (thin <= _LINE_OBSTACLE_THICKNESS
                        and span >= _LINE_OBSTACLE_MIN_SPAN):
                    obs.append(fitz.Rect(r.x0 - 1.0, r.y0 - 1.0,
                                         r.x1 + 1.0, r.y1 + 1.0))
                    continue

            if drawing.get("fill") is None:
                continue
            if drawing.get("fill_opacity", 1.0) < 0.1:
                continue
            if r.width * r.height < _OBSTACLE_MIN_AREA:
                continue
            obs.append(r)
    except Exception:
        pass

    return obs


def _free_y1(
    rect: fitz.Rect,
    page_height: float,
    obstacles: list[fitz.Rect],
    siblings: list[fitz.Rect],
) -> float:
    """Return the lowest y1 reachable from *rect* without colliding with anything."""
    limit = page_height - _CELL_BOUNDARY_PADDING
    for other in obstacles + siblings:
        # Only elements that horizontally overlap with our rect matter
        if other.x0 >= rect.x1 - 1 or other.x1 <= rect.x0 + 1:
            continue
        if other.y0 > rect.y1 and other.y0 < limit:
            limit = other.y0 - _CELL_BOUNDARY_PADDING
    return max(limit, rect.y1)


def _free_y0(
    rect: fitz.Rect,
    obstacles: list[fitz.Rect],
    siblings: list[fitz.Rect],
) -> float:
    """Return the highest safe y0 reachable from *rect* within local boundaries."""
    limit: float | None = None
    for other in obstacles + siblings:
        # Only elements that horizontally overlap with our rect matter
        if other.x0 >= rect.x1 - 1 or other.x1 <= rect.x0 + 1:
            continue
        if other.y1 < rect.y0:
            bound = other.y1 + _CELL_BOUNDARY_PADDING
            limit = bound if limit is None else max(limit, bound)
    if limit is None:
        return rect.y0
    return min(limit, rect.y0)


def _free_x1(
    rect: fitz.Rect,
    page_width: float,
    obstacles: list[fitz.Rect],
    siblings: list[fitz.Rect],
) -> float:
    # Return the rightmost x1 reachable from *rect* without colliding
    limit = page_width - _CELL_BOUNDARY_PADDING
    for other in obstacles + siblings:
        # Only elements that vertically overlap with our rect matter
        if other.y0 >= rect.y1 - 1 or other.y1 <= rect.y0 + 1:
            continue
        if other.x0 > rect.x1 and other.x0 < limit:
            limit = other.x0 - _CELL_BOUNDARY_PADDING
    return max(limit, rect.x1)


def _free_x0(
    rect: fitz.Rect,
    obstacles: list[fitz.Rect],
    siblings: list[fitz.Rect],
) -> float:
    """Return the leftmost safe x0 reachable from *rect* within local boundaries."""
    limit: float | None = None
    for other in obstacles + siblings:
        # Only elements that vertically overlap with our rect matter
        if other.y0 >= rect.y1 - 1 or other.y1 <= rect.y0 + 1:
            continue
        if other.x1 < rect.x0:
            bound = other.x1 + _CELL_BOUNDARY_PADDING
            limit = bound if limit is None else max(limit, bound)
    if limit is None:
        return rect.x0
    return min(limit, rect.x0)


_MAX_EXPAND_H = 120.0  # max points a block rect may grow rightward
_EXPAND_H_STEP = 8.0   # horizontal expansion increment (points)


def _expansion_values(max_expand: float, step: float) -> list[float]:
    max_expand = max(0.0, max_expand)
    values = [0.0]
    grown = 0.0
    while grown < max_expand - 0.5:
        grown += min(step, max_expand - grown)
        values.append(grown)
    return values


def _rect_growth_area(rect: fitz.Rect, base_rect: fitz.Rect) -> float:
    return max((rect.width * rect.height) - (base_rect.width * base_rect.height), 0.0)


def _choose_fit_candidate(
    candidates: list[dict[str, Any]],
    base_rect: fitz.Rect,
    preferred_scale: float,
) -> dict[str, Any]:
    if not candidates:
        return {"rect": fitz.Rect(base_rect), "spare": -1.0, "scale": 0.0}

    suitable = [
        c for c in candidates
        if c["spare"] >= 0 and c["scale"] >= preferred_scale
    ]
    if suitable:
        return min(
            suitable,
            key=lambda c: (
                _rect_growth_area(c["rect"], base_rect),
                -c["scale"],
                -c["spare"],
            ),
        )

    best_scale = max(c["scale"] for c in candidates)
    peers = [
        c for c in candidates
        if c["scale"] >= best_scale - _SCALE_EPSILON
    ]
    return min(
        peers,
        key=lambda c: (
            _rect_growth_area(c["rect"], base_rect),
            -c["spare"],
        ),
    )


def _choose_block_plan(
    plans: list[dict[str, Any]],
    base_rect: fitz.Rect,
    preferred_scale: float,
) -> dict[str, Any]:
    suitable = [
        p for p in plans
        if p["spare"] >= 0 and p["scale"] >= preferred_scale
    ]
    if suitable:
        return min(
            suitable,
            key=lambda p: (
                _rect_growth_area(p["rect"], base_rect),
                p["variant_priority"],
                -p["scale"],
                -p["spare"],
            ),
        )

    best_scale = max(p["scale"] for p in plans)
    peers = [
        p for p in plans
        if p["scale"] >= best_scale - _SCALE_EPSILON
    ]
    return min(
        peers,
        key=lambda p: (
            _rect_growth_area(p["rect"], base_rect),
            p["variant_priority"],
            -p["spare"],
        ),
    )


# ── fitting logic ─────────────────────────────────────────────────────────────

def _fit_block(
    html: str,
    css: str,
    rect: fitz.Rect,
    obstacles: list[fitz.Rect],
    siblings: list[fitz.Rect],
    page_height: float,
    page_width: float = 0.0,
    preferred_scale: float = _SCALE_THRESHOLD,
    rotate: int = 0,
) -> tuple[fitz.Rect, float, float]:
    # Return the best rect and probe metrics for this block's insert_htmlbox call
    if page_width <= 0:
        page_width = rect.x1 + 100

    probe_w = max(rect.width + _MAX_EXPAND_LEFT + _MAX_EXPAND_H + 20, 1)
    probe_h = max(rect.height + _MAX_EXPAND_UP + _MAX_EXPAND_DOWN + 20, 1)
    probe = fitz.open()
    probe.new_page(width=probe_w, height=probe_h)

    def _probe(r: fitz.Rect) -> tuple[float, float]:
        try:
            pg = probe[0]
            local = fitz.Rect(0, 0, r.width, r.height)
            res = pg.insert_htmlbox(local, html, css=css, scale_low=0, rotate=rotate)
            pg.clean_contents()
            return res[0], res[1]  # (spare_height, scale)
        except Exception:
            return -1.0, 0.0

    max_x1 = _free_x1(rect, page_width, obstacles, siblings)
    max_h_exp = min(_MAX_EXPAND_H, max_x1 - rect.x1)
    min_x0 = max(_free_x0(rect, obstacles, siblings), rect.x0 - _MAX_EXPAND_LEFT)
    min_y0 = max(_free_y0(rect, obstacles, siblings), rect.y0 - _MAX_EXPAND_UP)

    candidates: list[dict[str, Any]] = []
    seen: set[tuple[float, float, float, float]] = set()

    def _add_candidate(candidate: fitz.Rect) -> None:
        if candidate.is_empty or candidate.x1 <= candidate.x0 or candidate.y1 <= candidate.y0:
            return
        key = (
            round(candidate.x0, 3),
            round(candidate.y0, 3),
            round(candidate.x1, 3),
            round(candidate.y1, 3),
        )
        if key in seen:
            return
        seen.add(key)

        spare, scale = _probe(candidate)
        candidates.append({
            "rect": candidate,
            "spare": spare,
            "scale": scale,
        })

    for h_grown in _expansion_values(max_h_exp, _EXPAND_H_STEP):
        expanded = fitz.Rect(rect.x0, rect.y0, rect.x1 + h_grown, rect.y1)
        max_y1 = _free_y1(expanded, page_height, obstacles, siblings)
        max_v_exp = min(_MAX_EXPAND_DOWN, max_y1 - expanded.y1)

        for v_grown in _expansion_values(max_v_exp, _EXPAND_STEP):
            candidate = fitz.Rect(
                expanded.x0,
                expanded.y0,
                expanded.x1,
                expanded.y1 + v_grown,
            )
            _add_candidate(candidate)

    best = _choose_fit_candidate(candidates, rect, preferred_scale)
    if best["spare"] >= 0 and best["scale"] >= preferred_scale:
        probe.close()
        return fitz.Rect(best["rect"]), float(best["scale"]), float(best["spare"])

    if min_x0 < rect.x0:
        wide_seed = fitz.Rect(min_x0, rect.y0, rect.x1 + max_h_exp, rect.y1)
        wide_max_y1 = _free_y1(wide_seed, page_height, obstacles, siblings)
        _add_candidate(
            fitz.Rect(
                min_x0,
                rect.y0,
                rect.x1 + max_h_exp,
                wide_seed.y1 + min(_MAX_EXPAND_DOWN, wide_max_y1 - wide_seed.y1),
            )
        )

    if min_y0 < rect.y0:
        top_seed = fitz.Rect(rect.x0, min_y0, rect.x1, rect.y1)
        top_max_y1 = _free_y1(top_seed, page_height, obstacles, siblings)
        _add_candidate(
            fitz.Rect(
                rect.x0,
                min_y0,
                rect.x1,
                top_seed.y1 + min(_MAX_EXPAND_DOWN, top_max_y1 - top_seed.y1),
            )
        )

    if min_x0 < rect.x0 or min_y0 < rect.y0:
        cell_seed = fitz.Rect(min_x0, min_y0, rect.x1, rect.y1)
        cell_max_x1 = _free_x1(cell_seed, page_width, obstacles, siblings)
        cell_x1 = cell_seed.x1 + min(_MAX_EXPAND_H, cell_max_x1 - cell_seed.x1)
        cell_rect = fitz.Rect(min_x0, min_y0, cell_x1, rect.y1)
        cell_max_y1 = _free_y1(cell_rect, page_height, obstacles, siblings)
        _add_candidate(
            fitz.Rect(
                min_x0,
                min_y0,
                cell_x1,
                cell_rect.y1 + min(_MAX_EXPAND_DOWN, cell_max_y1 - cell_rect.y1),
            )
        )

    best = _choose_fit_candidate(candidates, rect, preferred_scale)
    probe.close()
    return fitz.Rect(best["rect"]), float(best["scale"]), float(best["spare"])


def _probe_htmlbox_fit(
    html: str,
    css: str,
    rect: fitz.Rect,
    *,
    rotate: int = 0,
) -> tuple[float, float]:
    probe = fitz.open()
    try:
        page = probe.new_page(width=max(rect.width + 2.0, 2.0), height=max(rect.height + 2.0, 2.0))
        result = page.insert_htmlbox(
            fitz.Rect(0, 0, max(rect.width, 1.0), max(rect.height, 1.0)),
            html,
            css=css,
            scale_low=0,
            rotate=rotate,
        )
        return float(result[1]), float(result[0])
    except Exception:
        return 0.0, -1.0
    finally:
        probe.close()


def _probe_literal_fit(
    translated_text: str,
    rect: fitz.Rect,
    line_styles: list[dict[str, Any]],
    *,
    rotate: int = 0,
) -> tuple[float, float]:
    html, css = _build_block_html(translated_text, line_styles)
    return _probe_htmlbox_fit(html, css, rect, rotate=rotate)


def _best_nonliteral_layout_plan(
    block: dict[str, Any],
    translated_text: str,
    orig_rect: fitz.Rect,
    line_styles: list[dict[str, Any]],
    obstacles: list[fitz.Rect],
    siblings: list[fitz.Rect],
    page_height: float,
    page_width: float,
    source_lang: str,
) -> dict[str, Any]:
    preferred_scale = _preferred_scale_for_block(block["text"], source_lang)
    rotation = int(block.get("rotation", 0))
    plans: list[dict[str, Any]] = []
    base_padding = float(block.get("padding_top_hint", 0.0) or 0.0)
    padding_options = [(base_padding, 0)]
    if base_padding > 1.0:
        padding_options.append((0.0, 2))

    for variant in _layout_variants_for_block(block["text"], source_lang):
        for padding_top, padding_priority in padding_options:
            html, css = _build_block_html(
                translated_text,
                line_styles,
                line_height=variant["line_height"],
                padding_top=padding_top,
            )
            fit_rect, probe_scale, probe_spare = _fit_block(
                html,
                css,
                orig_rect,
                obstacles,
                siblings,
                page_height,
                page_width=page_width,
                preferred_scale=preferred_scale,
                rotate=rotation,
            )
            plans.append({
                "html": html,
                "css": css,
                "rect": fit_rect,
                "scale": probe_scale,
                "spare": probe_spare,
                "variant_name": variant["name"],
                "variant_priority": variant["priority"] + padding_priority,
                "line_height": variant["line_height"],
                "padding_top": padding_top,
            })

    plan = _choose_block_plan(plans, orig_rect, preferred_scale)
    placement = _placement_overrides_for_block(
        orig_rect,
        plan["rect"],
        line_styles,
    )
    final_padding = max(
        float(plan.get("padding_top", 0.0) or 0.0),
        float(placement.get("padding_top", 0.0) or 0.0),
    )
    final_html, final_css = _build_block_html(
        translated_text,
        line_styles,
        line_height=plan["line_height"],
        align_override=placement.get("align"),
        padding_top=final_padding,
    )
    return {
        **plan,
        "html": final_html,
        "css": final_css,
        "padding_top": final_padding,
    }


def _probe_block_layout(
    block: dict[str, Any],
    translated_text: str,
    source_lang: str,
) -> tuple[float, float]:
    orig_rect = fitz.Rect(block["rect"])
    line_styles = block.get("line_styles", [])
    rotation = int(block.get("rotation", 0))
    if _strategy_for_block(block) == "literal":
        return _probe_literal_fit(
            translated_text,
            orig_rect,
            line_styles,
            rotate=rotation,
        )
    # Preflight should be cheap: probe the original container once using the
    # compact line-height variant. The final placement still runs the exhaustive
    # obstacle-aware expansion search. This avoids doubling a several-second
    # fit search for every at-risk block merely to decide whether to retry text.
    variants = _layout_variants_for_block(block["text"], source_lang)
    compact = variants[-1]
    html, css = _build_block_html(
        translated_text,
        line_styles,
        line_height=compact["line_height"],
        padding_top=float(block.get("padding_top_hint", 0.0) or 0.0),
    )
    return _probe_htmlbox_fit(html, css, orig_rect, rotate=rotation)


def _layout_fit_rank(scale: float, spare: float, preferred_scale: float) -> tuple[int, float, float]:
    fits = int(spare >= 0 and scale >= preferred_scale)
    return fits, float(scale), float(spare)


def _is_layout_retry_risk(block: dict[str, Any], translated_text: str) -> bool:
    capacity = max(int(block.get("capacity_chars") or 0), 0)
    if capacity <= 0:
        return False
    visible = len(_visible_block_text(translated_text))
    if visible <= 0:
        return False
    threshold = capacity * _LAYOUT_RETRY_RISK_RATIO
    if block.get("fixed_container"):
        threshold = capacity * 0.82
    return visible > threshold


def _layout_retry_unit_from_pdf_block(block: dict[str, Any]) -> SemanticTranslationUnit:
    base = _translation_unit_from_pdf_block(block)
    capacity = max(int(block.get("capacity_chars") or 0), 0)
    if capacity > 1:
        tightened_capacity = max(
            1,
            min(capacity - 1, int(capacity * _LAYOUT_RETRY_TARGET_RATIO)),
        )
    else:
        tightened_capacity = capacity or None
    return SemanticTranslationUnit(
        text=base.text,
        source_lang=base.source_lang,
        role=base.role,
        mode="concise",
        constraints=TranslationConstraints(
            max_chars=tightened_capacity,
            max_lines=base.constraints.max_lines,
            source_visible_chars=base.constraints.source_visible_chars,
            source_line_count=base.constraints.source_line_count,
        ),
        context=base.context,
        protected=base.protected,
        allow_unprotected_fallback=base.allow_unprotected_fallback,
        metadata={**dict(base.metadata), "layout_retry": "1"},
    )


def _supports_layout_guidance(translator: Any) -> bool:
    explicit = getattr(translator, "supports_layout_constraints", None)
    if explicit is not None:
        return bool(explicit)
    return any(
        callable(getattr(translator, name, None))
        for name in ("translate_units", "translate_blocks")
    )


def _retry_pdf_layout_overflow_blocks(
    page_blocks: list[tuple[int, list[dict[str, Any]]]],
    translator: Any,
    semantic_translator: Any | None,
    target_lang: str,
    *,
    source_lang: str = "auto",
    cancel_event: threading.Event | None = None,
    terminology: Mapping[str, str] | None = None,
    translation_memory: Any | None = None,
) -> dict[str, int]:
    stats = {"candidates": 0, "retried": 0, "accepted": 0}
    semantic_all = bool(getattr(semantic_translator, "_translate_all_blocks", False))
    groups: dict[str, list[dict[str, Any]]] = {"primary": [], "semantic": []}

    for _page_num, blocks in page_blocks:
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError("Translation cancelled")
        if not blocks:
            continue
        for block in blocks:
            translated = str(block.get("translated", block.get("text", "")))
            source = str(block.get("text", ""))
            if translated == source or not _is_layout_retry_risk(block, translated):
                continue
            preferred = _preferred_scale_for_block(source, source_lang)
            scale, spare = _probe_block_layout(
                block,
                translated,
                source_lang,
            )
            if spare >= 0 and scale >= preferred:
                continue
            stats["candidates"] += 1
            use_semantic = semantic_translator is not None and (
                _strategy_for_block(block) == "semantic" or semantic_all
            )
            group_name = "semantic" if use_semantic else "primary"
            active_translator = semantic_translator if use_semantic else translator
            if not _supports_layout_guidance(active_translator):
                continue
            groups[group_name].append({
                "block": block,
                "unit": _layout_retry_unit_from_pdf_block(block),
                "old_scale": scale,
                "old_spare": spare,
                "preferred": preferred,
            })

    for group_name, candidates in groups.items():
        if not candidates:
            continue
        active_translator = semantic_translator if group_name == "semantic" else translator
        fallback = translator if group_name == "semantic" else None
        service = TranslationService(
            active_translator,
            fallback_translator=fallback,
            terminology=terminology,
            translation_memory=translation_memory,
        )
        batch = service.translate_units(
            [item["unit"] for item in candidates],
            target_lang,
            cancel_event=cancel_event,
        )
        stats["retried"] += len(candidates)
        for item, candidate, engine in zip(candidates, batch.texts, batch.engines):
            if candidate is None:
                continue
            block = item["block"]
            new_scale, new_spare = _probe_block_layout(
                block,
                str(candidate),
                source_lang,
            )
            old_rank = _layout_fit_rank(item["old_scale"], item["old_spare"], item["preferred"])
            new_rank = _layout_fit_rank(new_scale, new_spare, item["preferred"])
            improved = new_rank[0] > old_rank[0] or (
                new_rank[0] == old_rank[0]
                and (
                    new_rank[1] > old_rank[1] + 0.02
                    or (
                        abs(new_rank[1] - old_rank[1]) <= 0.02
                        and new_rank[2] > old_rank[2] + 1.0
                    )
                )
            )
            if improved:
                block["translated"] = str(candidate)
                if engine:
                    block["translation_engine"] = engine
                stats["accepted"] += 1

    return stats


def _strategy_for_block(block: dict[str, Any]) -> str:
    strategy = str(block.get("strategy", "free")).strip().lower()
    if strategy in {"literal", "semantic", "free"}:
        return strategy
    return "free"


def _visible_block_text(text: str) -> str:
    return " ".join(str(text or "").split())


def _block_line_count(text: str) -> int:
    return len([line for line in str(text or "").splitlines() if line.strip()]) or 1


def _looks_numeric_or_tabular_text(text: str) -> bool:
    visible_chars = [ch for ch in str(text or "") if not ch.isspace()]
    if len(visible_chars) < 4:
        return False

    numericish = sum(
        1
        for ch in visible_chars
        if ch.isdigit() or ch in "%$€£¥:：/.,-+()[]{}<>|"
    )
    return any(ch.isdigit() for ch in visible_chars) and (numericish / len(visible_chars)) >= 0.55


def _looks_label_like_text(text: str) -> bool:
    visible = _visible_block_text(text)
    if not visible or len(visible) > _HEURISTIC_LABEL_MAX_CHARS:
        return False
    if _block_line_count(text) > 2:
        return False
    if visible.endswith((":", "：")):
        return True
    if any(mark in visible for mark in (":", "：")) and len(visible) <= 32:
        return True
    if len(visible.split()) <= 6 and re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9 /&+().,-]{0,41}", visible):
        return True
    return False


def _heuristic_strategy_for_block(
    block: dict[str, Any],
    *,
    source_lang: str,
    capacity_chars: int,
    content_hint: str,
) -> tuple[str, str] | None:
    text = str(block.get("text", ""))
    visible = _visible_block_text(text)
    line_count = _block_line_count(text)

    if not visible:
        return "free", "heuristic-empty"
    if content_hint == "heading":
        return "literal", "heuristic-heading"
    if line_count <= 2 and len(visible) <= _HEURISTIC_LITERAL_MAX_CHARS:
        return "literal", "heuristic-short"
    if _looks_numeric_or_tabular_text(text):
        return "literal", "heuristic-numeric"
    if _looks_label_like_text(text):
        return "literal", "heuristic-label"
    if content_hint != "body" or capacity_chars <= 0:
        return None

    visible_len = len(visible)
    if _is_cjk_source_block(text, source_lang):
        if line_count >= 2 and visible_len >= 16 and visible_len <= capacity_chars * _HEURISTIC_ROOMY_CJK_FREE_RATIO:
            return "free", "heuristic-roomy-cjk"
        return None

    if visible_len >= 24 and visible_len <= capacity_chars * _HEURISTIC_ROOMY_FREE_RATIO:
        return "free", "heuristic-roomy"
    return None


def _preroute_obvious_blocks(
    page_blocks: list[tuple[int, list[dict[str, Any]]]],
    source_lang: str,
) -> list[tuple[int, list[dict[str, Any]]]]:
    heuristic_advisor = NullAdvisor()

    for _, blocks in page_blocks:
        page_font_baseline = heuristic_advisor._page_font_baseline(blocks)
        for block in blocks:
            preset_reason = str(block.get("advisor_reason", "")).strip()
            if preset_reason and _strategy_for_block(block) in {"literal", "semantic", "free"}:
                continue

            capacity_chars = heuristic_advisor.estimate_capacity_chars(block)
            content_hint = heuristic_advisor.infer_content_hint(block, page_font_baseline)
            decision = _heuristic_strategy_for_block(
                block,
                source_lang=source_lang,
                capacity_chars=capacity_chars,
                content_hint=content_hint,
            )
            if decision is None:
                continue

            strategy, reason = decision
            block["capacity_chars"] = capacity_chars
            block["content_hint"] = content_hint
            block["strategy"] = strategy
            block["advisor_reason"] = reason

    return page_blocks


def _engine_name_for_translator(translator: Any) -> str:
    engine_name = getattr(translator, "_engine_name", "")
    if isinstance(engine_name, str) and engine_name.strip():
        return engine_name
    return translator.__class__.__name__.lower()


def _normalize_translation_results(
    texts: list[str],
    translated: object | None,
) -> tuple[list[str], list[int]]:
    results: list[str] = []
    failed_indices: list[int] = []
    items = translated if isinstance(translated, (list, tuple)) else []

    for idx, text in enumerate(texts):
        item = items[idx] if idx < len(items) else None
        if item is None:
            results.append(text)
            failed_indices.append(idx)
        else:
            results.append(str(item))

    return results, failed_indices


def _translate_units_with_fallback(
    translator: Any,
    texts: list[str],
    target_lang: str,
    *,
    source_lang: str = "auto",
    cancel_event: threading.Event | None = None,
    log_prefix: str = "Batch translation",
    terminology: Mapping[str, str] | None = None,
    translation_memory: Any | None = None,
    progress_callback: Callable[[int, int], None] | None = None,
) -> tuple[list[str], list[int]]:
    if not texts:
        return [], []

    service = TranslationService(translator, terminology=terminology, translation_memory=translation_memory)
    batch = service.translate_units(
        [
            SemanticTranslationUnit(text=text, source_lang=source_lang, mode="natural")
            for text in texts
        ],
        target_lang,
        cancel_event=cancel_event,
        progress_callback=progress_callback,
    )
    failed = batch.failed_indices
    results = [
        source if translated is None else str(translated)
        for source, translated in zip(texts, batch.texts)
    ]
    if failed:
        logger.warning("%s left %d item(s) untranslated", log_prefix, len(failed))
    return results, failed


def _translation_unit_from_pdf_block(block: dict[str, Any]) -> SemanticTranslationUnit:
    text = str(block.get("text", ""))
    capacity_chars = max(int(block.get("capacity_chars") or 0), 0)
    content_hint = str(block.get("content_hint", "body") or "body")
    strategy = _strategy_for_block(block)
    return SemanticTranslationUnit(
        text=text,
        source_lang=str(block.get("source_lang", "auto") or "auto"),
        role=content_hint,
        mode="concise" if strategy == "semantic" else "natural",
        constraints=TranslationConstraints(
            max_chars=capacity_chars or None,
            source_visible_chars=len(" ".join(text.split())),
            source_line_count=max(1, len(text.split("\n"))),
        ),
        metadata={"content_hint": content_hint, "strategy": strategy},
    )


def _translate_blocks_with_fallback(
    translator: Any,
    blocks: list[dict[str, Any]],
    target_lang: str,
    *,
    cancel_event: threading.Event | None = None,
    log_prefix: str = "Semantic translation",
    terminology: Mapping[str, str] | None = None,
    translation_memory: Any | None = None,
) -> tuple[list[str], list[int]]:
    if not blocks:
        return [], []

    units = [_translation_unit_from_pdf_block(block) for block in blocks]
    service = TranslationService(translator, terminology=terminology, translation_memory=translation_memory)
    batch = service.translate_units(units, target_lang, cancel_event=cancel_event)
    failed = batch.failed_indices
    results = [
        unit.text if translated is None else str(translated)
        for unit, translated in zip(units, batch.texts)
    ]
    if failed:
        logger.warning("%s left %d item(s) untranslated", log_prefix, len(failed))
    return results, failed


def _classify_blocks(
    page_blocks: list[tuple[int, list[dict[str, Any]]]],
    advisor: Any,
    source_lang: str,
    target_lang: str,
    *,
    cancel_event: threading.Event | None = None,
) -> list[tuple[int, list[dict[str, Any]]]]:
    active_advisor = advisor or NullAdvisor()
    return active_advisor.classify_blocks(
        page_blocks,
        source_lang,
        target_lang,
        cancel_event=cancel_event,
    )


def _insert_literal_block(
    page: fitz.Page,
    translated_text: str,
    orig_rect: fitz.Rect,
    line_styles: list[dict[str, Any]],
    *,
    rotate: int = 0,
) -> tuple[float, float]:
    html, css = _build_block_html(translated_text, line_styles)
    return page.insert_htmlbox(orig_rect, html, css=css, scale_low=0, rotate=rotate)


# ── main entry point ─────────────────────────────────────────────────────────

def _translate_pdf_open_document(
    src: Any,
    input_path: str,
    output_path: str,
    translator: Any,
    target_lang: str,
    *,
    cancel_event: threading.Event | None = None,
    source_lang: str = "auto",
    progress_callback: Callable[[int, int], None] | None = None,
    progress_event_callback: Callable[[ProgressUpdate], None] | None = None,
    advisor: Any | None = None,
    semantic_translator: Any | None = None,
    strict_errors: bool = False,
    terminology: Mapping[str, str] | None = None,
    translation_memory: Any | None = None,
) -> str | None:
    # Translate an already-open PDF while preserving the original layout.

    if _is_ocr_required(src):
        logger.warning(
            "Skipping '%s': scanned/image PDF (OCR required).",
            Path(input_path).name,
        )
        return "skipped-ocr"

    progress = ProgressReporter(progress_event_callback)
    progress.update("analyzing", 0, max(src.page_count + 1, 1))

    errors = 0

    # Progress: num_pages (extract) + 1 (classify) + 1 (translate) + num_pages (redact)
    _n_pages = src.page_count
    _total_steps = _n_pages + 2 + _n_pages
    _steps_done = 0

    def _report(done: int) -> None:
        nonlocal _steps_done
        _steps_done = done
        if progress_callback is not None:
            progress_callback(_steps_done, _total_steps)

    # ── Phase 1: extract ──────────────────────────────────────────────────────
    page_blocks: list[tuple[int, list[dict]]] = []
    for page_num in range(src.page_count):
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError("Translation cancelled")

        page   = src[page_num]
        blocks = _extract_blocks(page)

        pdi = len(page_blocks)
        page_blocks.append((page_num, blocks))

        _report(page_num + 1)
        progress.update("analyzing", page_num + 1, _n_pages + 1)

    classify_started = time.perf_counter()
    page_blocks = _classify_blocks(
        _preroute_obvious_blocks(page_blocks, source_lang),
        advisor,
        source_lang,
        target_lang,
        cancel_event=cancel_event,
    )
    classify_elapsed = time.perf_counter() - classify_started
    logger.info("PDF semantic classification finished in %.2fs", classify_elapsed)
    _report(_n_pages + 1)
    progress.update("analyzing", _n_pages + 1, _n_pages + 1)

    # ── Phase 2: translate ────────────────────────────────────────────────────
    primary_entries: list[tuple[int, int]] = []
    primary_units: list[str] = []
    semantic_entries: list[tuple[int, int]] = []
    semantic_blocks: list[dict[str, Any]] = []

    _semantic_all = getattr(semantic_translator, "_translate_all_blocks", False)
    for pdi, (_, blocks) in enumerate(page_blocks):
        for bi, block in enumerate(blocks):
            block["source_lang"] = source_lang
            strategy = _strategy_for_block(block)
            if semantic_translator is not None and (strategy == "semantic" or _semantic_all):
                semantic_entries.append((pdi, bi))
                semantic_blocks.append(block)
            else:
                primary_entries.append((pdi, bi))
                primary_units.append(block["text"])

    translation_total = len(primary_entries) + len(semantic_entries)
    progress.update(
        "translating",
        0,
        max(translation_total, 1),
        detail=f"{translation_total} text block(s)",
    )

    primary_results, primary_failures = _translate_units_with_fallback(
        translator,
        primary_units,
        target_lang,
        source_lang=source_lang,
        cancel_event=cancel_event,
        log_prefix="Primary PDF translation",
        terminology=terminology,
        translation_memory=translation_memory,
        progress_callback=(
            (lambda done, total: progress.update("translating", done, max(translation_total, 1)))
            if primary_entries
            else None
        ),
    )
    errors += len(primary_failures)
    primary_engine = _engine_name_for_translator(translator)
    for idx, (pdi, bi) in enumerate(primary_entries):
        blocks = page_blocks[pdi][1]
        blocks[bi]["translated"] = primary_results[idx]
        blocks[bi]["translation_engine"] = primary_engine

    semantic_started = time.perf_counter()
    if semantic_entries:
        semantic_units = [_translation_unit_from_pdf_block(block) for block in semantic_blocks]
        semantic_service = TranslationService(
            semantic_translator,
            fallback_translator=translator,
            terminology=terminology,
            translation_memory=translation_memory,
        )
        semantic_batch = semantic_service.translate_units(
            semantic_units,
            target_lang,
            cancel_event=cancel_event,
            progress_callback=lambda done, total: progress.update(
                "translating",
                len(primary_entries) + done,
                max(translation_total, 1),
            ),
        )
        errors += len(semantic_batch.failed_indices)

        for idx, (pdi, bi) in enumerate(semantic_entries):
            blocks = page_blocks[pdi][1]
            translated = semantic_batch.texts[idx]
            if translated is None:
                translated = semantic_blocks[idx]["text"]
            blocks[bi]["translated"] = translated
            blocks[bi]["translation_engine"] = (
                semantic_batch.engines[idx] or primary_engine
            )
    semantic_elapsed = time.perf_counter() - semantic_started
    logger.info(
        "PDF semantic translation finished in %.2fs for %s blocks",
        semantic_elapsed,
        len(semantic_entries),
    )
    progress.update("translating", max(translation_total, 1), max(translation_total, 1))

    layout_retry_stats = _retry_pdf_layout_overflow_blocks(
        page_blocks,
        translator,
        semantic_translator,
        target_lang,
        source_lang=source_lang,
        cancel_event=cancel_event,
        terminology=terminology,
        translation_memory=translation_memory,
    )
    if layout_retry_stats["candidates"]:
        logger.info(
            "PDF layout preflight found %d at-risk block(s), retried %d and accepted %d improved fit(s)",
            layout_retry_stats["candidates"],
            layout_retry_stats["retried"],
            layout_retry_stats["accepted"],
        )

    _report(_n_pages + 2)  # extraction + classification + translation done
    progress.update("layout", 0, max(_n_pages, 1))

    # ── Phase 3: redact + insert ──────────────────────────────────────────────
    for pdi, (page_num, blocks) in enumerate(page_blocks):
        if not blocks:
            _report(_n_pages + 2 + pdi + 1)
            progress.update("layout", pdi + 1, max(_n_pages, 1))
            continue

        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError("Translation cancelled")

        page        = src[page_num]
        page_height = page.rect.height
        page_width  = page.rect.width

        # Collect obstacles before redaction (image/drawing positions are stable)
        obstacles = _collect_obstacles(page)

        all_rects    = [fitz.Rect(b["rect"]) for b in blocks]
        placed_rects: list[fitz.Rect | None] = [None] * len(blocks)

        # Redact all source text in one pass *before* inserting any translations.
        # Repeated apply_redactions calls can touch content inserted earlier.
        # Explicit graphics=0 is important: PyMuPDF otherwise removes vector
        # graphics fully contained by a redaction rectangle (table borders,
        # underlines, small shapes). Images and graphics are preservation targets.
        changed_blocks = [
            block for block in blocks
            if str(block.get("translated", block["text"])) != str(block["text"])
        ]
        for block in changed_blocks:
            line_rects = block.get("line_rects", [])
            if line_rects:
                for lr in line_rects:
                    page.add_redact_annot(lr, fill=None)  # type: ignore[arg-type]
            else:
                page.add_redact_annot(block["rect"], fill=None)  # type: ignore[arg-type]
        if changed_blocks:
            page.apply_redactions(images=0, graphics=0, text=0)  # type: ignore[arg-type]

        # Process each block individually. If insertion fails, re-insert the
        # untranslated text so the area is not left blank.
        for bi, block in enumerate(blocks):
            tr_text     = block.get("translated", block["text"])
            orig_rect   = block["rect"]
            line_styles = block["line_styles"]
            line_rects  = block.get("line_rects", [])
            # Use the actual placed rect for each sibling when available so
            # expanded blocks are treated as obstacles by subsequent blocks.
            siblings    = [
                placed_rects[i] if placed_rects[i] is not None else all_rects[i]
                for i in range(len(all_rects)) if i != bi
            ]
            preferred_scale = _preferred_scale_for_block(block["text"], source_lang)
            strategy = _strategy_for_block(block)
            rotation = int(block.get("rotation", 0))

            # Pass-through / already-target-language blocks stay byte-for-byte in
            # the original page content stream. Re-rendering unchanged text is a
            # needless source of font, spacing and baseline drift.
            if str(tr_text) == str(block["text"]):
                placed_rects[bi] = orig_rect
                continue

            try:
                if strategy == "literal":
                    result = _insert_literal_block(
                        page,
                        tr_text,
                        orig_rect,
                        line_styles,
                        rotate=rotation,
                    )
                    placed_rects[bi] = orig_rect
                    if result[0] < 0:
                        logger.debug(
                            "literal insert_htmlbox overflow page %d block %d rect=%s",
                            page_num + 1, bi, orig_rect,
                        )
                    elif result[1] < preferred_scale:
                        logger.debug(
                            "literal insert_htmlbox compressed page %d block %d scale=%.3f preferred=%.3f rect=%s",
                            page_num + 1, bi, result[1], preferred_scale, orig_rect,
                        )
                    logger.debug(
                        "placed page=%d bi=%d strategy=literal engine=%s src_len=%d tr_len=%d cap=%d rect=%s",
                        page_num + 1, bi,
                        block.get("translation_engine", "-"),
                        len(block["text"]), len(tr_text),
                        block.get("capacity_chars", 0),
                        orig_rect,
                    )
                else:
                    plan = _best_nonliteral_layout_plan(
                        block,
                        tr_text,
                        orig_rect,
                        line_styles,
                        obstacles,
                        siblings,
                        page_height,
                        page_width,
                        source_lang,
                    )
                    result = page.insert_htmlbox(
                        plan["rect"], plan["html"], css=plan["css"], scale_low=0,
                        rotate=rotation,
                    )
                    placed_rects[bi] = plan["rect"]
                    if result[0] < 0:
                        logger.debug(
                            "insert_htmlbox overflow page %d block %d rect=%s",
                            page_num + 1, bi, plan["rect"],
                        )
                    elif result[1] < preferred_scale:
                        logger.debug(
                            "insert_htmlbox compressed page %d block %d variant=%s scale=%.3f probe=%.3f spare=%.2f preferred=%.3f rect=%s",
                            page_num + 1, bi, plan["variant_name"], result[1],
                            plan["scale"], plan["spare"], preferred_scale,
                            plan["rect"],
                        )
                    logger.debug(
                        "placed page=%d bi=%d strategy=%s engine=%s src_len=%d tr_len=%d cap=%d orig=%s placed=%s",
                        page_num + 1, bi, strategy,
                        block.get("translation_engine", "-"),
                        len(block["text"]), len(tr_text),
                        block.get("capacity_chars", 0),
                        orig_rect, plan["rect"],
                    )
            except Exception:
                logger.exception(
                    "Failed inserting translated block page %d block %d",
                    page_num + 1, bi,
                )
                # Fallback: re-insert original text so the area isn't blank
                placed_rects[bi] = orig_rect
                try:
                    fb_html, fb_css = _build_block_html(
                        block["text"], line_styles
                    )
                    page.insert_htmlbox(
                        orig_rect, fb_html, css=fb_css, scale_low=0,
                        rotate=rotation,
                    )
                except Exception:
                    logger.debug(
                        "Fallback insertion also failed page %d block %d",
                        page_num + 1, bi,
                    )
                errors += 1

        _report(_n_pages + 2 + pdi + 1)  # extract + classify + translate + pages redacted so far
        progress.update("layout", pdi + 1, max(_n_pages, 1))

    # ── Save ──────────────────────────────────────────────────────────────────
    if cancel_event is not None and cancel_event.is_set():
        raise CancelledError("Translation cancelled before saving")

    progress.update("saving", 0, 1)
    src.save(str(output_path), garbage=4, deflate=True, clean=True)
    progress.complete()

    if errors:
        message = (
            f"PDF translation completed with {errors} issue(s); "
            "affected blocks were restored from the original where possible"
        )
        if strict_errors:
            raise RuntimeError(message)
        logger.warning(message)
    return None

def translate_pdf(
    input_path: str,
    output_path: str,
    translator: Any,
    target_lang: str,
    *,
    cancel_event: threading.Event | None = None,
    source_lang: str = "auto",
    progress_callback: Callable[[int, int], None] | None = None,
    progress_event_callback: Callable[[ProgressUpdate], None] | None = None,
    advisor: Any | None = None,
    semantic_translator: Any | None = None,
    strict_errors: bool = False,
    terminology: Mapping[str, str] | None = None,
    translation_memory: Any | None = None,
) -> str | None:
    """Open, translate, and always close a PDF document."""
    src = fitz.open(input_path)
    try:
        return _translate_pdf_open_document(
            src, input_path, output_path, translator, target_lang,
            cancel_event=cancel_event, source_lang=source_lang,
            progress_callback=progress_callback,
            progress_event_callback=progress_event_callback,
            advisor=advisor,
            semantic_translator=semantic_translator, strict_errors=strict_errors,
            terminology=terminology, translation_memory=translation_memory,
        )
    finally:
        try:
            if not bool(getattr(src, "is_closed", False)):
                src.close()
        except Exception:
            logger.debug("Failed closing PDF document", exc_info=True)
