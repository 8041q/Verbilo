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
from typing import Any, Callable

import fitz  # PyMuPDF >= 1.24
from ..advisors import NullAdvisor
from ..utils import CancelledError

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

def _extract_blocks(page: fitz.Page) -> list[dict]:
    """Return block-info dicts for all text blocks on *page*.

    Each dict:
        rect        – fitz.Rect of the original block bounding box
        text        – full plain text (lines joined by \\n) for translation
        line_styles – list of style-dicts, one per original line, in order
        line_rects  – list of fitz.Rect, one per original line (for per-line redaction)
        page_width  – page width
        page_height – page height
    """
    page_w = page.rect.width
    page_h = page.rect.height
    d = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)

    out: list[dict] = []
    for block in d.get("blocks", []):
        if block.get("type") != 0:
            continue

        lines_data: list[tuple[str, dict, fitz.Rect]] = []

        for line in block.get("lines", []):
            spans = line.get("spans", [])
            valid = [s for s in spans if s.get("text", "").strip()]
            if not valid:
                continue

            # Join spans preserving inter-span spacing.  When the horizontal
            # gap between consecutive spans exceeds half the average character
            # width we insert a space so "100  50" does not become "10050".
            parts: list[str] = [valid[0]["text"]]
            for k in range(1, len(valid)):
                prev_bbox = fitz.Rect(valid[k - 1]["bbox"])
                cur_bbox  = fitz.Rect(valid[k]["bbox"])
                gap = cur_bbox.x0 - prev_bbox.x1
                avg_cw = prev_bbox.width / max(len(valid[k - 1]["text"]), 1)
                if gap > avg_cw * 0.5:
                    parts.append(" ")
                parts.append(valid[k]["text"])
            line_text = "".join(parts).strip()
            if not line_text:
                continue

            # Majority-vote font weight: use the weight that covers the most
            # characters.  Ties go to the lighter weight.
            weight_chars: dict[int, int] = {}
            for s in valid:
                w = _css_weight(s["font"], s["flags"])
                weight_chars[w] = weight_chars.get(w, 0) + len(s["text"])
            final_weight = min(
                weight_chars,
                key=lambda w: (-weight_chars[w], w),  # most chars wins; lighter on tie
            )

            dominant = max(valid, key=lambda s: len(s["text"]))
            lb = fitz.Rect(line["bbox"])

            lines_data.append((line_text, {
                "size":   dominant["size"],
                "weight": final_weight,
                "style":  _css_style(dominant["flags"]),
                "color":  _css_color(dominant["color"]),
                "align":  "left",  # placeholder — computed at block level below
            }, lb))

        if not lines_data:
            continue

        # ── Block-relative alignment inference ──
        # For multi-line blocks (≥3 lines): check variance of x0 and line
        # centres within the block.  For 1-2 line blocks: default to "left"
        # unless the block is clearly centred on the page.
        block_rect = fitz.Rect(block["bbox"])
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
            # 1-2 lines: default left, tight centering check
            bcx = (block_rect.x0 + block_rect.x1) / 2
            block_align = "left"
            if (block_rect.width < page_w * 0.50
                    and abs(bcx - page_w / 2) < page_w * 0.03):
                block_align = "center"

        for _, style, _ in lines_data:
            style["align"] = block_align

        out.append({
            "rect":        block_rect,
            "text":        "\n".join(t for t, _, _ in lines_data),
            "line_styles": [s for _, s, _ in lines_data],
            "line_rects":  [lr for _, _, lr in lines_data],
            "page_width":  page_w,
            "page_height": page_h,
        })

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
            res = pg.insert_htmlbox(local, html, css=css, scale_low=0)
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
    cancel_event: threading.Event | None = None,
    log_prefix: str = "Batch translation",
) -> tuple[list[str], list[int]]:
    if not texts:
        return [], []

    failed_indices: list[int]
    try:
        batch_results = translator.translate_batch(
            texts,
            target_lang,
            cancel_event=cancel_event,
        )
    except CancelledError:
        raise
    except Exception:
        logger.exception("%s failed; falling back to per-item", log_prefix)
        batch_results = None

    results, failed_indices = _normalize_translation_results(texts, batch_results)
    if not failed_indices:
        return results, []

    remaining_failures: list[int] = []
    for idx in failed_indices:
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError("Translation cancelled")
        try:
            item = translator.translate_text(texts[idx], target_lang)
        except CancelledError:
            raise
        except Exception:
            logger.exception("%s per-item fallback failed", log_prefix)
            remaining_failures.append(idx)
            results[idx] = texts[idx]
            continue

        if item is None:
            remaining_failures.append(idx)
            results[idx] = texts[idx]
        else:
            results[idx] = item

    return results, remaining_failures


def _translate_blocks_with_fallback(
    translator: Any,
    blocks: list[dict[str, Any]],
    target_lang: str,
    *,
    cancel_event: threading.Event | None = None,
    log_prefix: str = "Semantic translation",
) -> tuple[list[str], list[int]]:
    if not blocks:
        return [], []

    translate_blocks = getattr(translator, "translate_blocks", None)
    texts = [str(block.get("text", "")) for block in blocks]

    if callable(translate_blocks):
        try:
            batch_results = translate_blocks(
                blocks,
                target_lang,
                cancel_event=cancel_event,
            )
        except CancelledError:
            raise
        except Exception:
            logger.exception("%s failed; falling back to per-item", log_prefix)
        else:
            results, failed_indices = _normalize_translation_results(texts, batch_results)
            if not failed_indices:
                return results, []

            remaining_failures: list[int] = []
            for idx in failed_indices:
                if cancel_event is not None and cancel_event.is_set():
                    raise CancelledError("Translation cancelled")
                try:
                    item = translator.translate_text(texts[idx], target_lang)
                except CancelledError:
                    raise
                except Exception:
                    logger.exception("%s per-item fallback failed", log_prefix)
                    remaining_failures.append(idx)
                    results[idx] = texts[idx]
                    continue

                if item is None:
                    remaining_failures.append(idx)
                    results[idx] = texts[idx]
                else:
                    results[idx] = item

            return results, remaining_failures

    return _translate_units_with_fallback(
        translator,
        texts,
        target_lang,
        cancel_event=cancel_event,
        log_prefix=log_prefix,
    )


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
) -> tuple[float, float]:
    html, css = _build_block_html(translated_text, line_styles)
    return page.insert_htmlbox(orig_rect, html, css=css, scale_low=0)


# ── main entry point ─────────────────────────────────────────────────────────

def translate_pdf(
    input_path: str,
    output_path: str,
    translator: Any,
    target_lang: str,
    *,
    cancel_event: threading.Event | None = None,
    source_lang: str = "auto",
    progress_callback: Callable[[int, int], None] | None = None,
    advisor: Any | None = None,
    semantic_translator: Any | None = None,
) -> str | None:
    # Translate a PDF in-place while preserving the original layout
    src = fitz.open(input_path)

    if _is_ocr_required(src):
        logger.warning(
            "Skipping '%s': scanned/image PDF (OCR required).",
            Path(input_path).name,
        )
        src.close()
        return "skipped-ocr"

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
            src.close()
            raise CancelledError("Translation cancelled")

        page   = src[page_num]
        blocks = _extract_blocks(page)

        pdi = len(page_blocks)
        page_blocks.append((page_num, blocks))

        _report(page_num + 1)

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

    primary_results, primary_failures = _translate_units_with_fallback(
        translator,
        primary_units,
        target_lang,
        cancel_event=cancel_event,
        log_prefix="Primary PDF translation",
    )
    errors += len(primary_failures)
    primary_engine = _engine_name_for_translator(translator)
    for idx, (pdi, bi) in enumerate(primary_entries):
        blocks = page_blocks[pdi][1]
        blocks[bi]["translated"] = primary_results[idx]
        blocks[bi]["translation_engine"] = primary_engine

    semantic_started = time.perf_counter()
    if semantic_entries:
        semantic_results, semantic_failures = _translate_blocks_with_fallback(
            semantic_translator,
            semantic_blocks,
            target_lang,
            cancel_event=cancel_event,
            log_prefix="Semantic PDF translation",
        )
        semantic_engine = _engine_name_for_translator(semantic_translator)
        semantic_engines = [semantic_engine] * len(semantic_entries)

        if semantic_failures:
            fallback_units = [semantic_blocks[idx]["text"] for idx in semantic_failures]
            fallback_results, fallback_failures = _translate_units_with_fallback(
                translator,
                fallback_units,
                target_lang,
                cancel_event=cancel_event,
                log_prefix="Semantic fallback translation",
            )
            for local_idx, fallback_text in zip(semantic_failures, fallback_results):
                semantic_results[local_idx] = fallback_text
                semantic_engines[local_idx] = primary_engine
            errors += len(fallback_failures)

        for idx, (pdi, bi) in enumerate(semantic_entries):
            blocks = page_blocks[pdi][1]
            blocks[bi]["translated"] = semantic_results[idx]
            blocks[bi]["translation_engine"] = semantic_engines[idx]
    semantic_elapsed = time.perf_counter() - semantic_started
    logger.info(
        "PDF semantic translation finished in %.2fs for %s blocks",
        semantic_elapsed,
        len(semantic_entries),
    )

    _report(_n_pages + 2)  # extraction + classification + translation done

    # ── Phase 3: redact + insert ──────────────────────────────────────────────
    for pdi, (page_num, blocks) in enumerate(page_blocks):
        if not blocks:
            continue

        if cancel_event is not None and cancel_event.is_set():
            src.close()
            raise CancelledError("Translation cancelled")

        page        = src[page_num]
        page_height = page.rect.height
        page_width  = page.rect.width

        # Collect obstacles before redaction (image/drawing positions are stable)
        obstacles = _collect_obstacles(page)

        all_rects    = [fitz.Rect(b["rect"]) for b in blocks]
        placed_rects: list[fitz.Rect | None] = [None] * len(blocks)

        # Process each block individually: redact then insert.  If insertion
        # fails the original text is already gone so we attempt to re-insert
        # the untranslated text as a fallback.
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

            # Redact per-line rects instead of the whole block rect so that
            # vector drawings (table lines, shapes) within the block area
            # but outside the actual text lines are preserved.
            if line_rects:
                for lr in line_rects:
                    page.add_redact_annot(lr, fill=None)  # type: ignore[arg-type]
            else:
                page.add_redact_annot(orig_rect, fill=None)  # type: ignore[arg-type]
            page.apply_redactions(images=0)  # type: ignore[arg-type]

            try:
                if strategy == "literal":
                    result = _insert_literal_block(
                        page,
                        tr_text,
                        orig_rect,
                        line_styles,
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
                        block.get("engine", "-"),
                        len(block["text"]), len(tr_text),
                        block.get("capacity_chars", 0),
                        orig_rect,
                    )
                else:
                    plans: list[dict[str, Any]] = []
                    for variant in _layout_variants_for_block(block["text"], source_lang):
                        html, css = _build_block_html(
                            tr_text,
                            line_styles,
                            line_height=variant["line_height"],
                        )
                        fit_rect, probe_scale, probe_spare = _fit_block(
                            html, css, orig_rect,
                            obstacles, siblings, page_height,
                            page_width=page_width,
                            preferred_scale=preferred_scale,
                        )
                        plans.append({
                            "html": html,
                            "css": css,
                            "rect": fit_rect,
                            "scale": probe_scale,
                            "spare": probe_spare,
                            "variant_name": variant["name"],
                            "variant_priority": variant["priority"],
                            "line_height": variant["line_height"],
                        })

                    plan = _choose_block_plan(plans, orig_rect, preferred_scale)
                    placement = _placement_overrides_for_block(
                        orig_rect,
                        plan["rect"],
                        line_styles,
                    )
                    final_html, final_css = _build_block_html(
                        tr_text,
                        line_styles,
                        line_height=plan["line_height"],
                        align_override=placement.get("align"),
                        padding_top=placement.get("padding_top", 0.0),
                    )
                    result = page.insert_htmlbox(
                        plan["rect"], final_html, css=final_css, scale_low=0
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
                        block.get("engine", "-"),
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
                        orig_rect, fb_html, css=fb_css, scale_low=0
                    )
                except Exception:
                    logger.debug(
                        "Fallback insertion also failed page %d block %d",
                        page_num + 1, bi,
                    )
                errors += 1

        _report(_n_pages + 2 + pdi + 1)  # extract + classify + translate + pages redacted so far

    # ── Save ──────────────────────────────────────────────────────────────────
    if cancel_event is not None and cancel_event.is_set():
        src.close()
        raise CancelledError("Translation cancelled before saving")

    src.save(str(output_path), garbage=4, deflate=True, clean=True)
    src.close()

    if errors:
        raise RuntimeError(f"PDF translation completed with {errors} issue(s)")
    return None