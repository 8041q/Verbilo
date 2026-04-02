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
from pathlib import Path
from typing import Any

import fitz  # PyMuPDF >= 1.24
from ..utils import CancelledError

logger = logging.getLogger(__name__)

# ── tuneable constants ──────────────────────────────────────────────────────
_MIN_CHARS_PER_PAGE  = 20    # avg chars/page below which PDF is treated as scanned
_SCALE_THRESHOLD     = 0.78  # if insert_htmlbox scale < this, try rect expansion
_MAX_EXPAND_DOWN     = 60.0  # max points a block rect may grow downward
_EXPAND_STEP         = 6.0   # vertical expansion increment (points)
_OBSTACLE_MIN_AREA   = 200   # filled drawings smaller than this (pt²) are ignored
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

def _build_block_html(translated: str, line_styles: list[dict]) -> tuple[str, str]:
    # Build HTML + CSS for *translated* text using the original per-line styles.
    orig_lines = translated.split("\n")
    n = len(line_styles)
    max_size = max((s["size"] for s in line_styles), default=11.0)

    # Use the dominant alignment (most common across lines)
    align_counts: dict[str, int] = {}
    for s in line_styles:
        align_counts[s["align"]] = align_counts.get(s["align"], 0) + 1
    dominant_align = max(align_counts, key=align_counts.get)  # type: ignore[arg-type]

    parts: list[str] = []
    for i, line_text in enumerate(orig_lines):
        s = line_styles[min(i, n - 1)]
        inline = (
            f"font-size:{s['size']:.1f}px;"
            f"font-weight:{s['weight']};"
            f"font-style:{s['style']};"
            f"color:{s['color']};"
            f"text-decoration:none;"
        )
        safe = _html_escape(line_text)
        parts.append(f'<span style="{inline}">{safe}</span>')

    html = (
        f'<div style="text-align:{dominant_align}; margin:0; padding:0;">'
        + "<br/>".join(parts)
        + "</div>"
    )
    css = (
        f"* {{font-family:sans-serif; font-size:{max_size:.1f}px;"
        f" font-weight:normal; font-style:normal; text-decoration:none;"
        f" margin:0; padding:0; line-height:1.15;}}"
    )
    return html, css


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
            if drawing.get("fill") is None:
                continue
            if drawing.get("fill_opacity", 1.0) < 0.1:
                continue
            r = fitz.Rect(drawing["rect"])
            if r.is_empty or r.width * r.height < _OBSTACLE_MIN_AREA:
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
    limit = page_height - 2.0
    for other in obstacles + siblings:
        # Only elements that horizontally overlap with our rect matter
        if other.x0 >= rect.x1 - 1 or other.x1 <= rect.x0 + 1:
            continue
        if other.y0 > rect.y1 and other.y0 < limit:
            limit = other.y0 - 2.0
    return max(limit, rect.y1)


def _free_x1(
    rect: fitz.Rect,
    page_width: float,
    obstacles: list[fitz.Rect],
    siblings: list[fitz.Rect],
) -> float:
    # Return the rightmost x1 reachable from *rect* without colliding
    limit = page_width - 2.0  # minimum right margin
    for other in obstacles + siblings:
        # Only elements that vertically overlap with our rect matter
        if other.y0 >= rect.y1 - 1 or other.y1 <= rect.y0 + 1:
            continue
        if other.x0 > rect.x1 and other.x0 < limit:
            limit = other.x0 - 2.0
    return max(limit, rect.x1)


_MAX_EXPAND_H = 120.0  # max points a block rect may grow rightward
_EXPAND_H_STEP = 8.0   # horizontal expansion increment (points)


# ── fitting logic ─────────────────────────────────────────────────────────────

def _fit_block(
    html: str,
    css: str,
    rect: fitz.Rect,
    obstacles: list[fitz.Rect],
    siblings: list[fitz.Rect],
    page_height: float,
    page_width: float = 0.0,
) -> fitz.Rect:
    # Return the best rect for this block's insert_htmlbox call
    if page_width <= 0:
        page_width = rect.x1 + 100

    probe_w = max(rect.width + _MAX_EXPAND_H + 20, 1)
    probe_h = max(rect.height + _MAX_EXPAND_DOWN + 20, 1)
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

    # Step 1 — original rect
    spare, scale = _probe(rect)
    if spare >= 0 and scale >= _SCALE_THRESHOLD:
        probe.close()
        return rect

    # Step 2 — expand rightward only (x0 stays fixed)
    max_x1 = _free_x1(rect, page_width, obstacles, siblings)
    max_h_exp = min(_MAX_EXPAND_H, max_x1 - rect.x1)

    expanded = fitz.Rect(rect)
    h_grown = 0.0

    while h_grown < max_h_exp - 0.5:
        step = min(_EXPAND_H_STEP, max_h_exp - h_grown)
        expanded = fitz.Rect(expanded.x0, expanded.y0,
                             expanded.x1 + step, expanded.y1)
        h_grown += step
        spare, scale = _probe(expanded)
        if spare >= 0 and scale >= _SCALE_THRESHOLD:
            probe.close()
            return expanded

    # Step 3 — expand downward (y0 stays fixed)
    max_y1  = _free_y1(expanded, page_height, obstacles, siblings)
    max_exp = min(_MAX_EXPAND_DOWN, max_y1 - expanded.y1)
    v_grown = 0.0

    while v_grown < max_exp - 0.5:
        step = min(_EXPAND_STEP, max_exp - v_grown)
        expanded = fitz.Rect(expanded.x0, expanded.y0,
                             expanded.x1, expanded.y1 + step)
        v_grown += step
        spare, scale = _probe(expanded)
        if spare >= 0 and scale >= _SCALE_THRESHOLD:
            probe.close()
            return expanded

    probe.close()
    return expanded  # scale_low=0 on the real call will handle residual overflow


# ── main entry point ─────────────────────────────────────────────────────────

def translate_pdf(
    input_path: str,
    output_path: str,
    translator: Any,
    target_lang: str,
    *,
    cancel_event: threading.Event | None = None,
    source_lang: str = "auto",
    progress_callback: 'Callable[[int, int], None] | None' = None,
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

    # Progress: num_pages (extract) + 1 (translate) + num_pages (redact)
    _n_pages = src.page_count
    _total_steps = _n_pages + 1 + _n_pages
    _steps_done = 0

    def _report(done: int) -> None:
        nonlocal _steps_done
        _steps_done = done
        if progress_callback is not None:
            progress_callback(_steps_done, _total_steps)

    # ── Phase 1: extract ──────────────────────────────────────────────────────
    page_blocks: list[tuple[int, list[dict]]] = []
    all_units:   list[str] = []
    unit_map:    list[tuple[int, int]] = []

    for page_num in range(src.page_count):
        if cancel_event is not None and cancel_event.is_set():
            src.close()
            raise CancelledError("Translation cancelled")

        page   = src[page_num]
        blocks = _extract_blocks(page)

        pdi = len(page_blocks)
        page_blocks.append((page_num, blocks))

        for bi, block in enumerate(blocks):
            all_units.append(block["text"])
            unit_map.append((pdi, bi))

        _report(page_num + 1)

    # ── Phase 2: translate ────────────────────────────────────────────────────
    if all_units:
        try:
            all_translated = translator.translate_batch(
                all_units, target_lang, cancel_event=cancel_event
            )
        except CancelledError:
            src.close()
            raise
        except Exception:
            logger.exception("Batch translation failed; falling back to per-item")
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

    # Attach translated text to each block
    for unit_idx, (pdi, bi) in enumerate(unit_map):
        _, blocks = page_blocks[pdi]
        tr = all_translated[unit_idx] if unit_idx < len(all_translated) else None
        blocks[bi]["translated"] = tr if tr is not None else blocks[bi]["text"]
        if tr is None:
            errors += 1

    _report(_n_pages + 1)  # extraction + translation done

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

        all_rects = [b["rect"] for b in blocks]

        # Process each block individually: redact then insert.  If insertion
        # fails the original text is already gone so we attempt to re-insert
        # the untranslated text as a fallback.
        for bi, block in enumerate(blocks):
            tr_text     = block.get("translated", block["text"])
            orig_rect   = block["rect"]
            line_styles = block["line_styles"]
            line_rects  = block.get("line_rects", [])
            siblings    = [r for i, r in enumerate(all_rects) if i != bi]

            # Redact per-line rects instead of the whole block rect so that
            # vector drawings (table lines, shapes) within the block area
            # but outside the actual text lines are preserved.
            if line_rects:
                for lr in line_rects:
                    page.add_redact_annot(lr, fill=None)  # type: ignore[arg-type]
            else:
                page.add_redact_annot(orig_rect, fill=None)  # type: ignore[arg-type]
            page.apply_redactions(images=0)  # type: ignore[arg-type]

            html, css = _build_block_html(tr_text, line_styles)

            try:
                fit_rect = _fit_block(
                    html, css, orig_rect,
                    obstacles, siblings, page_height,
                    page_width=page_width,
                )
                result = page.insert_htmlbox(fit_rect, html, css=css, scale_low=0)
                if result[0] < 0:
                    logger.debug(
                        "insert_htmlbox overflow page %d block %d rect=%s",
                        page_num + 1, bi, fit_rect,
                    )
            except Exception:
                logger.exception(
                    "Failed inserting translated block page %d block %d",
                    page_num + 1, bi,
                )
                # Fallback: re-insert original text so the area isn't blank
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

        _report(_n_pages + 1 + pdi + 1)  # extract + translate + pages redacted so far

    # ── Save ──────────────────────────────────────────────────────────────────
    if cancel_event is not None and cancel_event.is_set():
        src.close()
        raise CancelledError("Translation cancelled before saving")

    src.save(str(output_path), garbage=4, deflate=True, clean=True)
    src.close()

    if errors:
        raise RuntimeError(f"PDF translation completed with {errors} issue(s)")
    return None
