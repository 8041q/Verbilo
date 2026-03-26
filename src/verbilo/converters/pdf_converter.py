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

def _css_weight(font_name: str, flags: int) -> int:
    """Map a PDF span font name + flags to a CSS font-weight integer.

    PyMuPDF flags bit 16 = bold.  Font name suffixes are also checked because
    some fonts (e.g. NotoSansCJKsc-Medium) encode weight only in the name.
    """
    name_lc = font_name.lower()
    if flags & 16 or "bold" in name_lc or "black" in name_lc or "heavy" in name_lc:
        return 700
    if "semibold" in name_lc or "demibold" in name_lc or "medium" in name_lc:
        return 500
    if "light" in name_lc or "thin" in name_lc:
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
        page_width  – page width (for alignment inference)
        page_height – page height
    """
    page_w = page.rect.width
    page_h = page.rect.height
    d = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)

    out: list[dict] = []
    for block in d.get("blocks", []):
        if block.get("type") != 0:
            continue

        lines_data: list[tuple[str, dict]] = []

        for line in block.get("lines", []):
            spans = line.get("spans", [])
            valid = [s for s in spans if s.get("text", "").strip()]
            if not valid:
                continue

            line_text = "".join(s["text"] for s in valid).strip()
            if not line_text:
                continue

            # Dominant style = span with the most characters
            dominant = max(valid, key=lambda s: len(s["text"]))

            # Infer alignment from line bbox vs page width
            lb = fitz.Rect(line["bbox"])
            cx = lb.x0 + lb.width / 2
            if abs(cx - page_w / 2) < page_w * 0.08:
                align = "center"
            elif lb.x0 > page_w * 0.60:
                align = "right"
            else:
                align = "left"

            lines_data.append((line_text, {
                "size":   dominant["size"],
                "weight": _css_weight(dominant["font"], dominant["flags"]),
                "style":  _css_style(dominant["flags"]),
                "color":  _css_color(dominant["color"]),
                "align":  align,
            }))

        if not lines_data:
            continue

        out.append({
            "rect":        fitz.Rect(block["bbox"]),
            "text":        "\n".join(t for t, _ in lines_data),
            "line_styles": [s for _, s in lines_data],
            "page_width":  page_w,
            "page_height": page_h,
        })

    return out


# ── HTML builder ─────────────────────────────────────────────────────────────

def _build_block_html(translated: str, line_styles: list[dict]) -> tuple[str, str]:
    """Build HTML + CSS for *translated* text using the original per-line styles.

    Splits *translated* on \\n and applies the i-th original style to the i-th
    translated line (last style reused for any extra lines).

    The CSS global rule hard-resets all inheritable properties so that only our
    inline span styles take effect — critical for correct bold/italic rendering
    inside insert_htmlbox.
    """
    orig_lines = translated.split("\n")
    n = len(line_styles)
    max_size = max((s["size"] for s in line_styles), default=11.0)

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
        parts.append(
            f'<div style="text-align:{s["align"]};">'
            f'<span style="{inline}">{safe}</span>'
            f"</div>"
        )

    html = "\n".join(parts)
    css = (
        f"* {{font-family:sans-serif; font-size:{max_size:.1f}px;"
        f" font-weight:normal; font-style:normal; text-decoration:none;}}"
    )
    return html, css


# ── obstacle / free-space helpers ────────────────────────────────────────────

def _collect_obstacles(page: fitz.Page) -> list[fitz.Rect]:
    """Return rects of all image blocks and significant filled drawings on *page*."""
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


# ── fitting logic ─────────────────────────────────────────────────────────────

def _fit_block(
    html: str,
    css: str,
    rect: fitz.Rect,
    obstacles: list[fitz.Rect],
    siblings: list[fitz.Rect],
    page_height: float,
) -> fitz.Rect:
    """Return the best rect for this block's insert_htmlbox call.

    1. Probe the original rect with scale_low=0.
       If scale >= SCALE_THRESHOLD — text fits well, return original rect unchanged.
    2. Otherwise expand downward in small steps into obstacle/sibling-free space,
       up to _MAX_EXPAND_DOWN points, until scale becomes acceptable.
    3. If still not good enough, return the largest allowed rect anyway;
       scale_low=0 will handle the remainder.

    All probes use a throw-away document (no writes to the real PDF).
    """
    probe = fitz.open()
    # Probe page height just needs to be tall enough for the expanded case
    probe.new_page(width=max(rect.width, 1), height=max(rect.height + _MAX_EXPAND_DOWN + 20, 1))

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

    # Step 2 — expand downward into free space
    max_y1   = _free_y1(rect, page_height, obstacles, siblings)
    max_exp  = min(_MAX_EXPAND_DOWN, max_y1 - rect.y1)
    expanded = fitz.Rect(rect)
    grown    = 0.0

    while grown < max_exp - 0.5:
        step     = min(_EXPAND_STEP, max_exp - grown)
        expanded = fitz.Rect(expanded.x0, expanded.y0, expanded.x1, expanded.y1 + step)
        grown   += step
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
) -> str | None:
    """Translate a PDF in-place while preserving the original layout.

    Phase 1 — Extract one translation unit per text block (all pages).
    Phase 2 — Batch-translate everything in a single API call.
    Phase 3 — Per page: redact originals, insert translations via insert_htmlbox
               with scale_low=0; expand block rect only when auto-scale is too aggressive.
    """
    src = fitz.open(input_path)

    if _is_ocr_required(src):
        logger.warning(
            "Skipping '%s': scanned/image PDF (OCR required).",
            Path(input_path).name,
        )
        src.close()
        return "skipped-ocr"

    errors = 0

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

    # ── Phase 3: redact + insert ──────────────────────────────────────────────
    for pdi, (page_num, blocks) in enumerate(page_blocks):
        if not blocks:
            continue

        if cancel_event is not None and cancel_event.is_set():
            src.close()
            raise CancelledError("Translation cancelled")

        page        = src[page_num]
        page_height = page.rect.height

        # Collect obstacles before redaction (image/drawing positions are stable)
        obstacles = _collect_obstacles(page)

        # Redact all original block rects in a single pass first
        for block in blocks:
            page.add_redact_annot(block["rect"], fill=None)  # type: ignore[arg-type]
        page.apply_redactions(images=0)  # type: ignore[arg-type]

        all_rects = [b["rect"] for b in blocks]

        for bi, block in enumerate(blocks):
            tr_text     = block.get("translated", block["text"])
            orig_rect   = block["rect"]
            line_styles = block["line_styles"]
            siblings    = [r for i, r in enumerate(all_rects) if i != bi]

            html, css = _build_block_html(tr_text, line_styles)

            try:
                fit_rect = _fit_block(
                    html, css, orig_rect,
                    obstacles, siblings, page_height,
                )
                result = page.insert_htmlbox(fit_rect, html, css=css, scale_low=0)
                if result[0] < 0:
                    logger.debug(
                        "insert_htmlbox overflow page %d block %d rect=%s",
                        page_num + 1, bi, fit_rect,
                    )
            except Exception:
                logger.exception(
                    "Failed inserting block page %d block %d", page_num + 1, bi
                )
                errors += 1

    # ── Save ──────────────────────────────────────────────────────────────────
    if cancel_event is not None and cancel_event.is_set():
        src.close()
        raise CancelledError("Translation cancelled before saving")

    src.save(str(output_path), garbage=4, deflate=True, clean=True)
    src.close()

    if errors:
        raise RuntimeError(f"PDF translation completed with {errors} issue(s)")
    return None
