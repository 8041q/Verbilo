"""PDF translator using PyMuPDF (fitz) for in-place text replacement.

Preserves images, vector graphics, page dimensions, and approximate text
positioning/formatting.  If a PDF appears to require OCR (scanned / image-
only pages), the file is **skipped** with a log message instead of producing
broken output.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

import fitz  # PyMuPDF

logger = logging.getLogger(__name__)

# Minimum average extractable characters per page to consider a PDF
# as containing real (non-OCR) text.
_MIN_CHARS_PER_PAGE = 20


def _is_ocr_required(doc: fitz.Document) -> bool:
    """Return *True* if the document appears to be scanned / image-only."""
    if doc.page_count == 0:
        return True
    total_chars = 0
    for page in doc:
        total_chars += len(page.get_text("text").strip())
    avg = total_chars / doc.page_count
    return avg < _MIN_CHARS_PER_PAGE


def translate_pdf(input_path: str, output_path: str, translator: Any, target_lang: str):
    """Translate text inside a PDF while preserving layout, images, and fonts."""
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

    for page_num in range(src.page_count):
        page = src[page_num]
        # Get structured text: blocks -> lines -> spans
        text_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)
        blocks = text_dict.get("blocks", [])

        # Collect all translatable text spans with their metadata
        span_infos: list[dict] = []  # each entry: {rect, text, size, color, flags, font}
        for block in blocks:
            if block.get("type") != 0:  # 0 = text block; skip image blocks
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    text = span.get("text", "").strip()
                    if not text:
                        continue
                    span_infos.append({
                        "rect": fitz.Rect(span["bbox"]),
                        "text": span["text"],
                        "size": span.get("size", 11),
                        "color": span.get("color", 0),
                        "flags": span.get("flags", 0),
                        "font": span.get("font", "helv"),
                    })

        if not span_infos:
            continue

        # Batch-translate all span texts for this page
        original_texts = [s["text"] for s in span_infos]
        try:
            translated_texts = translator.translate_batch(original_texts, target_lang)
        except Exception:
            logger.exception("Batch translation failed on page %d; falling back to per-item", page_num + 1)
            translated_texts = []
            for t in original_texts:
                try:
                    r = translator.translate_text(t, target_lang)
                    translated_texts.append(r if r is not None else t)
                except Exception:
                    logger.exception("Per-item fallback failed")
                    translated_texts.append(t)
                    errors += 1

        # Redact original text and insert translated text
        for info, tr_text in zip(span_infos, translated_texts):
            if tr_text is None:
                tr_text = info["text"]
                errors += 1

            rect = info["rect"]
            # Add redaction annotation to remove original text
            page.add_redact_annot(rect)

        # Apply all redactions at once (removes original text, keeps images)
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

        # Re-insert translated text
        for info, tr_text in zip(span_infos, translated_texts):
            if tr_text is None:
                tr_text = info["text"]
            rect = info["rect"]
            fontsize = info["size"]
            # Convert integer colour to (r, g, b) tuple for fitz
            c = info["color"]
            if isinstance(c, int):
                r_c = ((c >> 16) & 0xFF) / 255.0
                g_c = ((c >> 8) & 0xFF) / 255.0
                b_c = (c & 0xFF) / 255.0
                color = (r_c, g_c, b_c)
            else:
                color = (0, 0, 0)

            # Use a built-in font that supports broad character sets
            try:
                rc = page.insert_textbox(
                    rect,
                    tr_text,
                    fontsize=fontsize,
                    fontname="helv",
                    color=color,
                    align=fitz.TEXT_ALIGN_LEFT,
                )
                # rc < 0 means text didn't fit; try with smaller font
                if rc < 0:
                    page.insert_textbox(
                        rect,
                        tr_text,
                        fontsize=max(fontsize * 0.7, 5),
                        fontname="helv",
                        color=color,
                        align=fitz.TEXT_ALIGN_LEFT,
                    )
            except Exception:
                logger.exception("Failed to insert translated text on page %d", page_num + 1)
                errors += 1

    src.save(str(output_path), garbage=4, deflate=True)
    src.close()

    if errors:
        raise RuntimeError(f"PDF translation completed with {errors} issues")
