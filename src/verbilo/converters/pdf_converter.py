# in-place PDF translator using PyMuPDF; skips scanned/OCR-only PDFs
# Uses insert_htmlbox for automatic text fitting, wrapping, font selection,
# and RTL support.  No manual font management needed.

from __future__ import annotations

import logging
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

def _group_spans_by_line(blocks: list[dict]) -> list[dict]:
    # Group spans into lines and return per-line metadata
    lines_out: list[dict] = []
    for block in blocks:
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue
            valid_spans = [s for s in spans if s.get("text", "").strip()]
            if not valid_spans:
                continue

            combined_text = " ".join(s["text"].strip() for s in valid_spans)
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

            lines_out.append({
                "rect": line_rect,
                "text": combined_text,
                "size": best_size,
                "color": best_color,
                "flags": combined_flags,
                "spans": valid_spans,
            })
    return lines_out



# HTML builder for insert_htmlbox
def _build_html(text: str, fontsize: float, color: int, flags: int) -> tuple[str, str]:
    # Build an HTML snippet and CSS string for insert_htmlbox. Returns (html_text, css_string).
    # Convert integer colour to hex
    if isinstance(color, int):
        hex_color = f"#{color:06x}"
    else:
        hex_color = "#000000"

    bold = bool(flags & _FLAG_BOLD)
    italic = bool(flags & _FLAG_ITALIC)

    weight = "bold" if bold else "normal"
    style = "italic" if italic else "normal"

    # Escape HTML special characters
    safe_text = (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )
    # Convert newlines to <br>
    safe_text = safe_text.replace("\n", "<br>")

    html = f'<span style="color:{hex_color};font-weight:{weight};font-style:{style};">{safe_text}</span>'

    css = f"* {{font-size:{fontsize:.1f}px; font-family: sans-serif;}}"

    return html, css


# Main entry point

def translate_pdf(
    input_path: str,
    output_path: str,
    translator: Any,
    target_lang: str,
    *,
    cancel_event: threading.Event | None = None,
) -> str | None:
    """Translate text in a PDF while preserving layout.

    For each page:
    1.  Extract structured text (blocks -> lines -> spans).
    2.  Translate all lines in a single batch where possible.
    3.  Redact original text areas.
    4.  Re-insert translated text using insert_htmlbox, which handles
        automatic text wrapping, font sizing, RTL support, and complex
        script shaping via HarfBuzz.
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

    for page_num in range(src.page_count):
        # Honour cancellation between pages
        if cancel_event is not None and cancel_event.is_set():
            src.close()
            raise CancelledError("Translation cancelled")

        page = src[page_num]

        # Get structured text: blocks -> lines -> spans
        text_dict: dict[str, Any] = page.get_text(  # type: ignore[assignment]
            "dict", flags=fitz.TEXT_PRESERVE_WHITESPACE
        )
        blocks = text_dict.get("blocks", [])

        # Group spans by line for context-aware translation
        line_infos = _group_spans_by_line(blocks)
        if not line_infos:
            continue


        # Batch-translate all line texts for this page
        original_texts = [li["text"] for li in line_infos]
        try:
            translated_texts = translator.translate_batch(
                original_texts, target_lang, cancel_event=cancel_event
            )
        except CancelledError:
            src.close()
            raise
        except Exception:
            logger.exception(
                "Batch translation failed on page %d; falling back to per-item",
                page_num + 1,
            )
            translated_texts = []
            for t in original_texts:
                try:
                    r = translator.translate_text(t, target_lang)
                    translated_texts.append(r if r is not None else t)
                except Exception:
                    logger.exception("Per-item fallback failed")
                    translated_texts.append(t)
                    errors += 1

        # Redact original text per-line
        # fill=None -> do NOT paint white over the area (preserves background)
        for info in line_infos:
            page.add_redact_annot(info["rect"], fill=None)  # type: ignore[arg-type]
        # 0 = PDF_REDACT_IMAGE_NONE -> leave underlying images untouched
        page.apply_redactions(images=0)  # type: ignore[arg-type]


        # Insert translated text using insert_htmlbox
        for info, tr_text in zip(line_infos, translated_texts):
            if tr_text is None:
                tr_text = info["text"]
                errors += 1

            orig_rect: fitz.Rect = info["rect"]
            fontsize: float = info["size"]

            # Build the HTML + CSS for this line
            html, css = _build_html(tr_text, fontsize, info["color"], info["flags"])

            try:
                # insert_htmlbox with scale_low=0 (default) will automatically
                # shrink text to fit the box. It also handles text wrapping,
                # RTL scripts, and pulls in Noto fonts for any missing glyphs.
                result = page.insert_htmlbox(orig_rect, html, css=css)
                if result[0] < 0:
                    # Text could not fit even after scaling — log but continue
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
