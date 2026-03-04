# in-place PDF translator using PyMuPDF; skips scanned/OCR-only PDFs

from __future__ import annotations

import logging
import threading
from pathlib import Path
from typing import Any

import fitz  # PyMuPDF
from ..utils import CancelledError

logger = logging.getLogger(__name__)

_MIN_CHARS_PER_PAGE = 20  # min avg chars/page to treat as real text (not scanned)


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


# ---------------------------------------------------------------------------
# Font helpers – find a Unicode-capable font for text reinsertion
# ---------------------------------------------------------------------------

_UNICODE_FONT_PATH: str | None = None  # cached after first lookup


def _find_unicode_font() -> str | None:
    # Return the filesystem path to a Unicode-capable TrueType font
    global _UNICODE_FONT_PATH
    if _UNICODE_FONT_PATH is not None:
        return _UNICODE_FONT_PATH

    # --- Try pymupdf-fonts package (exposes font file paths) ---
    try:
        import pymupdf_fonts  # noqa: F401
        # If the package is importable, PyMuPDF will have extra built-in font
        # names like "notos" (Noto Sans Regular) available via insert_font().
        _UNICODE_FONT_PATH = "__pymupdf_fonts__"
        return _UNICODE_FONT_PATH
    except ImportError:
        pass

    # --- System font search ---
    import sys
    candidates: list[str] = []
    if sys.platform == "win32":
        import os
        windir = os.environ.get("WINDIR", r"C:\Windows")
        fonts_dir = Path(windir) / "Fonts"
        candidates = [
            str(fonts_dir / "arial.ttf"),
            str(fonts_dir / "calibri.ttf"),
            str(fonts_dir / "segoeui.ttf"),
            str(fonts_dir / "tahoma.ttf"),
        ]
    elif sys.platform == "darwin":
        candidates = [
            "/System/Library/Fonts/Helvetica.ttc",
            "/Library/Fonts/Arial.ttf",
            "/System/Library/Fonts/Supplemental/Arial.ttf",
        ]
    else:  # Linux / other
        candidates = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
            "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf",
            "/usr/share/fonts/TTF/DejaVuSans.ttf",
        ]

    for path in candidates:
        if Path(path).is_file():
            _UNICODE_FONT_PATH = path
            return _UNICODE_FONT_PATH

    _UNICODE_FONT_PATH = ""  # empty string → not found (avoid re-searching)
    return None


def _register_font(page: fitz.Page) -> str:
    # Register a Unicode font on *page* and return the fontname to use.
    font_path = _find_unicode_font()
    if font_path == "__pymupdf_fonts__":
        # pymupdf-fonts is installed; "notos" (Noto Sans) is available globally
        return "notos"
    if font_path:
        try:
            fontname = "unif0"
            page.insert_font(fontname=fontname, fontfile=font_path)
            return fontname
        except Exception:
            logger.debug("Failed to register font %s; falling back to helv", font_path)
    return "helv"


# ---------------------------------------------------------------------------
# Line-level span grouping helpers
# ---------------------------------------------------------------------------

def _group_spans_by_line(blocks: list[dict]) -> list[dict]:
    # Group spans into lines and return per-line metadata.
    lines_out: list[dict] = []
    for block in blocks:
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue
            # Filter out empty spans
            valid_spans = [s for s in spans if s.get("text", "").strip()]
            if not valid_spans:
                continue

            # Build combined text (preserve inter-span spacing)
            combined_text = " ".join(s["text"].strip() for s in valid_spans)
            if not combined_text.strip():
                continue

            # Determine predominant size/color by character count
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


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def translate_pdf(input_path: str, output_path: str, translator: Any, target_lang: str, *, cancel_event: threading.Event | None = None):
    # Translate text in-place, preserving layout and images
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
        text_dict: dict[str, Any] = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)  # type: ignore
        blocks = text_dict.get("blocks", [])

        # Group spans by line for context-aware translation
        line_infos = _group_spans_by_line(blocks)
        if not line_infos:
            continue

        # Batch-translate all line texts for this page
        original_texts = [li["text"] for li in line_infos]
        try:
            translated_texts = translator.translate_batch(original_texts, target_lang, cancel_event=cancel_event)
        except CancelledError:
            src.close()
            raise
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

        # Redact original text per-line (single rect per line → less collateral damage)
        # fill=None → do NOT paint white over the redacted area (preserves background/images)
        for info in line_infos:
            page.add_redact_annot(info["rect"], fill=None)  # type: ignore[arg-type]
        # 0 = PDF_REDACT_IMAGE_NONE → leave underlying images untouched
        page.apply_redactions(images=0)  # type: ignore[arg-type]

        # Register a Unicode-capable font on this page
        fontname = _register_font(page)

        # Re-insert translated text
        for info, tr_text in zip(line_infos, translated_texts):
            if tr_text is None:
                tr_text = info["text"]
                errors += 1

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

            # Try inserting with original font size; shrink iteratively if it
            # doesn't fit, with a broader rect height as last resort.
            inserted = False
            for scale in (1.0, 0.9, 0.8, 0.7, 0.6):
                try:
                    trial_size = max(fontsize * scale, 4)
                    # Slightly expand rect height when shrinking to avoid clipping
                    trial_rect = fitz.Rect(rect)
                    if scale < 1.0:
                        extra_h = rect.height * (1.0 - scale) * 0.5
                        trial_rect.y1 += extra_h
                    rc = page.insert_textbox(
                        trial_rect,
                        tr_text,
                        fontsize=trial_size,
                        fontname=fontname,
                        color=color,
                        align=fitz.TEXT_ALIGN_LEFT,
                    )
                    if rc >= 0:
                        inserted = True
                        break
                except Exception:
                    continue

            if not inserted:
                # Final fallback: just insert at whatever size, accepting overflow
                try:
                    page.insert_textbox(
                        rect,
                        tr_text,
                        fontsize=max(fontsize * 0.5, 4),
                        fontname=fontname,
                        color=color,
                        align=fitz.TEXT_ALIGN_LEFT,
                    )
                except Exception:
                    logger.exception("Failed to insert translated text on page %d", page_num + 1)
                    errors += 1

    # Check for cancellation before saving
    if cancel_event is not None and cancel_event.is_set():
        src.close()
        raise CancelledError("Translation cancelled before saving PDF")

    src.save(str(output_path), garbage=4, deflate=True, clean=True)
    src.close()

    if errors:
        raise RuntimeError(f"PDF translation completed with {errors} issues")
