from typing import Any
from pypdf import PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from pathlib import Path
import logging


def translate_pdf(input_path: str, output_path: str, translator: Any, target_lang: str):
    out_path = Path(output_path)
    width, height = letter
    errors = 0

    try:
        import pdfplumber

        use_pdfplumber = True
    except Exception:
        use_pdfplumber = False

    c = canvas.Canvas(str(out_path), pagesize=letter)

    if use_pdfplumber:
        with pdfplumber.open(input_path) as pdf:
            for page in pdf.pages:
                words = page.extract_words()
                words.sort(key=lambda w: (w.get("top", 0), w.get("x0", 0)))
                
                for w in words:
                    text = w.get("text") or ""
                    if not text.strip():
                        continue
                    
                    try:
                        translated = translator.translate_text(text, target_lang)
                        if translated is None:
                            translated = text
                    except Exception:
                        logging.exception("Failed to translate PDF word")
                        errors += 1
                        translated = text

                    # Coordinate mapping
                    top = w.get("top", 0)
                    y = height - (top / page.height) * height - 40
                    x = (w.get("x0", 0) / page.width) * width + 40
                    c.drawString(x, max(40, y), translated[:200])
                c.showPage()
        c.save()
        if errors:
            raise RuntimeError(f"Translation completed with {errors} failed words")
        return

    # Fallback: pypdf logic
    reader = PdfReader(input_path)
    for page in reader.pages:
        text = page.extract_text() or ""
        if text.strip():
            try:
                translated = translator.translate_text(text, target_lang)
            except Exception:
                logging.exception("Failed to translate PDF page text")
                errors += 1
                translated = text
            if translated is None:
                translated = text
        else:
            translated = ""
        y = height - 40
        for line in translated.splitlines():
            c.drawString(40, y, line[:200])
            y -= 14
            if y < 40:
                c.showPage()
                y = height - 40
        c.showPage()
    c.save()
    if errors:
        raise RuntimeError(f"Translation completed with {errors} failed pages/words")
