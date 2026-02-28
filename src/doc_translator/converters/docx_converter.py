from docx.api import Document
from typing import Any
import logging


def translate_docx(input_path: str, output_path: str, translator: Any, target_lang: str):
    doc = Document(input_path)
    errors = 0

    for para in doc.paragraphs:
        for run in para.runs:
            text = run.text
            if text and text.strip():
                try:
                    translated = translator.translate_text(text, target_lang)
                except Exception as e:
                    logging.exception("Failed to translate paragraph run")
                    errors += 1
                    translated = text
                if translated is None:
                    translated = text
                run.text = translated

    # tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        text = run.text
                        if text and text.strip():
                            try:
                                translated = translator.translate_text(text, target_lang)
                            except Exception:
                                logging.exception("Failed to translate table cell run")
                                errors += 1
                                translated = text
                            if translated is None:
                                translated = text
                            run.text = translated

    doc.save(output_path)
    if errors:
        raise RuntimeError(f"Translation completed with {errors} failed spans")
