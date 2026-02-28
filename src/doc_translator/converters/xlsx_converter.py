from openpyxl import load_workbook
from typing import Any
import logging


def translate_xlsx(input_path: str, output_path: str, translator: Any, target_lang: str):
    wb = load_workbook(filename=input_path)
    errors = 0
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(values_only=False):
            from openpyxl.cell.cell import MergedCell
            for cell in row:
                val = cell.value
                if isinstance(val, str) and val.strip():
                    try:
                        translated = translator.translate_text(val, target_lang)
                    except Exception:
                        logging.exception("Failed to translate cell")
                        errors += 1
                        translated = val
                    if translated is None:
                        translated = val
                    if not isinstance(cell, MergedCell):
                        cell.value = translated
    wb.save(output_path)
    if errors:
        raise RuntimeError(f"Translation completed with {errors} failed cells")
