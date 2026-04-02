from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from typing import Any, Callable
import logging
import threading
import unicodedata
import re
from ..utils import CancelledError

logger = logging.getLogger(__name__)

# Control-character pattern: matches C0/C1 control chars except tab, newline, carriage return
_CONTROL_CHAR_RE = re.compile(
    r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]'
)

# Separator used to group multiple cells into a single translation unit.
# Chosen to be extremely unlikely in real spreadsheet data.
_CELL_SEP = "\n\u27EASEP\u27EB\n"

# Maximum characters per grouped row before falling back to per-cell.
_GROUP_MAX_CHARS = 4000


def _sanitize_text(text: str) -> str:
    # Normalize Unicode and strip problematic control characters
    text = unicodedata.normalize("NFC", text)
    text = _CONTROL_CHAR_RE.sub("", text)
    return text


def translate_xlsx(input_path: str, output_path: str, translator: Any, target_lang: str, *, cancel_event: threading.Event | None = None, source_lang: str = "auto", progress_callback: 'Callable[[int, int], None] | None' = None):
    # batch-translate XLSX with row-level contextual grouping.
    # When source_lang=="auto" row grouping is skipped so each cell is its own translation
    # unit, letting the API auto-detect the language per cell.
    wb = load_workbook(filename=input_path)

    # --- collect every string cell that is writable, grouped by row ---
    # Each row group: list of (cell, sanitized_text) pairs
    RowGroup = list[tuple[Any, str]]
    row_groups: list[RowGroup] = []
    current_row: RowGroup = []
    current_row_key: tuple | None = None  # (sheet_title, row_number)

    for sheet in wb.worksheets:
        for row in sheet.iter_rows(values_only=False):
            row_cells: RowGroup = []
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                val = cell.value
                if isinstance(val, str) and val.strip() and not val.startswith("="):
                    row_cells.append((cell, _sanitize_text(val)))
            if row_cells:
                row_groups.append(row_cells)

    total_cells = sum(len(rg) for rg in row_groups)
    logger.info("XLSX '%s': collected %d translatable string cells in %d rows",
                input_path, total_cells, len(row_groups))

    if not row_groups:
        logger.warning("No translatable text found in XLSX file '%s'", input_path)
        wb.save(output_path)
        return

    # --- build translation units: group rows or send individually ---
    # A "unit" is either a joined row (multiple cells with separator) or a single cell.
    # After translation we split on separator to recover per-cell results.
    units: list[str] = []
    # Map: unit_index -> list of (cell, original_text) to write back
    unit_cells: list[list[tuple[Any, str]]] = []

    # In auto-detect mode do NOT group cells — every cell is its own unit so the
    # translation API sees a single-language segment and can auto-detect correctly.
    group_rows = source_lang != "auto"

    for rg in row_groups:
        row_texts = [t for _, t in rg]
        total_chars = sum(len(t) for t in row_texts)

        if group_rows and len(rg) > 1 and total_chars <= _GROUP_MAX_CHARS:
            # Group the row into a single unit with separators
            units.append(_CELL_SEP.join(row_texts))
            unit_cells.append(rg)
        else:
            # Send each cell individually (row too large, single cell, or auto mode)
            for cell, text in rg:
                units.append(text)
                unit_cells.append([(cell, text)])

    # --- batch-translate ---
    total_units = len(units)
    try:
        translated = translator.translate_batch(units, target_lang, cancel_event=cancel_event)
    except CancelledError:
        raise
    except Exception:
        logger.exception("Batch translation failed for XLSX; falling back to per-item")
        translated = []
        for t in units:
            try:
                r = translator.translate_text(t, target_lang)
                translated.append(r if r is not None else t)
            except Exception:
                logger.exception("Per-item fallback also failed")
                translated.append(t)

    if progress_callback is not None:
        progress_callback(total_units, total_units)

    # --- write results back ---
    errors = 0
    for cells_in_unit, tr_text in zip(unit_cells, translated):
        if tr_text is None:
            # Translation returned None — keep originals
            errors += 1
            continue

        if len(cells_in_unit) == 1:
            # Single cell unit — direct assignment
            cell, orig = cells_in_unit[0]
            cell.value = tr_text
        else:
            # Grouped row — split on separator
            parts = tr_text.split(_CELL_SEP)
            if len(parts) == len(cells_in_unit):
                for (cell, orig), part in zip(cells_in_unit, parts):
                    cell.value = part.strip() if part else orig
            else:
                # Separator was consumed/mangled by the model — fall back to
                # per-cell translation for this row
                logger.debug(
                    "XLSX row separator mismatch: expected %d, got %d; per-cell fallback",
                    len(cells_in_unit), len(parts),
                )
                for cell, orig in cells_in_unit:
                    try:
                        r = translator.translate_text(orig, target_lang)
                        cell.value = r if r is not None else orig
                    except Exception:
                        logger.exception("Per-cell fallback failed")
                        cell.value = orig
                        errors += 1

    # Check for cancellation before saving
    if cancel_event is not None and cancel_event.is_set():
        raise CancelledError("Translation cancelled before saving XLSX")

    wb.save(output_path)
    if errors:
        raise RuntimeError(f"Translation completed with {errors} failed cells")
