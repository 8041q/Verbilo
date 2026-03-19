from docx.api import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph as _DocxParagraph
from typing import Any
import logging
import re
import threading
from ..utils import CancelledError

logger = logging.getLogger(__name__)


# Run-formatting helpers

def _run_fmt(run) -> tuple:
    # Return a hashable tuple of the formatting properties that matter
    f = run.font
    return (
        f.bold,
        f.italic,
        f.underline,
        f.strike,
        str(f.size),
        str(f.color.rgb) if f.color and f.color.rgb else None,
        f.name,
    )


def _all_same_format(runs) -> bool:
    # True when every run in *runs* shares the same formatting
    fmts = {_run_fmt(r) for r in runs}
    return len(fmts) <= 1



# Paragraph-level text helpers

def _paragraph_full_text(runs) -> str:
    # Join all run texts into the full paragraph string.
    return "".join(r.text for r in runs)


def _redistribute_translated(runs, translated: str) -> None:
    #Write *translated* back into *runs*, preserving formatting boundaries.
    # Identify runs that held actual (non-whitespace) text
    text_runs = [(i, run) for i, run in enumerate(runs) if run.text and run.text.strip()]

    if not text_runs:
        return

    # ---- Single run or uniform formatting → simple assignment ----
    if len(text_runs) == 1 or _all_same_format([r for _, r in text_runs]):
        text_runs[0][1].text = translated
        for j, (i, run) in enumerate(text_runs):
            if j > 0:
                run.text = ""
        # Clear whitespace-only runs
        text_run_indices = {i for i, _ in text_runs}
        for i, run in enumerate(runs):
            if i not in text_run_indices and run.text:
                run.text = ""
        return

    # ---- Multiple runs with mixed formatting → proportional distribution ----
    words = translated.split()
    total_words = len(words)

    if total_words == 0:
        for _, run in text_runs:
            run.text = ""
        return

    # Proportions based on *original* character counts (stripped)
    orig_lengths = [max(len(run.text.strip()), 1) for _, run in text_runs]
    total_orig = sum(orig_lengths)

    assigned = 0
    for j, (i, run) in enumerate(text_runs):
        if j == len(text_runs) - 1:
            # Last run gets all remaining words
            run.text = " ".join(words[assigned:])
        else:
            proportion = orig_lengths[j] / total_orig
            n_words = max(1, round(proportion * total_words))
            end_idx = min(assigned + n_words, total_words)
            run.text = " ".join(words[assigned:end_idx])
            if end_idx < total_words:
                run.text += " "          # space separating from next run
            assigned = end_idx

    # Clear whitespace-only runs (no longer needed)
    text_run_indices = {i for i, _ in text_runs}
    for i, run in enumerate(runs):
        if i not in text_run_indices and run.text:
            run.text = ""


# Paragraph collectors

def _iter_all_paragraphs(doc):
    # Yield every paragraph in the document: body, tables, headers, footers, text boxes
    # Track yielded paragraph element ids to avoid duplicates
    seen: set[int] = set()

    def _track_and_yield(para):
        pid = id(para._p)
        if pid not in seen:
            seen.add(pid)
            return True
        return False

    # Body paragraphs
    for para in doc.paragraphs:
        if _track_and_yield(para):
            yield para

    # Table cell paragraphs
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if _track_and_yield(para):
                        yield para

    # Headers and footers in every section
    for section in doc.sections:
        for hf in (section.header, section.footer,
                    section.first_page_header, section.first_page_footer,
                    section.even_page_header, section.even_page_footer):
            try:
                if hf is None or hf.is_linked_to_previous:
                    continue
                for para in hf.paragraphs:
                    if _track_and_yield(para):
                        yield para
                for table in hf.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for para in cell.paragraphs:
                                if _track_and_yield(para):
                                    yield para
                # Text boxes inside header/footer
                for txbx in hf._element.iter(qn('w:txbxContent')):
                    for p_elem in txbx.findall(qn('w:p')):
                        if id(p_elem) not in seen:
                            seen.add(id(p_elem))
                            try:
                                yield _DocxParagraph(p_elem, doc.part)
                            except Exception:
                                pass
            except Exception:
                continue

    # Text boxes anywhere in the main document body
    for txbx in doc.element.iter(qn('w:txbxContent')):
        for p_elem in txbx.findall(qn('w:p')):
            if id(p_elem) not in seen:
                seen.add(id(p_elem))
                try:
                    yield _DocxParagraph(p_elem, doc.part)
                except Exception:
                    pass


# TOC / structural paragraph helpers

_TOC_TRAILING_RE = re.compile(r'\t\s*\d+\s*$')


def _is_toc_paragraph(para) -> bool:
    # Return True if *para* looks like a Table-of-Contents entry
    # Check paragraph style name
    try:
        style_name = (para.style.name or "").lower()
        if style_name.startswith("toc") or "table of content" in style_name:
            return True
    except Exception:
        pass
    # Check content: text + tab + page number
    try:
        full = para.text  # includes tab characters
        if _TOC_TRAILING_RE.search(full):
            return True
    except Exception:
        pass
    return False


def _get_translatable_runs(para, runs: list) -> list:
    # Return the subset of *runs* that should be translated.
    if not _is_toc_paragraph(para):
        return runs

    # Find the first run that contains a <w:tab/> element
    result: list = []
    for run in runs:
        if run._r.find(qn('w:tab')) is not None:
            break
        result.append(run)
    return result if result else runs  # fallback to all runs if no tab found



# Main entry point

def translate_docx(input_path: str, output_path: str, translator: Any, target_lang: str, *, cancel_event: threading.Event | None = None):
    # Batch-translate DOCX at paragraph level while preserving formatting
    doc = Document(input_path)

    # --- Collect paragraph texts (one entry per paragraph) ---
    ParaInfo = tuple  # (paragraph, runs, full_text)
    para_infos: list[tuple] = []
    para_texts: list[str] = []

    for para in _iter_all_paragraphs(doc):
        all_runs = list(para.runs)
        if not all_runs:
            continue
        # For TOC entries, only translate runs before the tab+page-number
        runs = _get_translatable_runs(para, all_runs)
        if not runs:
            continue
        full_text = _paragraph_full_text(runs)
        if not full_text or not full_text.strip():
            continue
        para_infos.append((para, runs, full_text))
        para_texts.append(full_text)

    if not para_texts:
        doc.save(output_path)
        return

    # --- Batch-translate all paragraph texts ---
    try:
        translated = translator.translate_batch(para_texts, target_lang, cancel_event=cancel_event)
    except CancelledError:
        raise
    except Exception:
        logger.exception("Batch translation failed for DOCX; falling back to per-item")
        translated = []
        for t in para_texts:
            try:
                r = translator.translate_text(t, target_lang)
                translated.append(r if r is not None else t)
            except Exception:
                logger.exception("Per-item fallback also failed")
                translated.append(t)

    # --- Write results back into the runs ---
    errors = 0
    for (para, runs, orig_text), tr in zip(para_infos, translated):
        if tr is None:
            tr = orig_text
            errors += 1
        _redistribute_translated(runs, tr)

    # Check for cancellation before saving
    if cancel_event is not None and cancel_event.is_set():
        raise CancelledError("Translation cancelled before saving DOCX")

    doc.save(output_path)
    if errors:
        raise RuntimeError(f"Translation completed with {errors} failed paragraphs")
