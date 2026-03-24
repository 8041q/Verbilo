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


# Tagged-run translation for mixed-format paragraphs

_TAG_OPEN = "\u27E8"   # ⟨  MATHEMATICAL LEFT ANGLE BRACKET
_TAG_CLOSE = "\u27E9"  # ⟩  MATHEMATICAL RIGHT ANGLE BRACKET

# BUG FIX #1 — Google Translate removes the ⟩ after the tag name (e.g. ⟨r0⟩ → ⟨r0 ).
# The old strict regex ⟨r(\d+)⟩ ... ⟨/r\1⟩ never matched the corrupted form.
# New regex:
#   \u27E9?  — make the opening ⟩ optional
#    ?       — absorb the single extra space Google inserts in place of the ⟩
#   (.*?)    — capture run content non-greedily (preserves meaningful trailing
#              whitespace that was part of the original run text)
#   \u27E9?  — make the closing ⟩ optional
# Using a single optional space (not \s+) avoids eating meaningful leading
# spaces that are part of run content, while still handling the corruption.
# The regex still works perfectly on properly-formed tags.
_TAG_RE = re.compile(r'\u27E8r(\d+)\u27E9? ?(.*?)\u27E8/r\1\u27E9?', re.DOTALL)

# Pattern used to strip any ⟨…⟩ tag artefacts that survive after a failed parse
# (e.g. when the translator mangles the tag structure beyond what the lenient
# regex can recover).
_TAG_ARTIFACT_RE = re.compile(
    r'\u27E8/?r\d+\u27E9?'   # ⟨rN⟩, ⟨/rN⟩, or corruption variants missing ⟩
    r'|\u27E9'                # stray ⟩ left over
    r'|\u27E8'                # stray ⟨ left over
)


def _strip_tag_artifacts(text: str) -> str:
    """Remove any leftover ⟨rN⟩ / ⟨/rN⟩ tag markers from translated text."""
    return _TAG_ARTIFACT_RE.sub('', text)


def _build_tagged_paragraph(runs: list) -> tuple[str, bool]:
    """Build tagged or plain text for a paragraph's runs.

    Returns (text, is_tagged).
    - Uniform formatting or ≤1 text run → plain joined text, is_tagged=False.
    - Mixed formatting → ``⟨rN⟩text⟨/rN⟩`` per content run, is_tagged=True.
    """
    text_runs = [(i, run) for i, run in enumerate(runs) if run.text and run.text.strip()]

    if len(text_runs) <= 1 or _all_same_format([r for _, r in text_runs]):
        return _paragraph_full_text(runs), False

    parts: list[str] = []
    for i, run in text_runs:
        parts.append(
            f"{_TAG_OPEN}r{i}{_TAG_CLOSE}{run.text}{_TAG_OPEN}/r{i}{_TAG_CLOSE}"
        )
    return "".join(parts), True


def _parse_and_assign_tagged(runs: list, translated: str) -> bool:
    """Try to parse tagged translation result and assign to runs.

    Returns True if tags were found and applied, False if tags were
    mangled (caller should fall back to proportional redistribution).
    """
    matches = list(_TAG_RE.finditer(translated))
    if not matches:
        return False

    text_run_indices = {i for i, run in enumerate(runs) if run.text and run.text.strip()}
    assigned: dict[int, str] = {}
    for m in matches:
        idx = int(m.group(1))
        if 0 <= idx < len(runs):
            assigned[idx] = m.group(2)

    # Require at least half of text runs to be matched
    if len(assigned) < max(1, len(text_run_indices) // 2):
        return False

    for i, run in enumerate(runs):
        if i in assigned:
            run.text = assigned[i]
        elif i in text_run_indices:
            run.text = ""
        elif run.text and not run.text.strip():
            run.text = ""

    return True


def _redistribute_translated(runs, translated: str, *, is_tagged: bool = False) -> None:
    # Write *translated* back into *runs*, preserving formatting boundaries.
    # If the paragraph was tagged, try to parse tags first.
    if is_tagged:
        if _parse_and_assign_tagged(runs, translated):
            return
        # BUG FIX #1 (continued) — tags were too mangled for even the lenient
        # regex to recover.  Strip any leftover ⟨⟩ artefacts before falling
        # back to proportional redistribution so they don't appear in the
        # final document text.
        translated = _strip_tag_artifacts(translated)

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
    """
    An aggressive XML-based iterator that finds paragraphs in:
    1. The main body
    2. Tables (including nested tables)
    3. Text boxes (including those in shapes/groups)
    4. Headers/Footers
    5. Content Controls (SDTs)
    """
    seen: set[int] = set()

    # Define the tags we are looking for
    P_TAG = qn('w:p')
    TXBX_TAG = qn('w:txbxContent')
    SDT_TAG = qn('w:sdtContent')

    def _yield_from_element(parent_elt):
        for p_elem in parent_elt.iter(P_TAG):
            pid = id(p_elem)
            if pid not in seen:
                seen.add(pid)
                try:
                    yield _DocxParagraph(p_elem, doc.part)
                except Exception:
                    continue

    # 1. Process Main Body (including Tables and SDTs)
    yield from _yield_from_element(doc.element.body)

    # 2. Process Headers and Footers 
    for section in doc.sections:
        for hf_type in ['header', 'footer', 'first_page_header', 'first_page_footer', 'even_page_header', 'even_page_footer']:
            hf = getattr(section, hf_type, None)
            if hf and not (hasattr(hf, 'is_linked_to_previous') and hf.is_linked_to_previous):
                yield from _yield_from_element(hf._element)

    # 3. Deep search for Text Boxes/Shapes that might be outside the standard flow
    # This catches text boxes wrapped in VML or DrawingML tags
    for txbx in doc.element.iter(TXBX_TAG):
        yield from _yield_from_element(txbx)

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



# Paragraph grouping for contextual translation

# Separator used to group consecutive paragraphs into a single translation unit.
_PARA_SEP = "\n\u27EASEP\u27EB\n"

# Maximum paragraphs per group and maximum characters per group.
_GROUP_MAX_PARAS = 10
_GROUP_MAX_CHARS = 4000


def _is_heading(para) -> bool:
    """Return True if *para* is a heading style (acts as a group boundary)."""
    try:
        style_name = (para.style.name or "").lower()
        return style_name.startswith("heading") or style_name.startswith("title")
    except Exception:
        return False


def _group_paragraphs(
    para_infos: list[tuple],
    para_texts: list[str],
) -> list[list[int]]:
    """Return groups of paragraph indices for contextual translation.

    Rules:
    - Headings are always their own group (boundary).
    - Consecutive non-heading paragraphs are grouped together.
    - Groups are capped at _GROUP_MAX_PARAS items or _GROUP_MAX_CHARS total chars.
    """
    groups: list[list[int]] = []
    current: list[int] = []
    current_chars = 0

    for i, (para, runs, full_text, is_tagged) in enumerate(para_infos):
        text_len = len(para_texts[i])

        # Heading → flush current group, then make heading its own group
        if _is_heading(para):
            if current:
                groups.append(current)
                current = []
                current_chars = 0
            groups.append([i])
            continue

        # Would exceed limits → flush current group first
        if current and (
            len(current) >= _GROUP_MAX_PARAS
            or current_chars + text_len > _GROUP_MAX_CHARS
        ):
            groups.append(current)
            current = []
            current_chars = 0

        current.append(i)
        current_chars += text_len

    if current:
        groups.append(current)

    return groups


# Main entry point

def translate_docx(input_path: str, output_path: str, translator: Any, target_lang: str, *, cancel_event: threading.Event | None = None, source_lang: str = "auto"):
    # Batch-translate DOCX with contextual paragraph grouping and formatting preservation.
    # When source_lang==="auto" (multi-language mode) grouping is skipped so each paragraph
    # is sent individually and the API can auto-detect per paragraph.
    doc = Document(input_path)

    # --- Collect paragraph texts (one entry per paragraph) ---
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
        full_text, is_tagged = _build_tagged_paragraph(runs)
        if not full_text or not full_text.strip():
            continue
        para_infos.append((para, runs, full_text, is_tagged))
        para_texts.append(full_text)

    if not para_texts:
        doc.save(output_path)
        return

    # --- Group paragraphs for contextual translation ---
    # In auto-detect mode do NOT group paragraphs — each paragraph is its own unit
    # so the translation API sees one language at a time and can auto-detect correctly.
    if source_lang == "auto":
        groups = [[i] for i in range(len(para_infos))]
    else:
        groups = _group_paragraphs(para_infos, para_texts)

    # Build translation units: join grouped paragraphs with separator
    units: list[str] = []
    for group in groups:
        if len(group) == 1:
            units.append(para_texts[group[0]])
        else:
            units.append(_PARA_SEP.join(para_texts[idx] for idx in group))

    # --- Batch-translate all units ---
    try:
        translated_units = translator.translate_batch(units, target_lang, cancel_event=cancel_event)
    except CancelledError:
        raise
    except Exception:
        logger.exception("Batch translation failed for DOCX; falling back to per-item")
        translated_units = []
        for t in units:
            try:
                r = translator.translate_text(t, target_lang)
                translated_units.append(r if r is not None else t)
            except Exception:
                logger.exception("Per-item fallback also failed")
                translated_units.append(t)

    # --- Split grouped results and write back into runs ---
    errors = 0
    for group, tr_unit in zip(groups, translated_units):
        if tr_unit is None:
            errors += 1
            continue

        if len(group) == 1:
            # Single paragraph — direct assignment
            para, runs, orig_text, is_tagged = para_infos[group[0]]
            _redistribute_translated(runs, tr_unit, is_tagged=is_tagged)
        else:
            # Split on separator to recover per-paragraph translations
            parts = tr_unit.split(_PARA_SEP)
            if len(parts) == len(group):
                for idx, part in zip(group, parts):
                    para, runs, orig_text, is_tagged = para_infos[idx]
                    _redistribute_translated(runs, part.strip() if part else orig_text, is_tagged=is_tagged)
            else:
                # Separator mangled — fall back to individual paragraph translation
                logger.debug(
                    "DOCX paragraph separator mismatch: expected %d, got %d; per-para fallback",
                    len(group), len(parts),
                )
                for idx in group:
                    para, runs, orig_text, is_tagged = para_infos[idx]
                    try:
                        r = translator.translate_text(orig_text, target_lang)
                        _redistribute_translated(runs, r if r is not None else orig_text, is_tagged=is_tagged)
                    except Exception:
                        logger.exception("Per-paragraph fallback failed")
                        errors += 1

    # Check for cancellation before saving
    if cancel_event is not None and cancel_event.is_set():
        raise CancelledError("Translation cancelled before saving DOCX")

    doc.save(output_path)
    if errors:
        raise RuntimeError(f"Translation completed with {errors} failed paragraphs")