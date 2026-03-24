from dataclasses import dataclass, field
from typing import Any, Callable
import logging
import re
import threading

from lxml import etree
from docx.api import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph as _DocxParagraph

from ..utils import CancelledError

logger = logging.getLogger(__name__)


# ── XML namespace URIs (for lxml-level access to non-python-docx content) ──────
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_VML = "urn:schemas-microsoft-com:vml"
_REL_FOOTNOTES = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
)
_REL_ENDNOTES = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
)
_DIAGRAM_DATA_CT = (
    "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"
)


@dataclass
class _TranslationUnit:
    """Uniform container for a text segment to be translated, regardless of origin."""
    source_text: str
    is_tagged: bool = False
    is_heading: bool = False
    write_back: Callable[[str], None] = field(default=lambda t: None, repr=False)


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


# ── Ancestor helper ────────────────────────────────────────────────────────────

def _has_ancestor_tag(elem, tag) -> bool:
    """Return True if *elem* has an ancestor element with the given *tag*."""
    parent = elem.getparent()
    while parent is not None:
        if parent.tag == tag:
            return True
        parent = parent.getparent()
    return False


# ── Content Collectors ─────────────────────────────────────────────────────────
# Each collector returns a list of _TranslationUnit objects.  Collectors that
# access non-XmlPart OPC parts also return a list of (Part, lxml_root) "dirty"
# pairs whose blobs must be serialised back before Document.save().

def _collect_paragraph_units(doc) -> list[_TranslationUnit]:
    """Collect TranslationUnits from standard python-docx paragraphs."""
    units: list[_TranslationUnit] = []
    for para in _iter_all_paragraphs(doc):
        all_runs = list(para.runs)
        if not all_runs:
            continue
        runs = _get_translatable_runs(para, all_runs)
        if not runs:
            continue
        full_text, is_tagged = _build_tagged_paragraph(runs)
        if not full_text or not full_text.strip():
            continue

        def _make_wb(r, t):
            return lambda translated: _redistribute_translated(r, translated, is_tagged=t)

        units.append(_TranslationUnit(
            source_text=full_text,
            is_tagged=is_tagged,
            is_heading=_is_heading(para),
            write_back=_make_wb(runs, is_tagged),
        ))
    return units


def _collect_footnote_endnote_units(doc):
    """Collect TranslationUnits from footnotes and endnotes.

    Returns ``(units, dirty_parts)``; *dirty_parts* is a list of
    ``(Part, lxml_root)`` tuples whose blobs must be serialised before save
    when the part is a plain OPC Part (not an XmlPart with a live element tree).
    """
    units: list[_TranslationUnit] = []
    dirty_parts: list[tuple] = []

    for rel_type in (_REL_FOOTNOTES, _REL_ENDNOTES):
        try:
            part = doc.part.part_related_by(rel_type)
        except (KeyError, ValueError):
            continue

        # Live XML tree (XmlPart) or parsed from blob (plain Part)
        root = getattr(part, '_element', None)
        if root is None:
            blob = getattr(part, 'blob', None)
            if not blob:
                continue
            root = etree.fromstring(blob)
            dirty_parts.append((part, root))

        W_P = qn('w:p')
        W_ID = qn('w:id')

        for note_elem in root:
            note_id = note_elem.get(W_ID)
            if note_id in ('0', '1', '-1'):
                continue  # separator / continuation separator

            for p_elem in note_elem.iter(W_P):
                try:
                    para = _DocxParagraph(p_elem, doc.part)
                except Exception:
                    continue

                all_runs = list(para.runs)
                if not all_runs:
                    continue
                full_text, is_tagged = _build_tagged_paragraph(all_runs)
                if not full_text or not full_text.strip():
                    continue

                def _make_wb(r, t):
                    return lambda translated: _redistribute_translated(
                        r, translated, is_tagged=t,
                    )

                units.append(_TranslationUnit(
                    source_text=full_text,
                    is_tagged=is_tagged,
                    write_back=_make_wb(all_runs, is_tagged),
                ))

    return units, dirty_parts


def _collect_drawingml_units(doc) -> list[_TranslationUnit]:
    """Collect TranslationUnits from DrawingML ``<a:p>`` paragraphs.

    Skips any ``<a:p>`` inside ``<w:txbxContent>`` (already handled by
    the python-docx paragraph pipeline).
    """
    units: list[_TranslationUnit] = []
    A_P_TAG = f'{{{_NS_A}}}p'
    A_T_TAG = f'{{{_NS_A}}}t'
    TXBX_TAG = qn('w:txbxContent')

    seen: set[int] = set()

    search_roots = [doc.element]
    for section in doc.sections:
        for hf_type in ('header', 'footer', 'first_page_header', 'first_page_footer',
                         'even_page_header', 'even_page_footer'):
            hf = getattr(section, hf_type, None)
            if hf and not (hasattr(hf, 'is_linked_to_previous') and hf.is_linked_to_previous):
                search_roots.append(hf._element)

    for root_elem in search_roots:
        for ap_elem in root_elem.iter(A_P_TAG):
            ap_id = id(ap_elem)
            if ap_id in seen:
                continue
            seen.add(ap_id)

            if _has_ancestor_tag(ap_elem, TXBX_TAG):
                continue

            all_at = list(ap_elem.iter(A_T_TAG))
            if not all_at:
                continue
            full_text = "".join((e.text or "") for e in all_at)
            if not full_text.strip():
                continue

            text_at = [e for e in all_at if e.text and e.text.strip()]
            orig_lens = [max(len((e.text or "").strip()), 1) for e in text_at]

            def _make_wb(t_elems, a_elems, olens):
                def wb(translated):
                    if len(t_elems) <= 1:
                        target = t_elems[0] if t_elems else (a_elems[0] if a_elems else None)
                        if target is not None:
                            target.text = translated
                        for e in a_elems:
                            if e is not target:
                                e.text = ""
                    else:
                        words = translated.split()
                        if not words:
                            for e in t_elems:
                                e.text = ""
                            return
                        total = sum(olens)
                        assigned = 0
                        for j, e in enumerate(t_elems):
                            if j == len(t_elems) - 1:
                                e.text = " ".join(words[assigned:])
                            else:
                                prop = olens[j] / total
                                n = max(1, round(prop * len(words)))
                                end = min(assigned + n, len(words))
                                e.text = " ".join(words[assigned:end])
                                if end < len(words):
                                    e.text += " "
                                assigned = end
                        for e in a_elems:
                            if e not in t_elems:
                                e.text = ""
                return wb

            units.append(_TranslationUnit(
                source_text=full_text,
                write_back=_make_wb(text_at, all_at, orig_lens),
            ))

    return units


def _collect_smartart_units(doc):
    """Collect TranslationUnits from SmartArt diagram data parts.

    Returns ``(units, dirty_parts)``.
    """
    units: list[_TranslationUnit] = []
    dirty_parts: list[tuple] = []
    A_T_TAG = f'{{{_NS_A}}}t'

    try:
        parts_iter = doc.part.package.iter_parts()
    except Exception:
        return units, dirty_parts

    for part in parts_iter:
        ct = getattr(part, 'content_type', '')
        if ct != _DIAGRAM_DATA_CT:
            continue

        root = getattr(part, '_element', None)
        if root is None:
            blob = getattr(part, 'blob', None)
            if not blob:
                continue
            root = etree.fromstring(blob)
            dirty_parts.append((part, root))

        for at_elem in root.iter(A_T_TAG):
            text = at_elem.text
            if not text or not text.strip():
                continue

            def _make_wb(e):
                def wb(translated):
                    e.text = translated
                return wb

            units.append(_TranslationUnit(
                source_text=text,
                write_back=_make_wb(at_elem),
            ))

    return units, dirty_parts


def _collect_vml_units(doc) -> list[_TranslationUnit]:
    """Collect TranslationUnits from legacy VML ``<v:textpath>`` elements.

    Skips elements inside ``<w:txbxContent>`` (already handled by the
    python-docx paragraph pipeline).
    """
    units: list[_TranslationUnit] = []
    TEXTPATH_TAG = f'{{{_NS_VML}}}textpath'
    TXBX_TAG = qn('w:txbxContent')

    seen: set[int] = set()

    search_roots = [doc.element]
    for section in doc.sections:
        for hf_type in ('header', 'footer', 'first_page_header', 'first_page_footer',
                         'even_page_header', 'even_page_footer'):
            hf = getattr(section, hf_type, None)
            if hf and not (hasattr(hf, 'is_linked_to_previous') and hf.is_linked_to_previous):
                search_roots.append(hf._element)

    for root_elem in search_roots:
        for tp_elem in root_elem.iter(TEXTPATH_TAG):
            tp_id = id(tp_elem)
            if tp_id in seen:
                continue
            seen.add(tp_id)

            if _has_ancestor_tag(tp_elem, TXBX_TAG):
                continue

            text = tp_elem.get('string', '')
            if not text or not text.strip():
                continue

            def _make_wb(e):
                def wb(translated):
                    e.set('string', translated)
                return wb

            units.append(_TranslationUnit(
                source_text=text,
                write_back=_make_wb(tp_elem),
            ))

    return units



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


def _group_units(
    units: list[_TranslationUnit],
    *,
    auto_detect: bool,
) -> list[list[int]]:
    """Group TranslationUnit indices for contextual translation.

    Same rules as ``_group_paragraphs`` but operates on ``_TranslationUnit``
    objects instead of raw tuples.
    """
    if auto_detect:
        return [[i] for i in range(len(units))]

    groups: list[list[int]] = []
    current: list[int] = []
    current_chars = 0

    for i, unit in enumerate(units):
        text_len = len(unit.source_text)

        if unit.is_heading:
            if current:
                groups.append(current)
                current = []
                current_chars = 0
            groups.append([i])
            continue

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


def _translate_and_writeback(
    units: list[_TranslationUnit],
    groups: list[list[int]],
    translator,
    target_lang: str,
    cancel_event: threading.Event | None,
) -> int:
    """Translate grouped units via *translator* and invoke each unit's write_back.

    Returns the number of failed paragraphs.
    """
    # Build translation strings (join grouped units with separator)
    batch: list[str] = []
    for group in groups:
        if len(group) == 1:
            batch.append(units[group[0]].source_text)
        else:
            batch.append(_PARA_SEP.join(units[idx].source_text for idx in group))

    # Batch translate
    try:
        translated = translator.translate_batch(
            batch, target_lang, cancel_event=cancel_event,
        )
    except CancelledError:
        raise
    except Exception:
        logger.exception("Batch translation failed; falling back to per-item")
        translated = []
        for t in batch:
            try:
                r = translator.translate_text(t, target_lang)
                translated.append(r if r is not None else t)
            except Exception:
                logger.exception("Per-item fallback also failed")
                translated.append(t)

    # Write back translated text into the document
    errors = 0
    for group, tr_text in zip(groups, translated):
        if tr_text is None:
            errors += 1
            continue

        if len(group) == 1:
            units[group[0]].write_back(tr_text)
        else:
            parts = tr_text.split(_PARA_SEP)
            if len(parts) == len(group):
                for idx, part in zip(group, parts):
                    units[idx].write_back(
                        part.strip() if part else units[idx].source_text,
                    )
            else:
                logger.debug(
                    "Separator mismatch: expected %d, got %d; per-unit fallback",
                    len(group), len(parts),
                )
                for idx in group:
                    try:
                        r = translator.translate_text(
                            units[idx].source_text, target_lang,
                        )
                        units[idx].write_back(
                            r if r is not None else units[idx].source_text,
                        )
                    except Exception:
                        logger.exception("Per-unit fallback failed")
                        errors += 1

    return errors


# Main entry point

def translate_docx(input_path: str, output_path: str, translator: Any, target_lang: str, *, cancel_event: threading.Event | None = None, source_lang: str = "auto"):
    """Batch-translate DOCX with formatting preservation.

    Extracts text from paragraphs (incl. tables, headers/footers, text boxes),
    footnotes/endnotes, DrawingML shapes, SmartArt diagrams, and legacy VML.
    When ``source_lang=="auto"`` grouping is skipped so each unit is sent
    individually and the API can auto-detect per unit.
    """
    doc = Document(input_path)
    auto_detect = source_lang == "auto"

    # ── Phase 1: python-docx paragraphs (body, tables, headers, text boxes) ──
    pdocx_units = _collect_paragraph_units(doc)

    # ── Phases 2–5: lxml-level content (isolated error handling) ─────────────
    lxml_pools: list[list[_TranslationUnit]] = []
    dirty_parts: list[tuple] = []  # (Part, lxml_root) to serialise before save

    for collector_name, collector_fn in (
        ("footnotes/endnotes", lambda: _collect_footnote_endnote_units(doc)),
        ("SmartArt", lambda: _collect_smartart_units(doc)),
    ):
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError("Translation cancelled during extraction")
        try:
            pool, dp = collector_fn()
            if pool:
                lxml_pools.append(pool)
            dirty_parts.extend(dp)
        except CancelledError:
            raise
        except Exception:
            logger.warning("Failed to extract %s; skipping", collector_name, exc_info=True)

    for collector_name, collector_fn in (
        ("DrawingML", lambda: _collect_drawingml_units(doc)),
        ("VML", lambda: _collect_vml_units(doc)),
    ):
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError("Translation cancelled during extraction")
        try:
            pool = collector_fn()
            if pool:
                lxml_pools.append(pool)
        except CancelledError:
            raise
        except Exception:
            logger.warning("Failed to extract %s; skipping", collector_name, exc_info=True)

    # ── Nothing to translate → save unchanged ────────────────────────────────
    if not pdocx_units and not lxml_pools:
        doc.save(output_path)
        return

    errors = 0

    # ── Translate python-docx paragraph units ────────────────────────────────
    if pdocx_units:
        groups = _group_units(pdocx_units, auto_detect=auto_detect)
        errors += _translate_and_writeback(
            pdocx_units, groups, translator, target_lang, cancel_event,
        )

    # ── Translate each lxml pool separately ──────────────────────────────────
    for pool in lxml_pools:
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError("Translation cancelled between pools")
        groups = _group_units(pool, auto_detect=auto_detect)
        errors += _translate_and_writeback(
            pool, groups, translator, target_lang, cancel_event,
        )

    # ── Serialise modified non-XmlPart blobs back to their OPC parts ─────────
    for part, root in dirty_parts:
        part._blob = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True,
        )

    # ── Final cancellation check before saving ───────────────────────────────
    if cancel_event is not None and cancel_event.is_set():
        raise CancelledError("Translation cancelled before saving DOCX")

    doc.save(output_path)
    if errors:
        raise RuntimeError(f"Translation completed with {errors} failed paragraphs")