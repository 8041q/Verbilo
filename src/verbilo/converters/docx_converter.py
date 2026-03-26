from __future__ import annotations

import copy
import logging
import shutil
import threading
import zipfile
import io
from dataclasses import dataclass, field
from typing import Any, Callable

from lxml import etree
from docx.oxml.ns import qn

from ..utils import CancelledError

logger = logging.getLogger(__name__)


# XML namespace constants

_NS_W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_NS_A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_VML = "urn:schemas-microsoft-com:vml"
_DIAGRAM_DATA_CT = (
    "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"
)

_WT_TAG      = qn('w:t')
_WTAB_TAG    = qn('w:tab')
_WVAL_ATTR   = qn('w:val')

# Tags that should be treated as opaque containers — text inside these is
# handled by the SDT collector or skipped intentionally.
_SDT_TAG     = qn('w:sdt')
_SDT_PR_TAG  = qn('w:sdtPr')
_SDT_CONTENT = qn('w:sdtContent')

# Tracked-change insertion wrapper — we translate text inside <w:ins> normally.
# Deleted text (<w:del> / <w:delText>) is intentionally skipped.
_WINS_TAG    = qn('w:ins')
_WDEL_TAG    = qn('w:del')

# XML parts inside the zip that carry translatable body text
_BODY_PARTS = [
    "word/document.xml",
]
# These parts are discovered dynamically from relationships, but we also check
# well-known paths as a fallback.
_FOOTNOTE_PARTS    = ["word/footnotes.xml"]
_ENDNOTE_PARTS     = ["word/endnotes.xml"]
_HEADER_GLOB       = "word/header"
_FOOTER_GLOB       = "word/footer"


# Shared translation-unit container

@dataclass
class _TranslationUnit:
    # Uniform container for a text segment to be translated
    source_text: str
    is_heading: bool = False
    write_back: Callable[[str], None] = field(default=lambda t: None, repr=False)


# Grouping + batch translate

_PARA_SEP = "\n\u27EASEP\u27EB\n"
_GROUP_MAX_UNITS = 30
_GROUP_MAX_CHARS = 4000


def _group_units(
    units: list[_TranslationUnit],
    *,
    auto_detect: bool,
) -> list[list[int]]:
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
            len(current) >= _GROUP_MAX_UNITS
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
    translator: Any,
    target_lang: str,
    cancel_event: threading.Event | None,
) -> int:
    #  Translate each group and call write_back on each unit. Returns error count. Uses ``translate_batch`` when available, falling back to ``translate_text`` for single-item groups
    
    has_batch = callable(getattr(translator, 'translate_batch', None))
    has_text  = callable(getattr(translator, 'translate_text', None))

    if not has_batch and not has_text:
        raise AttributeError(
            f"{type(translator).__name__} exposes neither translate_batch nor "
            "translate_text — cannot translate."
        )

    errors = 0
    for group in groups:
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError("Translation cancelled")

        texts = [units[i].source_text for i in group]

        # ── Batch path (preferred) ────────────────────────────────────────────
        if has_batch and len(group) > 1:
            try:
                translated_parts = translator.translate_batch(
                    texts, target_lang, cancel_event=cancel_event
                )
                if len(translated_parts) != len(texts):
                    raise ValueError("translate_batch returned wrong number of results")
                for idx, i in enumerate(group):
                    units[i].write_back(translated_parts[idx])
                continue
            except CancelledError:
                raise
            except Exception:
                logger.warning(
                    "translate_batch failed for group %s; falling back to translate_text",
                    group, exc_info=True,
                )
                # fall through to per-item path below

        # ── Per-item path (single items, or batch fallback) ───────────────────
        translate_fn = (
            translator.translate_text if has_text
            else lambda t, tl: translator.translate_batch([t], tl)[0]
        )
        for i in group:
            try:
                result = translate_fn(units[i].source_text, target_lang)
                units[i].write_back(result)
            except CancelledError:
                raise
            except Exception:
                logger.warning("Failed to translate unit %d", i, exc_info=True)
                errors += 1

    return errors


# Direct XML collectors — operate on lxml elements in-place

def _is_in_heading(elem: etree._Element) -> bool:
    # Return True if *elem* is inside a heading paragraph (w:pStyle starts with 'Heading')
    # Walk up to find w:p, then check w:pStyle
    node = elem
    while node is not None:
        if node.tag == qn('w:p'):
            ppr = node.find(qn('w:pPr'))
            if ppr is not None:
                pstyle = ppr.find(qn('w:pStyle'))
                if pstyle is not None:
                    val = pstyle.get(qn('w:val'), '')
                    if val.lower().startswith('heading'):
                        return True
            return False
        node = node.getparent()
    return False


def _is_inside_del(elem: etree._Element) -> bool:
    """Return True if *elem* is nested inside a ``<w:del>`` tracked-change block.

    Text inside ``<w:del>`` represents deleted content and must NOT be
    translated — it would corrupt the tracked-changes record.
    """
    node = elem.getparent()
    while node is not None:
        if node.tag == _WDEL_TAG:
            return True
        node = node.getparent()
    return False


def _is_inside_drawing_run(elem: etree._Element) -> bool:
    """Return True if *elem* is a ``<w:t>`` inside a ``<w:r>`` that also contains
    a ``<w:drawing>``.

    Word allows only one type of content per run — a run that holds a drawing
    element must not also carry translatable text.  When a source document has
    such a structure the text node is an artefact (e.g. leftover Chinese label
    that survived the drawing-group wrapper); translating it causes the result
    string to be rendered *on top of* the floating image instead of beside it.
    """
    _W_R       = qn('w:r')
    _W_DRAWING = qn('w:drawing')
    node = elem.getparent()
    while node is not None:
        if node.tag == _W_R:
            return node.find(_W_DRAWING) is not None
        node = node.getparent()
    return False


import re as _re

_RE_PURE_NUMBER = _re.compile(
    r"""
    ^[\s\u00a0]*           # optional leading whitespace / NBSP
    [+-]?                  # optional sign
    (?:
        \d{1,3}(?:[.,\s]\d{3})*  # thousands-grouped integer e.g. 1,234
        |\d+                      # plain integer
    )
    (?:[.,]\d+)?           # optional decimal part
    [\s\u00a0]*            # optional trailing whitespace
    (?:%|°|㎡|㎞|km|cm|mm|m²|m³|€|\$|£|¥|₹|元|円|₩)?  # optional unit
    [\s\u00a0]*$
    """,
    _re.VERBOSE,
)

# Matches strings that are MOSTLY digits — the text contains digits but also
# CJK/Latin characters that give the translator context to spell out the number.
# We extract and protect the numeric portion instead of skipping the whole unit.
_RE_HAS_DIGIT = _re.compile(r'\d')


def _is_numeric_only(text: str) -> bool:
    """Return True if *text* is a standalone number (possibly with units).

    These are never sent to the translator to prevent silent conversion of
    "42" → "forty-two" or decimal-separator reformatting.
    """
    return bool(_RE_PURE_NUMBER.match(text))


def _strip_protected_tokens(text: str) -> tuple[str, dict[str, str]]:
    """Replace numeric tokens in *text* with placeholders before translation.

    Returns (masked_text, placeholder_map).  The caller should substitute
    placeholders back after translation.

    We protect:
      - standalone integers and decimals (possibly thousands-grouped)
      - years expressed as 4-digit sequences
      - version strings like "3.14", "v2.0"
      - sequences that are >= 50% digits by character count
    """
    placeholders: dict[str, str] = {}
    counter = [0]

    def _replace(m: _re.Match) -> str:
        # Use a bracket format that is XML-safe, unlikely to appear in real text,
        # and opaque enough that translators treat it as a token to preserve.
        key = f'[[N{counter[0]}]]'
        placeholders[key] = m.group(0)
        counter[0] += 1
        return key

    # Protect numeric tokens: integers, decimals, version numbers, years
    masked = _re.sub(
        r'(?<!\w)(\d[\d,.\s]*\d|\d)(?!\w)',
        _replace,
        text,
    )
    return masked, placeholders


def _restore_protected_tokens(text: str, placeholders: dict[str, str]) -> str:
    """Substitute placeholders back into *text* after translation."""
    for key, original in placeholders.items():
        text = text.replace(key, original)
    return text


# ── Layout-protection passes ─────────────────────────────────────────────────
# Run these on each XML root AFTER translation so that text-expansion caused
# by longer translated strings doesn't break the document layout.

_WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


def _fix_table_row_heights(root: etree._Element) -> None:
    """Convert ``exact`` row-height constraints to ``atLeast``.

    When a table row uses ``w:trHeight w:hRule="exact"``, Word will clip any
    text that overflows the fixed height rather than expanding the row.  After
    translation the text is usually longer, so clipping becomes visible.
    Changing the rule to ``atLeast`` lets the row grow to fit its content while
    still respecting the designer's minimum-height intent.
    """
    TR_HEIGHT = qn('w:trHeight')
    H_RULE    = qn('w:hRule')
    for trh in root.iter(TR_HEIGHT):
        if trh.get(H_RULE, '').lower() == 'exact':
            trh.set(H_RULE, 'atLeast')
            logger.debug("Relaxed exact row height → atLeast")


def _fix_textbox_autofit(root: etree._Element) -> None:
    """Enable shape auto-fit (``spAutoFit``) on text-box body properties where
    the translated text might overflow the original fixed size.

    Handles both ``<a:bodyPr>`` (DrawingML shapes / SmartArt) and
    ``<wps:bodyPr>`` (modern Word text boxes via the wordprocessingShape
    namespace).  Each can carry one of three resize-mode children:

    * ``<a:noAutofit/>``  — clip text to the fixed box size  ← fixed ✗
    * ``<a:normAutofit/>`` — shrink font to fit the box        ← shrinks text ✗
    * ``<a:spAutoFit/>``  — grow the shape to fit the text    ← correct ✓
    * (no child)           — Word defaults to clip behaviour  ← also fixed ✗

    We replace ``noAutofit`` / ``normAutofit`` with ``spAutoFit`` and inject
    ``spAutoFit`` where no child exists, so the text box always expands to
    accommodate the (usually longer) translated text.
    """
    _NS_WPS      = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    A_BODY_PR    = f'{{{_NS_A}}}bodyPr'
    WPS_BODY_PR  = f'{{{_NS_WPS}}}bodyPr'
    A_NO_AUTOFIT = f'{{{_NS_A}}}noAutofit'
    A_NORM_AUTO  = f'{{{_NS_A}}}normAutofit'
    A_SP_AUTO    = f'{{{_NS_A}}}spAutoFit'
    _AUTOFIT_TAGS = {A_NO_AUTOFIT, A_NORM_AUTO, A_SP_AUTO}

    for body_pr in [*root.iter(A_BODY_PR), *root.iter(WPS_BODY_PR)]:
        # Find any existing autofit child
        existing = next(
            (child for child in body_pr if child.tag in _AUTOFIT_TAGS), None
        )
        if existing is None:
            # No autofit child — Word silently clips.  Inject spAutoFit.
            body_pr.append(etree.Element(A_SP_AUTO))
            logger.debug("Injected spAutoFit into text-box bodyPr (was missing)")
        elif existing.tag == A_SP_AUTO:
            pass  # Already correct
        else:
            idx = list(body_pr).index(existing)
            body_pr.remove(existing)
            body_pr.insert(idx, etree.Element(A_SP_AUTO))
            logger.debug("Replaced %s → spAutoFit in text-box bodyPr", existing.tag)


def _fix_frame_autosize(root: etree._Element) -> None:
    """Ensure legacy word-processing frames (``<w:framePr>``) can grow.

    Frames created with ``<w:framePr w:h="..." w:hRule="exact">`` clip content
    exactly like table rows.  We relax them to ``atLeast`` for the same reason.
    """
    FRAME_PR = qn('w:framePr')
    H_RULE   = qn('w:hRule')
    for fp in root.iter(FRAME_PR):
        if fp.get(H_RULE, '').lower() == 'exact':
            fp.set(H_RULE, 'atLeast')
            logger.debug("Relaxed framePr exact height → atLeast")


def _fix_vml_textbox_autosize(root: etree._Element) -> None:
    """Allow legacy VML text boxes to grow to fit translated (longer) text.

    VML shapes use a ``style`` attribute on ``<v:shape>`` that carries
    CSS-like ``key:value`` pairs separated by semicolons.  The property
    ``mso-fit-shape-to-text:t`` tells Word to expand the shape height to fit
    its content.  If it is absent or set to ``f`` (false) the shape clips.

    We walk up from every ``<v:textbox>`` to its parent ``<v:shape>`` and
    set / insert ``mso-fit-shape-to-text:t``.
    """
    V_TEXTBOX = f'{{{_NS_VML}}}textbox'
    V_SHAPE   = f'{{{_NS_VML}}}shape'

    for textbox in root.iter(V_TEXTBOX):
        shape = textbox.getparent()
        if shape is None or shape.tag != V_SHAPE:
            continue

        style = shape.get('style', '')
        # Parse into an ordered list of (key, value) pairs
        parts = [p.strip() for p in style.split(';') if p.strip()]
        pairs: list[tuple[str, str]] = []
        for part in parts:
            if ':' in part:
                k, _, v = part.partition(':')
                pairs.append((k.strip(), v.strip()))
            else:
                pairs.append((part, ''))

        # Replace or append mso-fit-shape-to-text
        updated = False
        new_pairs: list[tuple[str, str]] = []
        for k, v in pairs:
            if k.lower() == 'mso-fit-shape-to-text':
                new_pairs.append((k, 't'))
                updated = True
            else:
                new_pairs.append((k, v))
        if not updated:
            new_pairs.append(('mso-fit-shape-to-text', 't'))

        new_style = ';'.join(
            f'{k}:{v}' if v else k for k, v in new_pairs
        )
        shape.set('style', new_style)
        logger.debug("Set mso-fit-shape-to-text:t on VML shape")


def _expand_textbox_widths(root: etree._Element, expansion_ratio: float = 1.0) -> None:
    """Widen label-sized DrawingML/WPS text boxes to reduce unwanted line-wrapping.

    For CJK→Latin translation a phrase like "使用说明书" (5 chars) becomes
    "User instructions" (17 chars).  Because CJK characters render at ~1 em
    width and Latin characters at ~0.55 em on average, the translated text
    needs roughly ``expansion_ratio × 0.45`` times the original box width.
    ``spAutoFit`` handles vertical overflow when text wraps, but widening the
    box first lets short labels stay on a single line.

    Only narrow boxes (≤ ~2.2 in / 2 000 000 EMU) are widened — these are
    labels and callouts.  Wide content boxes already have room and rely on
    ``spAutoFit`` for vertical growth.

    Shapes covered:
    * ``<wps:wsp>`` containing ``<wps:txbx>`` — modern Word text boxes
    * ``<a:sp>`` containing ``<a:txBody>``    — DrawingML shapes with text

    Both the shape extents and the enclosing ``<wp:anchor>``/``<wp:inline>``
    ``<wp:extent cx>`` are updated so Word does not clip at the container boundary.
    """
    _NS_WPS    = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    WPS_WSP    = f'{{{_NS_WPS}}}wsp'
    WPS_SPR    = f'{{{_NS_WPS}}}spPr'
    WPS_TXB    = f'{{{_NS_WPS}}}txbx'
    A_SP       = f'{{{_NS_A}}}sp'
    A_TXBOD    = f'{{{_NS_A}}}txBody'
    A_SP_PR    = f'{{{_NS_A}}}spPr'
    A_XFRM     = f'{{{_NS_A}}}xfrm'
    A_EXT      = f'{{{_NS_A}}}ext'
    WP_ANCHOR  = f'{{{_WP_NS}}}anchor'
    WP_INLINE  = f'{{{_WP_NS}}}inline'
    WP_EXT     = f'{{{_WP_NS}}}extent'

    # Only expand label-sized boxes; large content boxes grow vertically instead.
    LABEL_MAX_CX = 2_000_000   # ~2.2 inches in EMU
    MAX_INCREASE = 1_200_000   # ~1.3 inches maximum extra width

    # Width expansion factor.
    # At expansion_ratio=1 (no change) the factor stays at 1.0 (no expansion).
    # For CJK→Latin with typical ratio 2.5–3.5 this gives factor ~1.7–2.0,
    # i.e. up to double the original width for small labels.
    if expansion_ratio <= 1.0:
        return  # Nothing to expand
    width_factor = min(1.0 + (expansion_ratio - 1.0) * 0.45, 2.0)

    def _expand_cx(elem: etree._Element) -> None:
        try:
            cx = int(elem.get('cx', 0))
        except ValueError:
            return
        if cx <= 0 or cx > LABEL_MAX_CX:
            return  # Skip very large (content) boxes
        increase = min(int(cx * (width_factor - 1.0)), MAX_INCREASE)
        if increase > 0:
            elem.set('cx', str(cx + increase))

    def _container_extent(shape_elem: etree._Element) -> etree._Element | None:
        """Walk up to the enclosing wp:anchor / wp:inline and return its extent."""
        node = shape_elem.getparent()
        while node is not None:
            if node.tag in (WP_ANCHOR, WP_INLINE):
                return node.find(WP_EXT)
            node = node.getparent()
        return None

    # WPS text boxes: <wps:wsp> with <wps:txbx>
    for wsp in root.iter(WPS_WSP):
        if wsp.find(WPS_TXB) is None:
            continue
        spr = wsp.find(WPS_SPR)
        if spr is None:
            continue
        xfrm = spr.find(A_XFRM)
        if xfrm is None:
            continue
        ext = xfrm.find(A_EXT)
        if ext is not None:
            _expand_cx(ext)

    # DrawingML shapes with a text body: <a:sp> containing <a:txBody>
    for sp in root.iter(A_SP):
        if sp.find(A_TXBOD) is None:
            continue
        spr = sp.find(A_SP_PR)
        if spr is None:
            continue
        xfrm = spr.find(A_XFRM)
        if xfrm is None:
            continue
        ext = xfrm.find(A_EXT)
        if ext is not None:
            _expand_cx(ext)


def _fix_anchor_wrapping(root: etree._Element) -> None:
    """Convert explicit ``wrapNone`` on standalone floating images to
    ``wrapTopAndBottom`` so that expanded translated text does not overlap them.

    An anchor image with ``wrapNone`` stays at its absolute page position even
    as translated text expands below it, causing visible overlap.  Changing the
    wrap mode to ``wrapTopAndBottom`` keeps the image exactly where it is but
    tells Word to push text above and below the image instead of behind it.

    The following anchors are **left untouched** to avoid breaking diagram
    layouts (labels, arrows, callout groups):

    * ``behindDoc="1"`` — background / watermark / diagram-underlay images
      whose overlap with text is intentional.
    * ``<wp:positionV relativeFrom>`` is ``paragraph``, ``line``, or
      ``character`` — the image is vertically anchored to the text flow, not
      the page.  Changing the wrap mode for these elements disrupts the visual
      grouping of diagram components (images + labels + arrows that all anchor
      to the same paragraph).
    * Anchors inside a ``<w:tc>`` table cell — cell-positioned images use
      cell-relative coordinates.
    * Anchors with no explicit ``<wp:wrapNone/>`` child — absence of a wrap
      element is another indicator of a deliberately-composed diagram group.
    * Anchors that already carry any non-None wrap mode.
    """
    W_TC         = qn('w:tc')
    WP_ANCHOR    = f'{{{_WP_NS}}}anchor'
    WP_POS_V     = f'{{{_WP_NS}}}positionV'
    WP_WRAP_NONE = f'{{{_WP_NS}}}wrapNone'
    WP_WRAP_TAB  = f'{{{_WP_NS}}}wrapTopAndBottom'
    _EXISTING_WRAP_TAGS = {
        f'{{{_WP_NS}}}wrapSquare',
        f'{{{_WP_NS}}}wrapTight',
        f'{{{_WP_NS}}}wrapThrough',
        f'{{{_WP_NS}}}wrapTopAndBottom',
    }
    # These relativeFrom values mean the image is part of a diagram group tied
    # to a specific paragraph/line/character in the text flow.
    _TEXT_RELATIVE = {'paragraph', 'line', 'character'}

    for anchor in root.iter(WP_ANCHOR):
        # Skip background / watermark / diagram-underlay images
        if anchor.get('behindDoc', '0') == '1':
            continue

        # Skip diagram elements: vertical position tied to the text flow
        pos_v = anchor.find(WP_POS_V)
        if pos_v is not None and pos_v.get('relativeFrom', '') in _TEXT_RELATIVE:
            continue

        # Skip anchors inside table cells
        node = anchor.getparent()
        in_tc = False
        while node is not None:
            if node.tag == W_TC:
                in_tc = True
                break
            node = node.getparent()
        if in_tc:
            continue

        children = list(anchor)
        child_tags = {c.tag for c in children}
        if child_tags & _EXISTING_WRAP_TAGS:
            continue

        # Only convert explicit wrapNone — anchors with no wrap child at all
        # are diagram components with deliberate absolute positioning.
        wrap_none = anchor.find(WP_WRAP_NONE)
        if wrap_none is not None:
            idx = children.index(wrap_none)
            anchor.remove(wrap_none)
            anchor.insert(idx, etree.Element(WP_WRAP_TAB))
            logger.debug("Replaced wrapNone → wrapTopAndBottom on anchor")


def _is_spacer_paragraph(para: etree._Element) -> bool:
    """Return True if *para* is a layout-spacer paragraph with no visible text.

    Paragraphs that host ``<wp:anchor>`` elements with
    ``positionV relativeFrom="paragraph"`` (or "line" / "character") are
    explicitly excluded even when they carry no visible ``<w:t>`` text.
    Their line-height / before / after spacing is the vertical reference point
    for the anchored drawing group; compressing it shifts the floating diagram
    away from its surrounding labels and arrows.
    """
    for wt in para.iter(_WT_TAG):
        if wt.text and wt.text.strip():
            return False

    # Protect paragraphs that anchor floating drawings to the text flow.
    # These look like spacers (no visible text) but their spacing determines
    # where the anchored image is rendered on the page.
    _WP_ANCHOR = f'{{{_WP_NS}}}anchor'
    _WP_POS_V  = f'{{{_WP_NS}}}positionV'
    _TEXT_RELATIVE = {'paragraph', 'line', 'character'}
    for anchor in para.iter(_WP_ANCHOR):
        pos_v = anchor.find(_WP_POS_V)
        if pos_v is not None and pos_v.get('relativeFrom', '') in _TEXT_RELATIVE:
            return False

    return True


def _para_has_page_break(para: etree._Element) -> bool:
    """Return True if *para* starts a new page via pageBreakBefore or a w:br page break."""
    W_P     = qn('w:p')
    W_PPR   = qn('w:pPr')
    W_PBB   = qn('w:pageBreakBefore')
    W_BR    = qn('w:br')
    W_TYPE  = qn('w:type')

    # Check w:pPr/w:pageBreakBefore
    ppr = para.find(W_PPR)
    if ppr is not None:
        pbb = ppr.find(W_PBB)
        if pbb is not None:
            val = pbb.get(qn('w:val'), '1')
            if val.lower() not in ('0', 'false'):
                return True

    # Check for explicit w:br type="page" inside any run
    for br in para.iter(W_BR):
        if br.get(W_TYPE, '').lower() == 'page':
            return True

    return False


def _collect_page_sections(root: etree._Element) -> list[list[etree._Element]]:
    """Split all body paragraphs into page sections divided by hard page breaks.

    Returns a list of sections, each section being a list of ``<w:p>`` elements
    that belong to the same logical page.  The paragraph that carries the page
    break starts the new section.
    """
    W_BODY = qn('w:body')
    W_P    = qn('w:p')
    W_TBL  = qn('w:tbl')

    body = root.find(W_BODY)
    if body is None:
        body = root  # headers/footers have no w:body wrapper

    sections: list[list[etree._Element]] = [[]]
    for child in body:
        if child.tag == W_P:
            if _para_has_page_break(child) and sections[-1]:
                sections.append([])
            sections[-1].append(child)
        # Tables count as bulk but we don't split inside them
    return [s for s in sections if s]


def _compress_section_spacing(
    section: list[etree._Element],
    expansion_ratio: float,
) -> None:
    """Compress spacing within one page section proportionally to expansion_ratio.

    Strategy
    --------
    1. Spacer paragraphs (no text): compress ``w:before/after`` by
       ``1/expansion_ratio``, clamped to [0.25, 1.0].  This reclaims the
       vertical space that would otherwise push content onto the next page.
    2. Content paragraphs: tiered compression based on expansion ratio:
       - ratio ≤ 1.3  →  no compression
       - 1.3 < ratio ≤ 2.0  →  ``max(0.50, 1/ratio)``
       - ratio > 2.0  →  ``max(0.38, 1/ratio)`` (aggressive: typical CJK→Latin)
    3. ``w:lineRule="exact"`` is always relaxed to ``atLeast`` regardless of
       ratio, because Latin descenders clip under exact CJK-sized line heights.
    4. Never reduce a spacing value below 0.
    """
    if expansion_ratio <= 1.0:
        # Relax exact line spacing even when text didn't expand
        for para in section:
            _relax_exact_line_spacing(para)
        return

    spacer_compression  = max(0.25, 1.0 / expansion_ratio)
    # Tiered content compression: harder for very high CJK→Latin expansion ratios
    if expansion_ratio > 2.0:
        content_compression = max(0.38, 1.0 / expansion_ratio)
    elif expansion_ratio > 1.3:
        content_compression = max(0.50, 1.0 / expansion_ratio)
    else:
        content_compression = 1.0

    W_PPR     = qn('w:pPr')
    W_SPACING = qn('w:spacing')
    W_BEFORE  = qn('w:before')
    W_AFTER   = qn('w:after')

    for para in section:
        _relax_exact_line_spacing(para)
        ppr = para.find(W_PPR)
        if ppr is None:
            continue
        spacing = ppr.find(W_SPACING)
        if spacing is None:
            continue

        is_spacer = _is_spacer_paragraph(para)
        factor = spacer_compression if is_spacer else content_compression

        for attr in (W_BEFORE, W_AFTER):
            raw = spacing.get(attr)
            if raw is None:
                continue
            try:
                val = int(raw)
            except ValueError:
                continue
            spacing.set(attr, str(max(0, int(val * factor))))


def _relax_exact_line_spacing(para: etree._Element) -> None:
    """Change ``w:lineRule="exact"`` to ``atLeast`` on a single paragraph."""
    W_PPR      = qn('w:pPr')
    W_SPACING  = qn('w:spacing')
    W_LINERULE = qn('w:lineRule')

    ppr = para.find(W_PPR)
    if ppr is None:
        return
    spacing = ppr.find(W_SPACING)
    if spacing is None:
        return
    if spacing.get(W_LINERULE, '').lower() == 'exact':
        spacing.set(W_LINERULE, 'atLeast')


def _compress_spacer_spacing(root: etree._Element, expansion_ratio: float) -> None:
    """Compress spacing section-by-section, respecting hard page boundaries."""
    if expansion_ratio <= 1.0:
        return
    for section in _collect_page_sections(root):
        _compress_section_spacing(section, expansion_ratio)


def _compress_paragraph_spacing(root: etree._Element, expansion_ratio: float) -> None:
    """Relax exact line spacing on all paragraphs (handled inside section pass)."""
    # This is now a no-op — the section-aware pass in _compress_spacer_spacing
    # handles both spacer and content paragraphs together so we don't double-apply.
    pass


def _ensure_list_indentation(root: etree._Element) -> None:
    """Add standard Western indentation to list paragraphs that have none.

    CJK source documents often format bullet/numbered lists without explicit
    ``w:ind`` because CJK characters have fixed-width cells.  After translation
    to Latin script the list items appear flush-left.  We apply standard Office
    indentation (720 twips per indent level, 360-twip hanging) to any paragraph
    that carries ``w:numPr`` but lacks a meaningful ``w:left`` indent
    (< 360 twips / 0.25 inch).
    """
    W_P       = qn('w:p')
    W_PPR     = qn('w:pPr')
    W_NUMPR   = qn('w:numPr')
    W_ILVL    = qn('w:ilvl')
    W_IND     = qn('w:ind')
    W_LEFT    = qn('w:left')
    W_HANGING = qn('w:hanging')

    TWIPS_PER_LEVEL = 720   # 0.5 inch per indent level (standard Office default)
    HANGING         = 360   # 0.25 inch hanging indent for the bullet/number
    MIN_EXISTING    = 360   # Only act when current left indent < this threshold

    for para in root.iter(W_P):
        ppr = para.find(W_PPR)
        if ppr is None:
            continue
        numpr = ppr.find(W_NUMPR)
        if numpr is None:
            continue  # Not a list paragraph

        # Determine indent level (0-based); default to 0 if absent
        ilvl_elem = numpr.find(W_ILVL)
        ilvl = 0
        if ilvl_elem is not None:
            try:
                ilvl = int(ilvl_elem.get(_WVAL_ATTR, '0'))
            except ValueError:
                pass

        desired_left = (ilvl + 1) * TWIPS_PER_LEVEL

        ind = ppr.find(W_IND)
        if ind is not None:
            try:
                current_left = int(ind.get(W_LEFT, '0'))
            except ValueError:
                current_left = 0
            if current_left >= MIN_EXISTING:
                continue  # Already has a meaningful indent — leave it alone
            ind.set(W_LEFT, str(desired_left))
            ind.set(W_HANGING, str(HANGING))
        else:
            # No w:ind at all — create one and insert right after w:numPr
            ind = etree.Element(W_IND)
            ind.set(W_LEFT, str(desired_left))
            ind.set(W_HANGING, str(HANGING))
            numpr_idx = list(ppr).index(numpr)
            ppr.insert(numpr_idx + 1, ind)

        logger.debug(
            "Applied list indent: level=%d left=%d hanging=%d",
            ilvl, desired_left, HANGING,
        )


def _apply_layout_fixes(root: etree._Element, expansion_ratio: float = 1.0) -> None:
    """Run all layout-protection passes on *root* in one call.

    Parameters
    ----------
    expansion_ratio:
        Ratio of total translated characters to total source characters for
        this XML part.  Values > 1 mean the translation is longer (common for
        CJK→Latin).  Used to scale down spacing on spacer and content
        paragraphs so page layout is preserved.
    """
    _fix_table_row_heights(root)
    _fix_textbox_autofit(root)
    _fix_frame_autosize(root)
    _fix_vml_textbox_autosize(root)
    _expand_textbox_widths(root, expansion_ratio)
    _compress_spacer_spacing(root, expansion_ratio)
    _compress_paragraph_spacing(root, expansion_ratio)
    _ensure_list_indentation(root)


def _collect_wt_units(root: etree._Element) -> list[_TranslationUnit]:
    #  Collect TranslationUnits from all ``<w:t>`` elements under *root*.
    #
    #  Skips:
    #    - empty / whitespace-only text nodes
    #    - text inside ``<w:del>`` tracked-change blocks (deleted text must be
    #      kept verbatim to preserve the revision history)
    #
    #  This is the primary collector for body text, headers, footers,
    #  footnotes and endnotes.  It mutates elements in-place via write_back.

    units: list[_TranslationUnit] = []

    for wt in root.iter(_WT_TAG):
        text = wt.text
        if not text or not text.strip():
            continue

        # Skip deleted text — translating it would corrupt tracked changes.
        if _is_inside_del(wt):
            continue

        # Skip <w:t> nodes that live inside a run which also contains a
        # <w:drawing>.  Such runs hold a floating image; the <w:t> is an
        # artefact that Word ignores visually but which we must not overwrite
        # with translated text (doing so renders text on top of the image).
        if _is_inside_drawing_run(wt):
            continue

        # Skip pure numbers — translators convert "42" → "forty-two", etc.
        if _is_numeric_only(text):
            continue

        # For mixed text (e.g. "2023年"), mask numeric tokens so they survive
        # translation unchanged, then restore them in write_back.
        masked_text, placeholders = _strip_protected_tokens(text.strip())
        is_heading = _is_in_heading(wt)

        def _make_wb(elem: etree._Element, orig_text: str, ph: dict[str, str]):
            def wb(translated: str) -> None:
                # Restore any numeric placeholders the translator may have mangled
                result = _restore_protected_tokens(translated, ph)
                # Preserve leading/trailing whitespace from the original
                leading  = orig_text[: len(orig_text) - len(orig_text.lstrip())]
                trailing = orig_text[len(orig_text.rstrip()):]
                elem.text = leading + result.strip() + trailing
                if leading or trailing:
                    elem.set(
                        '{http://www.w3.org/XML/1998/namespace}space', 'preserve'
                    )
            return wb

        units.append(_TranslationUnit(
            source_text=masked_text,
            is_heading=is_heading,
            write_back=_make_wb(wt, text, placeholders),
        ))

    return units


def _collect_drawingml_units(root: etree._Element) -> list[_TranslationUnit]:
    # Collect TranslationUnits from DrawingML ``<a:t>`` elements
    units: list[_TranslationUnit] = []
    A_T_TAG = f'{{{_NS_A}}}t'

    for at_elem in root.iter(A_T_TAG):
        text = at_elem.text
        if not text or not text.strip():
            continue

        if _is_numeric_only(text):
            continue

        masked_text, placeholders = _strip_protected_tokens(text.strip())

        def _make_wb(e: etree._Element, orig_text: str, ph: dict[str, str]):
            def wb(translated: str) -> None:
                result = _restore_protected_tokens(translated, ph)
                leading  = orig_text[: len(orig_text) - len(orig_text.lstrip())]
                trailing = orig_text[len(orig_text.rstrip()):]
                e.text = leading + result.strip() + trailing
                if leading or trailing:
                    e.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            return wb

        units.append(_TranslationUnit(
            source_text=masked_text,
            write_back=_make_wb(at_elem, text, placeholders),
        ))

    return units


def _collect_vml_units(root: etree._Element) -> list[_TranslationUnit]:
    # Collect TranslationUnits from legacy VML ``<v:textpath>`` elements
    units: list[_TranslationUnit] = []
    TEXTPATH_TAG = f'{{{_NS_VML}}}textpath'

    for tp_elem in root.iter(TEXTPATH_TAG):
        text = tp_elem.get('string', '')
        if not text or not text.strip():
            continue

        def _make_wb(e: etree._Element):
            def wb(translated: str) -> None:
                e.set('string', translated)
            return wb

        units.append(_TranslationUnit(
            source_text=text,
            write_back=_make_wb(tp_elem),
        ))

    return units


def _collect_smartart_units_from_blob(
    blob: bytes,
) -> tuple[list[_TranslationUnit], etree._Element]:
    # Parse a SmartArt diagramData blob and return (units, root_element)
    root = etree.fromstring(blob)
    units: list[_TranslationUnit] = []
    A_T_TAG = f'{{{_NS_A}}}t'

    for at_elem in root.iter(A_T_TAG):
        text = at_elem.text
        if not text or not text.strip():
            continue

        def _make_wb(e: etree._Element, orig_text: str):
            def wb(translated: str) -> None:
                leading  = orig_text[: len(orig_text) - len(orig_text.lstrip())]
                trailing = orig_text[len(orig_text.rstrip()):]
                e.text = leading + translated.strip() + trailing
                if leading or trailing:
                    e.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            return wb

        units.append(_TranslationUnit(
            source_text=text.strip(),
            write_back=_make_wb(at_elem, text),
        ))

    return units, root


# Zip-level patching — the key to preserving all formatting

def _parse_xml_part(data: bytes) -> etree._Element:
    return etree.fromstring(data)


def _serialise_xml_part(root: etree._Element, original_data: bytes) -> bytes:
    """Serialise *root* back to bytes, preserving the original XML declaration."""
    # Detect encoding from original declaration (default UTF-8)
    encoding = 'UTF-8'
    if original_data.startswith(b'<?xml'):
        decl_end = original_data.index(b'?>')
        decl = original_data[: decl_end + 2].decode('ascii', errors='replace')
        if 'encoding=' in decl:
            enc_start = decl.index('encoding=') + len('encoding=') + 1
            enc_end   = decl.index(decl[enc_start - 1], enc_start)
            encoding  = decl[enc_start:enc_end]

    return etree.tostring(
        root,
        xml_declaration=True,
        encoding=encoding,
        standalone=True,
    )


def _get_xml_part_names(zip_path: str) -> list[str]:
    # Return all XML part names inside the docx zip
    with zipfile.ZipFile(zip_path, 'r') as zf:
        return [n for n in zf.namelist() if n.endswith('.xml')]


def _patch_docx_in_place(
    docx_path: str,
    patches: dict[str, bytes],
) -> None:
    # Rewrite *patches* (part_name → new_bytes) inside the docx zip in-place.
    #  All other entries in the zip are left completely untouched - this is what preserves images, styles, themes, and relationships.

    # Read the whole zip into memory first (needed because ZipFile can't overwrite individual entries without rewriting the whole archive).
    with zipfile.ZipFile(docx_path, 'r') as zf:
        names    = zf.namelist()
        originals = {n: zf.read(n) for n in names}
        infos    = {info.filename: info for info in zf.infolist()}

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', compression=zipfile.ZIP_DEFLATED) as out:
        for name in names:
            data = patches.get(name, originals[name])
            out.writestr(infos[name], data)

    with open(docx_path, 'wb') as f:
        f.write(buf.getvalue())



# Main entry point

def translate_docx(
    input_path: str,
    output_path: str,
    translator: Any,
    target_lang: str,
    *,
    cancel_event: threading.Event | None = None,
    source_lang: str = "auto",
) -> None:
    auto_detect = source_lang == "auto"
    errors = 0

    if cancel_event is not None and cancel_event.is_set():
        raise CancelledError("Translation cancelled before starting")

    # ── Step 1: copy original to output ──────────────────────────────────────
    shutil.copy2(input_path, output_path)
    logger.debug("Copied %s → %s", input_path, output_path)

    # ── Step 2: discover all XML parts to patch ───────────────────────────────
    all_xml_parts = _get_xml_part_names(output_path)

    # Collect parts: body + headers + footers + footnotes + endnotes
    text_parts = []
    for name in all_xml_parts:
        lower = name.lower()
        if (
            lower == "word/document.xml"
            or lower.startswith("word/header")
            or lower.startswith("word/footer")
            or lower == "word/footnotes.xml"
            or lower == "word/endnotes.xml"
        ):
            text_parts.append(name)

    # SmartArt diagram data parts
    smartart_parts = []
    for name in all_xml_parts:
        # content type sniffing via [Content_Types].xml is more correct, but
        # the path convention is reliable enough for our purposes
        if 'diagrams/data' in name.lower() or 'diagramdata' in name.lower():
            smartart_parts.append(name)

    # word/settings.xml — patched to suppress paragraph spacing at page tops
    settings_part = next(
        (n for n in all_xml_parts if n.lower() == "word/settings.xml"), None
    )

    logger.debug("Text parts: %s", text_parts)
    logger.debug("SmartArt parts: %s", smartart_parts)

    # ── Step 3: parse, collect, translate, serialise ─────────────────────────
    patches: dict[str, bytes] = {}

    with zipfile.ZipFile(output_path, 'r') as zf:
        _extra = [settings_part] if settings_part else []
        raw_parts = {name: zf.read(name) for name in text_parts + smartart_parts + _extra}

    # --- text parts (w:t, a:t, VML) ---
    for part_name in text_parts:
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError("Translation cancelled during XML collection")

        raw = raw_parts[part_name]
        try:
            root = _parse_xml_part(raw)
        except etree.XMLSyntaxError:
            logger.warning("Skipping malformed XML part: %s", part_name)
            continue

        pool: list[_TranslationUnit] = []
        pool.extend(_collect_wt_units(root))
        pool.extend(_collect_drawingml_units(root))
        pool.extend(_collect_vml_units(root))

        if not pool:
            continue

        logger.debug("%s: %d units collected", part_name, len(pool))
        groups = _group_units(pool, auto_detect=auto_detect)

        # Snapshot source lengths before translation so we can compute the
        # expansion ratio used by the layout-fix passes below.
        source_chars = sum(len(u.source_text) for u in pool)

        part_errors = _translate_and_writeback(
            pool, groups, translator, target_lang, cancel_event,
        )
        errors += part_errors

        # Compute how much longer the translated text is compared to the source.
        # We measure the live elem.text values via write_back, but the simplest
        # proxy is to compare the unit texts before and after translation — the
        # write_back closures already updated the XML, so we re-read the <w:t>
        # nodes.  Instead, we use the translated strings that ended up in the
        # units.  Since write_back mutates the XML in-place we can't read them
        # back easily, so we estimate: collect current <w:t> text sum.
        translated_chars = sum(
            len(wt.text) for wt in root.iter(_WT_TAG)
            if wt.text and wt.text.strip()
        )
        expansion_ratio = (
            translated_chars / source_chars if source_chars > 0 else 1.0
        )
        logger.debug(
            "%s: expansion_ratio=%.2f (%d src → %d tgt chars)",
            part_name, expansion_ratio, source_chars, translated_chars,
        )

        # Relax fixed-size containers and compress spacing so that translated
        # (often longer) text doesn't overflow table rows, text boxes, frames,
        # or push spacer-separated sections out of place.
        _apply_layout_fixes(root, expansion_ratio)

        patches[part_name] = _serialise_xml_part(root, raw)

    # --- SmartArt parts ---
    for part_name in smartart_parts:
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError("Translation cancelled during SmartArt collection")

        raw = raw_parts[part_name]
        try:
            pool, root = _collect_smartart_units_from_blob(raw)
        except Exception:
            logger.warning("Failed to parse SmartArt part %s", part_name, exc_info=True)
            continue

        if not pool:
            continue

        logger.debug("%s: %d SmartArt units", part_name, len(pool))
        groups = _group_units(pool, auto_detect=auto_detect)
        errors += _translate_and_writeback(
            pool, groups, translator, target_lang, cancel_event,
        )
        patches[part_name] = _serialise_xml_part(root, raw)

    # --- settings.xml: suppress first-paragraph spacing at page tops ---
    if settings_part:
        raw_settings = raw_parts.get(settings_part)
        if raw_settings:
            try:
                settings_root = _parse_xml_part(raw_settings)
                SUPPRESS_TAG = qn('w:suppressFirstParagraphSpacing')
                if settings_root.find(SUPPRESS_TAG) is None:
                    settings_root.append(etree.Element(SUPPRESS_TAG))
                    patches[settings_part] = _serialise_xml_part(
                        settings_root, raw_settings
                    )
                    logger.debug(
                        "Injected w:suppressFirstParagraphSpacing into %s",
                        settings_part,
                    )
            except Exception:
                logger.warning(
                    "Failed to patch %s", settings_part, exc_info=True
                )

    # ── Step 4: write only modified parts back into the zip ──────────────────
    if patches:
        logger.debug("Patching %d XML parts in %s", len(patches), output_path)
        _patch_docx_in_place(output_path, patches)

    if errors:
        raise RuntimeError(f"Translation completed with {errors} failed units")
