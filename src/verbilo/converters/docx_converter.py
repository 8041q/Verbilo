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
_NS_WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
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
    # Return True if *elem* is nested inside a ``<w:del>`` tracked-change block.
    node = elem.getparent()
    while node is not None:
        if node.tag == _WDEL_TAG:
            return True
        node = node.getparent()
    return False


def _is_inside_drawing_run(elem: etree._Element) -> bool:
    # Return True if *elem* is a ``<w:t>`` inside a ``<w:r>`` that also contains a ``<w:drawing>``.
    _W_R       = qn('w:r')
    _W_DRAWING = qn('w:drawing')
    node = elem.getparent()
    while node is not None:
        if node.tag == _W_R:
            return node.find(_W_DRAWING) is not None
        node = node.getparent()
    return False


def _is_inside_wgp(elem: etree._Element) -> bool:
    # Return True if *elem* is nested inside a ``<wpg:wgp>`` Word Processing Group.
    node = elem.getparent()
    while node is not None:
        if node.tag == _WPG_WGP:
            return True
        node = node.getparent()
    return False


import re as _re

_RE_PURE_NUMBER = _re.compile(
    r"""
    ^[\s\u00a0]*           # optional leading whitespace / NBSP
    [+-﹢﹣]?              # optional sign (incl. fullwidth small-form ﹢﹣)
    (?:
        \d{1,3}(?:[.,\s]\d{3})*  # thousands-grouped integer e.g. 1,234
        |\d+                      # plain integer
    )
    (?:[.,]\d+)?           # optional decimal part
    [\s\u00a0]*            # optional trailing whitespace
    (?:%|°|℃|℉|㎡|㎞|km|cm|mm|m²|m³|€|\$|£|¥|₹|元|円|₩|V|A|W|Hz|kV|mA|kW|Ω|dB|MPa|kPa)?  # optional unit
    [\s\u00a0]*$
    """,
    _re.VERBOSE,
)

# Matches compound numeric/technical expressions that are language-independent
# and should be passed through untranslated.  Examples:
#   ﹢5℃～﹢40℃   0~75°   Φ20mm   DC24V   ≥75%RH   -20℃~+60℃
_RE_TECHNICAL_EXPR = _re.compile(
    r'^[\s\u00a0]*'
    r'[+-﹢﹣]?'
    r'[\d.,]+'
    r'[\s\u00a0~～\-–—/×x*Φφ°℃℉%㎡㎞VAwWΩ㏀㏁HzdBMPakK㎝㎜㎝㎞℃℉]*'
    r'(?:[+-﹢﹣]?[\d.,]+[\s\u00a0~～\-–—/×x*Φφ°℃℉%㎡㎞VAwWΩ㏀㏁HzdBMPakK㎝㎜]*)*'
    r'(?:mm|cm|km|kV|mA|kW|Hz|dB|MPa|kPa|RH|VA|pF|nF|μF|mH|pH|Nm)?'
    r'[\s\u00a0]*$'
)

# Matches standalone technical symbols with a number: Φ20, ≥75, ≤100
_RE_SYMBOL_NUMBER = _re.compile(
    r'^[\s\u00a0]*[Φφ≥≤><≧≦±∅][\s\u00a0]*\d[\d.,]*'
    r'(?:mm|cm|km|kV|mA|kW|Hz|dB|MPa|kPa|RH|VA|pF|nF|μF|%|°|℃|℉)?'
    r'[\s\u00a0]*$'
)

# Matches strings that are MOSTLY digits — the text contains digits but also
# CJK/Latin characters that give the translator context to spell out the number.
# We extract and protect the numeric portion instead of skipping the whole unit.
_RE_HAS_DIGIT = _re.compile(r'\d')


def _is_numeric_only(text: str) -> bool:
    # Return True if *text* is a standalone number (possibly with units)
    # or a compound technical expression (e.g. temperature ranges, dimensions).
    return bool(
        _RE_PURE_NUMBER.match(text)
        or _RE_TECHNICAL_EXPR.match(text)
        or _RE_SYMBOL_NUMBER.match(text)
    )


# Matches strings that are entirely punctuation/symbols with no translatable
# lexical content.  Sending these to OPUS-MT produces hallucinations because
# the model has nothing real to work with.
_RE_NO_TRANS = _re.compile(
    r'^[\s'
    r'\u0021-\u002F\u003A-\u0040\u005B-\u0060\u007B-\u007E'  # ASCII punct
    r'\u00A0-\u00BF'          # Latin-1 supplement (°, ±, ×, ÷ …)
    r'\u2000-\u206F'          # General punctuation (—, …, •, etc.)
    r'\u2190-\u21FF'          # Arrows
    r'\u2300-\u23FF'          # Misc technical
    r'\u2460-\u24FF'          # Enclosed alphanumerics (①②③)
    r'\u2500-\u257F'          # Box drawing
    r'\u2580-\u259F'          # Block elements
    r'\u25A0-\u25FF'          # Geometric shapes (●○■□)
    r'\u2600-\u26FF'          # Misc symbols (★☆)
    r'\u2700-\u27BF'          # Dingbats
    r'\u3000-\u303F'          # CJK punctuation (。，、；：？！…—～·【】《》)
    r'\uFE50-\uFE6F'          # Small Form Variants (\uff62﹢﹣﹤﹥ etc.)
    r'\uFF00-\uFF0F\uFF1A-\uFF20\uFF3B-\uFF40\uFF5B-\uFF65'  # Fullwidth punct
    r'\uFF66-\uFFEF'          # Halfwidth/fullwidth
    r']+$'
)


def _is_symbol_only(text: str) -> bool:
    # Return True if *text* contains no translatable lexical content —
    # only punctuation, symbols, whitespace, or enclosed numbers (①②③).
    # Such segments cause hallucinations in OPUS-MT and should be skipped.
    return bool(_RE_NO_TRANS.match(text))


def _is_symbol_char(c: str) -> bool:
    # Return True if the single character *c* is a non-translatable symbol.
    # Used by strip_symbol_frame to identify leading/trailing decoration chars.
    cp = ord(c)
    return (
        0x0021 <= cp <= 0x002F or 0x003A <= cp <= 0x0040 or
        0x005B <= cp <= 0x0060 or 0x007B <= cp <= 0x007E or
        0x00A0 <= cp <= 0x00BF or
        0x2000 <= cp <= 0x206F or 0x2190 <= cp <= 0x21FF or
        0x2300 <= cp <= 0x23FF or 0x2460 <= cp <= 0x24FF or
        0x2500 <= cp <= 0x257F or 0x2580 <= cp <= 0x259F or
        0x25A0 <= cp <= 0x25FF or 0x2600 <= cp <= 0x26FF or
        0x2700 <= cp <= 0x27BF or
        # CJK Symbols and Punctuation — exclude sentence-level punctuation
        # 、(U+3001) and 。(U+3002) so the zh→en model handles them and they
        # are not blindly reattached after translation, bypassing protection.
        (0x3000 <= cp <= 0x303F and cp not in (0x3001, 0x3002)) or
        0xFE50 <= cp <= 0xFE6F or
        # Fullwidth forms — exclude ％ (U+FF05) and ‰-adjacent chars that are
        # covered by _PROTECTED_UNITS; stripping them as suffix would bypass
        # placeholder protection and leave unrestored [[N…]] tokens in output.
        (0xFF00 <= cp <= 0xFF0F and cp != 0xFF05) or
        0xFF1A <= cp <= 0xFF20 or
        0xFF3B <= cp <= 0xFF40 or 0xFF5B <= cp <= 0xFF65 or
        0xFF66 <= cp <= 0xFFEF or
        c in ' \t\n\r'
    )


# Matches an enumeration index at the current position, e.g. "1)", "12.", "a)"
_RE_ENUM_INDEX = _re.compile(r'\d+[)\]>.] *|[a-zA-Z][)\]>.] *')


def _strip_symbol_frame(text: str) -> tuple[str, str, str]:
    # Split *text* into (prefix, core, suffix)
    i = 0
    while i < len(text) and _is_symbol_char(text[i]):
        i += 1
    # Include an enumeration index that may follow the symbol prefix,
    # e.g. "(1)" → prefix="(1)", not prefix="(" with core="1)…"
    m = _RE_ENUM_INDEX.match(text, i)
    if m:
        i = m.end()
        # After the enum index there may be more symbols (e.g. a space)
        while i < len(text) and _is_symbol_char(text[i]):
            i += 1
    j = len(text)
    while j > i and _is_symbol_char(text[j - 1]):
        j -= 1
    return text[:i], text[i:j], text[j:]


# Special symbols / units that OPUS-MT's SentencePiece vocabulary cannot
# represent.  These are protected as [[U0]], [[U1]], … placeholders before
# translation and restored verbatim afterwards.
_PROTECTED_UNITS = _re.compile(
    r'℃|℉'                            # temperature symbols
    r'|[﹢﹣]'                          # small-form plus/minus
    r'|[～]'                            # fullwidth tilde
    r'|[Φφ∅]'                          # diameter symbols
    r'|[≥≤≧≦]'                         # comparison operators
    r'|[±]'                            # plus-minus
    r'|[×÷]'                           # multiplication/division
    r'|[°](?!C|F)'                     # degree (not followed by C/F — those are ℃/℉)
    r'|㎡|㎞|㎝|㎜'                     # squared/cubed CJK units
    r'|㏀|㏁'                           # kilo-ohm / mega-ohm
    r'|[Ωμ]'                           # ohm, micro
    r'|％'                              # fullwidth percent
    r'|‰'                              # per mille
)


# Fullwidth structural punctuation mapped to ASCII equivalents.
# These tokenize poorly in SentencePiece; replacing them with ASCII gives the
# model cleaner context.  Sentence-level CJK punctuation (，。) is intentionally
# excluded — the zh-en model expects them.
_FULLWIDTH_STRUCT: dict[str, str] = {
    '\uff1a': ':',   # ：
    '\uff1b': ';',   # ；
    '\uff08': '(',   # （
    '\uff09': ')',   # ）
    '\uff01': '!',   # ！
    '\uff1f': '?',   # ？
}
_FULLWIDTH_STRUCT_RE = _re.compile(
    '[' + ''.join(_FULLWIDTH_STRUCT.keys()) + ']'
)


def _strip_protected_tokens(text: str) -> tuple[str, dict[str, str]]:
    # Replace numeric tokens and special symbols in *text* with placeholders
    # before translation so they survive the round-trip through OPUS-MT unchanged.
    
    placeholders: dict[str, str] = {}
    counter = [0]

    def _replace_num(m: _re.Match) -> str:
        key = f'[[N{counter[0]}]]'
        placeholders[key] = m.group(0)
        counter[0] += 1
        return key

    # Pass 0: normalise fullwidth structural punctuation to ASCII (one-way).
    # We do NOT restore these in _restore_protected_tokens — the translated
    # English output should use ASCII punctuation.
    normalized = _FULLWIDTH_STRUCT_RE.sub(
        lambda m: _FULLWIDTH_STRUCT[m.group(0)], text
    )

    # Pass 1: protect special symbols / units
    ucounter = [0]

    def _replace_unit(m: _re.Match) -> str:
        key = f'[[U{ucounter[0]}]]'
        placeholders[key] = m.group(0)
        ucounter[0] += 1
        return key

    masked = _PROTECTED_UNITS.sub(_replace_unit, normalized)

    # Pass 2: protect numeric tokens: integers, decimals, version numbers, years
    masked = _re.sub(
        r'(?<!\w)(\d[\d,.\s]*\d|\d)(?!\w)',
        _replace_num,
        masked,
    )

    # Guard: if masking left fewer than 2 CJK characters of real content,
    # the translator has too little context — undo ALL masking and return
    # the normalised original so the model sees actual text.
    stripped_for_check = _re.sub(r'\[\[[NU]\d+\]\]', '', masked)
    cjk_left = len(_RE_CJK.findall(stripped_for_check))
    if cjk_left < 2 and cjk_left > 0 and placeholders:
        return normalized, {}

    return masked, placeholders


def _is_placeholder_only(text: str) -> bool:
    # Return True if *text* is nothing but placeholder tokens and whitespace
    stripped = _re.sub(r'\[\[[NU]\d+\]\]', '', text).strip()
    return stripped == ''


def _restore_protected_tokens(text: str, placeholders: dict[str, str]) -> str:
    # Substitute placeholders back into *text* after translation
    # Pass 1: exact
    for key, original in placeholders.items():
        text = text.replace(key, original)

    # Pass 2: fuzzy — only needed when exact pass left tokens behind
    for key, original in placeholders.items():
        m = _re.match(r'\[\[([NU])(\d+)\]\]', key)
        if m is None:
            continue
        tag = m.group(1)  # 'N' or 'U'
        idx = _re.escape(m.group(2))
        # Pattern set 1: various bracket corruptions
        pattern = _re.compile(
            r'[\[\(]\s*[\[\(]?\s*[' + tag + tag.lower() + r']\s*' + idx + r'\s*[\]\)][\]\)]?'
        )
        text = pattern.sub(original, text)
        # Pattern set 2: single-bracket form like [N0] or (U1)
        pattern2 = _re.compile(
            r'\[\s*[' + tag + tag.lower() + r']\s*' + idx + r'\s*\]'
        )
        text = pattern2.sub(original, text)
        # Pattern set 3: completely bracketless form like N0 or U1
        # Only match when surrounded by non-alphanumeric chars
        pattern3 = _re.compile(
            r'(?<![a-zA-Z0-9])' + '[' + tag + tag.lower() + ']' + idx + r'(?![a-zA-Z0-9])'
        )
        text = pattern3.sub(original, text)

    return text


# Compiled patterns used by the hallucination guard
_RE_REPETITION  = _re.compile(r'(.{4,}?)\1{3,}')         # same chunk repeated 3+ times (adjacent)
_RE_SENT_REPEAT = _re.compile(r'(\b.{6,}?)(\s+\1){2,}') # phrase repeated 3+ times with whitespace gap
_RE_STUTTER     = _re.compile(r'(\b\S+\b)(\s+\1){4,}') # same word 5+ times in a row
_RE_GARBAGE     = _re.compile(                              # mangled placeholder debris
    r'U\s*[T]\s*\d'                                      # "UT95"-style artefacts
    r'|,\s*,\s*,\s*,',                                   # run of empty commas
)
# Matches CJK unified ideographs and common CJK punctuation/symbols
_RE_CJK = _re.compile(
    r'[\u2E80-\u2EFF\u2F00-\u2FDF\u3000-\u303F\u3040-\u309F\u30A0-\u30FF'
    r'\u3100-\u312F\u3200-\u32FF\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF'
    r'\uFE30-\uFE4F]'
)


def _max_translated_len(source: str) -> int:
    # Return the maximum plausible byte-length for a translation of *source*.
    # CJK characters expand 8-15x into Latin script
    cjk_chars = len(_RE_CJK.findall(source))
    other_chars = len(source) - cjk_chars
    return cjk_chars * 15 + other_chars * 5


def _is_hallucinated(translated: str, source: str) -> bool:
    """Return True if *translated* looks like an NMT hallucination.

    Heuristics (conservative — we would rather pass bad output than discard good):
    - Output exceeds the CJK-aware length budget (repetition loop / verbosity).
    - Output contains a chunk repeated 3+ times in a row.
    - Output contains a phrase repeated 3+ times with whitespace gaps.
    - Output contains 5+ consecutive repetitions of the same word.
    - Output contains characteristic garbage patterns (e.g. "UT95", comma runs).
    - Source has word characters but output has none (e.g. "* * * * * *").
    """
    if not translated.strip():
        return False  # empty is handled elsewhere

    if len(translated) > _max_translated_len(source):
        logger.debug(
            "Hallucination guard: output too long (%d vs budget %d, src=%r)",
            len(translated), _max_translated_len(source), source[:30],
        )
        return True

    if _RE_REPETITION.search(translated):
        logger.debug("Hallucination guard: chunk repetition detected")
        return True

    if _RE_SENT_REPEAT.search(translated):
        logger.debug("Hallucination guard: sentence repetition detected")
        return True

    if _RE_STUTTER.search(translated):
        logger.debug("Hallucination guard: word stutter detected")
        return True

    if _RE_GARBAGE.search(translated):
        logger.debug("Hallucination guard: garbage pattern detected in: %r", translated)
        return True

    # Flag output with no word characters only when the SOURCE had word characters.
    # Pure-punctuation sources (●, /, °, ：) legitimately pass through as punctuation.
    if _re.search(r'\w', source) and not _re.search(r'\w', translated):
        logger.debug("Hallucination guard: no word characters in output")
        return True

    return False


# ── Layout-protection passes ─────────────────────────────────────────────────
# Run these on each XML root AFTER translation so that text-expansion caused
# by longer translated strings doesn't break the document layout.

_WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

# wpg:wgp — Word Processing Group container (one anchor, many member shapes)
_WPG_WGP     = f'{{{_NS_WPG}}}wgp'

# wp: positional tags used for implicit-group anchor normalization
_WP_ANCHOR_TAG  = f'{{{_WP_NS}}}anchor'
_WP_POS_H_TAG   = f'{{{_WP_NS}}}positionH'
_WP_POS_V_TAG2  = f'{{{_WP_NS}}}positionV'
_WP_POS_OFFSET  = f'{{{_WP_NS}}}posOffset'
_WP_EXTENT_TAG  = f'{{{_WP_NS}}}extent'

# Bounding-box proximity threshold for implicit diagram group detection (EMU)
# 914 400 EMU = 1 inch.  Shapes further apart than this are NOT the same group.
_MAX_GROUP_GAP_EMU = 914_400


def _fix_table_row_heights(root: etree._Element) -> None:
    # Convert ``exact`` row-height constraints to ``atLeast``.
    TR_HEIGHT = qn('w:trHeight')
    H_RULE    = qn('w:hRule')
    for trh in root.iter(TR_HEIGHT):
        if trh.get(H_RULE, '').lower() == 'exact':
            trh.set(H_RULE, 'atLeast')
            logger.debug("Relaxed exact row height → atLeast")


def _fix_textbox_autofit(root: etree._Element) -> None:
    # Enable shape auto-fit (``spAutoFit``) on text-box body properties where the translated text might overflow the original fixed size.
    _NS_WPS      = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    A_BODY_PR    = f'{{{_NS_A}}}bodyPr'
    WPS_BODY_PR  = f'{{{_NS_WPS}}}bodyPr'
    A_NO_AUTOFIT = f'{{{_NS_A}}}noAutofit'
    A_NORM_AUTO  = f'{{{_NS_A}}}normAutofit'
    A_SP_AUTO    = f'{{{_NS_A}}}spAutoFit'
    _AUTOFIT_TAGS = {A_NO_AUTOFIT, A_NORM_AUTO, A_SP_AUTO}

    for body_pr in [*root.iter(A_BODY_PR), *root.iter(WPS_BODY_PR)]:
        # Skip shapes that are members of a wpg:wgp group — resizing them
        # independently would break the group's internal layout.
        if _is_inside_wgp(body_pr):
            continue
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
    # Ensure legacy word-processing frames (``<w:framePr>``) can grow.
    FRAME_PR = qn('w:framePr')
    H_RULE   = qn('w:hRule')
    for fp in root.iter(FRAME_PR):
        if fp.get(H_RULE, '').lower() == 'exact':
            fp.set(H_RULE, 'atLeast')
            logger.debug("Relaxed framePr exact height → atLeast")


def _fix_vml_textbox_autosize(root: etree._Element) -> None:
    # Allow legacy VML text boxes to grow to fit translated (longer) text.
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
    # Widen label-sized DrawingML/WPS text boxes to reduce unwanted line-wrapping.
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
        # Walk up to the enclosing wp:anchor / wp:inline and return its extent
        node = shape_elem.getparent()
        while node is not None:
            if node.tag in (WP_ANCHOR, WP_INLINE):
                return node.find(WP_EXT)
            node = node.getparent()
        return None

    # WPS text boxes: <wps:wsp> with <wps:txbx>
    for wsp in root.iter(WPS_WSP):
        # Skip group members — their extents are in group-local coordinates.
        if _is_inside_wgp(wsp):
            continue
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
        # Skip group members — their extents are in group-local coordinates.
        if _is_inside_wgp(sp):
            continue
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
    # Return True if *para* is a layout-spacer paragraph with no visible text.
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
    # Return True if *para* starts a new page via pageBreakBefore or a w:br page break.
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
    # Split all body paragraphs into page sections divided by hard page breaks.
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
    # Compress spacing within one page section proportionally to expansion_ratio.
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
    # Change ``w:lineRule="exact"`` to ``atLeast`` on a single paragraph.
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
    # Compress spacing section-by-section, respecting hard page boundaries.
    if expansion_ratio <= 1.0:
        return
    for section in _collect_page_sections(root):
        _compress_section_spacing(section, expansion_ratio)


def _compress_paragraph_spacing(root: etree._Element, expansion_ratio: float) -> None:
    # Relax exact line spacing on all paragraphs (handled inside section pass).
    # This is now a no-op — the section-aware pass in _compress_spacer_spacing
    # handles both spacer and content paragraphs together so we don't double-apply.
    pass


def _ensure_list_indentation(root: etree._Element) -> None:
    # Add standard Western indentation to list paragraphs that have none.
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


def _normalize_anchor_reference_frames(root: etree._Element) -> None:
    # Unify ``positionV`` reference frames for implicitly grouped floating anchors.
    W_P         = qn('w:p')
    _PARA_REL   = {'paragraph', 'line', 'character'}

    def _read_anchor_bbox(
        anchor: etree._Element,
    ) -> tuple[int, int, int, int] | None:
        # Return (x0, y0, x1, y1) bounding box in EMU, or None if unreadable
        try:
            pos_h = anchor.find(_WP_POS_H_TAG)
            pos_v = anchor.find(_WP_POS_V_TAG2)
            ext   = anchor.find(_WP_EXTENT_TAG)
            if pos_h is None or pos_v is None or ext is None:
                return None
            off_h_elem = pos_h.find(_WP_POS_OFFSET)
            off_v_elem = pos_v.find(_WP_POS_OFFSET)
            if off_h_elem is None or off_v_elem is None:
                return None
            x0 = int(off_h_elem.text or '0')
            y0 = int(off_v_elem.text or '0')
            cx = int(ext.get('cx', '0'))
            cy = int(ext.get('cy', '0'))
            return x0, y0, x0 + cx, y0 + cy
        except (ValueError, TypeError):
            return None

    def _bbox_gap(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> int:
        # Chebyshev gap between two bounding boxes (0 if they overlap)
        x_gap = max(0, max(a[0], b[0]) - min(a[2], b[2]))
        y_gap = max(0, max(a[1], b[1]) - min(a[3], b[3]))
        return max(x_gap, y_gap)

    for para in root.iter(W_P):
        # Collect all wp:anchor elements nested anywhere inside this paragraph.
        anchors = list(para.iter(_WP_ANCHOR_TAG))
        if len(anchors) < 2:
            continue

        # Read bounding boxes and relativeFrom values.
        bboxes: list[tuple[int, int, int, int] | None] = []
        rel_v:  list[str] = []
        for anchor in anchors:
            bboxes.append(_read_anchor_bbox(anchor))
            pv = anchor.find(_WP_POS_V_TAG2)
            rel_v.append(pv.get('relativeFrom', '') if pv is not None else '')

        # Only act if the paragraph has mixed relativeFrom values.
        has_page = any(r == 'page' for r in rel_v)
        has_para = any(r in _PARA_REL for r in rel_v)
        if not (has_page and has_para):
            continue

        # Build proximity clusters (single-linkage connected-components).
        n = len(anchors)
        cluster_id = list(range(n))          # each anchor starts in its own cluster

        def _find(i: int) -> int:
            while cluster_id[i] != i:
                cluster_id[i] = cluster_id[cluster_id[i]]
                i = cluster_id[i]
            return i

        def _union(i: int, j: int) -> None:
            ri, rj = _find(i), _find(j)
            if ri != rj:
                cluster_id[ri] = rj

        for i in range(n):
            if bboxes[i] is None:
                continue
            for j in range(i + 1, n):
                if bboxes[j] is None:
                    continue
                if _bbox_gap(bboxes[i], bboxes[j]) <= _MAX_GROUP_GAP_EMU:
                    _union(i, j)

        # Group indices by cluster root.
        from collections import defaultdict
        clusters: dict[int, list[int]] = defaultdict(list)
        for i in range(n):
            clusters[_find(i)].append(i)

        # For each cluster: if it mixes page-absolute + paragraph-relative
        # anchors, convert the page-absolute ones.
        for members in clusters.values():
            page_idxs = [i for i in members if rel_v[i] == 'page']
            para_idxs = [i for i in members if rel_v[i] in _PARA_REL]
            if not page_idxs or not para_idxs:
                continue  # Homogeneous cluster — nothing to do

            # Estimate the paragraph's page-Y by pairing the para-relative
            # anchor with the smallest absolute posV with the page-absolute
            # anchor that is closest to it vertically.
            # Both anchors are assumed to be at their correct visual positions
            # right now (i.e. this is the SOURCE document's state).
            # para_page_Y ≈ page_posV − para_posV
            best_para_y0: int | None = None
            best_page_y0: int | None = None

            for i in para_idxs:
                if bboxes[i] is not None:
                    y0 = bboxes[i][1]
                    if best_para_y0 is None or y0 < best_para_y0:
                        best_para_y0 = y0

            for i in page_idxs:
                if bboxes[i] is not None:
                    y0 = bboxes[i][1]
                    if best_page_y0 is None or y0 < best_page_y0:
                        best_page_y0 = y0

            if best_para_y0 is None or best_page_y0 is None:
                continue  # Can't estimate — skip cluster

            para_page_Y = best_page_y0 - best_para_y0

            # Convert each page-absolute anchor.
            for i in page_idxs:
                anchor = anchors[i]
                pv = anchor.find(_WP_POS_V_TAG2)
                if pv is None:
                    continue
                off_elem = pv.find(_WP_POS_OFFSET)
                if off_elem is None:
                    continue
                try:
                    current_y = int(off_elem.text or '0')
                except ValueError:
                    continue

                new_y = current_y - para_page_Y
                pv.set('relativeFrom', 'paragraph')
                off_elem.text = str(new_y)
                logger.debug(
                    "Converted anchor positionV: page@%d → paragraph@%d "
                    "(para_page_Y=%d)",
                    current_y, new_y, para_page_Y,
                )


def _apply_layout_fixes(root: etree._Element, expansion_ratio: float = 1.0) -> None:
    # Run all layout-protection passes on *root* in one call.
    _normalize_anchor_reference_frames(root)
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

        # Skip symbol/punctuation-only segments — no lexical content to translate,
        # and sending them to OPUS-MT causes hallucinations.
        if _is_symbol_only(text):
            continue

        # Strip leading/trailing symbol chars (e.g. '●' in '●拔出电源描头') so
        # the model only sees the translatable core.  The symbols are reattached
        # verbatim in write_back.  If nothing translatable remains, skip.
        sym_prefix, core_text, sym_suffix = _strip_symbol_frame(text.strip())
        if not core_text:
            continue

        # For mixed text (e.g. "2023年"), mask numeric tokens so they survive
        # translation unchanged, then restore them in write_back.
        masked_text, placeholders = _strip_protected_tokens(core_text)
        is_heading = _is_in_heading(wt)

        # If masking consumed everything, sending placeholder-only text to
        # OPUS-MT causes hallucinations — skip the segment entirely.
        if _is_placeholder_only(masked_text):
            continue

        def _make_wb(elem: etree._Element, orig_text: str, ph: dict[str, str],
                     src: str, s_pre: str, s_suf: str):
            def wb(translated: str) -> None:
                # Restore any numeric placeholders the translator may have mangled
                result = _restore_protected_tokens(translated, ph)
                # Reject hallucinated output — fall back to the original core
                if _is_hallucinated(result, src):
                    logger.warning(
                        "Hallucination detected; keeping original: %r -> %r",
                        orig_text, result,
                    )
                    result = _restore_protected_tokens(src, ph)
                # Reattach symbol frame and preserve original whitespace
                leading  = orig_text[: len(orig_text) - len(orig_text.lstrip())]
                trailing = orig_text[len(orig_text.rstrip()):]
                elem.text = leading + s_pre + result.strip() + s_suf + trailing
                if leading or trailing or s_pre or s_suf:
                    elem.set(
                        '{http://www.w3.org/XML/1998/namespace}space', 'preserve'
                    )
            return wb

        units.append(_TranslationUnit(
            source_text=masked_text,
            is_heading=is_heading,
            write_back=_make_wb(wt, text, placeholders, masked_text,
                                sym_prefix, sym_suffix),
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

        if _is_symbol_only(text):
            continue

        sym_prefix, core_text, sym_suffix = _strip_symbol_frame(text.strip())
        if not core_text:
            continue

        masked_text, placeholders = _strip_protected_tokens(core_text)

        # If masking consumed everything, skip to avoid OPUS-MT hallucinations.
        if _is_placeholder_only(masked_text):
            continue

        def _make_wb(e: etree._Element, orig_text: str, ph: dict[str, str],
                     src: str, s_pre: str, s_suf: str):
            def wb(translated: str) -> None:
                result = _restore_protected_tokens(translated, ph)
                if _is_hallucinated(result, src):
                    logger.warning(
                        "Hallucination detected; keeping original: %r -> %r",
                        orig_text, result,
                    )
                    result = _restore_protected_tokens(src, ph)
                leading  = orig_text[: len(orig_text) - len(orig_text.lstrip())]
                trailing = orig_text[len(orig_text.rstrip()):]
                e.text = leading + s_pre + result.strip() + s_suf + trailing
                if leading or trailing or s_pre or s_suf:
                    e.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            return wb

        units.append(_TranslationUnit(
            source_text=masked_text,
            write_back=_make_wb(at_elem, text, placeholders, masked_text,
                                sym_prefix, sym_suffix),
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
    # Serialise *root* back to bytes, preserving the original XML declaration
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