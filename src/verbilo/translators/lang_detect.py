# Multi-engine language detection with voting / confidence scoring.
#
# Supported detectors:
#   "auto"      – majority-vote across all available engines
#   "fasttext"  – fast, accurate, needs fasttext-langdetect
#   "lingua"    – designed for short text, needs lingua-language-detector
#   "langdetect"– legacy, needs langdetect
#
# When source_lang == "auto" the detector is never called — every cell
# is sent to the translator unconditionally.

from __future__ import annotations

import logging
import re
import unicodedata
from typing import Optional

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Thresholds
# ---------------------------------------------------------------------------

# Text shorter than this (after stripping numbers/punctuation) is always
# considered translatable — no detector is reliable on very short strings.
_MIN_DETECT_CHARS = 10

# Confidence floor for a single-engine result to be trusted on its own.
_CONFIDENCE_THRESHOLD = 0.65

# ISO 639-1 language code aliases — some detectors return longer codes.
_LANG_ALIASES: dict[str, str] = {
    "zh-cn": "zh", "zh-tw": "zh", "zho": "zh",
    "por": "pt", "eng": "en", "spa": "es", "fra": "fr", "deu": "de",
    "ita": "it", "nld": "nl", "rus": "ru", "jpn": "ja", "kor": "ko",
    "ara": "ar", "hin": "hi", "tur": "tr", "pol": "pl", "ukr": "uk",
    "ron": "ro", "ces": "cs", "hun": "hu", "swe": "sv", "dan": "da",
    "fin": "fi", "nor": "no", "nob": "no", "nno": "no",
}


def _norm_code(code: str) -> str:
    """Normalise a language code to lowercase ISO 639-1 (2 letters)."""
    code = code.strip().lower()
    code = _LANG_ALIASES.get(code, code)
    # strip region suffix  e.g. "pt-br" -> "pt"
    if len(code) > 2 and "-" in code:
        code = code.split("-")[0]
    if len(code) > 2 and "_" in code:
        code = code.split("_")[0]
    return code


# ---------------------------------------------------------------------------
# Text pre-processing for detection
# ---------------------------------------------------------------------------

_NON_LETTER_RE = re.compile(r"[^a-zA-Z\u00C0-\u024F\u0400-\u04FF"
                             r"\u0600-\u06FF\u4e00-\u9fff\u3040-\u30FF"
                             r"\uAC00-\uD7AF]")


def _clean_for_detection(text: str) -> str:
    """Strip numbers, punctuation, URLs, emails — keep only 'letter' content."""
    text = unicodedata.normalize("NFC", text)
    # remove URLs
    text = re.sub(r"https?://\S+", " ", text)
    # remove emails
    text = re.sub(r"\S+@\S+", " ", text)
    # collapse non-letter chars to spaces
    text = _NON_LETTER_RE.sub(" ", text)
    return " ".join(text.split())


# ---------------------------------------------------------------------------
# Individual detector back-ends
# ---------------------------------------------------------------------------

def _detect_fasttext(text: str) -> Optional[tuple[str, float]]:
    """Detect language using fasttext-langdetect.  Returns (code, conf) or None."""
    try:
        from ftlangdetect import detect as ft_detect  # type: ignore
        result = ft_detect(text)
        code = _norm_code(result.get("lang", ""))
        conf = float(result.get("score", 0.0))
        if code:
            return code, conf
    except Exception:
        pass
    return None


def _detect_lingua(text: str) -> Optional[tuple[str, float]]:
    """Detect language using lingua-language-detector.  Returns (code, conf) or None."""
    try:
        from lingua import LanguageDetectorBuilder  # type: ignore
        # Build a lightweight detector — cached at module level after first call.
        detector = _get_lingua_detector()
        result = detector.compute_language_confidence_values(text)
        if result:
            top = result[0]
            code = _norm_code(top.language.iso_code_639_1.name.lower())
            conf = float(top.value)
            if code:
                return code, conf
    except Exception:
        pass
    return None


# Module-level cache for the lingua detector (expensive to build).
_lingua_detector_instance = None


def _get_lingua_detector():
    global _lingua_detector_instance
    if _lingua_detector_instance is None:
        from lingua import LanguageDetectorBuilder  # type: ignore
        _lingua_detector_instance = (
            LanguageDetectorBuilder
            .from_all_languages()
            .with_preloaded_language_models()
            .build()
        )
    return _lingua_detector_instance


def _detect_langdetect(text: str) -> Optional[tuple[str, float]]:
    """Detect language using langdetect.  Returns (code, conf) or None."""
    try:
        import langdetect  # type: ignore
        langdetect.DetectorFactory.seed = 0
        probs = langdetect.detect_langs(text)
        if probs:
            top = probs[0]
            code = _norm_code(str(top.lang))
            conf = float(top.prob)
            if code:
                return code, conf
    except Exception:
        pass
    return None


# ---------------------------------------------------------------------------
# Detector registry — maps name → callable
# ---------------------------------------------------------------------------

_DETECTORS: dict[str, callable] = {
    "fasttext": _detect_fasttext,
    "fastText": _detect_fasttext,
    "lingua": _detect_lingua,
    "langdetect": _detect_langdetect,
}

# Preferred order for "auto" mode — most accurate first.
_AUTO_ORDER = ["fasttext", "lingua", "langdetect"]


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def detect_language(text: str, detector: str = "auto") -> tuple[str, float]:
    """Return (language_code, confidence) for *text*.

    *detector* can be ``"auto"``, ``"fasttext"``, ``"lingua"`` or
    ``"langdetect"``.  ``"auto"`` queries all available engines and returns
    the majority vote (or the highest-confidence single result).

    Returns ``("und", 0.0)`` when detection fails entirely.
    """
    cleaned = _clean_for_detection(text)
    if len(cleaned) < _MIN_DETECT_CHARS:
        return "und", 0.0

    if detector != "auto":
        fn = _DETECTORS.get(detector.lower())
        if fn is None:
            logger.warning("Unknown detector '%s'; falling back to auto", detector)
        else:
            result = fn(cleaned)
            if result:
                return result
            return "und", 0.0

    # Auto mode: collect votes from every available engine.
    votes: list[tuple[str, float]] = []
    for name in _AUTO_ORDER:
        fn = _DETECTORS[name]
        result = fn(cleaned)
        if result is not None:
            votes.append(result)

    if not votes:
        return "und", 0.0

    if len(votes) == 1:
        return votes[0]

    # Tally: count how many engines agree on each code.
    tally: dict[str, tuple[int, float]] = {}
    for code, conf in votes:
        count, best_conf = tally.get(code, (0, 0.0))
        tally[code] = (count + 1, max(best_conf, conf))

    # Pick the code with the most votes; break ties by confidence.
    winner = max(tally.items(), key=lambda kv: (kv[1][0], kv[1][1]))
    return winner[0], winner[1][1]


def is_source_language(
    text: str,
    source_lang: str,
    detector: str = "auto",
    strict: bool = False,
) -> bool:
    """Return ``True`` if *text* appears to be written in *source_lang*.

    When *source_lang* is ``"auto"`` this always returns ``True`` (translate
    everything).  For very short text (< ``_MIN_DETECT_CHARS`` letters) the
    result depends on *strict*: in lenient mode (default) returns ``True`` to
    avoid missing cells; in strict mode returns ``False`` to avoid wrongly
    translating unknown segments inside mixed-language cells.

    *strict=True* should only be used when evaluating individual segments of
    a mixed-language cell, where the safe default is to preserve rather than
    translate.
    """
    if source_lang == "auto":
        return True

    src = _norm_code(source_lang)

    cleaned = _clean_for_detection(text)
    if len(cleaned) < _MIN_DETECT_CHARS:
        # Too short to detect reliably.
        # Lenient: translate (avoid missing cells).
        # Strict: preserve (avoid translating ambiguous segments).
        return not strict

    detected_code, confidence = detect_language(text, detector=detector)

    if detected_code == "und":
        # Detection failed entirely.
        # Lenient: translate to be safe.  Strict: preserve to be safe.
        return not strict

    match = detected_code == src

    # Require decent confidence to *reject* a match.
    # Lenient: low-confidence non-match → translate (avoid skipping cells).
    # Strict:  low-confidence non-match → preserve (avoid clobbering segments).
    if not match and confidence < _CONFIDENCE_THRESHOLD:
        return not strict

    return match
