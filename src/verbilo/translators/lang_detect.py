# Language detection backends for source-language filtering.
#
# Supported detectors:
#   "fasttext"  – fast, accurate, needs fasttext-langdetect
#   "lingua"    – designed for short text, needs lingua-language-detector
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
    # Normalise a language code to lowercase ISO 639-1 (2 letters)
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
    # Strip numbers, punctuation, URLs, emails — keep only 'letter' content
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
    # Detect language using fasttext-langdetect.  Returns (code, conf) or None
    try:
        from fast_langdetect import detect as ft_detect  # type: ignore
        result = ft_detect(text)
        code = _norm_code(result.get("lang", ""))
        conf = float(result.get("score", 0.0))
        if code:
            return code, conf
    except Exception:
        pass
    return None


def _detect_lingua(text: str) -> Optional[tuple[str, float]]:
    # Detect language using lingua-language-detector.  Returns (code, conf) or None
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


# ---------------------------------------------------------------------------
# Detector registry — maps name → callable
# ---------------------------------------------------------------------------

_DETECTORS: dict[str, callable] = {
    "fasttext": _detect_fasttext,
    "fastText": _detect_fasttext,
    "lingua": _detect_lingua,
}


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def detect_language(text: str, detector: str = "fasttext") -> tuple[str, float]:
    # Return (language_code, confidence) for *text*
    cleaned = _clean_for_detection(text)
    if len(cleaned) < _MIN_DETECT_CHARS:
        return "und", 0.0

    fn = _DETECTORS.get(detector.lower())
    if fn is None:
        logger.warning("Unknown detector '%s'", detector)
        return "und", 0.0

    result = fn(cleaned)
    if result:
        return result
    return "und", 0.0


def is_source_language(
    text: str,
    source_lang: str,
    detector: str = "fasttext",
    strict: bool = False,
) -> bool:
    # Return ``True`` if *text* appears to be written in *source_lang*.
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
