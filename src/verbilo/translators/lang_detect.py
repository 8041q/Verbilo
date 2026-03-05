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

# Confidence floor for fasttext results to be trusted on their own.
# (Lingua uses its own built-in minimum_relative_distance instead.)
_CONFIDENCE_THRESHOLD = 0.65

# Lingua: minimum gap between the top-two language probabilities.
# If the gap is smaller than this, detect_language_of() returns None
# (= ambiguous).  Keep moderate so short cell text is not over-filtered.
_LINGUA_MIN_RELATIVE_DISTANCE = 0.25

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
        results = ft_detect(text, model="auto", k=1)
        if not results:
            return None
        first = results[0]
        lang_val = first.get("lang", "")
        if not isinstance(lang_val, str):
            lang_val = ""
        code = _norm_code(lang_val)
        conf = float(first.get("score", 0.0))
        if code:
            return code, conf
    except Exception:
        pass
    return None


def _detect_lingua(text: str) -> Optional[tuple[str, float]]:
    # Detect language using lingua-language-detector.  Returns (code, conf) or None.
    # Uses Lingua's built-in minimum_relative_distance for confidence filtering:
    # detect_language_of() returns None when the gap between top-two candidates
    # is too small (ambiguous).  When it returns a language, we trust it (1.0).
    try:
        detector = _get_lingua_detector()
        language = detector.detect_language_of(text)
        if language is None:
            return None  # ambiguous — Lingua's own confidence filter rejected it
        code = _norm_code(language.iso_code_639_1.name.lower())
        if code:
            return code, 1.0
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
            .with_minimum_relative_distance(_LINGUA_MIN_RELATIVE_DISTANCE)
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


def is_source_language_batch(
    texts: list[str],
    source_lang: str,
    detector: str = "fasttext",
    strict: bool = False,
) -> list[bool]:
    # Batch version of is_source_language.
    # For Lingua, uses detect_languages_in_parallel_of() across all CPU cores
    # in a single call instead of looping per-text.
    # For fasttext, falls back to per-item calls (no native batch API).
    if source_lang == "auto":
        return [True] * len(texts)

    src = _norm_code(source_lang)

    if detector.lower() == "lingua":
        return _is_source_language_batch_lingua(texts, src, strict)

    return [is_source_language(t, source_lang, detector=detector, strict=strict) for t in texts]


def _is_source_language_batch_lingua(
    texts: list[str],
    src: str,
    strict: bool,
) -> list[bool]:
    # Uses Lingua's detect_languages_in_parallel_of() — classifies all texts
    # simultaneously across all available CPU cores.
    # None results mean ambiguous (filtered by with_minimum_relative_distance).
    cleaned = [_clean_for_detection(t) for t in texts]
    results: list[bool] = [not strict] * len(texts)  # default for short/undecidable

    detection_indices = [i for i, c in enumerate(cleaned) if len(c) >= _MIN_DETECT_CHARS]
    if not detection_indices:
        return results

    detection_texts = [cleaned[i] for i in detection_indices]
    try:
        detector_instance = _get_lingua_detector()
        # Returns list[Language | None]
        detected = detector_instance.detect_languages_in_parallel_of(detection_texts)
        for list_idx, orig_idx in enumerate(detection_indices):
            language = detected[list_idx]
            if language is None:
                results[orig_idx] = not strict
            else:
                code = _norm_code(language.iso_code_639_1.name.lower())
                results[orig_idx] = bool(code and code == src)
    except Exception:
        logger.warning("Lingua parallel detection failed; falling back to per-item", exc_info=True)
        for list_idx, orig_idx in enumerate(detection_indices):
            results[orig_idx] = is_source_language(
                texts[orig_idx], src, detector="lingua", strict=strict
            )

    return results
