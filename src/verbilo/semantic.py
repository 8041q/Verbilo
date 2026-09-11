from __future__ import annotations

import hashlib
import json
import logging
import re
import threading
import unicodedata
from collections import Counter
from dataclasses import dataclass, field, replace
from typing import Any, Callable, Iterable, Literal, Mapping, Sequence

from .utils import CancelledError
from .translation_memory import TranslationMemory, TranslationMemoryEntry

logger = logging.getLogger(__name__)

TranslationMode = Literal["faithful", "natural", "concise"]
TranslationRole = Literal[
    "body",
    "heading",
    "table-cell",
    "label",
    "caption",
    "shape",
    "footnote",
    "header",
    "footer",
    "unknown",
]

# Shared placeholder syntax already used by the existing DOCX/XLSX/PDF paths.
# Keeping it here gives every format adapter the same parity rules.
_PROTECTED_TOKEN_RE = re.compile(r"\[\[[NU]\d+\]\]|⟦G\d+⟧|⟪SEP⟫")
_HAN_CHAR_RE = re.compile(r"[\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF]")
_LEXICAL_WORD_RE = re.compile(r"[^\W\d_]+", re.UNICODE)


_EMAIL_RE = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")
_URL_RE = re.compile(r"^(?:https?://|ftp://|www\.)\S+$", re.IGNORECASE)
_VERSION_RE = re.compile(r"^[vV]?\d+(?:\.\d+){1,}(?:[-+][A-Za-z0-9._-]+)?$")
_NUMERIC_UNIT_RE = re.compile(
    r"^[\s~≈<>≤≥+\-−]?[$€£¥₹]?\s*\d[\d\s.,:/%-]*\s*(?:%|°[CF]|[kmcgµun]?m|kg|lb|oz|ml|l|hz|khz|mhz|ghz|w|kw|v|a|mah|wh|pa|kpa|mpa|bar|psi|s|ms|min|h|hr|hrs|gb|mb|kb|tb)?\s*$",
    re.IGNORECASE,
)
_FILENAME_RE = re.compile(r"^[^\s/\\]+\.[A-Za-z0-9]{1,8}$")
_CODELIKE_RE = re.compile(r"^[A-Za-z0-9]+(?:[-_:/+.][A-Za-z0-9]+)+$")

# Strong, intentionally small Latin-language fingerprints.  These are used only
# when the evidence is unambiguous enough to justify *skipping* translation.
_LATIN_STOPWORDS: dict[str, frozenset[str]] = {
    "en": frozenset({"the", "and", "is", "are", "of", "to", "in", "for", "with", "this", "that", "not", "you", "your"}),
    "pt": frozenset({"o", "a", "os", "as", "de", "do", "da", "dos", "das", "e", "é", "para", "com", "não", "que", "uma", "um", "por", "se"}),
    "es": frozenset({"el", "la", "los", "las", "de", "del", "y", "es", "para", "con", "no", "que", "una", "un", "por"}),
    "fr": frozenset({"le", "la", "les", "de", "du", "des", "et", "est", "pour", "avec", "pas", "que", "une", "un"}),
    "de": frozenset({"der", "die", "das", "und", "ist", "für", "mit", "nicht", "ein", "eine", "von", "zu"}),
    "it": frozenset({"il", "lo", "la", "i", "gli", "le", "di", "e", "è", "per", "con", "non", "che", "una", "un"}),
    "nl": frozenset({"de", "het", "een", "en", "is", "voor", "met", "niet", "van", "op"}),
}


def _language_base(code: str | None) -> str:
    value = str(code or "").strip().lower().replace("_", "-")
    return value.split("-", 1)[0]


def _normalized_echo_text(text: str) -> str:
    value = unicodedata.normalize("NFKC", str(text or "")).casefold()
    return "".join(ch for ch in value if ch.isalnum())


def _translation_memory_normalize_text(text: str) -> str:
    value = unicodedata.normalize("NFKC", str(text or ""))
    return value.replace("\r\n", "\n").replace("\r", "\n").strip()


def _translation_memory_normalize_value(value: Any) -> Any:
    if isinstance(value, str):
        return _translation_memory_normalize_text(value)
    if isinstance(value, Mapping):
        return {str(key): _translation_memory_normalize_value(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_translation_memory_normalize_value(item) for item in value]
    return value


def _term_present(text: str, term: str) -> bool:
    haystack = unicodedata.normalize("NFKC", str(text or "")).casefold()
    needle = unicodedata.normalize("NFKC", str(term or "")).casefold().strip()
    if not needle:
        return False
    # CJK and non-ASCII terminology is commonly written without word separators,
    # so substring matching is more reliable than Unicode word boundaries there.
    if any(ord(ch) > 127 for ch in needle):
        return needle in haystack
    if needle[0].isalnum() and needle[-1].isalnum():
        return re.search(r"(?<!\w)" + re.escape(needle) + r"(?!\w)", haystack) is not None
    return needle in haystack


def _looks_like_source_echo(
    source_text: str,
    translated_text: str,
    source_lang: str,
    target_lang: str,
) -> bool:
    source_base = _language_base(source_lang)
    target_base = _language_base(target_lang)
    if not source_base or source_base == "auto" or not target_base or source_base == target_base:
        return False
    if _normalized_echo_text(source_text) != _normalized_echo_text(translated_text):
        return False

    source = str(source_text or "")
    han_count = len(_HAN_CHAR_RE.findall(source))
    if han_count >= 4:
        return True

    words = _LEXICAL_WORD_RE.findall(source)
    letters = sum(len(word) for word in words)
    # Deliberately conservative: do not reject unchanged product names, acronyms,
    # model identifiers, or short labels that may legitimately remain unchanged.
    return len(words) >= 2 and letters >= 12


@dataclass(frozen=True)
class LanguageDetection:
    language: str | None = None
    confidence: float = 0.0
    method: str = "none"


@dataclass(frozen=True)
class TranslationEligibilityDecision:
    translate: bool
    reason: str = "translate"
    detected_language: str | None = None
    confidence: float = 0.0
    method: str = "none"
    effective_source_lang: str | None = None


def _unicode_script_counts(text: str) -> dict[str, int]:
    counts = {
        "latin": 0,
        "cyrillic": 0,
        "han": 0,
        "kana": 0,
        "hangul": 0,
        "arabic": 0,
        "hebrew": 0,
        "greek": 0,
        "devanagari": 0,
        "thai": 0,
        "other": 0,
    }
    for ch in str(text or ""):
        cp = ord(ch)
        if 0x3040 <= cp <= 0x30FF or 0x31F0 <= cp <= 0x31FF:
            counts["kana"] += 1
        elif 0xAC00 <= cp <= 0xD7AF or 0x1100 <= cp <= 0x11FF:
            counts["hangul"] += 1
        elif 0x3400 <= cp <= 0x4DBF or 0x4E00 <= cp <= 0x9FFF or 0xF900 <= cp <= 0xFAFF:
            counts["han"] += 1
        elif 0x0400 <= cp <= 0x052F:
            counts["cyrillic"] += 1
        elif 0x0600 <= cp <= 0x06FF or 0x0750 <= cp <= 0x077F or 0x08A0 <= cp <= 0x08FF:
            counts["arabic"] += 1
        elif 0x0590 <= cp <= 0x05FF:
            counts["hebrew"] += 1
        elif 0x0370 <= cp <= 0x03FF or 0x1F00 <= cp <= 0x1FFF:
            counts["greek"] += 1
        elif 0x0900 <= cp <= 0x097F:
            counts["devanagari"] += 1
        elif 0x0E00 <= cp <= 0x0E7F:
            counts["thai"] += 1
        elif ch.isalpha():
            name = unicodedata.name(ch, "")
            if "LATIN" in name:
                counts["latin"] += 1
            else:
                counts["other"] += 1
    return counts


def _dominant_script(text: str) -> tuple[str | None, float, int]:
    counts = _unicode_script_counts(text)
    total = sum(counts.values())
    if total <= 0:
        return None, 0.0, 0
    script, count = max(counts.items(), key=lambda item: item[1])
    return script, count / total, total


def _language_script_family(code: str | None) -> str | None:
    base = _language_base(code)
    if base in {"zh"}:
        return "han"
    if base in {"ja"}:
        return "japanese"
    if base in {"ko"}:
        return "hangul"
    if base in {"ru", "uk", "bg", "be", "mk", "sr"}:
        return "cyrillic"
    if base in {"ar", "fa", "ur"}:
        return "arabic"
    if base in {"he", "yi"}:
        return "hebrew"
    if base in {"el"}:
        return "greek"
    if base in {"hi", "mr", "ne"}:
        return "devanagari"
    if base in {"th"}:
        return "thai"
    if base and base != "auto":
        return "latin"
    return None


def _is_nonlinguistic_text(text: str) -> bool:
    value = str(text or "").strip()
    if not value:
        return True
    if not any(ch.isalpha() for ch in value):
        return True
    if _EMAIL_RE.fullmatch(value) or _URL_RE.fullmatch(value):
        return True
    if _VERSION_RE.fullmatch(value) or _NUMERIC_UNIT_RE.fullmatch(value):
        return True
    if _FILENAME_RE.fullmatch(value):
        return True
    if len(value) <= 96 and not any(ch.isspace() for ch in value):
        if _CODELIKE_RE.fullmatch(value):
            letters = [ch for ch in value if ch.isalpha()]
            digits = [ch for ch in value if ch.isdigit()]
            punctuation = [ch for ch in value if not ch.isalnum()]
            if digits or (punctuation and letters and all(not ch.islower() for ch in letters)):
                return True
    return False


def _coerce_language_detection(value: Any, *, method: str) -> LanguageDetection | None:
    if value is None:
        return None
    language: Any = None
    confidence: Any = None
    if isinstance(value, str):
        language = value
        confidence = 1.0
    elif isinstance(value, Mapping):
        language = value.get("language") or value.get("lang") or value.get("code")
        confidence = value.get("confidence", value.get("score", value.get("probability", 1.0)))
    elif isinstance(value, (tuple, list)) and value:
        language = value[0]
        confidence = value[1] if len(value) > 1 else 1.0
    else:
        language = getattr(value, "language", getattr(value, "lang", None))
        confidence = getattr(value, "confidence", getattr(value, "score", 1.0))
    base = _language_base(language)
    if not base or base == "auto":
        return None
    try:
        score = max(0.0, min(float(confidence), 1.0))
    except (TypeError, ValueError):
        score = 0.0
    return LanguageDetection(base, score, method)


def _builtin_language_detection(text: str) -> LanguageDetection | None:
    value = unicodedata.normalize("NFKC", str(text or "")).strip()
    if not value:
        return None
    counts = _unicode_script_counts(value)
    total = sum(counts.values())
    if total <= 0:
        return None

    # Kana and Hangul are strong language-specific evidence. Han-only text is
    # deliberately a little less confident because Japanese can be kanji-only.
    if counts["kana"] >= 2:
        return LanguageDetection("ja", 0.995, "script")
    if counts["hangul"] >= 2:
        return LanguageDetection("ko", 0.995, "script")
    if counts["han"] >= 4 and counts["han"] / total >= 0.80:
        return LanguageDetection("zh", 0.94, "script")
    if counts["greek"] >= 3 and counts["greek"] / total >= 0.80:
        return LanguageDetection("el", 0.98, "script")
    if counts["hebrew"] >= 3 and counts["hebrew"] / total >= 0.80:
        return LanguageDetection("he", 0.97, "script")
    if counts["thai"] >= 3 and counts["thai"] / total >= 0.80:
        return LanguageDetection("th", 0.99, "script")

    words = [word.casefold() for word in _LEXICAL_WORD_RE.findall(value)]
    if len(words) >= 4 and counts["latin"] / total >= 0.80:
        scores = {lang: sum(word in stopwords for word in words) for lang, stopwords in _LATIN_STOPWORDS.items()}
        ranked = sorted(scores.items(), key=lambda item: item[1], reverse=True)
        if ranked and ranked[0][1] >= 2:
            winner, hits = ranked[0]
            runner_up = ranked[1][1] if len(ranked) > 1 else 0
            if hits >= runner_up + 2 or hits >= 4:
                confidence = min(0.99, 0.84 + 0.04 * hits)
                return LanguageDetection(winner, confidence, "lexical")

    # A few genuinely distinctive orthographic clues are useful for short text.
    folded = value.casefold()
    if any(ch in folded for ch in ("ã", "õ")) or re.search(r"(?:ção|ções)\b", folded):
        return LanguageDetection("pt", 0.98, "orthography")
    if "ñ" in folded or "¿" in value or "¡" in value:
        return LanguageDetection("es", 0.98, "orthography")
    if "ß" in folded:
        return LanguageDetection("de", 0.99, "orthography")
    if "œ" in folded:
        return LanguageDetection("fr", 0.99, "orthography")
    return None


class TranslationEligibilityPolicy:
    """Conservative format-neutral decision layer for selective translation."""

    def __init__(
        self,
        detector: Callable[[str], Any] | None = None,
        *,
        target_language_threshold: float = 0.92,
        non_source_threshold: float = 0.97,
        source_hint_threshold: float = 0.90,
    ) -> None:
        self.detector = detector
        self.target_language_threshold = float(target_language_threshold)
        self.non_source_threshold = float(non_source_threshold)
        self.source_hint_threshold = float(source_hint_threshold)

    def detect(self, text: str) -> LanguageDetection | None:
        if self.detector is not None:
            try:
                detected = _coerce_language_detection(self.detector(text), method="backend")
            except Exception:
                logger.debug("Language detector failed; using conservative built-in heuristics", exc_info=True)
            else:
                if detected is not None:
                    return detected
        return _builtin_language_detection(text)

    def decide(self, unit: "TranslationUnit", target_lang: str) -> TranslationEligibilityDecision:
        policy = str(unit.metadata.get("translation_policy", "") or "").strip().lower()
        if policy in {"never", "skip", "preserve"}:
            return TranslationEligibilityDecision(False, "explicit-skip")
        force = policy in {"always", "translate", "force"}

        visible_text = str(unit.text or "")
        if not force and _is_nonlinguistic_text(visible_text):
            return TranslationEligibilityDecision(False, "nonlinguistic")

        source_base = _language_base(unit.source_lang)
        target_base = _language_base(target_lang)
        detection = self.detect(visible_text)
        detected_lang = detection.language if detection else None
        confidence = detection.confidence if detection else 0.0
        method = detection.method if detection else "none"

        if not force and detection is not None and detected_lang == target_base and confidence >= self.target_language_threshold:
            return TranslationEligibilityDecision(
                False, "target-language", detected_lang, confidence, method, detected_lang
            )

        if not force and source_base and source_base != "auto":
            source_family = _language_script_family(source_base)
            dominant, ratio, letters = _dominant_script(visible_text)
            script_mismatch = False
            if letters >= 3 and ratio >= 0.80 and source_family is not None:
                if source_family == "japanese":
                    script_mismatch = dominant not in {"han", "kana"}
                elif source_family == "han":
                    script_mismatch = dominant not in {"han"}
                else:
                    script_mismatch = dominant != source_family
            if script_mismatch:
                return TranslationEligibilityDecision(
                    False, "non-source-language", detected_lang, max(confidence, ratio), "script-mismatch"
                )
            if detection is not None and detected_lang != source_base and confidence >= self.non_source_threshold:
                return TranslationEligibilityDecision(
                    False, "non-source-language", detected_lang, confidence, method
                )

        effective_source = source_base if source_base and source_base != "auto" else None
        if source_base == "auto" and detection is not None and confidence >= self.source_hint_threshold:
            effective_source = detected_lang
        return TranslationEligibilityDecision(
            True, "translate", detected_lang, confidence, method, effective_source
        )


@dataclass(frozen=True)
class TranslationConstraints:
    """Portable layout constraints attached to one translation unit."""

    max_chars: int | None = None
    max_lines: int | None = None
    source_visible_chars: int | None = None
    source_line_count: int | None = None

    def normalized(self, source_text: str) -> "TranslationConstraints":
        max_chars = max(int(self.max_chars or 0), 0) or None
        max_lines = max(int(self.max_lines or 0), 0) or None
        source_visible_chars = self.source_visible_chars
        if source_visible_chars is None:
            source_visible_chars = len(" ".join(source_text.split()))
        source_line_count = self.source_line_count
        if source_line_count is None:
            source_line_count = max(1, len(source_text.split("\n")))
        return TranslationConstraints(
            max_chars=max_chars,
            max_lines=max_lines,
            source_visible_chars=max(int(source_visible_chars or 0), 0),
            source_line_count=max(int(source_line_count or 1), 1),
        )

    def to_payload(self, source_text: str) -> dict[str, int]:
        value = self.normalized(source_text)
        payload: dict[str, int] = {}
        if value.max_chars is not None:
            payload["max_chars"] = value.max_chars
            # Backward-compatible aliases consumed by the current Ollama prompt.
            payload["capacity_chars"] = value.max_chars
            payload["source_visible_chars"] = int(value.source_visible_chars or 0)
            payload["line_count"] = int(value.source_line_count or 1)
        if value.max_lines is not None:
            payload["max_lines"] = value.max_lines
        return payload


@dataclass(frozen=True)
class TranslationContext:
    """Format-neutral semantic context. Phase 4 will populate this more deeply."""

    domain: str | None = None
    section: str | None = None
    before: tuple[str, ...] = ()
    after: tuple[str, ...] = ()
    metadata: tuple[tuple[str, str], ...] = ()
    terminology: tuple[tuple[str, str], ...] = ()

    @classmethod
    def from_values(
        cls,
        *,
        domain: str | None = None,
        section: str | None = None,
        before: Iterable[str] = (),
        after: Iterable[str] = (),
        metadata: Mapping[str, Any] | None = None,
        terminology: Mapping[str, Any] | None = None,
    ) -> "TranslationContext":
        meta = tuple(
            sorted((str(key), str(value)) for key, value in (metadata or {}).items())
        )
        terms = tuple(
            sorted(
                (str(key).strip(), str(value).strip())
                for key, value in (terminology or {}).items()
                if str(key).strip() and str(value).strip()
            )
        )
        return cls(
            domain=(domain or "").strip() or None,
            section=(section or "").strip() or None,
            before=tuple(str(item) for item in before if str(item).strip()),
            after=tuple(str(item) for item in after if str(item).strip()),
            metadata=meta,
            terminology=terms,
        )

    def to_payload(self) -> dict[str, Any]:
        payload: dict[str, Any] = {}
        if self.domain:
            payload["domain"] = self.domain
        if self.section:
            payload["section"] = self.section
        if self.before:
            payload["before"] = list(self.before)
        if self.after:
            payload["after"] = list(self.after)
        if self.metadata:
            payload["metadata"] = dict(self.metadata)
        if self.terminology:
            payload["terminology"] = dict(self.terminology)
        return payload


@dataclass(frozen=True)
class ProtectedText:
    """A masked source string plus the exact tokens that must survive translation."""

    text: str
    token_map: Mapping[str, str] = field(default_factory=dict)

    def expected_tokens(self) -> Counter[str]:
        expected = Counter(_PROTECTED_TOKEN_RE.findall(self.text))
        # Support adapter-defined opaque tokens even when they are outside the
        # built-in Verbilo token syntax.
        for token in self.token_map:
            if token not in expected:
                count = self.text.count(token)
                if count:
                    expected[token] = count
        return expected

    def tokens_match(self, translated_text: str) -> bool:
        expected = self.expected_tokens()
        actual = Counter(_PROTECTED_TOKEN_RE.findall(str(translated_text or "")))
        for token in self.token_map:
            if not _PROTECTED_TOKEN_RE.fullmatch(token):
                count = str(translated_text or "").count(token)
                if count:
                    actual[token] = count
        return actual == expected

    def restore(self, translated_text: str) -> str:
        restored = str(translated_text)
        for token, original in self.token_map.items():
            restored = restored.replace(token, original)
        return restored


@dataclass(frozen=True)
class TranslationQualityReport:
    hard_failures: tuple[str, ...] = ()
    warnings: tuple[str, ...] = ()
    terminology_required: int = 0
    terminology_matched: int = 0
    visible_chars: int = 0
    visible_lines: int = 0

    @property
    def valid(self) -> bool:
        return not self.hard_failures

    @property
    def failure_reason(self) -> str | None:
        return self.hard_failures[0] if self.hard_failures else None

    @property
    def terminology_score(self) -> float | None:
        if self.terminology_required <= 0:
            return None
        return self.terminology_matched / self.terminology_required


@dataclass(frozen=True)
class TranslationUnit:
    """The shared semantic unit passed between format adapters and translators."""

    text: str
    source_lang: str = "auto"
    role: TranslationRole | str = "body"
    mode: TranslationMode = "natural"
    constraints: TranslationConstraints = field(default_factory=TranslationConstraints)
    context: TranslationContext = field(default_factory=TranslationContext)
    protected: ProtectedText | None = None
    allow_unprotected_fallback: bool = False
    metadata: Mapping[str, Any] = field(default_factory=dict)

    @property
    def source_text(self) -> str:
        return self.protected.text if self.protected is not None else self.text

    def to_backend_block(self) -> dict[str, Any]:
        source_text = self.source_text
        role = str(self.role or "body")
        strategy = str(self.metadata.get("strategy", "semantic") or "semantic")
        content_hint = str(self.metadata.get("content_hint", role) or role)
        payload: dict[str, Any] = {
            "text": source_text,
            "source_lang": self.source_lang or "auto",
            "content_hint": content_hint,
            "role": role,
            "strategy": strategy,
            "translation_mode": self.mode,
        }
        payload.update(self.constraints.to_payload(source_text))
        context_payload = self.context.to_payload()
        if context_payload:
            payload["context"] = context_payload
        return payload

    def cache_identity(self, target_lang: str) -> str:
        return json.dumps(
            {
                "text": self.source_text,
                "source_lang": self.source_lang or "auto",
                "target_lang": target_lang,
                "role": str(self.role or "body"),
                "mode": self.mode,
                "constraints": self.constraints.to_payload(self.source_text),
                "context": self.context.to_payload(),
                "allow_unprotected_fallback": self.allow_unprotected_fallback,
                "strategy": str(self.metadata.get("strategy", "semantic") or "semantic"),
                "content_hint": str(
                    self.metadata.get("content_hint", self.role) or self.role or "body"
                ),
            },
            ensure_ascii=False,
            sort_keys=True,
        )

    def translation_memory_identity_json(self, target_lang: str) -> str:
        normalized_constraints = self.constraints.normalized(self.source_text)
        identity = {
            "schema": 1,
            "text": _translation_memory_normalize_text(self.source_text),
            "source_lang": _language_base(self.source_lang) or str(self.source_lang or "auto").lower(),
            "target_lang": _language_base(target_lang) or str(target_lang or "").lower(),
            "role": str(self.role or "body"),
            "mode": self.mode,
            "constraints": {
                "max_chars": normalized_constraints.max_chars,
                "max_lines": normalized_constraints.max_lines,
            },
            "context": _translation_memory_normalize_value(self.context.to_payload()),
            "strategy": str(self.metadata.get("strategy", "semantic") or "semantic"),
            "content_hint": str(self.metadata.get("content_hint", self.role) or self.role or "body"),
        }
        return json.dumps(identity, ensure_ascii=False, sort_keys=True, separators=(",", ":"))

    def translation_memory_identity(self, target_lang: str) -> str:
        identity_json = self.translation_memory_identity_json(target_lang)
        return hashlib.sha256(identity_json.encode("utf-8")).hexdigest()

    def output_text(self, translated_text: str) -> str:
        if self.protected is not None:
            return self.protected.restore(translated_text)
        return str(translated_text)

    def evaluate_translation(
        self,
        translated_text: str | None,
        target_lang: str = "",
        *,
        enforce_terminology: bool = False,
        detect_source_echo: bool = True,
    ) -> TranslationQualityReport:
        failures: list[str] = []
        warnings: list[str] = []

        raw = "" if translated_text is None else str(translated_text)
        if translated_text is None:
            failures.append("missing")
        elif self.source_text.strip() and not raw.strip():
            failures.append("empty")

        protected = self.protected or ProtectedText(self.source_text)
        tokens_ok = protected.tokens_match(raw)
        if translated_text is not None and not tokens_ok:
            failures.append("protected-token-mismatch")

        visible = self.output_text(raw) if translated_text is not None else ""
        if (
            translated_text is not None
            and raw.strip()
            and detect_source_echo
            and _looks_like_source_echo(self.text, visible, self.source_lang, target_lang)
        ):
            failures.append("source-echo")

        terminology_required = 0
        terminology_matched = 0
        missing_terms: list[tuple[str, str]] = []
        for source_term, target_term in self.context.terminology:
            if not _term_present(self.text, source_term):
                continue
            terminology_required += 1
            if _term_present(visible, target_term):
                terminology_matched += 1
            else:
                missing_terms.append((source_term, target_term))

        if missing_terms:
            if enforce_terminology:
                failures.append("terminology-mismatch")
            else:
                warnings.append("terminology-mismatch")

        normalized_constraints = self.constraints.normalized(self.source_text)
        visible_chars = len(" ".join(visible.split())) if visible else 0
        visible_lines = max(1, len(visible.splitlines())) if visible else 0
        if (
            normalized_constraints.max_chars is not None
            and visible_chars > normalized_constraints.max_chars
        ):
            warnings.append("max-chars-exceeded")
        if (
            normalized_constraints.max_lines is not None
            and visible_lines > normalized_constraints.max_lines
        ):
            warnings.append("max-lines-exceeded")

        return TranslationQualityReport(
            hard_failures=tuple(dict.fromkeys(failures)),
            warnings=tuple(dict.fromkeys(warnings)),
            terminology_required=terminology_required,
            terminology_matched=terminology_matched,
            visible_chars=visible_chars,
            visible_lines=visible_lines,
        )

    def validate_translation(
        self,
        translated_text: str | None,
        target_lang: str = "",
        *,
        enforce_terminology: bool = False,
        detect_source_echo: bool = True,
    ) -> str | None:
        return self.evaluate_translation(
            translated_text,
            target_lang,
            enforce_terminology=enforce_terminology,
            detect_source_echo=detect_source_echo,
        ).failure_reason


@dataclass
class TranslationMetrics:
    units: int = 0
    translated_units: int = 0
    skipped_units: int = 0
    nonlinguistic_skips: int = 0
    target_language_skips: int = 0
    non_source_skips: int = 0
    explicit_skips: int = 0
    detected_units: int = 0
    unique_units: int = 0
    cache_hits: int = 0
    backend_requests: int = 0
    retry_items: int = 0
    fallback_items: int = 0
    unprotected_retries: int = 0
    failed_items: int = 0
    terminology_mismatches: int = 0
    source_echo_rejections: int = 0
    constraint_warnings: int = 0
    tm_hits: int = 0
    tm_misses: int = 0
    tm_writes: int = 0
    tm_rejected: int = 0
    tm_errors: int = 0


@dataclass
class TranslationBatchResult:
    texts: list[str | None]
    engines: list[str | None]
    metrics: TranslationMetrics
    quality_reports: list[TranslationQualityReport | None] = field(default_factory=list)
    eligibility: list[TranslationEligibilityDecision | None] = field(default_factory=list)

    @property
    def failed_indices(self) -> list[int]:
        return [idx for idx, value in enumerate(self.texts) if value is None]


class TranslationService:
    """Shared translation orchestration for all format adapters.

    The service intentionally knows nothing about PDF rectangles, Word runs, or
    spreadsheet XML. Adapters provide TranslationUnit objects and keep write-back
    responsibility. Translator-specific persistent caches remain inside backends;
    this layer provides format-neutral dedupe/L1 reuse, persistent translation
    memory, validation, retry/fallback, engine attribution, and metrics.
    """

    def __init__(
        self,
        translator: Any,
        *,
        fallback_translator: Any | None = None,
        cache: dict[tuple[str, str], str] | None = None,
        terminology: Mapping[str, Any] | None = None,
        language_detector: Callable[[str], Any] | None = None,
        eligibility_policy: TranslationEligibilityPolicy | None = None,
        translation_memory: TranslationMemory | None = None,
    ) -> None:
        self.translator = translator
        self.fallback_translator = fallback_translator
        self.translation_memory = translation_memory
        self._cache = cache if cache is not None else {}
        self._terminology = {
            str(key).strip(): str(value).strip()
            for key, value in (terminology or {}).items()
            if str(key).strip() and str(value).strip()
        }
        if eligibility_policy is not None:
            self.eligibility_policy = eligibility_policy
        else:
            detector = language_detector
            if detector is None:
                candidate = getattr(translator, "detect_language", None)
                if callable(candidate):
                    detector = candidate
            self.eligibility_policy = TranslationEligibilityPolicy(detector)
        self._eligibility_cache: dict[tuple[str, str, str, str], TranslationEligibilityDecision] = {}
        self.last_metrics = TranslationMetrics()

    def _with_service_terminology(self, unit: TranslationUnit) -> TranslationUnit:
        if not self._terminology:
            return unit
        merged = dict(self._terminology)
        # Unit-level guidance is more specific and therefore wins.
        merged.update(dict(unit.context.terminology))
        normalized = tuple(sorted(merged.items()))
        if normalized == unit.context.terminology:
            return unit
        return replace(unit, context=replace(unit.context, terminology=normalized))

    @staticmethod
    def _engine_name(translator: Any) -> str:
        if translator is None:
            return ""
        for attr in ("_engine_name", "engine_name", "name"):
            value = getattr(translator, attr, None)
            if isinstance(value, str) and value.strip():
                return value.strip()
        return type(translator).__name__

    @staticmethod
    def _check_cancel(cancel_event: threading.Event | None) -> None:
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError("Translation cancelled")

    def _eligibility_decision(
        self, unit: TranslationUnit, target_lang: str
    ) -> TranslationEligibilityDecision:
        policy = str(unit.metadata.get("translation_policy", "") or "").strip().lower()
        key = (
            str(unit.text or ""),
            _language_base(unit.source_lang),
            _language_base(target_lang),
            policy,
        )
        cached = self._eligibility_cache.get(key)
        if cached is not None:
            return cached
        decision = self.eligibility_policy.decide(unit, target_lang)
        # TranslationService instances are normally document-scoped. Keep the
        # cache bounded anyway so a long-lived caller cannot grow it forever.
        if len(self._eligibility_cache) >= 8192:
            self._eligibility_cache.pop(next(iter(self._eligibility_cache)))
        self._eligibility_cache[key] = decision
        return decision

    def _dispatch_many(
        self,
        translator: Any,
        units: Sequence[TranslationUnit],
        target_lang: str,
        cancel_event: threading.Event | None,
    ) -> list[str | None]:
        self._check_cancel(cancel_event)
        translate_units = getattr(translator, "translate_units", None)
        if callable(translate_units):
            result = translate_units(list(units), target_lang, cancel_event=cancel_event)
        else:
            translate_blocks = getattr(translator, "translate_blocks", None)
            if callable(translate_blocks):
                result = translate_blocks(
                    [unit.to_backend_block() for unit in units],
                    target_lang,
                    cancel_event=cancel_event,
                )
            else:
                translate_batch = getattr(translator, "translate_batch", None)
                if not callable(translate_batch):
                    return [self._dispatch_one(translator, unit, target_lang, cancel_event) for unit in units]
                result = translate_batch(
                    [unit.source_text for unit in units],
                    target_lang,
                    cancel_event=cancel_event,
                )

        items = result if isinstance(result, (list, tuple)) else []
        return [
            None if idx >= len(items) or items[idx] is None else str(items[idx])
            for idx in range(len(units))
        ]

    def _dispatch_one(
        self,
        translator: Any,
        unit: TranslationUnit,
        target_lang: str,
        cancel_event: threading.Event | None,
    ) -> str | None:
        self._check_cancel(cancel_event)
        translate_units = getattr(translator, "translate_units", None)
        if callable(translate_units):
            result = translate_units([unit], target_lang, cancel_event=cancel_event)
            return str(result[0]) if result and result[0] is not None else None

        translate_blocks = getattr(translator, "translate_blocks", None)
        if callable(translate_blocks):
            result = translate_blocks(
                [unit.to_backend_block()], target_lang, cancel_event=cancel_event
            )
            return str(result[0]) if result and result[0] is not None else None

        translate_text = getattr(translator, "translate_text", None)
        if callable(translate_text):
            return_value = translate_text(unit.source_text, target_lang)
            return None if return_value is None else str(return_value)

        translate_batch = getattr(translator, "translate_batch", None)
        if callable(translate_batch):
            result = translate_batch([unit.source_text], target_lang, cancel_event=cancel_event)
            return str(result[0]) if result and result[0] is not None else None
        raise AttributeError(
            f"{type(translator).__name__} exposes no supported translation method"
        )

    def translate_units(
        self,
        units: Sequence[TranslationUnit],
        target_lang: str,
        *,
        cancel_event: threading.Event | None = None,
    ) -> TranslationBatchResult:
        metrics = TranslationMetrics(units=len(units))
        self.last_metrics = metrics
        if not units:
            return TranslationBatchResult([], [], metrics, [], [])

        units = [self._with_service_terminology(unit) for unit in units]
        texts: list[str | None] = [None] * len(units)
        engines: list[str | None] = [None] * len(units)
        quality_reports: list[TranslationQualityReport | None] = [None] * len(units)
        eligibility: list[TranslationEligibilityDecision | None] = [None] * len(units)
        primary_engine = self._engine_name(self.translator)
        fallback_engine = self._engine_name(self.fallback_translator)
        primary_can_batch = any(
            callable(getattr(self.translator, name, None))
            for name in ("translate_units", "translate_blocks", "translate_batch")
        )
        primary_enforces_terminology = bool(
            getattr(self.translator, "supports_terminology", False)
        )
        fallback_enforces_terminology = bool(
            getattr(self.fallback_translator, "supports_terminology", False)
        )

        def evaluate(
            unit: TranslationUnit,
            translated: str | None,
            *,
            structured: bool,
        ) -> TranslationQualityReport:
            return unit.evaluate_translation(
                translated,
                target_lang,
                enforce_terminology=structured,
                detect_source_echo=True,
            )

        def record_rejection(report: TranslationQualityReport, count: int = 1) -> None:
            if report.failure_reason == "terminology-mismatch":
                metrics.terminology_mismatches += count
            elif report.failure_reason == "source-echo":
                metrics.source_echo_rejections += count

        def record_final_quality(report: TranslationQualityReport, count: int) -> None:
            if any(
                warning in {"max-chars-exceeded", "max-lines-exceeded"}
                for warning in report.warnings
            ):
                metrics.constraint_warnings += count
            if "terminology-mismatch" in report.warnings:
                metrics.terminology_mismatches += count

        grouped: dict[str, list[int]] = {}
        unit_by_key: dict[str, TranslationUnit] = {}
        for idx, unit in enumerate(units):
            self._check_cancel(cancel_event)
            if not unit.source_text.strip():
                decision = TranslationEligibilityDecision(False, "nonlinguistic")
                eligibility[idx] = decision
                metrics.skipped_units += 1
                metrics.nonlinguistic_skips += 1
                raw_original = unit.source_text
                texts[idx] = unit.output_text(raw_original)
                engines[idx] = None
                quality_reports[idx] = unit.evaluate_translation(
                    raw_original, target_lang, enforce_terminology=False, detect_source_echo=False
                )
                continue

            decision = self._eligibility_decision(unit, target_lang)
            eligibility[idx] = decision
            if decision.detected_language is not None:
                metrics.detected_units += 1
            if not decision.translate:
                metrics.skipped_units += 1
                if decision.reason == "nonlinguistic":
                    metrics.nonlinguistic_skips += 1
                elif decision.reason == "target-language":
                    metrics.target_language_skips += 1
                elif decision.reason == "non-source-language":
                    metrics.non_source_skips += 1
                elif decision.reason == "explicit-skip":
                    metrics.explicit_skips += 1
                raw_original = unit.source_text
                texts[idx] = unit.output_text(raw_original)
                engines[idx] = None
                quality_reports[idx] = unit.evaluate_translation(
                    raw_original, target_lang, enforce_terminology=False, detect_source_echo=False
                )
                continue

            if (
                decision.effective_source_lang
                and _language_base(unit.source_lang) == "auto"
                and decision.effective_source_lang != "auto"
            ):
                unit = replace(unit, source_lang=decision.effective_source_lang)
                units[idx] = unit
            metrics.translated_units += 1
            key = unit.cache_identity(target_lang)
            grouped.setdefault(key, []).append(idx)
            unit_by_key.setdefault(key, unit)

        metrics.unique_units = len(grouped)
        pending_keys: list[str] = []
        for key, indices in grouped.items():
            cached = self._cache.get((primary_engine, key))
            unit = unit_by_key[key]
            cached_report = (
                evaluate(unit, cached, structured=primary_enforces_terminology)
                if cached is not None
                else None
            )
            if cached is not None and cached_report is not None and cached_report.valid:
                metrics.cache_hits += 1
                visible_cached = unit.output_text(cached)
                record_final_quality(cached_report, len(indices))
                for idx in indices:
                    texts[idx] = visible_cached
                    engines[idx] = primary_engine
                    quality_reports[idx] = cached_report
            else:
                if cached_report is not None:
                    record_rejection(cached_report, len(indices))
                pending_keys.append(key)

        if pending_keys and self.translation_memory is not None:
            tm_key_by_group = {
                key: unit_by_key[key].translation_memory_identity(target_lang)
                for key in pending_keys
            }
            try:
                tm_entries = self.translation_memory.get_many(tm_key_by_group.values())
            except Exception:
                logger.warning("Translation memory lookup failed; continuing without it", exc_info=True)
                metrics.tm_errors += 1
                tm_entries = {}
            still_pending: list[str] = []
            rejected_tm_keys: list[str] = []
            for key in pending_keys:
                tm_key = tm_key_by_group[key]
                entry = tm_entries.get(tm_key)
                if entry is None:
                    metrics.tm_misses += 1
                    still_pending.append(key)
                    continue
                unit = unit_by_key[key]
                report = evaluate(unit, entry.translation, structured=True)
                if not report.valid:
                    metrics.tm_rejected += 1
                    record_rejection(report)
                    rejected_tm_keys.append(tm_key)
                    still_pending.append(key)
                    continue
                metrics.tm_hits += 1
                record_final_quality(report, len(grouped[key]))
                visible_translation = unit.output_text(entry.translation)
                for idx in grouped[key]:
                    texts[idx] = visible_translation
                    engines[idx] = entry.engine or "translation-memory"
                    quality_reports[idx] = report
            if rejected_tm_keys:
                try:
                    self.translation_memory.delete_many(rejected_tm_keys)
                except Exception:
                    logger.debug("Could not delete rejected translation-memory entries", exc_info=True)
                    metrics.tm_errors += 1
            pending_keys = still_pending

        pending_tm_writes: list[TranslationMemoryEntry] = []
        if pending_keys:
            pending_units = [unit_by_key[key] for key in pending_keys]
            try:
                metrics.backend_requests += 1
                primary_results = self._dispatch_many(
                    self.translator, pending_units, target_lang, cancel_event
                )
            except CancelledError:
                raise
            except Exception:
                logger.warning(
                    "TranslationService batch dispatch failed; retrying units individually",
                    exc_info=True,
                )
                primary_results = [None] * len(pending_units)

            for local_idx, key in enumerate(pending_keys):
                self._check_cancel(cancel_event)
                unit = unit_by_key[key]
                translated = primary_results[local_idx] if local_idx < len(primary_results) else None
                report = evaluate(unit, translated, structured=primary_enforces_terminology)
                failure_reason = report.failure_reason
                if failure_reason is not None:
                    record_rejection(report)

                # A corrupted protected token should not be sent through the same
                # masked request again when the adapter explicitly allows an
                # unprotected recovery. XLSX uses this for glossary placeholders.
                retry_masked = not (
                    failure_reason == "protected-token-mismatch"
                    and unit.protected is not None
                    and unit.allow_unprotected_fallback
                    and unit.text != unit.source_text
                )
                if failure_reason is not None and retry_masked and primary_can_batch:
                    metrics.retry_items += 1
                    try:
                        metrics.backend_requests += 1
                        translated = self._dispatch_one(
                            self.translator, unit, target_lang, cancel_event
                        )
                    except CancelledError:
                        raise
                    except Exception:
                        logger.debug(
                            "TranslationService per-item retry failed", exc_info=True
                        )
                        translated = None
                    report = evaluate(unit, translated, structured=primary_enforces_terminology)
                    failure_reason = report.failure_reason
                    if failure_reason is not None:
                        record_rejection(report)

                used_engine = primary_engine
                used_structured = primary_enforces_terminology
                if (
                    failure_reason is not None
                    and unit.protected is not None
                    and unit.allow_unprotected_fallback
                    and unit.text != unit.source_text
                ):
                    metrics.unprotected_retries += 1
                    unprotected_unit = replace(
                        unit,
                        protected=None,
                        allow_unprotected_fallback=False,
                    )
                    try:
                        metrics.backend_requests += 1
                        translated = self._dispatch_one(
                            self.translator, unprotected_unit, target_lang, cancel_event
                        )
                    except CancelledError:
                        raise
                    except Exception:
                        logger.debug(
                            "TranslationService unprotected retry failed", exc_info=True
                        )
                        translated = None
                    report = evaluate(
                        unprotected_unit, translated, structured=primary_enforces_terminology
                    )
                    failure_reason = report.failure_reason
                    if failure_reason is not None:
                        record_rejection(report)
                    else:
                        # The successful result corresponds to the unprotected source,
                        # so do not attempt placeholder restoration below.
                        unit = unprotected_unit

                if failure_reason is not None and self.fallback_translator is not None:
                    metrics.fallback_items += 1
                    try:
                        metrics.backend_requests += 1
                        translated = self._dispatch_one(
                            self.fallback_translator, unit, target_lang, cancel_event
                        )
                    except CancelledError:
                        raise
                    except Exception:
                        logger.debug(
                            "TranslationService fallback translator failed", exc_info=True
                        )
                        translated = None
                    report = evaluate(unit, translated, structured=fallback_enforces_terminology)
                    failure_reason = report.failure_reason
                    if failure_reason is not None:
                        record_rejection(report)
                    used_engine = fallback_engine
                    used_structured = fallback_enforces_terminology

                if failure_reason is not None or translated is None:
                    metrics.failed_items += len(grouped[key])
                    for idx in grouped[key]:
                        quality_reports[idx] = report
                    continue

                translated = str(translated)
                original_unit = unit_by_key[key]
                if (
                    used_engine == primary_engine
                    and unit.source_text == original_unit.source_text
                ):
                    self._cache[(primary_engine, key)] = translated
                if (
                    self.translation_memory is not None
                    and unit.source_text == original_unit.source_text
                ):
                    identity = original_unit.translation_memory_identity(target_lang)
                    identity_json = original_unit.translation_memory_identity_json(target_lang)
                    pending_tm_writes.append(
                        TranslationMemoryEntry(
                            key=identity,
                            identity_json=identity_json,
                            source_text=original_unit.source_text,
                            source_lang=_language_base(original_unit.source_lang) or str(original_unit.source_lang or "auto"),
                            target_lang=_language_base(target_lang) or str(target_lang or ""),
                            role=str(original_unit.role or "body"),
                            mode=str(original_unit.mode or "natural"),
                            translation=translated,
                            engine=used_engine,
                        )
                    )
                visible_translation = unit.output_text(translated)
                # Re-evaluate the final chosen output once for diagnostics. This is
                # intentionally non-destructive for layout warnings.
                final_report = evaluate(unit, translated, structured=used_structured)
                record_final_quality(final_report, len(grouped[key]))
                for idx in grouped[key]:
                    texts[idx] = visible_translation
                    engines[idx] = used_engine
                    quality_reports[idx] = final_report

        if pending_tm_writes and self.translation_memory is not None:
            self._check_cancel(cancel_event)
            # Commit only after every candidate has passed validation. If the
            # operation is cancelled or raises earlier, nothing from this batch is
            # persisted to the translation memory.
            try:
                self.translation_memory.put_many(pending_tm_writes)
                metrics.tm_writes += len(pending_tm_writes)
            except Exception:
                logger.warning("Translation memory write failed; translation output is still valid", exc_info=True)
                metrics.tm_errors += 1

        return TranslationBatchResult(texts, engines, metrics, quality_reports, eligibility)
