from __future__ import annotations

import json
import logging
import re
import threading
from collections import Counter
from dataclasses import dataclass, field, replace
from typing import Any, Iterable, Literal, Mapping, Sequence

from .utils import CancelledError

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

    def output_text(self, translated_text: str) -> str:
        if self.protected is not None:
            return self.protected.restore(translated_text)
        return str(translated_text)

    def validate_translation(self, translated_text: str | None) -> str | None:
        if translated_text is None:
            return "missing"
        if self.source_text.strip() and not str(translated_text).strip():
            return "empty"
        protected = self.protected or ProtectedText(self.source_text)
        if not protected.tokens_match(str(translated_text)):
            return "protected-token-mismatch"
        return None


@dataclass
class TranslationMetrics:
    units: int = 0
    unique_units: int = 0
    cache_hits: int = 0
    backend_requests: int = 0
    retry_items: int = 0
    fallback_items: int = 0
    unprotected_retries: int = 0
    failed_items: int = 0


@dataclass
class TranslationBatchResult:
    texts: list[str | None]
    engines: list[str | None]
    metrics: TranslationMetrics

    @property
    def failed_indices(self) -> list[int]:
        return [idx for idx, value in enumerate(self.texts) if value is None]


class TranslationService:
    """Shared translation orchestration for all format adapters.

    The service intentionally knows nothing about PDF rectangles, Word runs, or
    spreadsheet XML. Adapters provide TranslationUnit objects and keep write-back
    responsibility. Translator-specific persistent caches remain inside backends;
    this layer provides format-neutral dedupe/L1 reuse, validation, retry/fallback,
    engine attribution, and metrics.
    """

    def __init__(
        self,
        translator: Any,
        *,
        fallback_translator: Any | None = None,
        cache: dict[tuple[str, str], str] | None = None,
        terminology: Mapping[str, Any] | None = None,
    ) -> None:
        self.translator = translator
        self.fallback_translator = fallback_translator
        self._cache = cache if cache is not None else {}
        self._terminology = {
            str(key).strip(): str(value).strip()
            for key, value in (terminology or {}).items()
            if str(key).strip() and str(value).strip()
        }
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
            return TranslationBatchResult([], [], metrics)

        units = [self._with_service_terminology(unit) for unit in units]
        texts: list[str | None] = [None] * len(units)
        engines: list[str | None] = [None] * len(units)
        primary_engine = self._engine_name(self.translator)
        fallback_engine = self._engine_name(self.fallback_translator)
        primary_can_batch = any(
            callable(getattr(self.translator, name, None))
            for name in ("translate_units", "translate_blocks", "translate_batch")
        )

        grouped: dict[str, list[int]] = {}
        unit_by_key: dict[str, TranslationUnit] = {}
        for idx, unit in enumerate(units):
            self._check_cancel(cancel_event)
            if not unit.source_text.strip():
                texts[idx] = unit.source_text
                engines[idx] = primary_engine
                continue
            key = unit.cache_identity(target_lang)
            grouped.setdefault(key, []).append(idx)
            unit_by_key.setdefault(key, unit)

        metrics.unique_units = len(grouped)
        pending_keys: list[str] = []
        for key, indices in grouped.items():
            cached = self._cache.get((primary_engine, key))
            unit = unit_by_key[key]
            if cached is not None and unit.validate_translation(cached) is None:
                metrics.cache_hits += 1
                visible_cached = unit.output_text(cached)
                for idx in indices:
                    texts[idx] = visible_cached
                    engines[idx] = primary_engine
            else:
                pending_keys.append(key)

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
                failure_reason = unit.validate_translation(translated)

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
                    failure_reason = unit.validate_translation(translated)

                used_engine = primary_engine
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
                    failure_reason = unprotected_unit.validate_translation(translated)
                    if failure_reason is None:
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
                    failure_reason = unit.validate_translation(translated)
                    used_engine = fallback_engine

                if failure_reason is not None or translated is None:
                    metrics.failed_items += len(grouped[key])
                    continue

                translated = str(translated)
                original_unit = unit_by_key[key]
                if (
                    used_engine == primary_engine
                    and unit.source_text == original_unit.source_text
                ):
                    self._cache[(primary_engine, key)] = translated
                visible_translation = unit.output_text(translated)
                for idx in grouped[key]:
                    texts[idx] = visible_translation
                    engines[idx] = used_engine

        return TranslationBatchResult(texts, engines, metrics)
