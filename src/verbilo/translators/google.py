# Google Translate backend via deep_translator — batching, caching, fallback

from __future__ import annotations

import re
import threading
from typing import Optional, Dict, Any
from .base import Translator
from ..utils import CancelledError
import logging

logger = logging.getLogger(__name__)

# Maximum texts per Google Translate batch request.
_BATCH_SIZE = 50

# Maximum total characters per batch – just under Google's ~5 000 char limit.
_BATCH_MAX_CHARS = 4900

# Valid detector names for the language-detection subsystem.
VALID_DETECTORS = ("auto", "fasttext", "lingua", "langdetect")


# Lightweight post-processing: fix common Google Translate artefacts
def post_process(text: str) -> str:
    # fix spacing around punctuation and collapse multiple spaces
    if not text:
        return text
    text = re.sub(r'\s+([,.:;!?\)}\]])', r'\1', text)
    text = re.sub(r'([,;])(?=[^\s\d])', r'\1 ', text)
    text = re.sub(r'(:)(?=[^\s\d/\\])', r'\1 ', text)
    text = re.sub(r'([.!?])(?=[A-Za-z\u00C0-\u024F\u0400-\u04FF\u4e00-\u9fff])', r'\1 ', text)
    text = re.sub(r'  +', ' ', text)
    return text.strip()


def _run_cancellable(fn, cancel_event: Optional[threading.Event], poll_interval: float = 0.05):
    result = [None]
    exc: list = [None]

    def _target():
        try:
            result[0] = fn()
        except Exception as e:
            exc[0] = e

    t = threading.Thread(target=_target, daemon=True)
    t.start()
    while t.is_alive():
        t.join(timeout=poll_interval)
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError("Translation cancelled")
    if exc[0] is not None:
        raise exc[0]
    return result[0]


# returns text unchanged — for testing
class IdentityTranslator:

    def translate_text(self, text: str, target_lang: str) -> str:
        return text

    def translate_batch(
        self,
        texts: list[str],
        target_lang: str,
        *,
        cancel_event: Optional[threading.Event] = None,
    ) -> list[str]:
        return list(texts)


class DeepTranslatorWrapper:

    def __init__(self, source_lang: str = "auto", detector: str = "auto"):
        self._source_lang = source_lang
        self._detector = detector
        try:
            from deep_translator import GoogleTranslator
            self._impl_cls = GoogleTranslator
        except Exception:
            self._impl_cls = None
        self._instances: Dict[str, Any] = {}
        self._cache: Dict[str, Dict[str, str]] = {}

    def _get_instance(self, target_lang: str):
        inst = self._instances.get(target_lang)
        if inst is None:
            if self._impl_cls is None:
                raise RuntimeError("deep_translator is not available")
            # Always let Google auto-detect the source; pre-filtering handles
            # source-language selection so we never feed the wrong lang code.
            inst = self._impl_cls(source="auto", target=target_lang)
            self._instances[target_lang] = inst
        return inst

    def _should_translate(self, text: str) -> bool:
        """Return True if *text* should be sent to the translator.

        When source_lang == "auto" every cell is translated.  Otherwise the
        multi-engine detector decides whether the text is in the source
        language.
        """
        if self._source_lang == "auto":
            return True
        from .lang_detect import is_source_language
        return is_source_language(text, self._source_lang, detector=self._detector)

    # Pattern for splitting mixed-language cells exactly.
    _SEGMENT_RE = re.compile(r'(\n|\r\n|\r|/)')

    def _translate_segments(self, text: str, target_lang: str) -> str:
        """Split *text* on ``/`` and newlines, translate only the segments
        that are in the source language, and reassemble with the original
        separators.  Only used when ``source_lang != 'auto'``.
        """
        from .lang_detect import is_source_language
        parts = self._SEGMENT_RE.split(text)
        changed = False
        result_parts: list[str] = []
        for i, part in enumerate(parts):
            if i % 2 == 1:
                # separator — keep as-is
                result_parts.append(part)
                continue
            stripped = part.strip()
            if not stripped:
                result_parts.append(part)
                continue
            if is_source_language(stripped, self._source_lang, detector=self._detector, strict=True):
                translated = self._translate_single(stripped, target_lang)
                # Preserve leading/trailing whitespace from the original part
                leading = part[:len(part) - len(part.lstrip())]
                trailing = part[len(part.rstrip()):]
                result_parts.append(leading + translated + trailing)
                changed = True
            else:
                result_parts.append(part)
        return "".join(result_parts) if changed else text

    def _translate_single(self, text: str, target_lang: str) -> str:
        """Translate a single piece of text (no segment splitting)."""
        tgt_cache = self._cache.setdefault(target_lang, {})
        if text in tgt_cache:
            return tgt_cache[text]
        try:
            translator = self._get_instance(target_lang)
            result = translator.translate(text)
            result = result if result is not None else text
            result = post_process(result)
            tgt_cache[text] = result
            return result
        except Exception:
            logger.exception("DeepTranslator segment failed for target '%s'", target_lang)
            return text

    # ----- single-item convenience ----- #

    def translate_text(self, text: str, target_lang: str) -> str:
        if not self._impl_cls or not text or not text.strip():
            return text
        # When a specific source language is set, handle mixed-language cells
        # by splitting on / and newlines and translating only matching segments.
        if self._source_lang != "auto":
            if self._SEGMENT_RE.search(text):
                return self._translate_segments(text, target_lang)
            if not self._should_translate(text):
                return text
        tgt_cache = self._cache.setdefault(target_lang, {})
        if text in tgt_cache:
            return tgt_cache[text]
        try:
            translator = self._get_instance(target_lang)
            result = translator.translate(text)
            result = result if result is not None else text
            result = post_process(result)
            tgt_cache[text] = result
            return result
        except Exception:
            logger.exception("DeepTranslator failed for target '%s'", target_lang)
            raise

    # ----- batch (the fast path) ----- #

    def translate_batch(
        self,
        texts: list[str],
        target_lang: str,
        *,
        cancel_event: Optional[threading.Event] = None,
    ) -> list[str]:
        # batches requests; empty strings and already-target-lang text pass through
        if not self._impl_cls:
            return list(texts)

        # copy originals so untouched indices keep their original text
        results: list[str] = list(texts)
        tgt_cache = self._cache.setdefault(target_lang, {})

        # Separate translatable vs pass-through, resolving cache hits inline.
        to_translate: list[tuple[int, str]] = []
        for i, t in enumerate(texts):
            if not t or not t.strip():
                continue
            # Mixed-language cell: handle per-segment (not batchable)
            if self._source_lang != "auto" and self._SEGMENT_RE.search(t):
                results[i] = self._translate_segments(t, target_lang)
                continue
            if not self._should_translate(t):
                continue
            if t in tgt_cache:
                results[i] = tgt_cache[t]
            else:
                to_translate.append((i, t))

        if not to_translate:
            return results

        # Deduplicate: translate each unique string only once
        unique_texts: dict[str, list[int]] = {}  # text -> [indices]
        for idx, t in to_translate:
            unique_texts.setdefault(t, []).append(idx)

        dedup_items: list[tuple[str, list[int]]] = list(unique_texts.items())

        translator = self._get_instance(target_lang)

        # Split into chunks respecting both count and character limits
        chunks = self._make_chunks(dedup_items)

        for chunk in chunks:
            if cancel_event is not None and cancel_event.is_set():
                raise CancelledError("Translation cancelled")
            chunk_texts = [t for t, _ in chunk]
            try:
                translated = _run_cancellable(
                    lambda texts=chunk_texts: translator.translate_batch(texts),
                    cancel_event,
                )
                if translated is None:
                    translated = chunk_texts
                if len(translated) < len(chunk):
                    translated.extend(
                        chunk_texts[i] for i in range(len(translated), len(chunk))
                    )
                for (orig_text, indices), tr_text in zip(chunk, translated):
                    tr_text = tr_text if tr_text is not None else orig_text
                    tr_text = post_process(tr_text)
                    tgt_cache[orig_text] = tr_text
                    for idx in indices:
                        results[idx] = tr_text
            except CancelledError:
                raise
            except Exception:
                logger.exception("Batch translation failed; falling back to sub-batches")
                self._subbatch_fallback(chunk, translator, results, tgt_cache, cancel_event)

        return results

    def _subbatch_fallback(
        self,
        chunk: list[tuple[str, list[int]]],
        translator,
        results: list[str],
        tgt_cache: dict[str, str],
        cancel_event: Optional[threading.Event] = None,
    ) -> None:
        if len(chunk) <= 2:
            for orig_text, indices in chunk:
                if cancel_event is not None and cancel_event.is_set():
                    raise CancelledError("Translation cancelled")
                try:
                    r = _run_cancellable(
                        lambda text=orig_text: translator.translate(text),
                        cancel_event,
                    )
                    r = r if r is not None else orig_text
                    r = post_process(r)
                    tgt_cache[orig_text] = r
                    for idx in indices:
                        results[idx] = r
                except CancelledError:
                    raise
                except Exception:
                    logger.exception("Per-item fallback also failed")
            return
        mid = len(chunk) // 2
        for half in (chunk[:mid], chunk[mid:]):
            if cancel_event is not None and cancel_event.is_set():
                raise CancelledError("Translation cancelled")
            half_texts = [t for t, _ in half]
            try:
                translated = _run_cancellable(
                    lambda texts=half_texts: translator.translate_batch(texts),
                    cancel_event,
                )
                if translated is None:
                    translated = half_texts
                if len(translated) < len(half):
                    translated.extend(
                        half_texts[i] for i in range(len(translated), len(half))
                    )
                for (orig_text, indices), tr_text in zip(half, translated):
                    tr_text = tr_text if tr_text is not None else orig_text
                    tr_text = post_process(tr_text)
                    tgt_cache[orig_text] = tr_text
                    for idx in indices:
                        results[idx] = tr_text
            except CancelledError:
                raise
            except Exception:
                logger.exception("Sub-batch failed; recursing")
                self._subbatch_fallback(half, translator, results, tgt_cache, cancel_event)

    @staticmethod
    def _make_chunks(items: list[tuple[str, list[int]]]) -> list[list[tuple[str, list[int]]]]:
        chunks: list[list[tuple[str, list[int]]]] = []
        current: list[tuple[str, list[int]]] = []
        current_chars = 0
        for item in items:
            text_len = len(item[0])
            if current and (len(current) >= _BATCH_SIZE or current_chars + text_len > _BATCH_MAX_CHARS):
                chunks.append(current)
                current = []
                current_chars = 0
            current.append(item)
            current_chars += text_len
        if current:
            chunks.append(current)
        return chunks


class TranslatorFactory:
    @staticmethod
    def get(name: Optional[str] = None, source_lang: str = "auto", detector: str = "auto") -> Translator:
        if name is None:
            try:
                from deep_translator import GoogleTranslator  # type: ignore  # noqa: F401
                return DeepTranslatorWrapper(source_lang=source_lang, detector=detector)
            except Exception:
                logger.warning(
                    "deep_translator is not available — returning IdentityTranslator "
                    "(text will NOT be translated). Install it with: pip install deep-translator"
                )
                return IdentityTranslator()
        if name.lower() == "identity":
            return IdentityTranslator()
        if name.lower() in ("deep", "deep-translator", "google"):
            return DeepTranslatorWrapper(source_lang=source_lang, detector=detector)
        return IdentityTranslator()
