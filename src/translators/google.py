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
# Minimum text length (chars) for target-language detection to be trusted.
_SKIP_MIN_CHARS = 20


# ---------------------------------------------------------------------------
# Lightweight post-processing: fix common Google Translate artefacts
# ---------------------------------------------------------------------------

def post_process(text: str) -> str:
    # fix spacing around punctuation and collapse multiple spaces
    if not text:
        return text

    # Space before comma / period / colon / semicolon / closing paren
    text = re.sub(r'\s+([,.:;!?\)}\]])', r'\1', text)

    # Missing space after comma, semicolon, colon (but not inside numbers like "1,000" or "12:30")
    text = re.sub(r'([,;])(?=[^\s\d])', r'\1 ', text)
    text = re.sub(r'(:)(?=[^\s\d/\\])', r'\1 ', text)

    # Missing space after sentence-ending punctuation followed by a letter
    text = re.sub(r'([.!?])(?=[A-Za-z\u00C0-\u024F\u0400-\u04FF\u4e00-\u9fff])', r'\1 ', text)

    # Collapse multiple spaces into one
    text = re.sub(r'  +', ' ', text)

    return text.strip()


def _run_cancellable(fn, cancel_event: Optional[threading.Event], poll_interval: float = 0.05):
    # run fn() in a daemon thread; raises CancelledError if cancel_event fires mid-run
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


# wraps GoogleTranslator with batching, dedup cache, and post-processing
# always sends source="auto" to the API; source_lang only skips already-target-lang text
class DeepTranslatorWrapper:

    def __init__(self, source_lang: str = "auto"):
        self._source_lang = source_lang
        try:
            from deep_translator import GoogleTranslator  # type: ignore
            self._impl_cls = GoogleTranslator
        except Exception:
            self._impl_cls = None
        self._instances: Dict[str, Any] = {}
        # Per-target translation cache   text -> translated text
        self._cache: Dict[str, Dict[str, str]] = {}
        # Optional langdetect for target-language skip
        self._langdetect = None
        try:
            import langdetect as _ld
            _ld.DetectorFactory.seed = 0  # deterministic
            self._langdetect = _ld
        except ImportError:
            pass

    def _get_instance(self, target_lang: str):
        inst = self._instances.get(target_lang)
        if inst is None:
            if self._impl_cls is None:
                raise RuntimeError("deep_translator is not available")
            # always source="auto"; source_lang only for the target-skip check
            inst = self._impl_cls(source="auto", target=target_lang)
            self._instances[target_lang] = inst
        return inst

    def _is_already_target_lang(self, text: str, target_lang: str) -> bool:
        # True if text is already in target_lang (only when langdetect available and text long enough)
        if self._langdetect is None:
            return False
        if len(text.strip()) < _SKIP_MIN_CHARS:
            return False
        try:
            detected = self._langdetect.detect(text)
            return detected.lower().startswith(target_lang.lower())
        except Exception:
            return False

    # ----- single-item convenience ----- #

    def translate_text(self, text: str, target_lang: str) -> str:
        if not self._impl_cls or not text or not text.strip():
            return text
        # Check cache
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
        to_translate: list[tuple[int, str]] = []  # (original_index, text)
        for i, t in enumerate(texts):
            if not t or not t.strip():
                continue  # results[i] already == t
            if self._is_already_target_lang(t, target_lang):
                continue  # already in the target language
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
            # Honour cancellation before starting the next HTTP request
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
                # Pad if Google Translate returned fewer items than sent
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
        # on batch failure, retry in halves; per-item only for the smallest failing chunk
        if len(chunk) <= 2:
            # Small enough — do per-item
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
                    # results already contains the original text
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
        # split into chunks capped at _BATCH_SIZE items and _BATCH_MAX_CHARS chars
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
    def get(name: Optional[str] = None, source_lang: str = "auto") -> Translator:
        if name is None:
            try:
                from deep_translator import GoogleTranslator  # type: ignore  # noqa: F401
                return DeepTranslatorWrapper(source_lang=source_lang)
            except Exception:
                logger.warning(
                    "deep_translator is not available — returning IdentityTranslator "
                    "(text will NOT be translated). Install it with: pip install deep-translator"
                )
                return IdentityTranslator()
        if name.lower() == "identity":
            return IdentityTranslator()
        if name.lower() in ("deep", "deep-translator", "google"):
            return DeepTranslatorWrapper(source_lang=source_lang)
        return IdentityTranslator()
