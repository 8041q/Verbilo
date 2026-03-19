# Baidu Translate backend via deep_translator — batching, caching, fallback

from __future__ import annotations

import re
import threading
import types
from typing import Optional, Dict, Any
from .base import Translator
from .http_session import make_session, resolve_proxies
from ..utils import CancelledError
import logging

logger = logging.getLogger(__name__)

# Baidu free-tier QPS limit (1 query per second).
_BAIDU_QPS_DELAY = 1.0

# Maximum texts per batch request — Baidu processes one text per API call
# so we keep smaller batches to avoid hitting rate limits hard.
_BATCH_SIZE = 30

# Maximum total characters per batch.
_BATCH_MAX_CHARS = 4900

# Baidu uses non-standard language codes.  Map ISO 639-1 base codes (used by
# the rest of the app) to the codes Baidu's API actually expects.
_ISO_TO_BAIDU: dict[str, str] = {
    "ar": "ara", "bg": "bul", "cs": "cs", "da": "dan", "de": "de",
    "el": "el", "en": "en", "es": "spa", "et": "est", "fi": "fin",
    "fr": "fra", "hu": "hu", "it": "it", "ja": "jp", "ko": "kor",
    "nl": "nl", "pl": "pl", "pt": "pt", "ro": "ro", "ru": "ru",
    "sl": "slo", "sv": "swe", "th": "th", "vi": "vie",
    "zh-CN": "zh", "zh": "zh", "zh-TW": "cht",
}

# The set of ISO codes that Baidu supports (for language dropdown filtering).
BAIDU_LANG_CODES: frozenset[str] = frozenset(
    _ISO_TO_BAIDU.keys() | {
        # Add the Baidu-native codes so direct entries also match
        "ara", "bul", "dan", "est", "fin", "fra", "jp", "kor",
        "slo", "spa", "swe", "vie", "cht", "wyw", "yue",
    }
)


def _baidu_code(iso_code: str) -> str:
    """Convert an ISO 639-1 code to the Baidu API code."""
    return _ISO_TO_BAIDU.get(iso_code, iso_code)


# Re-use the same post-processing as the Google wrapper.
from .google import post_process, _run_cancellable


def _patch_baidu_requests(session):
    """Monkey-patch ``deep_translator.baidu.requests`` so that all HTTP calls
    go through *session* (which carries retry, timeout, and proxy settings)."""
    try:
        import deep_translator.baidu as _mod
        # Create a thin shim module that delegates .post / .get to the session
        shim = types.ModuleType("requests_shim")
        shim.post = session.post  # type: ignore[attr-defined]
        shim.get = session.get    # type: ignore[attr-defined]
        _mod.requests = shim      # type: ignore[attr-defined]
        logger.debug("Patched deep_translator.baidu.requests → resilient session")
    except Exception:
        logger.debug("Could not patch deep_translator.baidu.requests", exc_info=True)


class BaiduTranslatorWrapper:
    """Wraps :class:`deep_translator.BaiduTranslator` with caching, batching,
    fallback, cancellation, and resilient HTTP."""

    _engine_name = "baidu"

    def __init__(
        self,
        appid: str,
        appkey: str,
        source_lang: str = "auto",
        detector: str = "fasttext",
        proxies: Optional[dict] = None,
        tier: str = "standard",
    ):
        self._source_lang = source_lang
        self._detector = detector
        self._appid = appid
        self._appkey = appkey
        self._proxies = proxies
        self._tier = tier
        # Premium tier records under a separate engine key so it does not count
        # against the Standard free-tier limit (50K chars/month).
        self._engine_name = "baidu-premium" if tier == "premium" else "baidu"

        # Set up a resilient session and inject it into deep_translator.baidu
        self._session = make_session(proxies=proxies)
        _patch_baidu_requests(self._session)

        try:
            from deep_translator import BaiduTranslator
            self._impl_cls = BaiduTranslator
        except Exception:
            self._impl_cls = None

        self._instances: Dict[str, Any] = {}
        self._cache: Dict[str, Dict[str, str]] = {}

    @property
    def tier(self) -> str:
        return self._tier

    @tier.setter
    def tier(self, value: str) -> None:
        if value == self._tier:
            return
        logger.debug("Baidu tier changed: %s → %s; clearing cached instances", self._tier, value)
        self._tier = value
        # Update the engine key used for cache storage and usage tracking.
        self._engine_name = "baidu-premium" if value == "premium" else "baidu"
        self._instances.clear()
        # L1 in-memory translation cache is intentionally kept
    
    def _get_instance(self, target_lang: str):
        inst = self._instances.get(target_lang)
        if inst is None:
            if self._impl_cls is None:
                raise RuntimeError("deep_translator.BaiduTranslator is not available")
            src = "auto" if self._source_lang == "auto" else _baidu_code(self._source_lang)
            inst = self._impl_cls(
                source=src,
                target=_baidu_code(target_lang),
                appid=self._appid,
                appkey=self._appkey,
            )
            self._instances[target_lang] = inst
        return inst

    # --- language filtering ----------------------------------------------------

    def _should_translate(self, text: str) -> bool:
        if self._source_lang == "auto":
            return True
        from .lang_detect import is_source_language
        return is_source_language(text, self._source_lang, detector=self._detector)

    _SEGMENT_RE = re.compile(r'(\n|\r\n|\r|/)')

    def _translate_segments(self, text: str, target_lang: str) -> str:
        from .lang_detect import is_source_language
        parts = self._SEGMENT_RE.split(text)
        changed = False
        result_parts: list[str] = []
        for i, part in enumerate(parts):
            if i % 2 == 1:
                result_parts.append(part)
                continue
            stripped = part.strip()
            if not stripped:
                result_parts.append(part)
                continue
            if is_source_language(stripped, self._source_lang, detector=self._detector, strict=True):
                translated = self._translate_single(stripped, target_lang)
                leading = part[:len(part) - len(part.lstrip())]
                trailing = part[len(part.rstrip()):]
                result_parts.append(leading + translated + trailing)
                changed = True
            else:
                result_parts.append(part)
        return "".join(result_parts) if changed else text

    def _translate_single(self, text: str, target_lang: str) -> str:
        tgt_cache = self._cache.setdefault(target_lang, {})
        if text in tgt_cache:
            return tgt_cache[text]
        # L2 SQLite cache lookup
        try:
            from .cache import get_cache
            cached = get_cache().get(self._engine_name, text, target_lang)
            if cached is not None:
                tgt_cache[text] = cached
                return cached
        except Exception:
            pass
        try:
            import time
            if self._tier == "standard":
                time.sleep(_BAIDU_QPS_DELAY)
            translator = self._get_instance(target_lang)
            result = translator.translate(text)
            result = result if result is not None else text
            result = post_process(result)
            tgt_cache[text] = result
            try:
                from .cache import get_cache
                from .usage import get_tracker
                get_cache().put(self._engine_name, text, target_lang, result)
                get_tracker().record(self._engine_name, len(text))
            except Exception:
                pass
            return result
        except Exception:
            logger.exception("Baidu segment translation failed for target '%s'", target_lang)
            return text

    # ----- single-item convenience ----- #

    def translate_text(self, text: str, target_lang: str) -> str:
        if not self._impl_cls or not text or not text.strip():
            return text
        if self._source_lang != "auto":
            if self._SEGMENT_RE.search(text):
                return self._translate_segments(text, target_lang)
            if not self._should_translate(text):
                return text
        tgt_cache = self._cache.setdefault(target_lang, {})
        if text in tgt_cache:
            return tgt_cache[text]
        # L2 SQLite cache lookup
        try:
            from .cache import get_cache
            cached = get_cache().get(self._engine_name, text, target_lang)
            if cached is not None:
                tgt_cache[text] = cached
                return cached
        except Exception:
            pass
        try:
            import time
            if self._tier == "standard":
                time.sleep(_BAIDU_QPS_DELAY)
            translator = self._get_instance(target_lang)
            result = translator.translate(text)
            result = result if result is not None else text
            result = post_process(result)
            tgt_cache[text] = result
            try:
                from .cache import get_cache
                from .usage import get_tracker
                get_cache().put(self._engine_name, text, target_lang, result)
                get_tracker().record(self._engine_name, len(text))
            except Exception:
                pass
            return result
        except Exception:
            logger.exception("Baidu translation failed for target '%s'", target_lang)
            raise

    # ----- batch (the fast path) ----- #

    def translate_batch(
        self,
        texts: list[str],
        target_lang: str,
        *,
        cancel_event: Optional[threading.Event] = None,
    ) -> list[str]:
        if not self._impl_cls:
            return list(texts)

        results: list[str] = list(texts)
        tgt_cache = self._cache.setdefault(target_lang, {})

        _lingua_batch: dict[int, bool] = {}
        if self._source_lang != "auto" and self._detector == "lingua":
            from .lang_detect import is_source_language_batch
            candidates = [
                (i, t) for i, t in enumerate(texts)
                if t and t.strip() and not self._SEGMENT_RE.search(t)
            ]
            if candidates:
                batch_flags = is_source_language_batch(
                    [t for _, t in candidates],
                    self._source_lang,
                    detector="lingua",
                )
                _lingua_batch = {i: flag for (i, _), flag in zip(candidates, batch_flags)}

        to_translate: list[tuple[int, str]] = []
        for i, t in enumerate(texts):
            if not t or not t.strip():
                continue
            if self._source_lang != "auto" and self._SEGMENT_RE.search(t):
                results[i] = self._translate_segments(t, target_lang)
                continue
            should = _lingua_batch[i] if i in _lingua_batch else self._should_translate(t)
            if not should:
                continue
            if t in tgt_cache:
                results[i] = tgt_cache[t]
            else:
                to_translate.append((i, t))

        if not to_translate:
            return results

        # L2 SQLite cache lookup for L1 misses
        try:
            from .cache import get_cache
            l2_hits = get_cache().get_batch(
                self._engine_name, [t for _, t in to_translate], target_lang
            )
            if l2_hits:
                still: list[tuple[int, str]] = []
                for i, t in to_translate:
                    if t in l2_hits:
                        tgt_cache[t] = l2_hits[t]
                        results[i] = l2_hits[t]
                    else:
                        still.append((i, t))
                to_translate = still
        except Exception:
            pass

        if not to_translate:
            return results

        # Deduplicate
        unique_texts: dict[str, list[int]] = {}
        for idx, t in to_translate:
            unique_texts.setdefault(t, []).append(idx)

        dedup_items: list[tuple[str, list[int]]] = list(unique_texts.items())
        translator = self._get_instance(target_lang)
        chunks = self._make_chunks(dedup_items)

        for chunk in chunks:
            if cancel_event is not None and cancel_event.is_set():
                raise CancelledError("Translation cancelled")
            # Baidu's translate_batch is just a sequential loop internally,
            # but we respect the QPS delay between items.
            for orig_text, indices in chunk:
                if cancel_event is not None and cancel_event.is_set():
                    raise CancelledError("Translation cancelled")
                try:
                    import time
                    if self._tier == "standard":
                        time.sleep(_BAIDU_QPS_DELAY)

                    def _do(text=orig_text):
                        return translator.translate(text)

                    r = _run_cancellable(_do, cancel_event)
                    r = r if r is not None else orig_text
                    r = post_process(r)
                    tgt_cache[orig_text] = r
                    for idx in indices:
                        results[idx] = r
                    try:
                        from .cache import get_cache
                        from .usage import get_tracker
                        get_cache().put(self._engine_name, orig_text, target_lang, r)
                        get_tracker().record(self._engine_name, len(orig_text))
                    except Exception:
                        pass
                except CancelledError:
                    raise
                except Exception:
                    logger.exception("Baidu per-item translation failed")

        return results

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
