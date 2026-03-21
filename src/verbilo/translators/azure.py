# Microsoft Azure Cognitive Services Translator (v3) backend
#
# REST endpoint: https://api.cognitive.microsofttranslator.com/translate?api-version=3.0
# Auth headers : Ocp-Apim-Subscription-Key + Ocp-Apim-Subscription-Region
#
# Free tier    : 2 000 000 characters / month
# Batch limits : up to 100 elements or 50 000 chars per request

from __future__ import annotations

import re
import threading
from typing import Optional, Dict

from .base import Translator, has_inline_tags, unicode_tags_to_html, html_tags_to_unicode
from .http_session import make_session, is_transient_error
from ..utils import CancelledError
import logging

logger = logging.getLogger(__name__)

# ── API constants ─────────────────────────────────────────────────────────────
_API_URL = "https://api.cognitive.microsofttranslator.com/translate"
_API_VERSION = "3.0"

# Keep well under Azure's 100-element / 50 000-char limits per request.
_BATCH_SIZE = 100
_BATCH_MAX_CHARS = 40_000

# Monthly free-tier character limit (used by UsageTracker).
AZURE_MONTHLY_LIMIT = 2_000_000

# ── Supported languages ───────────────────────────────────────────────────────
# ISO 639-1 (and a handful of regional) codes that Azure Translator supports.
# Source: https://learn.microsoft.com/azure/cognitive-services/translator/language-support
AZURE_LANG_CODES: frozenset[str] = frozenset({
    "af", "am", "ar", "as", "az", "ba", "bg", "bn", "bo", "bs",
    "ca", "cs", "cy", "da", "de", "el", "en", "es", "et", "eu",
    "fa", "fi", "fj", "fr", "ga", "gl", "gu", "he", "hi", "hr",
    "ht", "hu", "hy", "id", "ig", "is", "it", "ja", "ka", "kk",
    "km", "ko", "ku", "ky", "lo", "lt", "lv", "mg", "mi", "mk",
    "ml", "mn", "ms", "mt", "my", "ne", "nl", "no", "or", "pa",
    "pl", "pt", "ro", "ru", "sk", "sl", "sm", "sn", "so", "sq",
    "sr", "st", "sv", "sw", "ta", "te", "th", "ti", "tk", "tl",
    "tn", "tr", "tt", "ug", "uk", "ur", "uz", "vi", "xh", "yo",
    "zh", "zu",
    # Regional codes present in the language dropdown
    "zh-CN", "zh-TW",
})

# Codes where Azure expects a different string from plain ISO 639-1.
_AZURE_TARGET_MAP: dict[str, str] = {
    "zh":    "zh-Hans",
    "zh-cn": "zh-Hans",
    "zh-tw": "zh-Hant",
    "sr":    "sr-Cyrl",   # default to Cyrillic Serbian
    "no":    "nb",        # Norwegian → Bokmål
}


def _azure_target_lang(iso_code: str) -> str:
    # Map an ISO 639-1 (or regional) code to the Azure API language code
    return _AZURE_TARGET_MAP.get(iso_code.lower(), iso_code)


# ── Translator wrapper ────────────────────────────────────────────────────────

class AzureTranslatorWrapper:
    """Wraps the Azure Cognitive Services Translator v3 REST API.

    Parameters
    ----------
    api_key:     Azure subscription key.
    region:      Azure region slug, e.g. ``"eastus"``.
    source_lang: ISO 639-1 code, or ``"auto"`` to translate everything.
    detector:    ``"fasttext"`` or ``"lingua"`` for source-language detection.
    proxies:     Optional proxy dict forwarded to :func:`make_session`.
    """

    _SEGMENT_RE = re.compile(r'(\n|\r\n|\r|/)')
    _engine_name = "azure"

    def __init__(
        self,
        api_key: str,
        region: str,
        source_lang: str = "auto",
        detector: str = "fasttext",
        proxies: Optional[dict] = None,
    ):
        self._api_key = api_key
        self._region = region
        self._source_lang = source_lang
        self._detector = detector
        self._session = make_session(proxies=proxies)
        # L1 in-memory cache: {target_lang: {source_text: translated_text}}
        self._cache: Dict[str, Dict[str, str]] = {}

    # ── Language filtering (same pattern as GoogleCloudTranslatorWrapper) ──

    def _should_translate(self, text: str) -> bool:
        if self._source_lang == "auto":
            return True
        from .lang_detect import is_source_language
        return is_source_language(text, self._source_lang, detector=self._detector)

    def _translate_segments(self, text: str, target_lang: str) -> str:
        # Translate only the segments that match source_lang (splits on / and newlines)
        from .lang_detect import is_source_language
        parts = self._SEGMENT_RE.split(text)
        changed = False
        result_parts: list[str] = []
        for i, part in enumerate(parts):
            if i % 2 == 1:          # separator — keep as-is
                result_parts.append(part)
                continue
            stripped = part.strip()
            if not stripped:
                result_parts.append(part)
                continue
            if is_source_language(stripped, self._source_lang, detector=self._detector, strict=True):
                translated = self._azure_translate_single(stripped, target_lang)
                leading  = part[: len(part) - len(part.lstrip())]
                trailing = part[len(part.rstrip()):]
                result_parts.append(leading + translated + trailing)
                changed = True
            else:
                result_parts.append(part)
        return "".join(result_parts) if changed else text

    # ── Core API call ─────────────────────────────────────────────────────────

    def _azure_translate_texts(self, texts: list[str], target_lang: str) -> list[str]:
        # POST to Azure Translator v3 and return one translated string per input.
        api_target = _azure_target_lang(target_lang)
        params: dict = {"api-version": _API_VERSION, "to": api_target}
        if self._source_lang != "auto":
            params["from"] = self._source_lang

        # Enable HTML mode when texts contain inline formatting tags so that
        # Azure preserves the <span class="rN"> elements we inject.
        use_html = any(has_inline_tags(t) for t in texts)
        if use_html:
            params["textType"] = "html"
            body = [{"Text": unicode_tags_to_html(t)} for t in texts]
        else:
            body = [{"Text": t} for t in texts]

        headers = {
            "Ocp-Apim-Subscription-Key": self._api_key,
            "Ocp-Apim-Subscription-Region": self._region,
            "Content-Type": "application/json",
        }
        resp = self._session.post(_API_URL, params=params, json=body, headers=headers)
        resp.raise_for_status()
        data = resp.json()
        results: list[str] = []
        for item, orig in zip(data, texts):
            translations = item.get("translations", [])
            results.append(translations[0].get("text", orig) if translations else orig)

        if use_html:
            results = [html_tags_to_unicode(r) for r in results]

        return results

    def _azure_translate_single(self, text: str, target_lang: str) -> str:
        from .google import post_process
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
            results = self._azure_translate_texts([text], target_lang)
            result = post_process(results[0]) if results else text
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
            logger.exception("Azure single translation failed for target '%s'", target_lang)
            return text

    # ── Public interface (Translator protocol) ────────────────────────────────

    def translate_text(self, text: str, target_lang: str) -> str:
        if not text or not text.strip():
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
        from .google import post_process
        try:
            results = self._azure_translate_texts([text], target_lang)
            result = post_process(results[0]) if results else text
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
            logger.exception("Azure translate_text failed for target '%s'", target_lang)
            raise

    def translate_batch(
        self,
        texts: list[str],
        target_lang: str,
        *,
        cancel_event: Optional[threading.Event] = None,
    ) -> list[str]:
        from .google import post_process, _run_cancellable

        results: list[str] = list(texts)
        tgt_cache = self._cache.setdefault(target_lang, {})

        # Pre-compute Lingua batch detection (uses all CPU cores).
        _lingua_batch: dict[int, bool] = {}
        if self._source_lang != "auto" and self._detector == "lingua":
            from .lang_detect import is_source_language_batch
            candidates = [
                (i, t) for i, t in enumerate(texts)
                if t and t.strip() and not self._SEGMENT_RE.search(t)
            ]
            if candidates:
                batch_flags = is_source_language_batch(
                    [t for _, t in candidates], self._source_lang, detector="lingua",
                )
                _lingua_batch = {i: flag for (i, _), flag in zip(candidates, batch_flags)}

        # Step 1: Separate translatable vs. pass-through; resolve L1 cache hits.
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

        # Step 2: Check L2 (SQLite) for remaining items.
        try:
            from .cache import get_cache
            l2_hits = get_cache().get_batch(
                self._engine_name, [t for _, t in to_translate], target_lang
            )
            if l2_hits:
                still_to_translate: list[tuple[int, str]] = []
                for i, t in to_translate:
                    if t in l2_hits:
                        tgt_cache[t] = l2_hits[t]
                        results[i] = l2_hits[t]
                    else:
                        still_to_translate.append((i, t))
                to_translate = still_to_translate
        except Exception:
            pass

        if not to_translate:
            return results

        # Step 3: Deduplicate and batch-translate via API.
        unique_texts: dict[str, list[int]] = {}
        for idx, t in to_translate:
            unique_texts.setdefault(t, []).append(idx)
        dedup_items: list[tuple[str, list[int]]] = list(unique_texts.items())
        chunks = self._make_chunks(dedup_items)

        for chunk in chunks:
            if cancel_event is not None and cancel_event.is_set():
                raise CancelledError("Translation cancelled")
            chunk_texts = [t for t, _ in chunk]
            try:
                translated = _run_cancellable(
                    lambda texts=chunk_texts: self._azure_translate_texts(texts, target_lang),
                    cancel_event,
                )
                if translated is None:
                    translated = chunk_texts
                l2_pairs: list[tuple[str, str]] = []
                chars_sent = 0
                for (orig_text, indices), tr_text in zip(chunk, translated):
                    tr_text = post_process(tr_text) if tr_text else orig_text
                    tgt_cache[orig_text] = tr_text
                    for idx in indices:
                        results[idx] = tr_text
                    l2_pairs.append((orig_text, tr_text))
                    chars_sent += len(orig_text)
                # Persist to L2 cache and record usage
                try:
                    from .cache import get_cache
                    from .usage import get_tracker
                    get_cache().put_batch(self._engine_name, l2_pairs, target_lang)
                    get_tracker().record(self._engine_name, chars_sent)
                except Exception:
                    pass
            except CancelledError:
                raise
            except Exception as exc:
                if not is_transient_error(exc):
                    raise
                logger.exception("Azure batch failed; falling back to sub-batches")
                self._subbatch_fallback(
                    chunk, target_lang, results, tgt_cache, cancel_event,
                )
        return results

    def _subbatch_fallback(
        self,
        chunk: list[tuple[str, list[int]]],
        target_lang: str,
        results: list[str],
        tgt_cache: dict[str, str],
        cancel_event=None,
    ) -> None:
        # Binary-halving fallback — split failing chunk until per-item
        from .google import post_process, _run_cancellable

        if len(chunk) <= 2:
            for orig_text, indices in chunk:
                if cancel_event is not None and cancel_event.is_set():
                    raise CancelledError("Translation cancelled")
                try:
                    r = self._azure_translate_single(orig_text, target_lang)
                    for idx in indices:
                        results[idx] = r
                except CancelledError:
                    raise
                except Exception:
                    logger.exception("Azure per-item fallback also failed")
            return
        mid = len(chunk) // 2
        for half in (chunk[:mid], chunk[mid:]):
            if cancel_event is not None and cancel_event.is_set():
                raise CancelledError("Translation cancelled")
            half_texts = [t for t, _ in half]
            try:
                translated = _run_cancellable(
                    lambda texts=half_texts: self._azure_translate_texts(texts, target_lang),
                    cancel_event,
                )
                if translated is None:
                    translated = half_texts
                for (orig_text, i_list), tr_text in zip(half, translated):
                    tr_text = post_process(tr_text) if tr_text else orig_text
                    tgt_cache[orig_text] = tr_text
                    for idx in i_list:
                        results[idx] = tr_text
            except CancelledError:
                raise
            except Exception:
                logger.exception("Azure sub-batch failed; recursing")
                self._subbatch_fallback(half, target_lang, results, tgt_cache, cancel_event)

    @staticmethod
    def _make_chunks(
        items: list[tuple[str, list[int]]],
    ) -> list[list[tuple[str, list[int]]]]:
        chunks: list[list[tuple[str, list[int]]]] = []
        current: list[tuple[str, list[int]]] = []
        current_chars = 0
        for item in items:
            text_len = len(item[0])
            if current and (
                len(current) >= _BATCH_SIZE
                or current_chars + text_len > _BATCH_MAX_CHARS
            ):
                chunks.append(current)
                current = []
                current_chars = 0
            current.append(item)
            current_chars += text_len
        if current:
            chunks.append(current)
        return chunks
