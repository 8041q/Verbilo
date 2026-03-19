# DeepL Free REST API backend
#
# Endpoint : https://api-free.deepl.com/v2/translate
# Auth     : Authorization: DeepL-Auth-Key <key>
#
# Free tier : 500 000 characters / month
# Batch     : up to 50 "text" items per request

from __future__ import annotations

import re
import threading
from typing import Optional, Dict

from .base import Translator
from .http_session import make_session
from ..utils import CancelledError
import logging

logger = logging.getLogger(__name__)

# ── API constants ─────────────────────────────────────────────────────────────
_FREE_API_URL = "https://api-free.deepl.com/v2/translate"
_PRO_API_URL  = "https://api.deepl.com/v2/translate"

_BATCH_SIZE      = 50
_BATCH_MAX_CHARS = 40_000

# Monthly free-tier character limit (used by UsageTracker).
DEEPL_MONTHLY_LIMIT = 500_000

# ── Language code mapping ─────────────────────────────────────────────────────
# DeepL uses uppercase codes with regional variants for targets.
# Source: https://developers.deepl.com/docs/resources/supported-languages
_ISO_TO_DEEPL_TARGET: dict[str, str] = {
    "ar":    "AR",
    "bg":    "BG",
    "cs":    "CS",
    "da":    "DA",
    "de":    "DE",
    "el":    "EL",
    "en":    "EN-US",   # default to American English
    "es":    "ES",
    "et":    "ET",
    "fi":    "FI",
    "fr":    "FR",
    "hu":    "HU",
    "id":    "ID",
    "it":    "IT",
    "ja":    "JA",
    "ko":    "KO",
    "lt":    "LT",
    "lv":    "LV",
    "nb":    "NB",
    "no":    "NB",      # Norwegian → Bokmål
    "nl":    "NL",
    "pl":    "PL",
    "pt":    "PT-BR",   # default to Brazilian Portuguese
    "pt-br": "PT-BR",
    "pt-pt": "PT-PT",
    "ro":    "RO",
    "ru":    "RU",
    "sk":    "SK",
    "sl":    "SL",
    "sv":    "SV",
    "tr":    "TR",
    "uk":    "UK",
    "zh":    "ZH",
    "zh-cn": "ZH",
    "zh-tw": "ZH",      # DeepL does not have a separate Traditional-Chinese target
}

# ISO 639-1 codes that DeepL supports (used for language-list filtering in the GUI).
DEEPL_LANG_CODES: frozenset[str] = frozenset({
    "ar", "bg", "cs", "da", "de", "el", "en", "es", "et", "fi",
    "fr", "hu", "id", "it", "ja", "ko", "lt", "lv", "nb", "no",
    "nl", "pl", "pt", "ro", "ru", "sk", "sl", "sv", "tr", "uk", "zh",
    # Regional codes present in the language dropdown
    "zh-CN", "zh-TW",
})


def _deepl_target_lang(iso_code: str) -> str:
    # Convert an ISO 639-1 (or regional) code to the DeepL target language code
    return _ISO_TO_DEEPL_TARGET.get(iso_code.lower(), iso_code.upper())


def _deepl_source_lang(iso_code: str) -> str:
    # Convert an ISO 639-1 code to the DeepL source language code (plain uppercase)
    code = iso_code.lower()
    if code in ("zh", "zh-cn", "zh-tw"):
        return "ZH"
    # Strip regional suffix: "pt-br" → "PT"
    return code.split("-")[0].upper()


# ── Translator wrapper ────────────────────────────────────────────────────────

class DeepLTranslatorWrapper:
    """Wraps the DeepL Free (or Pro) REST API.

    Parameters
    ----------
    api_key:     DeepL authentication key.
    source_lang: ISO 639-1 code, or ``"auto"`` to translate everything.
    detector:    ``"fasttext"`` or ``"lingua"`` for source-language detection.
    proxies:     Optional proxy dict forwarded to :func:`make_session`.
    pro:         Use the DeepL Pro endpoint instead of the free one.
    """

    _SEGMENT_RE = re.compile(r'(\n|\r\n|\r|/)')
    _engine_name = "deepl"

    def __init__(
        self,
        api_key: str,
        source_lang: str = "auto",
        detector: str = "fasttext",
        proxies: Optional[dict] = None,
        pro: bool = False,
    ):
        self._api_key = api_key
        self._source_lang = source_lang
        self._detector = detector
        # Auto-detect endpoint from the key suffix: Free API keys always end in ':fx'.
        # Key suffix takes precedence over the 'pro' parameter so a Pro key pasted into
        # the "DeepL Free" settings field still hits api.deepl.com and avoids a 403.
        if api_key.endswith(":fx"):
            self._api_url = _FREE_API_URL
        else:
            self._api_url = _PRO_API_URL
        self._session = make_session(proxies=proxies)
        # L1 in-memory cache: {target_lang: {source_text: translated_text}}
        self._cache: Dict[str, Dict[str, str]] = {}

    # ── Language filtering ────────────────────────────────────────────────────

    def _should_translate(self, text: str) -> bool:
        if self._source_lang == "auto":
            return True
        from .lang_detect import is_source_language
        return is_source_language(text, self._source_lang, detector=self._detector)

    def _translate_segments(self, text: str, target_lang: str) -> str:
        # Translate only segments that match source_lang (splits on / and newlines)
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
                translated = self._deepl_translate_single(stripped, target_lang)
                leading  = part[: len(part) - len(part.lstrip())]
                trailing = part[len(part.rstrip()):]
                result_parts.append(leading + translated + trailing)
                changed = True
            else:
                result_parts.append(part)
        return "".join(result_parts) if changed else text

    # ── Core API call ─────────────────────────────────────────────────────────

    def _deepl_translate_texts(self, texts: list[str], target_lang: str) -> list[str]:
        # Call the DeepL translate endpoint and return one result per input text
        payload: dict = {
            "text": texts,
            "target_lang": _deepl_target_lang(target_lang),
        }
        if self._source_lang != "auto":
            payload["source_lang"] = _deepl_source_lang(self._source_lang)
        headers = {
            "Authorization": f"DeepL-Auth-Key {self._api_key}",
            "Content-Type": "application/json",
        }
        resp = self._session.post(self._api_url, json=payload, headers=headers)
        resp.raise_for_status()
        data = resp.json()
        translations = data.get("translations", [])
        return [t.get("text", orig) for t, orig in zip(translations, texts)]

    def _deepl_translate_single(self, text: str, target_lang: str) -> str:
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
            results = self._deepl_translate_texts([text], target_lang)
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
            logger.exception("DeepL single translation failed for target '%s'", target_lang)
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
            results = self._deepl_translate_texts([text], target_lang)
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
            logger.exception("DeepL translate_text failed for target '%s'", target_lang)
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
                    lambda texts=chunk_texts: self._deepl_translate_texts(texts, target_lang),
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
            except Exception:
                logger.exception("DeepL batch failed; falling back to per-item")
                for orig_text, indices in chunk:
                    if cancel_event is not None and cancel_event.is_set():
                        raise CancelledError("Translation cancelled")
                    try:
                        r = self._deepl_translate_single(orig_text, target_lang)
                        for idx in indices:
                            results[idx] = r
                    except Exception:
                        logger.exception("DeepL per-item fallback also failed")
        return results

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
