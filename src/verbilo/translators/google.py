# Google Translate backend via deep_translator — batching, caching, fallback

from __future__ import annotations

import re
import types
import threading
from typing import Optional, Dict, Any
from .base import Translator
from .http_session import make_session, resolve_proxies
from ..utils import CancelledError
import logging

logger = logging.getLogger(__name__)


def _patch_google_requests(session):
    try:
        import deep_translator.google as _mod
        shim = types.ModuleType("requests_shim")
        shim.get = session.get    # type: ignore[attr-defined]
        shim.post = session.post  # type: ignore[attr-defined]
        _mod.requests = shim      # type: ignore[attr-defined]
        logger.debug("Patched deep_translator.google.requests → resilient session")
    except Exception:
        logger.debug("Could not patch deep_translator.google.requests", exc_info=True)

# Maximum texts per Google Translate batch request.
_BATCH_SIZE = 50

# Maximum total characters per batch – just under Google's ~5 000 char limit.
_BATCH_MAX_CHARS = 4900

# Valid detector names for the language-detection subsystem.
VALID_DETECTORS = ("fasttext", "lingua")


# Lightweight post-processing: fix common Google Translate artefacts
def post_process(text: str) -> str:
    # fix spacing around punctuation and collapse multiple spaces
    if not text:
        return text
    # Preserve leading/trailing whitespace (meaningful for DOCX runs)
    leading = text[:len(text) - len(text.lstrip())]
    trailing = text[len(text.rstrip()):]
    inner = text.strip()
    if not inner:
        return text
    inner = re.sub(r'\s+([,.:;!?\)}\]])', r'\1', inner)
    inner = re.sub(r'([,;])(?=[^\s\d])', r'\1 ', inner)
    inner = re.sub(r'(:)(?=[^\s\d/\\])', r'\1 ', inner)
    inner = re.sub(r'([.!?])(?=[A-Za-z\u00C0-\u024F\u0400-\u04FF\u4e00-\u9fff])', r'\1 ', inner)
    inner = re.sub(r'  +', ' ', inner)
    return leading + inner + trailing


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

    _engine_name = "google"

    def __init__(self, source_lang: str = "auto", detector: str = "fasttext",
                 proxies: Optional[dict] = None):
        self._source_lang = source_lang
        self._detector = detector
        self._proxies = proxies

        # Set up a resilient session and inject it into deep_translator.google
        self._session = make_session(proxies=proxies)
        _patch_google_requests(self._session)

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
            inst = self._impl_cls(
                source="auto", target=target_lang,
                proxies=resolve_proxies(self._proxies),
            )
            self._instances[target_lang] = inst
        return inst

    def _should_translate(self, text: str) -> bool:
        # Return True if *text* should be sent to the translator.
        if self._source_lang == "auto":
            return True
        from .lang_detect import is_source_language
        return is_source_language(text, self._source_lang, detector=self._detector)

    # Pattern for splitting mixed-language cells exactly.
    _SEGMENT_RE = re.compile(r'(\n|\r\n|\r|/)')

    def _translate_segments(self, text: str, target_lang: str) -> str:
        # Split *text* on ``/`` and newlines, translate only the segments
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
        # Translate a single piece of text (no segment splitting)
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
            translator = self._get_instance(target_lang)
            result = translator.translate(text)
            result = result if result is not None else text
            result = post_process(result)
            tgt_cache[text] = result
            try:
                from .cache import get_cache
                get_cache().put(self._engine_name, text, target_lang, result)
            except Exception:
                pass
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
            translator = self._get_instance(target_lang)
            result = translator.translate(text)
            result = result if result is not None else text
            result = post_process(result)
            tgt_cache[text] = result
            try:
                from .cache import get_cache
                get_cache().put(self._engine_name, text, target_lang, result)
            except Exception:
                pass
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

        # Pre-compute batch language detection for Lingua (uses all CPU cores).
        # For fasttext or source_lang=="auto", per-item detection is used instead.
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

        # Separate translatable vs pass-through, resolving cache hits inline.
        to_translate: list[tuple[int, str]] = []
        for i, t in enumerate(texts):
            if not t or not t.strip():
                continue
            # Mixed-language cell: handle per-segment (not batchable)
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
                l2_pairs: list[tuple[str, str]] = []
                for (orig_text, indices), tr_text in zip(chunk, translated):
                    tr_text = tr_text if tr_text is not None else orig_text
                    tr_text = post_process(tr_text)
                    tgt_cache[orig_text] = tr_text
                    for idx in indices:
                        results[idx] = tr_text
                    l2_pairs.append((orig_text, tr_text))
                try:
                    from .cache import get_cache
                    get_cache().put_batch(self._engine_name, l2_pairs, target_lang)
                except Exception:
                    pass
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


class GoogleCloudTranslatorWrapper:
    # Uses the official Google Cloud Translation API v2 (requires an API key).

    _API_URL = "https://translation.googleapis.com/language/translate/v2"
    _engine_name = "google-cloud"

    def __init__(
        self,
        api_key: str,
        source_lang: str = "auto",
        detector: str = "fasttext",
        proxies: Optional[dict] = None,
    ):
        self._api_key = api_key
        self._source_lang = source_lang
        self._detector = detector
        self._session = make_session(proxies=proxies)
        self._cache: Dict[str, Dict[str, str]] = {}

    # --- language filtering (same logic as DeepTranslatorWrapper) -----------

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
                translated = self._cloud_translate_single(stripped, target_lang)
                leading = part[:len(part) - len(part.lstrip())]
                trailing = part[len(part.rstrip()):]
                result_parts.append(leading + translated + trailing)
                changed = True
            else:
                result_parts.append(part)
        return "".join(result_parts) if changed else text

    # --- core API call ---------------------------------------------------

    def _cloud_translate_texts(self, texts: list[str], target_lang: str) -> list[str]:
        # Call the Cloud Translation v2 REST API for a list of texts
        params: dict = {
            "key": self._api_key,
            "target": target_lang,
            "format": "text",
        }
        if self._source_lang != "auto":
            params["source"] = self._source_lang
        # The v2 API accepts multiple 'q' params in one POST
        data = {"q": texts}
        data.update(params)
        resp = self._session.post(self._API_URL, json=data)
        resp.raise_for_status()
        body = resp.json()
        translations = body.get("data", {}).get("translations", [])
        return [
            t.get("translatedText", orig)
            for t, orig in zip(translations, texts)
        ]

    def _cloud_translate_single(self, text: str, target_lang: str) -> str:
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
            results = self._cloud_translate_texts([text], target_lang)
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
            logger.exception("Google Cloud single translation failed for target '%s'", target_lang)
            return text

    # --- public interface (Translator protocol) ----------------------------

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
        try:
            results = self._cloud_translate_texts([text], target_lang)
            result = post_process(results[0]) if results else text
            tgt_cache[text] = result
            return result
        except Exception:
            logger.exception("Google Cloud translate failed for target '%s'", target_lang)
            raise

    def translate_batch(
        self,
        texts: list[str],
        target_lang: str,
        *,
        cancel_event: Optional[threading.Event] = None,
    ) -> list[str]:
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

        # L2 SQLite cache lookup for items that missed L1
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

        # Cloud API can handle up to ~128 segments per request;  use the same
        # chunk helper but with a larger batch size.
        chunks = DeepTranslatorWrapper._make_chunks(dedup_items)

        for chunk in chunks:
            if cancel_event is not None and cancel_event.is_set():
                raise CancelledError("Translation cancelled")
            chunk_texts = [t for t, _ in chunk]
            try:
                translated = _run_cancellable(
                    lambda texts=chunk_texts: self._cloud_translate_texts(texts, target_lang),
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
                logger.exception("Cloud batch failed; falling back to per-item")
                for orig_text, indices in chunk:
                    if cancel_event is not None and cancel_event.is_set():
                        raise CancelledError("Translation cancelled")
                    try:
                        r = self._cloud_translate_single(orig_text, target_lang)
                        for idx in indices:
                            results[idx] = r
                    except Exception:
                        logger.exception("Cloud per-item fallback also failed")
        return results


class GoogleCloudV3TranslatorWrapper:
    # Uses the Google Cloud Translation API v3 (Advanced).

    _engine_name = "google-cloud-v3"

    def __init__(
        self,
        project_id: str,
        sa_json: str = "",
        source_lang: str = "auto",
        detector: str = "fasttext",
        proxies: Optional[dict] = None,   # accepted but no-op for gRPC
    ):
        self._project_id = project_id
        self._sa_json = sa_json
        self._source_lang = source_lang
        self._detector = detector
        self._client = None    # lazy-initialised in _get_client()
        self._cache: Dict[str, Dict[str, str]] = {}

    # ── Client ─────────────────────────────────────────────────────────────────────────────

    def _get_client(self):
        if self._client is not None:
            return self._client
        try:
            from google.cloud import translate_v3
        except ImportError:
            raise RuntimeError(
                "google-cloud-translate is not installed. "
                'Run: pip install "google-cloud-translate>=3.0.0"'
            )
        if self._sa_json:
            import json as _json
            import os as _os
            from google.oauth2 import service_account
            if _os.path.isfile(self._sa_json):
                with open(self._sa_json, encoding="utf-8") as fh:
                    info = _json.load(fh)
            else:
                info = _json.loads(self._sa_json)
            creds = service_account.Credentials.from_service_account_info(
                info,
                scopes=["https://www.googleapis.com/auth/cloud-translation"],
            )
            self._client = translate_v3.TranslationServiceClient(credentials=creds)
        else:
            self._client = translate_v3.TranslationServiceClient()
        return self._client

    # ── Language filtering ──────────────────────────────────────────────────────────

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
                translated = self._v3_translate_single(stripped, target_lang)
                leading = part[:len(part) - len(part.lstrip())]
                trailing = part[len(part.rstrip()):]
                result_parts.append(leading + translated + trailing)
                changed = True
            else:
                result_parts.append(part)
        return "".join(result_parts) if changed else text

    # ── Core API call ─────────────────────────────────────────────────────────────────────

    def _v3_translate_texts(self, texts: list[str], target_lang: str) -> list[str]:
        # Call the Cloud Translation v3 API and return one result per input
        client = self._get_client()
        parent = f"projects/{self._project_id}/locations/global"
        request: dict = {
            "parent": parent,
            "contents": texts,
            "target_language_code": target_lang,
            "mime_type": "text/plain",
        }
        if self._source_lang != "auto":
            request["source_language_code"] = self._source_lang
        response = client.translate_text(request=request)
        return [t.translated_text for t in response.translations]

    def _v3_translate_single(self, text: str, target_lang: str) -> str:
        tgt_cache = self._cache.setdefault(target_lang, {})
        if text in tgt_cache:
            return tgt_cache[text]
        try:
            from .cache import get_cache
            cached = get_cache().get(self._engine_name, text, target_lang)
            if cached is not None:
                tgt_cache[text] = cached
                return cached
        except Exception:
            pass
        try:
            results = self._v3_translate_texts([text], target_lang)
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
            logger.exception("Google Cloud v3 single translation failed for target '%s'", target_lang)
            return text

    # ── Public interface (Translator protocol) ───────────────────────────────────────

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
        try:
            from .cache import get_cache
            cached = get_cache().get(self._engine_name, text, target_lang)
            if cached is not None:
                tgt_cache[text] = cached
                return cached
        except Exception:
            pass
        try:
            results = self._v3_translate_texts([text], target_lang)
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
            logger.exception("Google Cloud v3 translate_text failed for target '%s'", target_lang)
            raise

    def translate_batch(
        self,
        texts: list[str],
        target_lang: str,
        *,
        cancel_event: Optional[threading.Event] = None,
    ) -> list[str]:
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
                    [t for _, t in candidates], self._source_lang, detector="lingua",
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

        unique_texts: dict[str, list[int]] = {}
        for idx, t in to_translate:
            unique_texts.setdefault(t, []).append(idx)
        dedup_items: list[tuple[str, list[int]]] = list(unique_texts.items())
        chunks = DeepTranslatorWrapper._make_chunks(dedup_items)

        for chunk in chunks:
            if cancel_event is not None and cancel_event.is_set():
                raise CancelledError("Translation cancelled")
            chunk_texts = [t for t, _ in chunk]
            try:
                translated = _run_cancellable(
                    lambda texts=chunk_texts: self._v3_translate_texts(texts, target_lang),
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
                logger.exception("Google Cloud v3 batch failed; falling back to per-item")
                for orig_text, indices in chunk:
                    if cancel_event is not None and cancel_event.is_set():
                        raise CancelledError("Translation cancelled")
                    try:
                        r = self._v3_translate_single(orig_text, target_lang)
                        for idx in indices:
                            results[idx] = r
                    except Exception:
                        logger.exception("Google Cloud v3 per-item fallback also failed")
        return results
