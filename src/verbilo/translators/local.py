# Local offline translation backend — OPUS-MT models via CTranslate2 + SentencePiece
#
# Each model pair lives on disk under ``<model_dir>/<src>-<tgt>/`` and must
# contain the CTranslate2 converted model plus ``source.spm`` / ``target.spm``.
# A ``converted.ok`` sentinel file written by ``download_models.py`` is checked
# before loading; missing sentinel → FileNotFoundError with actionable message.

from __future__ import annotations

import logging
import os
import re
import threading
from collections import OrderedDict
from pathlib import Path
from typing import Dict, Optional

from ..utils import CancelledError

logger = logging.getLogger(__name__)

# Maximum number of CTranslate2 model instances kept resident in memory.
_MAX_MODELS = 3

# CTranslate2 batch‐token budget (roughly matches Google's 4 900‐char limit in
# spirit: keep individual chunks small so cancellation is responsive).
_BATCH_SIZE = 64

_SENTINEL = "converted.ok"


class OpusMTTranslator:
    """Offline OPUS-MT translator backed by CTranslate2 + SentencePiece.

    Satisfies the ``Translator`` protocol defined in ``base.py``.

    When **source_lang is "auto"**, the source language is detected **once per
    batch** by sampling the first non-empty text segment.  This means that
    mixed-language documents will be routed through a single model determined
    by that sample.  This is a deliberate trade-off: OPUS-MT requires a fixed
    source language to select the correct ``opus-mt-{src}-{tgt}`` model, and
    per-string detection would cause constant model swapping.

    Model lifecycle
    ---------------
    Models are loaded lazily on first use and cached in an LRU
    ``OrderedDict`` capped at 3 entries.  When a 4th model would be loaded,
    the least-recently-used model is unloaded (``ct2.Translator.unload_model``)
    to free memory.
    """

    _engine_name = "local"
    _SEGMENT_RE = re.compile(r'(\n|\r\n|\r|/)')

    def __init__(
        self,
        model_dir: str,
        source_lang: str = "auto",
        detector: str = "fasttext",
    ):
        self._model_dir = Path(model_dir)
        self._source_lang = source_lang
        self._detector = detector
        # LRU cache: key = "{src}-{tgt}", value = (ct2.Translator, sp_source, sp_target)
        self._models: OrderedDict[str, tuple] = OrderedDict()
        # L1 in-memory translation cache: {target_lang: {source_text: translated_text}}
        self._cache: Dict[str, Dict[str, str]] = {}

    # ------------------------------------------------------------------
    # Model management (LRU cap = 3)
    # ------------------------------------------------------------------

    def _load_model(self, src: str, tgt: str):
        """Return ``(ct2_translator, sp_source, sp_target)`` for the pair,
        loading from disk if not already cached."""
        import ctranslate2
        import sentencepiece as spm

        key = f"{src}-{tgt}"

        if key in self._models:
            self._models.move_to_end(key)
            return self._models[key]

        pair_dir = self._model_dir / key

        # Decision 3: check sentinel before attempting to load.
        if not (pair_dir / _SENTINEL).exists():
            raise FileNotFoundError(
                f"OPUS-MT model '{key}' is missing or incomplete at {pair_dir}. "
                f"Run: python scripts/download_models.py opus-mt {src} {tgt}"
            )

        # Evict LRU if at capacity.
        if len(self._models) >= _MAX_MODELS:
            evict_key, (old_translator, _, _) = self._models.popitem(last=False)
            try:
                old_translator.unload_model()
            except Exception:
                pass
            logger.debug("Evicted model %s from LRU cache", evict_key)

        translator = ctranslate2.Translator(str(pair_dir), device="cpu")

        sp_source = spm.SentencePieceProcessor()
        sp_source.Load(str(pair_dir / "source.spm"))

        sp_target = spm.SentencePieceProcessor()
        sp_target.Load(str(pair_dir / "target.spm"))

        entry = (translator, sp_source, sp_target)
        self._models[key] = entry
        logger.info("Loaded OPUS-MT model %s from %s", key, pair_dir)
        return entry

    # ------------------------------------------------------------------
    # Source language resolution
    # ------------------------------------------------------------------

    def _resolve_src(self, text: str) -> str:
        """Detect the source language of *text*, falling back to ``"en"``."""
        from .lang_detect import detect_language
        code, _conf = detect_language(text, detector=self._detector)
        return code if code != "und" else "en"

    def _resolve_batch_src(self, texts: list[str]) -> str:
        """Detect source language once from the first non-empty element."""
        if self._source_lang != "auto":
            return self._source_lang
        for t in texts:
            if t and t.strip():
                return self._resolve_src(t)
        return "en"

    # ------------------------------------------------------------------
    # Low-level translation helpers
    # ------------------------------------------------------------------

    def _tokenize(self, sp, text: str) -> list[str]:
        return sp.Encode(text, out_type=str) + ["</s>"]

    def _detokenize(self, sp, tokens: list[str]) -> str:
        return sp.Decode(tokens)

    def _translate_tokens(self, model, tokens_batch: list[list[str]]) -> list[list[str]]:
        results = model.translate_batch(tokens_batch, beam_size=2, max_decoding_length=512)
        return [r.hypotheses[0] for r in results]

    def _translate_single_raw(self, text: str, src: str, tgt: str) -> str:
        """Translate a single string through the CTranslate2 model."""
        translator, sp_source, sp_target = self._load_model(src, tgt)
        tokens = self._tokenize(sp_source, text)
        translated = self._translate_tokens(translator, [tokens])
        return self._detokenize(sp_target, translated[0])

    # ------------------------------------------------------------------
    # Source-language filtering (mirrors google.py patterns)
    # ------------------------------------------------------------------

    def _should_translate(self, text: str) -> bool:
        if self._source_lang == "auto":
            return True
        from .lang_detect import is_source_language
        return is_source_language(text, self._source_lang, detector=self._detector)

    def _translate_segments(self, text: str, target_lang: str, src: str) -> str:
        """Split on ``/`` and newlines, translate only source-language segments."""
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
                translated = self._cached_translate(stripped, src, target_lang)
                leading = part[: len(part) - len(part.lstrip())]
                trailing = part[len(part.rstrip()) :]
                result_parts.append(leading + translated + trailing)
                changed = True
            else:
                result_parts.append(part)
        return "".join(result_parts) if changed else text

    # ------------------------------------------------------------------
    # Caching wrapper (L1 dict + L2 SQLite)
    # ------------------------------------------------------------------

    def _cached_translate(self, text: str, src: str, tgt: str) -> str:
        tgt_cache = self._cache.setdefault(tgt, {})
        if text in tgt_cache:
            return tgt_cache[text]
        # L2 SQLite lookup
        try:
            from .cache import get_cache
            cached = get_cache().get(self._engine_name, text, tgt)
            if cached is not None:
                tgt_cache[text] = cached
                return cached
        except Exception:
            pass
        result = self._translate_single_raw(text, src, tgt)
        tgt_cache[text] = result
        # L2 write
        try:
            from .cache import get_cache
            get_cache().put(self._engine_name, text, tgt, result)
        except Exception:
            pass
        return result

    # ------------------------------------------------------------------
    # Public interface — Translator protocol
    # ------------------------------------------------------------------

    def translate_text(self, text: str, target_lang: str) -> str:
        if not text or not text.strip():
            return text

        src = self._resolve_batch_src([text])

        if self._source_lang != "auto":
            if self._SEGMENT_RE.search(text):
                return self._translate_segments(text, target_lang, src)
            if not self._should_translate(text):
                return text

        return self._cached_translate(text, src, target_lang)

    def translate_batch(
        self,
        texts: list[str],
        target_lang: str,
        *,
        cancel_event: Optional[threading.Event] = None,
    ) -> list[str]:
        """Translate a list of strings, returning one result per input.

        When ``source_lang`` is ``"auto"``, the source language is detected
        **once** from the first non-empty text and used for the entire batch.
        """
        if not texts:
            return []

        results: list[str] = list(texts)

        # Decision 1: resolve source language once for the whole batch.
        src = self._resolve_batch_src(texts)
        tgt_cache = self._cache.setdefault(target_lang, {})

        # Collect items that need translation.
        to_translate: list[tuple[int, str]] = []
        for i, t in enumerate(texts):
            if cancel_event is not None and cancel_event.is_set():
                raise CancelledError("Translation cancelled")
            if not t or not t.strip():
                continue
            # Mixed-language segment handling
            if self._source_lang != "auto" and self._SEGMENT_RE.search(t):
                results[i] = self._translate_segments(t, target_lang, src)
                continue
            if self._source_lang != "auto" and not self._should_translate(t):
                continue
            if t in tgt_cache:
                results[i] = tgt_cache[t]
                continue
            to_translate.append((i, t))

        if not to_translate:
            return results

        # L2 SQLite bulk lookup
        try:
            from .cache import get_cache
            l2_hits = get_cache().get_batch(
                self._engine_name, [t for _, t in to_translate], target_lang,
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

        # Load model once for the batch
        translator, sp_source, sp_target = self._load_model(src, target_lang)

        # Translate in chunks for responsive cancellation
        l2_pairs: list[tuple[str, str]] = []
        for chunk_start in range(0, len(dedup_items), _BATCH_SIZE):
            if cancel_event is not None and cancel_event.is_set():
                raise CancelledError("Translation cancelled")

            chunk = dedup_items[chunk_start : chunk_start + _BATCH_SIZE]
            chunk_texts = [t for t, _ in chunk]

            tokens_batch = [self._tokenize(sp_source, t) for t in chunk_texts]
            translated_tokens = self._translate_tokens(translator, tokens_batch)

            for (orig_text, indices), out_tokens in zip(chunk, translated_tokens):
                tr_text = self._detokenize(sp_target, out_tokens)
                tgt_cache[orig_text] = tr_text
                for idx in indices:
                    results[idx] = tr_text
                l2_pairs.append((orig_text, tr_text))

        # Bulk L2 cache write
        if l2_pairs:
            try:
                from .cache import get_cache
                get_cache().put_batch(self._engine_name, l2_pairs, target_lang)
            except Exception:
                pass

        return results
