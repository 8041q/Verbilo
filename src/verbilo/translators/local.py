# Local offline translation backend — OPUS-MT models via CTranslate2 + SentencePiece
#
# Each model pair lives on disk under ``<model_dir>/<src>-<tgt>/`` and must
# contain the CTranslate2 converted model plus ``source.spm`` / ``target.spm``.
# A ``converted.ok`` sentinel file written by ``download_models.py`` is checked
# before loading; missing sentinel → FileNotFoundError with actionable message.

from __future__ import annotations

import json
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
_REQUIRED_MODEL_FILES = ("model.bin", "source.spm", "target.spm")

# OPUS catalogue/folder codes are not always the same codes used by the GUI.
# Keep this normalization in the backend too, because the loader must be able
# to open a legacy folder such as ``eng-fra`` when the GUI asks for ``en-fr``.
_OPUS_CODE_MAP: dict[str, str] = {
    "eng": "en", "fra": "fr", "deu": "de", "spa": "es", "por": "pt",
    "ita": "it", "nld": "nl", "rus": "ru", "zho": "zh", "jpn": "ja",
    "jap": "ja", "kor": "ko", "ara": "ar", "pol": "pl", "tur": "tr",
    "swe": "sv", "dan": "da", "fin": "fi", "ukr": "uk", "ces": "cs",
    "ron": "ro", "hun": "hu", "nor": "no", "bul": "bg", "hrv": "hr",
    "ell": "el", "heb": "he", "hin": "hi", "tha": "th", "vie": "vi",
    "cat": "ca", "ind": "id", "msa": "ms", "slk": "sk", "slv": "sl",
    "est": "et", "lav": "lv", "lit": "lt", "srp": "sr",
}


def normalize_model_code(code: str) -> str:
    """Normalize an OPUS/folder language code to the GUI's canonical code."""
    normalized = (code or "").strip().lower().replace("_", "-")
    return _OPUS_CODE_MAP.get(normalized, normalized)


def _validated_pair_dir(path: Path) -> bool:
    """True only when a model is both marked ready and actually loadable."""
    if not (path / _SENTINEL).is_file():
        return False
    return all(
        (path / filename).is_file() and (path / filename).stat().st_size > 0
        for filename in _REQUIRED_MODEL_FILES
    )


def _pair_from_ready_dir(path: Path) -> tuple[str, str] | None:
    """Read the declared pair from the sentinel, with legacy-name fallback.

    New downloaders store JSON metadata in ``converted.ok``. Older releases
    wrote only ``ok`` and therefore have to infer the pair from the folder name.
    """
    sentinel = path / _SENTINEL
    try:
        payload = json.loads(sentinel.read_text(encoding="utf-8"))
        if isinstance(payload, dict):
            src = str(payload.get("source") or "").strip()
            tgt = str(payload.get("target") or "").strip()
            if src and tgt:
                return src, tgt
            pair = str(payload.get("pair") or "").strip()
            if pair and "-" in pair:
                src, tgt = pair.split("-", 1)
                if src and tgt:
                    return src, tgt
    except Exception:
        # Legacy sentinel contains plain text (normally ``ok``).
        pass

    name = path.name
    # Legacy downloader variants could accidentally retain a model-family
    # prefix. Handle the known forms without guessing at arbitrary BCP-47 tags.
    if "_" in name:
        name = name.split("_", 1)[-1]
    for prefix in ("tc-big-", "opus-mt-"):
        if name.startswith(prefix):
            name = name[len(prefix):]
            break
    parts = name.split("-", 1)
    if len(parts) != 2 or not parts[0] or not parts[1]:
        return None
    return parts[0], parts[1]


def list_downloaded_pairs(model_dir: str) -> list[tuple[str, str]]:
    """Return every *loadable* ``(src, tgt)`` model pair on disk.

    This intentionally uses the same readiness rules as ``_load_model`` so the
    GUI cannot advertise a model that the translator would reject.
    """
    root = Path(model_dir)
    if not root.is_dir():
        return []
    pairs: set[tuple[str, str]] = set()
    for child in root.iterdir():
        if not child.is_dir() or not _validated_pair_dir(child):
            continue
        pair = _pair_from_ready_dir(child)
        if pair is not None:
            pairs.add(pair)
    return sorted(pairs)


def find_downloaded_pair_dir(model_dir: str | Path, src: str, tgt: str) -> Path | None:
    """Resolve a requested GUI pair to its actual ready directory on disk.

    Exact canonical folders win. Legacy OPUS-code folders (for example
    ``eng-fra``) remain usable through normalized-code matching.
    """
    root = Path(model_dir)
    exact = root / f"{src}-{tgt}"
    if _validated_pair_dir(exact):
        return exact
    if not root.is_dir():
        return None

    want = (normalize_model_code(src), normalize_model_code(tgt))
    for child in sorted(root.iterdir(), key=lambda item: item.name.lower()):
        if not child.is_dir() or not _validated_pair_dir(child):
            continue
        pair = _pair_from_ready_dir(child)
        if pair is None:
            continue
        have = (normalize_model_code(pair[0]), normalize_model_code(pair[1]))
        if have == want:
            return child
    return None


def _same_language_code(source: str, target: str) -> bool:
    """Return True when translating would be a same-language round trip.

    Regional variants are treated as the same language except Chinese, where
    Simplified/Traditional conversion can be intentional.
    """
    src = (source or "").strip().lower().replace("_", "-")
    tgt = (target or "").strip().lower().replace("_", "-")
    if not src or not tgt:
        return False
    if src == tgt:
        return True
    src_base = src.split("-", 1)[0]
    tgt_base = tgt.split("-", 1)[0]
    return src_base == tgt_base and src_base != "zh"


class OpusMTTranslator:
    # Offline OPUS-MT translator backed by CTranslate2 + SentencePiece.

    _engine_name = "local"
    _CACHE_VERSION = "v2"
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

     
    # Model management (LRU cap = 3)
     
    def _load_model(self, src: str, tgt: str):
        # Return ``(ct2_translator, sp_source, sp_target)`` for the pair, loading from disk if not already cached
        import ctranslate2
        import sentencepiece as spm

        key = f"{src}-{tgt}"

        if key in self._models:
            self._models.move_to_end(key)
            return self._models[key]

        pair_dir = find_downloaded_pair_dir(self._model_dir, src, tgt)

        # Use exactly the same readiness rules as model discovery.
        if pair_dir is None:
            expected = self._model_dir / key
            raise FileNotFoundError(
                f"OPUS-MT model '{key}' is missing or incomplete under {self._model_dir}. "
                f"Expected a ready model such as {expected}. "
                f"Run: python scripts/download_models.py opus-mt {src}-{tgt}"
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

     
    # Source language resolution
     
    def _resolve_src(self, text: str) -> str:
        # Detect the source language of *text*, falling back to "en"
        from .lang_detect import detect_language
        code, _conf = detect_language(text, detector=self._detector)
        return code if code != "und" else "en"

    def _resolve_batch_src(self, texts: list[str]) -> str:
        # Detect source language once from the first non-empty element
        if self._source_lang != "auto":
            return self._source_lang
        for t in texts:
            if t and t.strip():
                return self._resolve_src(t)
        return "en"

     
    # Low-level translation helpers
     
    # Per-language-pair generation parameters for CTranslate2.  The zh-en pair
    # gets a shorter-favoring length_penalty because Chinese topic-prominent
    # constructions expand significantly into English subject-prominent prose.
    _PAIR_PARAMS: dict[str, dict] = {
        "zh-en": {
            "beam_size": 2,
            "max_decoding_length": 256,
            "length_penalty": 0.8,
            "repetition_penalty": 1.2,
            "no_repeat_ngram_size": 3,
        },
    }
    _DEFAULT_PARAMS: dict = {
        "beam_size": 2,
        "max_decoding_length": 512,
        "length_penalty": 1.0,
        "repetition_penalty": 1.0,
        "no_repeat_ngram_size": 0,
    }

    def _tokenize(self, sp, text: str) -> list[str]:
        return sp.Encode(text, out_type=str) + ["</s>"]

    def _detokenize(self, sp, tokens: list[str]) -> str:
        return sp.Decode(tokens)

    def _translate_tokens(
        self, model, tokens_batch: list[list[str]], src: str = "", tgt: str = "",
    ) -> list[list[str]]:
        pair_key = f"{src}-{tgt}" if src and tgt else ""
        params = self._PAIR_PARAMS.get(pair_key, self._DEFAULT_PARAMS)
        results = model.translate_batch(
            tokens_batch,
            beam_size=params["beam_size"],
            max_decoding_length=params["max_decoding_length"],
            length_penalty=params["length_penalty"],
            repetition_penalty=params["repetition_penalty"],
            no_repeat_ngram_size=params["no_repeat_ngram_size"],
        )
        return [r.hypotheses[0] for r in results]

    def _translate_single_raw(self, text: str, src: str, tgt: str) -> str:
        # Translate a single string through the CTranslate2 model
        translator, sp_source, sp_target = self._load_model(src, tgt)
        tokens = self._tokenize(sp_source, text)
        translated = self._translate_tokens(translator, [tokens], src=src, tgt=tgt)
        return self._detokenize(sp_target, translated[0])

     
    # Source-language filtering (mirrors google.py patterns)
     
    def _should_translate(self, text: str) -> bool:
        if self._source_lang == "auto":
            return True
        from .lang_detect import is_source_language
        return is_source_language(text, self._source_lang, detector=self._detector)

    def _translate_segments(self, text: str, target_lang: str, src: str) -> str:
        # Split on `/` and newlines, translate only source-language segments
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

     
    # Caching wrapper (L1 dict + L2 SQLite)
     
    def _cached_translate(self, text: str, src: str, tgt: str) -> str:
        cache_scope = f"{src.lower()}:{tgt.lower()}"
        tgt_cache = self._cache.setdefault(cache_scope, {})
        if text in tgt_cache:
            return tgt_cache[text]
        # L2 SQLite lookup
        try:
            from .cache import get_cache
            cached = get_cache().get(self._cache_engine(src, tgt), text, tgt)
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
            get_cache().put(self._cache_engine(src, tgt), text, tgt, result)
        except Exception:
            pass
        return result

    def _cache_engine(self, src: str, tgt: str) -> str:
        # Keep model-pair results isolated from one another and old cache data
        return f"{self._engine_name}:{src.lower()}-{tgt.lower()}:{self._CACHE_VERSION}"

    def _translate_auto_batch(
        self,
        texts: list[str],
        target_lang: str,
        cancel_event: Optional[threading.Event],
    ) -> list[str]:
        # Detect every unit and translate homogeneous language buckets.

        results = list(texts)
        buckets: dict[str, dict[str, list[int]]] = {}
        for index, text in enumerate(texts):
            if cancel_event is not None and cancel_event.is_set():
                raise CancelledError("Translation cancelled")
            if not text or not text.strip():
                continue
            src = self._resolve_src(text)
            # Bilingual spreadsheets/documents often contain cells that are
            # already in the target language.  Do not try to load a non-existent
            # en-en / fr-fr / etc. OPUS model for those cells.
            if _same_language_code(src, target_lang):
                continue
            buckets.setdefault(src, {}).setdefault(text, []).append(index)

        for src, unique_texts in buckets.items():
            cache_scope = f"{src.lower()}:{target_lang.lower()}"
            l1 = self._cache.setdefault(cache_scope, {})
            pending: dict[str, list[int]] = {}
            for text, indices in unique_texts.items():
                cached = l1.get(text)
                if cached is not None:
                    for index in indices:
                        results[index] = cached
                else:
                    pending[text] = indices

            if pending:
                try:
                    from .cache import get_cache
                    l2_hits = get_cache().get_batch(
                        self._cache_engine(src, target_lang), list(pending), target_lang,
                    )
                except Exception:
                    l2_hits = {}
                for text, translated in l2_hits.items():
                    l1[text] = translated
                    for index in pending.pop(text, []):
                        results[index] = translated

            if not pending:
                continue

            try:
                translator, sp_source, sp_target = self._load_model(src, target_lang)
            except FileNotFoundError:
                # Auto-detection can occasionally identify a short SKU/label as a
                # language for which the user has no model installed.  Keep those
                # units unchanged instead of failing the whole mixed document.
                logger.warning(
                    "No local OPUS-MT model for auto-detected %s->%s; keeping %d unique unit(s) unchanged",
                    src, target_lang, len(pending),
                )
                continue
            saved: list[tuple[str, str]] = []
            items = list(pending.items())
            for start in range(0, len(items), _BATCH_SIZE):
                if cancel_event is not None and cancel_event.is_set():
                    raise CancelledError("Translation cancelled")
                chunk = items[start:start + _BATCH_SIZE]
                tokenized = [self._tokenize(sp_source, text) for text, _ in chunk]
                outputs = self._translate_tokens(translator, tokenized, src=src, tgt=target_lang)
                if len(outputs) != len(chunk):
                    raise RuntimeError("Local translator returned an incomplete batch")
                for (text, indices), tokens in zip(chunk, outputs):
                    translated = self._detokenize(sp_target, tokens)
                    if not translated:
                        raise RuntimeError("Local translator returned an empty translation")
                    l1[text] = translated
                    saved.append((text, translated))
                    for index in indices:
                        results[index] = translated
            try:
                from .cache import get_cache
                get_cache().put_batch(self._cache_engine(src, target_lang), saved, target_lang)
            except Exception:
                pass
        return results

     
    # Public interface — Translator protocol
     
    def translate_text(self, text: str, target_lang: str) -> str:
        if not text or not text.strip():
            return text

        src = self._resolve_batch_src([text])
        if _same_language_code(src, target_lang):
            return text

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
        # Translate a list of strings, returning one result per input.

        # When 'source_lang' is 'auto', each non-empty unit is detected and grouped with units
        # that use the same installed source-to-target model.
        if not texts:
            return []

        if self._source_lang == "auto":
            return self._translate_auto_batch(texts, target_lang, cancel_event)

        results: list[str] = list(texts)

        # Decision 1: resolve source language once for the whole batch.
        src = self._resolve_batch_src(texts)
        if _same_language_code(src, target_lang):
            return list(texts)
        cache_scope = f"{src.lower()}:{target_lang.lower()}"
        tgt_cache = self._cache.setdefault(cache_scope, {})

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
                self._cache_engine(src, target_lang), [t for _, t in to_translate], target_lang,
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
            translated_tokens = self._translate_tokens(
                translator, tokens_batch, src=src, tgt=target_lang,
            )

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
                get_cache().put_batch(self._cache_engine(src, target_lang), l2_pairs, target_lang)
            except Exception:
                pass

        return results
