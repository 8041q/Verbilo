from typing import Optional, Dict, Any, List
from .base import Translator
import logging

logger = logging.getLogger(__name__)

# Maximum texts per Google Translate batch request.
_BATCH_SIZE = 50
# Maximum total characters per batch to stay within free-tier limits.
_BATCH_MAX_CHARS = 4500


class IdentityTranslator:
    """Returns text unchanged — useful for testing."""

    def translate_text(self, text: str, target_lang: str) -> str:
        return text

    def translate_batch(self, texts: list[str], target_lang: str) -> list[str]:
        return list(texts)


class DeepTranslatorWrapper:
    """Wraps *deep_translator.GoogleTranslator* with batching and optional
    source-language filtering via *langdetect*.

    Parameters
    ----------
    source_lang : str
        Language code for the source language (e.g. ``"en"``).
        ``"auto"`` means translate everything regardless of detected language.
    """

    def __init__(self, source_lang: str = "auto"):
        self._source_lang = source_lang
        try:
            from deep_translator import GoogleTranslator  # type: ignore
            self._impl_cls = GoogleTranslator
        except Exception:
            self._impl_cls = None
        self._instances: Dict[str, Any] = {}
        # Lazy-init langdetect
        self._langdetect = None
        if source_lang and source_lang != "auto":
            try:
                import langdetect as _ld
                _ld.DetectorFactory.seed = 0  # deterministic
                self._langdetect = _ld
            except ImportError:
                logger.warning("langdetect not installed; source-language filtering disabled")

    def _get_instance(self, target_lang: str):
        inst = self._instances.get(target_lang)
        if inst is None:
            src = self._source_lang if self._source_lang != "auto" else "auto"
            inst = self._impl_cls(source=src, target=target_lang)
            self._instances[target_lang] = inst
        return inst

    def _should_translate(self, text: str) -> bool:
        """Return *True* if the text appears to be in the source language."""
        if self._source_lang == "auto" or self._langdetect is None:
            return True
        try:
            detected = self._langdetect.detect(text)
            # langdetect returns ISO-639-1 codes, e.g. "en", "zh-cn"
            return detected.lower().startswith(self._source_lang.lower())
        except Exception:
            # If detection fails (very short text, etc.) translate it anyway
            return True

    # ----- single-item convenience ----- #

    def translate_text(self, text: str, target_lang: str) -> str:
        if not self._impl_cls or not text or not text.strip():
            return text
        if not self._should_translate(text):
            return text
        try:
            translator = self._get_instance(target_lang)
            result = translator.translate(text)
            return result if result is not None else text
        except Exception:
            logger.exception("DeepTranslator failed for target '%s'", target_lang)
            raise

    # ----- batch (the fast path) ----- #

    def translate_batch(self, texts: list[str], target_lang: str) -> list[str]:
        """Translate a list of strings, batching HTTP requests for speed.

        Empty / whitespace-only strings and strings not in the source language
        are returned unchanged without consuming an API call.
        """
        if not self._impl_cls:
            return list(texts)

        results: list[str] = [""] * len(texts)
        # Separate translatable vs pass-through
        to_translate: list[tuple[int, str]] = []  # (original_index, text)
        for i, t in enumerate(texts):
            if not t or not t.strip():
                results[i] = t
            elif not self._should_translate(t):
                results[i] = t
            else:
                to_translate.append((i, t))

        if not to_translate:
            return results

        translator = self._get_instance(target_lang)

        # Split into chunks respecting both count and character limits
        chunks = self._make_chunks(to_translate)

        for chunk in chunks:
            chunk_texts = [t for _, t in chunk]
            try:
                translated = translator.translate_batch(chunk_texts)
                if translated is None:
                    translated = chunk_texts
                for (orig_idx, orig_text), tr_text in zip(chunk, translated):
                    results[orig_idx] = tr_text if tr_text is not None else orig_text
            except Exception:
                logger.exception("Batch translation failed; falling back to per-item")
                for orig_idx, orig_text in chunk:
                    try:
                        r = translator.translate(orig_text)
                        results[orig_idx] = r if r is not None else orig_text
                    except Exception:
                        logger.exception("Per-item fallback also failed")
                        results[orig_idx] = orig_text

        return results

    @staticmethod
    def _make_chunks(items: list[tuple[int, str]]) -> list[list[tuple[int, str]]]:
        """Split *items* into chunks of at most ``_BATCH_SIZE`` items and
        ``_BATCH_MAX_CHARS`` total characters."""
        chunks: list[list[tuple[int, str]]] = []
        current: list[tuple[int, str]] = []
        current_chars = 0
        for item in items:
            text_len = len(item[1])
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
                from deep_translator import GoogleTranslator  # type: ignore
                return DeepTranslatorWrapper(source_lang=source_lang)
            except Exception:
                return IdentityTranslator()
        if name.lower() == "identity":
            return IdentityTranslator()
        if name.lower() in ("deep", "deep-translator", "google"):
            return DeepTranslatorWrapper(source_lang=source_lang)
        return IdentityTranslator()

