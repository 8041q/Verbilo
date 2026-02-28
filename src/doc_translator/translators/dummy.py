from typing import Optional, Dict, Any
from .base import Translator
import logging

logger = logging.getLogger(__name__)


class IdentityTranslator:
    def translate_text(self, text: str, target_lang: str) -> str:
        return text


class DeepTranslatorWrapper:
    # Wrapper around deep_translator.GoogleTranslator.
    
    # Caches translator instances per target language.
    # If deep-translator is not installed, behaves as identity (returns original text).
    # If instantiation or translation fails, logs and re-raises the exception.
    
    def __init__(self):
        try:
            from deep_translator import GoogleTranslator  # type: ignore

            self._impl_cls = GoogleTranslator
            self._instances: Dict[str, Any] = {}
        except Exception:
            # deep-translator not available; behave as identity
            self._impl_cls = None
            self._instances = {}

    def translate_text(self, text: str, target_lang: str) -> str:
        if not self._impl_cls:
            # deep-translator not installed; return original text
            return text
        try:
            # reuse translator instance for the same target_lang when possible
            translator = self._instances.get(target_lang)
            if translator is None:
                translator = self._impl_cls(source="auto", target=target_lang)
                self._instances[target_lang] = translator
            return translator.translate(text)
        except Exception as e:
            logger.exception("DeepTranslator failed for target '%s'", target_lang)
            raise


class TranslatorFactory:
    @staticmethod
    def get(name: Optional[str] = None) -> Translator:
        if name is None:
            # try deep-translator, else identity
            try:
                from deep_translator import GoogleTranslator  # type: ignore

                return DeepTranslatorWrapper()
            except Exception:
                return IdentityTranslator()
        if name.lower() == "identity":
            return IdentityTranslator()
        if name.lower() in ("deep", "deep-translator", "google"):
            return DeepTranslatorWrapper()
        return IdentityTranslator()
