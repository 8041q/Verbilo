from typing import Protocol, runtime_checkable


@runtime_checkable
class Translator(Protocol):
    """Minimal interface that every translation backend must satisfy."""

    def translate_text(self, text: str, target_lang: str) -> str:
        """Translate a single text string to *target_lang*."""
        ...

    def translate_batch(self, texts: list[str], target_lang: str) -> list[str]:
        """Translate a list of text strings to *target_lang*.

        Implementations should handle chunking internally.  The default
        behaviour (when not overridden) falls back to per-item calls.
        """
        ...
