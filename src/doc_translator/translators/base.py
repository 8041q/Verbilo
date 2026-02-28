from typing import Protocol


class Translator(Protocol):
    def translate_text(self, text: str, target_lang: str) -> str:
        ...
