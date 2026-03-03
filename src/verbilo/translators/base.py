import threading
from typing import Optional, Protocol, runtime_checkable


# all backends must implement translate_text and translate_batch
@runtime_checkable
class Translator(Protocol):

    def translate_text(self, text: str, target_lang: str) -> str:
        ...

    def translate_batch(
        self,
        texts: list[str],
        target_lang: str,
        *,
        cancel_event: Optional[threading.Event] = None,
    ) -> list[str]:
        # handle chunking internally; raises CancelledError if cancel_event fires
        ...
