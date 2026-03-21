import re
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


# ── Inline-tag conversion for HTML-aware translation APIs ─────────────────────
# The DOCX and PDF converters wrap formatting runs/spans in Unicode bracket
# tags: ⟨rN⟩...⟨/rN⟩ (DOCX) and ⟨sN⟩...⟨/sN⟩ (PDF).  APIs that support
# HTML tag preservation (Azure textType="html", DeepL tag_handling="html")
# can protect these markers by converting them to real HTML span elements
# before the API call, then converting back afterwards.

_TAG_OPEN = "\u27E8"
_TAG_CLOSE = "\u27E9"
_UNICODE_TAG_RE = re.compile(
    r"\u27E8([rs]\d+)\u27E9(.*?)\u27E8/\1\u27E9", re.DOTALL
)
_HTML_SPAN_TAG_RE = re.compile(
    r'<span class="([rs]\d+)">(.*?)</span>', re.DOTALL
)


def has_inline_tags(text: str) -> bool:
    # Return True if *text* contains our Unicode bracket tags
    return _TAG_OPEN in text and _TAG_CLOSE in text


def unicode_tags_to_html(text: str) -> str:
    # Convert ``⟨rN⟩...⟨/rN⟩`` to ``<span class="rN">...</span>``
    return _UNICODE_TAG_RE.sub(
        lambda m: f'<span class="{m.group(1)}">{m.group(2)}</span>',
        text,
    )


def html_tags_to_unicode(text: str) -> str:
    # Convert ``<span class="rN">...</span>`` back to ``⟨rN⟩...⟨/rN⟩``
    return _HTML_SPAN_TAG_RE.sub(
        lambda m: f"{_TAG_OPEN}{m.group(1)}{_TAG_CLOSE}"
        f"{m.group(2)}"
        f"{_TAG_OPEN}/{m.group(1)}{_TAG_CLOSE}",
        text,
    )
