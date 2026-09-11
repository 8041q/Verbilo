from __future__ import annotations

import logging
import re
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Mapping

from ..semantic import ProtectedText, TranslationContext, TranslationService, TranslationUnit
from ..progress import ProgressReporter, ProgressUpdate
from ..utils import CancelledError
from .text_converter import _decode_text, _split_line_ending

logger = logging.getLogger(__name__)

_FENCE_RE = re.compile(r"^\s{0,3}(`{3,}|~{3,})")
_HEADING_RE = re.compile(r"^(\s{0,3}#{1,6}[ \t]+)(.*?)([ \t]+#+[ \t]*)?$")
_LIST_RE = re.compile(r"^(\s*(?:>\s*)*(?:(?:[-+*]|\d+[.)])[ \t]+(?:\[[ xX]\][ \t]+)?))(.*)$")
_QUOTE_RE = re.compile(r"^(\s*(?:>\s*)+)(.*)$")
_TABLE_SEPARATOR_RE = re.compile(r"^\s*\|?\s*:?-{3,}:?\s*(?:\|\s*:?-{3,}:?\s*)+\|?\s*$")
_THEMATIC_RE = re.compile(r"^\s{0,3}(?:(?:\*\s*){3,}|(?:-\s*){3,}|(?:_\s*){3,})$")
_REFERENCE_DEF_RE = re.compile(r"^\s{0,3}\[[^\]]+\]:\s*\S+")
_HTML_COMMENT_RE = re.compile(r"^\s*<!--.*-->\s*$")
_SETEXT_RE = re.compile(r"^\s{0,3}(?:=+|-+)\s*$")
_EXISTING_MD_TOKEN_RE = re.compile(r"⟦G(\d+)⟧")


@dataclass
class _MarkdownSegment:
    prefix: str
    core: str
    suffix: str
    newline: str
    role: str
    mode: str
    protected: ProtectedText
    context: TranslationContext = TranslationContext()


def _protect_markdown_inline(text: str) -> ProtectedText:
    token_map: dict[str, str] = {}
    existing = [int(m.group(1)) for m in _EXISTING_MD_TOKEN_RE.finditer(text)]
    counter = max(existing, default=-1) + 1

    def protect(match: re.Match[str]) -> str:
        nonlocal counter
        token = f"⟦G{counter}⟧"
        counter += 1
        token_map[token] = match.group(0)
        return token

    masked = text
    patterns = [
        # Inline code and images are non-translatable payloads.
        re.compile(r"`+[^`\n]+`+"),
        re.compile(r"!\[[^\]\n]*\]\([^\n)]*\)"),
        re.compile(r"<https?://[^>\s]+>"),
        re.compile(r"</?[A-Za-z][^>\n]*>"),
        # Keep Markdown link destinations intact while allowing labels to translate.
        re.compile(r"(?<=\])\([^\n)]*\)"),
        re.compile(r"https?://[^\s<>()]+"),
        # Backslash escapes and structural punctuation must survive exactly.
        re.compile(r"\\[\\`*_{}\[\]()#+.!|>-]"),
        re.compile(r"(?:\*\*|__|~~|\*|_|\||\[|\])"),
    ]
    for pattern in patterns:
        masked = pattern.sub(protect, masked)
    return ProtectedText(masked, token_map)


def _line_segment(body: str, newline: str) -> _MarkdownSegment | None:
    if not body.strip():
        return None
    if _TABLE_SEPARATOR_RE.match(body) or _THEMATIC_RE.match(body):
        return None
    if _REFERENCE_DEF_RE.match(body) or _HTML_COMMENT_RE.match(body):
        return None

    prefix = ""
    suffix = ""
    core = body
    role = "body"
    mode = "natural"

    heading = _HEADING_RE.match(body)
    if heading:
        prefix = heading.group(1)
        core = heading.group(2)
        suffix = heading.group(3) or ""
        role = "heading"
        mode = "faithful"
    else:
        listing = _LIST_RE.match(body)
        quote = _QUOTE_RE.match(body) if listing is None else None
        if listing:
            prefix, core = listing.group(1), listing.group(2)
        elif quote:
            prefix, core = quote.group(1), quote.group(2)
        else:
            leading = body[: len(body) - len(body.lstrip(" \t"))]
            trailing = body[len(body.rstrip(" \t")) :]
            prefix = leading
            suffix = trailing
            core = body[len(leading): len(body) - len(trailing) if trailing else len(body)]

    if not core.strip():
        return None
    if "|" in core:
        role = "table-cell"
        mode = "concise"
    protected = _protect_markdown_inline(core)
    return _MarkdownSegment(prefix, core, suffix, newline, role, mode, protected)


def _prepare_markdown(text: str) -> tuple[list[str | _MarkdownSegment], list[_MarkdownSegment]]:
    rendered: list[str | _MarkdownSegment] = []
    segments: list[_MarkdownSegment] = []
    lines = text.splitlines(keepends=True)
    fence: tuple[str, int] | None = None
    front_matter = False
    html_comment = False

    for index, raw in enumerate(lines):
        body, newline = _split_line_ending(raw)

        # YAML/TOML-style front matter at the start of a Markdown document is
        # metadata, not prose. Preserve it byte-for-byte.
        if index == 0 and body.strip() == "---":
            front_matter = True
            rendered.append(raw)
            continue
        if front_matter:
            rendered.append(raw)
            if body.strip() in {"---", "..."}:
                front_matter = False
            continue

        fence_match = _FENCE_RE.match(body)
        if fence_match:
            marker = fence_match.group(1)
            if fence is None:
                fence = (marker[0], len(marker))
                rendered.append(raw)
                continue
            if marker[0] == fence[0] and len(marker) >= fence[1]:
                fence = None
                rendered.append(raw)
                continue
        if fence is not None:
            rendered.append(raw)
            continue

        # Preserve indented code blocks and multi-line HTML comments.
        if html_comment:
            rendered.append(raw)
            if "-->" in body:
                html_comment = False
            continue
        if "<!--" in body and "-->" not in body:
            html_comment = True
            rendered.append(raw)
            continue
        if body.startswith("    ") or body.startswith("\t"):
            rendered.append(raw)
            continue
        if _SETEXT_RE.match(body) and rendered and isinstance(rendered[-1], _MarkdownSegment):
            rendered.append(raw)
            continue

        next_body = ""
        if index + 1 < len(lines):
            next_body = _split_line_ending(lines[index + 1])[0]
        setext_heading = bool(body.strip() and _SETEXT_RE.match(next_body))

        segment = _line_segment(body, newline)
        if segment is None:
            rendered.append(raw)
        else:
            if setext_heading:
                segment.role = "heading"
                segment.mode = "faithful"
            rendered.append(segment)
            segments.append(segment)

    section: str | None = None
    for idx, segment in enumerate(segments):
        segment.context = TranslationContext.from_values(
            section=section,
            before=(segments[idx - 1].core,) if idx > 0 else (),
            after=(segments[idx + 1].core,) if idx + 1 < len(segments) else (),
            metadata={"format": "markdown"},
        )
        if segment.role == "heading":
            section = segment.core
    return rendered, segments


def translate_markdown(
    input_path: str,
    output_path: str,
    translator: Any,
    target_lang: str,
    *,
    cancel_event: threading.Event | None = None,
    source_lang: str = "auto",
    progress_callback: Callable[[int, int], None] | None = None,
    progress_event_callback: Callable[[ProgressUpdate], None] | None = None,
    terminology: Mapping[str, str] | None = None,
    strict_errors: bool = False,
    translation_memory: Any | None = None,
) -> None:
    progress = ProgressReporter(progress_event_callback)
    progress.update("analyzing", 0, 1)
    if cancel_event is not None and cancel_event.is_set():
        raise CancelledError("Translation cancelled before starting")

    raw = Path(input_path).read_bytes()
    text, encoding, bom = _decode_text(raw)
    rendered, segments = _prepare_markdown(text)
    progress.update("analyzing", 1, 1, detail=f"{len(segments)} text unit(s)")
    if not segments:
        progress.update("saving", 0, 1)
        Path(output_path).write_bytes(raw)
        progress.complete()
        return

    units = [
        TranslationUnit(
            text=segment.core,
            source_lang=source_lang,
            role=segment.role,
            mode=segment.mode,  # type: ignore[arg-type]
            context=segment.context,
            protected=segment.protected,
            metadata={"format": "markdown"},
        )
        for segment in segments
    ]
    progress.update("translating", 0, len(units), detail=f"{len(units)} text unit(s)")
    result = TranslationService(translator, terminology=terminology, translation_memory=translation_memory).translate_units(
        units,
        target_lang,
        cancel_event=cancel_event,
        progress_callback=lambda done, total: progress.update("translating", done, total),
    )

    translated_by_id: dict[int, str] = {}
    errors = 0
    progress.update("layout", 0, len(segments))
    for index, (segment, translated) in enumerate(zip(segments, result.texts), start=1):
        if translated is None:
            errors += 1
            translated_by_id[id(segment)] = segment.core
        else:
            translated_by_id[id(segment)] = translated
        if progress_callback is not None:
            progress_callback(index, len(segments))
        progress.update("layout", index, len(segments))

    output_parts: list[str] = []
    for item in rendered:
        if isinstance(item, str):
            output_parts.append(item)
        else:
            output_parts.append(
                item.prefix + translated_by_id[id(item)] + item.suffix + item.newline
            )
    progress.update("saving", 0, 1)
    Path(output_path).write_bytes(bom + "".join(output_parts).encode(encoding))
    progress.complete()

    if errors:
        message = f"Translation completed with {errors} failed Markdown lines; originals were preserved"
        if strict_errors:
            raise RuntimeError(message)
        logger.warning(message)
