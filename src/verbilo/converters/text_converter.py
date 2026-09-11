from __future__ import annotations

import codecs
import logging
import re
import threading
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Any, Callable, Mapping

from ..semantic import TranslationContext, TranslationService, TranslationUnit
from ..utils import CancelledError

logger = logging.getLogger(__name__)


@dataclass
class _TextSegment:
    prefix: str
    core: str
    suffix: str
    newline: str
    role: str = "body"
    mode: str = "natural"
    context: TranslationContext = TranslationContext()


def _decode_text(data: bytes) -> tuple[str, str, bytes]:
    if data.startswith(codecs.BOM_UTF8):
        return data[len(codecs.BOM_UTF8):].decode("utf-8"), "utf-8", codecs.BOM_UTF8
    if data.startswith(codecs.BOM_UTF16_LE):
        return data[len(codecs.BOM_UTF16_LE):].decode("utf-16-le"), "utf-16-le", codecs.BOM_UTF16_LE
    if data.startswith(codecs.BOM_UTF16_BE):
        return data[len(codecs.BOM_UTF16_BE):].decode("utf-16-be"), "utf-16-be", codecs.BOM_UTF16_BE
    return data.decode("utf-8"), "utf-8", b""


def _split_line_ending(line: str) -> tuple[str, str]:
    if line.endswith("\r\n"):
        return line[:-2], "\r\n"
    if line.endswith("\n") or line.endswith("\r"):
        return line[:-1], line[-1:]
    return line, ""


def _looks_like_heading(text: str, *, next_is_blank: bool) -> bool:
    stripped = text.strip()
    if not stripped or len(stripped) > 100:
        return False
    if stripped.endswith((".", "?", "!", ";", ":", "。", "！", "？")):
        return False
    alpha = [ch for ch in stripped if ch.isalpha()]
    all_caps = bool(alpha) and all(not ch.islower() for ch in alpha)
    return all_caps or next_is_blank


def _prepare_segments(text: str) -> tuple[list[str | _TextSegment], list[_TextSegment]]:
    lines = text.splitlines(keepends=True)
    if not lines and text == "":
        return [], []
    if text and (not lines or "".join(lines) != text):
        lines = text.splitlines(keepends=True)

    rendered: list[str | _TextSegment] = []
    segments: list[_TextSegment] = []
    for idx, raw in enumerate(lines):
        body, newline = _split_line_ending(raw)
        if not body.strip():
            rendered.append(raw)
            continue
        leading = body[: len(body) - len(body.lstrip(" \t"))]
        trailing = body[len(body.rstrip(" \t")) :]
        core = body[len(leading): len(body) - len(trailing) if trailing else len(body)]
        next_is_blank = idx + 1 < len(lines) and not _split_line_ending(lines[idx + 1])[0].strip()
        heading = _looks_like_heading(core, next_is_blank=next_is_blank)
        segment = _TextSegment(
            prefix=leading,
            core=core,
            suffix=trailing,
            newline=newline,
            role="heading" if heading else "body",
            mode="faithful" if heading else "natural",
        )
        rendered.append(segment)
        segments.append(segment)

    _populate_context(segments)
    return rendered, segments


def _populate_context(segments: list[_TextSegment]) -> None:
    section: str | None = None
    for idx, segment in enumerate(segments):
        before = (segments[idx - 1].core,) if idx > 0 else ()
        after = (segments[idx + 1].core,) if idx + 1 < len(segments) else ()
        segment.context = TranslationContext.from_values(
            section=section,
            before=before,
            after=after,
            metadata={"format": "text"},
        )
        if segment.role == "heading":
            section = segment.core


def translate_txt(
    input_path: str,
    output_path: str,
    translator: Any,
    target_lang: str,
    *,
    cancel_event: threading.Event | None = None,
    source_lang: str = "auto",
    progress_callback: Callable[[int, int], None] | None = None,
    terminology: Mapping[str, str] | None = None,
    strict_errors: bool = False,
) -> None:
    if cancel_event is not None and cancel_event.is_set():
        raise CancelledError("Translation cancelled before starting")

    raw = Path(input_path).read_bytes()
    text, encoding, bom = _decode_text(raw)
    rendered, segments = _prepare_segments(text)
    if not segments:
        Path(output_path).write_bytes(raw)
        return

    units = [
        TranslationUnit(
            text=segment.core,
            source_lang=source_lang,
            role=segment.role,
            mode=segment.mode,  # type: ignore[arg-type]
            context=segment.context,
            metadata={"format": "txt"},
        )
        for segment in segments
    ]
    result = TranslationService(translator, terminology=terminology).translate_units(
        units, target_lang, cancel_event=cancel_event
    )

    translated_by_id: dict[int, str] = {}
    errors = 0
    for index, (segment, translated) in enumerate(zip(segments, result.texts), start=1):
        if translated is None:
            errors += 1
            translated_by_id[id(segment)] = segment.core
        else:
            translated_by_id[id(segment)] = translated
        if progress_callback is not None:
            progress_callback(index, len(segments))

    output_parts: list[str] = []
    for item in rendered:
        if isinstance(item, str):
            output_parts.append(item)
        else:
            output_parts.append(
                item.prefix + translated_by_id[id(item)] + item.suffix + item.newline
            )
    output_text = "".join(output_parts)
    Path(output_path).write_bytes(bom + output_text.encode(encoding))

    if errors:
        message = f"Translation completed with {errors} failed text lines; originals were preserved"
        if strict_errors:
            raise RuntimeError(message)
        logger.warning(message)
