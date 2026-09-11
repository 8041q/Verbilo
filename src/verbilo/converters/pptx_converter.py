from __future__ import annotations

import logging
import re
import tempfile
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Mapping
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from ..semantic import TranslationConstraints, TranslationContext, TranslationService, TranslationUnit
from ..utils import CancelledError

logger = logging.getLogger(__name__)

_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
_A_P = f"{{{_A_NS}}}p"
_A_R = f"{{{_A_NS}}}r"
_A_FLD = f"{{{_A_NS}}}fld"
_A_BR = f"{{{_A_NS}}}br"
_A_T = f"{{{_A_NS}}}t"
_A_RPR = f"{{{_A_NS}}}rPr"
_A_HLINK = f"{{{_A_NS}}}hlinkClick"
_A_TC = f"{{{_A_NS}}}tc"
_P_SP = f"{{{_P_NS}}}sp"
_P_NVPR = f"{{{_P_NS}}}nvPr"
_P_PH = f"{{{_P_NS}}}ph"
_P_SPPR = f"{{{_P_NS}}}spPr"
_A_XFRM = f"{{{_A_NS}}}xfrm"
_A_EXT = f"{{{_A_NS}}}ext"
_A_TXBODY = f"{{{_P_NS}}}txBody"
_A_BODYPR = f"{{{_A_NS}}}bodyPr"
_EMU_PER_POINT = 12700.0


@dataclass
class _PptxUnit:
    nodes: list[etree._Element]
    original_parts: list[str]
    source_text: str
    role: str
    mode: str
    constraints: TranslationConstraints
    part_name: str
    context: TranslationContext = field(default_factory=TranslationContext)
    translated_text: str | None = None


def _local_name(tag: str) -> str:
    try:
        return etree.QName(tag).localname
    except Exception:
        return tag


def _ancestor(node: etree._Element, tag: str) -> etree._Element | None:
    current = node.getparent()
    while current is not None:
        if current.tag == tag:
            return current
        current = current.getparent()
    return None


def _paragraph_islands(paragraph: etree._Element) -> list[list[etree._Element]]:
    islands: list[list[etree._Element]] = []
    current: list[etree._Element] = []

    def flush() -> None:
        nonlocal current
        if current:
            islands.append(current)
            current = []

    for child in paragraph:
        if child.tag == _A_BR:
            flush()
            continue
        if child.tag not in {_A_R, _A_FLD}:
            continue
        text_node = child.find(_A_T)
        if text_node is None or text_node.text is None:
            continue
        rpr = child.find(_A_RPR)
        has_hyperlink = rpr is not None and rpr.find(_A_HLINK) is not None
        if has_hyperlink:
            flush()
            islands.append([text_node])
            continue
        current.append(text_node)
    flush()

    if not islands:
        nodes = [node for node in paragraph.iter(_A_T) if node.text is not None]
        if nodes:
            islands.append(nodes)
    return islands


def _shape_role(node: etree._Element, part_name: str) -> tuple[str, str]:
    if part_name.startswith("ppt/notesSlides/"):
        return "footnote", "natural"
    if part_name.startswith("ppt/charts/") or part_name.startswith("ppt/diagrams/"):
        return "shape", "concise"
    if _ancestor(node, _A_TC) is not None:
        return "table-cell", "concise"

    shape = _ancestor(node, _P_SP)
    if shape is not None:
        nvpr = shape.find(f".//{_P_NVPR}")
        ph = nvpr.find(_P_PH) if nvpr is not None else None
        ph_type = (ph.get("type") if ph is not None else "") or ""
        if ph_type in {"title", "ctrTitle"}:
            return "heading", "faithful"
    return "shape", "concise"


def _font_points(nodes: list[etree._Element]) -> float:
    for node in nodes:
        parent = node.getparent()
        if parent is None:
            continue
        rpr = parent.find(_A_RPR)
        if rpr is not None:
            try:
                size = int(rpr.get("sz", "0"))
            except ValueError:
                size = 0
            if size > 0:
                return max(size / 100.0, 6.0)
    return 18.0


def _shape_constraints(nodes: list[etree._Element], source_text: str) -> TranslationConstraints:
    if not nodes:
        return TranslationConstraints()
    shape = _ancestor(nodes[0], _P_SP)
    if shape is None:
        return TranslationConstraints()
    sppr = shape.find(_P_SPPR)
    xfrm = sppr.find(_A_XFRM) if sppr is not None else None
    ext = xfrm.find(_A_EXT) if xfrm is not None else None
    if ext is None:
        return TranslationConstraints()
    try:
        cx = int(ext.get("cx", "0"))
        cy = int(ext.get("cy", "0"))
    except ValueError:
        return TranslationConstraints()
    if cx <= 0 or cy <= 0:
        return TranslationConstraints()

    # Account for text-body insets when they are explicitly present.
    txbody = shape.find(_A_TXBODY)
    bodypr = txbody.find(_A_BODYPR) if txbody is not None else None
    if bodypr is not None:
        for attr in ("lIns", "rIns"):
            try:
                cx -= max(int(bodypr.get(attr, "0")), 0)
            except ValueError:
                pass
        for attr in ("tIns", "bIns"):
            try:
                cy -= max(int(bodypr.get(attr, "0")), 0)
            except ValueError:
                pass
    if cx <= 0 or cy <= 0:
        return TranslationConstraints()

    font_pt = _font_points(nodes)
    width_pt = cx / _EMU_PER_POINT
    height_pt = cy / _EMU_PER_POINT
    estimated_lines = max(1, int(height_pt / max(font_pt * 1.2, 1.0)))
    chars_per_line = max(1, int(width_pt / max(font_pt * 0.55, 1.0)))
    source_visible = len(" ".join(source_text.split()))
    source_lines = max(1, source_text.count("\n") + 1)
    return TranslationConstraints(
        max_chars=max(source_visible, estimated_lines * chars_per_line),
        max_lines=max(source_lines, estimated_lines),
        source_visible_chars=source_visible,
        source_line_count=source_lines,
    )


def _part_kind(part_name: str) -> str:
    if part_name.startswith("ppt/slides/"):
        return "slide"
    if part_name.startswith("ppt/notesSlides/"):
        return "notes"
    if part_name.startswith("ppt/charts/"):
        return "chart"
    if part_name.startswith("ppt/diagrams/"):
        return "diagram"
    return "presentation"


def _collect_part_units(root: etree._Element, part_name: str) -> list[_PptxUnit]:
    units: list[_PptxUnit] = []
    for paragraph in root.iter(_A_P):
        for nodes in _paragraph_islands(paragraph):
            parts = [node.text or "" for node in nodes]
            original = "".join(parts)
            if not original or not original.strip():
                continue
            leading = original[: len(original) - len(original.lstrip())]
            trailing = original[len(original.rstrip()) :]
            core = original[len(leading): len(original) - len(trailing) if trailing else len(original)]
            if not core.strip():
                continue
            role, mode = _shape_role(nodes[0], part_name)
            units.append(
                _PptxUnit(
                    nodes=nodes,
                    original_parts=parts,
                    source_text=core,
                    role=role,
                    mode=mode,
                    constraints=_shape_constraints(nodes, core),
                    part_name=part_name,
                )
            )

    section: str | None = None
    for idx, unit in enumerate(units):
        unit.context = TranslationContext.from_values(
            section=section,
            before=(units[idx - 1].source_text,) if idx > 0 else (),
            after=(units[idx + 1].source_text,) if idx + 1 < len(units) else (),
            metadata={"format": "pptx", "part": part_name, "part_kind": _part_kind(part_name)},
        )
        if unit.role == "heading":
            section = unit.source_text
    return units


def _redistribute_text(parts: list[str], translated: str) -> list[str]:
    if not parts:
        return []
    if len(parts) == 1:
        return [translated]
    visible = [idx for idx, value in enumerate(parts) if value]
    if not visible:
        return [translated] + [""] * (len(parts) - 1)
    total = sum(len(parts[idx]) for idx in visible) or 1
    target_len = len(translated)
    boundaries: list[int] = []
    cumulative = 0
    previous = 0
    safe = set(" \t,.;:!?/)-–—")
    for idx in visible[:-1]:
        cumulative += len(parts[idx])
        ideal = round(target_len * cumulative / total)
        best = max(previous, min(target_len, ideal))
        best_distance = abs(best - ideal)
        radius = min(20, target_len)
        for pos in range(max(previous, ideal - radius), min(target_len, ideal + radius) + 1):
            if pos <= previous:
                continue
            if (pos > 0 and translated[pos - 1] in safe) or (pos < target_len and translated[pos] in safe):
                distance = abs(pos - ideal)
                if distance < best_distance:
                    best = pos
                    best_distance = distance
        boundaries.append(best)
        previous = best
    starts = [0] + boundaries
    ends = boundaries + [target_len]
    chunks = [translated[a:b] for a, b in zip(starts, ends)]
    output = [""] * len(parts)
    for idx, chunk in zip(visible, chunks):
        output[idx] = chunk
    return output


def _serialize_xml(root: etree._Element, original: bytes) -> bytes:
    encoding = "UTF-8"
    if original.startswith(b"<?xml"):
        try:
            declaration = original[: original.index(b"?>") + 2].decode("ascii", errors="ignore")
            match = re.search(r"encoding=[\"']([^\"']+)", declaration, re.I)
            if match:
                encoding = match.group(1)
        except Exception:
            pass
    standalone = b'standalone="yes"' in original[:200] or b"standalone='yes'" in original[:200]
    return etree.tostring(
        root,
        xml_declaration=True,
        encoding=encoding,
        standalone=True if standalone else None,
    )


def _rebuild_package(input_path: str, output_path: str, patches: dict[str, bytes]) -> None:
    if not patches and Path(input_path).resolve() != Path(output_path).resolve():
        import shutil
        shutil.copy2(input_path, output_path)
        return
    out_dir = str(Path(output_path).resolve().parent)
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(prefix=".verbilo-pptx-", suffix=".tmp", dir=out_dir, delete=False) as tmp:
        temp_path = tmp.name
    try:
        with ZipFile(input_path, "r") as zin, ZipFile(temp_path, "w") as zout:
            zout.comment = zin.comment
            for info in zin.infolist():
                zout.writestr(info, patches.get(info.filename, zin.read(info.filename)))
        Path(temp_path).replace(output_path)
    except Exception:
        try:
            Path(temp_path).unlink(missing_ok=True)
        finally:
            raise


def translate_pptx(
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
    translation_memory: Any | None = None,
) -> None:
    if cancel_event is not None and cancel_event.is_set():
        raise CancelledError("Translation cancelled before starting")

    roots: dict[str, tuple[etree._Element, bytes]] = {}
    units: list[_PptxUnit] = []
    with ZipFile(input_path, "r") as zin:
        names = set(zin.namelist())
        candidates = sorted(
            name for name in names
            if name.endswith(".xml") and (
                name.startswith("ppt/slides/")
                or name.startswith("ppt/notesSlides/")
                or name.startswith("ppt/charts/")
                or name.startswith("ppt/diagrams/")
            )
        )
        for name in candidates:
            raw = zin.read(name)
            if b":t>" not in raw and b"<a:t" not in raw:
                continue
            try:
                root = etree.fromstring(raw, parser=etree.XMLParser(resolve_entities=False, remove_blank_text=False))
            except Exception:
                logger.warning("Skipping malformed PPTX XML part %s", name, exc_info=True)
                continue
            part_units = _collect_part_units(root, name)
            if part_units:
                roots[name] = (root, raw)
                units.extend(part_units)

    if not units:
        _rebuild_package(input_path, output_path, {})
        return

    semantic_units = [
        TranslationUnit(
            text=unit.source_text,
            source_lang=source_lang,
            role=unit.role,
            mode=unit.mode,  # type: ignore[arg-type]
            constraints=unit.constraints,
            context=unit.context,
            metadata={"format": "pptx", "part": unit.part_name},
        )
        for unit in units
    ]
    result = TranslationService(translator, terminology=terminology, translation_memory=translation_memory).translate_units(
        semantic_units, target_lang, cancel_event=cancel_event
    )

    changed_parts: set[str] = set()
    errors = 0
    for index, (unit, translated) in enumerate(zip(units, result.texts), start=1):
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError("Translation cancelled")
        if translated is None:
            errors += 1
        else:
            original_joined = "".join(unit.original_parts)
            leading = original_joined[: len(original_joined) - len(original_joined.lstrip())]
            trailing = original_joined[len(original_joined.rstrip()) :]
            final_text = leading + translated.strip() + trailing
            if final_text != original_joined:
                chunks = _redistribute_text(unit.original_parts, final_text)
                for node, chunk in zip(unit.nodes, chunks):
                    node.text = chunk
                    xml_space = "{http://www.w3.org/XML/1998/namespace}space"
                    if chunk[:1].isspace() or chunk[-1:].isspace():
                        node.set(xml_space, "preserve")
                changed_parts.add(unit.part_name)
                unit.translated_text = translated
        if progress_callback is not None:
            progress_callback(index, len(units))

    patches = {
        name: _serialize_xml(root, original)
        for name, (root, original) in roots.items()
        if name in changed_parts
    }
    if cancel_event is not None and cancel_event.is_set():
        raise CancelledError("Translation cancelled before saving PPTX")
    _rebuild_package(input_path, output_path, patches)

    if errors:
        message = f"Translation completed with {errors} failed PPTX text units; originals were preserved"
        if strict_errors:
            raise RuntimeError(message)
        logger.warning(message)
