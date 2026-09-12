from __future__ import annotations

import logging
import math
import re
import tempfile
import threading
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Any, Callable, Mapping
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from ..semantic import TranslationConstraints, TranslationContext, TranslationService, TranslationUnit
from ..progress import ProgressReporter, ProgressUpdate
from ..output_validation import OutputValidationMetrics
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
_A_TCPR = f"{{{_A_NS}}}tcPr"
_A_TBL = f"{{{_A_NS}}}tbl"
_A_TBLGRID = f"{{{_A_NS}}}tblGrid"
_A_GRIDCOL = f"{{{_A_NS}}}gridCol"
_A_TR = f"{{{_A_NS}}}tr"
_A_DRAWING_TXBODY = f"{{{_A_NS}}}txBody"
_P_SP = f"{{{_P_NS}}}sp"
_P_NVPR = f"{{{_P_NS}}}nvPr"
_P_PH = f"{{{_P_NS}}}ph"
_P_SPPR = f"{{{_P_NS}}}spPr"
_A_XFRM = f"{{{_A_NS}}}xfrm"
_A_EXT = f"{{{_A_NS}}}ext"
_A_TXBODY = f"{{{_P_NS}}}txBody"
_A_BODYPR = f"{{{_A_NS}}}bodyPr"
_A_NOAUTOFIT = f"{{{_A_NS}}}noAutofit"
_A_NORMAUTOFIT = f"{{{_A_NS}}}normAutofit"
_A_SPAUTOFIT = f"{{{_A_NS}}}spAutoFit"
_A_SCENE3D = f"{{{_A_NS}}}scene3d"
_A_SP3D = f"{{{_A_NS}}}sp3d"
_A_EXTLST = f"{{{_A_NS}}}extLst"
_EMU_PER_POINT = 12700.0
_LAYOUT_RETRY_CHAR_RATIO = 0.88
_AUTOFIT_MIN_SCALE = 0.72
_AUTOFIT_TRIGGER_SCALE = 0.985


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
    txbody: etree._Element | None = None
    if shape is not None:
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
        txbody = shape.find(_A_TXBODY)
    else:
        tc = _ancestor(nodes[0], _A_TC)
        row = tc.getparent() if tc is not None else None
        table = _ancestor(tc, _A_TBL) if tc is not None else None
        if tc is None or row is None or row.tag != _A_TR or table is None:
            return TranslationConstraints()
        grid = table.find(_A_TBLGRID)
        cols = list(grid.findall(_A_GRIDCOL)) if grid is not None else []
        cells = [child for child in row if child.tag == _A_TC]
        try:
            cell_idx = cells.index(tc)
        except ValueError:
            return TranslationConstraints()
        start_col = 0
        for prior in cells[:cell_idx]:
            try:
                start_col += max(int(prior.get("gridSpan", "1")), 1)
            except ValueError:
                start_col += 1
        try:
            span = max(int(tc.get("gridSpan", "1")), 1)
            cy = int(row.get("h", "0"))
            cx = sum(int(col.get("w", "0")) for col in cols[start_col:start_col + span])
        except ValueError:
            return TranslationConstraints()
        txbody = tc.find(_A_DRAWING_TXBODY)
        tcpr = tc.find(_A_TCPR)
        if tcpr is not None:
            for attr in ("marL", "marR"):
                try:
                    cx -= max(int(tcpr.get(attr, "0")), 0)
                except ValueError:
                    pass
            for attr in ("marT", "marB"):
                try:
                    cy -= max(int(tcpr.get(attr, "0")), 0)
                except ValueError:
                    pass

    if cx <= 0 or cy <= 0:
        return TranslationConstraints()

    # Account for text-body insets when they are explicitly present.
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


def _constraint_load_components(
    text: str,
    constraints: TranslationConstraints,
) -> tuple[float, float]:
    normalized = constraints.normalized(text)
    char_load = 0.0
    line_load = 0.0
    if normalized.max_chars:
        visible_chars = len(" ".join(str(text or "").split()))
        char_load = visible_chars / max(normalized.max_chars, 1)
    if normalized.max_lines:
        visible_lines = max(1, len(str(text or "").splitlines()))
        line_load = visible_lines / max(normalized.max_lines, 1)
    return char_load, line_load


def _constraint_load(text: str, constraints: TranslationConstraints) -> float:
    char_load, line_load = _constraint_load_components(text, constraints)
    return max(char_load, line_load)


def _tighten_constraints(
    constraints: TranslationConstraints,
    source_text: str,
) -> TranslationConstraints:
    normalized = constraints.normalized(source_text)
    max_chars = normalized.max_chars
    if max_chars is not None and max_chars > 1:
        max_chars = max(
            1,
            min(max_chars - 1, int(math.floor(max_chars * _LAYOUT_RETRY_CHAR_RATIO))),
        )
    return TranslationConstraints(
        max_chars=max_chars,
        max_lines=normalized.max_lines,
        source_visible_chars=normalized.source_visible_chars,
        source_line_count=normalized.source_line_count,
    )


def _text_container_for_nodes(nodes: list[etree._Element]) -> etree._Element | None:
    if not nodes:
        return None
    shape = _ancestor(nodes[0], _P_SP)
    if shape is not None:
        return shape
    return _ancestor(nodes[0], _A_TC)


def _supports_layout_guidance(translator: Any) -> bool:
    explicit = getattr(translator, "supports_layout_constraints", None)
    if explicit is not None:
        return bool(explicit)
    return any(
        callable(getattr(translator, name, None))
        for name in ("translate_units", "translate_blocks")
    )


def _uses_shape_autofit(container: etree._Element) -> bool:
    txbody = (
        container.find(_A_TXBODY)
        if container.tag == _P_SP
        else container.find(_A_DRAWING_TXBODY)
    )
    bodypr = txbody.find(_A_BODYPR) if txbody is not None else None
    return bodypr is not None and bodypr.find(_A_SPAUTOFIT) is not None


def _apply_normal_autofit(container: etree._Element, scale: float) -> bool:
    txbody = (
        container.find(_A_TXBODY)
        if container.tag == _P_SP
        else container.find(_A_DRAWING_TXBODY)
    )
    bodypr = txbody.find(_A_BODYPR) if txbody is not None else None
    if bodypr is None:
        return False

    scale = max(_AUTOFIT_MIN_SCALE, min(float(scale), 1.0))
    if scale >= _AUTOFIT_TRIGGER_SCALE:
        return False

    if bodypr.find(_A_SPAUTOFIT) is not None:
        # Respect an existing author choice to grow the shape itself.
        return False

    desired = int(round(scale * 100000))
    existing_normal = bodypr.find(_A_NORMAUTOFIT)
    if existing_normal is not None:
        try:
            current = int(existing_normal.get("fontScale", "100000"))
        except ValueError:
            current = 100000
        existing_normal.set("fontScale", str(min(current, desired)))
        return True

    existing_none = bodypr.find(_A_NOAUTOFIT)
    if existing_none is not None:
        bodypr.remove(existing_none)

    autofit = etree.Element(_A_NORMAUTOFIT)
    autofit.set("fontScale", str(desired))
    insert_at = len(bodypr)
    for idx, child in enumerate(bodypr):
        if child.tag in {_A_SCENE3D, _A_SP3D, _A_EXTLST}:
            insert_at = idx
            break
    bodypr.insert(insert_at, autofit)
    return True


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
    progress_event_callback: Callable[[ProgressUpdate], None] | None = None,
    terminology: Mapping[str, str] | None = None,
    strict_errors: bool = False,
    translation_memory: Any | None = None,
    translation_cache: Any | None = None,
    metrics_callback: Callable[[Any], None] | None = None,
) -> None:
    progress = ProgressReporter(progress_event_callback)
    progress.update("analyzing", 0, 1)
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

    progress.update("analyzing", 1, 1, detail=f"{len(units)} text unit(s)")

    if not units:
        progress.update("saving", 0, 1)
        _rebuild_package(input_path, output_path, {})
        progress.complete()
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
    service = TranslationService(
        translator,
        terminology=terminology,
        translation_memory=translation_memory, persistent_cache=translation_cache,
        metrics_callback=metrics_callback,
    )
    progress.update("translating", 0, len(semantic_units), detail=f"{len(semantic_units)} text unit(s)")
    result = service.translate_units(
        semantic_units,
        target_lang,
        cancel_event=cancel_event,
        progress_callback=lambda done, total: progress.update("translating", done, total),
    )

    translated_texts = list(result.texts)
    visual_metrics = OutputValidationMetrics()
    retry_indices: list[int] = []
    retry_units: list[TranslationUnit] = []

    # Slide text that exceeds the estimated shape capacity gets one more request
    # with a tighter concise budget before we touch visual formatting. Notes,
    # charts, and diagram parts intentionally keep their existing behavior.
    for idx, (pptx_unit, semantic_unit, translated) in enumerate(
        zip(units, semantic_units, translated_texts)
    ):
        if translated is None or not pptx_unit.part_name.startswith("ppt/slides/"):
            continue
        if semantic_unit.constraints.max_chars is None:
            continue
        if _constraint_load(str(translated), semantic_unit.constraints) <= 1.0:
            continue
        tightened = _tighten_constraints(semantic_unit.constraints, semantic_unit.source_text)
        retry_indices.append(idx)
        retry_units.append(
            replace(
                semantic_unit,
                mode="concise",
                constraints=tightened,
                metadata={**dict(semantic_unit.metadata), "layout_retry": "1"},
            )
        )

    visual_metrics.layout_retry_candidates = len(retry_units)
    accepted_retries = 0
    if retry_units and _supports_layout_guidance(translator):
        layout_total = len(retry_units) + len(units)
        progress.update("layout", 0, layout_total, detail=f"{len(retry_units)} compact retry unit(s)")
        retry_service = TranslationService(
            translator, terminology=terminology, translation_memory=translation_memory, persistent_cache=translation_cache
        )
        retry_result = retry_service.translate_units(
            retry_units,
            target_lang,
            cancel_event=cancel_event,
            progress_callback=lambda done, total: progress.update("layout", done, layout_total),
        )
        for source_idx, candidate in zip(retry_indices, retry_result.texts):
            if candidate is None:
                continue
            current = translated_texts[source_idx]
            if current is None:
                translated_texts[source_idx] = candidate
                accepted_retries += 1
                continue
            constraints = semantic_units[source_idx].constraints
            if _constraint_load(str(candidate), constraints) + 0.02 < _constraint_load(str(current), constraints):
                translated_texts[source_idx] = candidate
                accepted_retries += 1
        visual_metrics.layout_retry_accepted = accepted_retries
        logger.info(
            "PPTX layout retry examined %d unit(s) and accepted %d shorter fit(s)",
            len(retry_units),
            accepted_retries,
        )
        layout_base = len(retry_units)
    else:
        layout_total = len(units)
        layout_base = 0
        progress.update("layout", 0, layout_total)

    # If a slide shape still overflows after the concise retry, add bounded
    # DrawingML normal-autofit.  Use the most conservative requested scale when
    # multiple translation islands share the same text box.
    shape_loads: dict[int, dict[str, Any]] = {}
    for pptx_unit, semantic_unit, translated in zip(units, semantic_units, translated_texts):
        if translated is None or not pptx_unit.part_name.startswith("ppt/slides/"):
            continue
        if semantic_unit.constraints.max_chars is None:
            continue
        container = _text_container_for_nodes(pptx_unit.nodes)
        if container is None:
            continue
        key = id(container)
        char_load, line_load = _constraint_load_components(
            str(translated), semantic_unit.constraints
        )
        state = shape_loads.setdefault(
            key,
            {"container": container, "char_load": 0.0, "line_load": 0.0},
        )
        state["char_load"] += char_load
        state["line_load"] += line_load

    autofit_applied = 0
    for state in shape_loads.values():
        load = max(float(state["char_load"]), float(state["line_load"]))
        if load <= 1.0:
            continue
        # If even the minimum allowed font scale cannot compensate for the
        # estimated load, keep the output but surface it as a visual-risk warning.
        if (
            load * (_AUTOFIT_MIN_SCALE ** 2) > 1.0
            and not _uses_shape_autofit(state["container"])
        ):
            visual_metrics.visual_overflow_warnings += 1
        scale = max(_AUTOFIT_MIN_SCALE, min(1.0, 1.0 / math.sqrt(load)))
        if scale >= _AUTOFIT_TRIGGER_SCALE:
            continue
        container = state["container"]
        if _apply_normal_autofit(container, scale):
            autofit_applied += 1
    visual_metrics.pptx_autofit_adjustments = autofit_applied
    if autofit_applied:
        logger.info("PPTX applied bounded autofit to %d slide shape(s)", autofit_applied)

    changed_parts: set[str] = set()
    errors = 0
    for index, (unit, translated) in enumerate(zip(units, translated_texts), start=1):
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
        progress.update("layout", layout_base + index, layout_total)

    patches = {
        name: _serialize_xml(root, original)
        for name, (root, original) in roots.items()
        if name in changed_parts
    }
    if cancel_event is not None and cancel_event.is_set():
        raise CancelledError("Translation cancelled before saving PPTX")
    progress.update("saving", 0, 1)
    _rebuild_package(input_path, output_path, patches)
    progress.complete()
    if metrics_callback is not None:
        metrics_callback(visual_metrics)

    if errors:
        message = f"Translation completed with {errors} failed PPTX text units; originals were preserved"
        if strict_errors:
            raise RuntimeError(message)
        logger.warning(message)
