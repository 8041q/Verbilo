from __future__ import annotations

import sys
from pathlib import Path

import fitz

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from verbilo.converters.pdf_converter import (
    _build_block_html,
    _choose_block_plan,
    _choose_fit_candidate,
    _fit_block,
    _placement_overrides_for_block,
    _preferred_scale_for_block,
)


def _html_for(text: str, *, line_height: float = 1.15) -> tuple[str, str]:
    styles = [{
        "size": 12.0,
        "weight": 400,
        "style": "normal",
        "color": "#000000",
        "align": "left",
    }]
    return _build_block_html(text, styles, line_height=line_height)


def test_fit_block_expands_to_use_available_width() -> None:
    html, css = _html_for("This translated sentence needs more width than the source block had.")
    rect = fitz.Rect(0, 0, 80, 20)

    fit_rect, scale, spare = _fit_block(
        html,
        css,
        rect,
        [],
        [],
        160,
        page_width=220,
        preferred_scale=0.86,
    )

    assert fit_rect.x1 > rect.x1
    assert scale >= 0.86
    assert spare >= 0


def test_fit_block_stays_inside_table_like_boundaries() -> None:
    html, css = _html_for("A longer English label should grow inside the local cell but not cross the borders.")
    rect = fitz.Rect(10, 10, 50, 24)
    obstacles = [
        fitz.Rect(60, 0, 61, 120),
        fitz.Rect(0, 40, 120, 41),
    ]

    fit_rect, _, _ = _fit_block(
        html,
        css,
        rect,
        obstacles,
        [],
        120,
        page_width=120,
        preferred_scale=0.86,
    )

    assert fit_rect.x1 <= 58.0
    assert fit_rect.y1 <= 38.0


def test_fit_block_can_use_left_and_top_slack_inside_bounded_cell() -> None:
    html, css = _html_for(
        "This longer English label needs the surrounding slack inside the bordered cell to stay readable."
    )
    rect = fitz.Rect(38, 24, 56, 30)
    obstacles = [
        fitz.Rect(10, 10, 11, 44),
        fitz.Rect(80, 10, 81, 44),
        fitz.Rect(10, 10, 80, 11),
        fitz.Rect(10, 44, 80, 45),
    ]

    fit_rect, _, spare = _fit_block(
        html,
        css,
        rect,
        obstacles,
        [],
        60,
        page_width=100,
        preferred_scale=0.90,
    )

    assert fit_rect.x0 >= 14.0
    assert fit_rect.y0 >= 14.0
    assert fit_rect.x1 <= 77.0
    assert fit_rect.y1 <= 41.0
    assert fit_rect.x0 < rect.x0 or fit_rect.y0 < rect.y0
    assert spare >= 0


def test_fit_block_keeps_anchored_position_when_anchored_fit_is_readable() -> None:
    html, css = _html_for("Label")
    rect = fitz.Rect(38, 24, 56, 30)
    obstacles = [
        fitz.Rect(10, 10, 11, 44),
        fitz.Rect(80, 10, 81, 44),
        fitz.Rect(10, 10, 80, 11),
        fitz.Rect(10, 44, 80, 45),
    ]

    fit_rect, scale, spare = _fit_block(
        html,
        css,
        rect,
        obstacles,
        [],
        60,
        page_width=100,
        preferred_scale=0.78,
    )

    assert scale >= 0.78
    assert spare >= 0
    assert fit_rect.x0 == rect.x0
    assert fit_rect.y0 == rect.y0


def test_placement_overrides_center_short_label_in_expanded_cell() -> None:
    orig_rect = fitz.Rect(38, 24, 56, 30)
    fit_rect = fitz.Rect(14, 14, 77, 41)
    overrides = _placement_overrides_for_block(
        orig_rect,
        fit_rect,
        [{"align": "left"}],
    )

    assert overrides["align"] == "center"
    assert overrides["padding_top"] > 0


def test_placement_overrides_skip_anchored_short_label() -> None:
    rect = fitz.Rect(38, 24, 56, 30)
    overrides = _placement_overrides_for_block(
        rect,
        rect,
        [{"align": "left"}],
    )

    assert overrides == {}


def test_fit_block_respects_neighboring_siblings() -> None:
    html, css = _html_for("Long text should use the gap before the next block, not overlap it.")
    rect = fitz.Rect(0, 0, 48, 18)
    sibling = fitz.Rect(70, 0, 130, 30)

    fit_rect, _, _ = _fit_block(
        html,
        css,
        rect,
        [],
        [sibling],
        120,
        page_width=180,
        preferred_scale=0.78,
    )

    assert fit_rect.x1 <= 68.0


def test_choose_fit_candidate_prefers_smallest_readable_rect() -> None:
    base_rect = fitz.Rect(0, 0, 100, 50)
    result = _choose_fit_candidate(
        [
            {"rect": fitz.Rect(0, 0, 100, 50), "scale": 0.80, "spare": 5.0},
            {"rect": fitz.Rect(0, 0, 140, 50), "scale": 0.92, "spare": 8.0},
        ],
        base_rect,
        0.78,
    )

    assert result["rect"] == base_rect


def test_choose_block_plan_prefers_smaller_readable_plan() -> None:
    base_rect = fitz.Rect(0, 0, 100, 50)
    result = _choose_block_plan(
        [
            {
                "rect": fitz.Rect(0, 0, 100, 50),
                "scale": 0.80,
                "spare": 5.0,
                "variant_priority": 1,
            },
            {
                "rect": fitz.Rect(0, 0, 140, 50),
                "scale": 0.92,
                "spare": 8.0,
                "variant_priority": 0,
            },
        ],
        base_rect,
        0.78,
    )

    assert result["rect"] == base_rect


def test_preferred_scale_is_higher_for_cjk_source_blocks() -> None:
    assert _preferred_scale_for_block("中文测试内容", "auto") > _preferred_scale_for_block("English source text", "auto")