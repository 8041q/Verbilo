from __future__ import annotations

import json
import logging
import threading
from abc import ABC, abstractmethod
from typing import Any, Literal, TypedDict

from ..utils import CancelledError

logger = logging.getLogger(__name__)

BlockStrategy = Literal["literal", "semantic", "free"]
ContentHint = Literal["heading", "body"]
PageBlocks = list[tuple[int, list[dict[str, Any]]]]


class AdvisorDecision(TypedDict):
    strategy: BlockStrategy
    reason: str


class AdvisorBase(ABC):
    def classify_blocks(
        self,
        page_blocks: PageBlocks,
        source_lang: str,
        target_lang: str,
        *,
        cancel_event: threading.Event | None = None,
    ) -> PageBlocks:
        decision_cache: dict[tuple[str, str, str, str, int], AdvisorDecision] = {}
        metrics = {
            "preset_blocks": 0,
            "cache_hits": 0,
            "llm_calls": 0,
            "fallbacks": 0,
        }

        for _, blocks in page_blocks:
            page_font_baseline = self._page_font_baseline(blocks)
            for block in blocks:
                if cancel_event is not None and cancel_event.is_set():
                    raise CancelledError("Translation cancelled")

                capacity_chars = self.estimate_capacity_chars(block)
                content_hint = self.infer_content_hint(block, page_font_baseline)
                preset_reason = str(block.get("advisor_reason", "")).strip()
                preset_strategy = str(block.get("strategy", "")).strip().lower()
                if preset_reason and preset_strategy in {"literal", "semantic", "free"}:
                    normalized = self.normalize_decision(
                        {
                            "strategy": preset_strategy,
                            "reason": preset_reason,
                        }
                    )
                    metrics["preset_blocks"] += 1
                    block["capacity_chars"] = capacity_chars
                    block["content_hint"] = content_hint
                    block["strategy"] = normalized["strategy"]
                    block["advisor_reason"] = normalized["reason"]
                    continue

                cache_key = self._decision_cache_key(
                    text=block.get("text", ""),
                    source_lang=source_lang,
                    target_lang=target_lang,
                    capacity_chars=capacity_chars,
                    content_hint=content_hint,
                )

                try:
                    decision = decision_cache.get(cache_key)
                    if decision is None:
                        metrics["llm_calls"] += 1
                        decision = self.classify_block(
                            text=block.get("text", ""),
                            source_lang=source_lang,
                            target_lang=target_lang,
                            capacity_chars=capacity_chars,
                            content_hint=content_hint,
                        )
                    else:
                        metrics["cache_hits"] += 1
                except CancelledError:
                    raise
                except Exception:
                    logger.exception("Advisor classification failed")
                    metrics["fallbacks"] += 1
                    decision = {
                        "strategy": "free",
                        "reason": "advisor-fallback",
                    }

                normalized = self.normalize_decision(decision)
                decision_cache.setdefault(cache_key, normalized)
                block["capacity_chars"] = capacity_chars
                block["content_hint"] = content_hint
                block["strategy"] = normalized["strategy"]
                block["advisor_reason"] = normalized["reason"]

            logger.info(
                "Advisor metrics: preset=%s cache_hits=%s llm_calls=%s fallbacks=%s",
                metrics["preset_blocks"],
                metrics["cache_hits"],
                metrics["llm_calls"],
                metrics["fallbacks"],
            )
        return page_blocks

    @abstractmethod
    def classify_block(
        self,
        *,
        text: str,
        source_lang: str,
        target_lang: str,
        capacity_chars: int,
        content_hint: ContentHint,
    ) -> AdvisorDecision:
        raise NotImplementedError

    def estimate_capacity_chars(self, block: dict[str, Any]) -> int:
        rect = block.get("rect")
        if rect is None:
            return 1

        font_size = self._block_font_size(block)
        chars_per_line = max(rect.width / max(font_size * 0.55, 1.0), 1.0)
        line_count = max(rect.height / max(font_size * 1.15, 1.0), 1.0)
        estimated = int(chars_per_line * line_count * 0.88)
        return max(estimated, 1)

    def infer_content_hint(
        self,
        block: dict[str, Any],
        page_font_baseline: float,
    ) -> ContentHint:
        text = str(block.get("text", ""))
        visible_text = text.replace("\n", " ").strip()
        line_count = len([line for line in text.splitlines() if line.strip()]) or 1
        block_font_size = self._block_font_size(block)
        align = self._block_alignment(block)

        is_short = len(visible_text) <= 48
        is_heading_sized = block_font_size >= max(page_font_baseline * 1.18, 12.0)
        if line_count <= 2 and is_short and (align == "center" or is_heading_sized):
            return "heading"
        return "body"

    def normalize_decision(self, raw: dict[str, Any] | AdvisorDecision) -> AdvisorDecision:
        strategy = str(raw.get("strategy", "free")).strip().lower()
        if strategy not in {"literal", "semantic", "free"}:
            strategy = "free"

        reason = str(raw.get("reason", "advisor-fallback")).strip()
        if not reason:
            reason = "advisor-fallback"

        return {
            "strategy": strategy,  # type: ignore[return-value]
            "reason": reason[:200],
        }

    def _decision_cache_key(
        self,
        *,
        text: str,
        source_lang: str,
        target_lang: str,
        capacity_chars: int,
        content_hint: ContentHint,
    ) -> tuple[str, str, str, str, int]:
        normalized_text = " ".join(str(text or "").split())
        return (
            normalized_text,
            source_lang.strip().lower(),
            target_lang.strip().lower(),
            content_hint,
            max(int(capacity_chars), 1),
        )

    def build_prompt(
        self,
        *,
        text: str,
        source_lang: str,
        target_lang: str,
        capacity_chars: int,
        content_hint: ContentHint,
    ) -> tuple[str, str]:
        system_prompt = (
            "You are a PDF translation strategy advisor. Do not translate the text. "
            "Choose exactly one strategy for the block: literal, semantic, or free. "
            "literal means preserve wording closely for labels, short headings, numbers, "
            "or wording-sensitive text. semantic means preserve meaning but prefer concise "
            "phrasing for tight layouts. free means preserve meaning naturally when the layout "
            "has room. Chinese-to-English often expands when translated literally, so prefer "
            "semantic for tight zh-to-en blocks unless the wording is obviously sensitive. "
            "Return JSON only with keys strategy and reason."
        )
        user_prompt = json.dumps(
            {
                "source_lang": source_lang,
                "target_lang": target_lang,
                "block_text": text,
                "estimated_capacity_chars": capacity_chars,
                "content_hint": content_hint,
            },
            ensure_ascii=False,
        )
        return system_prompt, user_prompt

    def _page_font_baseline(self, blocks: list[dict[str, Any]]) -> float:
        sizes = sorted(self._block_font_size(block) for block in blocks)
        if not sizes:
            return 11.0
        return sizes[len(sizes) // 2]

    def _block_font_size(self, block: dict[str, Any]) -> float:
        sizes = [
            float(style.get("size", 0.0))
            for style in block.get("line_styles", [])
            if style.get("size")
        ]
        if not sizes:
            return 11.0
        sizes.sort()
        return sizes[len(sizes) // 2]

    def _block_alignment(self, block: dict[str, Any]) -> str:
        align_counts: dict[str, int] = {}
        for style in block.get("line_styles", []):
            align = str(style.get("align", "left")).strip().lower() or "left"
            align_counts[align] = align_counts.get(align, 0) + 1
        if not align_counts:
            return "left"
        return max(align_counts, key=align_counts.get)