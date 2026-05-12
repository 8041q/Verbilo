from __future__ import annotations

from .base import AdvisorBase, AdvisorDecision, ContentHint


class NullAdvisor(AdvisorBase):
    def classify_block(
        self,
        *,
        text: str,
        source_lang: str,
        target_lang: str,
        capacity_chars: int,
        content_hint: ContentHint,
    ) -> AdvisorDecision:
        return {
            "strategy": "free",
            "reason": "null-advisor",
        }