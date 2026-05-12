from __future__ import annotations

import json
import re
from typing import Literal

from .base import AdvisorBase, AdvisorDecision, ContentHint
from ..translators.http_session import make_session


class OllamaAdvisor(AdvisorBase):
    def __init__(
        self,
        model: Literal["qwen3.5:4b"] = "qwen3.5:4b",
        base_url: str = "http://127.0.0.1:11434",
        timeout: float = 20.0,
        proxies: dict | None = None,
    ) -> None:
        self._model = model
        self._base_url = base_url.rstrip("/")
        self._session = make_session(proxies=proxies, timeout=timeout)

    def classify_block(
        self,
        *,
        text: str,
        source_lang: str,
        target_lang: str,
        capacity_chars: int,
        content_hint: ContentHint,
    ) -> AdvisorDecision:
        system_prompt, user_prompt = self.build_prompt(
            text=text,
            source_lang=source_lang,
            target_lang=target_lang,
            capacity_chars=capacity_chars,
            content_hint=content_hint,
        )
        return self._request_decision(system_prompt, user_prompt)

    def _request_decision(self, system_prompt: str, user_prompt: str) -> AdvisorDecision:
        response = self._session.post(
            f"{self._base_url}/api/chat",
            json={
                "model": self._model,
                "stream": False,
                "think": False,
                "format": "json",
                "options": {
                    "temperature": 0.1,
                },
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
            },
        )
        response.raise_for_status()

        payload = response.json()
        content = payload.get("message", {}).get("content", "")
        content = re.sub(r"<think>.*?</think>", "", content, flags=re.DOTALL).strip()
        if not content:
            raise ValueError("Ollama returned empty content for advisor classification")
        raw = json.loads(content)
        return self.normalize_decision(raw)