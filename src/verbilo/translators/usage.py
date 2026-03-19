# Per-engine monthly character usage tracker
#
# Usage data is stored in ~/.verbilo_usage.json (resolved via platformdirs).
#
# Schema:
#   { "2026-03": { "azure": 1500000, "deepl": 123000 }, "2026-02": {...} }
#
# Old months are retained for history but never contribute to limit checks.

from __future__ import annotations

import json
import logging
import os
import threading
from datetime import datetime
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

try:
    from platformdirs import user_config_dir as _user_config_dir
except ImportError:
    def _user_config_dir(appname: str, **_kw) -> str:   # type: ignore[misc]
        return str(Path.home() / f".{appname.lower()}")

_USAGE_FILENAME = ".verbilo_usage.json"

# Monthly character limits per engine; None = not tracked / unlimited.
ENGINE_LIMITS: dict[str, Optional[int]] = {
    "azure":            2_000_000,  # Azure free tier: 2 M chars/month
    "deepl":              500_000,  # DeepL Free:    500 K chars/month
    "google-cloud":       500_000,  # Basic (v2) free tier: 500 K chars/month
    "google-cloud-v3":        None, # Advanced (v3): pay-per-use, informational only
    "google":                 None, # Rate-limited, not char-limited
    "baidu":               50_000,  # Standard free tier: 50 K chars/month
    "baidu-premium":          None, # Premium: QPS-limited, not char-limited
}


def _usage_path() -> Path:
    return Path(_user_config_dir("verbilo", appauthor=False)) / _USAGE_FILENAME


class UsageTracker:
    # Thread-safe, persistent monthly character-count tracker.

    def __init__(self, path: Optional[Path] = None):
        self._path = path or _usage_path()
        self._lock = threading.RLock()
        self._data: dict[str, dict[str, int]] = {}
        self._load()

    # ── Persistence ──────────────────────────────────────────────────────────

    def _load(self) -> None:
        try:
            if self._path.exists():
                raw = self._path.read_text(encoding="utf-8")
                self._data = json.loads(raw)
        except Exception:
            logger.warning("Could not load usage file at %s; starting fresh.", self._path)
            self._data = {}

    def _save(self) -> None:
        try:
            self._path.parent.mkdir(parents=True, exist_ok=True)
            text = json.dumps(self._data, indent=2, ensure_ascii=False)
            with open(self._path, "w", encoding="utf-8") as fh:
                fh.write(text)
                fh.flush()
                try:
                    os.fsync(fh.fileno())
                except Exception:
                    pass
        except Exception:
            logger.warning("Could not save usage file at %s.", self._path, exc_info=True)

    # ── Month key ─────────────────────────────────────────────────────────────

    @staticmethod
    def _month_key() -> str:
        return datetime.now().strftime("%Y-%m")

    # ── Public API ────────────────────────────────────────────────────────────

    def record(self, engine: str, char_count: int) -> None:
        # Add *char_count* characters to this month's tally for *engine*
        if char_count <= 0:
            return
        with self._lock:
            month = self._month_key()
            bucket = self._data.setdefault(month, {})
            bucket[engine] = bucket.get(engine, 0) + char_count
            self._save()

    def get_usage(self, engine: str) -> int:
        # Return total characters used this month for *engine*
        with self._lock:
            return self._data.get(self._month_key(), {}).get(engine, 0)

    def get_limit(self, engine: str) -> Optional[int]:
        # Return the monthly character limit for *engine*, or ``None`` if unlimited
        return ENGINE_LIMITS.get(engine)

    def get_remaining(self, engine: str) -> Optional[int]:
        # Return remaining characters this month, or ``None`` if engine is unlimited
        limit = self.get_limit(engine)
        if limit is None:
            return None
        return max(0, limit - self.get_usage(engine))

    def check_warning(self, engine: str) -> Optional[str]:
        """Return a warning level string if usage is above a threshold
        Return values
        -------------
        ``"limit"``  — at or above 100 %
        ``"warn"``   — 90 – 100 %
        ``"info"``   — 80 –  90 %
        ``None``     — below 80 % or engine has no limit
        """
        limit = self.get_limit(engine)
        if limit is None or limit <= 0:
            return None
        used = self.get_usage(engine)
        pct = used / limit
        if pct >= 1.0:
            return "limit"
        if pct >= 0.90:
            return "warn"
        if pct >= 0.80:
            return "info"
        return None

    def format_usage(self, engine: str) -> Optional[str]:
        # Return a human-readable usage string or ``None`` for unlimited engines.
        # Example: ``"1.2M / 2M chars (60%)"``
        
        limit = self.get_limit(engine)
        if limit is None:
            return None
        used = self.get_usage(engine)
        pct  = int(used / limit * 100) if limit else 0

        def _fmt(n: int) -> str:
            if n >= 1_000_000:
                return f"{n / 1_000_000:.1f}M"
            if n >= 1_000:
                return f"{n / 1_000:.0f}K"
            return str(n)

        return f"{_fmt(used)} / {_fmt(limit)} chars ({pct}%)"

    def reset(self, engine: Optional[str] = None) -> None:
        # Reset this month's usage - all engines when *engine* is ``None``
        with self._lock:
            month = self._month_key()
            if engine is None:
                self._data[month] = {}
            else:
                self._data.setdefault(month, {}).pop(engine, None)
            self._save()


# ── Module-level singleton ────────────────────────────────────────────────────

_tracker: Optional[UsageTracker] = None
_tracker_lock = threading.Lock()


def get_tracker() -> UsageTracker:
    # Return the global :class:`UsageTracker` instance (created on first call)
    global _tracker
    if _tracker is None:
        with _tracker_lock:
            if _tracker is None:
                _tracker = UsageTracker()
    return _tracker
