from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any, Mapping


_METRIC_FIELDS = (
    "units", "unique_units", "cache_hits", "persistent_cache_hits",
    "persistent_cache_writes", "persistent_cache_rejected", "persistent_cache_errors",
    "backend_requests", "retry_items",
    "fallback_items", "unprotected_retries", "failed_items", "terminology_mismatches",
    "source_echo_rejections", "constraint_warnings", "translated_units", "skipped_units",
    "nonlinguistic_skips", "target_language_skips", "non_source_skips", "explicit_skips",
    "detected_units", "tm_hits", "tm_misses", "tm_writes", "tm_rejected", "tm_errors",
)


@dataclass
class TranslationFileReport:
    path: str
    status: str = "pending"
    elapsed_seconds: float = 0.0
    output_path: str | None = None
    metrics: dict[str, int] = field(default_factory=dict)
    error: str | None = None

    def add_metrics(self, metrics: Any) -> None:
        for name in _METRIC_FIELDS:
            value = int(getattr(metrics, name, 0) or 0)
            if value:
                self.metrics[name] = self.metrics.get(name, 0) + value


@dataclass
class TranslationBatchReport:
    files: list[TranslationFileReport] = field(default_factory=list)
    terminology_conflicts: tuple[str, ...] = ()
    terminology_entries: int = 0

    @property
    def totals(self) -> dict[str, int]:
        result: dict[str, int] = {}
        for item in self.files:
            for key, value in item.metrics.items():
                result[key] = result.get(key, 0) + int(value)
        result["files_total"] = len(self.files)
        result["files_finished"] = sum(1 for f in self.files if f.status == "finished")
        result["files_failed"] = sum(1 for f in self.files if f.status == "error")
        result["files_cancelled"] = sum(1 for f in self.files if f.status == "cancelled")
        result["files_skipped"] = sum(1 for f in self.files if f.status == "skipped")
        return result

    @property
    def warning_count(self) -> int:
        t = self.totals
        return (
            t.get("failed_items", 0)
            + t.get("terminology_mismatches", 0)
            + t.get("source_echo_rejections", 0)
            + t.get("constraint_warnings", 0)
            + t.get("persistent_cache_rejected", 0)
            + t.get("persistent_cache_errors", 0)
            + t.get("tm_rejected", 0)
            + t.get("tm_errors", 0)
            + t.get("files_failed", 0)
            + len(self.terminology_conflicts)
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "files": [asdict(item) for item in self.files],
            "totals": self.totals,
            "terminology_conflicts": list(self.terminology_conflicts),
            "terminology_entries": int(self.terminology_entries),
            "warning_count": self.warning_count,
        }
