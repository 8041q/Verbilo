from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Mapping


@dataclass(frozen=True)
class ProgressUpdate:
    """One measurable progress update for a document translation.

    ``overall_fraction`` is task progress, not an ETA.  Stage progress is based
    on completed document work (analysis items, translation units, write-back
    items and the final save) rather than elapsed time or source file bytes.
    """

    stage: str
    label: str
    completed: int
    total: int
    stage_fraction: float
    overall_fraction: float
    detail: str | None = None


_STAGE_BANDS: Mapping[str, tuple[float, float, str]] = {
    "analyzing": (0.00, 0.10, "Analyzing document"),
    "translating": (0.10, 0.88, "Translating text"),
    "layout": (0.88, 0.97, "Applying translated content"),
    "saving": (0.97, 1.00, "Saving output"),
}


class ProgressReporter:
    """Emit monotonic, stage-aware document progress.

    The bands intentionally reserve most of the bar for translation because it
    is normally the dominant operation, but movement inside each band is driven
    only by real completed work.  Callers may report any integer unit count.
    """

    def __init__(self, callback: Callable[[ProgressUpdate], None] | None = None) -> None:
        self.callback = callback
        self._last_fraction = 0.0
        self._last_stage = ""

    def update(
        self,
        stage: str,
        completed: int,
        total: int,
        *,
        detail: str | None = None,
    ) -> ProgressUpdate | None:
        if self.callback is None:
            return None
        if stage not in _STAGE_BANDS:
            raise ValueError(f"Unknown progress stage: {stage!r}")

        start, end, label = _STAGE_BANDS[stage]
        safe_total = max(int(total), 1)
        safe_completed = min(max(int(completed), 0), safe_total)
        stage_fraction = safe_completed / safe_total
        fraction = start + (end - start) * stage_fraction
        # Never let retries, a recalculated total, or duplicate callbacks make
        # a GUI bar move backwards.
        fraction = max(self._last_fraction, min(1.0, fraction))
        self._last_fraction = fraction
        self._last_stage = stage

        update = ProgressUpdate(
            stage=stage,
            label=label,
            completed=safe_completed,
            total=safe_total,
            stage_fraction=stage_fraction,
            overall_fraction=fraction,
            detail=detail,
        )
        self.callback(update)
        return update

    def complete(self, *, detail: str | None = None) -> ProgressUpdate | None:
        return self.update("saving", 1, 1, detail=detail)
