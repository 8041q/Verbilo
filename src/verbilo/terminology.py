from __future__ import annotations

import csv
import json
import os
import tempfile
import threading
import uuid
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable, Mapping

try:
    from platformdirs import user_config_dir as _user_config_dir
except ImportError:  # pragma: no cover - fallback for bare installs
    def _user_config_dir(appname: str, **_kw) -> str:
        return str(Path.home() / f".{appname.lower()}")

_SCHEMA_VERSION = 1
_LOCK = threading.RLock()


def default_terminology_path() -> Path:
    return Path(_user_config_dir("verbilo", appauthor=False)) / "terminology.json"


def _norm_lang(value: str | None) -> str:
    return str(value or "").strip().replace("_", "-").lower()


def _lang_base(value: str | None) -> str:
    return _norm_lang(value).split("-", 1)[0]


def _norm_term(value: str | None) -> str:
    return " ".join(str(value or "").strip().split())


@dataclass(frozen=True)
class TerminologyEntry:
    id: str
    source_lang: str
    target_lang: str
    source_text: str
    target_text: str
    enabled: bool = True

    @classmethod
    def create(
        cls,
        source_lang: str,
        target_lang: str,
        source_text: str,
        target_text: str,
        *,
        enabled: bool = True,
        entry_id: str | None = None,
    ) -> "TerminologyEntry":
        source_lang = _norm_lang(source_lang)
        target_lang = _norm_lang(target_lang)
        source_text = _norm_term(source_text)
        target_text = _norm_term(target_text)
        if not source_lang or source_lang == "auto":
            raise ValueError("source language must be an explicit language code")
        if not target_lang or target_lang == "auto":
            raise ValueError("target language must be an explicit language code")
        if not source_text:
            raise ValueError("source term cannot be empty")
        if not target_text:
            raise ValueError("target term cannot be empty")
        return cls(
            id=str(entry_id or uuid.uuid4().hex),
            source_lang=source_lang,
            target_lang=target_lang,
            source_text=source_text,
            target_text=target_text,
            enabled=bool(enabled),
        )


@dataclass(frozen=True)
class TerminologySnapshot:
    mapping: Mapping[str, str]
    conflicts: tuple[str, ...] = ()
    matched_entries: int = 0


class TerminologyStore:
    """Atomic, versioned terminology store used by the GUI and batch worker.

    The store is deliberately independent from GUI configuration so a bad settings
    write cannot destroy terminology data. Jobs receive a plain mapping snapshot at
    start; editing the store while a job runs therefore only affects the next job.
    """

    def __init__(self, path: str | os.PathLike[str] | None = None) -> None:
        self.path = Path(path or default_terminology_path()).expanduser()
        self.last_error: str | None = None

    def load(self) -> list[TerminologyEntry]:
        self.last_error = None
        if not self.path.exists():
            return []
        try:
            payload = json.loads(self.path.read_text(encoding="utf-8-sig"))
            if isinstance(payload, list):  # tolerate early/unversioned exports
                raw_entries = payload
            elif isinstance(payload, dict):
                raw_entries = payload.get("entries", [])
            else:
                raise ValueError("terminology file root must be an object or list")
            if not isinstance(raw_entries, list):
                raise ValueError("terminology entries must be a list")
            entries: list[TerminologyEntry] = []
            seen_ids: set[str] = set()
            for index, item in enumerate(raw_entries, start=1):
                if not isinstance(item, dict):
                    raise ValueError(f"entry {index} is not an object")
                entry = TerminologyEntry.create(
                    item.get("source_lang", ""),
                    item.get("target_lang", ""),
                    item.get("source_text", item.get("source", "")),
                    item.get("target_text", item.get("target", "")),
                    enabled=item.get("enabled", True),
                    entry_id=str(item.get("id") or uuid.uuid4().hex),
                )
                if entry.id in seen_ids:
                    entry = TerminologyEntry.create(
                        entry.source_lang, entry.target_lang, entry.source_text,
                        entry.target_text, enabled=entry.enabled,
                    )
                seen_ids.add(entry.id)
                entries.append(entry)
            return entries
        except Exception as exc:
            self.last_error = str(exc)
            return []

    def save(self, entries: Iterable[TerminologyEntry]) -> None:
        items = list(entries)
        # Validate everything before touching disk.
        validated = [
            TerminologyEntry.create(
                entry.source_lang, entry.target_lang, entry.source_text,
                entry.target_text, enabled=entry.enabled, entry_id=entry.id,
            )
            for entry in items
        ]
        payload = {
            "version": _SCHEMA_VERSION,
            "entries": [asdict(entry) for entry in validated],
        }
        self.path.parent.mkdir(parents=True, exist_ok=True)
        text = json.dumps(payload, ensure_ascii=False, indent=2) + "\n"
        with _LOCK:
            fd, tmp_name = tempfile.mkstemp(
                prefix=f".{self.path.name}.", suffix=".tmp", dir=self.path.parent
            )
            try:
                with os.fdopen(fd, "w", encoding="utf-8", newline="") as handle:
                    handle.write(text)
                    handle.flush()
                    try:
                        os.fsync(handle.fileno())
                    except OSError:
                        pass
                os.replace(tmp_name, self.path)
            finally:
                try:
                    Path(tmp_name).unlink(missing_ok=True)
                except OSError:
                    pass

    def snapshot(self, source_lang: str, target_lang: str) -> TerminologySnapshot:
        entries = [entry for entry in self.load() if entry.enabled]
        target = _norm_lang(target_lang)
        source = _norm_lang(source_lang)
        target_base = _lang_base(target)

        target_candidates: list[tuple[TerminologyEntry, int]] = []
        for entry in entries:
            entry_target = _norm_lang(entry.target_lang)
            if entry_target == target:
                target_candidates.append((entry, 2))
            elif _lang_base(entry_target) == target_base and "-" not in entry_target:
                target_candidates.append((entry, 1))

        if source and source != "auto":
            source_base = _lang_base(source)
            # One winner per source term. Exact source locale and exact target
            # locale outrank generic base-language entries deterministically.
            winners: dict[str, tuple[tuple[int, int], TerminologyEntry]] = {}
            order: list[str] = []
            for entry, target_score in target_candidates:
                entry_source = _norm_lang(entry.source_lang)
                if entry_source == source:
                    source_score = 2
                elif _lang_base(entry_source) == source_base and "-" not in entry_source:
                    source_score = 1
                else:
                    continue
                key = entry.source_text.casefold()
                if key not in winners:
                    order.append(key)
                score = (source_score, target_score)
                if key not in winners or score >= winners[key][0]:
                    winners[key] = (score, entry)
            mapping = {winners[key][1].source_text: winners[key][1].target_text for key in order}
            return TerminologySnapshot(mapping=mapping, matched_entries=len(winners))

        # Auto-source: first choose the most specific target-locale rule within
        # each source language, then compare across source languages. This avoids
        # a generic-target rule falsely conflicting with a regional override.
        per_language: dict[tuple[str, str], tuple[int, TerminologyEntry]] = {}
        source_order: list[str] = []
        for entry, target_score in target_candidates:
            term_key = entry.source_text.casefold()
            source_key = _norm_lang(entry.source_lang)
            pair_key = (term_key, source_key)
            if term_key not in source_order:
                source_order.append(term_key)
            if pair_key not in per_language or target_score >= per_language[pair_key][0]:
                per_language[pair_key] = (target_score, entry)

        mapping: dict[str, str] = {}
        conflicts: list[str] = []
        matched = 0
        for term_key in source_order:
            chosen = [value[1] for key, value in per_language.items() if key[0] == term_key]
            targets = {entry.target_text.casefold() for entry in chosen}
            if len(targets) > 1:
                conflicts.append(chosen[-1].source_text)
                continue
            if chosen:
                mapping[chosen[-1].source_text] = chosen[-1].target_text
                matched += len(chosen)
        return TerminologySnapshot(mapping=mapping, conflicts=tuple(conflicts), matched_entries=matched)

    def import_file(self, path: str | os.PathLike[str], *, replace: bool = False) -> tuple[int, int]:
        import_path = Path(path)
        suffix = import_path.suffix.lower()
        if suffix == ".csv":
            imported = self._read_csv(import_path)
        elif suffix == ".json":
            other = TerminologyStore(import_path)
            imported = other.load()
            if other.last_error:
                raise ValueError(other.last_error)
        else:
            raise ValueError("terminology import must be a .csv or .json file")

        current = [] if replace else self.load()
        by_key: dict[tuple[str, str, str], TerminologyEntry] = {
            (_norm_lang(e.source_lang), _norm_lang(e.target_lang), e.source_text.casefold()): e
            for e in current
        }
        added = 0
        updated = 0
        for entry in imported:
            key = (_norm_lang(entry.source_lang), _norm_lang(entry.target_lang), entry.source_text.casefold())
            if key in by_key:
                old = by_key[key]
                entry = TerminologyEntry.create(
                    entry.source_lang, entry.target_lang, entry.source_text, entry.target_text,
                    enabled=entry.enabled, entry_id=old.id,
                )
                updated += 1
            else:
                added += 1
            by_key[key] = entry
        self.save(by_key.values())
        return added, updated

    @staticmethod
    def _read_csv(path: Path) -> list[TerminologyEntry]:
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            fields = {str(name or "").strip().lower() for name in (reader.fieldnames or [])}
            required = {"source_lang", "target_lang"}
            if not required.issubset(fields):
                raise ValueError("CSV requires source_lang and target_lang columns")
            entries: list[TerminologyEntry] = []
            for row_no, row in enumerate(reader, start=2):
                source_text = row.get("source_text") or row.get("source") or ""
                target_text = row.get("target_text") or row.get("target") or ""
                raw_enabled = str(row.get("enabled", "true")).strip().lower()
                enabled = raw_enabled not in {"0", "false", "no", "off"}
                try:
                    entries.append(TerminologyEntry.create(
                        row.get("source_lang", ""), row.get("target_lang", ""),
                        source_text, target_text, enabled=enabled,
                    ))
                except ValueError as exc:
                    raise ValueError(f"CSV row {row_no}: {exc}") from exc
            return entries

    def export_file(self, path: str | os.PathLike[str]) -> int:
        output = Path(path)
        entries = self.load()
        output.parent.mkdir(parents=True, exist_ok=True)
        if output.suffix.lower() == ".csv":
            fd, tmp_name = tempfile.mkstemp(prefix=f".{output.name}.", suffix=".tmp", dir=output.parent)
            try:
                with os.fdopen(fd, "w", encoding="utf-8-sig", newline="") as handle:
                    writer = csv.DictWriter(handle, fieldnames=[
                        "source_lang", "target_lang", "source_text", "target_text", "enabled"
                    ])
                    writer.writeheader()
                    for entry in entries:
                        writer.writerow({
                            "source_lang": entry.source_lang,
                            "target_lang": entry.target_lang,
                            "source_text": entry.source_text,
                            "target_text": entry.target_text,
                            "enabled": "true" if entry.enabled else "false",
                        })
                os.replace(tmp_name, output)
            finally:
                Path(tmp_name).unlink(missing_ok=True)
        elif output.suffix.lower() == ".json":
            payload = {"version": _SCHEMA_VERSION, "entries": [asdict(e) for e in entries]}
            fd, tmp_name = tempfile.mkstemp(prefix=f".{output.name}.", suffix=".tmp", dir=output.parent)
            try:
                with os.fdopen(fd, "w", encoding="utf-8", newline="") as handle:
                    handle.write(json.dumps(payload, ensure_ascii=False, indent=2) + "\n")
                    handle.flush()
                    try:
                        os.fsync(handle.fileno())
                    except OSError:
                        pass
                os.replace(tmp_name, output)
            finally:
                Path(tmp_name).unlink(missing_ok=True)
        else:
            raise ValueError("terminology export must end in .csv or .json")
        return len(entries)
