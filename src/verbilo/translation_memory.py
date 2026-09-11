from __future__ import annotations

import sqlite3
import threading
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable, Sequence

_SCHEMA_LOCK = threading.Lock()


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def default_translation_memory_path() -> Path:
    return Path.home() / ".verbilo" / "translation_memory.sqlite3"


@dataclass(frozen=True)
class TranslationMemoryEntry:
    key: str
    identity_json: str
    source_text: str
    source_lang: str
    target_lang: str
    role: str
    mode: str
    translation: str
    engine: str = ""
    created_at: str | None = None
    updated_at: str | None = None
    last_used_at: str | None = None
    use_count: int = 0


class TranslationMemory:
    """Persistent exact-match translation memory backed by SQLite.

    Connections are short-lived and WAL mode is used so separate translation
    workers/processes can safely share the same database. The memory stores only
    validated raw translations; protected placeholder restoration remains the
    responsibility of ``TranslationUnit``.
    """

    def __init__(self, path: str | Path | None = None, *, timeout: float = 10.0) -> None:
        self.path = Path(path or default_translation_memory_path()).expanduser()
        self.timeout = float(timeout)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._ensure_schema()

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.path, timeout=self.timeout)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA busy_timeout = 10000")
        connection.execute("PRAGMA foreign_keys = ON")
        return connection

    def _ensure_schema(self) -> None:
        with _SCHEMA_LOCK:
            with self._connect() as connection:
                connection.execute("PRAGMA journal_mode = WAL")
                connection.execute(
                    """
                    CREATE TABLE IF NOT EXISTS translations (
                        key TEXT PRIMARY KEY,
                        identity_json TEXT NOT NULL,
                        source_text TEXT NOT NULL,
                        source_lang TEXT NOT NULL,
                        target_lang TEXT NOT NULL,
                        role TEXT NOT NULL,
                        mode TEXT NOT NULL,
                        translation TEXT NOT NULL,
                        engine TEXT NOT NULL DEFAULT '',
                        created_at TEXT NOT NULL,
                        updated_at TEXT NOT NULL,
                        last_used_at TEXT,
                        use_count INTEGER NOT NULL DEFAULT 0
                    )
                    """
                )
                connection.execute(
                    "CREATE INDEX IF NOT EXISTS idx_tm_langs ON translations(source_lang, target_lang)"
                )
                connection.execute("PRAGMA user_version = 1")

    @staticmethod
    def _entry_from_row(row: sqlite3.Row) -> TranslationMemoryEntry:
        return TranslationMemoryEntry(
            key=str(row["key"]),
            identity_json=str(row["identity_json"]),
            source_text=str(row["source_text"]),
            source_lang=str(row["source_lang"]),
            target_lang=str(row["target_lang"]),
            role=str(row["role"]),
            mode=str(row["mode"]),
            translation=str(row["translation"]),
            engine=str(row["engine"] or ""),
            created_at=str(row["created_at"]),
            updated_at=str(row["updated_at"]),
            last_used_at=(str(row["last_used_at"]) if row["last_used_at"] else None),
            use_count=int(row["use_count"] or 0),
        )

    def get(self, key: str, *, touch: bool = True) -> TranslationMemoryEntry | None:
        found = self.get_many([key], touch=touch)
        return found.get(key)

    def get_many(
        self,
        keys: Sequence[str] | Iterable[str],
        *,
        touch: bool = True,
    ) -> dict[str, TranslationMemoryEntry]:
        unique_keys = list(dict.fromkeys(str(key) for key in keys if str(key)))
        if not unique_keys:
            return {}
        found: dict[str, TranslationMemoryEntry] = {}
        now = _utc_now()
        with self._connect() as connection:
            for offset in range(0, len(unique_keys), 500):
                chunk = unique_keys[offset : offset + 500]
                placeholders = ",".join("?" for _ in chunk)
                rows = connection.execute(
                    f"SELECT * FROM translations WHERE key IN ({placeholders})",
                    chunk,
                ).fetchall()
                for row in rows:
                    entry = self._entry_from_row(row)
                    found[entry.key] = entry
            if touch and found:
                for offset in range(0, len(found), 500):
                    chunk = list(found)[offset : offset + 500]
                    placeholders = ",".join("?" for _ in chunk)
                    connection.execute(
                        f"UPDATE translations SET use_count = use_count + 1, last_used_at = ? "
                        f"WHERE key IN ({placeholders})",
                        [now, *chunk],
                    )
        return found

    def put(self, entry: TranslationMemoryEntry) -> None:
        self.put_many([entry])

    def put_many(self, entries: Sequence[TranslationMemoryEntry] | Iterable[TranslationMemoryEntry]) -> None:
        items = list(entries)
        if not items:
            return
        now = _utc_now()
        rows = [
            (
                entry.key,
                entry.identity_json,
                entry.source_text,
                entry.source_lang,
                entry.target_lang,
                entry.role,
                entry.mode,
                entry.translation,
                entry.engine or "",
                entry.created_at or now,
                now,
                entry.last_used_at,
                int(entry.use_count or 0),
            )
            for entry in items
        ]
        with self._connect() as connection:
            connection.executemany(
                """
                INSERT INTO translations (
                    key, identity_json, source_text, source_lang, target_lang,
                    role, mode, translation, engine, created_at, updated_at,
                    last_used_at, use_count
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(key) DO UPDATE SET
                    identity_json = excluded.identity_json,
                    source_text = excluded.source_text,
                    source_lang = excluded.source_lang,
                    target_lang = excluded.target_lang,
                    role = excluded.role,
                    mode = excluded.mode,
                    translation = excluded.translation,
                    engine = excluded.engine,
                    updated_at = excluded.updated_at
                """,
                rows,
            )

    def delete_many(self, keys: Sequence[str] | Iterable[str]) -> None:
        unique_keys = list(dict.fromkeys(str(key) for key in keys if str(key)))
        if not unique_keys:
            return
        with self._connect() as connection:
            for offset in range(0, len(unique_keys), 500):
                chunk = unique_keys[offset : offset + 500]
                placeholders = ",".join("?" for _ in chunk)
                connection.execute(
                    f"DELETE FROM translations WHERE key IN ({placeholders})",
                    chunk,
                )

    def count(self) -> int:
        with self._connect() as connection:
            row = connection.execute("SELECT COUNT(*) FROM translations").fetchone()
        return int(row[0] if row else 0)

    def clear(self) -> None:
        with self._connect() as connection:
            connection.execute("DELETE FROM translations")
