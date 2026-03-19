# Persistent SQLite-backed translation cache
#
# DB location : resolved via platformdirs  (~/.verbilo_cache.db)
# Table key   : (engine, source_text, target_lang)
# Eviction    : LRU, keeps at most *max_entries* rows (default 500 000)
# Concurrency : WAL mode + per-thread connections (safe for multi-threaded use)

from __future__ import annotations

import logging
import sqlite3
import threading
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

try:
    from platformdirs import user_config_dir as _user_config_dir
except ImportError:
    def _user_config_dir(appname: str, **_kw) -> str:   # type: ignore[misc]
        return str(Path.home() / f".{appname.lower()}")

_CACHE_FILENAME = ".verbilo_cache.db"
_MAX_ENTRIES    = 500_000     # LRU eviction threshold


def _cache_path() -> Path:
    return Path(_user_config_dir("verbilo", appauthor=False)) / _CACHE_FILENAME


_DDL = """
CREATE TABLE IF NOT EXISTS translations (
    engine          TEXT    NOT NULL,
    source_text     TEXT    NOT NULL,
    target_lang     TEXT    NOT NULL,
    translated_text TEXT    NOT NULL,
    created_at      INTEGER NOT NULL DEFAULT (strftime('%s','now')),
    PRIMARY KEY (engine, source_text, target_lang)
);
CREATE INDEX IF NOT EXISTS idx_tl_created ON translations (created_at);
"""


class TranslationCache:
    """Thread-safe SQLite translation cache.

    Two-level usage pattern (not enforced here — translators do it):
    L1 = in-memory ``dict`` per wrapper instance (fast, session-scoped)
    L2 = this SQLite cache (persistent, shared across sessions)

    Instances are normally obtained via :func:`get_cache`.
    """

    def __init__(
        self,
        db_path: Optional[Path] = None,
        max_entries: int = _MAX_ENTRIES,
    ):
        self._db_path    = db_path or _cache_path()
        self._max_entries = max_entries
        self._local = threading.local()          # per-thread connection
        self._write_lock = threading.Lock()      # serialise writes / eviction
        self._init_db()

    # ── Connection management ─────────────────────────────────────────────────

    def _get_conn(self) -> sqlite3.Connection:
        """Return (creating if necessary) the per-thread SQLite connection."""
        conn: Optional[sqlite3.Connection] = getattr(self._local, "conn", None)
        if conn is None:
            self._db_path.parent.mkdir(parents=True, exist_ok=True)
            conn = sqlite3.connect(
                str(self._db_path),
                check_same_thread=False,
                isolation_level=None,   # autocommit
            )
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("PRAGMA synchronous=NORMAL")
            self._local.conn = conn
        return conn

    def _init_db(self) -> None:
        try:
            db = self._get_conn()
            db.executescript(_DDL)
        except Exception:
            logger.warning("Could not initialise translation cache DB.", exc_info=True)

    # ── Read API ──────────────────────────────────────────────────────────────

    def get(self, engine: str, text: str, target_lang: str) -> Optional[str]:
        """Return the cached translation or ``None`` on a miss."""
        if not text:
            return None
        try:
            row = self._get_conn().execute(
                "SELECT translated_text FROM translations "
                "WHERE engine=? AND source_text=? AND target_lang=?",
                (engine, text, target_lang),
            ).fetchone()
            return row[0] if row else None
        except Exception:
            logger.debug("Cache.get failed", exc_info=True)
            return None

    def get_batch(
        self, engine: str, texts: list[str], target_lang: str
    ) -> dict[str, str]:
        """Return ``{source_text: translated_text}`` for every cached hit in *texts*."""
        if not texts:
            return {}
        try:
            placeholders = ",".join("?" * len(texts))
            rows = self._get_conn().execute(
                f"SELECT source_text, translated_text FROM translations "
                f"WHERE engine=? AND target_lang=? AND source_text IN ({placeholders})",
                (engine, target_lang, *texts),
            ).fetchall()
            return {src: tgt for src, tgt in rows}
        except Exception:
            logger.debug("Cache.get_batch failed", exc_info=True)
            return {}

    # ── Write API ─────────────────────────────────────────────────────────────

    def put(
        self, engine: str, text: str, target_lang: str, translated_text: str
    ) -> None:
        """Store a single translation; silently overwrites existing entries."""
        if not text or not translated_text:
            return
        try:
            with self._write_lock:
                db = self._get_conn()
                db.execute(
                    "INSERT OR REPLACE INTO translations "
                    "(engine, source_text, target_lang, translated_text, created_at) "
                    "VALUES (?, ?, ?, ?, strftime('%s','now'))",
                    (engine, text, target_lang, translated_text),
                )
                self._maybe_evict(db)
        except Exception:
            logger.debug("Cache.put failed", exc_info=True)

    def put_batch(
        self,
        engine: str,
        pairs: list[tuple[str, str]],   # (source_text, translated_text)
        target_lang: str,
    ) -> None:
        """Bulk-insert translations; silently overwrites existing entries."""
        valid = [(src, tgt) for src, tgt in pairs if src and tgt]
        if not valid:
            return
        try:
            with self._write_lock:
                db = self._get_conn()
                db.executemany(
                    "INSERT OR REPLACE INTO translations "
                    "(engine, source_text, target_lang, translated_text, created_at) "
                    "VALUES (?, ?, ?, ?, strftime('%s','now'))",
                    [(engine, src, target_lang, tgt) for src, tgt in valid],
                )
                self._maybe_evict(db)
        except Exception:
            logger.debug("Cache.put_batch failed", exc_info=True)

    # ── Maintenance ───────────────────────────────────────────────────────────

    def _maybe_evict(self, db: sqlite3.Connection) -> None:
        """Remove the oldest rows if the cache exceeds *max_entries*."""
        try:
            count = db.execute("SELECT COUNT(*) FROM translations").fetchone()[0]
            if count > self._max_entries:
                excess = count - self._max_entries
                db.execute(
                    "DELETE FROM translations WHERE rowid IN "
                    "(SELECT rowid FROM translations ORDER BY created_at ASC LIMIT ?)",
                    (excess,),
                )
        except Exception:
            logger.debug("Cache eviction failed", exc_info=True)

    def clear(self, engine: Optional[str] = None) -> None:
        """Delete all cache entries, or only entries for *engine* if specified."""
        try:
            with self._write_lock:
                db = self._get_conn()
                if engine is None:
                    db.execute("DELETE FROM translations")
                else:
                    db.execute("DELETE FROM translations WHERE engine=?", (engine,))
                # Try to reclaim file space after deleting rows
                try:
                    # checkpoint WAL and then VACUUM to shrink file on disk
                    db.execute("PRAGMA wal_checkpoint(FULL)")
                    db.execute("VACUUM")
                except Exception:
                    # best-effort; ignore failures
                    pass
        except Exception:
            logger.debug("Cache.clear failed", exc_info=True)

    def size(self, engine: Optional[str] = None) -> int:
        """Return the number of cached translations (optionally filtered to *engine*)."""
        try:
            if engine is None:
                row = self._get_conn().execute(
                    "SELECT COUNT(*) FROM translations"
                ).fetchone()
            else:
                row = self._get_conn().execute(
                    "SELECT COUNT(*) FROM translations WHERE engine=?", (engine,)
                ).fetchone()
            return row[0] if row else 0
        except Exception:
            return 0

    def disk_usage_bytes(self) -> int:
        """Return the approximate on-disk size, uses SQLite PRAGMA to compute `page_count * page_size`. Falls back to the file size on disk if PRAGMA fails."""
        try:
            db = self._get_conn()
            pc_row = db.execute("PRAGMA page_count").fetchone()
            ps_row = db.execute("PRAGMA page_size").fetchone()
            if pc_row and ps_row:
                page_count = int(pc_row[0] or 0)
                page_size = int(ps_row[0] or 0)
                if page_count and page_size:
                    return page_count * page_size
        except Exception:
            logger.debug("Could not read PRAGMA page_count/page_size", exc_info=True)
        try:
            if self._db_path.exists():
                return int(self._db_path.stat().st_size)
        except Exception:
            logger.debug("Could not stat cache DB file", exc_info=True)
        return 0


# ── Module-level singleton ────────────────────────────────────────────────────

_cache: Optional[TranslationCache] = None
_cache_lock = threading.Lock()


def get_cache() -> TranslationCache:
    """Return the global :class:`TranslationCache` instance (created on first call)."""
    global _cache
    if _cache is None:
        with _cache_lock:
            if _cache is None:
                _cache = TranslationCache()
    return _cache
