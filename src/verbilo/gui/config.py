# loads/saves GUI defaults from .verbilo_gui.json in the platform user-config dir

from __future__ import annotations

import json
import os
import logging
import tempfile
from pathlib import Path
from typing import Dict, Any

try:
    from platformdirs import user_config_dir as _user_config_dir
except ImportError:
    # Graceful fallback when platformdirs is not installed (e.g. bare clone without pip install)
    def _user_config_dir(appname: str, **_kw) -> str:  # type: ignore[misc]
        return str(Path.home() / f".{appname.lower()}")

CONFIG_FILENAME = ".verbilo_gui.json"
SENSITIVE_CONFIG_KEYS = frozenset({
    "google_api_key", "google_sa_json", "baidu_appkey", "azure_key", "deepl_api_key",
})


def redact_sensitive_text(text: str) -> str:
    # Remove values for known credentials before text reaches the GUI log
    import re
    redacted = str(text)
    for key in SENSITIVE_CONFIG_KEYS:
        redacted = re.sub(
            rf'(["\']?{re.escape(key)}["\']?\s*[:=]\s*)[^,\s\]\}}]+',
            r'\1***', redacted, flags=re.IGNORECASE,
        )
    return re.sub(r'(DeepL-Auth-Key\s+)[^\s]+', r'\1***', redacted, flags=re.IGNORECASE)

_DEFAULT_CONFIG: Dict[str, Any] = {
    "debug_mode": False,
    "ui_locale": "en",
    "ollama_enabled": False,
    "ollama_model": "qwen3.5:4b",
    "ollama_base_url": "http://127.0.0.1:11434",
    "translation_memory_enabled": False,
    "translation_memory_path": "",
}


def _config_path() -> Path:
    return Path(_user_config_dir("verbilo", appauthor=False)) / CONFIG_FILENAME


def load_config() -> Dict[str, Any]:
    p = _config_path()
    if not p.exists():
        return dict(_DEFAULT_CONFIG)
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return dict(_DEFAULT_CONFIG)
    if not isinstance(data, dict):
        return dict(_DEFAULT_CONFIG)
    cfg = dict(_DEFAULT_CONFIG)
    cfg.update(data)
    if "ollama_enabled" not in data and "pdf_semantic_enabled" in data:
        cfg["ollama_enabled"] = bool(data["pdf_semantic_enabled"])
    return cfg


def save_config(cfg: Dict[str, Any]) -> bool:
    p = _config_path()
    try:
        serialized = dict(cfg)
        if "ollama_enabled" in serialized:
            serialized.pop("pdf_semantic_enabled", None)
        text = json.dumps(serialized, indent=2, ensure_ascii=False)
        # ensure parent folder exists
        parent = p.parent
        if not parent.exists():
            parent.mkdir(parents=True, exist_ok=True)
        # A same-directory temporary file plus replace prevents a crash from
        # leaving the user's credentials/settings as truncated JSON.
        fd, tmp_name = tempfile.mkstemp(prefix=f".{CONFIG_FILENAME}.", suffix=".tmp", dir=parent)
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as fh:
                fh.write(text)
                fh.flush()
                try:
                    os.fsync(fh.fileno())
                except OSError:
                    pass
            os.replace(tmp_name, p)
        finally:
            try:
                Path(tmp_name).unlink(missing_ok=True)
            except OSError:
                pass
        return True
    except Exception:
        # best-effort, don't crash the GUI; log for visibility during development
        try:
            logging.exception("Failed to write GUI config %s", p)
        except Exception:
            pass
        return False
