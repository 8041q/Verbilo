# loads/saves GUI defaults from .verbilo_gui.json in the platform user-config dir

from __future__ import annotations

import json
import os
import logging
from pathlib import Path
from typing import Dict, Any

try:
    from platformdirs import user_config_dir as _user_config_dir
except ImportError:
    # Graceful fallback when platformdirs is not installed (e.g. bare clone without pip install)
    def _user_config_dir(appname: str, **_kw) -> str:  # type: ignore[misc]
        return str(Path.home() / f".{appname.lower()}")

CONFIG_FILENAME = ".verbilo_gui.json"


def _config_path() -> Path:
    return Path(_user_config_dir("verbilo", appauthor=False)) / CONFIG_FILENAME


def load_config() -> Dict[str, Any]:
    p = _config_path()
    if not p.exists():
        # first-run defaults
        return {"debug_mode": False}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_config(cfg: Dict[str, Any]) -> None:
    p = _config_path()
    try:
        text = json.dumps(cfg, indent=2, ensure_ascii=False)
        # ensure parent folder exists
        parent = p.parent
        if not parent.exists():
            parent.mkdir(parents=True, exist_ok=True)
        # write to ensure the file appears on disk
        with open(p, "w", encoding="utf-8") as fh:
            fh.write(text)
            fh.flush()
            try:
                os.fsync(fh.fileno())
            except Exception:
                # fsync may not be available on some platforms or filesystems; ignore
                pass
    except Exception:
        # best-effort, don't crash the GUI; log for visibility during development
        try:
            logging.exception("Failed to write GUI config %s", p)
        except Exception:
            pass
