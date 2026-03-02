# loads/saves GUI defaults from .verbilo_gui.json in cwd

from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, Any

CONFIG_FILENAME = ".verbilo_gui.json"


def _config_path() -> Path:
    return Path.cwd() / CONFIG_FILENAME


def load_config() -> Dict[str, Any]:
    p = _config_path()
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_config(cfg: Dict[str, Any]) -> None:
    p = _config_path()
    try:
        p.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")
    except Exception:
        # best-effort, don't crash the GUI
        pass
