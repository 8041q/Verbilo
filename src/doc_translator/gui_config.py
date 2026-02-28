# Small config helper for GUI defaults.

# Stores per-project GUI defaults in `.doc_translator_gui.json` in the current working directory.
# Keys: `default_input`, `default_output`.

from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, Any

CONFIG_FILENAME = ".doc_translator_gui.json"


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
        # best-effort, do not crash the GUI
        pass
