# src/verbilo/__init__.py
from pathlib import Path

_pyproject = Path(__file__).parent.parent.parent / "pyproject.toml"

if _pyproject.exists():
    try:
        import tomllib
    except ImportError:
        import tomli as tomllib

    with open(_pyproject, "rb") as f:
        _data = tomllib.load(f)

    __version__ = _data["tool"]["poetry"]["version"]
    __build_date__ = _data["tool"]["poetry"]["build_date"]

else:
    try:
        from ._version import __version__, __build_date__
    except ImportError:
        __version__ = "0.0.0-error"
        __build_date__ = "unknown"