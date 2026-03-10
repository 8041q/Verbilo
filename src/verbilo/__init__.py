
# src/verbilo/__init__.py
try:
    from ._version import __version__, __build_date__
except ImportError:
    try:
        from pathlib import Path
        import tomllib  # Python 3.11+

        # Locate pyproject.toml
        root = Path(__file__).resolve().parents[2]
        pyproject_path = root / "pyproject.toml"
        with pyproject_path.open("rb") as f:
            pyproject = tomllib.load(f)

        # Poetry-specific fallback
        poetry = pyproject.get("tool", {}).get("poetry", {})
        __version__ = poetry.get("version", "0.0.0-dev")
        __build_date__ = poetry.get("build_date", "Development-Dynamic")
    except Exception:
        __version__ = "0.0.0-error"
        __build_date__ = "unknown"
