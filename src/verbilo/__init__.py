try:
    from importlib.metadata import version, PackageNotFoundError
except Exception:
    try:
        from importlib_metadata import version, PackageNotFoundError
    except Exception:
        version = None
        class PackageNotFoundError(Exception):
            pass

# Development fallback value (kept so importing from source still shows something)
__version__ = "0.1.0"

# Prefer installed distribution metadata when available
if version:
    try:
        __version__ = version("verbilo")
    except PackageNotFoundError:
        pass