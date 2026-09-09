import os
from pathlib import Path


def _translated_name(path: Path) -> str:
    return path.stem + ".translated" + path.suffix


def _available_path(directory: Path, filename: str) -> Path:
    # Return a non-existent filename in *directory* without overwriting output
    candidate = directory / filename
    if not candidate.exists():
        return candidate
    stem = Path(filename).stem
    suffix = Path(filename).suffix
    index = 2
    while True:
        candidate = directory / f"{stem} ({index}){suffix}"
        if not candidate.exists():
            return candidate
        index += 1


def resolve_output_path(input_path: Path | str, out_arg: str | None = None) -> str:
    # None → .translated next to input; directory → collision-safe translated name;
    # explicit filename → use as-is (overwrite policy belongs to the caller).
    p = Path(input_path).resolve()
    if out_arg is None:
        return str(_available_path(p.parent, _translated_name(p)))

    out_p = Path(out_arg)
    # If out_arg explicitly ends with a separator, treat as directory
    if str(out_arg).endswith(os.path.sep) or str(out_arg).endswith("/") or str(out_arg).endswith("\\"):
        out_p.mkdir(parents=True, exist_ok=True)
        return str(_available_path(out_p, _translated_name(p)))

    if out_p.exists() and out_p.is_dir():
        return str(_available_path(out_p, _translated_name(p)))

    # Otherwise treat as file path
    return str(out_p)


def format_bytes(n: int) -> str:
    # Format bytes as human-readable string (B, KB, MB, GB).
    try:
        n = int(n or 0)
    except Exception:
        return "0 B"
    if n < 1024:
        return f"{n} B"
    for unit in ("KB", "MB", "GB", "TB"):
        n /= 1024.0
        if n < 1024.0:
            return f"{n:.1f} {unit}"
    return f"{n:.1f} PB"
