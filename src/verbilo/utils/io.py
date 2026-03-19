from pathlib import Path
import os


def resolve_output_path(input_path: Path | str, out_arg: str | None = None) -> str:
    # None → .translated next to input; dir → original name inside it; else use as-is
    p = Path(input_path).resolve()
    if out_arg is None:
        return str(p.with_name(p.stem + ".translated" + p.suffix))

    out_p = Path(out_arg)
    # If out_arg explicitly ends with a separator, treat as directory
    if str(out_arg).endswith(os.path.sep) or str(out_arg).endswith("/") or str(out_arg).endswith("\\"):
        out_p.mkdir(parents=True, exist_ok=True)
        return str(out_p / p.name)

    if out_p.exists() and out_p.is_dir():
        return str(out_p / p.name)

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
