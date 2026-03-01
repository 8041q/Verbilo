import threading
from pathlib import Path
from typing import Callable, Iterable
import tkinter as tk
import traceback

from .main import translate_file

SUPPORTED_EXTS = (".docx", ".pdf", ".xlsx", ".xls")


def list_supported_files(path: str) -> list[str]:
    p = Path(path)
    if p.is_file():
        return [str(p)] if p.suffix.lower() in SUPPORTED_EXTS else []
    if not p.exists():
        return []
    files = []
    for f in p.iterdir():
        if f.is_file() and f.suffix.lower() in SUPPORTED_EXTS:
            files.append(str(f))
    return files


def center_window(window: tk.Tk | tk.Toplevel, width: int, height: int, parent: tk.Widget | None = None) -> None:
    # Center a window on screen or over a parent widget.

    # If `parent` is provided, center `window` over the parent widget.
    # Otherwise center on the primary screen.
    
    window.update_idletasks()
    if parent is not None:
        parent.update_idletasks()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        x = px + (pw - width) // 2
        y = py + (ph - height) // 2
    else:
        sw = window.winfo_screenwidth()
        sh = window.winfo_screenheight()
        x = (sw - width) // 2
        y = (sh - height) // 2
    window.geometry(f"{width}x{height}+{x}+{y}")


class Worker:
    """Background worker that translates a list of files sequentially.

    Call ``start()`` to run in a background thread.  Provide
    ``progress_cb(file, status)`` and ``log_cb(message)`` callbacks.
    Call ``stop()`` to request cancellation.
    """

    def __init__(self):
        self._thread = None
        self._stop = threading.Event()

    def start(
        self,
        files: Iterable[str],
        target_lang: str,
        output_dir: str | None,
        translator_name: str | None,
        progress_cb: Callable[[str, str], None],
        log_cb: Callable[[str], None],
        source_lang: str = "auto",
    ):
        if self._thread and self._thread.is_alive():
            raise RuntimeError("Worker already running")
        if not target_lang or not isinstance(target_lang, str) or not target_lang.strip():
            raise ValueError("target_lang must be a non-empty language code (e.g. 'en')")
        self._stop.clear()
        self._thread = threading.Thread(
            target=self._run,
            args=(list(files), target_lang, output_dir, translator_name, progress_cb, log_cb, source_lang),
            daemon=True,
        )
        self._thread.start()

    def stop(self):
        self._stop.set()

    def _run(self, files, target_lang, output_dir, translator_name, progress_cb, log_cb, source_lang):
        for f in files:
            if self._stop.is_set():
                log_cb("Cancelled by user")
                break
            try:
                progress_cb(f, "started")
                from pathlib import Path
                name = Path(f).name
                log_cb(f"Translating {name} ...")
                out = translate_file(f, target_lang, output_dir, translator_name, source_lang=source_lang)
                if out == "skipped-ocr":
                    progress_cb(f, "finished")
                    log_cb(f"Skipped {name} (scanned/image PDF requiring OCR)")
                else:
                    progress_cb(f, "finished")
                    try:
                        if out:
                            log_cb(f"Finished {name} -> {out}")
                        else:
                            log_cb(f"Finished {name}")
                    except Exception:
                        log_cb(f"Finished {name}")
            except Exception as e:
                progress_cb(f, "error")
                from pathlib import Path
                name = Path(f).name
                tb = traceback.format_exc()
                log_cb(f"Error translating {name}: {e}\n{tb}")
