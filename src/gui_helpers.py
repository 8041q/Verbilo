import threading
import time
from pathlib import Path
from typing import Callable, Iterable
import tkinter as tk
import traceback

from .main import translate_file
from .utils import CancelledError

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


def center_window(window, width, height=None, parent=None):
    window.update_idletasks()
    if height is None:
        height = window.winfo_reqheight()
    if parent is not None:
        parent.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - width) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - height) // 2
    else:
        x = (window.winfo_screenwidth() - width) // 2
        y = (window.winfo_screenheight() - height) // 2
    window.geometry(f"{width}x{height}+{x}+{y}")


# runs translation in a background thread; call start() to begin, stop() to cancel
class Worker:

    def __init__(self):
        self._thread: threading.Thread | None = None
        self._stop = threading.Event()

    @property
    def cancelled(self) -> bool:
        return self._stop.is_set()

    @property
    def alive(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def start(
        self,
        files: Iterable[str],
        target_lang: str,
        output_dir: str | None,
        translator_name: str | None,
        progress_cb: Callable[[str, str, float | None], None],
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
            name = Path(f).name
            t0 = time.perf_counter()
            try:
                progress_cb(f, "started", None)
                log_cb(f"Translating {name} ...")
                out = translate_file(
                    f, target_lang, output_dir, translator_name,
                    source_lang=source_lang,
                    cancel_event=self._stop,
                )
                elapsed = time.perf_counter() - t0

                # Check cancellation right after translate_file returns
                if self._stop.is_set():
                    progress_cb(f, "cancelled", None)
                    log_cb(f"Cancelled during {name}")
                    break

                if out == "skipped-ocr":
                    progress_cb(f, "finished", elapsed)
                    log_cb(f"Skipped {name} (scanned/image PDF requiring OCR)")
                else:
                    progress_cb(f, "finished", elapsed)
                    try:
                        if out:
                            log_cb(f"Finished {name} -> {out}")
                        else:
                            log_cb(f"Finished {name}")
                    except Exception:
                        log_cb(f"Finished {name}")
            except CancelledError:
                progress_cb(f, "cancelled", None)
                log_cb(f"Cancelled during {name}")
                break
            except Exception as e:
                elapsed = time.perf_counter() - t0
                progress_cb(f, "error", elapsed)
                tb = traceback.format_exc()
                log_cb(f"Error translating {name}: {e}\n{tb}")

        # Signal that the worker loop has exited
        log_cb("__worker_done__")
