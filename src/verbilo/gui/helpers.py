import threading
import time
import logging
import re
from pathlib import Path
from typing import Callable, Iterable, Union, Optional, Any
import tkinter as tk
import traceback

from ..main import translate_file
from ..utils import CancelledError

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

def center_window(window, width=None, height=None, parent=None):
    # Center `window` on screen (parent=None) or over `parent`
    window.update_idletasks()

    if width is not None and height is not None:
        # logical -> physical for centering math only
        try:
            sf = window._get_window_scaling()
        except AttributeError:
            sf = 1.0
        sf = sf if sf > 0 else 1.0
        phys_w = round(width * sf)
        phys_h = round(height * sf)
    else:
        phys_w = window.winfo_width()
        phys_h = window.winfo_height()

    if parent is not None:
        parent.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width()  - phys_w) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - phys_h) // 2
    else:
        x = (window.winfo_screenwidth()  - phys_w) // 2
        y = (window.winfo_screenheight() - phys_h) // 2

    x = max(0, x)
    y = max(0, y)

    if width is not None and height is not None:
        window.geometry(f"{width}x{height}+{x}+{y}")
    else:
        # position-only: CTk._apply_geometry_scaling passes +x+y unchanged
        window.geometry(f"+{x}+{y}")


# runs translation in a background thread; call start() to begin, stop() to cancel
class Worker:

    def __init__(self):
        self._thread: Optional[threading.Thread] = None
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
        output_dir: Optional[str],
        translator_name: Optional[str],
        progress_cb: Callable[[str, str, Optional[float]], None],
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


class GuiLoggingHandler(logging.Handler):
    # Logging handler that forwards formatted log records to a GUI log callback.

    def __init__(self, log_cb: Callable, debug_getter: Optional[Callable] = None):
        super().__init__()
        self.log_cb = log_cb
        self._debug_getter = debug_getter or (lambda: False)

    def _sanitize_warning_text(self, text: str) -> str:
        try:
            # If the captured text contains 'UserWarning:'
            if "UserWarning:" in text:
                idx = text.find("UserWarning:")
                after = text[idx + len("UserWarning:"):]
            else:
                after = text

            # Split into lines and pick the first non-empty line
            for line in after.splitlines():
                s = line.strip()
                if s:
                    return s
            return after.strip()
        except Exception:
            return text

    def emit(self, record: logging.LogRecord) -> None:
        try:
            orig_msg = record.getMessage()

            # Skip noisy informational messages about collected cells
            if record.levelno == logging.INFO and re.search(r"collected \d+ translatable string cells", orig_msg, re.IGNORECASE):
                return

            # Determine current debug mode
            debug_enabled = False
            try:
                debug_enabled = bool(self._debug_getter())
            except Exception:
                debug_enabled = False

            if not debug_enabled and record.levelno < logging.WARNING:
                return

            # Sanitize warning messages
            sanitized = orig_msg
            if record.name == "py.warnings" or "UserWarning:" in orig_msg:
                sanitized = self._sanitize_warning_text(orig_msg)

            final = None
            if self.formatter is not None:
                try:
                    formatted = self.format(record)
                    if orig_msg:
                        final = formatted.replace(orig_msg, sanitized, 1)
                    else:
                        final = f"{formatted} {sanitized}"
                except Exception:
                    final = sanitized
            else:
                final = sanitized

            if final:
                try:
                    self.log_cb(final)
                except Exception:
                    # Swallow GUI callback errors to avoid breaking logging
                    pass
        except Exception:
            try:
                self.log_cb(f"Logging error: {record.getMessage()}")
            except Exception:
                pass
