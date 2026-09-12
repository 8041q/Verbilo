from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Callable, Iterable

logger = logging.getLogger(__name__)


@dataclass
class FileDropRegistration:
    backend: str
    _cleanup: Callable[[], None] | None = None

    def close(self) -> None:
        if self._cleanup is not None:
            try:
                self._cleanup()
            except Exception:
                logger.debug("Failed to clean up file-drop registration", exc_info=True)
            self._cleanup = None


def _install_tkinterdnd(root, widget, on_files) -> FileDropRegistration | None:
    """Register a Tk-aware file-drop target when tkinterdnd2 is installed.

    Do not install a native Windows WNDPROC hook here. Tk/CustomTkinter owns the
    window procedure and Python callbacks invoked directly from a foreign/native
    callback can race Tk's interpreter/GIL state. tkinterdnd2 keeps delivery
    inside Tk's event machinery and is the only supported drag/drop backend.
    """
    try:
        from tkinterdnd2 import COPY, DND_FILES, TkinterDnD  # type: ignore

        # Importing tkinterdnd2 adds DnD methods to tkinter.BaseWidget. Loading
        # tkdnd into the existing CTk interpreter lets us keep the CTk root.
        TkinterDnD._require(root)  # type: ignore[attr-defined]
        widget.drop_target_register(DND_FILES)

        def _drop(event):
            try:
                paths = list(widget.tk.splitlist(event.data))
            except Exception:
                raw = str(getattr(event, "data", "") or "").strip()
                paths = [raw] if raw else []
            if paths:
                # TkinterDnD dispatches this callback on Tk's event thread.
                on_files(paths)
            return getattr(event, "action", COPY) or COPY

        widget.dnd_bind("<<Drop>>", _drop)

        def _cleanup():
            try:
                widget.drop_target_unregister()
            except Exception:
                pass

        return FileDropRegistration("tkinterdnd2", _cleanup)
    except Exception:
        logger.debug(
            "Drag-and-drop unavailable; install/package tkinterdnd2 to enable it",
            exc_info=True,
        )
        return None


def install_file_drop(
    root,
    widget,
    on_files: Callable[[Iterable[str]], None],
) -> FileDropRegistration | None:
    """Enable invisible file drag-and-drop when a safe Tk DnD backend exists.

    Drag-and-drop is an enhancement, not a requirement: when tkinterdnd2 is not
    available the app keeps normal file/folder picker behavior with no visible
    placeholder/drop-zone UI.
    """
    return _install_tkinterdnd(root, widget, on_files)
