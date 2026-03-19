# icons.py — centralised icon loading via pytablericons (Tabler Icons)

from __future__ import annotations
import io
from functools import lru_cache
from typing import TYPE_CHECKING
from pathlib import Path
from importlib import resources

try:
    import customtkinter as ctk
except Exception:
    ctk = None  # type: ignore[assignment]

try:
    from pytablericons import TablerIcons, OutlineIcon, FilledIcon  # type: ignore[import-untyped]
    from PIL import Image as PILImage, ImageTk  # type: ignore[import-untyped]

    _HAS_ICONS = True
except Exception:  
    TablerIcons = None  # type: ignore[assignment,misc]
    OutlineIcon = None  # type: ignore[assignment]
    FilledIcon = None  # type: ignore[assignment]
    PILImage = None  # type: ignore[assignment]
    ImageTk = None  # type: ignore[assignment]
    _HAS_ICONS = False

if TYPE_CHECKING:
    from customtkinter import CTkImage

_ICON_MAP: dict[str, object] = {}

if OutlineIcon is not None:
    _ICON_MAP = {
        # Sidebar / navigation
        "language":         OutlineIcon.LANGUAGE,
        "world":            OutlineIcon.WORLD,
        "settings":         OutlineIcon.SETTINGS,

        # Actions
        "play":             OutlineIcon.PLAYER_PLAY,
        "stop":             OutlineIcon.PLAYER_STOP,
        "add-file":         OutlineIcon.FILE_PLUS,
        "open-folder":      OutlineIcon.FOLDER_OPEN,
        "folder":           OutlineIcon.FOLDER,
        "trash":            OutlineIcon.TRASH,
        "search":           OutlineIcon.SEARCH,
        "chevron-down":     OutlineIcon.CHEVRON_DOWN,
        "x":                OutlineIcon.X,

        # File types
        "file-docx":        OutlineIcon.FILE_TYPE_DOCX,
        "file-pdf":         OutlineIcon.FILE_TYPE_PDF,
        "file-xls":         OutlineIcon.FILE_TYPE_XLS,
        "file-unknown":     OutlineIcon.FILE_UNKNOWN,
        "file":             OutlineIcon.FILE,

        # Info / about
        "info":             OutlineIcon.INFO_CIRCLE,
    }


def _icon_colors() -> tuple[str, str]:
    # Return (light_color, dark_color) for icon rendering
    return "#3B3F46", "#C8CCD2"


def _accent_icon_colors() -> tuple[str, str]:
    # Colours for icons on accent-coloured buttons (light, dark).
    return "#FFFFFF", "#FFFFFF"


# Public API 

@lru_cache(maxsize=256)
def _render(icon_enum: object, size: int, color: str) -> object:
    # Render a single icon to a PIL Image (cached).
    if TablerIcons is None:
        return None
    return TablerIcons.load(icon_enum, size=size, color=color, stroke_width=1.8)


def get_icon(
    name: str,
    size: int = 18,
    *,
    on_accent: bool = False,
) -> "CTkImage | None":

    if ctk is None or TablerIcons is None:
        return None

    icon_enum = _ICON_MAP.get(name)
    if icon_enum is None:
        return None

    if on_accent:
        light_c, dark_c = _accent_icon_colors()
    else:
        light_c, dark_c = _icon_colors()

    light_img = _render(icon_enum, size, light_c)
    dark_img = _render(icon_enum, size, dark_c)

    if light_img is None or dark_img is None:
        return None

    return ctk.CTkImage(light_image=light_img, dark_image=dark_img, size=(size, size))


def get_photo_image(name: str, size: int = 18, color: str = "#C8CCD2") -> object | None:
    # Return a tkinter PhotoImage for use in ttk widgets (single image)
    if TablerIcons is None or PILImage is None or ImageTk is None:
        return None

    icon_enum = _ICON_MAP.get(name)
    if icon_enum is None:
        return None

    pil = _render(icon_enum, size, color)
    if pil is None:
        return None
    return ImageTk.PhotoImage(pil)


def get_app_icon(size: int = 64) -> object | None:
    # Load favicon from package resources (verbilo/assets/favicon.ico)
    if PILImage is not None:
        # Strategy 1: __file__-relative path (reliable in editable installs & Nuitka)
        try:
            _favicon = Path(__file__).resolve().parent.parent / "assets" / "favicon.ico"
            if _favicon.exists():
                pil = PILImage.open(_favicon).convert("RGBA")
                if pil.size != (size, size):
                    try:
                        resample = PILImage.Resampling.LANCZOS
                    except AttributeError:
                        resample = PILImage.LANCZOS  # type: ignore[attr-defined]
                    pil = pil.resize((size, size), resample)
                return pil
        except Exception:
            pass

        # Strategy 2: importlib.resources (wheel installs)
        try:
            asset_trav = resources.files("verbilo.assets").joinpath("favicon.ico")
            data = asset_trav.read_bytes()
            pil = PILImage.open(io.BytesIO(data)).convert("RGBA")
            if pil.size != (size, size):
                try:
                    resample = PILImage.Resampling.LANCZOS
                except AttributeError:
                    resample = PILImage.LANCZOS  # type: ignore[attr-defined]
                pil = pil.resize((size, size), resample)
            return pil
        except Exception:
            pass

    # Fallback to SVG icon
    if TablerIcons is None or OutlineIcon is None:
        return None
    return TablerIcons.load(OutlineIcon.LANGUAGE, size=size, color="#5B9BD5", stroke_width=1.6)

def apply_window_icon(root: object, size: int = 64) -> bool:
    # Apply the app icon using in-memory PhotoImage. Avoids requiring a filesystem path
    try:
        # Lazy import so icons module stays importable without tkinter/Pillow
        from PIL import ImageTk
    except Exception:
        ImageTk = None  # type: ignore[assignment]

    try:
        icon_pil = get_app_icon(size=size)
        if icon_pil is None:
            return False

            # Create an in-memory PhotoImage and set with iconphoto.
        if ImageTk is not None:
            try:
                icon_tk = ImageTk.PhotoImage(icon_pil)
                try:
                    # Prefer low-level call so it works even if CTk monkey-patches things
                    root.tk.call("wm", "iconphoto", root._w, "-default", icon_tk._PhotoImage__photo)
                except Exception:
                    try:
                        root.iconphoto(True, icon_tk)
                    except Exception:
                        pass
                # Keep a reference so the image isn't garbage-collected
                setattr(root, "_app_icon_ref", icon_tk)
            except Exception:
                return False

        def _locked_iconbitmap(*args, **kwargs):
            # No disk-based ico; we re-apply the in-memory icon if needed.
            try:
                if ImageTk is not None and getattr(root, "_app_icon_ref", None) is not None:
                    # Re-apply in-memory icon
                    try:
                        root.tk.call("wm", "iconphoto", root._w, "-default", root._app_icon_ref._PhotoImage__photo)
                    except Exception:
                        try:
                            root.iconphoto(True, root._app_icon_ref)
                        except Exception:
                            pass
            except Exception:
                pass

        try:
            root.iconbitmap = _locked_iconbitmap  # type: ignore[method-assign]
        except Exception:
            pass

        return True
    except Exception:
        return False
    