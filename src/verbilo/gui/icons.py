# icons.py — centralised icon loading via pytablericons (Tabler Icons, MIT)
#
# Every icon used in the app is loaded through get_icon().  It returns a CTkImage
# with separate light/dark variants so icons adapt to the active theme.
# Results are cached to avoid re-rendering the same SVG multiple times.

from __future__ import annotations

from functools import lru_cache
from typing import TYPE_CHECKING

from pathlib import Path

try:
    import customtkinter as ctk
except Exception:
    ctk = None  # type: ignore[assignment]

try:
    from pytablericons import TablerIcons, OutlineIcon, FilledIcon  # type: ignore[import-untyped]
    from PIL import Image as PILImage, ImageTk  # type: ignore[import-untyped]

    _HAS_ICONS = True
except Exception:  # pragma: no cover
    TablerIcons = None  # type: ignore[assignment,misc]
    OutlineIcon = None  # type: ignore[assignment]
    FilledIcon = None  # type: ignore[assignment]
    PILImage = None  # type: ignore[assignment]
    ImageTk = None  # type: ignore[assignment]
    _HAS_ICONS = False

if TYPE_CHECKING:
    from customtkinter import CTkImage

# Semantic icon map
# Maps human-readable names to OutlineIcon enum members.

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


# Colour tokens for light / dark icons 

def _icon_colors() -> tuple[str, str]:
    # Return (light_color, dark_color) for icon rendering.
    # light_color: used on light theme (choose a dark colour).
    # dark_color: used on dark theme (choose a light colour).
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
    # Return a CTkImage (light + dark variants) for `name`.
    # Returns None when required libraries or icon are unavailable.
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
    # Return a tkinter PhotoImage for use in ttk widgets (single image).
    # Use the provided colour for the current theme.
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
    asset_path = Path(__file__).resolve().parents[3] / "assets" / "favicon.ico"

    try:
        if asset_path.exists() and PILImage is not None:
            pil = PILImage.open(asset_path).convert("RGBA")
            if pil.size != (size, size):
                try:
                    resample = PILImage.Resampling.LANCZOS
                except AttributeError:
                    resample = PILImage.LANCZOS
                pil = pil.resize((size, size), resample)
            return pil
    except Exception:
        return None

    # Fallback to existing logic
    if TablerIcons is None or OutlineIcon is None:
        return None
    return TablerIcons.load(OutlineIcon.LANGUAGE, size=size, color="#5B9BD5", stroke_width=1.6)


def apply_window_icon(root: object, size: int = 64) -> bool:
    """Apply the app icon to a Tk/Toplevel window.

    Tries to set a .ico file on Windows (via `iconbitmap`) and then sets
    a `PhotoImage` via `iconphoto` so the icon appears in titlebar and
    taskbar. Stores a reference on the window as `_app_icon_ref` to prevent
    garbage collection.

    After successfully setting the icon, monkey-patches `root.iconbitmap` on
    the instance level so CustomTkinter cannot overwrite it with its own logo
    in any subsequent scheduled `after()` calls.

    Returns True on success, False otherwise.
    """
    try:
        # Lazy import so icons module stays importable without tkinter/Pillow
        from PIL import ImageTk
    except Exception:
        ImageTk = None  # type: ignore[assignment]

    try:
        icon_pil = get_app_icon(size=size)
        if icon_pil is None:
            return False

        asset = Path(__file__).resolve().parents[3] / "assets" / "favicon.ico"
        asset_str = str(asset) if asset.exists() else None

        def _set_ico():
            if asset_str:
                try:
                    # Low-level call bypasses CTk's Python-level iconbitmap override
                    root.tk.call("wm", "iconbitmap", root._w, asset_str)
                except Exception:
                    pass

        _set_ico()

        if ImageTk is not None:
            try:
                icon_tk = ImageTk.PhotoImage(icon_pil)
                try:
                    root.tk.call("wm", "iconphoto", root._w, "-default", icon_tk._PhotoImage__photo)
                except Exception:
                    try:
                        root.iconphoto(True, icon_tk)
                    except Exception:
                        pass
                setattr(root, "_app_icon_ref", icon_tk)
            except Exception:
                return False

        # Monkey-patch iconbitmap on the *instance* so CTk's scheduled calls
        # (e.g. root.after(200, _windows_set_titlebar_icon)) become no-ops.
        # We still re-apply our own .ico each time so nothing can dislodge it.
        def _locked_iconbitmap(*args, **kwargs):
            _set_ico()

        try:
            root.iconbitmap = _locked_iconbitmap  # type: ignore[method-assign]
        except Exception:
            pass

        return True
    except Exception:
        return False