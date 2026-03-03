# icons.py — centralised icon loading via pytablericons (Tabler Icons, MIT)
#
# Every icon used in the app is loaded through get_icon().  It returns a CTkImage
# with separate light/dark variants so icons adapt to the active theme.
# Results are cached to avoid re-rendering the same SVG multiple times.

from __future__ import annotations

from functools import lru_cache
from typing import TYPE_CHECKING

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
    # Return a PIL Image suitable for the window/taskbar icon.
    if TablerIcons is None or OutlineIcon is None:
        return None
    return TablerIcons.load(OutlineIcon.LANGUAGE, size=size, color="#5B9BD5", stroke_width=1.6)
