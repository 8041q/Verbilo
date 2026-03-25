# icons.py — centralised icon loading via pytablericons (Tabler Icons)

from __future__ import annotations
import io
from functools import lru_cache
from typing import TYPE_CHECKING
from pathlib import Path
from importlib import resources
from .theme import get_mode

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


def _load_asset_image(filename: str, size: int | None = None):
    if PILImage is None:
        return None
    # filesystem first
    assets_dir = Path(__file__).resolve().parent.parent / "assets"
    try:
        p = assets_dir / filename
        if p.exists():
            pil = PILImage.open(p).convert("RGBA")
            if size is not None and pil.size != (size, size):
                try:
                    resample = PILImage.Resampling.LANCZOS
                except AttributeError:
                    resample = PILImage.LANCZOS
                pil = pil.resize((size, size), resample)
            return pil
    except Exception:
        pass

    # try any basename-*.png in assets (filesystem only — good for dev)
    try:
        basename = filename.rsplit(".", 1)[0]
        for f in sorted(assets_dir.glob(f"{basename}-*.png")):
            try:
                pil = PILImage.open(f).convert("RGBA")
                if size is not None and pil.size != (size, size):
                    try:
                        resample = PILImage.Resampling.LANCZOS
                    except AttributeError:
                        resample = PILImage.LANCZOS
                    pil = pil.resize((size, size), resample)
                return pil
            except Exception:
                continue
    except Exception:
        pass

    # packaged resource fallback (original logic)
    try:
        asset_trav = resources.files("verbilo.assets").joinpath(filename)
        data = asset_trav.read_bytes()
        pil = PILImage.open(io.BytesIO(data)).convert("RGBA")
        if size is not None and pil.size != (size, size):
            try:
                resample = PILImage.Resampling.LANCZOS
            except AttributeError:
                resample = PILImage.LANCZOS
            pil = pil.resize((size, size), resample)
        return pil
    except Exception:
        pass

    return None


def _load_asset_pair(name: str, size: int | None = None) -> tuple[object | None, object | None]:
    """Return (pil_light, pil_dark) for the given base name.

    Tries filesystem first then packaged resources. If only a single image
    exists (no -light/-dark variants), that image is returned for both
    light and dark as a fallback.
    """
    if PILImage is None:
        return None, None

    assets_dir = Path(__file__).resolve().parent.parent / "assets"

    def _open_file(p: Path):
        try:
            pil = PILImage.open(p).convert("RGBA")
            if size is not None and pil.size != (size, size):
                try:
                    resample = PILImage.Resampling.LANCZOS
                except AttributeError:
                    resample = PILImage.LANCZOS
                pil = pil.resize((size, size), resample)
            return pil
        except Exception:
            return None

    light = None
    dark = None

    # Filesystem explicit candidates
    try:
        # try size-specific then generic
        if size is not None:
            for fn in (f"{name}-{size}-light.png", f"{name}-light.png"):
                p = assets_dir / fn
                if p.exists():
                    light = _open_file(p)
                    break
            for fn in (f"{name}-{size}-dark.png", f"{name}-dark.png"):
                p = assets_dir / fn
                if p.exists():
                    dark = _open_file(p)
                    break
        else:
            for fn in (f"{name}-light.png",):
                p = assets_dir / fn
                if p.exists():
                    light = _open_file(p)
                    break
            for fn in (f"{name}-dark.png",):
                p = assets_dir / fn
                if p.exists():
                    dark = _open_file(p)
                    break

        # glob-style fallback (basename-*-light.png / basename-*-dark.png)
        if light is None:
            for f in sorted(assets_dir.glob(f"{name}-*-light.png")):
                light = _open_file(f)
                if light is not None:
                    break
        if dark is None:
            for f in sorted(assets_dir.glob(f"{name}-*-dark.png")):
                dark = _open_file(f)
                if dark is not None:
                    break
    except Exception:
        pass

    # Packaged resource fallback
    try:
        if light is None:
            if size is not None:
                for fn in (f"{name}-{size}-light.png", f"{name}-light.png"):
                    try:
                        asset_trav = resources.files("verbilo.assets").joinpath(fn)
                        data = asset_trav.read_bytes()
                        light = PILImage.open(io.BytesIO(data)).convert("RGBA")
                        if size is not None and light.size != (size, size):
                            try:
                                resample = PILImage.Resampling.LANCZOS
                            except AttributeError:
                                resample = PILImage.LANCZOS
                            light = light.resize((size, size), resample)
                        break
                    except Exception:
                        continue
            else:
                try:
                    asset_trav = resources.files("verbilo.assets").joinpath(f"{name}-light.png")
                    data = asset_trav.read_bytes()
                    light = PILImage.open(io.BytesIO(data)).convert("RGBA")
                except Exception:
                    pass

        if dark is None:
            if size is not None:
                for fn in (f"{name}-{size}-dark.png", f"{name}-dark.png"):
                    try:
                        asset_trav = resources.files("verbilo.assets").joinpath(fn)
                        data = asset_trav.read_bytes()
                        dark = PILImage.open(io.BytesIO(data)).convert("RGBA")
                        if size is not None and dark.size != (size, size):
                            try:
                                resample = PILImage.Resampling.LANCZOS
                            except AttributeError:
                                resample = PILImage.LANCZOS
                            dark = dark.resize((size, size), resample)
                        break
                    except Exception:
                        continue
            else:
                try:
                    asset_trav = resources.files("verbilo.assets").joinpath(f"{name}-dark.png")
                    data = asset_trav.read_bytes()
                    dark = PILImage.open(io.BytesIO(data)).convert("RGBA")
                except Exception:
                    pass
    except Exception:
        pass

    # If we found neither light nor dark specifically, try single-file fallback
    if light is None and dark is None:
        for fn in ([f"{name}-{size}.png"] if size is not None else []) + [f"{name}.png"]:
            pil = _load_asset_image(fn, size=size)
            if pil is not None:
                light = dark = pil
                break

    return light, dark
    

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
    # prefer packaged/static asset variants (light + dark). Try size-specific
    # then generic names. For the special `language` icon, also try the
    # 2logo name.
    cand_bases = ["2logo", "language"] if name == "language" else [name]
    for base in cand_bases:
        pil_light, pil_dark = _load_asset_pair(base, size=size)
        if (pil_light is not None or pil_dark is not None) and ctk is not None:
            # If only one variant exists, use it for both light and dark to
            # preserve previous behaviour.
            if pil_light is None:
                pil_light = pil_dark
            if pil_dark is None:
                pil_dark = pil_light
            return ctk.CTkImage(light_image=pil_light, dark_image=pil_dark, size=(size, size))

    # existing guard (fallback to TablerIcons)
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


def get_photo_image(name: str, size: int = 18, color: str = "#C8CCD2", mode: str | None = None) -> object | None:
    # prefer static asset for PhotoImage. PhotoImage is a single-image
    # object, so pick the variant matching the current theme mode.
    if mode is None:
        try:
            mode = get_mode()
        except Exception:
            mode = "Light"

    cand_bases = ["2logo", "language"] if name == "language" else [name]
    for base in cand_bases:
        pil_light, pil_dark = _load_asset_pair(base, size=size)
        if pil_light is not None or pil_dark is not None:
            # choose by mode (prefer dark when mode contains 'dark')
            mode_l = (mode or "").lower()
            preferred = pil_dark if "dark" in mode_l else pil_light
            if preferred is None:
                preferred = pil_light or pil_dark
            if preferred is not None and ImageTk is not None:
                return ImageTk.PhotoImage(preferred)

    # Return a tkinter PhotoImage for use in ttk widgets (single image) —
    # fall back to TablerIcons rendering when no static asset found.
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
