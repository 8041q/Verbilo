# gui_theme.py — design tokens (colours, fonts, spacing); edit here to restyle everything

from __future__ import annotations

import platform
import sys
from dataclasses import dataclass
from typing import Any, TYPE_CHECKING, Literal

try:
    import customtkinter as ctk
except Exception:
    ctk = None  # type: ignore[assignment]

if TYPE_CHECKING:
    from customtkinter import CTkFrame, CTkButton, CTkLabel, CTkEntry, CTkFont


#  DPI / scaling
_scale_factor: float = 1.0

def init_dpi(root: Any = None) -> float:
    # Detect and apply DPI scaling.
    global _scale_factor

    if root is not None:
        try:
            root.update_idletasks()
            dpi = root.winfo_fpixels("1i")  # pixels per inch
            _scale_factor = dpi / 96.0
        except Exception:
            _scale_factor = 1.0
    else:
        _scale_factor = 1.0

    return _scale_factor


def scale(value: int | float) -> int:
    # Scale a pixel value by the current DPI factor.
    return max(1, round(value * _scale_factor))


#  Layout constants (logical pixels — call scale() at usage sites)
PADDING: int = 18
CARD_CORNER_RADIUS: int = 10
CARD_BORDER_WIDTH: int = 1
SIDEBAR_WIDTH: int = 260
BUTTON_CORNER_RADIUS: int = 8
WINDOW_WIDTH: int = 1100
WINDOW_HEIGHT: int = 735
WINDOW_MIN_WIDTH: int = 850
WINDOW_MIN_HEIGHT: int = 735

#  Fonts — proper hierarchy, OS-native family first
def _detect_font_family() -> str:
    _sys = platform.system()
    if _sys == "Windows":
        return "Segoe UI"
    if _sys == "Darwin":
        return "SF Pro Text"
    return "Roboto"


FONT_FAMILY: str = _detect_font_family()

# (family, size, ?weight) — sizes are logical
FONT_HEADING:    tuple[str, int, str] = (FONT_FAMILY, 23, "bold") # app title
FONT_SUBHEADING: tuple[str, int, str] = (FONT_FAMILY, 15, "bold") # headers titles
FONT_SECTION:    tuple[str, int, str] = (FONT_FAMILY, 11, "bold")  # sidebar sections (uppercased)
FONT_BODY:       tuple[str, int]      = (FONT_FAMILY, 12) # dropdown list
FONT_SMALL:      tuple[str, int]      = (FONT_FAMILY, 12) # identification (table headers, settings option)
FONT_TINY:       tuple[str, int]      = (FONT_FAMILY, 11) # info and notes


#  Palette — semantic colour tokens

@dataclass(frozen=True)
class Palette:
    # Backgrounds
    bg_main: str
    bg_card: str
    bg_sidebar: str
    bg_input: str
    bg_popup: str
    bg_row_even: str
    bg_row_odd: str
    bg_heading: str

    # Text
    text_primary: str
    text_secondary: str
    text_muted: str
    text_on_accent: str

    # Accent / brand
    accent: str
    accent_hover: str
    accent_pressed: str

    # Borders & dividers
    border: str
    divider: str

    # Semantic status (shared)
    status_success: str = "#4CAF50"
    status_error: str   = "#F44336"
    status_warning: str = "#FF9800"
    status_info: str    = "#4DA6FF"
    status_pending: str = "#888888"


# Dark mode

DARK = Palette(
    bg_main="#181A1F",
    bg_card="#22252B",
    bg_sidebar="#14161A",
    bg_input="#2A2D34",
    bg_popup="#22252B",
    bg_row_even="#22252B",
    bg_row_odd="#282B32",
    bg_heading="#1C1E24",

    text_primary="#E8EAED",
    text_secondary="#B0B3B8",
    text_muted="#6B6F78",
    text_on_accent="#CECECE",

    accent="#6C8878",
    accent_hover="#97B9A5",
    accent_pressed="#8DAC9A",

    border="#33363D",
    divider="#2C2F36",
)

# Light mode

LIGHT = Palette(
    bg_main="#F4F5F7",
    bg_card="#FFFFFF",
    bg_sidebar="#EBEDF0",
    bg_input="#FFFFFF",
    bg_popup="#FFFFFF",
    bg_row_even="#FFFFFF",
    bg_row_odd="#F7F8FA",
    bg_heading="#EBEDF0",

    text_primary="#1D1F24",
    text_secondary="#4A4E57",
    text_muted="#9198A1",
    text_on_accent="#FFFFFF",

    accent="#82AD95",
    accent_hover="#97B9A5",
    accent_pressed="#8DAC9A",

    border="#D0D3D9",
    divider="#E2E4E8",
)


#  Runtime state

_current_mode: str = "Dark"


def set_mode(mode: str) -> None:
    global _current_mode
    _current_mode = mode
    if ctk is not None:
        ctk.set_appearance_mode(mode)


def get_mode() -> str:
    return _current_mode


def get() -> Palette:
    return DARK if _current_mode == "Dark" else LIGHT


#  Widget factory helpers
def make_card(parent: Any, **overrides: Any) -> Any:
    # Themed card frame.
    if ctk is None:
        import tkinter as _tk
        return _tk.Frame(parent)
    p = get()
    return ctk.CTkFrame(
        parent,
        corner_radius=overrides.pop("corner_radius", CARD_CORNER_RADIUS),
        border_width=overrides.pop("border_width", CARD_BORDER_WIDTH),
        fg_color=overrides.pop("fg_color", p.bg_card),
        border_color=overrides.pop("border_color", p.border),
        **overrides,
    )


def make_button(
    parent: Any,
    text: str,
    command: Any = None,
    style: str = "primary",
    image: Any = None,
    **overrides: Any,
) -> Any:
    # Themed button. style: "primary" / "secondary" / "ghost".
    if ctk is None:
        import tkinter as _tk
        p = get()
        # Determine disabled text colour by style so disabled text remains readable
        if style == "primary":
            disabled_color = p.text_on_accent
        elif style == "secondary":
            disabled_color = p.text_secondary
        else:
            disabled_color = p.text_muted

        # Allow callers to override `disabledforeground`; default to style mapping
        disabled_fg = overrides.pop("disabledforeground", disabled_color)
        btn = _tk.Button(parent, text=text, command=command, disabledforeground=disabled_fg, **overrides)
        try:
            setattr(btn, "_verbilo_style", style)
        except Exception:
            pass
        return btn
    p = get()
    _font = ctk.CTkFont(family=FONT_FAMILY, size=FONT_BODY[1])

    if style == "primary":
        kw: dict[str, Any] = dict(
            fg_color=p.accent,
            hover_color=p.accent_hover,
            text_color=p.text_on_accent,
            border_color=p.accent_pressed,
            border_width=0,
        )
    elif style == "secondary":
        kw = dict(
            fg_color="transparent",
            hover_color=p.bg_sidebar,
            text_color=p.text_secondary,
            border_color=p.border,
            border_width=1,
        )
    else:  # ghost
        kw = dict(
            fg_color="transparent",
            hover_color=p.bg_sidebar,
            text_color=p.text_secondary,
            border_color=p.bg_sidebar,
            border_width=0,
        )

    kw.update(overrides)

    # Ensure CTkButton shows a readable disabled text color by default (style-aware)
    if style == "primary":
        disabled_color = p.text_on_accent
    elif style == "secondary":
        disabled_color = p.text_secondary
    else:
        disabled_color = p.text_muted

    kw.setdefault("text_color_disabled", disabled_color)

    btn_args: dict[str, Any] = dict(
        text=text,
        command=command,
        corner_radius=kw.pop("corner_radius", BUTTON_CORNER_RADIUS),
        font=kw.pop("font", _font),
    )
    if image is not None:
        btn_args["image"] = image
        btn_args["compound"] = "left"

    btn = ctk.CTkButton(parent, **btn_args, **kw)
    try:
        setattr(btn, "_verbilo_style", style)
    except Exception:
        pass
    return btn


def make_label(parent: Any, text: str, level: str = "body", **overrides: Any) -> Any:
    # Themed label. Levels: heading, subheading, section, body, small, tiny, muted.
    if ctk is None:
        import tkinter as _tk
        return _tk.Label(parent, text=text)
    p = get()

    font_map: dict[str, tuple[int, str]] = {
        "heading":    (FONT_HEADING[1],    "bold"),
        "subheading": (FONT_SUBHEADING[1], "bold"),
        "section":    (FONT_SECTION[1],    "bold"),
        "body":       (FONT_BODY[1],       "normal"),
        "small":      (FONT_SMALL[1],      "normal"),
        "tiny":       (FONT_TINY[1],       "normal"),
        "muted":      (FONT_SMALL[1],      "normal"),
    }
    size, weight = font_map.get(level, font_map["body"])

    color_map: dict[str, str] = {
        "heading":    p.text_primary,
        "subheading": p.text_primary,
        "section":    p.text_muted,
        "body":       p.text_secondary,
        "small":      p.text_secondary,
        "tiny":       p.text_muted,
        "muted":      p.text_muted,
    }
    text_color = color_map.get(level, p.text_secondary)
    display_text = text.upper() if level == "section" else text

    _font = ctk.CTkFont(family=FONT_FAMILY, size=size, weight=weight)
    overrides.setdefault("anchor", "w")
    return ctk.CTkLabel(
        parent,
        text=display_text,
        font=_font,
        text_color=overrides.pop("text_color", text_color),
        **overrides,
    )


def make_entry(parent: Any, **overrides: Any) -> Any:
    # Themed text entry.
    if ctk is None:
        import tkinter as _tk
        return _tk.Entry(parent)
    p = get()
    _font = ctk.CTkFont(family=FONT_FAMILY, size=FONT_BODY[1])
    return ctk.CTkEntry(
        parent,
        fg_color=overrides.pop("fg_color", p.bg_input),
        border_color=overrides.pop("border_color", p.border),
        text_color=overrides.pop("text_color", p.text_secondary),
        placeholder_text_color=overrides.pop("placeholder_text_color", p.text_muted),
        corner_radius=overrides.pop("corner_radius", BUTTON_CORNER_RADIUS),
        border_width=overrides.pop("border_width", 1),
        font=overrides.pop("font", _font),
        **overrides,
    )


def make_divider(parent: Any, orientation: str = "horizontal") -> Any:
    # Thin coloured divider line.
    if ctk is None:
        import tkinter as _tk
        return _tk.Frame(parent, height=1, bg="#555555")
    p = get()
    if orientation == "horizontal":
        return ctk.CTkFrame(parent, height=1, fg_color=p.divider, corner_radius=0)
    return ctk.CTkFrame(parent, width=1, fg_color=p.divider, corner_radius=0)
