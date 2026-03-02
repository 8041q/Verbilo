# gui_theme.py — design tokens (colours, fonts, spacing); edit here to restyle everything

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, TYPE_CHECKING, Literal

try:
    import customtkinter as ctk
except Exception:
    ctk = None  # type: ignore[assignment]

if TYPE_CHECKING:
    from customtkinter import CTkFrame, CTkButton, CTkLabel, CTkEntry, CTkFont


# layout

PADDING: int = 20
CARD_CORNER_RADIUS: int = 12
CARD_BORDER_WIDTH: int = 1
SIDEBAR_WIDTH: int = 250
BUTTON_CORNER_RADIUS: int = 8
WINDOW_WIDTH: int = 1100
WINDOW_HEIGHT: int = 700
WINDOW_MIN_WIDTH: int = 950
WINDOW_MIN_HEIGHT: int = 650

# fonts

FONT_FAMILY: str = "Roboto"

FONT_HEADING: tuple[str, int, str] = (FONT_FAMILY, 18, "bold")
FONT_SUBHEADING: tuple[str, int, str] = (FONT_FAMILY, 14, "bold")
FONT_BODY: tuple[str, int] = (FONT_FAMILY, 13)
FONT_SMALL: tuple[str, int] = (FONT_FAMILY, 13)
FONT_TINY: tuple[str, int] = (FONT_FAMILY, 16)


# palette

# all colours are semantic — bg/text/accent/border/status
@dataclass(frozen=True)
class Palette:

    # Backgrounds
    bg_main: str            # root window / main area
    bg_card: str            # card surfaces (slightly lighter/darker than main)
    bg_sidebar: str         # sidebar background
    bg_input: str           # entry / combo-box background
    bg_popup: str           # popup overlay background
    bg_row_even: str        # treeview even row
    bg_row_odd: str         # treeview odd row
    bg_heading: str         # treeview heading row

    # Text
    text_primary: str       # headings, important labels
    text_secondary: str     # body text, descriptions
    text_muted: str         # placeholders, disabled text
    text_on_accent: str     # text rendered on accent-coloured surfaces

    # Accent / brand
    accent: str             # primary action colour (buttons, selections)
    accent_hover: str       # hover state
    accent_pressed: str     # pressed / active state

    # Borders & dividers
    border: str             # card / button border
    divider: str            # thin section dividers

    # Semantic status (shared across modes)
    status_success: str = "#4CAF50"
    status_error: str   = "#F44336"
    status_warning: str = "#FF9800"
    status_info: str    = "#4DA6FF"
    status_pending: str = "#888888"


# brand: #4A4A4A dark grey, #CBCBCB silver, #FFFFE3 cream, #6D8196 steel-blue

DARK = Palette(
    bg_main="#1E1E1E",
    bg_card="#2B2B2B",
    bg_sidebar="#171717",
    bg_input="#333333",
    bg_popup="#2B2B2B",
    bg_row_even="#2B2B2B",
    bg_row_odd="#323232",
    bg_heading="#242424",

    text_primary="#FFFFE3",
    text_secondary="#CBCBCB",
    text_muted="#6B6B6B",
    text_on_accent="#FFFFE3",

    accent="#6D8196",
    accent_hover="#7E95AB",
    accent_pressed="#5A6E80",

    border="#4A4A4A",
    divider="#3A3A3A",
)

LIGHT = Palette(
    bg_main="#F5F5F0",
    bg_card="#FFFFFF",
    bg_sidebar="#EAEAE5",
    bg_input="#FFFFFF",
    bg_popup="#FFFFFF",
    bg_row_even="#FFFFFF",
    bg_row_odd="#F7F7F2",
    bg_heading="#EAEAE5",

    text_primary="#1E1E1E",
    text_secondary="#4A4A4A",
    text_muted="#999999",
    text_on_accent="#FFFFE3",

    accent="#6D8196",
    accent_hover="#5A6E80",
    accent_pressed="#4A5E70",

    border="#CBCBCB",
    divider="#E0E0DA",
)


# runtime state

_current_mode: str = "Dark"


def set_mode(mode: str) -> None:
    # "Dark" or "Light"
    global _current_mode
    _current_mode = mode
    if ctk is not None:
        ctk.set_appearance_mode(mode)


def get_mode() -> str:
    return _current_mode


def get() -> Palette:
    return DARK if _current_mode == "Dark" else LIGHT


# widget helpers

def make_card(parent: Any, **overrides: Any) -> Any:
    # themed card frame
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


def make_button(parent: Any, text: str, command: Any = None, style: str = "primary",
                **overrides: Any) -> Any:
    # style: "primary" (filled), "secondary" (outline), "ghost" (transparent)
    if ctk is None:
        import tkinter as _tk
        return _tk.Button(parent, text=text, command=command)
    p = get()
    _font = ctk.CTkFont(family=FONT_FAMILY, size=FONT_BODY[1])
    if style == "primary":
        kw: dict[str, Any] = dict(
            fg_color=p.accent,
            hover_color=p.accent_hover,
            text_color=p.text_on_accent,
            border_color=p.border,
            border_width=1,
        )
    elif style == "secondary":
        kw: dict[str, Any] = dict(
            fg_color="transparent",
            hover_color=p.bg_card,
            text_color=p.text_secondary,
            border_color=p.border,
            border_width=1,
        )
    else:  # ghost
        kw: dict[str, Any] = dict(
            fg_color="transparent",
            hover_color=p.bg_card,
            text_color=p.text_secondary,
            border_color=p.bg_sidebar,
            border_width=0,
        )
    kw.update(overrides)
    return ctk.CTkButton(
        parent, text=text, command=command,
        corner_radius=kw.pop("corner_radius", BUTTON_CORNER_RADIUS),  # type: ignore[arg-type]
        font=kw.pop("font", _font),  # type: ignore[arg-type]
        **kw,  # type: ignore[arg-type]
    )


def make_label(parent: Any, text: str, level: str = "body", **overrides: Any) -> Any:
    # level: "heading", "subheading", "body", "small", "muted", "section"
    if ctk is None:
        import tkinter as _tk
        return _tk.Label(parent, text=text)
    p = get()
    font_map = {
        "heading":    (FONT_HEADING[1],    FONT_HEADING[2] if len(FONT_HEADING) > 2 else "normal"),
        "subheading": (FONT_SUBHEADING[1], FONT_SUBHEADING[2] if len(FONT_SUBHEADING) > 2 else "normal"),
        "body":       (FONT_BODY[1],       "normal"),
        "small":      (FONT_SMALL[1],      "normal"),
        "muted":      (FONT_SMALL[1],      "normal"),
        "section":    (FONT_TINY[1],       "bold"),
    }
    size, weight = font_map.get(level, font_map["body"])
    color_map = {
        "heading":    p.text_primary,
        "subheading": p.text_primary,
        "body":       p.text_secondary,
        "small":      p.text_secondary,
        "muted":      p.text_muted,
        "section":    p.text_muted,
    }
    text_color = color_map.get(level, p.text_secondary)
    display_text = text.upper() if level == "section" else text

    _font = ctk.CTkFont(family=FONT_FAMILY, size=size, weight=weight)  # type: ignore[arg-type]
    overrides.setdefault("anchor", "w")
    return ctk.CTkLabel(
        parent,
        text=display_text,
        font=_font,
        text_color=overrides.pop("text_color", text_color),
        **overrides,
    )


def make_entry(parent: Any, **overrides: Any) -> Any:
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
    if ctk is None:
        import tkinter as _tk
        return _tk.Frame(parent, height=1, bg="#555555")
    p = get()
    if orientation == "horizontal":
        return ctk.CTkFrame(parent, height=1, fg_color=p.divider, corner_radius=0)
    return ctk.CTkFrame(parent, width=1, fg_color=p.divider, corner_radius=0)
