# gui_customtk.py — dashboard GUI; sidebar (controls) + content (file table, progress, log)

from __future__ import annotations

import os
import json
import threading
import urllib.request
import urllib.error
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import logging

import customtkinter as ctk

import webbrowser

from . import theme
from .helpers import Worker, list_supported_files, center_window, GuiLoggingHandler, SUPPORTED_EXTS
from .config import load_config, save_config
from .icons import get_icon, get_photo_image, get_app_icon, apply_window_icon

logger = logging.getLogger(__name__)

# --- version & about constants ---
from .. import __version__ as APP_VERSION
from .. import __build_date__ as APP_BUILD_DATE

# Github URLs
GITHUB_URL = "https://github.com/8041q/Verbilo"
RELEASES_URL = "https://github.com/8041q/Verbilo/releases"
ISSUE_URL = "https://github.com/8041q/Verbilo/issues"

# Default folder name
DEFAULT_OUTPUT_FOLDER = "Output"
DEFAULT_INPUT_FOLDER = "Input"


def _try_make_relative(absolute_path: str) -> str:
    # Return a path relative to cwd if possible, otherwise return the original.
    try:
        return str(Path(absolute_path).relative_to(Path.cwd()))
    except ValueError:
        return absolute_path

# --- language helpers ---

# ISO 639-1 base codes supported by each detector backend.
# Used to filter the source-language dropdown when the detector changes.
_LINGUA_LANG_CODES: frozenset[str] = frozenset({
    # 75 languages supported by lingua-language-detector
    "af", "sq", "ar", "hy", "az", "eu", "be", "bn", "no", "bs", "bg", "ca",
    "zh", "hr", "cs", "da", "nl", "en", "eo", "et", "fi", "fr",
    "lg",  # Ganda
    "ka", "de",
    "el", "gu", "he", "hi", "hu", "is", "id", "ga", "it", "ja", "kk", "ko",
    "la", "lv", "lt", "mk", "ms", "mi", "mr", "mn", "fa", "pl", "pt", "pa",
    "ro", "ru", "sr", "sn", "sk", "sl", "so", "st", "es", "sw", "sv", "tl",
    "ta", "te", "th",
    "ts",  # Tsonga
    "tn",  # Tswana
    "tr", "uk", "ur", "vi", "cy", "xh", "yo", "zu",
})

_FASTTEXT_LANG_CODES: frozenset[str] = frozenset({
    # 176-language FastText lid.176 model (ISO 639-1 / BCP-47 base codes)
    "af", "am", "an", "ar", "as", "az", "ba", "be", "bg", "bn", "br", "bs",
    "ca", "ce", "co", "cs", "cv", "cy", "da", "de", "dv", "el", "en", "eo",
    "es", "et", "eu", "fa", "fi", "fr", "fy", "ga", "gd", "gl", "gn", "gu",
    "gv", "ha",  # Hausa
    "he", "hi", "hr", "ht", "hu", "hy", "ia", "id", "ig",  # Igbo
    "io", "is", "it",
    "ja", "jv", "jw",  # Javanese (jv canonical; jw = Google's alias)
    "ka", "kk", "km", "kn", "ko", "ku", "kw", "ky", "la", "lb",
    "li", "lo", "lt", "lv", "mg", "mi", "mk", "ml", "mn", "mr", "ms", "mt",
    "my", "ne", "nl", "nn", "no", "oc", "or", "os", "pa", "pl", "ps", "pt",
    "qu", "rm", "ro", "ru", "rw",  # Kinyarwanda
    "sa", "sc", "sd", "si", "sk", "sl", "sm",  # Samoan
    "sn", "so",
    "sq", "sr", "st", "su", "sv", "sw", "ta", "te", "tg", "th", "tk", "tl",
    "tr", "tt", "ug", "uk", "ur", "uz", "vi", "vo", "wa", "xh", "yi", "yo",
    "zh", "zu",
})

# ISO 639-1 codes that Baidu Translate supports (mapped from its native codes).
_BAIDU_LANG_CODES: frozenset[str] = frozenset({
    "ar", "bg", "cs", "da", "de", "el", "en", "es", "et", "fi",
    "fr", "hu", "it", "ja", "ko", "nl", "pl", "pt", "ro", "ru",
    "sl", "sv", "th", "vi", "zh", "zh-CN", "zh-TW",
})

# Translation engine choices (display name → internal key).
_ENGINE_OPTIONS: list[tuple[str, str]] = [
    ("Google Translate (free)", "google"),
    ("Google Cloud API", "google-cloud"),
    ("Baidu Translate", "baidu"),
]
_ENGINE_DISPLAY = [name for name, _ in _ENGINE_OPTIONS]
_ENGINE_MAP = {name: key for name, key in _ENGINE_OPTIONS}
_ENGINE_REVERSE = {key: name for name, key in _ENGINE_OPTIONS}


def _filter_by_detector(
    opts: list[tuple[str, str]], detector: str,
) -> list[tuple[str, str]]:
    # Return only (code, name) pairs the given detector can identify.
    codes = _LINGUA_LANG_CODES if detector == "lingua" else _FASTTEXT_LANG_CODES
    result = []
    for code, name in opts:
        base = code.lower().split("-")[0].split("_")[0]
        if base in codes:
            result.append((code, name))
    return result


def _filter_by_engine(
    opts: list[tuple[str, str]], engine: str,
) -> list[tuple[str, str]]:
    """Return only (code, name) pairs that the given translation engine supports.
    Google (free & cloud) supports all languages, so no filtering needed."""
    if engine == "baidu":
        result = []
        for code, name in opts:
            base = code.lower().split("_")[0]
            if base in _BAIDU_LANG_CODES or code in _BAIDU_LANG_CODES:
                result.append((code, name))
        return result
    return list(opts)  # google / google-cloud: no restriction


_cached_language_options: list[tuple[str, str]] | None = None


def _get_language_options() -> list[tuple[str, str]]:
    # returns (code, name) pairs for all supported languages
    global _cached_language_options
    if _cached_language_options is not None:
        return _cached_language_options

    fallback: dict[str, str] = {
        "af": "Afrikaans", "sq": "Albanian", "am": "Amharic", "ar": "Arabic",
        "hy": "Armenian", "az": "Azerbaijani", "eu": "Basque", "be": "Belarusian",
        "bn": "Bengali", "bs": "Bosnian", "bg": "Bulgarian", "ca": "Catalan",
        "ceb": "Cebuano", "zh-CN": "Chinese (Simplified)", "zh-TW": "Chinese (Traditional)",
        "co": "Corsican", "hr": "Croatian", "cs": "Czech", "da": "Danish",
        "nl": "Dutch", "en": "English", "eo": "Esperanto", "et": "Estonian",
        "fi": "Finnish", "fr": "French", "fy": "Frisian", "gl": "Galician",
        "ka": "Georgian", "de": "German", "el": "Greek", "gu": "Gujarati",
        "ht": "Haitian Creole", "ha": "Hausa", "haw": "Hawaiian", "he": "Hebrew",
        "hi": "Hindi", "hmn": "Hmong", "hu": "Hungarian", "is": "Icelandic",
        "ig": "Igbo", "id": "Indonesian", "ga": "Irish", "it": "Italian",
        "ja": "Japanese", "jw": "Javanese", "kn": "Kannada", "kk": "Kazakh",
        "km": "Khmer", "rw": "Kinyarwanda", "ko": "Korean", "ku": "Kurdish",
        "ky": "Kyrgyz", "lo": "Lao", "la": "Latin", "lv": "Latvian",
        "lt": "Lithuanian", "lb": "Luxembourgish", "mk": "Macedonian",
        "mg": "Malagasy", "ms": "Malay", "ml": "Malayalam", "mt": "Maltese",
        "mi": "Maori", "mr": "Marathi", "mn": "Mongolian", "my": "Myanmar (Burmese)",
        "ne": "Nepali", "no": "Norwegian", "ny": "Nyanja (Chichewa)",
        "or": "Odia (Oriya)", "ps": "Pashto", "fa": "Persian", "pl": "Polish",
        "pt": "Portuguese", "pa": "Punjabi", "ro": "Romanian", "ru": "Russian",
        "sm": "Samoan", "gd": "Scots Gaelic", "sr": "Serbian", "st": "Sesotho",
        "sn": "Shona", "sd": "Sindhi", "si": "Sinhala (Sinhalese)", "sk": "Slovak",
        "sl": "Slovenian", "so": "Somali", "es": "Spanish", "su": "Sundanese",
        "sw": "Swahili", "sv": "Swedish", "tl": "Tagalog (Filipino)", "tg": "Tajik",
        "ta": "Tamil", "tt": "Tatar", "te": "Telugu", "th": "Thai", "tr": "Turkish",
        "tk": "Turkmen", "uk": "Ukrainian", "ur": "Urdu", "ug": "Uyghur",
        "uz": "Uzbek", "vi": "Vietnamese", "cy": "Welsh", "xh": "Xhosa",
        "yi": "Yiddish", "yo": "Yoruba", "zu": "Zulu",
    }

    try:
        from deep_translator import GoogleTranslator

        langs = None
        getlangs = getattr(GoogleTranslator, "get_supported_languages", None)
        if callable(getlangs):
            try:
                langs = getlangs()
            except TypeError:
                try:
                    langs = GoogleTranslator().get_supported_languages()
                except Exception:
                    langs = None
        elif hasattr(GoogleTranslator, "SUPPORTED_LANGUAGES"):
            langs = getattr(GoogleTranslator, "SUPPORTED_LANGUAGES")

        if isinstance(langs, dict):
            first_key = next(iter(langs), "")
            if len(first_key) <= 5 and first_key.isascii() and first_key.islower():
                result = [(str(k), str(v).title()) for k, v in langs.items()]
            else:
                result = [(str(v), str(k).title()) for k, v in langs.items()]
            result.sort(key=lambda x: x[1].lower())
            _cached_language_options = result
            return result

        if isinstance(langs, (list, tuple)) and langs:
            result = []
            name_to_code = {v.lower(): k for k, v in fallback.items()}
            for entry in langs:
                entry = str(entry).strip()
                if not entry:
                    continue
                low = entry.lower()
                if low in name_to_code:
                    result.append((name_to_code[low], entry.title()))
                elif entry in fallback:
                    result.append((entry, fallback[entry]))
                else:
                    result.append((low, entry.title()))
            result.sort(key=lambda x: x[1].lower())
            _cached_language_options = result
            return result

    except Exception:
        logger.exception("Failed to probe deep_translator for supported languages")

    result = sorted(fallback.items(), key=lambda x: x[1].lower())
    _cached_language_options = result
    return result

# --- searchable dropdown ---

class SimpleComboBox:
    """Non-searchable dropdown — same visual style as SearchableComboBox but read-only.
    Used for small fixed option sets like the language detector picker."""

    _POPUP_ROWS = 8

    def __init__(self, parent, values, variable, command=None, **kw):
        p = theme.get()
        self._parent = parent
        self._all_values = list(values)
        self._variable = variable
        self._last_valid = variable.get() or (values[0] if values else "")
        self._popup = None
        self._suppress_open = False
        self._command = command

        # Outer frame — identical styling to SearchableComboBox
        self._frame = ctk.CTkFrame(
            parent,
            fg_color=p.bg_input,
            corner_radius=theme.BUTTON_CORNER_RADIUS,
            border_width=0,
            border_color=p.bg_input,
        ) if ctk else tk.Frame(parent, bd=0, highlightthickness=0)
        self._frame.grid_columnconfigure(0, weight=1)

        _font = ctk.CTkFont(family=theme.FONT_FAMILY, size=theme.FONT_BODY[1]) if ctk else None

        # Read-only label displaying the current value
        self._label_var = tk.StringVar(value=self._last_valid)
        if ctk:
            self._label = ctk.CTkLabel(
                self._frame,
                textvariable=self._label_var,
                font=_font,
                text_color=p.text_secondary,
                anchor="w",
                fg_color="transparent",
                height=32,
            )
        else:
            self._label = tk.Label(
                self._frame,
                textvariable=self._label_var,
                anchor="w",
                bg=p.bg_input,
                fg=p.text_secondary,
            )
        self._label.grid(row=0, column=0, sticky="ew", padx=(10, 0))

        # Arrow button — identical to SearchableComboBox
        arrow_img = get_icon("chevron-down", size=14)
        btn_kw = dict(
            master=self._frame,
            width=28, height=28,
            fg_color="transparent",
            hover_color=p.bg_input,
            corner_radius=4,
            command=self._on_arrow,
            border_width=0,
        )
        if arrow_img and ctk:
            self._btn = ctk.CTkButton(text="", image=arrow_img, **btn_kw)
        elif ctk:
            self._btn = ctk.CTkButton(
                text="\u25BC",
                font=ctk.CTkFont(family=theme.FONT_FAMILY, size=10),
                text_color=p.text_muted, **btn_kw,
            )
        else:
            self._btn = tk.Button(
                self._frame,
                text="\u25BC",
                command=self._on_arrow,
                bd=0, relief="flat", highlightthickness=0,
                background=p.bg_input,
                activebackground=p.bg_input,
                takefocus=0,
            )
        try:
            self._btn.configure(takefocus=False)
        except Exception:
            pass
        self._btn.grid(row=0, column=1, padx=(0, 2), pady=2)

        # Click on the label area also toggles the popup
        self._label.bind("<Button-1>", lambda _e: self._on_arrow())

        # Close popup when user clicks anywhere outside
        self._frame.after_idle(self._bind_root_click)

    # -- Geometry passthrough ------------------------------------------

    def grid(self, **kw):  self._frame.grid(**kw)
    def pack(self, **kw):  self._frame.pack(**kw)
    def place(self, **kw): self._frame.place(**kw)

    # -- Public API ----------------------------------------------------

    def get(self):
        return self._last_valid

    def set(self, value):
        if value in self._all_values:
            self._last_valid = value
            self._label_var.set(value)
            self._variable.set(value)

    def configure(self, **kw):
        # Accept state= so callers can disable/enable just like a normal widget
        state = kw.get("state")
        if state == "disabled":
            try:
                self._btn.configure(state="disabled")
                self._label.configure(state="disabled")
            except Exception:
                pass
        elif state == "normal":
            try:
                self._btn.configure(state="normal")
                self._label.configure(state="normal")
            except Exception:
                pass

    # -- Arrow / popup toggle ------------------------------------------

    def _on_arrow(self):
        if self._popup and self._popup.winfo_exists():
            self._close()
        else:
            self._open()

    # -- Root click detection ------------------------------------------

    def _bind_root_click(self):
        try:
            self._frame.winfo_toplevel().bind("<Button-1>", self._root_click, "+")
        except Exception:
            pass

    def _root_click(self, event):
        if not (self._popup and self._popup.winfo_exists()):
            return
        for container in (self._frame, self._popup):
            w = event.widget
            while w is not None:
                if w is container:
                    return
                w = getattr(w, "master", None)
        self._close()

    # -- Popup ---------------------------------------------------------

    def _open(self):
        if self._suppress_open:
            return
        if self._popup and self._popup.winfo_exists():
            return

        p = theme.get()
        self._popup = tk.Toplevel(self._frame)
        self._popup.wm_overrideredirect(True)
        self._popup.wm_attributes("-topmost", True)

        outer = tk.Frame(self._popup, bg=p.bg_popup, bd=0, highlightthickness=0)
        outer.pack(fill="both", expand=True)

        self._listbox = tk.Listbox(
            outer,
            height=min(self._POPUP_ROWS, len(self._all_values)),
            font=(theme.FONT_FAMILY, theme.FONT_BODY[1]),
            activestyle="none",
            selectbackground=p.accent,
            selectforeground=p.text_on_accent,
            bg=p.bg_popup, fg=p.text_secondary,
            relief="flat", borderwidth=0, highlightthickness=0,
        )
        self._listbox.pack(fill="both", expand=True, padx=4, pady=4)

        for item in self._all_values:
            self._listbox.insert(tk.END, item)

        if self._last_valid in self._all_values:
            idx = self._all_values.index(self._last_valid)
            self._listbox.selection_set(idx)
            self._listbox.see(idx)

        self._listbox.bind("<ButtonRelease-1>", self._on_select)
        self._listbox.bind("<Return>",          self._on_select)
        self._listbox.bind("<Escape>",          lambda _e: self._close())
        self._listbox.bind("<FocusOut>",        lambda _e: self._frame.after(100, self._check_focus))

        self._position_popup()

    def _position_popup(self):
        self._frame.update_idletasks()
        x = self._frame.winfo_rootx()
        y = self._frame.winfo_rooty() + self._frame.winfo_height() + 2
        w = self._frame.winfo_width()
        rows = min(self._POPUP_ROWS, max(1, len(self._all_values)))
        row_px = theme.scale(theme.FONT_BODY[1] + 10)
        h = rows * row_px + 8
        self._popup.geometry(f"{w}x{h}+{x}+{y}")

    def _on_select(self, _event=None):
        if not hasattr(self, "_listbox"):
            return
        sel = self._listbox.curselection()
        if not sel:
            return
        value = self._listbox.get(sel[0])
        self._last_valid = value
        self._label_var.set(value)
        self._variable.set(value)
        self._close(suppress_ms=150)
        if self._command:
            try:
                self._command(value)
            except Exception:
                pass

    def _check_focus(self):
        # Close only if focus has truly left both the frame and the popup
        try:
            focused = self._frame.focus_get()
            if focused and self._popup and self._popup.winfo_exists():
                if focused.winfo_toplevel() is self._popup:
                    return
        except Exception:
            pass
        self._close()

    def _close(self, suppress_ms=0):
        if self._popup and self._popup.winfo_exists():
            self._popup.destroy()
        self._popup = None
        if suppress_ms:
            self._suppress_open = True
            self._frame.after(suppress_ms, lambda: setattr(self, "_suppress_open", False))

class SearchableComboBox:
    _POPUP_ROWS = 8  # max visible rows before scrolling

    def __init__(self, parent, values, variable, **kw):
        p = theme.get()
        self._parent = parent
        self._all_values = list(values)
        self._variable = variable
        self._last_valid = variable.get() or (values[0] if values else "")
        self._popup = None
        self._suppress_open = False

        # Styled outer frame
        self._frame = ctk.CTkFrame(
            parent,
            fg_color=p.bg_input,
            corner_radius=theme.BUTTON_CORNER_RADIUS,
            border_width=0,
            border_color=p.bg_input,
        ) if ctk else tk.Frame(parent, bd=0, highlightthickness=0)
        self._frame.grid_columnconfigure(0, weight=1)

        _font = ctk.CTkFont(family=theme.FONT_FAMILY, size=theme.FONT_BODY[1]) if ctk else None

        # Entry (display + search input)
        self._var = tk.StringVar(value=self._last_valid)
        self._entry = ctk.CTkEntry(
            self._frame,
            textvariable=self._var,
            fg_color="transparent",
            border_width=0,
            font=_font,
            text_color=p.text_secondary,
            height=32,
        ) if ctk else tk.Entry(self._frame, textvariable=self._var)
        self._entry.grid(row=0, column=0, sticky="ew", padx=(6, 0))

        # Arrow button (borderless, non-focusable to avoid outline)
        arrow_img = get_icon("chevron-down", size=14)
        btn_kw = dict(
            master=self._frame,
            width=28, height=28,
            fg_color="transparent",
            hover_color=p.bg_input,
            corner_radius=4,
            command=self._on_arrow,
            border_width=0,
        )
        if arrow_img and ctk:
            self._btn = ctk.CTkButton(text="", image=arrow_img, **btn_kw)
        elif ctk:
            self._btn = ctk.CTkButton(
                text="\u25BC",
                font=ctk.CTkFont(family=theme.FONT_FAMILY, size=10),
                text_color=p.text_muted, **btn_kw,
            )
        else:
            self._btn = tk.Button(
                self._frame,
                text="\u25BC",
                command=self._on_arrow,
                bd=0,
                relief="flat",
                highlightthickness=0,
                background=p.bg_input,
                activebackground=p.bg_input,
                takefocus=0,
            )

        # Ensure button doesn't draw focus highlight
        try:
            self._btn.configure(takefocus=False)
        except Exception:
            pass

        self._btn.grid(row=0, column=1, padx=(0, 2), pady=2)

        # Get the inner tk.Entry for select_range / low-level bindings
        self._tk_entry = self._entry
        if ctk and isinstance(self._entry, ctk.CTkEntry):
            for child in self._entry.winfo_children():
                if isinstance(child, tk.Entry):
                    self._tk_entry = child
                    break

        self._tk_entry.bind("<FocusIn>",    self._on_focus_in)
        self._tk_entry.bind("<Button-1>",   self._on_click, "+")
        self._tk_entry.bind("<KeyRelease>", self._on_key)
        self._tk_entry.bind("<FocusOut>",   self._on_focus_out)
        self._tk_entry.bind("<Return>",     self._on_enter)
        self._tk_entry.bind("<Escape>",     self._on_escape)
        self._tk_entry.bind("<Down>",       self._focus_list)

        # Root-level click to close on non-focusable area clicks
        self._frame.after_idle(self._bind_root_click)

    # -- Geometry passthrough ------------------------------------------

    def grid(self, **kw):  self._frame.grid(**kw)
    def pack(self, **kw):  self._frame.pack(**kw)
    def place(self, **kw): self._frame.place(**kw)

    # -- Public API ----------------------------------------------------

    def get(self):
        return self._last_valid

    def set(self, value):
        self._last_valid = value
        self._var.set(value)
        self._variable.set(value)

    def update_values(self, values: list[str]) -> None:
        # Replace the option list; revert selection if current value no longer exists.
        self._all_values = list(values)
        if self._last_valid not in self._all_values:
            fallback = self._all_values[0] if self._all_values else ""
            self.set(fallback)
        if self._popup and self._popup.winfo_exists():
            self._close()
    

    def refresh_colors(self):
        pass

    # -- Entry event handlers ------------------------------------------

    def _on_focus_in(self, _event=None):
        # Tab/programmatic focus: select-all and open.
        if self._suppress_open:
            return
        if _event == "NotifyPointer":
            return  # mouse click handled by _on_click instead
        self._select_all()
        self._open()

    def _on_click(self, _event=None):
        # Mouse click on entry (may already have focus).
        if self._suppress_open:
            return
        self._select_all()
        self._open()

    def _on_key(self, event=None):
        # Live filter and open/refresh popup on each keystroke.
        if event and event.keysym in (
            "Shift_L", "Shift_R", "Control_L", "Control_R",
            "Alt_L", "Alt_R", "Caps_Lock", "Return", "Escape",
            "Up", "Down", "Left", "Right", "Tab",
        ):
            return
        query = self._var.get().lower()
        filtered = [v for v in self._all_values if query in v.lower()] if query else self._all_values
        display = filtered if filtered else self._all_values
        if self._popup and self._popup.winfo_exists():
            self._populate(display)
        else:
            self._open(display)

    def _on_focus_out(self, _event=None):
        # Delay so a listbox click can land before validation.
        self._frame.after(150, self._validate_or_revert)

    def _on_enter(self, _event=None):
        if self._popup and self._popup.winfo_exists() and hasattr(self, "_listbox"):
            sel = self._listbox.curselection()
            if sel:
                self._confirm(self._listbox.get(sel[0]))
            else:
                self._validate_or_revert()
        else:
            self._validate_or_revert()

    def _on_escape(self, _event=None):
        self._revert()

    def _focus_list(self, _event=None):
        if self._popup and self._popup.winfo_exists() and hasattr(self, "_listbox"):
            self._listbox.focus_set()
            if not self._listbox.curselection() and self._listbox.size():
                self._listbox.selection_set(0)
                self._listbox.activate(0)

    def _on_arrow(self):
        if self._popup and self._popup.winfo_exists():
            self._revert()
        else:
            self._open()
            self._select_all()

    # -- Root click detection ------------------------------------------

    def _bind_root_click(self):
        try:
            self._frame.winfo_toplevel().bind("<Button-1>", self._root_click, "+")
        except Exception:
            pass

    def _root_click(self, event):
        if not (self._popup and self._popup.winfo_exists()):
            return
        # If click is inside our frame OR popup -> leave open
        for container in (self._frame, self._popup):
            w = event.widget
            while w is not None:
                if w is container:
                    return
                w = getattr(w, "master", None)
        self._revert()

    # -- Popup ---------------------------------------------------------

    def _open(self, items=None):
        if items is None:
            items = self._all_values
        if self._popup and self._popup.winfo_exists():
            self._populate(items)
            return

        p = theme.get()
        self._popup = tk.Toplevel(self._frame)
        self._popup.wm_overrideredirect(True)
        self._popup.wm_attributes("-topmost", True)

        outer = tk.Frame(
            self._popup, bg=p.bg_popup,
            bd=0, highlightthickness=0,
        )
        outer.pack(fill="both", expand=True)

        scrollbar = tk.Scrollbar(outer, orient="vertical")
        self._listbox = tk.Listbox(
            outer,
            yscrollcommand=scrollbar.set,
            height=self._POPUP_ROWS,
            font=(theme.FONT_FAMILY, theme.FONT_BODY[1]),
            activestyle="none",
            selectbackground=p.accent,
            selectforeground=p.text_on_accent,
            bg=p.bg_popup, fg=p.text_secondary,
            relief="flat", borderwidth=0, highlightthickness=0,
        )
        scrollbar.config(command=self._listbox.yview)
        self._listbox.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=4)
        scrollbar.pack(side="right", fill="y", pady=4, padx=(0, 2))

        self._listbox.bind("<ButtonRelease-1>", self._on_list_select)
        self._listbox.bind("<Return>",          self._on_list_select)
        self._listbox.bind("<Escape>",          self._on_escape)
        self._listbox.bind("<FocusOut>",        self._on_focus_out)

        self._populate(items)

    def _populate(self, items):
        self._listbox.delete(0, tk.END)
        for item in items:
            self._listbox.insert(tk.END, item)
        if self._last_valid in items:
            idx = items.index(self._last_valid)
            self._listbox.selection_set(idx)
            self._listbox.see(idx)
        self._position_popup()

    def _position_popup(self):
        self._frame.update_idletasks()
        x = self._frame.winfo_rootx()
        y = self._frame.winfo_rooty() + self._frame.winfo_height() + 2
        w = self._frame.winfo_width()
        rows = min(self._POPUP_ROWS, max(1, self._listbox.size()))
        row_px = theme.scale(theme.FONT_BODY[1] + 10)
        h = rows * row_px + 8
        self._popup.geometry(f"{w}x{h}+{x}+{y}")

    # -- Selection / validation ----------------------------------------

    def _on_list_select(self, _event=None):
        if hasattr(self, "_listbox"):
            sel = self._listbox.curselection()
            if sel:
                self._confirm(self._listbox.get(sel[0]))

    def _confirm(self, value):
        self._last_valid = value
        self._var.set(value)
        self._variable.set(value)
        self._close(suppress_ms=150)
        try:
            self._frame.master.focus_set()
        except Exception:
            pass

    def _validate_or_revert(self):
        # Don't act if focus moved into the popup
        try:
            focused = self._frame.focus_get()
            if focused and self._popup and self._popup.winfo_exists():
                try:
                    if focused.winfo_toplevel() is self._popup:
                        return
                except Exception:
                    pass
        except Exception:
            pass
        current = self._var.get().strip()
        if current in self._all_values:
            self._last_valid = current
            self._variable.set(current)
            self._close()
        else:
            self._revert()

    def _revert(self):
        self._var.set(self._last_valid)
        self._variable.set(self._last_valid)
        self._close(suppress_ms=150)

    def _close(self, suppress_ms=0):
        if self._popup and self._popup.winfo_exists():
            self._popup.destroy()
        self._popup = None
        if suppress_ms:
            self._suppress_open = True
            self._frame.after(suppress_ms, lambda: setattr(self, "_suppress_open", False))

    # -- Helpers -------------------------------------------------------

    def _select_all(self):
        try:
            self._tk_entry.select_range(0, tk.END)
            self._tk_entry.icursor(tk.END)
        except Exception:
            pass


# --- main app ---

class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.worker = Worker()
        self.files: list[str] = []
        self.cfg = load_config() or {}
        self.total_files = 0
        self.completed_files = 0
        self._running = False

        # Mapping: treeview iid -> filepath
        self._tree_ids: dict[str, str] = {}
        # Mapping: filepath -> treeview iid
        self._file_to_iid: dict[str, str] = {}
        # Track per-file timing
        self._file_start_times: dict[str, float] = {}

        # Apply saved appearance mode (default: Light)
        saved_mode = self.cfg.get("appearance_mode", "Light")
        theme.set_mode(saved_mode)

        self._build_ui()

        def _reapply_icon():
            try:
                apply_window_icon(self.root, size=64)
            except Exception:
                pass

        # Schedule at multiple points to beat any late CTk icon re-application
        self.root.after(100, _reapply_icon)
        self.root.after(500, _reapply_icon)

        # Update check state (populated by background thread on startup)
        self._update_check_result: dict | None = None
        if self.cfg.get("auto_check_updates", True):
            self._run_update_check(startup=True)

        # Apply defaults from config
        default_out = self.cfg.get("default_output") or DEFAULT_OUTPUT_FOLDER
        self.output_entry.insert(0, default_out)
        default_in = self.cfg.get("default_input")
        
        if default_in:
            found = list_supported_files(default_in)
            for f in found:
                if f not in self.files:
                    self._add_file_to_table(f)

    # --- directory helpers ---

    def _set_button_disabled(self, button: object, disabled: bool) -> None:
        # Centralized enable/disable helper: sets state and ensures CTk disabled text color
        try:
            state = "disabled" if disabled else "normal"
            button.configure(state=state)
            # Do not override CTk disabled colour here; it is set at creation time
        except Exception:
            pass

    def _initialdir_for_input(self) -> str:
        try:
            if self.files:
                return str(Path(self.files[0]).parent.resolve())
        except Exception:
            pass
        if self.cfg.get("default_input"):
            val = self.cfg["default_input"]
            p = Path(val)
            return str(p.resolve() if p.is_absolute() else (Path.cwd() / p).resolve())
        return str(Path.cwd())

    def _initialdir_for_output(self) -> str:
        candidate = None
        try:
            if hasattr(self, "output_entry"):
                val = self.output_entry.get().strip()
                if val:
                    p = Path(val)
                    candidate = p.resolve() if p.is_absolute() else (Path.cwd() / p).resolve()
        except Exception:
            pass
        if candidate is None and self.cfg.get("default_output"):
            candidate = (Path.cwd() / self.cfg["default_output"]).resolve()
        if candidate is None:
            candidate = Path.cwd()
        # Walk up to the nearest existing ancestor so Windows doesn't fall back
        # to the last-used directory (which may be the input file's directory).
        while not candidate.exists() and candidate.parent != candidate:
            candidate = candidate.parent
        return str(candidate)

    # --- UI construction ---

    def _build_ui(self):
        p = theme.get()

        self.root.title("Verbilo")
        if isinstance(self.root, ctk.CTk):
            self.root.configure(fg_color=p.bg_main)
        else:
            self.root.configure(bg=p.bg_main)
        self.root.geometry(f"{theme.WINDOW_WIDTH}x{theme.WINDOW_HEIGHT}")
        try:
            self.root.minsize(theme.WINDOW_MIN_WIDTH, theme.WINDOW_MIN_HEIGHT)
            self.root.resizable(True, True)
        except Exception:
            pass

        self.root.grid_columnconfigure(0, weight=0, minsize=theme.scale(theme.SIDEBAR_WIDTH))
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

        # Sidebar
        self._build_sidebar()

        # Main content
        self._build_content()

    # --- sidebar ---

    def _build_sidebar(self):
        PAD = theme.PADDING  # CTk widgets self-scale padx/pady — do NOT pre-scale
        p = theme.get()

        self.sidebar = ctk.CTkFrame(
            self.root, width=theme.SIDEBAR_WIDTH,
            fg_color=p.bg_sidebar, corner_radius=0,
        )
        self.sidebar.grid(row=0, column=0, sticky="ns")
        self.sidebar.grid_propagate(False)
        self.sidebar.grid_columnconfigure(0, weight=1)

        row = 0

        # App title with icon
        title_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        title_frame.grid(row=row, column=0, sticky="ew", padx=PAD, pady=(PAD, PAD + 8))
        lang_icon = get_icon("language", size=24)
        if lang_icon:
            ctk.CTkLabel(title_frame, text="", image=lang_icon, width=24).pack(side=tk.LEFT, padx=(0, 8))
        theme.make_label(
            title_frame, "Verbilo", level="heading",
        ).pack(side=tk.LEFT)
        row += 1

        # TRANSLATION section
        theme.make_label(
            self.sidebar, "Translation", level="section",
        ).grid(row=row, column=0, sticky="w", padx=PAD, pady=(0, 6))
        row += 1

        # Source language
        self._source_lang_label = theme.make_label(
            self.sidebar, "Source language", level="small",
        )
        self._source_lang_label.grid(row=row, column=0, sticky="w", padx=PAD, pady=(4, 2))
        row += 1

        lang_opts = _get_language_options()
        self._lang_map = {f"{name} ({code})": code for code, name in lang_opts}
        display_values = list(self._lang_map.keys())
        if not display_values:
            display_values = ["English (en)"]

        # Source language: filtered to languages the initial detector (fasttext) can identify.
        _src_filtered = _filter_by_detector(lang_opts, "fasttext")
        _src_display = [f"{name} ({code})" for code, name in _src_filtered]
        self._source_lang_label.configure(text=f"Source language ({len(_src_display)})")
        self.source_lang_var = tk.StringVar(value="Auto-detect (translate all)")
        source_values = ["Auto-detect (translate all)"] + _src_display
        self._source_lang_map = {"Auto-detect (translate all)": "auto"}
        self._source_lang_map.update({k: v for k, v in self._lang_map.items() if k in set(_src_display)})

        self.source_lang_box = SearchableComboBox(
            self.sidebar, source_values, self.source_lang_var,
        )
        self.source_lang_box.grid(row=row, column=0, sticky="ew", padx=PAD, pady=(0, 6))
        row += 1

        # Target language
        theme.make_label(
            self.sidebar, "Target language", level="small",
        ).grid(row=row, column=0, sticky="w", padx=PAD, pady=(6, 2))
        row += 1

        default_target = "English (en)"
        if default_target not in display_values and display_values:
            default_target = display_values[0]
        self.lang_var = tk.StringVar(value=default_target)
        self.target_lang_box = SearchableComboBox(
            self.sidebar, display_values, self.lang_var,
        )
        self.target_lang_box.grid(row=row, column=0, sticky="ew", padx=PAD, pady=(0, 6))
        row += 1

        # Language detector
        theme.make_label(
            self.sidebar, "Language detector", level="small",
        ).grid(row=row, column=0, sticky="w", padx=PAD, pady=(6, 2))
        row += 1

        self.detector_var = tk.StringVar(value="fasttext")
        self.detector_menu = SimpleComboBox(
            self.sidebar,
            values=["fasttext", "lingua"],
            variable=self.detector_var,
            command=self._on_detector_changed,
        )
        self.detector_menu.grid(row=row, column=0, sticky="ew", padx=PAD, pady=(0, 6))
        row += 1

        # Translation engine
        theme.make_label(
            self.sidebar, "Translation engine", level="small",
        ).grid(row=row, column=0, sticky="w", padx=PAD, pady=(6, 2))
        row += 1

        saved_engine = self.cfg.get("translation_engine", "google")
        default_engine_display = _ENGINE_REVERSE.get(saved_engine, _ENGINE_DISPLAY[0])
        self.engine_var = tk.StringVar(value=default_engine_display)
        self.engine_menu = SimpleComboBox(
            self.sidebar,
            values=_ENGINE_DISPLAY,
            variable=self.engine_var,
            command=self._on_engine_changed,
        )
        self.engine_menu.grid(row=row, column=0, sticky="ew", padx=PAD, pady=(0, 6))
        row += 1

        # Divider
        theme.make_divider(self.sidebar).grid(
            row=row, column=0, sticky="ew", padx=PAD, pady=8,
        )
        row += 1

        # OUTPUT section
        theme.make_label(
            self.sidebar, "Output", level="section",
        ).grid(row=row, column=0, sticky="w", padx=PAD, pady=(0, 6))
        row += 1

        theme.make_label(
            self.sidebar, "Output folder", level="small",
        ).grid(row=row, column=0, sticky="w", padx=PAD, pady=(2, 2))
        row += 1

        self.output_entry = theme.make_entry(self.sidebar, height=32)
        self.output_entry.grid(row=row, column=0, sticky="ew", padx=PAD, pady=(0, 4))
        row += 1

        browse_icon = get_icon("folder", size=16, on_accent=False)
        theme.make_button(
            self.sidebar, "Browse", command=self._select_output, style="secondary",
            image=browse_icon, height=30,
        ).grid(row=row, column=0, sticky="ew", padx=PAD, pady=(0, 8))
        row += 1

        # Divider
        theme.make_divider(self.sidebar).grid(
            row=row, column=0, sticky="ew", padx=PAD, pady=4,
        )
        row += 1

        # Action buttons
        play_icon = get_icon("play", size=16, on_accent=True)
        self.start_btn = theme.make_button(
            self.sidebar, "Start Translation", command=self._start, style="primary",
            height=38, image=play_icon,
        )
        self.start_btn.grid(row=row, column=0, sticky="ew", padx=PAD, pady=(8, 4))
        row += 1

        stop_icon = get_icon("stop", size=16)
        self.cancel_btn = theme.make_button(
            self.sidebar, "Cancel", command=self._cancel, style="secondary",
            height=32, state="disabled", image=stop_icon,
        )
        self.cancel_btn.grid(row=row, column=0, sticky="ew", padx=PAD, pady=(0, 4))
        row += 1

        # Spacer row (pushes settings to bottom)
        self.sidebar.grid_rowconfigure(row, weight=1)
        row += 1

        # Settings at the very bottom
        theme.make_divider(self.sidebar).grid(
            row=row, column=0, sticky="ew", padx=PAD, pady=4,
        )
        row += 1

        settings_icon = get_icon("settings", size=16)
        theme.make_button(
            self.sidebar, "Settings", command=self._open_settings, style="ghost",
            anchor="w", image=settings_icon,
        ).grid(row=row, column=0, sticky="ew", padx=PAD, pady=(4, 4))
        row += 1

        info_icon = get_icon("info", size=16)
        about_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        about_frame.grid(row=row, column=0, sticky="ew", padx=PAD, pady=(0, PAD))
        about_frame.grid_columnconfigure(0, weight=1)
        theme.make_button(
            about_frame, "About", command=self._open_about, style="ghost",
            anchor="w", image=info_icon,
        ).grid(row=0, column=0, sticky="ew")
        theme.make_label(
            about_frame, "(beta)", level="tiny", text_color=p.text_muted,
        ).grid(row=0, column=1, sticky="e", padx=(8, 0))

    # --- content area ---

    def _build_content(self):
        PAD = theme.PADDING  # CTk widgets self-scale padx/pady — do NOT pre-scale
        p = theme.get()

        content = ctk.CTkFrame(self.root, fg_color=p.bg_main, corner_radius=0)
        content.grid(row=0, column=1, sticky="nsew", padx=PAD, pady=PAD)
        self._content = content

        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(0, weight=0)   # toolbar
        content.grid_rowconfigure(1, weight=3)   # file table
        content.grid_rowconfigure(2, weight=0)   # progress
        content.grid_rowconfigure(3, weight=1)   # log

        # Toolbar card
        toolbar = theme.make_card(content)
        toolbar.grid(row=0, column=0, sticky="ew", pady=(0, PAD))
        toolbar.grid_columnconfigure(0, weight=1)

        theme.make_label(toolbar, "Files", level="subheading").grid(
            row=0, column=0, sticky="w", padx=PAD, pady=8,
        )
        btn_frame = ctk.CTkFrame(toolbar, fg_color="transparent")
        btn_frame.grid(row=0, column=1, sticky="e", padx=PAD, pady=8)

        add_icon = get_icon("add-file", size=16, on_accent=True)
        folder_icon = get_icon("open-folder", size=16, on_accent=True)
        trash_icon = get_icon("trash", size=16)

        theme.make_button(
            btn_frame, "Add Files", command=self._add_files, style="primary",
            image=add_icon, height=30,
        ).pack(side=tk.LEFT, padx=(0, 6))
        theme.make_button(
            btn_frame, "Select Folder", command=self._select_folder, style="primary",
            image=folder_icon, height=30,
        ).pack(side=tk.LEFT, padx=(0, 6))
        theme.make_button(
            btn_frame, "Clear", command=self._clear_files, style="secondary",
            image=trash_icon, height=30,
        ).pack(side=tk.LEFT)

        # File table card
        table_card = theme.make_card(content)
        table_card.grid(row=1, column=0, sticky="nsew", pady=(0, PAD))
        self._build_file_table(table_card)

        # Progress card
        progress_card = theme.make_card(content)
        progress_card.grid(row=2, column=0, sticky="ew", pady=(0, PAD))
        progress_card.grid_columnconfigure(0, weight=1)

        self.progress_label = theme.make_label(
            progress_card, "Ready", level="small",
        )
        self.progress_label.grid(
            row=0, column=0, sticky="w", padx=PAD, pady=(10, 4),
        )

        self.progress = ctk.CTkProgressBar(
            progress_card,
            progress_color=p.accent,
            fg_color=p.bg_main,
            corner_radius=4,
            height=8,
        )
        self.progress.grid(row=1, column=0, sticky="ew", padx=PAD, pady=(0, 10))
        self._set_progress(0.0)

        # Log card
        log_card = theme.make_card(content)
        log_card.grid(row=3, column=0, sticky="nsew")
        log_card.grid_columnconfigure(0, weight=1)
        log_card.grid_rowconfigure(1, weight=1)

        theme.make_label(log_card, "Log", level="subheading").grid(
            row=0, column=0, sticky="w", padx=PAD, pady=(10, 4),
        )

        self.log = ctk.CTkTextbox(
            log_card,
            fg_color=p.bg_main,
            text_color=p.text_secondary,
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=theme.FONT_SMALL[1]),
            corner_radius=6,
            border_width=0,
            wrap="word",
            activate_scrollbars=True,
        )
        self.log.grid(row=1, column=0, sticky="nsew", padx=PAD, pady=(0, PAD))
        self.log.configure(state="disabled")  # read-only until we insert

    # --- file table ---

    def _build_file_table(self, parent):
        p = theme.get()
        style = ttk.Style()
        style.theme_use("clam")

        body_font = (theme.FONT_FAMILY, theme.FONT_BODY[1])
        heading_font = (theme.FONT_FAMILY, theme.FONT_SMALL[1], "bold")

        style.configure(
            "FileTable.Treeview",
            rowheight=32,
            font=body_font,
            background=p.bg_card,
            foreground=p.text_secondary,
            fieldbackground=p.bg_card,
            borderwidth=0,
        )
        style.configure(
            "FileTable.Treeview.Heading",
            font=heading_font,
            background=p.bg_heading,
            foreground=p.text_muted,
            borderwidth=0,
            relief="flat",
            padding=(8, 6),
        )
        style.map(
            "FileTable.Treeview",
            background=[("selected", p.accent)],
            foreground=[("selected", p.text_on_accent)],
        )
        style.map(
            "FileTable.Treeview.Heading",
            background=[("active", p.bg_input)],
        )
        # Slim, styled scrollbar
        style.configure(
            "Slim.Vertical.TScrollbar",
            gripcount=0,
            background=p.bg_card,
            darkcolor=p.bg_card,
            lightcolor=p.bg_card,
            troughcolor=p.bg_card,
            bordercolor=p.bg_card,
            arrowcolor=p.text_muted,
            relief="flat",
            borderwidth=0,
            arrowsize=12,
            width=10,
        )
        style.map(
            "Slim.Vertical.TScrollbar",
            background=[
                ("active",   p.border),
                ("!active",  p.divider),
                ("disabled", p.bg_card),
            ],
            arrowcolor=[("disabled", p.bg_card)],
        )

        # Load file-type icons for the treeview (PhotoImage for ttk)
        icon_color = p.text_muted
        self._file_icons: dict[str, object] = {}
        for ext, icon_name in ((".docx", "file-docx"), (".pdf", "file-pdf"),
                                (".xlsx", "file-xls"), (".xls", "file-xls")):
            img = get_photo_image(icon_name, size=18, color=icon_color)
            if img:
                self._file_icons[ext] = img
        # Fallback icon
        fallback = get_photo_image("file", size=18, color=icon_color)
        if fallback:
            self._file_icons["_default"] = fallback

        container = tk.Frame(parent, bg=p.bg_card)
        container.pack(
            fill=tk.BOTH, expand=True,
            # raw tk.Frame does NOT self-scale — must use scale() explicitly
            padx=theme.scale(theme.PADDING), pady=theme.scale(theme.PADDING),
        )

        self.file_table = ttk.Treeview(
            container,
            columns=("status", "time"),
            show="headings",
            style="FileTable.Treeview",
            selectmode="browse",
        )
        self.file_table.heading("status", text="Status", anchor="center")
        self.file_table.heading("time", text="Time", anchor="center")

        self.file_table["show"] = ("tree", "headings")
        self.file_table.heading("#0", text="File", anchor="w")
        self.file_table.column("#0", width=300, minwidth=150, stretch=True)
        self.file_table.column("status", width=100, minwidth=80, stretch=False, anchor="center")
        self.file_table.column("time", width=80, minwidth=60, stretch=False, anchor="center")

        scrollbar = ttk.Scrollbar(
            container, orient=tk.VERTICAL,
            command=self.file_table.yview,
            style="Slim.Vertical.TScrollbar",
        )
        self.file_table.configure(yscrollcommand=scrollbar.set)

        self.file_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 2))

        # Status colour tags
        self.file_table.tag_configure("pending",   foreground=p.status_pending)
        self.file_table.tag_configure("started",   foreground=p.status_info)
        self.file_table.tag_configure("finished",  foreground=p.status_success)
        self.file_table.tag_configure("error",     foreground=p.status_error)
        self.file_table.tag_configure("cancelled", foreground=p.status_warning)
        self.file_table.tag_configure("even", background=p.bg_row_even)
        self.file_table.tag_configure("odd",  background=p.bg_row_odd)

    def _get_file_icon(self, filepath: str):
        # Return the appropriate PhotoImage icon for a file extension.
        ext = Path(filepath).suffix.lower()
        return self._file_icons.get(ext, self._file_icons.get("_default"))

    # --- table helpers ---

    def _add_file_to_table(self, filepath: str, status: str = "pending"):
        self.files.append(filepath)
        name = os.path.basename(filepath)
        idx = len(self.files) - 1
        row_tag = "even" if idx % 2 == 0 else "odd"
        icon = self._get_file_icon(filepath)
        kw: dict[str, object] = {}
        if icon:
            kw["image"] = icon
        iid = self.file_table.insert(
            "", tk.END, text=f"  {name}", values=(status, ""), tags=(status, row_tag), **kw,
        )
        self._tree_ids[iid] = filepath
        self._file_to_iid[filepath] = iid

    def _update_file_status(self, filepath: str, status: str, elapsed: float | None = None):
        iid = self._file_to_iid.get(filepath)
        if iid is None:
            return
        name = os.path.basename(filepath)
        time_str = f"{elapsed:.1f}s" if elapsed is not None else ""
        idx = list(self._file_to_iid.keys()).index(filepath)
        row_tag = "even" if idx % 2 == 0 else "odd"
        self.file_table.item(iid, text=name, values=(status, time_str), tags=(status, row_tag))

    def _update_all_statuses(self, status: str):
        for filepath, iid in self._file_to_iid.items():
            name = os.path.basename(filepath)
            idx = list(self._file_to_iid.keys()).index(filepath)
            row_tag = "even" if idx % 2 == 0 else "odd"
            self.file_table.item(iid, text=name, values=(status, ""), tags=(status, row_tag))

    def _retag_rows(self):
        for idx, iid in enumerate(self.file_table.get_children()):
            current_tags = list(self.file_table.item(iid, "tags"))
            new_tags = [t for t in current_tags if t not in ("even", "odd")]
            new_tags.append("even" if idx % 2 == 0 else "odd")
            self.file_table.item(iid, tags=tuple(new_tags))

    # --- settings dialog ---

    def _open_settings(self):
        p = theme.get()
        PAD = theme.PADDING  # CTk widgets self-scale padx/pady — do NOT pre-scale

        win = ctk.CTkToplevel(self.root)
        win.wm_attributes("-alpha", 0)  # keep invisible until centered
        apply_window_icon(win)
        win.title("Settings")
        win.transient(self.root)
        win.grab_set()
        win.configure(fg_color=p.bg_main)

        win.grid_columnconfigure(0, weight=1)

        # Card wrapper
        card = theme.make_card(win)
        card.grid(row=0, column=0, sticky="nsew", padx=PAD, pady=PAD)
        card.grid_columnconfigure(1, weight=1)

        settings_icon = get_icon("settings", size=20)
        title_frame = ctk.CTkFrame(card, fg_color="transparent")
        title_frame.grid(row=0, column=0, columnspan=3, sticky="w", padx=PAD, pady=(PAD, PAD))
        if settings_icon:
            ctk.CTkLabel(title_frame, text="", image=settings_icon, width=20).pack(side=tk.LEFT, padx=(0, 8))
        theme.make_label(title_frame, "Settings", level="heading").pack(side=tk.LEFT)

        # Default input folder
        theme.make_label(card, "Default input folder", level="small").grid(
            row=1, column=0, sticky="w", padx=PAD, pady=(0, 6),
        )
        in_entry = theme.make_entry(card, height=32)
        in_entry.grid(row=1, column=1, sticky="ew", padx=4, pady=(0, 6))
        in_entry.insert(0, self.cfg.get("default_input", ""))

        def _browse_default_input():
            raw = in_entry.get().strip() or self.cfg.get("default_input") or ""
            if raw:
                _p = Path(raw)
                candidate = _p.resolve() if _p.is_absolute() else (Path.cwd() / _p).resolve()
            else:
                candidate = Path.cwd()
            while not candidate.exists() and candidate.parent != candidate:
                candidate = candidate.parent
            d = filedialog.askdirectory(title="Select default input folder", parent=win, initialdir=str(candidate))
            if d:
                in_entry.delete(0, tk.END)
                in_entry.insert(0, str(Path(d).resolve()))

        browse_icon_s = get_icon("folder", size=14)
        theme.make_button(card, "Browse", command=_browse_default_input, style="secondary",
                          image=browse_icon_s, height=28).grid(
            row=1, column=2, padx=(4, PAD), pady=(0, 6),
        )

        # Default output folder
        theme.make_label(card, "Default output folder", level="small").grid(
            row=2, column=0, sticky="w", padx=PAD, pady=(0, 6),
        )
        out_entry = theme.make_entry(card, height=32)
        out_entry.grid(row=2, column=1, sticky="ew", padx=4, pady=(0, 6))
        out_entry.insert(0, self.cfg.get("default_output", DEFAULT_OUTPUT_FOLDER))

        def _browse_default_output():
            raw = out_entry.get().strip() or self.cfg.get("default_output") or DEFAULT_OUTPUT_FOLDER
            if raw:
                _p = Path(raw)
                candidate = _p.resolve() if _p.is_absolute() else (Path.cwd() / _p).resolve()
            else:
                candidate = Path.cwd()
            while not candidate.exists() and candidate.parent != candidate:
                candidate = candidate.parent
            d = filedialog.askdirectory(title="Select default output folder", parent=win, initialdir=str(candidate))
            if d:
                out_entry.delete(0, tk.END)
                out_entry.insert(0, str(Path(d).resolve()))

        theme.make_button(card, "Browse", command=_browse_default_output, style="secondary",
                          image=browse_icon_s, height=28).grid(
            row=2, column=2, padx=(4, PAD), pady=(0, 6),
        )

        # Appearance mode toggle
        theme.make_label(card, "Appearance", level="small").grid(
            row=3, column=0, sticky="w", padx=PAD, pady=(8, 4),
        )

        mode_switch_var = tk.BooleanVar(value=(theme.get_mode() == "Dark"))
        mode_switch = ctk.CTkSwitch(
            card,
            text="Dark mode" if theme.get_mode() == "Dark" else "Light mode",
            variable=mode_switch_var,
            onvalue=True,
            offvalue=False,
            progress_color=p.accent,
            button_color=p.accent,
            button_hover_color=p.accent_hover,
            fg_color=p.bg_input,
            text_color=p.text_secondary,
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=theme.FONT_BODY[1]),
        )
        mode_switch.grid(row=3, column=1, sticky="w", padx=4, pady=(8, 4))

        def _on_mode_switch(*_):
            mode_switch.configure(text="Dark mode" if mode_switch_var.get() else "Light mode")

        mode_switch_var.trace_add("write", _on_mode_switch)

        # Restart note
        theme.make_label(
            card, "Appearance changes require a restart to take effect.",
            level="tiny",
        ).grid(row=4, column=0, columnspan=3, sticky="w", padx=PAD, pady=(0, 8))

        # --- Network section ---
        theme.make_divider(card).grid(row=5, column=0, columnspan=3, sticky="ew", padx=PAD, pady=(4, 8))

        theme.make_label(card, "NETWORK", level="section").grid(
            row=6, column=0, sticky="w", padx=PAD, pady=(0, 6),
        )

        theme.make_label(card, "HTTPS Proxy", level="small").grid(
            row=7, column=0, sticky="w", padx=PAD, pady=(0, 4),
        )
        proxy_entry = theme.make_entry(card, height=32)
        proxy_entry.grid(row=7, column=1, columnspan=2, sticky="ew", padx=(4, PAD), pady=(0, 4))
        proxy_entry.insert(0, self.cfg.get("proxy_url", ""))
        theme.make_label(
            card, "e.g. http://127.0.0.1:7890  —  also reads HTTPS_PROXY env var",
            level="tiny",
        ).grid(row=8, column=0, columnspan=3, sticky="w", padx=PAD, pady=(0, 8))

        # --- API Keys section ---
        theme.make_divider(card).grid(row=9, column=0, columnspan=3, sticky="ew", padx=PAD, pady=(4, 8))

        theme.make_label(card, "API KEYS", level="section").grid(
            row=10, column=0, sticky="w", padx=PAD, pady=(0, 6),
        )

        # Google Cloud API key
        theme.make_label(card, "Google Cloud API key", level="small").grid(
            row=11, column=0, sticky="w", padx=PAD, pady=(0, 4),
        )
        google_key_entry = theme.make_entry(card, height=32)
        google_key_entry.grid(row=11, column=1, columnspan=2, sticky="ew", padx=(4, PAD), pady=(0, 4))
        google_key_entry.insert(0, self.cfg.get("google_api_key", ""))
        google_key_entry.configure(show="•")
        theme.make_label(
            card, "Optional — enables \"Google Cloud API\" engine. Leave empty to use free Google Translate.",
            level="tiny",
        ).grid(row=12, column=0, columnspan=3, sticky="w", padx=PAD, pady=(0, 8))

        # Baidu API credentials
        theme.make_label(card, "Baidu App ID", level="small").grid(
            row=13, column=0, sticky="w", padx=PAD, pady=(0, 4),
        )
        baidu_id_entry = theme.make_entry(card, height=32)
        baidu_id_entry.grid(row=13, column=1, columnspan=2, sticky="ew", padx=(4, PAD), pady=(0, 4))
        baidu_id_entry.insert(0, self.cfg.get("baidu_appid", ""))

        theme.make_label(card, "Baidu App Key", level="small").grid(
            row=14, column=0, sticky="w", padx=PAD, pady=(0, 4),
        )
        baidu_key_entry = theme.make_entry(card, height=32)
        baidu_key_entry.grid(row=14, column=1, columnspan=2, sticky="ew", padx=(4, PAD), pady=(0, 4))
        baidu_key_entry.insert(0, self.cfg.get("baidu_appkey", ""))
        baidu_key_entry.configure(show="•")
        theme.make_label(
            card, "Get credentials at fanyi-api.baidu.com/choose — required for Baidu engine.",
            level="tiny",
        ).grid(row=15, column=0, columnspan=3, sticky="w", padx=PAD, pady=(0, 8))

        # --- Updates section ---
        theme.make_divider(card).grid(row=16, column=0, columnspan=3, sticky="ew", padx=PAD, pady=(4, 8))

        theme.make_label(card, "UPDATES", level="section").grid(
            row=17, column=0, sticky="w", padx=PAD, pady=(0, 6),
        )

        auto_updates_var = tk.BooleanVar(value=self.cfg.get("auto_check_updates", True))
        auto_updates_cb = ctk.CTkCheckBox(
            card,
            text="Automatically check for updates",
            variable=auto_updates_var,
            onvalue=True,
            offvalue=False,
            checkmark_color=p.bg_main,
            fg_color=p.accent,
            hover_color=p.accent_hover,
            border_color=p.border,
            text_color=p.text_secondary,
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=theme.FONT_BODY[1]),
        )
        auto_updates_cb.grid(row=17, column=1, columnspan=2, sticky="w", padx=4, pady=(0, 6))

        # --- DEBUG section ---
        theme.make_label(card, "DEBUG", level="section").grid(
            row=18, column=0, sticky="w", padx=PAD, pady=(0, 6),
        )

        debug_var = tk.BooleanVar(value=bool(self.cfg.get("debug_mode", False)))
        debug_cb = ctk.CTkCheckBox(
            card,
            text="Debug mode (show debug/info messages)",
            variable=debug_var,
            onvalue=True,
            offvalue=False,
            checkmark_color=p.bg_main,
            fg_color=p.accent,
            hover_color=p.accent_hover,
            border_color=p.border,
            text_color=p.text_secondary,
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=theme.FONT_BODY[1]),
        )
        # Place on its own row so it doesn't overlap the updates checkbox
        debug_cb.grid(row=18, column=1, columnspan=2, sticky="w", padx=4, pady=(8, 6))

        # Inline validation error
        self._settings_error = theme.make_label(
            card, "", level="tiny",
            text_color=p.status_error,
        )
        self._settings_error.grid(row=19, column=0, columnspan=3, sticky="w", padx=PAD, pady=(0, 2))

        # Button row
        btn_frame = ctk.CTkFrame(card, fg_color="transparent")
        btn_frame.grid(row=20, column=0, columnspan=3, pady=(4, PAD))

        def _save_and_close():
            inp = in_entry.get().strip()
            out = out_entry.get().strip()
            if not out:
                self._settings_error.configure(text="Input or Output path cannot be empty")
                return
            self._settings_error.configure(text="")
            self.cfg["default_input"] = inp
            self.cfg["default_output"] = out
            new_mode = "Dark" if mode_switch_var.get() else "Light"
            self.cfg["appearance_mode"] = new_mode
            self.cfg["auto_check_updates"] = auto_updates_var.get()
            self.cfg["debug_mode"] = debug_var.get()
            # Network & API keys
            self.cfg["proxy_url"] = proxy_entry.get().strip()
            self.cfg["google_api_key"] = google_key_entry.get().strip()
            self.cfg["baidu_appid"] = baidu_id_entry.get().strip()
            self.cfg["baidu_appkey"] = baidu_key_entry.get().strip()
            save_config(self.cfg)
            try:
                self._apply_debug_mode()
            except Exception:
                pass
            win.destroy()

        theme.make_button(btn_frame, "Save", command=_save_and_close, style="primary",
                          height=32).pack(
            side=tk.LEFT, padx=(0, 8),
        )
        theme.make_button(btn_frame, "Cancel", command=win.destroy, style="secondary",
                          height=32).pack(
            side=tk.LEFT,
        )

        # Title-bar X acts as Cancel (discard changes)
        win.protocol("WM_DELETE_WINDOW", win.destroy)
        win.update_idletasks()
        def _show_settings():
            center_window(win, parent=self.root)
            win.wm_attributes("-alpha", 1)
        win.after(20, _show_settings)
        try:
            win.resizable(False, False)
        except Exception:
            pass

    # --- update check ---

    def _run_update_check(self, startup: bool = False, callback=None):
        # Launch a background thread to check for the latest GitHub release.
        def _check():
            result: dict = {"status": "error", "version": None, "url": None}
            try:
                api_url = "https://api.github.com/repos/8041q/Verbilo/releases/latest"
                req = urllib.request.Request(
                    api_url, headers={"User-Agent": f"Verbilo/{APP_VERSION}"}
                )
                with urllib.request.urlopen(req, timeout=8) as resp:
                    data = json.loads(resp.read().decode())
                tag = data.get("tag_name", "").lstrip("v")
                html_url = data.get("html_url", RELEASES_URL)
                if tag and tag != APP_VERSION:
                    result = {"status": "update", "version": tag, "url": html_url}
                else:
                    result = {"status": "latest", "version": APP_VERSION, "url": None}
            except Exception:
                result = {"status": "error", "version": None, "url": None}

            self._update_check_result = result
            if callback:
                self.root.after(0, lambda r=result: callback(r))
            elif result["status"] == "update":
                delay = 100 if startup else 0
                self.root.after(delay, lambda r=result: self._show_update_dialog(r))

        threading.Thread(target=_check, daemon=True).start()

    def _show_update_dialog(self, result: dict):
        # Show a small dialog when a newer release is available.
        if result.get("status") != "update":
            return
        p = theme.get()
        PAD = theme.PADDING  # CTk widgets self-scale padx/pady — do NOT pre-scale

        dlg = ctk.CTkToplevel(self.root)
        dlg.wm_attributes("-alpha", 0)  # keep invisible until centered
        apply_window_icon(dlg)
        dlg.title("Update Available")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.configure(fg_color=p.bg_main)
        dlg.grid_columnconfigure(0, weight=1)

        card = theme.make_card(dlg)
        card.grid(row=0, column=0, sticky="nsew", padx=PAD, pady=PAD)
        card.grid_columnconfigure(0, weight=1)

        theme.make_label(
            card, f"A new version is available: v{result['version']}", level="subheading",
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=PAD, pady=(PAD, 4))
        theme.make_label(
            card, "Download the latest release from GitHub to update.", level="small",
        ).grid(row=1, column=0, columnspan=2, sticky="w", padx=PAD, pady=(0, PAD))

        dlg_btn_frame = ctk.CTkFrame(card, fg_color="transparent")
        dlg_btn_frame.grid(row=2, column=0, columnspan=2, pady=(0, PAD))

        def _download():
            webbrowser.open(result.get("url") or RELEASES_URL)
            dlg.destroy()

        theme.make_button(dlg_btn_frame, "Download", command=_download, style="primary",
                          height=32).pack(side=tk.LEFT, padx=(0, 8))
        theme.make_button(dlg_btn_frame, "Later", command=dlg.destroy, style="secondary",
                          height=32).pack(side=tk.LEFT)

        dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)
        dlg.update_idletasks()
        def _show_update():
            center_window(dlg, parent=self.root)
            dlg.wm_attributes("-alpha", 1)
        dlg.after(20, _show_update)
        try:
            dlg.resizable(False, False)
        except Exception:
            pass

    # --- about dialog ---

    def _open_about(self):
        # Open the standalone About dialog.
        p = theme.get()
        PAD = theme.PADDING  # CTk widgets self-scale padx/pady — do NOT pre-scale

        win = ctk.CTkToplevel(self.root)
        win.wm_attributes("-alpha", 0)  # keep invisible until centered
        apply_window_icon(win)
        win.title("About Verbilo")
        win.transient(self.root)
        win.grab_set()
        win.configure(fg_color=p.bg_main)
        win.grid_columnconfigure(0, weight=1)

        card = theme.make_card(win)
        card.grid(row=0, column=0, sticky="nsew", padx=PAD, pady=PAD)
        card.grid_columnconfigure(0, weight=1)

        # --- Brand area ---
        brand_frame = ctk.CTkFrame(card, fg_color="transparent")
        brand_frame.grid(row=0, column=0, sticky="ew", padx=PAD, pady=(PAD, 8))
        brand_frame.grid_columnconfigure(1, weight=1)

        try:
            pil_img = get_app_icon(size=256)
            if pil_img is not None and ctk is not None:
                logo_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(65, 65))
                ctk.CTkLabel(brand_frame, text="", image=logo_img, width=48).grid(
                    row=0, column=0, rowspan=2, padx=(0, 12),
                )
        except Exception:
            pass

        theme.make_label(brand_frame, "Verbilo", level="heading").grid(
            row=0, column=1, sticky="sw",
        )
        # Small muted beta badge beside the product heading
        theme.make_label(
            brand_frame, "Made by crt_ (8041q)", level="tiny", text_color=p.text_muted,
        ).grid(row=0, column=2, sticky="sw", padx=(6, 0))
        theme.make_label(brand_frame, "File translation tool", level="small").grid(
            row=1, column=1, sticky="nw",
        )

        # --- Version & date ---
        theme.make_label(card, f"Version   v{APP_VERSION}", level="body").grid(
            row=1, column=0, sticky="w", padx=PAD, pady=(0, 2),
        )
        theme.make_label(card, f"Build date   {APP_BUILD_DATE}", level="body").grid(
            row=2, column=0, sticky="w", padx=PAD, pady=(0, 2),
        )

        # --- Copyright ---
        theme.make_label(
            card, "\u00a9 2026 crt_ (8041q) \u2014  Released under the MIT License",
            level="tiny",
        ).grid(row=3, column=0, sticky="w", padx=PAD, pady=(0, PAD))

        # --- Divider ---
        theme.make_divider(card).grid(row=4, column=0, sticky="ew", padx=PAD, pady=(0, 8))

        # --- Check for updates ---
        check_frame = ctk.CTkFrame(card, fg_color="transparent")
        check_frame.grid(row=5, column=0, sticky="w", padx=PAD, pady=(0, 4))

        update_status_var = tk.StringVar(value="")

        def _do_check():
            update_status_var.set("Checking\u2026")
            self._set_button_disabled(check_btn, True)

            def _on_result(result):
                self._set_button_disabled(check_btn, False)
                if result["status"] == "update":
                    update_status_var.set(f"Update available: v{result['version']}")
                    self._show_update_dialog(result)
                elif result["status"] == "latest":
                    update_status_var.set("\u2713  You are up to date")
                else:
                    update_status_var.set("Could not check for updates.")

            self._run_update_check(startup=False, callback=_on_result)

        check_btn = theme.make_button(
            check_frame, "Check for Updates", command=_do_check, style="ghost", height=28,
        )
        check_btn.pack(side=tk.LEFT)

        update_status_lbl = theme.make_label(check_frame, "", level="small")
        update_status_lbl.pack(side=tk.LEFT, padx=(10, 0))

        def _sync_status(*_):
            val = update_status_var.get()
            if val.startswith("\u2713"):
                color = p.status_success
            elif "Update available" in val:
                color = p.text_secondary
            elif "Could not" in val:
                color = p.status_error
            else:
                color = p.text_muted
            update_status_lbl.configure(text=val, text_color=color)

        update_status_var.trace_add("write", _sync_status)

        # Surface result from an already-completed startup check
        if self._update_check_result:
            r = self._update_check_result
            if r["status"] == "latest":
                update_status_var.set("\u2713  You are up to date")
            elif r["status"] == "update":
                update_status_var.set(f"Update available: v{r['version']}")
            elif r["status"] == "error":
                update_status_var.set("Could not check for updates.")

        # --- Divider ---
        theme.make_divider(card).grid(row=6, column=0, sticky="ew", padx=PAD, pady=(4, 8))

        # --- Links ---
        links_frame = ctk.CTkFrame(card, fg_color="transparent")
        links_frame.grid(row=7, column=0, sticky="w", padx=PAD, pady=(0, PAD))

        def _open_github():
            webbrowser.open(GITHUB_URL)

        def _open_releases():
            webbrowser.open(RELEASES_URL)

        def _open_issue():
            webbrowser.open(ISSUE_URL)

        theme.make_button(links_frame, "View on GitHub", command=_open_github, style="ghost",
                          height=28).pack(side=tk.LEFT, padx=(0, theme.scale(8)))
        theme.make_button(links_frame, "Release Notes", command=_open_releases, style="ghost",
                          height=28).pack(side=tk.LEFT, padx=(0, theme.scale(8)))
        theme.make_button(links_frame, "Report a Bug", command=_open_issue, style="ghost",
                          height=28).pack(side=tk.LEFT, padx=(0, theme.scale(8)))

        win.protocol("WM_DELETE_WINDOW", win.destroy)
        win.update_idletasks()
        def _show_about():
            center_window(win, parent=self.root)
            win.wm_attributes("-alpha", 1)
        win.after(20, _show_about)
        try:
            win.resizable(False, False)
        except Exception:
            pass

    # --- file management ---

    def _add_files(self):
        init = self._initialdir_for_input()
        filetypes = [
            ("All supported", ("*.docx", "*.pdf", "*.xlsx", "*.xls")),
            ("Excel spreadsheets", ("*.xlsx", "*.xls")),
            ("Word documents", ("*.docx",)),
            ("PDF documents", ("*.pdf",)),
            ("All files", "*.*"),
        ]
        paths = filedialog.askopenfilenames(title="Select files", parent=self.root, initialdir=init, filetypes=filetypes)
        if not paths:
            return

        from pathlib import Path as _Path
        allowed = {ext.lower() for ext in SUPPORTED_EXTS}
        # normalize current files to resolved absolute strings to detect duplicates reliably
        existing = {str(_Path(p).resolve()) for p in self.files}
        invalid = []

        for p in paths:
            suf = _Path(p).suffix.lower()
            resolved = str(_Path(p).resolve())
            if suf in allowed:
                if resolved not in existing:
                    self._add_file_to_table(resolved)
                    existing.add(resolved)
            else:
                invalid.append(p)

        if invalid:
            messagebox.showwarning(
                "Unsupported files",
                "Some selected files were not supported and were ignored.\n\n"
                "Supported types: .docx, .pdf, .xlsx, .xls"
            )

    def _select_folder(self):
        init = self._initialdir_for_input()
        d = filedialog.askdirectory(title="Select folder containing files", parent=self.root, initialdir=init)
        if not d:
            return
        found = list_supported_files(d)
        if not found:
            messagebox.showinfo("No files", f"No supported files found in {d}")
            return
        
        from pathlib import Path as _Path
        existing = {str(_Path(p).resolve()) for p in self.files}
        for f in found:
            resolved = str(_Path(f).resolve())
            if resolved not in existing:
                self._add_file_to_table(resolved)
                existing.add(resolved)

    def _clear_files(self):
        # removes selected file, or clears all if nothing selected
        selected = self.file_table.selection()
        if selected:
            iid = selected[0]
            filepath = self._tree_ids.pop(iid, None)
            if filepath:
                self._file_to_iid.pop(filepath, None)
                if filepath in self.files:
                    self.files.remove(filepath)
            self.file_table.delete(iid)
            self._retag_rows()
        else:
            self.files.clear()
            self._tree_ids.clear()
            self._file_to_iid.clear()
            self._file_start_times.clear()
            for iid in self.file_table.get_children():
                self.file_table.delete(iid)
            self.completed_files = 0
            self.total_files = 0
            self._set_progress(0.0)
            self._update_progress_label("Ready")
            try:
                self.log.configure(state="normal")
                self.log.delete("1.0", "end")
                self.log.configure(state="disabled")
            except Exception:
                pass

    def _select_output(self):
        init = self._initialdir_for_output()
        d = filedialog.askdirectory(title="Select output folder", parent=self.root, initialdir=init)
        if d:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, _try_make_relative(d))

    def _on_detector_changed(self, *_) -> None:
        # Repopulate the source language dropdown to only show languages
        # the newly selected detector can actually identify.
        # Called by trace_add (args ignored) OR by CTkOptionMenu command=.
        self._refresh_language_dropdowns()

    def _on_engine_changed(self, *_) -> None:
        # Repopulate language dropdowns when the translation engine changes.
        self._refresh_language_dropdowns()
        engine_key = _ENGINE_MAP.get(self.engine_var.get(), "google")
        self.cfg["translation_engine"] = engine_key
        save_config(self.cfg)
        self._log(f"Translation engine changed to: {engine_key!r}")

    def _refresh_language_dropdowns(self) -> None:
        """Recompute source and target language lists based on current engine + detector."""
        detector = (self.detector_var.get() or "fasttext").strip().lower()
        engine_key = _ENGINE_MAP.get(self.engine_var.get(), "google")
        lang_opts = _get_language_options()

        # Target: filter by engine only (detector doesn't restrict target)
        tgt_filtered = _filter_by_engine(lang_opts, engine_key)
        tgt_display = [f"{name} ({code})" for code, name in tgt_filtered]
        self._lang_map = {f"{name} ({code})": code for code, name in tgt_filtered}
        self.target_lang_box.update_values(tgt_display)

        # Source: intersection of engine-supported and detector-supported
        src_filtered = _filter_by_engine(_filter_by_detector(lang_opts, detector), engine_key)
        src_display = [f"{name} ({code})" for code, name in src_filtered]
        src_display_set = set(src_display)
        source_values = ["Auto-detect (translate all)"] + src_display
        self._source_lang_map = {"Auto-detect (translate all)": "auto"}
        self._source_lang_map.update(
            {f"{name} ({code})": code for code, name in src_filtered}
        )
        self.source_lang_box.update_values(source_values)
        self._source_lang_label.configure(text=f"Source language ({len(src_display)})")
        self._log(f"Language lists updated for engine={engine_key!r}, detector={detector!r} "
                  f"(target: {len(tgt_display)}, source: {len(src_display)})")

    # --- progress helpers ---

    def _set_progress(self, fraction: float):
        try:
            if ctk and isinstance(self.progress, ctk.CTkProgressBar):
                self.progress.set(max(0.0, min(1.0, fraction)))
            else:
                self.progress["maximum"] = 100
                self.progress["value"] = max(0, min(100, fraction * 100))
        except Exception:
            pass

    def _update_progress_label(self, text: str):
        try:
            self.progress_label.configure(text=text)
        except Exception:
            pass

    # --- start / cancel ---

    def _start(self):
        if self._running:
            return

        # Resolve target language
        sel = self.lang_var.get()
        lang = self._lang_map.get(sel)
        if not lang:
            typed = self.target_lang_box.get()
            lang = self._lang_map.get(typed)
        if not lang:
            messagebox.showwarning("Missing language", "Please select a target language from the dropdown.")
            return
        if not self.files:
            messagebox.showwarning("No files", "Please add files or select a folder first.")
            return

        # Resolve source language
        source_sel = self.source_lang_var.get()
        source_lang = self._source_lang_map.get(source_sel)
        if not source_lang:
            typed = self.source_lang_box.get()
            source_lang = self._source_lang_map.get(typed, "auto")

        # Resolve output path; relative paths are resolved against cwd and auto-created
        output = self.output_entry.get().strip() or DEFAULT_OUTPUT_FOLDER
        out_path = Path(output)
        is_relative = not out_path.is_absolute()
        if is_relative:
            out_path = Path.cwd() / out_path
        if not out_path.exists():
            if is_relative:
                out_path.mkdir(parents=True, exist_ok=True)
            else:
                self._log(f"Error: Output path \u201c{output}\u201d does not exist \u2014 translation cancelled.")
                return
        output = str(out_path)

        # Resolve language detector selection
        detector = (self.detector_var.get() or "fasttext").strip().lower()
        if detector not in ("fasttext", "lingua"):
            detector = "fasttext"

        # Resolve translation engine and credentials from config
        engine = _ENGINE_MAP.get(self.engine_var.get(), "google")
        proxy_url = self.cfg.get("proxy_url", "").strip()
        proxies = {"https": proxy_url, "http": proxy_url} if proxy_url else None
        google_api_key = self.cfg.get("google_api_key", "")
        baidu_appid = self.cfg.get("baidu_appid", "")
        baidu_appkey = self.cfg.get("baidu_appkey", "")

        # Validate credentials for engines that require them
        if engine == "baidu" and (not baidu_appid or not baidu_appkey):
            messagebox.showwarning(
                "Missing credentials",
                "Baidu Translate requires an App ID and App Key.\n"
                "Please configure them in Settings → API Keys.",
            )
            return
        if engine == "google-cloud" and not google_api_key:
            messagebox.showwarning(
                "Missing API key",
                "Google Cloud API requires an API key.\n"
                "Please configure it in Settings → API Keys,\n"
                "or switch to \"Google Translate (free)\".",
            )
            return

        try:
            self._log(f"Starting: engine={engine!r}, source={source_lang!r}, target={lang!r}, detector={detector!r}")
        except Exception:
            pass

        # Update UI state
        self._running = True
        self._set_button_disabled(self.start_btn, True)
        self._set_button_disabled(self.cancel_btn, False)

        self._update_all_statuses("pending")
        self.total_files = len(self.files)
        self.completed_files = 0
        self._file_start_times.clear()
        self._set_progress(0.0)
        self._update_progress_label(f"0 / {self.total_files} files (0%)")

        self.worker.start(
            self.files, lang, output, None,
            self._progress_cb, self._log,
            source_lang=source_lang,
            detector=detector,
            engine=engine,
            proxies=proxies,
            google_api_key=google_api_key,
            baidu_appid=baidu_appid,
            baidu_appkey=baidu_appkey,
        )

    def _cancel(self):
        if not self._running:
            return
        self.worker.stop()
        self._set_button_disabled(self.cancel_btn, True)
        self._log("Cancelling\u2026")

    # --- callbacks (run on main thread) ---

    def _progress_cb(self, filepath: str, status: str, elapsed: float | None):
        def _update():
            import time as _time
            if status == "started":
                self._file_start_times[filepath] = _time.perf_counter()
                self._update_file_status(filepath, "translating\u2026")
            elif status in ("finished", "error"):
                self.completed_files += 1
                self._update_file_status(filepath, status, elapsed)
                pct = self.completed_files / max(1, self.total_files)
                self._set_progress(pct)
                self._update_progress_label(
                    f"{self.completed_files} / {self.total_files} files ({int(pct * 100)}%)",
                )
                if self.completed_files >= self.total_files:
                    self._finish_run()
            elif status == "cancelled":
                self._update_file_status(filepath, "cancelled")
                for f in self.files:
                    iid = self._file_to_iid.get(f)
                    if iid:
                        vals = self.file_table.item(iid, "values")
                        if vals and vals[0] == "pending":
                            self._update_file_status(f, "cancelled")
                self._finish_run(cancelled=True)

        self.root.after(0, _update)

    def _log(self, msg: str):
        # thread-safe log; "__worker_done__" signals the run is done
        if msg == "__worker_done__":
            self.root.after(0, lambda: self._finish_run() if self._running else None)
            return

        def append():
            self.log.configure(state="normal")
            self.log.insert("end", msg + "\n")
            self.log.see("end")
            self.log.configure(state="disabled")

        self.root.after(0, append)

    def _finish_run(self, cancelled: bool = False):
        if not self._running:
            return
        self._running = False
        self._set_button_disabled(self.start_btn, False)
        self._set_button_disabled(self.cancel_btn, True)
        if cancelled:
            pct = self.completed_files / max(1, self.total_files)
            self._update_progress_label(
                f"Cancelled \u2014 {self.completed_files} / {self.total_files} files ({int(pct * 100)}%)",
            )

    def _apply_debug_mode(self):
        # Apply the current debug mode immediately so logging reflects changes
        try:
            import logging as _logging

            debug = bool(self.cfg.get("debug_mode", False))
            root_logger = _logging.getLogger()

            try:
                root_logger.setLevel(_logging.DEBUG if debug else _logging.INFO)
            except Exception:
                pass

            for h in list(root_logger.handlers):
                try:
                    h.setLevel(_logging.DEBUG if debug else _logging.INFO)
                except Exception:
                    pass

            noisy_loggers = [
                "PIL", "PIL.PngImagePlugin", "PIL.Image", "PIL.ImageFile",
                "urllib3", "urllib3.connectionpool", "urllib3.util.retry",
                "requests", "http.client",
            ]
            for name in noisy_loggers:
                try:
                    _logging.getLogger(name).setLevel(_logging.WARNING)
                except Exception:
                    pass

            try:
                _logging.captureWarnings(True)
            except Exception:
                pass

        except Exception:
            try:
                logger.exception("Failed to apply debug mode")
            except Exception:
                pass

# --- entry point ---

def main():
    if ctk is None:
        tk.Tk().withdraw()
        messagebox.showerror(
            "Missing dependency",
            "customtkinter is required for GUI. Install with:\n\n"
            "    pip install customtkinter",
        )
        return


    root = ctk.CTk()
    root.wm_attributes("-alpha", 0)  # keep compositor-invisible during CTk init
    root.withdraw()  # hide until centered to avoid visible position jump
    theme.init_dpi(root)
    app = App(root)

    try:
        apply_window_icon(root, size=64)
    except Exception:
        pass

    try:
        import logging as _logging
        import re

        raw_log = app._log
        def _filtered_log(msg: str):
            try:
                if msg == "__worker_done__":
                    raw_log(msg)
                    return

                # If debug mode is disabled, skip noisy collected-info lines
                try:
                    debug_enabled = bool(app.cfg.get("debug_mode", False))
                except Exception:
                    debug_enabled = False

                if not debug_enabled and re.search(r"collected \d+ translatable string cells", msg, re.I):
                    return

                raw_log(msg)
            except Exception:
                try:
                    raw_log(msg)
                except Exception:
                    pass

        _handler = GuiLoggingHandler(
            _filtered_log,
            debug_getter=lambda: bool(app.cfg.get("debug_mode", False)),
        )
        _handler.setFormatter(_logging.Formatter("%(asctime)s %(levelname)s: %(message)s"))

        _handler.setLevel(_logging.DEBUG if app.cfg.get("debug_mode", False) else _logging.INFO)
        root_logger = _logging.getLogger()

        if not any(isinstance(h, GuiLoggingHandler) for h in root_logger.handlers):
            root_logger.addHandler(_handler)

        _logging.captureWarnings(True)
        app._apply_debug_mode()

        # Replace app._log so GUI's own log calls go through the same filter
        app._log = _filtered_log
    except Exception:
        logger.exception("Failed to install GUI logging handler")

    def _show():
        root.wm_attributes("-alpha", 1)
        try:
            center_window(root)
        except Exception:
            pass
        root.deiconify()

    root.after(10, _show)
    root.mainloop()

if __name__ == "__main__":
    main()
