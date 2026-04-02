# gui_customtk.py — dashboard GUI; sidebar (controls) + content (file table, progress, log)

from __future__ import annotations

import os
import json
import sys
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


def _get_app_root() -> Path:
    """Return the application root directory, aware of Nuitka frozen builds.

    In dev mode ``__file__`` lives at ``src/verbilo/gui/app.py`` so
    ``parents[3]`` reaches the repo root.  In a Nuitka standalone build the
    directory structure is flatter (``launch.dist/verbilo/gui/…``) so the same
    traversal overshoots.  For frozen builds we anchor on the executable's
    directory which is always the dist root.
    """
    if getattr(sys, "frozen", False) or "__compiled__" in globals():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[3]


def _try_make_relative(absolute_path: str) -> str:
    repo_root = _get_app_root()
    try:
        return str(Path(absolute_path).resolve().relative_to(repo_root))
    except Exception:
        try:
            return os.path.relpath(str(Path(absolute_path).resolve()), str(repo_root))
        except Exception:
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

# ISO 639-1 codes supported by Azure Translator.
_AZURE_LANG_CODES: frozenset[str] = frozenset({
    "af", "am", "ar", "as", "az", "ba", "bg", "bn", "bo", "bs",
    "ca", "cs", "cy", "da", "de", "el", "en", "es", "et", "eu",
    "fa", "fi", "fj", "fr", "ga", "gl", "gu", "he", "hi", "hr",
    "ht", "hu", "hy", "id", "ig", "is", "it", "ja", "ka", "kk",
    "km", "ko", "ku", "ky", "lo", "lt", "lv", "mg", "mi", "mk",
    "ml", "mn", "ms", "mt", "my", "ne", "nl", "no", "or", "pa",
    "pl", "pt", "ro", "ru", "sk", "sl", "sm", "sn", "so", "sq",
    "sr", "st", "sv", "sw", "ta", "te", "th", "ti", "tk", "tl",
    "tn", "tr", "tt", "ug", "uk", "ur", "uz", "vi", "xh", "yo",
    "zh", "zu", "zh-CN", "zh-TW",
})

# ISO 639-1 codes supported by DeepL
_DEEPL_LANG_CODES: frozenset[str] = frozenset({
    "ar", "bg", "cs", "da", "de", "el", "en", "es", "et", "fi",
    "fr", "hu", "id", "it", "ja", "ko", "lt", "lv", "nb", "no",
    "nl", "pl", "pt", "ro", "ru", "sk", "sl", "sv", "tr", "uk",
    "zh", "zh-CN", "zh-TW",
})

# Translation engine choices (display name → internal key).
_ENGINE_OPTIONS: list[tuple[str, str]] = [
    ("Google Translate (free)",  "google"),
    ("Google Cloud API (v2)",    "google-cloud"),
    ("Google Cloud API (v3)",    "google-cloud-v3"),
    ("Baidu Translate",          "baidu"),
    ("Microsoft Azure",          "azure"),
    ("DeepL",               "deepl"),
    ("Local (Offline)",          "local"),
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
    # Return only (code, name) pairs that the given translation engine supports. Google (free & cloud) supports all languages, so no filtering needed
    _ENGINE_CODES: dict[str, frozenset[str]] = {
        "baidu":           _BAIDU_LANG_CODES,
        "baidu-premium":   _BAIDU_LANG_CODES,
        "azure":           _AZURE_LANG_CODES,
        "deepl":           _DEEPL_LANG_CODES,
        "deepl-pro":       _DEEPL_LANG_CODES,
    }
    codes = _ENGINE_CODES.get(engine)
    if codes is None:
        return list(opts)  # google / google-cloud: no restriction
    result = []
    for code, name in opts:
        base = code.lower().split("_")[0]
        if base in codes or code in codes:
            result.append((code, name))
    return result


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
        "ceb": "Cebuano", "zh": "Chinese (Simplified)",
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


# --- OPUS-MT model helpers (local engine) ---------------------------------

# Map OPUS folder codes (2- or 3-letter) to ISO 639-1 codes used by the GUI.
_OPUS_CODE_MAP: dict[str, str] = {
    "eng": "en", "fra": "fr", "deu": "de", "spa": "es", "por": "pt",
    "ita": "it", "nld": "nl", "rus": "ru", "zho": "zh", "jpn": "ja",
    "jap": "ja", "kor": "ko", "ara": "ar", "pol": "pl", "tur": "tr",
    "swe": "sv", "dan": "da", "fin": "fi", "ukr": "uk", "ces": "cs",
    "ron": "ro", "hun": "hu", "nor": "no", "bul": "bg", "hrv": "hr",
    "ell": "el", "heb": "he", "hin": "hi", "tha": "th", "vie": "vi",
    "cat": "ca", "ind": "id", "msa": "ms", "slk": "sk", "slv": "sl",
    "est": "et", "lav": "lv", "lit": "lt", "srp": "sr",
}


def _opus_code_to_iso(code: str) -> str:
    # Convert an OPUS model code (e.g. ``eng``, ``fra``) to ISO 639-1
    return _OPUS_CODE_MAP.get(code, code)


def _get_default_model_dir() -> str:
    # Return the default OPUS-MT models directory (same logic as factory.py)
    return str(_get_app_root() / "models" / "opus-mt")


def _load_models_catalogue() -> list[dict]:
    # Load the bundled models catalogue from assets/models_catalogue.json
    cat_path = Path(__file__).resolve().parent.parent / "assets" / "models_catalogue.json"
    try:
        return json.loads(cat_path.read_text(encoding="utf-8"))
    except Exception:
        logger.warning("Could not load models catalogue from %s", cat_path)
        return []


def _get_local_model_dir_from_cfg(cfg: dict) -> str:
    return cfg.get("local_model_dir", "").strip() or _get_default_model_dir()


# --- searchable dropdown ---

class SimpleComboBox:
    # Non-searchable dropdown — same visual style as SearchableComboBox but read-only

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
        if self._suppress_open:
            return
        widget = getattr(_event, "widget", None) or self._tk_entry
        def _do():
            try:
                self._select_all()
                self._open()
            except Exception:
                pass
        try:
            widget.after_idle(_do)
        except Exception:
            _do()

    def _on_click(self, _event=None):
        if self._suppress_open:
            return
        widget = getattr(_event, "widget", None) or self._tk_entry
        try:
            widget.focus_set()
        except Exception:
            pass
        def _do():
            try:
                self._select_all()
                self._open()
            except Exception:
                pass
        try:
            widget.after_idle(_do)
        except Exception:
            _do()

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

        # Apply saved translation engine so language dropdowns reflect it
        try:
            self._on_engine_changed()
        except Exception:
            logger.exception("Failed to apply saved translation engine on startup")

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
        repo_root = _get_app_root()
        candidate = None
        try:
            if hasattr(self, "output_entry"):
                val = self.output_entry.get().strip()
                if val:
                    p = Path(val)
                    candidate = p.resolve() if p.is_absolute() else (repo_root / p).resolve()
        except Exception:
            pass
        if candidate is None and self.cfg.get("default_output"):
            candidate = (repo_root / self.cfg["default_output"]).resolve()
        if candidate is None:
            candidate = repo_root
        # Walk up to the nearest existing ancestor so the file dialog has a valid start dir.
        while not candidate.exists() and candidate.parent != candidate:
            candidate = candidate.parent
        return str(candidate)

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
        # Overlay badge for visual testing (non-clickable)
        try:
            title_badge = theme.make_label(title_frame, "(beta)", level="tiny", text_color=p.text_muted)
            # Place at the far right of the title row without affecting layout
            title_badge.place(relx=1.0, x=-8, rely=0.5, anchor="e")
            try:
                title_badge.configure(state="disabled")
            except Exception:
                pass
        except Exception:
            pass
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
        self.source_lang_box.grid(row=row, column=0, sticky="ew", padx=PAD, pady=(0, 2))
        row += 1

        # Listen for source-language changes (filters target list for local engine)
        self.source_lang_var.trace_add("write", self._on_source_lang_changed)

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
        self.engine_menu.grid(row=row, column=0, sticky="ew", padx=PAD, pady=(0, 2))
        row += 1

        # Usage label — shows monthly quota for engines that have one
        self._engine_usage_label = theme.make_label(self.sidebar, "", level="tiny")
        self._engine_usage_label_grid_kw = dict(row=row, column=0, sticky="w", padx=PAD, pady=(0, 6))
        self._engine_usage_label.grid(**self._engine_usage_label_grid_kw)
        row += 1
        self._update_usage_label(_ENGINE_MAP.get(default_engine_display, "google"))
      
        # Spacer row (pushes everything to bottom)
        self.sidebar.grid_rowconfigure(row, weight=1)
        row += 1

        # OUTPUT section
        theme.make_label(
            self.sidebar, "Output folder", level="small",
        ).grid(row=row, column=0, sticky="w", padx=PAD, pady=(2, 2))
        row += 1

        _out_frame = tk.Frame(self.sidebar, bg=p.bg_sidebar, bd=0, highlightthickness=0)
        _out_frame.columnconfigure(0, weight=1)
        _out_frame.columnconfigure(1, weight=0)
        _out_frame.grid(row=row, column=0, sticky="ew", padx=PAD, pady=(0, 8))
        row += 1

        self.output_entry = theme.make_entry(_out_frame, height=32)
        self.output_entry.grid(row=0, column=0, sticky="ew", pady=(0, 0))

        browse_icon = get_icon("folder", size=16, on_accent=False)
        theme.make_button(
            _out_frame, "", command=self._select_output, style="secondary",
            image=browse_icon, height=30, width=40,
        ).grid(row=0, column=1, sticky="ew", padx=(4, 0))

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
        about_btn = theme.make_button(
            about_frame, "About", command=self._open_about, style="ghost",
            anchor="w", image=info_icon,
        )
        about_btn.grid(row=0, column=0, sticky="ew")

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
            relief="flat",
            # "clam" theme draws its border via these color keys even when
            # borderwidth=0; setting them to the card background makes the
            # outline invisible without removing any padding.
            bordercolor=p.bg_card,
            lightcolor=p.bg_card,
            darkcolor=p.bg_card,
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
            background=[
                ("selected", p.accent),
                ("active", p.bg_card),
                ("!active", p.bg_card),
                ("focus", p.bg_card),
            ],
            foreground=[
                ("selected", p.text_on_accent),
                ("active", p.text_secondary),
                ("!active", p.text_secondary),
            ],
            bordercolor=[
                ("active", p.bg_card),
                ("focus", p.bg_card),
                ("!active", p.bg_card),
            ],
            lightcolor=[
                ("active", p.bg_card),
                ("focus", p.bg_card),
                ("!active", p.bg_card),
            ],
            darkcolor=[
                ("active", p.bg_card),
                ("focus", p.bg_card),
                ("!active", p.bg_card),
            ],
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
        # Suppress the Tk-level focus-ring outline (drawn outside ttk style)
        try:
            self.file_table.configure(takefocus=False)
            # highlightthickness is a raw Tk option not exposed by ttk but
            # accessible via the underlying tk widget call
            self.file_table.tk.call(str(self.file_table), "configure", "-highlightthickness", 0)
        except Exception:
            pass
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

        # Deselect when clicking on empty space in the table
        def _on_table_click(event):
            if not self.file_table.identify_row(event.y):
                self.file_table.selection_set([])
        self.file_table.bind("<Button-1>", _on_table_click, "+")

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

    def _format_elapsed_time(self, elapsed: float | None) -> str:
        if elapsed is None:
            return ""
        if elapsed < 60:
            return f"{elapsed:.1f}s"
        if elapsed < 3600:
            mins = int(elapsed // 60)
            secs = elapsed - mins * 60
            return f"{mins}m {secs:.0f}s" if secs >= 1 else f"{mins}m"
        hours = int(elapsed // 3600)
        mins = int((elapsed - hours * 3600) // 60)
        return f"{hours}h {mins}m" if mins else f"{hours}h"

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
        time_str = self._format_elapsed_time(elapsed)
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

        # Outer card wrapper — holds title + two-column body
        card = theme.make_card(win)
        card.configure(width=theme.scale(820))
        win.minsize(theme.scale(820), theme.scale(440))
        card.grid(row=0, column=0, sticky="nsew", padx=PAD, pady=PAD)
        card.grid_columnconfigure(0, weight=1)

        # ── Title row ────────────────────────────────────────────────────
        settings_icon = get_icon("settings", size=20)
        title_frame = ctk.CTkFrame(card, fg_color="transparent")
        title_frame.grid(row=0, column=0, sticky="w", padx=PAD, pady=(PAD, PAD))
        if settings_icon:
            ctk.CTkLabel(title_frame, text="", image=settings_icon, width=20).pack(side=tk.LEFT, padx=(0, 8))
        theme.make_label(title_frame, "Settings", level="heading").pack(side=tk.LEFT)

        # ── Two-column body ───────────────────────────────────────────────
        body = ctk.CTkFrame(card, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew", padx=PAD, pady=(0, PAD))
        # col 0 = left, col 1 = vertical divider (fixed width), col 2 = right
        body.grid_columnconfigure(0, weight=0, minsize=theme.scale(260))
        body.grid_columnconfigure(1, weight=0, minsize=theme.scale(1))
        body.grid_columnconfigure(2, weight=1)
        body.grid_rowconfigure(0, weight=0)

        # Vertical divider — use a plain tk.Frame so the fixed 1-px width is respected
        p_now = theme.get()
        _vdiv = tk.Frame(body, width=1, bg=p_now.divider)
        _vdiv.grid(row=0, column=1, sticky="ns", padx=(PAD, PAD))
        _vdiv.grid_propagate(False)

        # ── LEFT COLUMN: Folders + Appearance + Updates + Debug ──────────
        left = ctk.CTkFrame(body, fg_color="transparent")
        # Anchor left column to north-west and avoid vertical stretching
        left.grid(row=0, column=0, sticky="nw")
        
        left.grid_columnconfigure(0, weight=1)

        _lrow = 0

        # FOLDERS section
        theme.make_label(left, "FOLDERS", level="section").grid(
            row=_lrow, column=0, columnspan=2, sticky="w", pady=(0, 4),
        )
        _lrow += 1

        # Default input folder
        theme.make_label(left, "Default input folder", level="small").grid(
            row=_lrow, column=0, sticky="w", pady=(0, 2),
        )
        _lrow += 1
        _input_row = ctk.CTkFrame(left, fg_color="transparent")
        _input_row.grid(row=_lrow, column=0, columnspan=2, sticky="ew", pady=(0, 1))
        _input_row.grid_columnconfigure(0, weight=1)
        in_entry = theme.make_entry(_input_row, height=28)
        in_entry.grid(row=0, column=0, sticky="ew")
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
        theme.make_button(_input_row, "Browse", command=_browse_default_input, style="secondary",
                          image=browse_icon_s, height=28).grid(row=0, column=1, padx=(6, 0))
        _lrow += 1

        # Default output folder
        theme.make_label(left, "Default output folder", level="small").grid(
            row=_lrow, column=0, sticky="w", pady=(6, 4),
        )
        _lrow += 1
        _output_row = ctk.CTkFrame(left, fg_color="transparent")
        _output_row.grid(row=_lrow, column=0, columnspan=2, sticky="ew", pady=(0, 1))
        _output_row.grid_columnconfigure(0, weight=1)
        out_entry = theme.make_entry(_output_row, height=28)
        out_entry.grid(row=0, column=0, sticky="ew")
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

        theme.make_button(_output_row, "Browse", command=_browse_default_output, style="secondary",
                          image=browse_icon_s, height=28).grid(row=0, column=1, padx=(6, 0))
        _lrow += 1

        # Divider
        theme.make_divider(left).grid(row=_lrow, column=0, columnspan=2, sticky="ew", pady=(10, 8))
        _lrow += 1

        # APPEARANCE section
        theme.make_label(left, "APPEARANCE", level="section").grid(
            row=_lrow, column=0, columnspan=2, sticky="w", pady=(0, 4),
        )
        _lrow += 1

        mode_switch_var = tk.BooleanVar(value=(theme.get_mode() == "Dark"))
        mode_switch = ctk.CTkSwitch(
            left,
            text="Dark mode" if theme.get_mode() == "Dark" else "Light mode",
            variable=mode_switch_var,
            onvalue=True,
            offvalue=False,
            progress_color=p.accent,
            button_color=p.accent,
            button_hover_color=p.accent_hover,
            fg_color=p.divider,
            text_color=p.text_secondary,
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=theme.FONT_BODY[1]),
        )
        mode_switch.grid(row=_lrow, column=0, columnspan=2, sticky="w", pady=(0, 2))
        _lrow += 1

        def _on_mode_switch(*_):
            mode_switch.configure(text="Dark mode" if mode_switch_var.get() else "Light mode")

        mode_switch_var.trace_add("write", _on_mode_switch)

        theme.make_label(
            left, "Appearance changes require a restart to take effect.",
            level="tiny",
        ).grid(row=_lrow, column=0, columnspan=2, sticky="w", pady=(0, 2))
        _lrow += 1

        # Divider
        theme.make_divider(left).grid(row=_lrow, column=0, columnspan=2, sticky="ew", pady=(10, 8))
        _lrow += 1

        # UPDATES section
        theme.make_label(left, "UPDATES", level="section").grid(
            row=_lrow, column=0, columnspan=2, sticky="w", pady=(0, 4),
        )
        _lrow += 1

        auto_updates_var = tk.BooleanVar(value=self.cfg.get("auto_check_updates", True))
        auto_updates_cb = ctk.CTkCheckBox(
            left,
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
        auto_updates_cb.grid(row=_lrow, column=0, columnspan=2, sticky="w", pady=(0, 4))
        _lrow += 1

        # Divider
        theme.make_divider(left).grid(row=_lrow, column=0, columnspan=2, sticky="ew", pady=(10, 8))
        _lrow += 1

        # DEBUG section
        theme.make_label(left, "DEBUG", level="section").grid(
            row=_lrow, column=0, columnspan=2, sticky="w", pady=(0, 4),
        )
        _lrow += 1

        debug_var = tk.BooleanVar(value=bool(self.cfg.get("debug_mode", False)))
        debug_cb = ctk.CTkCheckBox(
            left,
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
        debug_cb.grid(row=_lrow, column=0, columnspan=2, sticky="w", pady=(0, 4))
        _lrow += 1

        # Divider
        theme.make_divider(left).grid(row=_lrow, column=0, columnspan=2, sticky="ew", pady=(10, 8))
        _lrow += 1

        # LOCAL MODELS section
        theme.make_label(left, "LOCAL MODELS", level="section").grid(
            row=_lrow, column=0, columnspan=2, sticky="w", pady=(0, 4),
        )
        _lrow += 1

        def _open_manager_from_settings():
            win.destroy()
            self.root.after(50, self._open_model_manager)

        theme.make_button(
            left, "Open Model Manager",
            command=_open_manager_from_settings,
            style="secondary", height=28,
        ).grid(row=_lrow, column=0, columnspan=2, sticky="w", pady=(0, 4))
        _lrow += 1

        # ── RIGHT COLUMN: Network + API Keys ─────────────────────────────
        # Use a scrollable frame for the right column
        try:
            right = ctk.CTkScrollableFrame(body, fg_color="transparent")
            try:
                right.configure(height=theme.scale(350))
            except Exception:
                pass
        except Exception:
            right = ctk.CTkFrame(body, fg_color="transparent")
        right.grid(row=0, column=2, sticky="nsew")
        right.grid_columnconfigure(0, weight=1)

        _rrow = 0

        # NETWORK section
        theme.make_label(right, "NETWORK", level="section").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 6),
        )
        _rrow += 1

        theme.make_label(right, "HTTPS Proxy", level="small").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 4),
        )
        _rrow += 1
        proxy_entry = theme.make_entry(right, height=28)
        proxy_entry.grid(row=_rrow, column=0, sticky="ew", pady=(0, 1))
        proxy_entry.insert(0, self.cfg.get("proxy_url", ""))
        _rrow += 1
        theme.make_label(
            right, "e.g. http://127.0.0.1:7890  —  also reads HTTPS_PROXY env var",
            level="tiny",
        ).grid(row=_rrow, column=0, sticky="w", pady=(0, 8))
        _rrow += 1

        # Divider
        theme.make_divider(right).grid(row=_rrow, column=0, sticky="ew", pady=(4, 8))
        _rrow += 1

        # Google Cloud section
        theme.make_label(right, "GOOGLE CLOUD", level="section").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 6),
        )
        _rrow += 1

        # Google Cloud API key
        theme.make_label(right, "Google Cloud API key (v2)", level="small").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 4),
        )
        _rrow += 1
        google_key_entry = theme.make_entry(right, height=28)
        google_key_entry.grid(row=_rrow, column=0, sticky="ew", pady=(0, 1))
        google_key_entry.insert(0, self.cfg.get("google_api_key", ""))
        google_key_entry.configure(show="•")
        _rrow += 1

        lbl1 = theme.make_label(
            right, "For Google Cloud API (v2). Get key at cloud.google.com/translate\n500K chars/month free",
            level="tiny",
        )
        lbl1.configure(anchor="w", justify="left")
        lbl1.grid(row=_rrow, column=0, sticky="w", pady=(0, 8))
        _rrow += 1

        # Google Cloud v3 Project ID
        theme.make_label(right, "Google Cloud Project ID (v3)", level="small").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 2),
        )
        _rrow += 1
        google_project_entry = theme.make_entry(right, height=28)
        google_project_entry.grid(row=_rrow, column=0, sticky="ew", pady=(0, 2))
        google_project_entry.insert(0, self.cfg.get("google_project_id", ""))
        _rrow += 1

        # Google Cloud v3 Service Account JSON
        theme.make_label(right, "Service Account JSON key (v3)", level="small").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 2),
        )
        _rrow += 1
        google_sa_entry = theme.make_entry(right, height=28)
        google_sa_entry.grid(row=_rrow, column=0, sticky="ew", pady=(0, 1))
        google_sa_entry.insert(0, self.cfg.get("google_sa_json", ""))
        _rrow += 1
        lbl_v3 = theme.make_label(
            right, "For Google Cloud API (v3). Project ID required; SA JSON = file path or raw JSON.\nLeave SA JSON blank to use Application Default Credentials.",
            level="tiny",
        )
        lbl_v3.configure(anchor="w", justify="left")
        lbl_v3.grid(row=_rrow, column=0, sticky="w", pady=(0, 8))
        _rrow += 1

        # Divider
        theme.make_divider(right).grid(row=_rrow, column=0, sticky="ew", pady=(4, 8))
        _rrow += 1

        # Baidu section
        theme.make_label(right, "BAIDU TRANSLATE", level="section").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 6),
        )
        _rrow += 1

        # Baidu App ID
        theme.make_label(right, "Baidu App ID", level="small").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 2),
        )
        _rrow += 1
        baidu_id_entry = theme.make_entry(right, height=28)
        baidu_id_entry.grid(row=_rrow, column=0, sticky="ew", pady=(0, 2))
        baidu_id_entry.insert(0, self.cfg.get("baidu_appid", ""))
        _rrow += 1

        # Baidu App Key
        theme.make_label(right, "Baidu App Key", level="small").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 2),
        )
        _rrow += 1
        baidu_key_entry = theme.make_entry(right, height=28)
        baidu_key_entry.grid(row=_rrow, column=0, sticky="ew", pady=(0, 1))
        baidu_key_entry.insert(0, self.cfg.get("baidu_appkey", ""))
        baidu_key_entry.configure(show="•")
        _rrow += 1

        lbl = theme.make_label(
            right, "Get credentials at fanyi-api.baidu.com/choose -> 通用文本翻译 / 文本翻译 API\nStandard: 50K chars/month free (1 req/s QPS limit)\nPremium: no char limit, higher QPS (requires approved account)",
            level="tiny",
        )
        lbl.configure(anchor="w", justify="left")
        lbl.grid(row=_rrow, column=0, sticky="w", pady=(0, 4))
        _rrow += 1

        # Baidu API tier selection
        theme.make_label(right, "API Tier", level="small").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 4),
        )
        _rrow += 1
        baidu_tier_var = ctk.StringVar(value=self.cfg.get("baidu_tier", "standard"))

        _tier_frame = ctk.CTkFrame(right, fg_color="transparent")
        _tier_frame.grid(row=_rrow, column=0, sticky="w", pady=(0, 8))

        def _make_tier_btn(parent, label, value):
            def _select():
                baidu_tier_var.set(value)
                _refresh_tier_btns()
            btn = ctk.CTkButton(
                parent, text=label, width=90, height=28,
                corner_radius=theme.BUTTON_CORNER_RADIUS,
                border_width=1,
                font=ctk.CTkFont(family=theme.FONT_FAMILY, size=theme.FONT_SMALL[1]),
                command=_select,
            )
            btn.pack(side=tk.LEFT, padx=(0, 6))
            return btn

        _tier_btn_standard = _make_tier_btn(_tier_frame, "Standard", "standard")
        _tier_btn_premium  = _make_tier_btn(_tier_frame, "Premium",  "premium")

        def _refresh_tier_btns():
            p_now = theme.get()
            selected = baidu_tier_var.get()
            for btn, val in ((_tier_btn_standard, "standard"), (_tier_btn_premium, "premium")):
                if val == selected:
                    btn.configure(
                        fg_color=p_now.accent,
                        hover_color=p_now.accent_hover,
                        text_color=p_now.text_on_accent,
                        border_color=p_now.accent_pressed,
                    )
                else:
                    btn.configure(
                        fg_color="transparent",
                        hover_color=p_now.bg_card,
                        text_color=p_now.text_secondary,
                        border_color=p_now.border,
                    )

        _refresh_tier_btns()
        _rrow += 1

        # Divider
        theme.make_divider(right).grid(row=_rrow, column=0, sticky="ew", pady=(4, 8))
        _rrow += 1

        # Azure section
        theme.make_label(right, "AZURE TRANSLATOR", level="section").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 6),
        )
        _rrow += 1

        theme.make_label(right, "Subscription Key", level="small").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 2),
        )
        _rrow += 1
        azure_key_entry = theme.make_entry(right, height=28)
        azure_key_entry.grid(row=_rrow, column=0, sticky="ew", pady=(0, 2))
        azure_key_entry.insert(0, self.cfg.get("azure_key", ""))
        azure_key_entry.configure(show="•")
        _rrow += 1

        theme.make_label(right, "Region (e.g. eastus)", level="small").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 2),
        )
        _rrow += 1
        azure_region_entry = theme.make_entry(right, height=28)
        azure_region_entry.grid(row=_rrow, column=0, sticky="ew", pady=(0, 1))
        azure_region_entry.insert(0, self.cfg.get("azure_region", ""))
        _rrow += 1
        lbl_az = theme.make_label(
            right, "Get key at portal.azure.com → Cognitive Services → Translator\n2M chars free/month",
            level="tiny",
        )
        lbl_az.configure(anchor="w", justify="left")
        lbl_az.grid(row=_rrow, column=0, sticky="w", pady=(0, 8))
        _rrow += 1

        # Divider
        theme.make_divider(right).grid(row=_rrow, column=0, sticky="ew", pady=(4, 8))
        _rrow += 1

        # DeepL section
        theme.make_label(right, "DeepL", level="section").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 6),
        )
        _rrow += 1

        theme.make_label(right, "DeepL API Key", level="small").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 2),
        )
        _rrow += 1
        deepl_key_entry = theme.make_entry(right, height=28)
        deepl_key_entry.grid(row=_rrow, column=0, sticky="ew", pady=(0, 1))
        deepl_key_entry.insert(0, self.cfg.get("deepl_api_key", ""))
        deepl_key_entry.configure(show="•")
        _rrow += 1
        lbl_dl = theme.make_label(
            right, "Get key at deepl.com/pro-api (free tier)\n500K chars free/month",
            level="tiny",
        )
        lbl_dl.configure(anchor="w", justify="left")
        lbl_dl.grid(row=_rrow, column=0, sticky="w", pady=(0, 8))
        _rrow += 1

        # Divider
        theme.make_divider(right).grid(row=_rrow, column=0, sticky="ew", pady=(4, 8))
        _rrow += 1

        # Usage summary + cache controls
        theme.make_label(right, "USAGE THIS MONTH", level="section").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 4),
        )
        _rrow += 1

        def _get_usage_text() -> str:
            try:
                from ..translators.usage import get_tracker
                t = get_tracker()
                lines: list[str] = []
                for display_name, eng_key in _ENGINE_OPTIONS:
                    usage_str = t.format_usage(eng_key)
                    if usage_str:
                        lines.append(f"{display_name}: {usage_str}")
                return "\n".join(lines) if lines else "No usage data yet."
            except Exception:
                return ""

        usage_text = _get_usage_text()
        usage_lbl = theme.make_label(right, usage_text or "No usage data yet.", level="tiny")
        usage_lbl.configure(anchor="w", justify="left")
        usage_lbl.grid(row=_rrow, column=0, sticky="w", pady=(0, 6))
        _rrow += 1

        # Create a small row with the button and a ghost label beside it
        try:
            # compute initial text (always visible, even when 0)
            from ..translators.cache import get_cache
            from ..utils.io import format_bytes
            n = get_cache().size()
            b = get_cache().disk_usage_bytes()
            initial_lbl = f"({n:,} entries, {format_bytes(b)})"
        except Exception:
            initial_lbl = "(0 entries, 0 B)"

        _cache_row = ctk.CTkFrame(right, fg_color="transparent")
        _cache_row.grid(row=_rrow, column=0, sticky="w", pady=(0, 4))

        cache_lbl = theme.make_label(_cache_row, initial_lbl, level="tiny")
        cache_lbl.configure(anchor="w", justify="left", text_color=p.text_secondary)

        # Helper to compute and update the cache info label (entries + human-readable size)
        def _update_cache_label():
            try:
                from ..translators.cache import get_cache
                from ..utils.io import format_bytes
                n = get_cache().size()
                b = get_cache().disk_usage_bytes()
                txt = f"({n:,} entries, {format_bytes(b)})"
            except Exception:
                txt = "(0 entries, 0 B)"
            try:
                cache_lbl.configure(text=txt)
                cache_lbl.update_idletasks()
            except Exception:
                pass

        def _clear_cache():
            try:
                from ..translators.cache import get_cache
                get_cache().clear()
                self._log("Translation cache cleared.")
            except Exception as e:
                self._log(f"Error clearing cache: {e}")
            try:
                _update_cache_label()
            except Exception:
                pass

        clear_btn = theme.make_button(
            _cache_row, "Clear translation cache",
            command=_clear_cache, style="secondary", height=26,
        )
        clear_btn.pack(side=tk.LEFT)
        cache_lbl.pack(side=tk.LEFT, padx=(8, 0))

        # ── Bottom row: error label + buttons ────────────────────────────
        bottom = ctk.CTkFrame(card, fg_color="transparent")
        bottom.grid(row=2, column=0, sticky="ew", padx=PAD, pady=(0, PAD//2))
        bottom.grid_columnconfigure(0, weight=1)

        # Inline validation error (spans full width)
        self._settings_error = theme.make_label(
            bottom, "", level="tiny",
            text_color=p.status_error,
        )
        self._settings_error.grid(row=1, column=2, sticky="w", pady=(0, 6))

        # Button row
        btn_frame = ctk.CTkFrame(bottom, fg_color="transparent")
        btn_frame.grid(row=1, column=0, sticky="w", pady=(0, 0))

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
            self.cfg["baidu_tier"] = baidu_tier_var.get()
            self.cfg["google_project_id"] = google_project_entry.get().strip()
            self.cfg["google_sa_json"] = google_sa_entry.get().strip()
            self.cfg["azure_key"] = azure_key_entry.get().strip()
            self.cfg["azure_region"] = azure_region_entry.get().strip()
            self.cfg["deepl_api_key"] = deepl_key_entry.get().strip()
            save_config(self.cfg)

            # Refresh engine-dependent UI after config changes
            try:
                self._refresh_language_dropdowns()
                self._update_usage_label(_ENGINE_MAP.get(self.engine_var.get(), "google"))
            except Exception:
                logger.exception("Failed to refresh language dropdowns after saving settings")

            try:
                self._apply_debug_mode()
            except Exception:
                pass
            win.destroy()

        theme.make_button(btn_frame, "Save", command=_save_and_close, style="primary",
                          height=28).pack(
            side=tk.LEFT, padx=(0, 6),
        )
        theme.make_button(btn_frame, "Cancel", command=win.destroy, style="secondary",
                          height=28).pack(
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

    # --- model manager ---

    _model_manager_open = False

    def _open_model_manager(self):
        # Open the Local Model Manager as a modal window
        import queue
        import shutil
        import subprocess
        import threading

        if self._model_manager_open:
            return
        self._model_manager_open = True

        from ..translators.local import list_downloaded_pairs

        p = theme.get()
        PAD = theme.PADDING

        win = ctk.CTkToplevel(self.root)
        win.wm_attributes("-alpha", 0)
        apply_window_icon(win)
        win.title("Local Model Manager")
        win.transient(self.root)
        win.grab_set()
        win.configure(fg_color=p.bg_main)
        win.grid_columnconfigure(0, weight=1)
        win.grid_rowconfigure(0, weight=1)

        model_dir = _get_local_model_dir_from_cfg(self.cfg)
        catalogue = _load_models_catalogue()
        scripts_dir = str(_get_app_root() / "scripts")

        # --- state ---
        # dl_procs: canonical_name -> subprocess.Popen
        dl_procs: dict[str, subprocess.Popen] = {}
        # dl_state: canonical_name -> "idle"/"downloading"/"done"/"error"
        dl_state: dict[str, str] = {}
        # dl_queues: canonical_name -> queue.Queue of stdout lines
        dl_queues: dict[str, queue.Queue] = {}
        # dl_cumulative: canonical_name -> {"last_received": int, "offset": int}
        dl_cumulative: dict[str, dict] = {}
        row_widgets: dict[str, dict] = {}  # canonical_name -> widget dict
        check_vars: dict[str, tk.BooleanVar] = {}

        # Build a lookup for catalogue size_mb by canonical_name
        _cat_size_mb: dict[str, int] = {}
        for _e in catalogue:
            _cat_size_mb[_e["canonical_name"]] = _e.get("size_mb", 0)

        def _is_downloaded(canonical_name: str) -> bool:
            # Check if a model pair is downloaded (by canonical name or ISO equivalent)
            pair_dir = Path(model_dir) / canonical_name
            if (pair_dir / "converted.ok").exists():
                return True
            # Also match via ISO-normalised codes (e.g. catalogue "en-fr" vs disk "eng-fra")
            parts = canonical_name.split("-", 1)
            if len(parts) == 2:
                iso_src = _opus_code_to_iso(parts[0])
                iso_tgt = _opus_code_to_iso(parts[1])
                for s, t in list_downloaded_pairs(model_dir):
                    if _opus_code_to_iso(s) == iso_src and _opus_code_to_iso(t) == iso_tgt:
                        return True
            return False

        # --- outer card ---
        card = theme.make_card(win)
        card.configure(width=theme.scale(900))
        win.minsize(theme.scale(900), theme.scale(560))
        card.grid(row=0, column=0, sticky="nsew", padx=PAD, pady=PAD)
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(2, weight=1)  # table area stretches

        # --- header ---
        hdr = ctk.CTkFrame(card, fg_color="transparent")
        hdr.grid(row=0, column=0, sticky="w", padx=PAD, pady=(PAD, 4))
        lang_icon = get_icon("language", size=20)
        if lang_icon:
            ctk.CTkLabel(hdr, text="", image=lang_icon, width=20).pack(side=tk.LEFT, padx=(0, 8))
        theme.make_label(hdr, "Local Model Manager", level="heading").pack(side=tk.LEFT)

        # --- info note ---
        theme.make_label(
            card,
            "OPUS-MT models cover one source \u2192 target pair each. "
            "Download the pairs you need below.",
            level="small", text_color=p.text_muted,
        ).grid(row=1, column=0, sticky="w", padx=PAD, pady=(0, 8))

        # --- scrollable table ---
        # --- table container with sticky header + scrollable body ---
        table_container = ctk.CTkFrame(card, fg_color="transparent")
        table_container.grid(row=2, column=0, sticky="nsew", padx=PAD, pady=(0, 4))
        table_container.grid_columnconfigure(0, weight=1)
        table_container.grid_rowconfigure(1, weight=1)

        # Header row (fixed)
        header_frame = ctk.CTkFrame(table_container, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew")
        _SCROLLBAR_W = 16
        for ci, w in enumerate([30, 1, 1, 80, 120, 160]):
            weight = 0 if ci in (0, 3, 4, 5) else 1
            header_frame.grid_columnconfigure(ci, weight=weight, minsize=w)
        header_frame.grid_columnconfigure(6, weight=0, minsize=_SCROLLBAR_W)

        _hdr_labels = ["", "Source", "Target", "Size (MB)", "Progress", ""]
        for ci, txt in enumerate(_hdr_labels):
            if txt:
                theme.make_label(header_frame, txt, level="section").grid(
                    row=0, column=ci, sticky="w", padx=(4, 8), pady=(0, 6))

        # Scrollable body (model rows)
        try:
            body_frame = ctk.CTkScrollableFrame(table_container, fg_color="transparent")
            try:
                body_frame.configure(height=theme.scale(340))
            except Exception:
                pass
        except Exception:
            body_frame = ctk.CTkFrame(table_container, fg_color="transparent")
        body_frame.grid(row=1, column=0, sticky="nsew")
        # Column layout for body also aligned with header
        for ci, w in enumerate([30, 1, 1, 80, 120, 160]):
            weight = 0 if ci in (0, 3, 4, 5) else 1
            body_frame.grid_columnconfigure(ci, weight=weight, minsize=w)

        lang_opts_map = {code: name for code, name in _get_language_options()}

        def _lang_display(code: str) -> str:
            iso = _opus_code_to_iso(code)
            name = lang_opts_map.get(iso, code)
            return f"{name} ({iso})"

        def _update_row_status(cname: str):
            # Refresh the progress label and action button for one row
            rw = row_widgets.get(cname)
            if not rw:
                return
            downloaded = _is_downloaded(cname)
            state = dl_state.get(cname, "idle")

            # Progress label
            if state == "downloading":
                rw["progress_label"].configure(text="0%", text_color=p.status_info)
            elif state == "error":
                rw["progress_label"].configure(text="Error", text_color=p.status_error)
            elif downloaded:
                rw["progress_label"].configure(text="100%", text_color=p.status_success)
            else:
                rw["progress_label"].configure(text="\u2014", text_color=p.text_muted)

            # Action frame — clear and rebuild
            for child in rw["action_frame"].winfo_children():
                child.destroy()

            if state == "downloading":
                theme.make_button(
                    rw["action_frame"], "Cancel",
                    command=lambda cn=cname: _cancel_download(cn),
                    style="ghost", height=22,
                ).grid(row=0, column=0)
            elif state == "error":
                theme.make_button(
                    rw["action_frame"], "Retry",
                    command=lambda cn=cname: _start_download(cn),
                    style="secondary", height=22,
                ).grid(row=0, column=0)
            elif downloaded:
                theme.make_button(
                    rw["action_frame"], "Delete",
                    command=lambda cn=cname: _delete_model(cn),
                    style="ghost", height=22,
                ).grid(row=0, column=0)
            else:
                theme.make_button(
                    rw["action_frame"], "Download",
                    command=lambda cn=cname: _start_download(cn),
                    style="primary", height=22,
                ).grid(row=0, column=0)

        def _build_rows():
            # Clear existing rows from body_only (keep header fixed)
            for child in body_frame.winfo_children():
                child.destroy()
            row_widgets.clear()
            check_vars.clear()

            r = 0
            for entry in catalogue:
                cname = entry["canonical_name"]
                src_display = _lang_display(entry["source"])
                tgt_display = _lang_display(entry["target"])

                cv = tk.BooleanVar(value=False)
                check_vars[cname] = cv
                cb = ctk.CTkCheckBox(
                    body_frame, text="", variable=cv, width=24,
                    checkmark_color=p.bg_main, fg_color=p.accent,
                    hover_color=p.accent_hover, border_color=p.border,
                )
                cb.grid(row=r, column=0, padx=(4, 0), pady=2)

                theme.make_label(body_frame, src_display, level="body").grid(
                    row=r, column=1, sticky="w", padx=(4, 8), pady=2)
                theme.make_label(body_frame, tgt_display, level="body").grid(
                    row=r, column=2, sticky="w", padx=(4, 8), pady=2)
                theme.make_label(body_frame, str(entry.get("size_mb", "?")), level="body").grid(
                    row=r, column=3, sticky="w", padx=(4, 8), pady=2)

                progress_lbl = theme.make_label(body_frame, "", level="small")
                progress_lbl.grid(row=r, column=4, sticky="w", padx=(4, 8), pady=2)

                action_fr = ctk.CTkFrame(body_frame, fg_color="transparent")
                action_fr.grid(row=r, column=5, sticky="w", padx=(4, 8), pady=2)

                row_widgets[cname] = {
                    "progress_label": progress_lbl,
                    "action_frame": action_fr,
                }
                _update_row_status(cname)
                r += 1

        # --- download / delete / cancel actions ---

        def _find_slug(cname: str) -> str:
            for e in catalogue:
                if e["canonical_name"] == cname:
                    return e["slug"]
            return cname

        def _find_ct2_repo(cname: str) -> str | None:
            for e in catalogue:
                if e["canonical_name"] == cname:
                    return e.get("ct2_repo")
            return None

        def _find_hf_repo(cname: str) -> str | None:
            """Extract HuggingFace repo name from catalogue download_url."""
            for e in catalogue:
                if e["canonical_name"] == cname:
                    url = e.get("download_url", "")
                    if "huggingface.co/" in url:
                        return url.rsplit("huggingface.co/", 1)[-1]
            return None

        def _start_download(cname: str):
            if dl_state.get(cname) == "downloading":
                return
            dl_state[cname] = "downloading"
            dl_cumulative[cname] = {
                "last_received": 0, "last_total": 0, "completed_bytes": 0,
            }
            _update_row_status(cname)

            slug = _find_slug(cname)
            ct2_repo = _find_ct2_repo(cname)
            hf_repo = _find_hf_repo(cname)

            q: queue.Queue = queue.Queue()
            dl_queues[cname] = q

            total_model_bytes = _cat_size_mb.get(cname, 0) * 1024 * 1024

            def _poll():
                if cname not in dl_queues:
                    return
                # Drain all available lines without blocking
                sentinel_received = False
                error_received = False
                while True:
                    try:
                        line = q.get_nowait()
                    except queue.Empty:
                        break
                    if line is None:
                        # Stream ended — process/thread finished
                        sentinel_received = True
                        break
                    line = line.strip()
                    if line.startswith("ERROR:"):
                        error_received = True
                    elif line.startswith("PHASE converting"):
                        rw = row_widgets.get(cname)
                        if rw and "progress_label" in rw:
                            rw["progress_label"].configure(
                                text="Converting\u2026", text_color=p.status_info)
                    elif line.startswith("PROGRESS "):
                        parts = line.split()
                        if len(parts) == 3:
                            try:
                                received = int(parts[1])
                                total = int(parts[2])
                                cum = dl_cumulative.get(cname)
                                if cum is not None:
                                    # Detect new file: received drops
                                    if received < cum["last_received"]:
                                        cum["completed_bytes"] += cum["last_total"]
                                    cum["last_received"] = received
                                    cum["last_total"] = total
                                    total_received = cum["completed_bytes"] + received
                                    aggregate_total = cum["completed_bytes"] + total
                                    # Use catalogue size as cap when available,
                                    # otherwise use actual download totals.
                                    denom = max(total_model_bytes, aggregate_total) if total_model_bytes > 0 else max(1, aggregate_total)
                                    pct = min(99, int(total_received / denom * 100))
                                    rw = row_widgets.get(cname)
                                    if rw and "progress_label" in rw:
                                        rw["progress_label"].configure(
                                            text=f"{pct}%", text_color=p.status_info)
                            except ValueError:
                                pass
                p_obj = dl_procs.get(cname)
                if p_obj is None:
                    # Frozen/in-process path: wait for the sentinel from the download thread
                    if sentinel_received:
                        dl_procs.pop(cname, None)
                        dl_queues.pop(cname, None)
                        dl_cumulative.pop(cname, None)
                        dl_state[cname] = "error" if error_received else "done"
                        _update_row_status(cname)
                        if not error_received:
                            self._refresh_language_dropdowns()
                    else:
                        win.after(100, _poll)  # keep polling until the thread finishes
                    return
                rc = p_obj.poll()
                if rc is None:
                    win.after(100, _poll)
                else:
                    dl_procs.pop(cname, None)
                    dl_queues.pop(cname, None)
                    dl_cumulative.pop(cname, None)
                    dl_state[cname] = "done" if rc == 0 else "error"
                    _update_row_status(cname)
                    if rc == 0:
                        self._refresh_language_dropdowns()

            is_frozen = getattr(sys, "frozen", False) or "__compiled__" in globals()
            if is_frozen:
                dl_procs[cname] = None  # no subprocess

                def _frozen_download():
                    import io
                    from scripts.download_models import download_opus_mt
                    old_stdout = sys.stdout
                    old_stderr = sys.stderr
                    class _QueueWriter(io.TextIOBase):
                        def write(self, s):
                            if s.strip():
                                q.put(s.rstrip())
                            return len(s)
                    writer = _QueueWriter()
                    sys.stdout = writer
                    sys.stderr = writer
                    try:
                        download_opus_mt(slug, model_dir, ct2_repo=ct2_repo,
                                         hf_repo=hf_repo)
                    except SystemExit as e:
                        if e.code != 0:
                            q.put(f"ERROR: download failed (exit code {e.code})")
                    except Exception as exc:
                        q.put(f"ERROR: {exc}")
                    finally:
                        sys.stdout = old_stdout
                        sys.stderr = old_stderr
                        q.put(None)  # sentinel

                threading.Thread(target=_frozen_download, daemon=True).start()
                win.after(100, _poll)
                return

            scripts_dir = str(_get_app_root() / "scripts")

            cmd = [
                sys.executable,
                os.path.join(scripts_dir, "download_models.py"),
                "opus-mt", slug, "--dest-dir", model_dir,
            ]
            if ct2_repo:
                cmd.extend(["--ct2-repo", ct2_repo])
            if hf_repo:
                cmd.extend(["--hf-repo", hf_repo])

            try:
                proc = subprocess.Popen(
                    cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, bufsize=1,
                )
            except Exception as exc:
                dl_state[cname] = "error"
                _update_row_status(cname)
                logger.error("Failed to start download for %s: %s", cname, exc)
                return

            dl_procs[cname] = proc

            def _reader():
                try:
                    for line in proc.stdout:
                        q.put(line)
                except Exception:
                    pass
                finally:
                    q.put(None)  # sentinel: stream ended

            t = threading.Thread(target=_reader, daemon=True)
            t.start()

            win.after(100, _poll)

        def _cancel_download(cname: str):
            proc = dl_procs.pop(cname, None)
            dl_queues.pop(cname, None)
            dl_cumulative.pop(cname, None)
            if proc:
                try:
                    proc.kill()
                except Exception:
                    pass
            dl_state[cname] = "idle"
            _update_row_status(cname)

        def _delete_model(cname: str):
            # Delete the exact canonical dir first
            pair_dir = Path(model_dir) / cname
            if pair_dir.is_dir():
                shutil.rmtree(pair_dir, ignore_errors=True)
            # Also delete any ISO-equivalent dir (e.g. "eng-fra" for "en-fr")
            parts = cname.split("-", 1)
            if len(parts) == 2:
                iso_src = _opus_code_to_iso(parts[0])
                iso_tgt = _opus_code_to_iso(parts[1])
                for s, t in list_downloaded_pairs(model_dir):
                    if _opus_code_to_iso(s) == iso_src and _opus_code_to_iso(t) == iso_tgt:
                        alt_dir = Path(model_dir) / f"{s}-{t}"
                        if alt_dir.is_dir():
                            shutil.rmtree(alt_dir, ignore_errors=True)
            dl_state[cname] = "idle"
            _update_row_status(cname)
            self._refresh_language_dropdowns()

        # --- batch action buttons ---
        batch_frame = ctk.CTkFrame(card, fg_color="transparent")
        batch_frame.grid(row=3, column=0, sticky="w", padx=PAD, pady=(4, PAD))

        def _batch_download():
            for cname, cv in check_vars.items():
                if cv.get() and not _is_downloaded(cname):
                    _start_download(cname)

        def _batch_delete():
            for cname, cv in check_vars.items():
                if cv.get() and _is_downloaded(cname):
                    _delete_model(cname)

        theme.make_button(
            batch_frame, "Download selected",
            command=_batch_download, style="primary", height=28,
        ).pack(side=tk.LEFT, padx=(0, 8))
        theme.make_button(
            batch_frame, "Delete selected",
            command=_batch_delete, style="ghost", height=28,
        ).pack(side=tk.LEFT)

        # --- build initial rows ---
        _build_rows()

        # --- close handler ---
        def _on_close():
            # Terminate any running downloads
            for cn, proc in list(dl_procs.items()):
                try:
                    proc.kill()
                except Exception:
                    pass
            dl_procs.clear()
            dl_queues.clear()
            dl_cumulative.clear()
            self._model_manager_open = False
            self._refresh_language_dropdowns()
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", _on_close)
        win.update_idletasks()

        def _show_mm():
            center_window(win, parent=self.root)
            win.wm_attributes("-alpha", 1)
        win.after(20, _show_mm)
        try:
            win.resizable(True, True)
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
            card, "\u00a9 2026 crt_ (8041q) \u2014  Released under the AGPL-3.0 License",
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
        self._update_usage_label(engine_key)

    def _on_source_lang_changed(self, *_) -> None:
        # Only local engine uses source-based target filtering.
        engine_key = _ENGINE_MAP.get(self.engine_var.get(), "google")
        if engine_key != "local":
            return

        from ..translators.local import list_downloaded_pairs
        model_dir = _get_local_model_dir_from_cfg(self.cfg)
        pairs = list_downloaded_pairs(model_dir)
        lang_opts = _get_language_options()

        source_code = self._source_lang_map.get(self.source_lang_var.get())

        if not source_code:
            tgt_codes = {_opus_code_to_iso(t) for _, t in pairs}
        else:
            tgt_codes = set()
            for s, t in pairs:
                if _opus_code_to_iso(s) == source_code:
                    tgt_codes.add(_opus_code_to_iso(t))

        tgt_filtered = [(c, n) for c, n in lang_opts if c in tgt_codes]
        tgt_display = [f"{name} ({code})" for code, name in tgt_filtered]
        self._lang_map = {f"{name} ({code})": code for code, name in tgt_filtered}
        self.target_lang_box.update_values(tgt_display)

    def _update_usage_label(self, engine_key: str) -> None:
        # Update the small quota label below the engine dropdown
        try:
            from ..translators.usage import get_tracker
            tracker = get_tracker()
            # For Baidu, use the tier-specific usage key
            usage_key = engine_key
            if engine_key == "baidu":
                tier = self.cfg.get("baidu_tier", "standard")
                if tier != "standard":
                    usage_key = "baidu-premium"
            text = tracker.format_usage(usage_key)
            if not text:
                self._engine_usage_label.configure(text="")
                self._engine_usage_label.grid_remove()
                return
            p = theme.get()
            warning = tracker.check_warning(usage_key)
            if warning == "limit":
                color = p.status_error
            elif warning == "warn":
                color = p.status_warn if hasattr(p, "status_warn") else "#e67e22"
            elif warning == "info":
                color = "#f39c12"
            else:
                color = p.text_secondary
            self._engine_usage_label.configure(text=text, text_color=color)
            self._engine_usage_label.grid(**self._engine_usage_label_grid_kw)
        except Exception:
            pass

    def _refresh_language_dropdowns(self) -> None:
        # Recompute source and target language lists based on current engine + detector
        detector = (self.detector_var.get() or "fasttext").strip().lower()
        engine_key = _ENGINE_MAP.get(self.engine_var.get(), "google")
        lang_opts = _get_language_options()
        lang_name_by_code = {code: name for code, name in lang_opts}

        if engine_key == "local":
            # Local engine: only show languages where a downloaded model exists
            from ..translators.local import list_downloaded_pairs
            model_dir = _get_local_model_dir_from_cfg(self.cfg)
            pairs = list_downloaded_pairs(model_dir)
            self._local_downloaded_pairs = pairs

            # Build a map from ISO codes found on disk → display name
            src_codes_set: set[str] = set()
            tgt_codes_set: set[str] = set()
            for s, t in pairs:
                src_codes_set.add(_opus_code_to_iso(s))
                tgt_codes_set.add(_opus_code_to_iso(t))

            # Source list — languages which appear as source in at least one pair
            # No Auto-detect for local engine (each pair needs an explicit model)
            src_filtered = [(c, n) for c, n in lang_opts if c in src_codes_set]
            src_display = [f"{name} ({code})" for code, name in src_filtered]
            source_values = src_display
            self._source_lang_map = {
                f"{name} ({code})": code for code, name in src_filtered
            }
            self.source_lang_box.update_values(source_values)
            self._source_lang_label.configure(text=f"Source language ({len(src_display)})")

            # Target list — all targets reachable from any source
            tgt_filtered = [(c, n) for c, n in lang_opts if c in tgt_codes_set]
            tgt_display = [f"{name} ({code})" for code, name in tgt_filtered]
            self._lang_map = {f"{name} ({code})": code for code, name in tgt_filtered}
            self.target_lang_box.update_values(tgt_display)

            self._log(f"Language lists updated for engine={engine_key!r} "
                      f"(target: {len(tgt_display)}, source: {len(src_display)}, "
                      f"downloaded pairs: {len(pairs)})")
            return

        # Non-local engines: original logic
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
        repo_root = _get_app_root()
        output = self.output_entry.get().strip() or DEFAULT_OUTPUT_FOLDER
        out_path = Path(output)
        is_relative = not out_path.is_absolute()
        if is_relative:
            out_path = (repo_root / out_path).resolve()
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
        google_project_id = self.cfg.get("google_project_id", "")
        google_sa_json = self.cfg.get("google_sa_json", "")
        baidu_appid = self.cfg.get("baidu_appid", "")
        baidu_appkey = self.cfg.get("baidu_appkey", "")
        baidu_tier = self.cfg.get("baidu_tier", "standard")
        azure_key = self.cfg.get("azure_key", "")
        azure_region = self.cfg.get("azure_region", "")
        deepl_api_key = self.cfg.get("deepl_api_key", "")
        local_model_dir = _get_local_model_dir_from_cfg(self.cfg)

        # Validate credentials for engines that require them
        if engine == "baidu" and (not baidu_appid or not baidu_appkey):
            messagebox.showwarning(
                "Missing credentials",
                "Baidu Translate requires an App ID and App Key.\n"
                "Please configure them in Settings → API Keys.",
            )
            return
        if engine == "google-cloud-v3" and not google_project_id:
            messagebox.showwarning(
                "Missing credentials",
                "Google Cloud API (v3) requires a Project ID.\n"
                "Please configure it in Settings → API Keys.",
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
        if engine == "azure" and (not azure_key or not azure_region):
            messagebox.showwarning(
                "Missing credentials",
                "Microsoft Azure Translator requires a Subscription Key and Region.\n"
                "Please configure them in Settings → Azure Translator.",
            )
            return
        if engine == "deepl" and not deepl_api_key:
            messagebox.showwarning(
                "Missing API key",
                "DeepL requires an API Key.\n"
                "Please configure it in Settings \u2192 DeepL.",
            )
            return

        # Local engine: verify required model is downloaded
        if engine == "local":
            from ..translators.local import list_downloaded_pairs
            _lm_dir = local_model_dir or _get_default_model_dir()
            _local_pairs = list_downloaded_pairs(_lm_dir)
            if not _local_pairs:
                if messagebox.askyesno(
                    "No models downloaded",
                    "No local OPUS-MT models found.\n\n"
                    "Open Settings to download models?",
                ):
                    self._open_settings()
                return
            _pair_set = {(_opus_code_to_iso(s), _opus_code_to_iso(t))
                         for s, t in _local_pairs}
            if (source_lang, lang) not in _pair_set:
                if messagebox.askyesno(
                    "Missing model",
                    f"No local model for {source_lang} \u2192 {lang}.\n\n"
                    "Open Settings to download it?",
                ):
                    self._open_settings()
                return

        # Warn if usage is close to or at the monthly limit.
        # For Baidu, the usage key depends on the configured tier.
        _usage_key = engine
        if engine == "baidu" and baidu_tier != "standard":
            _usage_key = "baidu-premium"
        if engine in ("azure", "deepl", "google-cloud", "baidu"):
            try:
                from ..translators.usage import get_tracker
                tracker = get_tracker()
                warning = tracker.check_warning(_usage_key)
                if warning in ("warn", "limit"):
                    used_str = tracker.format_usage(_usage_key) or ""
                    engine_display = _ENGINE_REVERSE.get(engine, engine)
                    if warning == "limit":
                        msg = (
                            f"You have used {used_str} of your {engine_display} quota.\n\n"
                            "You may have exceeded your monthly limit. "
                            "The API will likely reject requests.\n\nContinue anyway?"
                        )
                    else:
                        msg = (
                            f"You have used {used_str} of your {engine_display} quota "
                            "for this month.\n\nYou are approaching your monthly limit. Continue?"
                        )
                    if not messagebox.askyesno("Usage Warning", msg):
                        return
                    self._update_usage_label(engine)
            except Exception:
                pass

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
        self._update_progress_label("Starting\u2026")
        self._current_file_name = ""

        self.worker.start(
            self.files, lang, output, None,
            self._progress_cb, self._log,
            source_lang=source_lang,
            detector=detector,
            engine=engine,
            proxies=proxies,
            google_api_key=google_api_key,
            google_project_id=google_project_id,
            google_sa_json=google_sa_json,
            baidu_appid=baidu_appid,
            baidu_appkey=baidu_appkey,
            baidu_tier=baidu_tier,
            azure_key=azure_key,
            azure_region=azure_region,
            deepl_api_key=deepl_api_key,
            local_model_dir=local_model_dir,
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
                self._current_file_name = Path(filepath).name
                self._update_progress_label(
                    f"Translating {self._current_file_name}\u2026 0%"
                )
            elif status == "progress":
                # elapsed carries the global fraction (0.0–1.0) for intra-file progress
                frac = elapsed if elapsed is not None else 0.0
                frac = max(0.0, min(1.0, frac))
                self._set_progress(frac)
                name = getattr(self, "_current_file_name", "")
                self._update_progress_label(
                    f"Translating {name}\u2026 {int(frac * 100)}%"
                )
            elif status in ("finished", "error"):
                self.completed_files += 1
                self._update_file_status(filepath, status, elapsed)
                pct = self.completed_files / max(1, self.total_files)
                self._set_progress(pct)
                remaining = self.total_files - self.completed_files
                if remaining > 0:
                    self._update_progress_label(
                        f"{self.completed_files} / {self.total_files} files ({int(pct * 100)}%)"
                    )
                else:
                    self._update_progress_label(
                        f"{self.completed_files} / {self.total_files} files (100%)"
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
        try:
            engine_key = _ENGINE_MAP.get(self.engine_var.get(), "google")
            self._update_usage_label(engine_key)
        except Exception:
            pass
        if cancelled:
            pct = self.completed_files / max(1, self.total_files)
            self._update_progress_label(
                f"Cancelled \u2014 {int(pct * 100)}%",
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

    # --- responsive scaling (MUST run before ctk.CTk() is instantiated) ---
    try:
        from customtkinter.windows.widgets.scaling.scaling_tracker import ScalingTracker

        _probe = tk.Tk()
        _probe.withdraw()
        _dpi            = _probe.winfo_fpixels("1i")   # physical pixels per inch
        _screen_w_phys  = _probe.winfo_screenwidth()
        _screen_h_phys  = _probe.winfo_screenheight()
        _probe.destroy()
        del _probe

        ctk_dpi_scale    = _dpi / 96.0                 # 96 dpi == 100 %
        screen_w_logical = _screen_w_phys / ctk_dpi_scale
        screen_h_logical = _screen_h_phys / ctk_dpi_scale

        margin    = 0.95
        ratio_w   = (screen_w_logical * margin) / theme.WINDOW_WIDTH
        ratio_h   = (screen_h_logical * margin) / theme.WINDOW_HEIGHT
        layout_scale = min(1.0, ratio_w, ratio_h)

        if layout_scale < 0.98:
            # Write directly to the class-level attributes
            ScalingTracker.widget_scaling = max(layout_scale, 0.4)
            ScalingTracker.window_scaling = max(layout_scale, 0.4)
    except Exception:
        pass

    root = ctk.CTk()
    root.wm_attributes("-alpha", 0)  # keep invisible until centered
    root.withdraw()

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
