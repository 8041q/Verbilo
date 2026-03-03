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

try:
    import customtkinter as ctk
except Exception:
    ctk = None

import tomllib
import webbrowser

from . import gui_theme as theme
from .gui_helpers import Worker, list_supported_files, center_window, GuiLoggingHandler
from .gui_config import load_config, save_config
from .icons import get_icon, get_photo_image, get_app_icon

logger = logging.getLogger(__name__)

# --- version & about constants ---

def _read_pyproject_meta() -> tuple[str, str]:
    # Read version and build_date from pyproject.toml. Returns (version, build_date).
    try:
        toml_path = Path(__file__).parent.parent / "pyproject.toml"
        with open(toml_path, "rb") as fh:
            data = tomllib.load(fh)
        poetry = data.get("tool", {}).get("poetry", {})
        version = poetry.get("version", "unknown")
        build_date = poetry.get("build_date", "unknown")
        return version, build_date
    except Exception:
        return "unknown", "unknown"

APP_VERSION, APP_BUILD_DATE = _read_pyproject_meta()

# Github URLs
GITHUB_URL = "https://github.com/8041q/Verbilo"
RELEASES_URL = "https://github.com/8041q/Verbilo/releases"

# Default output folder name (relative to cwd)
DEFAULT_OUTPUT_FOLDER = "Output"


def _try_make_relative(absolute_path: str) -> str:
    # Return a path relative to cwd if possible, otherwise return the original.
    try:
        return str(Path(absolute_path).relative_to(Path.cwd()))
    except ValueError:
        return absolute_path

# --- language helpers ---

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


class SearchableComboBox:
    # Entry + scrollable Toplevel listbox.
    # - Click/Tab: select-all and open dropdown.
    # - Type: live filter, keep dropdown open.
    # - Click item or Enter: confirm and close.
    # - Click outside/Escape/focus-out: revert to last valid and close.

    _POPUP_ROWS = 8  # max visible rows before scrolling

    def __init__(self, parent, values, variable, **kw):
        p = theme.get()
        self._parent = parent
        self._all_values = list(values)
        self._variable = variable
        self._last_valid = variable.get() or (values[0] if values else "")
        self._popup = None
        # Auto-clears after a short delay so a stale True never blocks reopening.
        self._suppress_open = False

        # Styled outer frame
        self._frame = ctk.CTkFrame(
            parent,
            fg_color=p.bg_input,
            corner_radius=theme.BUTTON_CORNER_RADIUS,
            border_width=1,
            border_color=p.border,
        ) if ctk else tk.Frame(parent)
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

        # Arrow button
        arrow_img = get_icon("chevron-down", size=14)
        btn_kw = dict(
            master=self._frame, width=28, height=28,
            fg_color="transparent", hover_color=p.bg_card,
            corner_radius=4, command=self._on_arrow,
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
            self._btn = tk.Button(self._frame, text="\u25BC", command=self._on_arrow)
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
        self._frame.after(200, self._bind_root_click)

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

    def refresh_colors(self):
        pass

    # -- Entry event handlers ------------------------------------------

    def _on_focus_in(self, _event=None):
        # Tab/programmatic focus: select-all and open.
        if self._suppress_open:
            return
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
            self._confirm(self._listbox.get(sel[0] if sel else 0))
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
            highlightbackground=p.border, highlightthickness=1,
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
        row_px = theme.FONT_BODY[1] + 10
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
                if str(focused).startswith(str(self._popup)):
                    return
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

        # Apply saved appearance mode (default: Dark)
        saved_mode = self.cfg.get("appearance_mode", "Dark")
        theme.set_mode(saved_mode)

        self._build_ui()

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

    def _initialdir_for_input(self) -> str:
        try:
            if self.files:
                return str(Path(self.files[0]).parent.resolve())
        except Exception:
            pass
        if self.cfg.get("default_input"):
            return str((Path.cwd() / self.cfg["default_input"]).resolve())
        return str(Path.cwd())

    def _initialdir_for_output(self) -> str:
        try:
            if hasattr(self, "output_entry"):
                val = self.output_entry.get().strip()
                if val:
                    return str((Path.cwd() / val).resolve())
        except Exception:
            pass
        if self.cfg.get("default_output"):
            return str((Path.cwd() / self.cfg["default_output"]).resolve())
        return str(Path.cwd())

    # --- UI construction ---

    def _build_ui(self):
        p = theme.get()

        self.root.title("Verbilo")
        if isinstance(self.root, ctk.CTk):
            self.root.configure(fg_color=p.bg_main)
        else:
            self.root.configure(bg=p.bg_main)
        center_window(self.root, theme.WINDOW_WIDTH, theme.WINDOW_HEIGHT)
        try:
            self.root.minsize(theme.WINDOW_MIN_WIDTH, theme.WINDOW_MIN_HEIGHT)
            self.root.resizable(True, True)
        except Exception:
            pass

        # Root grid: sidebar (col 0, fixed)  |  content (col 1, expands)
        self.root.grid_columnconfigure(0, weight=0, minsize=theme.SIDEBAR_WIDTH)
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

        # Sidebar
        self._build_sidebar()

        # Main content
        self._build_content()

    # --- sidebar ---

    def _build_sidebar(self):
        PAD = theme.PADDING
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
        theme.make_label(
            self.sidebar, "Source language", level="small",
        ).grid(row=row, column=0, sticky="w", padx=PAD, pady=(4, 2))
        row += 1

        lang_opts = _get_language_options()
        self._lang_map = {f"{name} ({code})": code for code, name in lang_opts}
        display_values = list(self._lang_map.keys())
        if not display_values:
            display_values = ["English (en)"]

        self.source_lang_var = tk.StringVar(value="Auto-detect (translate all)")
        source_values = ["Auto-detect (translate all)"] + display_values
        self._source_lang_map = {"Auto-detect (translate all)": "auto"}
        self._source_lang_map.update(self._lang_map)

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

        # Translator
        theme.make_label(
            self.sidebar, "Translator", level="small",
        ).grid(row=row, column=0, sticky="w", padx=PAD, pady=(6, 2))
        row += 1

        self.translator_var = tk.StringVar(value="auto")
        self.translator_menu = ctk.CTkOptionMenu(
            self.sidebar,
            values=["auto", "identity", "deep"],
            variable=self.translator_var,
            fg_color=p.bg_input,
            button_color=p.accent,
            button_hover_color=p.accent_hover,
            text_color=p.text_secondary,
            dropdown_fg_color=p.bg_popup,
            dropdown_hover_color=p.accent,
            dropdown_text_color=p.text_secondary,
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=theme.FONT_BODY[1]),
            dropdown_font=ctk.CTkFont(family=theme.FONT_FAMILY, size=theme.FONT_BODY[1]),
            corner_radius=theme.BUTTON_CORNER_RADIUS,
            height=32,
        )
        self.translator_menu.grid(row=row, column=0, sticky="ew", padx=PAD, pady=(0, 6))
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
        # About row: keep the main button text as 'About' and show a small
        # muted '(beta)' badge to the right so the clickable area remains
        # the same and the badge is visually distinct.
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
        PAD = theme.PADDING
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
            padx=theme.PADDING, pady=theme.PADDING,
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

        scrollbar = ttk.Scrollbar(container, orient=tk.VERTICAL, command=self.file_table.yview)
        self.file_table.configure(yscrollcommand=scrollbar.set)

        self.file_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

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
        PAD = theme.PADDING

        win = ctk.CTkToplevel(self.root)
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
            init = str((Path.cwd() / raw).resolve()) if raw else str(Path.cwd())
            d = filedialog.askdirectory(title="Select default input folder", initialdir=init)
            if d:
                in_entry.delete(0, tk.END)
                in_entry.insert(0, _try_make_relative(d))

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
            init = str((Path.cwd() / raw).resolve()) if raw else str(Path.cwd())
            d = filedialog.askdirectory(title="Select default output folder", initialdir=init)
            if d:
                out_entry.delete(0, tk.END)
                out_entry.insert(0, _try_make_relative(d))

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

        # --- Updates section ---
        theme.make_divider(card).grid(row=5, column=0, columnspan=3, sticky="ew", padx=PAD, pady=(4, 8))

        theme.make_label(card, "UPDATES", level="section").grid(
            row=6, column=0, sticky="w", padx=PAD, pady=(0, 6),
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
        auto_updates_cb.grid(row=6, column=1, columnspan=2, sticky="w", padx=4, pady=(0, 6))

        # Debug mode (show extra informational messages)
        debug_var = tk.BooleanVar(value=self.cfg.get("debug_mode", True))
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
        debug_cb.grid(row=7, column=1, columnspan=2, sticky="w", padx=4, pady=(8, 6))

        # Inline validation error
        self._settings_error = theme.make_label(
            card, "", level="tiny",
            text_color=p.status_error,
        )
        self._settings_error.grid(row=8, column=0, columnspan=3, sticky="w", padx=PAD, pady=(0, 2))

        # Button row
        btn_frame = ctk.CTkFrame(card, fg_color="transparent")
        btn_frame.grid(row=9, column=0, columnspan=3, pady=(4, PAD))

        def _save_and_close():
            inp = in_entry.get().strip()
            out = out_entry.get().strip()
            if not inp or not out:
                self._settings_error.configure(text="Input or Output path cannot be empty")
                return
            self._settings_error.configure(text="")
            self.cfg["default_input"] = inp
            self.cfg["default_output"] = out
            new_mode = "Dark" if mode_switch_var.get() else "Light"
            self.cfg["appearance_mode"] = new_mode
            self.cfg["auto_check_updates"] = auto_updates_var.get()
            self.cfg["debug_mode"] = debug_var.get()
            save_config(self.cfg)
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

        # Centre the dialog
        win.update_idletasks()
        center_window(win, max(win.winfo_reqwidth(), 520), parent=self.root)
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
            elif not startup and result["status"] == "update":
                self.root.after(0, lambda r=result: self._show_update_dialog(r))

        threading.Thread(target=_check, daemon=True).start()

    def _show_update_dialog(self, result: dict):
        # Show a small dialog when a newer release is available.
        if result.get("status") != "update":
            return
        p = theme.get()
        PAD = theme.PADDING

        dlg = ctk.CTkToplevel(self.root)
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
        center_window(dlg, max(dlg.winfo_reqwidth(), 400), parent=self.root)
        try:
            dlg.resizable(False, False)
        except Exception:
            pass

    # --- about dialog ---

    def _open_about(self):
        # Open the standalone About dialog.
        p = theme.get()
        PAD = theme.PADDING

        win = ctk.CTkToplevel(self.root)
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
            pil_img = get_app_icon(size=48)
            if pil_img is not None and ctk is not None:
                logo_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(48, 48))
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
            brand_frame, "(beta)", level="tiny", text_color=p.text_muted,
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
            card, "\u00a9 2026 crt_  \u2014  Released under the MIT License",
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
            check_btn.configure(state="disabled")

            def _on_result(result):
                check_btn.configure(state="normal")
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

        theme.make_button(links_frame, "View on GitHub", command=_open_github, style="ghost",
                          height=28).pack(side=tk.LEFT, padx=(0, 6))
        theme.make_button(links_frame, "Release Notes", command=_open_releases, style="ghost",
                          height=28).pack(side=tk.LEFT, padx=(0, 16))
        theme.make_label(links_frame, "Made by crt_", level="tiny").pack(side=tk.LEFT)

        win.protocol("WM_DELETE_WINDOW", win.destroy)
        win.update_idletasks()
        center_window(win, max(win.winfo_reqwidth(), 460), parent=self.root)
        try:
            win.resizable(False, False)
        except Exception:
            pass

    # --- file management ---

    def _add_files(self):
        init = self._initialdir_for_input()
        paths = filedialog.askopenfilenames(title="Select files", initialdir=init)
        for p in paths:
            if p not in self.files:
                self._add_file_to_table(p)

    def _select_folder(self):
        init = self._initialdir_for_input()
        d = filedialog.askdirectory(title="Select folder containing files", initialdir=init)
        if not d:
            return
        found = list_supported_files(d)
        if not found:
            messagebox.showinfo("No files", f"No supported files found in {d}")
            return
        for f in found:
            if f not in self.files:
                self._add_file_to_table(f)

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
        d = filedialog.askdirectory(title="Select output folder", initialdir=init)
        if d:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, _try_make_relative(d))

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

        # Normalize translator selection
        sel_trans = (self.translator_var.get() or "").strip()
        norm = sel_trans.lower()
        if not sel_trans or norm in ("auto", "none"):
            translator = None
        else:
            translator = sel_trans

        try:
            self._log(f"Starting: source={source_lang!r}, target={lang!r}, translator={translator!r}")
        except Exception:
            pass

        # Update UI state
        self._running = True
        self.start_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")

        self._update_all_statuses("pending")
        self.total_files = len(self.files)
        self.completed_files = 0
        self._file_start_times.clear()
        self._set_progress(0.0)
        self._update_progress_label(f"0 / {self.total_files} files (0%)")

        self.worker.start(
            self.files, lang, output, translator,
            self._progress_cb, self._log,
            source_lang=source_lang,
        )

    def _cancel(self):
        if not self._running:
            return
        self.worker.stop()
        self.cancel_btn.configure(state="disabled")
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
        self.start_btn.configure(state="normal")
        self.cancel_btn.configure(state="disabled")
        if cancelled:
            pct = self.completed_files / max(1, self.total_files)
            self._update_progress_label(
                f"Cancelled \u2014 {self.completed_files} / {self.total_files} files ({int(pct * 100)}%)",
            )

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

    # DPI awareness (must be called before creating the root window on Windows)
    theme.init_dpi()

    root = ctk.CTk()

    # Finalize DPI scaling now that root exists
    theme.init_dpi(root)

    # Set app icon (title bar + taskbar)
    try:
        from PIL import ImageTk
        icon_pil = get_app_icon(size=64)
        if icon_pil:
            icon_tk = ImageTk.PhotoImage(icon_pil)
            root.iconphoto(True, icon_tk)
            root._app_icon_ref = icon_tk  # prevent garbage collection
    except Exception:
        pass

    app = App(root)  # noqa: F841

    # Install GUI logging handler so Python `logging` and captured `warnings`
    # are forwarded into the app's log pane. Wrap the app log callable so
    # debug/info lines can be toggled via the Settings `debug_mode` option.
    try:
        import logging as _logging
        import re

        raw_log = app._log

        def _filtered_log(msg: str):
            try:
                # Always pass through the worker-done sentinel
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

                # Forward the (possibly filtered) message to the raw GUI logger.
                # Content sanitization for warnings is handled centrally by the
                # GuiLoggingHandler so we don't mutate message text here.
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
        _handler.setLevel(_logging.INFO)
        root_logger = _logging.getLogger()
        # Avoid adding duplicate handlers on repeated starts
        if not any(isinstance(h, GuiLoggingHandler) for h in root_logger.handlers):
            root_logger.addHandler(_handler)
        root_logger.setLevel(_logging.INFO)
        _logging.captureWarnings(True)

        # Replace app._log so GUI's own log calls go through the same filter
        app._log = _filtered_log
    except Exception:
        logger.exception("Failed to install GUI logging handler")

    root.mainloop()


if __name__ == "__main__":
    main()
