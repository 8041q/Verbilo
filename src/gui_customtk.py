# gui_customtk.py — dashboard GUI; sidebar (controls) + content (file table, progress, log)

from __future__ import annotations

import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import logging

try:
    import customtkinter as ctk
except Exception:
    ctk = None

from . import gui_theme as theme
from .gui_helpers import Worker, list_supported_files, center_window
from .gui_config import load_config, save_config

logger = logging.getLogger(__name__)

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

_ComboBoxBase = ctk.CTkFrame if ctk else tk.Frame

# themed searchable dropdown — click to open a filtered popup listbox
class SearchableComboBox(_ComboBoxBase):  # type: ignore[misc]

    def __init__(self, parent, values: list[str], variable: tk.StringVar,
                 width: int = 28, **kw):
        p = theme.get()
        if ctk:
            super().__init__(parent, fg_color="transparent", **kw)
        else:
            super().__init__(parent, **kw)

        self._values = values
        self._variable = variable
        self._popup: tk.Toplevel | None = None

        # --- Display row (looks like a dropdown selector) ---
        display_font = (theme.FONT_FAMILY, theme.FONT_BODY[1])
        arrow_font = (theme.FONT_FAMILY, theme.FONT_SMALL[1])

        self._display = tk.Label(
            self,
            text=variable.get() or (values[0] if values else ""),
            anchor="w", padx=8, pady=5,
            font=display_font,
            bg=p.bg_input, fg=p.text_secondary,
            relief="flat", borderwidth=0,
            cursor="hand2",
        )
        self._display.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self._arrow = tk.Label(
            self,
            text="\u25bc",
            font=arrow_font,
            bg=p.bg_input, fg=p.text_muted,
            relief="flat", borderwidth=0,
            padx=6, pady=5,
            cursor="hand2",
        )
        self._arrow.pack(side=tk.RIGHT)

        self._display.bind("<Button-1>", lambda e: self._toggle_popup())
        self._arrow.bind("<Button-1>", lambda e: self._toggle_popup())

    def get(self) -> str:
        return self._display.cget("text")

    def set(self, value: str):
        self._display.configure(text=value)
        self._variable.set(value)

    def refresh_colors(self):
        p = theme.get()
        self._display.configure(bg=p.bg_input, fg=p.text_secondary)
        self._arrow.configure(bg=p.bg_input, fg=p.text_muted)

    def _toggle_popup(self):
        if self._popup and self._popup.winfo_exists():
            self._close_popup()
        else:
            self._open_popup()

    def _open_popup(self):
        if self._popup and self._popup.winfo_exists():
            self._popup.destroy()

        p = theme.get()
        self._popup = tk.Toplevel(self)
        self._popup.wm_overrideredirect(True)
        self._popup.wm_attributes("-topmost", True)

        self.update_idletasks()
        x = self._display.winfo_rootx()
        y = self._display.winfo_rooty() + self._display.winfo_height()
        w = self._display.winfo_width() + self._arrow.winfo_width()

        popup_frame = tk.Frame(
            self._popup, bg=p.bg_popup,
            highlightbackground=p.border, highlightthickness=1,
        )
        popup_frame.pack(fill=tk.BOTH, expand=True)

        # --- Search entry ---
        search_frame = tk.Frame(popup_frame, bg=p.bg_popup)
        search_frame.pack(fill=tk.X, padx=6, pady=(6, 3))

        tk.Label(
            search_frame, text="\U0001F50D", bg=p.bg_popup,
            fg=p.text_muted, font=(theme.FONT_FAMILY, theme.FONT_TINY[1]),
        ).pack(side=tk.LEFT, padx=(2, 0))

        self._search_var = tk.StringVar()
        self._search_entry = tk.Entry(
            search_frame, textvariable=self._search_var,
            font=(theme.FONT_FAMILY, theme.FONT_BODY[1]),
            bg=p.bg_input, fg=p.text_secondary,
            insertbackground=p.text_secondary,
            relief="flat", borderwidth=2,
        )
        self._search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)

        # --- Separator ---
        tk.Frame(popup_frame, height=1, bg=p.divider).pack(fill=tk.X, padx=6)

        # --- Scrollable listbox ---
        list_frame = tk.Frame(popup_frame, bg=p.bg_popup)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=(3, 6))

        scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL)
        self._listbox = tk.Listbox(
            list_frame,
            width=0,
            height=min(10, max(4, len(self._values))),
            yscrollcommand=scrollbar.set,
            font=(theme.FONT_FAMILY, theme.FONT_BODY[1]),
            activestyle="none",
            selectbackground=p.accent,
            selectforeground=p.text_on_accent,
            bg=p.bg_popup, fg=p.text_secondary,
            relief="flat", borderwidth=0,
            highlightthickness=0,
        )
        scrollbar.config(command=self._listbox.yview)
        self._listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self._populate_listbox(self._values)

        current = self.get()
        if current in self._values:
            idx = self._values.index(current)
            self._listbox.selection_set(idx)
            self._listbox.see(idx)

        # Bindings
        self._search_var.trace_add("write", self._on_search_changed)
        self._search_entry.bind("<Return>", self._on_search_enter)
        self._search_entry.bind("<Escape>", self._close_popup)
        self._search_entry.bind("<Down>", self._focus_listbox)
        self._listbox.bind("<ButtonRelease-1>", self._on_select)
        self._listbox.bind("<Return>", self._on_select)
        self._listbox.bind("<Escape>", self._close_popup)

        row_h = 22
        list_rows = min(10, max(4, len(self._values)))
        popup_h = min(340, list_rows * row_h + 50)
        self._popup.geometry(f"{w}x{popup_h}+{x}+{y}")

        self._search_entry.focus_set()
        self._popup.bind("<FocusOut>", self._on_popup_focus_out)

    def _populate_listbox(self, items: list[str]):
        self._listbox.delete(0, tk.END)
        for item in items:
            self._listbox.insert(tk.END, item)
        self._filtered = items

    def _on_search_changed(self, *_args):
        query = self._search_var.get().lower()
        if query:
            filtered = [v for v in self._values if query in v.lower()]
        else:
            filtered = self._values
        self._populate_listbox(filtered if filtered else self._values)
        if self._listbox.size() > 0:
            self._listbox.selection_set(0)
            self._listbox.see(0)

    def _focus_listbox(self, event=None):
        if self._popup and self._popup.winfo_exists() and hasattr(self, "_listbox"):
            self._listbox.focus_set()
            if self._listbox.size() > 0 and not self._listbox.curselection():
                self._listbox.selection_set(0)
                self._listbox.activate(0)

    def _on_search_enter(self, event=None):
        if hasattr(self, "_listbox") and self._listbox.size() > 0:
            sel = self._listbox.curselection()
            idx = sel[0] if sel else 0
            value = self._listbox.get(idx)
            self.set(value)
        self._close_popup()

    def _on_select(self, event=None):
        if not hasattr(self, "_listbox"):
            return
        sel = self._listbox.curselection()
        if sel:
            value = self._listbox.get(sel[0])
            self.set(value)
        self._close_popup()

    def _close_popup(self, event=None):
        if self._popup and self._popup.winfo_exists():
            self._popup.destroy()
        self._popup = None

    def _on_popup_focus_out(self, event=None):
        self.after(150, self._maybe_close)

    def _maybe_close(self):
        try:
            focused = self.focus_get()
            if focused is None:
                self._close_popup()
                return
            if self._popup and self._popup.winfo_exists():
                try:
                    if str(focused).startswith(str(self._popup)):
                        return
                except Exception:
                    pass
                if hasattr(self, "_search_entry") and focused == self._search_entry:
                    return
                if hasattr(self, "_listbox") and focused == self._listbox:
                    return
            self._close_popup()
        except Exception:
            self._close_popup()


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

        # Apply defaults from config
        default_out = self.cfg.get("default_output")
        if default_out:
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
                return str(Path(self.files[0]).parent)
        except Exception:
            pass
        if self.cfg.get("default_input"):
            return self.cfg.get("default_input") or ""
        return str(Path.cwd())

    def _initialdir_for_output(self) -> str:
        try:
            if hasattr(self, "output_entry"):
                val = self.output_entry.get().strip()
                if val:
                    return val
        except Exception:
            pass
        if self.cfg.get("default_output"):
            return str(self.cfg.get("default_output") or "")
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

        # App title
        theme.make_label(
            self.sidebar, "Verbilo", level="heading",
        ).grid(row=row, column=0, sticky="w", padx=PAD, pady=(PAD, PAD * 2))
        row += 1

        # TRANSLATION section
        theme.make_label(
            self.sidebar, "Translation", level="section",
        ).grid(row=row, column=0, sticky="w", padx=PAD, pady=(0, 4))
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
        self.source_lang_box.grid(row=row, column=0, sticky="ew", padx=PAD, pady=(0, 4))
        row += 1

        # Target language
        theme.make_label(
            self.sidebar, "Target language", level="small",
        ).grid(row=row, column=0, sticky="w", padx=PAD, pady=(8, 2))
        row += 1

        self.lang_var = tk.StringVar(
            value=display_values[0] if display_values else "English (en)",
        )
        self.target_lang_box = SearchableComboBox(
            self.sidebar, display_values, self.lang_var,
        )
        self.target_lang_box.grid(row=row, column=0, sticky="ew", padx=PAD, pady=(0, 4))
        row += 1

        # Translator
        theme.make_label(
            self.sidebar, "Translator", level="small",
        ).grid(row=row, column=0, sticky="w", padx=PAD, pady=(8, 2))
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
        )
        self.translator_menu.grid(row=row, column=0, sticky="ew", padx=PAD, pady=(0, 4))
        row += 1

        # Divider
        theme.make_divider(self.sidebar).grid(
            row=row, column=0, sticky="ew", padx=PAD, pady=8,
        )
        row += 1

        # OUTPUT section
        theme.make_label(
            self.sidebar, "Output", level="section",
        ).grid(row=row, column=0, sticky="w", padx=PAD, pady=(0, 4))
        row += 1

        theme.make_label(
            self.sidebar, "Output folder", level="small",
        ).grid(row=row, column=0, sticky="w", padx=PAD, pady=(4, 2))
        row += 1

        self.output_entry = theme.make_entry(self.sidebar)
        self.output_entry.grid(row=row, column=0, sticky="ew", padx=PAD, pady=(0, 4))
        row += 1

        theme.make_button(
            self.sidebar, "Browse\u2026", command=self._select_output, style="secondary",
        ).grid(row=row, column=0, sticky="ew", padx=PAD, pady=(0, 8))
        row += 1

        # Divider
        theme.make_divider(self.sidebar).grid(
            row=row, column=0, sticky="ew", padx=PAD, pady=4,
        )
        row += 1

        # Action buttons
        self.start_btn = theme.make_button(
            self.sidebar, "\u25b6  Start Translation", command=self._start, style="primary",
            height=36,
        )
        self.start_btn.grid(row=row, column=0, sticky="ew", padx=PAD, pady=(8, 4))
        row += 1

        self.cancel_btn = theme.make_button(
            self.sidebar, "Cancel", command=self._cancel, style="secondary",
            height=32, state="disabled",
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

        theme.make_button(
            self.sidebar, "\u2699  Settings", command=self._open_settings, style="ghost",
            anchor="w",
        ).grid(row=row, column=0, sticky="ew", padx=PAD, pady=(4, PAD))

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
            row=0, column=0, sticky="w", padx=PAD, pady=6,
        )
        btn_frame = ctk.CTkFrame(toolbar, fg_color="transparent")
        btn_frame.grid(row=0, column=1, sticky="e", padx=PAD, pady=6)

        theme.make_button(
            btn_frame, "Add Files", command=self._add_files, style="primary",
        ).pack(side=tk.LEFT, padx=(0, 8))
        theme.make_button(
            btn_frame, "Select Folder", command=self._select_folder, style="primary",
        ).pack(side=tk.LEFT, padx=(0, 8))
        theme.make_button(
            btn_frame, "Clear", command=self._clear_files, style="secondary",
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
            progress_card, "Ready", level="body",
        )
        self.progress_label.grid(
            row=0, column=0, sticky="w", padx=PAD, pady=(10, 4),
        )

        self.progress = ctk.CTkProgressBar(
            progress_card,
            progress_color=p.accent,
            fg_color=p.bg_main,
            corner_radius=6,
            height=10,
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
            corner_radius=8,
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
        heading_font = (theme.FONT_FAMILY, theme.FONT_BODY[1], "bold")

        style.configure(
            "FileTable.Treeview",
            rowheight=30,
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
            foreground=p.text_secondary,
            borderwidth=0,
            relief="flat",
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

    # --- table helpers ---

    def _add_file_to_table(self, filepath: str, status: str = "pending"):
        self.files.append(filepath)
        name = os.path.basename(filepath)
        idx = len(self.files) - 1
        row_tag = "even" if idx % 2 == 0 else "odd"
        iid = self.file_table.insert(
            "", tk.END, text=name, values=(status, ""), tags=(status, row_tag),
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

        theme.make_label(card, "Settings", level="heading").grid(
            row=0, column=0, columnspan=3, sticky="w", padx=PAD, pady=(PAD, PAD),
        )

        # Default input folder
        theme.make_label(card, "Default input folder:", level="body").grid(
            row=1, column=0, sticky="w", padx=PAD, pady=(0, 6),
        )
        in_entry = theme.make_entry(card)
        in_entry.grid(row=1, column=1, sticky="ew", padx=4, pady=(0, 6))
        in_entry.insert(0, self.cfg.get("default_input", ""))

        def _browse_default_input():
            init = in_entry.get().strip() or self.cfg.get("default_input") or str(Path.cwd())
            d = filedialog.askdirectory(title="Select default input folder", initialdir=init)
            if d:
                in_entry.delete(0, tk.END)
                in_entry.insert(0, d)

        theme.make_button(card, "Browse", command=_browse_default_input, style="secondary").grid(
            row=1, column=2, padx=(4, PAD), pady=(0, 6),
        )

        # Default output folder
        theme.make_label(card, "Default output folder:", level="body").grid(
            row=2, column=0, sticky="w", padx=PAD, pady=(0, 6),
        )
        out_entry = theme.make_entry(card)
        out_entry.grid(row=2, column=1, sticky="ew", padx=4, pady=(0, 6))
        out_entry.insert(0, self.cfg.get("default_output", ""))

        def _browse_default_output():
            init = out_entry.get().strip() or self.cfg.get("default_output") or str(Path.cwd())
            d = filedialog.askdirectory(title="Select default output folder", initialdir=init)
            if d:
                out_entry.delete(0, tk.END)
                out_entry.insert(0, d)

        theme.make_button(card, "Browse", command=_browse_default_output, style="secondary").grid(
            row=2, column=2, padx=(4, PAD), pady=(0, 6),
        )

        # Appearance mode toggle
        theme.make_label(card, "Appearance:", level="body").grid(
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
        tk.Label(
            card,
            text="Appearance changes require a restart to take effect.",
            font=(theme.FONT_FAMILY, theme.FONT_SMALL[1] - 1),
            fg=p.text_muted,
            bg=p.bg_card,
            anchor="w",
        ).grid(row=4, column=0, columnspan=3, sticky="w", padx=PAD, pady=(0, 8))

        # Inline validation error
        error_label = tk.Label(
            card,
            text="",
            font=(theme.FONT_FAMILY, theme.FONT_SMALL[1]),
            fg=p.status_error,
            bg=p.bg_card,
            anchor="w",
        )
        error_label.grid(row=5, column=0, columnspan=3, sticky="w", padx=PAD, pady=(0, 2))

        # Button row
        btn_frame = ctk.CTkFrame(card, fg_color="transparent")
        btn_frame.grid(row=6, column=0, columnspan=3, pady=(4, PAD))

        def _save_and_close():
            inp = in_entry.get().strip()
            out = out_entry.get().strip()
            if not inp or not out:
                error_label.configure(text="Input or Output path cannot be empty")
                return
            error_label.configure(text="")
            self.cfg["default_input"] = inp
            self.cfg["default_output"] = out
            new_mode = "Dark" if mode_switch_var.get() else "Light"
            self.cfg["appearance_mode"] = new_mode
            save_config(self.cfg)
            win.destroy()

        theme.make_button(btn_frame, "Save", command=_save_and_close, style="primary").pack(
            side=tk.LEFT, padx=(0, 8),
        )
        theme.make_button(btn_frame, "Cancel", command=win.destroy, style="secondary").pack(
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
            self.output_entry.insert(0, d)

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

        # use typed path if present; else fall back to config default
        output = self.output_entry.get().strip()
        if not output:
            output = self.cfg.get("default_output") or str(Path.cwd() / "output")
            Path(str(output)).mkdir(parents=True, exist_ok=True)
        elif not Path(str(output)).is_dir():
            self._log(f"Error: Output path \u201c{output}\u201d does not exist \u2014 translation cancelled.")
            return

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

    root = ctk.CTk()
    app = App(root)  # noqa: F841
    root.mainloop()


if __name__ == "__main__":
    main()
