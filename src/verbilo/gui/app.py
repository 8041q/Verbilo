# gui_customtk.py — dashboard GUI; sidebar (controls) + content (file table, progress, log)

from __future__ import annotations

import os
import json
import queue
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
from .dnd import install_file_drop
from .config import load_config, save_config
from .icons import get_icon, get_photo_image, get_app_icon, apply_window_icon
from .dropdowns import SelectDropdown as _SelectDropdown, SearchableDropdown as _SearchableDropdown
from .i18n import DEFAULT_UI_LOCALE, get_supported_ui_locales, load_ui_localizer, resolve_ui_locale
from ..terminology import TerminologyEntry, TerminologyStore
from ..translation_memory import TranslationMemory, default_translation_memory_path

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

_ENGINE_LABEL_KEYS: list[tuple[str, str]] = [
    ("google", "engine.google"),
    ("google-cloud", "engine.google_cloud"),
    ("google-cloud-v3", "engine.google_cloud_v3"),
    ("baidu", "engine.baidu"),
    ("azure", "engine.azure"),
    ("deepl", "engine.deepl"),
    ("local", "engine.local"),
]


def _get_engine_options(localizer) -> list[tuple[str, str]]:
    return [(localizer.t(label_key), engine_key) for engine_key, label_key in _ENGINE_LABEL_KEYS]


def _get_engine_display_name(localizer, engine_key: str) -> str:
    for key, label_key in _ENGINE_LABEL_KEYS:
        if key == engine_key:
            return localizer.t(label_key)
    return engine_key


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


_cached_language_options: dict[str, list[tuple[str, str]]] = {}


def _get_language_options(locale: str | None = None) -> list[tuple[str, str]]:
    # returns (code, name) pairs for all supported languages
    resolved_locale = resolve_ui_locale(locale)
    cached = _cached_language_options.get(resolved_locale)
    if cached is not None:
        return cached

    base_localizer = load_ui_localizer(DEFAULT_UI_LOCALE)
    localizer = load_ui_localizer(resolved_locale)
    fallback = dict(base_localizer.language_names)
    canonical_codes = {code.lower(): code for code in fallback}

    def _canon(code: str) -> str:
        return canonical_codes.get(code.strip().lower(), code.strip())

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
            result = []
            seen_codes: set[str] = set()
            first_key = next(iter(langs), "")
            if len(first_key) <= 5 and first_key.isascii() and first_key.islower():
                codes = [str(k) for k in langs]
            else:
                codes = [str(v) for v in langs.values()]
            for code in codes:
                canonical = _canon(code)
                if not canonical or canonical in seen_codes:
                    continue
                seen_codes.add(canonical)
                result.append((canonical, localizer.language_name(canonical)))
            result.sort(key=lambda x: x[1].lower())
            _cached_language_options[resolved_locale] = result
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
                    code = name_to_code[low]
                elif entry in fallback:
                    code = entry
                else:
                    code = _canon(low)
                result.append((code, localizer.language_name(code)))
            result.sort(key=lambda x: x[1].lower())
            _cached_language_options[resolved_locale] = result
            return result

    except Exception:
        logger.exception("Failed to probe deep_translator for supported languages")

    result = sorted(
        ((code, localizer.language_name(code)) for code in fallback),
        key=lambda x: x[1].lower(),
    )
    _cached_language_options[resolved_locale] = result
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
    # Convert an OPUS model code (e.g. ``eng``, ``fra``) to the GUI code.
    normalized = (code or "").strip().lower().replace("_", "-")
    return _OPUS_CODE_MAP.get(normalized, normalized)


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


def _get_local_language_options(codes: set[str], locale: str | None = None) -> list[tuple[str, str]]:
    """Build display options directly from installed local-model codes.

    Local models must not disappear merely because ``deep_translator`` does not
    list the same language. Use its names when available, then fall back to the
    app localizer (and ultimately to the code itself).
    """
    normalized_codes = {_opus_code_to_iso(code) for code in codes if code}
    known_names = dict(_get_language_options(locale))
    localizer = load_ui_localizer(resolve_ui_locale(locale))
    result: list[tuple[str, str]] = []
    for code in normalized_codes:
        name = known_names.get(code)
        if not name:
            try:
                name = localizer.language_name(code)
            except Exception:
                name = code
        if not name:
            name = code
        result.append((code, name))
    result.sort(key=lambda item: (item[1].lower(), item[0]))
    return result


# --- searchable dropdown ---

def _install_tree_hover(tree, hover_color: str) -> None:
    """Add non-invasive row hover feedback to a ttk.Treeview."""
    tree.tag_configure("_hover", background=hover_color)
    state = {"row": None}

    def _set_row(row):
        previous = state["row"]
        if previous == row:
            return
        state["row"] = row
        if previous and tree.exists(previous):
            tags = tuple(tag for tag in tree.item(previous, "tags") if tag != "_hover")
            tree.item(previous, tags=tags)
        if row and tree.exists(row):
            tags = tuple(tree.item(row, "tags"))
            if "_hover" not in tags:
                tree.item(row, tags=tags + ("_hover",))

    tree.bind("<Motion>", lambda event: _set_row(tree.identify_row(event.y) or None), "+")
    tree.bind("<Leave>", lambda _event: _set_row(None), "+")


def _configure_app_table_styles(p) -> None:
    """Configure the shared table style used by the queue and manager dialogs."""
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass

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
            ("active", p.border),
            ("!active", p.divider),
            ("disabled", p.bg_card),
        ],
        arrowcolor=[("disabled", p.bg_card)],
    )


def _style_app_table(tree, p) -> None:
    """Apply the main queue table look to any application Treeview."""
    _configure_app_table_styles(p)
    tree.configure(style="FileTable.Treeview")
    tree.tag_configure("even", background=p.bg_row_even)
    tree.tag_configure("odd", background=p.bg_row_odd)
    _install_tree_hover(tree, p.bg_heading)


def _restripe_tree(tree) -> None:
    """Refresh alternating row tags without disturbing semantic/hover tags."""
    for index, iid in enumerate(tree.get_children()):
        tags = [tag for tag in tree.item(iid, "tags") if tag not in ("even", "odd")]
        tags.insert(0, "even" if index % 2 == 0 else "odd")
        tree.item(iid, tags=tuple(tags))


class SimpleComboBox(_SelectDropdown):
    """Backward-compatible name for the shared read-only dropdown control."""


class SearchableComboBox(_SearchableDropdown):
    """Backward-compatible name for the shared searchable dropdown control."""


# --- main app ---


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.worker = Worker()
        self.files: list[str] = []
        self._file_status: dict[str, str] = {}
        self._file_attempts: dict[str, int] = {}
        self._active_run_files: tuple[str, ...] = ()
        self.cfg = load_config() or {}
        self.ui = load_ui_localizer(self.cfg.get("ui_locale"))
        self.t = self.ui.t
        self.total_files = 0
        self.completed_files = 0
        self._running = False
        self._ollama_pull_lock = threading.Lock()
        self._ollama_install_cancel: threading.Event | None = None
        self._terminology_store = TerminologyStore()
        self._last_run_report: dict | None = None
        self._last_run_cancelled = False
        self._closing = False

        # Language-list refreshes can be requested from dropdown callbacks.
        # Keep them out of the originating popup event and prevent trace-driven
        # nested refreshes while source/target values are being synchronized.
        self._language_refresh_after_id = None
        self._refreshing_language_dropdowns = False

        # Thread-safe log queue: worker threads put messages here; main thread drains it
        self._log_queue: queue.SimpleQueue = queue.SimpleQueue()

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
        self._install_keyboard_shortcuts()
        self._drop_registration = self._install_drag_and_drop()
        self._sync_file_empty_state()
        try:
            self.root.protocol("WM_DELETE_WINDOW", self._request_close)
        except Exception:
            pass

        def _reapply_icon():
            try:
                apply_window_icon(self.root, size=64)
            except Exception:
                pass

        # Schedule at multiple points to beat any late CTk icon re-application
        self.root.after(100, _reapply_icon)
        self.root.after(500, _reapply_icon)

        # Start log queue drain loop (runs every 50 ms on the main thread)
        self._poll_log_queue()

        # Update check state (populated by background thread on startup)
        self._update_check_result: dict | None = None
        if self.cfg.get("auto_check_updates", True):
            self._run_update_check(startup=True)

        # Apply defaults from config
        default_out = self.cfg.get("default_output") or DEFAULT_OUTPUT_FOLDER
        self.output_entry.insert(0, default_out)
        default_in = self.cfg.get("default_input")
        
        if default_in:
            self._add_paths([default_in], announce=False)

        # Apply saved translation engine so language dropdowns reflect it
        try:
            self._on_engine_changed()
        except Exception:
            logger.exception("Failed to apply saved translation engine on startup")

    def _format_language_option(self, code: str, name: str | None = None) -> str:
        display_name = name if name is not None else self.ui.language_name(code)
        return f"{display_name} ({code})"

    def _engine_key_for_display(self, display_name: str) -> str:
        return self._engine_map.get(display_name, "google")

    def _engine_display_name(self, engine_key: str) -> str:
        return _get_engine_display_name(self.ui, engine_key)

    def _status_text(self, status: str) -> str:
        status_key = {
            "pending": "table.status.pending",
            "started": "table.status.translating",
            "finished": "table.status.finished",
            "error": "table.status.error",
            "cancelled": "table.status.cancelled",
            "skipped": "table.status.skipped",
            "retrying": "table.status.retrying",
            "translating": "table.status.translating",
        }.get(status)
        if status_key is None:
            return status
        return self.t(status_key)

    def _install_keyboard_shortcuts(self) -> None:
        """Install non-destructive application shortcuts without stealing text input."""
        bindings = {
            "<Control-o>": self._add_files,
            "<Control-Shift-o>": self._select_folder,
            "<Control-comma>": self._open_settings,
            "<Control-g>": self._open_terminology_manager,
            "<Control-Return>": self._start,
            "<F1>": self._open_about,
            # macOS equivalents; harmless on other Tk platforms.
            "<Command-o>": self._add_files,
            "<Command-Shift-o>": self._select_folder,
            "<Command-comma>": self._open_settings,
            "<Command-g>": self._open_terminology_manager,
            "<Command-Return>": self._start,
        }
        for sequence, callback in bindings.items():
            def _invoke(_event=None, _callback=callback, _sequence=sequence):
                try:
                    _callback()
                except Exception:
                    logger.exception("Keyboard shortcut failed: %s", _sequence)
                return "break"
            try:
                self.root.bind(sequence, _invoke, "+")
            except Exception:
                pass

    def _bind_modal_keys(self, window, close_callback, accept_callback=None, initial_focus=None) -> None:
        """Give modal dialogs predictable Escape/keyboard behavior and initial focus."""
        def _close(_event=None):
            close_callback()
            return "break"

        try:
            window.bind("<Escape>", _close, "+")
        except Exception:
            pass
        if accept_callback is not None:
            def _accept(_event=None):
                accept_callback()
                return "break"
            for sequence in ("<Control-Return>", "<Command-Return>"):
                try:
                    window.bind(sequence, _accept, "+")
                except Exception:
                    pass
        if initial_focus is not None:
            def _focus():
                try:
                    initial_focus.focus_set()
                except Exception:
                    pass
            try:
                window.after(80, _focus)
            except Exception:
                pass

    def _destroy_root(self) -> None:
        try:
            registration = getattr(self, "_drop_registration", None)
            if registration is not None:
                registration.close()
        except Exception:
            pass
        try:
            self.root.destroy()
        except Exception:
            pass

    def _request_close(self) -> None:
        if getattr(self, "_closing", False):
            return
        if getattr(self, "_running", False):
            try:
                confirmed = messagebox.askyesno(
                    self.t("app.exit_running_title"),
                    self.t("app.exit_running_body"),
                    parent=self.root,
                )
            except Exception:
                confirmed = False
            if not confirmed:
                return
            self._closing = True
            try:
                self.worker.stop()
            except Exception:
                pass
            try:
                self._update_progress_label(self.t("app.cancelling_exit"))
                self._set_button_disabled(self.start_btn, True)
                self._set_button_disabled(self.cancel_btn, True)
            except Exception:
                pass

            def _wait_for_worker():
                try:
                    if self.worker.alive:
                        self.root.after(50, _wait_for_worker)
                        return
                except Exception:
                    pass
                self._destroy_root()

            try:
                self.root.after(20, _wait_for_worker)
            except Exception:
                self._destroy_root()
            return
        self._closing = True
        self._destroy_root()

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
            title_badge = theme.make_label(
                title_frame, self.t("sidebar.beta"), level="tiny", text_color=p.text_muted,
            )
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
            self.sidebar, self.t("sidebar.source_language"), level="small",
        )
        self._source_lang_label.grid(row=row, column=0, sticky="w", padx=PAD, pady=(4, 2))
        row += 1

        self._engine_options = _get_engine_options(self.ui)
        self._engine_display = [name for name, _ in self._engine_options]
        self._engine_map = {name: key for name, key in self._engine_options}
        self._engine_reverse = {key: name for name, key in self._engine_options}

        lang_opts = _get_language_options(self.ui.locale)
        self._lang_map = {self._format_language_option(code, name): code for code, name in lang_opts}
        display_values = list(self._lang_map.keys())
        if not display_values:
            display_values = [self._format_language_option("en")]

        # Source language: filtered to languages the initial detector (fasttext) can identify.
        _src_filtered = _filter_by_detector(lang_opts, "fasttext")
        _src_display = [self._format_language_option(code, name) for code, name in _src_filtered]
        self._source_lang_label.configure(
            text=self.t("sidebar.source_language_count", count=len(_src_display)),
        )
        auto_detect_label = self.t("sidebar.auto_detect")
        self.source_lang_var = tk.StringVar(value=auto_detect_label)
        source_values = [auto_detect_label] + _src_display
        self._source_lang_map = {auto_detect_label: "auto"}
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
            self.sidebar, self.t("sidebar.target_language"), level="small",
        ).grid(row=row, column=0, sticky="w", padx=PAD, pady=(6, 2))
        row += 1

        default_target = self._format_language_option("en")
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
            self.sidebar, self.t("sidebar.language_detector"), level="small",
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
            self.sidebar, self.t("sidebar.translation_engine"), level="small",
        ).grid(row=row, column=0, sticky="w", padx=PAD, pady=(6, 2))
        row += 1

        saved_engine = self.cfg.get("translation_engine", "google")
        default_engine_display = self._engine_reverse.get(saved_engine, self._engine_display[0])
        self.engine_var = tk.StringVar(value=default_engine_display)
        self.engine_menu = SimpleComboBox(
            self.sidebar,
            values=self._engine_display,
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
        self._update_usage_label(self._engine_key_for_display(default_engine_display))
      
        # Spacer row (pushes everything to bottom)
        self.sidebar.grid_rowconfigure(row, weight=1)
        row += 1

        # OUTPUT section
        theme.make_label(
            self.sidebar, self.t("sidebar.output_folder"), level="small",
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
        self.output_browse_btn = theme.make_button(
            _out_frame, "", command=self._select_output, style="secondary",
            image=browse_icon, height=30, width=40,
        )
        self.output_browse_btn.grid(row=0, column=1, sticky="ew", padx=(4, 0))

        # Divider
        theme.make_divider(self.sidebar).grid(
            row=row, column=0, sticky="ew", padx=PAD, pady=4,
        )
        row += 1

        # Action buttons
        play_icon = get_icon("play", size=16, on_accent=True)
        self.start_btn = theme.make_button(
            self.sidebar, self.t("sidebar.start_translation"), command=self._start, style="primary",
            height=38, image=play_icon,
        )
        self.start_btn.grid(row=row, column=0, sticky="ew", padx=PAD, pady=(8, 4))
        row += 1

        stop_icon = get_icon("stop", size=16)
        self.cancel_btn = theme.make_button(
            self.sidebar, self.t("sidebar.cancel"), command=self._cancel, style="secondary",
            height=32, state="disabled", image=stop_icon,
        )
        self.cancel_btn.grid(row=row, column=0, sticky="ew", padx=PAD, pady=(0, 4))
        row += 1

        # Utility actions at the bottom share the same icon/text layout.
        theme.make_divider(self.sidebar).grid(
            row=row, column=0, sticky="ew", padx=PAD, pady=4,
        )
        row += 1

        terminology_icon = get_icon("terminology", size=16)
        self.terminology_btn = theme.make_button(
            self.sidebar, self.t("sidebar.terminology"), command=self._open_terminology_manager,
            style="ghost", anchor="w", image=terminology_icon,
        )
        self.terminology_btn.grid(row=row, column=0, sticky="ew", padx=PAD, pady=(4, 2))
        row += 1

        settings_icon = get_icon("settings", size=16)
        self.settings_btn = theme.make_button(
            self.sidebar, self.t("sidebar.settings"), command=self._open_settings, style="ghost",
            anchor="w", image=settings_icon,
        )
        self.settings_btn.grid(row=row, column=0, sticky="ew", padx=PAD, pady=(0, 2))
        row += 1

        info_icon = get_icon("info", size=16)
        self.about_btn = theme.make_button(
            self.sidebar, self.t("sidebar.about"), command=self._open_about, style="ghost",
            anchor="w", image=info_icon,
        )
        self.about_btn.grid(row=row, column=0, sticky="ew", padx=PAD, pady=(0, PAD))

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

        theme.make_label(toolbar, self.t("content.files"), level="subheading").grid(
            row=0, column=0, sticky="w", padx=PAD, pady=8,
        )
        btn_frame = ctk.CTkFrame(toolbar, fg_color="transparent")
        btn_frame.grid(row=0, column=1, sticky="e", padx=PAD, pady=8)

        add_icon = get_icon("add-file", size=16, on_accent=True)
        folder_icon = get_icon("open-folder", size=16, on_accent=True)
        trash_icon = get_icon("trash", size=16)

        self.add_files_btn = theme.make_button(
            btn_frame, self.t("content.add_files"), command=self._add_files, style="primary",
            image=add_icon, height=30,
        )
        self.add_files_btn.pack(side=tk.LEFT, padx=(0, 6))
        self.select_folder_btn = theme.make_button(
            btn_frame, self.t("content.select_folder"), command=self._select_folder, style="primary",
            image=folder_icon, height=30,
        )
        self.select_folder_btn.pack(side=tk.LEFT, padx=(0, 6))
        self.clear_files_btn = theme.make_button(
            btn_frame, self.t("content.clear"), command=self._clear_files, style="secondary",
            image=trash_icon, height=30,
        )
        self.clear_files_btn.pack(side=tk.LEFT)

        # File table card
        table_card = theme.make_card(content)
        table_card.grid(row=1, column=0, sticky="nsew", pady=(0, PAD))
        self._build_file_table(table_card)

        # Progress card
        progress_card = theme.make_card(content)
        progress_card.grid(row=2, column=0, sticky="ew", pady=(0, PAD))
        progress_card.grid_columnconfigure(0, weight=1)
        progress_card.grid_columnconfigure(1, weight=0)
        progress_card.grid_columnconfigure(2, weight=0)

        self.progress_label = theme.make_label(
            progress_card, self.t("content.ready"), level="small",
        )
        self.progress_label.grid(
            row=0, column=0, sticky="w", padx=PAD, pady=(10, 4),
        )
        self.report_btn = theme.make_button(
            progress_card, self.t("report.view"), command=self._open_translation_report,
            style="ghost", height=24, state="disabled",
        )
        self.report_btn.grid(row=0, column=1, sticky="e", padx=(4, 4), pady=(6, 2))
        # A report is a post-run artifact. Keep the button completely out of the
        # layout until a completed run has produced one.
        self.report_btn.grid_remove()

        self.retry_failed_btn = theme.make_button(
            progress_card, self.t("queue.retry_failed"), command=self._retry_failed,
            style="ghost", height=24, state="disabled",
        )
        self.retry_failed_btn.grid(row=0, column=2, sticky="e", padx=(4, PAD), pady=(6, 2))
        self.retry_failed_btn.grid_remove()

        self.progress = ctk.CTkProgressBar(
            progress_card,
            progress_color=p.accent,
            fg_color=p.bg_main,
            corner_radius=4,
            height=8,
        )
        self.progress.grid(row=1, column=0, columnspan=3, sticky="ew", padx=PAD, pady=(0, 10))
        self._set_progress(0.0)

        # Log card
        log_card = theme.make_card(content)
        log_card.grid(row=3, column=0, sticky="nsew")
        log_card.grid_columnconfigure(0, weight=1)
        log_card.grid_rowconfigure(1, weight=1)

        theme.make_label(log_card, self.t("content.log"), level="subheading").grid(
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
        _configure_app_table_styles(p)

        # Load file-type icons for the treeview (PhotoImage for ttk)
        icon_color = p.text_muted
        self._file_icons: dict[str, object] = {}
        for ext, icon_name in ((".docx", "file-docx"), (".pdf", "file-pdf"),
                                (".xlsx", "file-xls")):
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
            selectmode="extended",
            takefocus=True,
        )
        self.file_table.heading("status", text=self.t("table.status"), anchor="center")
        self.file_table.heading("time", text=self.t("table.time"), anchor="center")

        self.file_table["show"] = ("tree", "headings")
        self.file_table.heading("#0", text=self.t("table.file"), anchor="w")
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

        self._file_empty_state = theme.make_label(
            container, self.t("content.empty_queue"), level="small",
            anchor="center", justify="center", text_color=p.text_muted,
        )
        self._file_empty_state.place(relx=0.5, rely=0.5, anchor="center")

        def _delete_selected(_event=None):
            self._remove_selected_files()
            return "break"

        def _select_all(_event=None):
            self._select_all_queue_rows()
            return "break"

        self.file_table.bind("<Delete>", _delete_selected, "+")
        self.file_table.bind("<Control-a>", _select_all, "+")
        self.file_table.bind("<Control-A>", _select_all, "+")
        self.file_table.bind("<Command-a>", _select_all, "+")
        self.file_table.bind("<<TreeviewSelect>>", lambda _e: self._sync_clear_button_label(), "+")

        # Deselect when clicking on empty space in the table
        def _on_table_click(event):
            if not self.file_table.identify_row(event.y):
                self.file_table.selection_set([])
        self.file_table.bind("<Button-1>", _on_table_click, "+")

        # Lightweight row hover makes the table feel interactive without
        # changing the selected/status foreground colours.
        self.file_table.tag_configure("hover", background=p.bg_heading)
        self._hovered_file_row = None

        def _on_table_motion(event):
            row = self.file_table.identify_row(event.y) or None
            if row == self._hovered_file_row:
                return
            previous = self._hovered_file_row
            self._hovered_file_row = row
            if previous and self.file_table.exists(previous):
                tags = tuple(tag for tag in self.file_table.item(previous, "tags") if tag != "hover")
                self.file_table.item(previous, tags=tags)
            if row and self.file_table.exists(row):
                tags = tuple(self.file_table.item(row, "tags"))
                if "hover" not in tags:
                    self.file_table.item(row, tags=tags + ("hover",))

        def _on_table_leave(_event=None):
            previous = self._hovered_file_row
            self._hovered_file_row = None
            if previous and self.file_table.exists(previous):
                tags = tuple(tag for tag in self.file_table.item(previous, "tags") if tag != "hover")
                self.file_table.item(previous, tags=tags)

        self.file_table.bind("<Motion>", _on_table_motion, "+")
        self.file_table.bind("<Leave>", _on_table_leave, "+")

        # Status colour tags
        self.file_table.tag_configure("pending",   foreground=p.status_pending)
        self.file_table.tag_configure("started",   foreground=p.status_info)
        self.file_table.tag_configure("finished",  foreground=p.status_success)
        self.file_table.tag_configure("error",     foreground=p.status_error)
        self.file_table.tag_configure("cancelled", foreground=p.status_warning)
        self.file_table.tag_configure("skipped", foreground=p.status_warning)
        self.file_table.tag_configure("retrying", foreground=p.status_info)
        self.file_table.tag_configure("even", background=p.bg_row_even)
        self.file_table.tag_configure("odd",  background=p.bg_row_odd)
        self._sync_file_empty_state()
        self._sync_clear_button_label()

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

    def _sync_file_empty_state(self) -> None:
        label = getattr(self, "_file_empty_state", None)
        table = getattr(self, "file_table", None)
        if label is None or table is None:
            return
        try:
            if table.get_children():
                label.place_forget()
            else:
                label.place(relx=0.5, rely=0.5, anchor="center")
                label.lift()
        except Exception:
            pass

    def _sync_clear_button_label(self) -> None:
        button = getattr(self, "clear_files_btn", None)
        table = getattr(self, "file_table", None)
        if button is None or table is None:
            return
        try:
            key = "content.remove_selected" if table.selection() else "content.clear_all"
            button.configure(text=self.t(key))
        except Exception:
            pass

    def _select_all_queue_rows(self) -> None:
        if getattr(self, "_running", False):
            return
        try:
            rows = self.file_table.get_children()
            if rows:
                self.file_table.selection_set(rows)
                self.file_table.focus(rows[0])
                self._sync_clear_button_label()
        except Exception:
            pass

    def _remove_selected_files(self) -> None:
        """Remove only selected queue rows; never interpret no selection as clear-all."""
        if getattr(self, "_running", False):
            return
        try:
            selected = tuple(self.file_table.selection())
        except Exception:
            selected = ()
        if not selected:
            return
        for iid in selected:
            filepath = self._tree_ids.pop(iid, None)
            if filepath:
                self._file_to_iid.pop(filepath, None)
                self._file_status.pop(filepath, None)
                self._file_attempts.pop(filepath, None)
                self._file_start_times.pop(filepath, None)
                if filepath in self.files:
                    self.files.remove(filepath)
            try:
                self.file_table.delete(iid)
            except Exception:
                pass
        self._retag_rows()
        self._sync_file_empty_state()
        self._sync_clear_button_label()
        self._sync_retry_button_visibility()

    def _add_file_to_table(self, filepath: str, status: str = "pending"):
        self.files.append(filepath)
        self._file_status[filepath] = status
        self._file_attempts.setdefault(filepath, 0)
        name = os.path.basename(filepath)
        idx = len(self.files) - 1
        row_tag = "even" if idx % 2 == 0 else "odd"
        icon = self._get_file_icon(filepath)
        kw: dict[str, object] = {}
        if icon:
            kw["image"] = icon
        iid = self.file_table.insert(
            "", tk.END, text=name, values=(self._status_text(status), ""), tags=(status, row_tag), **kw,
        )
        self._tree_ids[iid] = filepath
        self._file_to_iid[filepath] = iid
        self._sync_file_empty_state()
        self._sync_clear_button_label()

    def _update_file_status(self, filepath: str, status: str, elapsed: float | None = None):
        self._file_status[filepath] = status
        iid = self._file_to_iid.get(filepath)
        if iid is None:
            return
        name = os.path.basename(filepath)
        time_str = self._format_elapsed_time(elapsed)
        idx = list(self._file_to_iid.keys()).index(filepath)
        row_tag = "even" if idx % 2 == 0 else "odd"
        self.file_table.item(iid, text=name, values=(self._status_text(status), time_str), tags=(status, row_tag))

    def _update_all_statuses(self, status: str):
        for filepath, iid in self._file_to_iid.items():
            self._file_status[filepath] = status
            name = os.path.basename(filepath)
            idx = list(self._file_to_iid.keys()).index(filepath)
            row_tag = "even" if idx % 2 == 0 else "odd"
            self.file_table.item(iid, text=name, values=(self._status_text(status), ""), tags=(status, row_tag))

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
        win.title(self.t("settings.title"))
        win.transient(self.root)
        win.grab_set()
        win.configure(fg_color=p.bg_main)

        win.grid_columnconfigure(0, weight=1)
        win.grid_rowconfigure(0, weight=1)

        # Outer card wrapper — holds title + two-column body
        card = theme.make_card(win)
        card.configure(width=theme.scale(900))
        win.minsize(theme.scale(900), theme.scale(650))
        card.grid(row=0, column=0, sticky="nsew", padx=PAD, pady=PAD)
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(1, weight=1)

        # ── Title row ────────────────────────────────────────────────────
        settings_icon = get_icon("settings", size=20)
        title_frame = ctk.CTkFrame(card, fg_color="transparent")
        title_frame.grid(row=0, column=0, sticky="w", padx=PAD, pady=(PAD, PAD))
        if settings_icon:
            ctk.CTkLabel(title_frame, text="", image=settings_icon, width=20).pack(side=tk.LEFT, padx=(0, 8))
        theme.make_label(title_frame, self.t("settings.title"), level="heading").pack(side=tk.LEFT)

        # ── Two-column body ───────────────────────────────────────────────
        body = ctk.CTkFrame(card, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew", padx=PAD, pady=(0, PAD))
        # col 0 = left, col 1 = vertical divider (fixed width), col 2 = right
        body.grid_columnconfigure(0, weight=0, minsize=theme.scale(320))
        body.grid_columnconfigure(1, weight=0, minsize=theme.scale(1))
        body.grid_columnconfigure(2, weight=1, minsize=theme.scale(460))
        body.grid_rowconfigure(0, weight=1)

        # Vertical divider — use a plain tk.Frame so the fixed 1-px width is respected
        p_now = theme.get()
        _vdiv = tk.Frame(body, width=1, bg=p_now.divider)
        _vdiv.grid(row=0, column=1, sticky="ns", padx=(PAD, PAD))
        _vdiv.grid_propagate(False)

        # ── LEFT COLUMN: Folders + Appearance + Updates + Debug ──────────
        def _hide_settings_scrollbar(frame):
            """Hide CTk's visual scrollbar without disabling wheel/trackpad scrolling."""
            scrollbar = getattr(frame, "_scrollbar", None)
            if scrollbar is None:
                return
            # CTkScrollableFrame keeps scrolling on its canvas; the scrollbar is
            # only a visual/drag control, so removing it from the geometry manager
            # preserves mouse-wheel and trackpad scrolling.
            for forget in ("grid_remove", "grid_forget", "pack_forget", "place_forget"):
                method = getattr(scrollbar, forget, None)
                if method is None:
                    continue
                try:
                    method()
                    break
                except Exception:
                    continue

        try:
            left = ctk.CTkScrollableFrame(body, fg_color="transparent")
            left.configure(width=theme.scale(320), height=theme.scale(390))
            _hide_settings_scrollbar(left)
        except Exception:
            left = ctk.CTkFrame(body, fg_color="transparent")
        left.grid(row=0, column=0, sticky="nsew")
        left.grid_columnconfigure(0, weight=1)

        # A CTkScrollableFrame does not propagate its children's requested width
        # the same way the old plain CTkFrame did. Keep the left pane readable and
        # explicitly wrap its labels to the *actual* available width so translated
        # UI strings do not get clipped at higher DPI or with longer locales.
        _left_wrapped_labels: list[object] = []

        def _left_settings_label(text: str, *, level: str = "body", **kwargs):
            label = theme.make_label(left, text, level=level, **kwargs)
            try:
                label.configure(anchor="w", justify="left")
            except Exception:
                pass
            _left_wrapped_labels.append(label)
            return label

        def _sync_left_settings_wrap(_event=None):
            try:
                available = int(left.winfo_width()) - theme.scale(24)
            except Exception:
                available = theme.scale(296)
            wrap = max(theme.scale(220), available)
            for label in tuple(_left_wrapped_labels):
                try:
                    if label.winfo_exists():
                        label.configure(wraplength=wrap)
                except Exception:
                    pass

        try:
            left.bind("<Configure>", _sync_left_settings_wrap, add="+")
        except Exception:
            pass

        _lrow = 0

        # FOLDERS section
        _left_settings_label(self.t("settings.section.folders"), level="section").grid(
            row=_lrow, column=0, columnspan=2, sticky="w", pady=(0, 4),
        )
        _lrow += 1

        # Default input folder
        _left_settings_label(self.t("settings.default_input_folder"), level="small").grid(
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
            d = filedialog.askdirectory(title=self.t("dialog.select_default_input_folder"), parent=win, initialdir=str(candidate))
            if d:
                in_entry.delete(0, tk.END)
                in_entry.insert(0, str(Path(d).resolve()))

        browse_icon_s = get_icon("folder", size=14)
        theme.make_button(_input_row, self.t("settings.browse"), command=_browse_default_input, style="secondary",
                          image=browse_icon_s, height=28).grid(row=0, column=1, padx=(6, 0))
        _lrow += 1

        # Default output folder
        _left_settings_label(self.t("settings.default_output_folder"), level="small").grid(
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
            d = filedialog.askdirectory(title=self.t("dialog.select_default_output_folder"), parent=win, initialdir=str(candidate))
            if d:
                out_entry.delete(0, tk.END)
                out_entry.insert(0, str(Path(d).resolve()))

        theme.make_button(_output_row, self.t("settings.browse"), command=_browse_default_output, style="secondary",
                          image=browse_icon_s, height=28).grid(row=0, column=1, padx=(6, 0))
        _lrow += 1

        # Divider
        theme.make_divider(left).grid(row=_lrow, column=0, columnspan=2, sticky="ew", pady=(10, 8))
        _lrow += 1

        # APPEARANCE section
        _left_settings_label(self.t("settings.section.appearance"), level="section").grid(
            row=_lrow, column=0, columnspan=2, sticky="w", pady=(0, 4),
        )
        _lrow += 1

        mode_switch_var = tk.BooleanVar(value=(theme.get_mode() == "Dark"))
        mode_switch = ctk.CTkSwitch(
            left,
            text=self.t("settings.appearance.dark_mode") if theme.get_mode() == "Dark" else self.t("settings.appearance.light_mode"),
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
            mode_switch.configure(text=self.t("settings.appearance.dark_mode") if mode_switch_var.get() else self.t("settings.appearance.light_mode"))

        mode_switch_var.trace_add("write", _on_mode_switch)

        _left_settings_label(
            self.t("settings.appearance.restart_required"),
            level="tiny",
        ).grid(row=_lrow, column=0, columnspan=2, sticky="ew", pady=(0, 2))
        _lrow += 1

        # Divider
        theme.make_divider(left).grid(row=_lrow, column=0, columnspan=2, sticky="ew", pady=(10, 8))
        _lrow += 1

        # LANGUAGE section
        _left_settings_label(self.t("settings.section.language"), level="section").grid(
            row=_lrow, column=0, columnspan=2, sticky="w", pady=(0, 4),
        )
        _lrow += 1

        _locale_options = get_supported_ui_locales(self.ui.locale)
        _locale_display = [name for _, name in _locale_options]
        _current_locale = resolve_ui_locale(self.cfg.get("ui_locale"))
        _initial_locale_display = next(
            (name for code, name in _locale_options if code == _current_locale),
            _locale_display[0],
        )
        ui_lang_var = tk.StringVar(value=_initial_locale_display)
        SimpleComboBox(
            left,
            values=_locale_display,
            variable=ui_lang_var,
        ).grid(row=_lrow, column=0, columnspan=2, sticky="ew", pady=(0, 4))
        _lrow += 1

        _left_settings_label(
            self.t("settings.language.restart_required"),
            level="tiny",
        ).grid(row=_lrow, column=0, columnspan=2, sticky="ew", pady=(0, 2))
        _lrow += 1

        # Divider
        theme.make_divider(left).grid(row=_lrow, column=0, columnspan=2, sticky="ew", pady=(10, 8))
        _lrow += 1

        # UPDATES section
        _left_settings_label(self.t("settings.section.updates"), level="section").grid(
            row=_lrow, column=0, columnspan=2, sticky="w", pady=(0, 4),
        )
        _lrow += 1

        auto_updates_var = tk.BooleanVar(value=self.cfg.get("auto_check_updates", True))
        auto_updates_cb = ctk.CTkCheckBox(
            left,
            text=self.t("settings.auto_check_updates"),
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
        _left_settings_label(self.t("settings.section.debug"), level="section").grid(
            row=_lrow, column=0, columnspan=2, sticky="w", pady=(0, 4),
        )
        _lrow += 1

        debug_var = tk.BooleanVar(value=bool(self.cfg.get("debug_mode", False)))
        debug_cb = ctk.CTkCheckBox(
            left,
            text=self.t("settings.debug_mode"),
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
        _left_settings_label(self.t("settings.section.local_models"), level="section").grid(
            row=_lrow, column=0, columnspan=2, sticky="w", pady=(0, 4),
        )
        _lrow += 1

        def _open_manager_from_settings():
            win.destroy()
            self.root.after(50, self._open_model_manager)

        theme.make_button(
            left, self.t("settings.open_model_manager"),
            command=_open_manager_from_settings,
            style="secondary", height=28,
        ).grid(row=_lrow, column=0, columnspan=2, sticky="w", pady=(0, 4))
        _lrow += 1

        try:
            left.after_idle(_sync_left_settings_wrap)
        except Exception:
            pass

        # ── RIGHT COLUMN: Network + API Keys ─────────────────────────────
        # Use a scrollable frame for the right column
        try:
            right = ctk.CTkScrollableFrame(body, fg_color="transparent")
            try:
                right.configure(width=theme.scale(460), height=theme.scale(350))
            except Exception:
                pass
            _hide_settings_scrollbar(right)
        except Exception:
            right = ctk.CTkFrame(body, fg_color="transparent")
        right.grid(row=0, column=2, sticky="nsew")
        right.grid_columnconfigure(0, weight=1)

        # Keep the API/network side intentionally wider than the left settings pane.
        # CTkScrollableFrame does not reliably propagate child-requested width, so
        # explanatory text must wrap against the rendered pane width rather than a
        # fixed pixel value.
        _right_wrapped_labels: list[object] = []

        def _right_settings_label(text: str, *, level: str = "body", **kwargs):
            label = theme.make_label(right, text, level=level, **kwargs)
            try:
                label.configure(anchor="w", justify="left")
            except Exception:
                pass
            _right_wrapped_labels.append(label)
            return label

        def _sync_right_settings_wrap(_event=None):
            try:
                available = int(right.winfo_width()) - theme.scale(28)
            except Exception:
                available = theme.scale(430)
            wrap = max(theme.scale(300), available)
            for label in tuple(_right_wrapped_labels):
                try:
                    if label.winfo_exists():
                        label.configure(wraplength=wrap)
                except Exception:
                    pass

        try:
            right.bind("<Configure>", _sync_right_settings_wrap, add="+")
        except Exception:
            pass

        _rrow = 0

        # NETWORK section
        _right_settings_label(self.t("settings.section.network"), level="section").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 6),
        )
        _rrow += 1

        _right_settings_label(self.t("settings.https_proxy"), level="small").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 4),
        )
        _rrow += 1
        proxy_entry = theme.make_entry(right, height=28)
        proxy_entry.grid(row=_rrow, column=0, sticky="ew", pady=(0, 1))
        proxy_entry.insert(0, self.cfg.get("proxy_url", ""))
        _rrow += 1
        _right_settings_label(self.t("settings.proxy_hint"),
            level="tiny",
        ).grid(row=_rrow, column=0, sticky="w", pady=(0, 8))
        _rrow += 1

        # Divider
        theme.make_divider(right).grid(row=_rrow, column=0, sticky="ew", pady=(4, 8))
        _rrow += 1

        # Ollama / Qwen section
        _right_settings_label(self.t("settings.section.ollama"), level="section").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 6),
        )
        _rrow += 1

        ollama_enabled_var = tk.BooleanVar(value=bool(self.cfg.get("ollama_enabled", False)))
        ollama_enabled_cb = ctk.CTkCheckBox(
            right,
            text=self.t("settings.ollama_enabled"),
            variable=ollama_enabled_var,
            onvalue=True,
            offvalue=False,
            checkmark_color=p.bg_main,
            fg_color=p.accent,
            hover_color=p.accent_hover,
            border_color=p.border,
            text_color=p.text_secondary,
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=theme.FONT_BODY[1]),
        )
        ollama_enabled_cb.grid(row=_rrow, column=0, sticky="w", pady=(0, 4))
        _rrow += 1

        _right_settings_label(self.t("settings.ollama_base_url"), level="small").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 4),
        )
        _rrow += 1
        ollama_base_url_entry = theme.make_entry(right, height=28)
        ollama_base_url_entry.grid(row=_rrow, column=0, sticky="ew", pady=(0, 1))
        ollama_base_url_entry.insert(0, self.cfg.get("ollama_base_url", "http://127.0.0.1:11434"))
        _rrow += 1

        ollama_url_hint_lbl = _right_settings_label(self.t("settings.ollama_base_url_hint"), level="tiny",
        )
        ollama_url_hint_lbl.configure(anchor="w", justify="left")
        ollama_url_hint_lbl.grid(row=_rrow, column=0, sticky="ew", pady=(0, 6))
        _rrow += 1

        _right_settings_label(self.t("settings.ollama.runtime"), level="small").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 3),
        )
        _rrow += 1

        ollama_runtime_row = ctk.CTkFrame(right, fg_color="transparent")
        ollama_runtime_row.grid(row=_rrow, column=0, sticky="ew", pady=(0, 8))
        ollama_runtime_row.grid_columnconfigure(1, weight=1)
        ollama_runtime_dot = theme.make_label(
            ollama_runtime_row, "●", level="small", text_color=p.text_muted,
        )
        ollama_runtime_dot.grid(row=0, column=0, sticky="w", padx=(0, 6))
        ollama_runtime_status_lbl = theme.make_label(
            ollama_runtime_row, self.t("settings.ollama.runtime.checking"), level="small",
        )
        ollama_runtime_status_lbl.grid(row=0, column=1, sticky="w")
        _rrow += 1

        _right_settings_label(self.t("settings.ollama.model_heading"), level="small").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 4),
        )
        _rrow += 1

        _OLLAMA_MODEL_META = {
            "qwen3.5:4b": (
                self.t("settings.ollama.model.qwen"),
                self.t("settings.ollama.model_desc.qwen"),
            ),
            "demonbyron/HY-MT1.5-1.8B": (
                self.t("settings.ollama.model.hymt"),
                self.t("settings.ollama.model_desc.hymt"),
            ),
            "translategemma:4b": (
                self.t("settings.ollama.model.translategemma"),
                self.t("settings.ollama.model_desc.translategemma"),
            ),
            "mistral:7b": (
                self.t("settings.ollama.model.mistral"),
                self.t("settings.ollama.model_desc.mistral"),
            ),
        }
        _OLLAMA_DISPLAY_TO_ID = {display: model for model, (display, _desc) in _OLLAMA_MODEL_META.items()}
        _OLLAMA_ID_TO_DISPLAY = {model: display for model, (display, _desc) in _OLLAMA_MODEL_META.items()}
        _saved_ollama_model = self.cfg.get("ollama_model", "qwen3.5:4b")
        if _saved_ollama_model not in _OLLAMA_MODEL_META:
            _saved_ollama_model = "qwen3.5:4b"
        ollama_model_var = ctk.StringVar(value=_saved_ollama_model)
        ollama_model_display_var = ctk.StringVar(value=_OLLAMA_ID_TO_DISPLAY[_saved_ollama_model])

        ollama_model_box = SimpleComboBox(
            right,
            values=[meta[0] for meta in _OLLAMA_MODEL_META.values()],
            variable=ollama_model_display_var,
            command=lambda value: _on_ollama_model_selected(value),
        )
        ollama_model_box.grid(row=_rrow, column=0, sticky="ew", pady=(0, 4))
        _rrow += 1

        ollama_model_desc_lbl = _right_settings_label(
            _OLLAMA_MODEL_META[_saved_ollama_model][1], level="tiny",
        )
        ollama_model_desc_lbl.grid(row=_rrow, column=0, sticky="ew", pady=(0, 8))
        _rrow += 1

        ollama_model_card = theme.make_card(right, fg_color=p.bg_main)
        ollama_model_card.grid(row=_rrow, column=0, sticky="ew", pady=(0, 8))
        ollama_model_card.grid_columnconfigure(0, weight=1)

        ollama_model_state_lbl = theme.make_label(
            ollama_model_card, self.t("settings.ollama.state.checking"), level="section",
        )
        ollama_model_state_lbl.grid(row=0, column=0, sticky="w", padx=10, pady=(9, 1))
        ollama_model_detail_lbl = theme.make_label(
            ollama_model_card, "", level="tiny", text_color=p.text_muted,
        )
        ollama_model_detail_lbl.configure(anchor="w", justify="left")
        _right_wrapped_labels.append(ollama_model_detail_lbl)
        ollama_model_detail_lbl.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 7))

        ollama_model_progress = ctk.CTkProgressBar(
            ollama_model_card,
            progress_color=p.accent,
            fg_color=p.bg_card,
            corner_radius=4,
            height=7,
        )
        ollama_model_progress.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 7))
        ollama_model_progress.set(0)
        ollama_model_progress.grid_remove()

        ollama_model_actions = ctk.CTkFrame(ollama_model_card, fg_color="transparent")
        ollama_model_actions.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 9))
        ollama_model_actions.grid_columnconfigure(2, weight=1)

        ollama_running_note_lbl = theme.make_label(
            ollama_model_card, self.t("settings.ollama.running_note"),
            level="tiny", text_color=p.status_warning,
        )
        ollama_running_note_lbl.configure(anchor="w", justify="left")
        _right_wrapped_labels.append(ollama_running_note_lbl)
        ollama_running_note_lbl.grid(row=4, column=0, sticky="ew", padx=10, pady=(0, 8))
        if not self._running:
            ollama_running_note_lbl.grid_remove()

        ollama_ui_state = {
            "runtime": "checking",
            "model": "checking",
            "busy": False,
            "operation": None,
            "generation": 0,
            "error": None,
            "progress": None,
            "stage": None,
            "failed_action": None,
            "owns_pull_lock": False,
        }

        def _selected_ollama_model() -> str:
            return ollama_model_var.get() or "qwen3.5:4b"

        def _selected_ollama_display() -> str:
            model = _selected_ollama_model()
            return _OLLAMA_ID_TO_DISPLAY.get(model, model)

        def _ollama_proxies() -> dict | None:
            proxy_url = proxy_entry.get().strip()
            return {"https": proxy_url, "http": proxy_url} if proxy_url else None

        def _ollama_base_url() -> str:
            return ollama_base_url_entry.get().strip() or "http://127.0.0.1:11434"

        def _set_ollama_runtime_visual(state: str):
            mapping = {
                "checking": ("settings.ollama.runtime.checking", p.text_muted),
                "connected": ("settings.ollama.runtime.connected", p.status_success),
                "stopped": ("settings.ollama.runtime.stopped", p.status_warning),
                "not_installed": ("settings.ollama.runtime.not_installed", p.status_warning),
                "unavailable": ("settings.ollama.runtime.unavailable", p.status_error),
            }
            key, color = mapping.get(state, mapping["unavailable"])
            ollama_runtime_dot.configure(text_color=color)
            ollama_runtime_status_lbl.configure(text=self.t(key), text_color=color)

        def _ollama_stage_text(stage: str | None) -> str:
            key = {
                "manifest": "settings.ollama.progress.manifest",
                "layers": "settings.ollama.progress.layers",
                "verifying": "settings.ollama.progress.verifying",
                "writing_manifest": "settings.ollama.progress.writing_manifest",
                "cleanup": "settings.ollama.progress.cleanup",
                "downloading": "settings.ollama.progress.downloading",
            }.get(stage, "settings.ollama.progress.downloading")
            return self.t(key)

        def _set_action_button(button, *, text: str, command, style: str = "primary", visible: bool = True, enabled: bool = True):
            try:
                button.configure(text=text, command=command)
                # Preserve the existing button widget while refreshing style-specific colours.
                p_now = theme.get()
                if style == "primary":
                    button.configure(
                        fg_color=p_now.accent,
                        hover_color=p_now.accent_hover,
                        text_color=p_now.text_on_accent,
                        border_width=0,
                    )
                elif style == "secondary":
                    button.configure(
                        fg_color=p_now.bg_input,
                        hover_color=p_now.bg_heading,
                        text_color=p_now.text_secondary,
                        border_width=1,
                        border_color=p_now.border,
                    )
                else:
                    button.configure(
                        fg_color="transparent",
                        hover_color=p_now.bg_heading,
                        text_color=p_now.text_secondary,
                        border_width=0,
                    )
                self._set_button_disabled(button, not enabled)
                if visible:
                    button.grid()
                else:
                    button.grid_remove()
            except Exception:
                pass

        def _render_ollama_model_state():
            runtime = str(ollama_ui_state["runtime"])
            model_state = str(ollama_ui_state["model"])
            busy = bool(ollama_ui_state["busy"])
            operation = ollama_ui_state["operation"]
            translation_running = bool(self._running)
            try:
                if translation_running:
                    ollama_running_note_lbl.grid()
                else:
                    ollama_running_note_lbl.grid_remove()
            except Exception:
                pass
            error = ollama_ui_state.get("error")
            progress = ollama_ui_state.get("progress")
            stage = ollama_ui_state.get("stage")
            display = _selected_ollama_display()

            _set_ollama_runtime_visual(runtime)
            try:
                self._set_button_disabled(ollama_refresh_btn, busy)
            except Exception:
                pass
            try:
                ollama_model_box.configure(state="disabled" if busy else "normal")
                ollama_base_url_entry.configure(state="disabled" if busy else "normal")
            except Exception:
                pass

            if busy and operation == "download":
                if stage == "verifying":
                    title = self.t("settings.ollama.state.verifying")
                else:
                    title = self.t("settings.ollama.state.downloading", model=display)
                ollama_model_state_lbl.configure(text=title, text_color=p.status_info)
                stage_text = _ollama_stage_text(stage)
                if isinstance(progress, int):
                    detail = self.t("settings.ollama.progress.percent", stage=stage_text, percent=progress)
                    ollama_model_progress.set(max(0.0, min(1.0, progress / 100.0)))
                else:
                    detail = stage_text
                    ollama_model_progress.set(0.04)
                ollama_model_detail_lbl.configure(text=detail, text_color=p.text_muted)
                ollama_model_progress.grid()
                _set_action_button(
                    ollama_primary_action_btn,
                    text=self.t("settings.ollama.action.cancel"),
                    command=_cancel_ollama_download,
                    style="secondary",
                )
                _set_action_button(ollama_remove_btn, text=self.t("settings.ollama.action.remove"), command=_remove_ollama_model_from_settings, visible=False)
            elif busy and operation == "start":
                ollama_model_state_lbl.configure(text=self.t("settings.ollama.state.starting"), text_color=p.status_info)
                ollama_model_detail_lbl.configure(text=self.t("settings.ollama.runtime.checking"), text_color=p.text_muted)
                ollama_model_progress.grid_remove()
                _set_action_button(ollama_primary_action_btn, text=self.t("settings.ollama.action.start"), command=lambda: None, enabled=False)
                _set_action_button(ollama_remove_btn, text=self.t("settings.ollama.action.remove"), command=_remove_ollama_model_from_settings, visible=False)
            elif busy and operation == "remove":
                ollama_model_state_lbl.configure(text=self.t("settings.ollama.state.removing", model=display), text_color=p.status_info)
                ollama_model_detail_lbl.configure(text="", text_color=p.text_muted)
                ollama_model_progress.grid_remove()
                _set_action_button(ollama_primary_action_btn, text=self.t("settings.ollama.action.retry"), command=lambda: None, visible=False)
                _set_action_button(ollama_remove_btn, text=self.t("settings.ollama.action.remove"), command=lambda: None, enabled=False)
            else:
                ollama_model_progress.grid_remove()
                if error:
                    ollama_model_state_lbl.configure(text=self.t("settings.ollama.state.error"), text_color=p.status_error)
                    ollama_model_detail_lbl.configure(text=str(error), text_color=p.status_error)
                    failed_action = ollama_ui_state.get("failed_action")
                    retry_command = _start_ollama_runtime if failed_action == "start" else _pull_ollama_model_from_settings if failed_action == "download" else _refresh_ollama_status
                    _set_action_button(ollama_primary_action_btn, text=self.t("settings.ollama.action.retry"), command=retry_command, style="primary")
                    _set_action_button(ollama_remove_btn, text=self.t("settings.ollama.action.remove"), command=_remove_ollama_model_from_settings, visible=False)
                elif runtime == "connected" and model_state == "installed":
                    ollama_model_state_lbl.configure(text=self.t("settings.ollama.state.ready"), text_color=p.status_success)
                    ollama_model_detail_lbl.configure(text=self.t("settings.ollama.state.ready_detail"), text_color=p.text_muted)
                    _set_action_button(ollama_primary_action_btn, text=self.t("settings.ollama.action.download"), command=_pull_ollama_model_from_settings, visible=False)
                    _set_action_button(
                        ollama_remove_btn, text=self.t("settings.ollama.action.remove"),
                        command=_remove_ollama_model_from_settings, style="secondary",
                        visible=True, enabled=not translation_running,
                    )
                elif runtime == "connected" and model_state == "not_installed":
                    ollama_model_state_lbl.configure(text=self.t("settings.ollama.state.not_downloaded"), text_color=p.status_warning)
                    ollama_model_detail_lbl.configure(text=self.t("settings.ollama.state.not_downloaded_detail"), text_color=p.text_muted)
                    _set_action_button(ollama_primary_action_btn, text=self.t("settings.ollama.action.download"), command=_pull_ollama_model_from_settings, style="primary")
                    _set_action_button(ollama_remove_btn, text=self.t("settings.ollama.action.remove"), command=_remove_ollama_model_from_settings, visible=False)
                elif runtime == "stopped":
                    ollama_model_state_lbl.configure(text=self.t("settings.ollama.state.unknown"), text_color=p.status_warning)
                    ollama_model_detail_lbl.configure(text=self.t("settings.ollama.state.unknown_detail"), text_color=p.text_muted)
                    _set_action_button(ollama_primary_action_btn, text=self.t("settings.ollama.action.start"), command=_start_ollama_runtime, style="primary")
                    _set_action_button(ollama_remove_btn, text=self.t("settings.ollama.action.remove"), command=_remove_ollama_model_from_settings, visible=False)
                elif runtime == "not_installed":
                    ollama_model_state_lbl.configure(text=self.t("settings.ollama.state.unknown"), text_color=p.status_warning)
                    ollama_model_detail_lbl.configure(text=self.t("settings.ollama.runtime.not_installed"), text_color=p.text_muted)
                    _set_action_button(ollama_primary_action_btn, text=self.t("settings.ollama.action.get"), command=lambda: webbrowser.open("https://ollama.com/download"), style="primary")
                    _set_action_button(ollama_remove_btn, text=self.t("settings.ollama.action.remove"), command=_remove_ollama_model_from_settings, visible=False)
                elif runtime == "unavailable":
                    ollama_model_state_lbl.configure(text=self.t("settings.ollama.state.unknown"), text_color=p.status_error)
                    ollama_model_detail_lbl.configure(text=self.t("settings.ollama.state.unknown_detail"), text_color=p.text_muted)
                    _set_action_button(ollama_primary_action_btn, text=self.t("settings.ollama.action.retry"), command=_refresh_ollama_status, style="primary")
                    _set_action_button(ollama_remove_btn, text=self.t("settings.ollama.action.remove"), command=_remove_ollama_model_from_settings, visible=False)
                else:
                    ollama_model_state_lbl.configure(text=self.t("settings.ollama.state.checking"), text_color=p.text_secondary)
                    ollama_model_detail_lbl.configure(text="", text_color=p.text_muted)
                    _set_action_button(ollama_primary_action_btn, text=self.t("settings.ollama.action.retry"), command=_refresh_ollama_status, visible=False)
                    _set_action_button(ollama_remove_btn, text=self.t("settings.ollama.action.remove"), command=_remove_ollama_model_from_settings, visible=False)

        def _apply_ollama_event(event):
            def _update():
                if not win.winfo_exists():
                    return
                kind = getattr(event, "kind", "")
                if kind == "checking_runtime":
                    ollama_ui_state["runtime"] = "checking"
                elif kind in ("starting_runtime", "waiting_runtime"):
                    ollama_ui_state["runtime"] = "stopped"
                elif kind == "runtime_ready":
                    ollama_ui_state["runtime"] = "connected"
                elif kind == "downloading":
                    ollama_ui_state["runtime"] = "connected"
                    ollama_ui_state["model"] = "downloading"
                    ollama_ui_state["progress"] = getattr(event, "progress", None)
                    ollama_ui_state["stage"] = getattr(event, "stage", None)
                elif kind == "verifying":
                    ollama_ui_state["model"] = "downloading"
                    ollama_ui_state["progress"] = getattr(event, "progress", 100)
                    ollama_ui_state["stage"] = "verifying"
                elif kind == "ready":
                    ollama_ui_state["runtime"] = "connected"
                    ollama_ui_state["model"] = "installed"
                    ollama_ui_state["progress"] = 100
                elif kind == "removed":
                    ollama_ui_state["runtime"] = "connected"
                    ollama_ui_state["model"] = "not_installed"
                elif kind == "cancelled":
                    ollama_ui_state["model"] = "not_installed"
                _render_ollama_model_state()
            try:
                self.root.after(0, _update)
            except Exception:
                pass

        def _refresh_ollama_status(_event=None):
            if ollama_ui_state["busy"]:
                return
            ollama_ui_state["generation"] += 1
            generation = ollama_ui_state["generation"]
            ollama_ui_state.update({
                "runtime": "checking",
                "model": "checking",
                "error": None,
                "progress": None,
                "stage": None,
                "failed_action": None,
            })
            _render_ollama_model_state()
            model_name = _selected_ollama_model()
            base_url = _ollama_base_url()
            proxies = _ollama_proxies()

            def _worker():
                from ..translators.ollama import inspect_ollama_model
                inspection = inspect_ollama_model(model_name, base_url, proxies=proxies)

                def _apply():
                    if not win.winfo_exists() or generation != ollama_ui_state["generation"] or ollama_ui_state["busy"]:
                        return
                    ollama_ui_state["runtime"] = inspection.runtime
                    ollama_ui_state["model"] = inspection.model_state
                    ollama_ui_state["error"] = None
                    _render_ollama_model_state()
                try:
                    self.root.after(0, _apply)
                except Exception:
                    pass

            threading.Thread(target=_worker, daemon=True).start()

        def _on_ollama_model_selected(display_value: str):
            if ollama_ui_state["busy"]:
                return
            model_name = _OLLAMA_DISPLAY_TO_ID.get(display_value)
            if not model_name:
                return
            ollama_model_var.set(model_name)
            ollama_model_desc_lbl.configure(text=_OLLAMA_MODEL_META[model_name][1])
            _refresh_ollama_status()

        def _finish_ollama_operation(*, error: Exception | str | None = None, failed_action: str | None = None, refresh: bool = False):
            def _finish():
                # Always release non-UI resources, even if Settings was closed
                # while a download/removal thread was finishing.
                ollama_ui_state["busy"] = False
                ollama_ui_state["operation"] = None
                ollama_ui_state["failed_action"] = failed_action if error else None
                ollama_ui_state["error"] = str(error) if error else None
                self._ollama_install_cancel = None
                if ollama_ui_state.get("owns_pull_lock"):
                    try:
                        self._ollama_pull_lock.release()
                    except Exception:
                        pass
                    ollama_ui_state["owns_pull_lock"] = False
                try:
                    exists = bool(win.winfo_exists())
                except Exception:
                    exists = False
                if not exists:
                    return
                if refresh and error is None:
                    _refresh_ollama_status()
                else:
                    _render_ollama_model_state()
            try:
                self.root.after(0, _finish)
            except Exception:
                pass

        def _start_ollama_runtime():
            if ollama_ui_state["busy"]:
                return
            ollama_ui_state.update({"busy": True, "operation": "start", "error": None, "failed_action": None})
            _render_ollama_model_state()
            base_url = _ollama_base_url()
            proxies = _ollama_proxies()

            def _worker():
                try:
                    from ..translators.ollama import ensure_ollama_runtime
                    ensure_ollama_runtime(base_url, proxies=proxies, event_callback=_apply_ollama_event)
                except Exception as exc:
                    logger.exception("Could not start Ollama runtime")
                    _finish_ollama_operation(error=exc, failed_action="start")
                    return
                _finish_ollama_operation(refresh=True)

            threading.Thread(target=_worker, daemon=True).start()

        def _pull_ollama_model_from_settings():
            if ollama_ui_state["busy"]:
                return
            if not self._ollama_pull_lock.acquire(blocking=False):
                ollama_ui_state["error"] = self.t("settings.ollama_pull.already_running")
                ollama_ui_state["failed_action"] = "download"
                _render_ollama_model_state()
                return

            ollama_ui_state["owns_pull_lock"] = True
            model_name = _selected_ollama_model()
            base_url = _ollama_base_url()
            proxies = _ollama_proxies()
            cancel_event = threading.Event()
            self._ollama_install_cancel = cancel_event
            ollama_ui_state.update({
                "busy": True,
                "operation": "download",
                "runtime": "checking",
                "model": "downloading",
                "error": None,
                "failed_action": None,
                "progress": 0,
                "stage": "manifest",
            })
            _render_ollama_model_state()

            def _worker():
                try:
                    from ..translators.ollama import ollama_required_models, pull_ollama_models
                    pull_ollama_models(
                        ollama_required_models(model_name),
                        base_url=base_url,
                        proxies=proxies,
                        event_callback=_apply_ollama_event,
                        cancel_event=cancel_event,
                    )
                except Exception as exc:
                    from ..utils import CancelledError
                    if isinstance(exc, CancelledError) or cancel_event.is_set():
                        _finish_ollama_operation(refresh=True)
                        return
                    logger.exception("Ollama model download failed")
                    _finish_ollama_operation(error=exc, failed_action="download")
                    return
                _finish_ollama_operation(refresh=True)

            threading.Thread(target=_worker, daemon=True).start()

        def _cancel_ollama_download():
            cancel_event = self._ollama_install_cancel
            if cancel_event is None:
                return
            cancel_event.set()
            ollama_model_state_lbl.configure(text=self.t("settings.ollama.cancel_requested"), text_color=p.status_warning)
            self._set_button_disabled(ollama_primary_action_btn, True)

        def _remove_ollama_model_from_settings():
            if ollama_ui_state["busy"] or self._running:
                return
            display = _selected_ollama_display()
            if not messagebox.askyesno(
                self.t("settings.ollama.remove_confirm_title"),
                self.t("settings.ollama.remove_confirm_body", model=display),
                parent=win,
            ):
                return
            model_name = _selected_ollama_model()
            base_url = _ollama_base_url()
            proxies = _ollama_proxies()
            ollama_ui_state.update({"busy": True, "operation": "remove", "error": None, "failed_action": None})
            _render_ollama_model_state()

            def _worker():
                try:
                    from ..translators.ollama import remove_ollama_model
                    remove_ollama_model(model_name, base_url=base_url, proxies=proxies, event_callback=_apply_ollama_event)
                except Exception as exc:
                    logger.exception("Ollama model removal failed")
                    _finish_ollama_operation(error=exc)
                    return
                _finish_ollama_operation(refresh=True)

            threading.Thread(target=_worker, daemon=True).start()

        ollama_primary_action_btn = theme.make_button(
            ollama_model_actions,
            self.t("settings.ollama.action.retry"),
            command=_refresh_ollama_status,
            style="primary",
            height=27,
        )
        ollama_primary_action_btn.grid(row=0, column=0, sticky="w", padx=(0, 6))

        ollama_remove_btn = theme.make_button(
            ollama_model_actions,
            self.t("settings.ollama.action.remove"),
            command=_remove_ollama_model_from_settings,
            style="secondary",
            height=27,
        )
        ollama_remove_btn.grid(row=0, column=1, sticky="w", padx=(0, 6))
        ollama_remove_btn.grid_remove()

        ollama_refresh_btn = theme.make_button(
            ollama_model_actions,
            self.t("settings.ollama.action.refresh"),
            command=_refresh_ollama_status,
            style="ghost",
            height=27,
        )
        ollama_refresh_btn.grid(row=0, column=3, sticky="e")

        _rrow += 1

        try:
            ollama_base_url_entry.bind("<Return>", _refresh_ollama_status, add="+")
            ollama_base_url_entry.bind("<FocusOut>", _refresh_ollama_status, add="+")
        except Exception:
            pass

        _refresh_ollama_status()

        # Divider
        theme.make_divider(right).grid(row=_rrow, column=0, sticky="ew", pady=(4, 8))
        _rrow += 1

        # Google Cloud section
        _right_settings_label(self.t("settings.section.google_cloud"), level="section").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 6),
        )
        _rrow += 1

        # Google Cloud API key
        _right_settings_label(self.t("settings.google_cloud_api_key_v2"), level="small").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 4),
        )
        _rrow += 1
        google_key_entry = theme.make_entry(right, height=28)
        google_key_entry.grid(row=_rrow, column=0, sticky="ew", pady=(0, 1))
        google_key_entry.insert(0, self.cfg.get("google_api_key", ""))
        google_key_entry.configure(show="•")
        _rrow += 1

        lbl1 = _right_settings_label(self.t("settings.google_cloud_v2_hint"),
            level="tiny",
        )
        lbl1.configure(anchor="w", justify="left")
        lbl1.grid(row=_rrow, column=0, sticky="w", pady=(0, 8))
        _rrow += 1

        # Google Cloud v3 Project ID
        _right_settings_label(self.t("settings.google_cloud_project_id_v3"), level="small").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 2),
        )
        _rrow += 1
        google_project_entry = theme.make_entry(right, height=28)
        google_project_entry.grid(row=_rrow, column=0, sticky="ew", pady=(0, 2))
        google_project_entry.insert(0, self.cfg.get("google_project_id", ""))
        _rrow += 1

        # Google Cloud v3 Service Account JSON
        _right_settings_label(self.t("settings.google_cloud_sa_json_v3"), level="small").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 2),
        )
        _rrow += 1
        google_sa_entry = theme.make_entry(right, height=28)
        google_sa_entry.grid(row=_rrow, column=0, sticky="ew", pady=(0, 1))
        google_sa_entry.insert(0, self.cfg.get("google_sa_json", ""))
        _rrow += 1
        lbl_v3 = _right_settings_label(self.t("settings.google_cloud_v3_hint"),
            level="tiny",
        )
        lbl_v3.configure(anchor="w", justify="left")
        lbl_v3.grid(row=_rrow, column=0, sticky="w", pady=(0, 8))
        _rrow += 1

        # Divider
        theme.make_divider(right).grid(row=_rrow, column=0, sticky="ew", pady=(4, 8))
        _rrow += 1

        # Baidu section
        _right_settings_label(self.t("settings.section.baidu"), level="section").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 6),
        )
        _rrow += 1

        # Baidu App ID
        _right_settings_label(self.t("settings.baidu_app_id"), level="small").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 2),
        )
        _rrow += 1
        baidu_id_entry = theme.make_entry(right, height=28)
        baidu_id_entry.grid(row=_rrow, column=0, sticky="ew", pady=(0, 2))
        baidu_id_entry.insert(0, self.cfg.get("baidu_appid", ""))
        _rrow += 1

        # Baidu App Key
        _right_settings_label(self.t("settings.baidu_app_key"), level="small").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 2),
        )
        _rrow += 1
        baidu_key_entry = theme.make_entry(right, height=28)
        baidu_key_entry.grid(row=_rrow, column=0, sticky="ew", pady=(0, 1))
        baidu_key_entry.insert(0, self.cfg.get("baidu_appkey", ""))
        baidu_key_entry.configure(show="•")
        _rrow += 1

        lbl = _right_settings_label(self.t("settings.baidu_hint"),
            level="tiny",
        )
        lbl.configure(anchor="w", justify="left")
        lbl.grid(row=_rrow, column=0, sticky="w", pady=(0, 4))
        _rrow += 1

        # Baidu API tier selection
        _right_settings_label(self.t("settings.api_tier"), level="small").grid(
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

        _tier_btn_standard = _make_tier_btn(_tier_frame, self.t("settings.api_tier.standard"), "standard")
        _tier_btn_premium  = _make_tier_btn(_tier_frame, self.t("settings.api_tier.premium"),  "premium")

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
        _right_settings_label(self.t("settings.section.azure"), level="section").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 6),
        )
        _rrow += 1

        _right_settings_label(self.t("settings.azure_subscription_key"), level="small").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 2),
        )
        _rrow += 1
        azure_key_entry = theme.make_entry(right, height=28)
        azure_key_entry.grid(row=_rrow, column=0, sticky="ew", pady=(0, 2))
        azure_key_entry.insert(0, self.cfg.get("azure_key", ""))
        azure_key_entry.configure(show="•")
        _rrow += 1

        _right_settings_label(self.t("settings.azure_region"), level="small").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 2),
        )
        _rrow += 1
        azure_region_entry = theme.make_entry(right, height=28)
        azure_region_entry.grid(row=_rrow, column=0, sticky="ew", pady=(0, 1))
        azure_region_entry.insert(0, self.cfg.get("azure_region", ""))
        _rrow += 1
        lbl_az = _right_settings_label(self.t("settings.azure_hint"),
            level="tiny",
        )
        lbl_az.configure(anchor="w", justify="left")
        lbl_az.grid(row=_rrow, column=0, sticky="w", pady=(0, 8))
        _rrow += 1

        # Divider
        theme.make_divider(right).grid(row=_rrow, column=0, sticky="ew", pady=(4, 8))
        _rrow += 1

        # DeepL section
        _right_settings_label(self.t("settings.section.deepl"), level="section").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 6),
        )
        _rrow += 1

        _right_settings_label(self.t("settings.deepl_api_key"), level="small").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 2),
        )
        _rrow += 1
        deepl_key_entry = theme.make_entry(right, height=28)
        deepl_key_entry.grid(row=_rrow, column=0, sticky="ew", pady=(0, 1))
        deepl_key_entry.insert(0, self.cfg.get("deepl_api_key", ""))
        deepl_key_entry.configure(show="•")
        _rrow += 1
        lbl_dl = _right_settings_label(self.t("settings.deepl_hint"),
            level="tiny",
        )
        lbl_dl.configure(anchor="w", justify="left")
        lbl_dl.grid(row=_rrow, column=0, sticky="w", pady=(0, 8))
        _rrow += 1

        # Divider
        theme.make_divider(right).grid(row=_rrow, column=0, sticky="ew", pady=(4, 8))
        _rrow += 1

        # Usage summary + cache controls
        _right_settings_label(self.t("settings.section.usage"), level="section").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 4),
        )
        _rrow += 1

        def _get_usage_text() -> str:
            try:
                from ..translators.usage import get_tracker
                t = get_tracker()
                lines: list[str] = []
                for eng_key, _label_key in _ENGINE_LABEL_KEYS:
                    usage_str = t.format_usage(eng_key)
                    if usage_str:
                        display_name = self._engine_display_name(eng_key)
                        lines.append(f"{display_name}: {usage_str}")
                return "\n".join(lines) if lines else self.t("settings.no_usage_data")
            except Exception:
                return ""

        usage_text = _get_usage_text()
        usage_lbl = _right_settings_label(usage_text or self.t("settings.no_usage_data"), level="tiny")
        usage_lbl.configure(anchor="w", justify="left")
        usage_lbl.grid(row=_rrow, column=0, sticky="w", pady=(0, 6))
        _rrow += 1

        _right_settings_label(self.t("settings.backend_cache.title"), level="small",
        ).grid(row=_rrow, column=0, sticky="w", pady=(0, 2))
        _rrow += 1

        cache_hint = _right_settings_label(self.t("settings.backend_cache.hint"), level="tiny",
        )
        cache_hint.configure(anchor="w", justify="left")
        cache_hint.grid(row=_rrow, column=0, sticky="ew", pady=(0, 5))
        _rrow += 1

        # Create a small row with the button and a ghost label beside it
        try:
            # compute initial text (always visible, even when 0)
            from ..translators.cache import get_cache
            from ..utils.io import format_bytes
            n = get_cache().size()
            b = get_cache().disk_usage_bytes()
            initial_lbl = self.t("settings.cache.entries", entries=f"{n:,}", size=format_bytes(b))
        except Exception:
            initial_lbl = self.t("settings.cache.entries", entries="0", size="0 B")

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
                txt = self.t("settings.cache.entries", entries=f"{n:,}", size=format_bytes(b))
            except Exception:
                txt = self.t("settings.cache.entries", entries="0", size="0 B")
            try:
                cache_lbl.configure(text=txt)
                cache_lbl.update_idletasks()
            except Exception:
                pass

        def _clear_cache():
            try:
                from ..translators.cache import get_cache
                get_cache().clear()
                self._log(self.t("log.cache_cleared"))
            except Exception as e:
                self._log(self.t("log.cache_clear_error", error=e))
            try:
                _update_cache_label()
            except Exception:
                pass

        clear_btn = theme.make_button(
            _cache_row, self.t("settings.backend_cache.clear"),
            command=_clear_cache, style="secondary", height=26,
        )
        clear_btn.pack(side=tk.LEFT)
        cache_lbl.pack(side=tk.LEFT, padx=(8, 0))
        _rrow += 1

        # Translation Memory controls. These settings are snapshotted when a
        # worker starts, so edits here never mutate an in-flight job.
        theme.make_divider(right).grid(row=_rrow, column=0, sticky="ew", pady=(8, 8))
        _rrow += 1
        _right_settings_label(self.t("settings.section.translation_memory"), level="section").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 4),
        )
        _rrow += 1

        tm_enabled_var = tk.BooleanVar(value=bool(self.cfg.get("translation_memory_enabled", False)))
        tm_enabled_cb = ctk.CTkCheckBox(
            right, text=self.t("settings.translation_memory.enable"),
            variable=tm_enabled_var, onvalue=True, offvalue=False,
            checkmark_color=p.bg_main, fg_color=p.accent, hover_color=p.accent_hover,
            border_color=p.border, text_color=p.text_secondary,
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=theme.FONT_BODY[1]),
        )
        tm_enabled_cb.grid(row=_rrow, column=0, sticky="w", pady=(0, 3))
        _rrow += 1
        tm_explain = _right_settings_label(self.t("settings.translation_memory.hint"), level="tiny",
        )
        tm_explain.configure(anchor="w", justify="left")
        tm_explain.grid(row=_rrow, column=0, sticky="ew", pady=(0, 6))
        _rrow += 1

        _right_settings_label(self.t("settings.translation_memory_path"), level="small").grid(
            row=_rrow, column=0, sticky="w", pady=(0, 2),
        )
        _rrow += 1
        _tm_path_row = ctk.CTkFrame(right, fg_color="transparent")
        _tm_path_row.grid(row=_rrow, column=0, sticky="ew", pady=(0, 2))
        _tm_path_row.grid_columnconfigure(0, weight=1)
        tm_path_entry = theme.make_entry(_tm_path_row, height=28)
        tm_path_entry.grid(row=0, column=0, sticky="ew")
        _configured_tm_path = str(self.cfg.get("translation_memory_path") or "").strip()
        tm_path_entry.insert(0, _configured_tm_path or str(default_translation_memory_path()))

        def _browse_tm_path():
            raw = tm_path_entry.get().strip() or str(default_translation_memory_path())
            candidate = Path(raw).expanduser()
            initialdir = candidate.parent if candidate.parent.exists() else Path.home()
            chosen = filedialog.asksaveasfilename(
                title=self.t("dialog.select_translation_memory"), parent=win,
                initialdir=str(initialdir), initialfile=candidate.name or "translation_memory.sqlite3",
                defaultextension=".sqlite3", filetypes=[("SQLite", "*.sqlite3"), ("All files", "*.*")],
            )
            if chosen:
                tm_path_entry.delete(0, tk.END)
                tm_path_entry.insert(0, chosen)
                _update_tm_label()

        theme.make_button(
            _tm_path_row, self.t("settings.browse"), command=_browse_tm_path,
            style="secondary", image=browse_icon_s, height=28,
        ).grid(row=0, column=1, padx=(6, 0))
        _rrow += 1

        _right_settings_label(self.t("settings.translation_memory_path_hint"), level="tiny",
        ).grid(row=_rrow, column=0, sticky="w", pady=(0, 4))
        _rrow += 1

        _tm_actions = ctk.CTkFrame(right, fg_color="transparent")
        _tm_actions.grid(row=_rrow, column=0, sticky="w", pady=(0, 4))
        tm_info_lbl = theme.make_label(_tm_actions, "", level="tiny")

        def _tm_selected_path() -> Path:
            raw = tm_path_entry.get().strip()
            return Path(raw).expanduser() if raw else default_translation_memory_path()

        def _update_tm_label():
            try:
                path = _tm_selected_path()
                if path.exists():
                    from ..utils.io import format_bytes
                    stats = TranslationMemory(path, timeout=0.75).stats()
                    text = self.t(
                        "settings.translation_memory.entries",
                        entries=f"{stats['entries']:,}", size=format_bytes(stats["bytes"]),
                    )
                else:
                    text = self.t("settings.translation_memory.entries", entries="0", size="0 B")
            except Exception as exc:
                text = self.t("settings.error.translation_memory", error=exc)
            tm_info_lbl.configure(text=text)

        def _manage_tm_from_settings():
            path = _tm_selected_path()
            try:
                win.grab_release()
            except Exception:
                pass
            self._open_translation_memory_manager(str(path), parent=win)

        def _clear_tm_from_settings():
            if self._running:
                messagebox.showwarning(
                    self.t("settings.section.translation_memory"),
                    self.t("settings.translation_memory.running_note"), parent=win,
                )
                return
            if not messagebox.askyesno(
                self.t("settings.translation_memory.confirm_clear_title"),
                self.t("settings.translation_memory.confirm_clear_body"), parent=win,
            ):
                return
            try:
                TranslationMemory(_tm_selected_path(), timeout=0.75).clear(vacuum=True)
                _update_tm_label()
            except Exception as exc:
                self._settings_error.configure(
                    text=self.t("settings.error.translation_memory", error=exc)
                )

        tm_manage_btn = theme.make_button(
            _tm_actions, self.t("settings.translation_memory.manage"),
            command=_manage_tm_from_settings, style="secondary", height=26,
        )
        tm_manage_btn.pack(side=tk.LEFT)
        tm_clear_btn = theme.make_button(
            _tm_actions, self.t("settings.translation_memory.clear"),
            command=_clear_tm_from_settings, style="secondary", height=26,
        )
        tm_clear_btn.pack(side=tk.LEFT, padx=(6, 0))
        if self._running:
            self._set_button_disabled(tm_clear_btn, True)
        tm_info_lbl.pack(side=tk.LEFT, padx=(8, 0))
        _update_tm_label()
        _rrow += 1

        try:
            right.after_idle(_sync_right_settings_wrap)
        except Exception:
            pass

        # ── Bottom row: error label + buttons ────────────────────────────
        bottom = ctk.CTkFrame(card, fg_color="transparent")
        bottom.grid(row=2, column=0, sticky="ew", padx=PAD, pady=(0, PAD//2))
        bottom.grid_columnconfigure(0, weight=1)

        # Inline validation error (spans full width)
        self._settings_error = theme.make_label(
            bottom, "", level="tiny",
            text_color=p.status_error,
        )
        self._settings_error.configure(anchor="w", justify="left", wraplength=760)
        self._settings_error.grid(row=0, column=0, sticky="ew", pady=(0, 6))

        # Button row
        btn_frame = ctk.CTkFrame(bottom, fg_color="transparent")
        btn_frame.grid(row=1, column=0, sticky="w", pady=(0, 0))

        def _save_and_close():
            inp = in_entry.get().strip()
            out = out_entry.get().strip()
            if not out:
                self._settings_error.configure(text=self.t("settings.error.input_or_output_empty"))
                return
            self._settings_error.configure(text="")
            tm_path_value = tm_path_entry.get().strip()
            if tm_enabled_var.get():
                try:
                    TranslationMemory(tm_path_value or default_translation_memory_path(), timeout=0.75).count()
                except Exception as exc:
                    self._settings_error.configure(
                        text=self.t("settings.error.translation_memory", error=exc)
                    )
                    return
            previous_cfg = dict(self.cfg)
            self.cfg["default_input"] = inp
            self.cfg["default_output"] = out
            new_mode = "Dark" if mode_switch_var.get() else "Light"
            self.cfg["appearance_mode"] = new_mode
            selected_locale_name = ui_lang_var.get()
            self.cfg["ui_locale"] = next(
                (code for code, name in _locale_options if name == selected_locale_name),
                _current_locale,
            )
            self.cfg["auto_check_updates"] = auto_updates_var.get()
            self.cfg["debug_mode"] = debug_var.get()
            self.cfg["ollama_enabled"] = ollama_enabled_var.get()
            self.cfg["ollama_base_url"] = ollama_base_url_entry.get().strip()
            self.cfg["ollama_model"] = ollama_model_var.get()
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
            self.cfg["translation_memory_enabled"] = bool(tm_enabled_var.get())
            default_tm = str(default_translation_memory_path())
            self.cfg["translation_memory_path"] = "" if not tm_path_value or tm_path_value == default_tm else tm_path_value
            if not save_config(self.cfg):
                self.cfg = previous_cfg
                self._settings_error.configure(text=self.t("settings.error.save_failed"))
                return

            # Refresh engine-dependent UI after config changes
            try:
                self._refresh_language_dropdowns()
                self._update_usage_label(self._engine_key_for_display(self.engine_var.get()))
            except Exception:
                logger.exception("Failed to refresh language dropdowns after saving settings")

            try:
                self._apply_debug_mode()
            except Exception:
                pass
            if self._ollama_install_cancel is not None:
                self._ollama_install_cancel.set()
            win.destroy()

        def _on_settings_close():
            if self._ollama_install_cancel is not None:
                self._ollama_install_cancel.set()
            win.destroy()

        theme.make_button(btn_frame, self.t("settings.save"), command=_save_and_close, style="primary",
                          height=28).pack(
            side=tk.LEFT, padx=(0, 6),
        )
        theme.make_button(btn_frame, self.t("sidebar.cancel"), command=_on_settings_close, style="secondary",
                          height=28).pack(
            side=tk.LEFT,
        )

        # Title-bar X / Escape act as Cancel; Ctrl+Enter saves.
        win.protocol("WM_DELETE_WINDOW", _on_settings_close)
        self._bind_modal_keys(win, _on_settings_close, _save_and_close, in_entry)
        win.update_idletasks()
        def _show_settings():
            center_window(win, parent=self.root)
            win.wm_attributes("-alpha", 1)
        win.after(20, _show_settings)
        try:
            win.resizable(True, True)
        except Exception:
            pass


    # --- terminology / translation-memory management ---

    def _open_terminology_manager(self):
        p = theme.get()
        win = ctk.CTkToplevel(self.root)
        win.wm_attributes("-alpha", 0)
        apply_window_icon(win)
        win.title(self.t("terminology.title"))
        win.transient(self.root)
        win.grab_set()
        win.configure(fg_color=p.bg_main)
        win.geometry("900x560")
        win.minsize(760, 460)
        win.grid_columnconfigure(0, weight=1)
        win.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(win, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=theme.PADDING, pady=(theme.PADDING, 8))
        header.grid_columnconfigure(0, weight=1)
        theme.make_label(header, self.t("terminology.title"), level="heading").grid(row=0, column=0, sticky="w")
        theme.make_label(header, self.t("terminology.intro"), level="tiny").grid(row=1, column=0, sticky="w", pady=(2, 0))
        if self._running:
            theme.make_label(
                header, self.t("terminology.running_note"), level="tiny", text_color=p.status_warning,
            ).grid(row=2, column=0, sticky="w", pady=(2, 0))

        card = theme.make_card(win)
        card.grid(row=1, column=0, sticky="nsew", padx=theme.PADDING, pady=(0, 8))
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(0, weight=1)

        columns = ("enabled", "source_lang", "target_lang", "source", "target")
        tree = ttk.Treeview(
            card, columns=columns, show="headings", selectmode="browse",
            takefocus=True, style="FileTable.Treeview",
        )
        _style_app_table(tree, p)
        tree.heading("enabled", text=self.t("terminology.column.enabled"))
        tree.heading("source_lang", text=self.t("terminology.column.source_language"))
        tree.heading("target_lang", text=self.t("terminology.column.target_language"))
        tree.heading("source", text=self.t("terminology.column.source"))
        tree.heading("target", text=self.t("terminology.column.target"))
        tree.column("enabled", width=55, stretch=False, anchor="center")
        tree.column("source_lang", width=85, stretch=False, anchor="center")
        tree.column("target_lang", width=85, stretch=False, anchor="center")
        tree.column("source", width=260, stretch=True)
        tree.column("target", width=260, stretch=True)
        scrollbar = ttk.Scrollbar(
            card, orient=tk.VERTICAL, command=tree.yview,
            style="Slim.Vertical.TScrollbar",
        )
        tree.configure(yscrollcommand=scrollbar.set)
        tree.grid(row=0, column=0, sticky="nsew", padx=(12, 0), pady=12)
        scrollbar.grid(row=0, column=1, sticky="ns", padx=(0, 12), pady=12)

        status = theme.make_label(win, "", level="tiny")
        status.grid(row=2, column=0, sticky="w", padx=theme.PADDING, pady=(0, 4))

        entries_by_id: dict[str, TerminologyEntry] = {}

        def _refresh():
            nonlocal entries_by_id
            entries = self._terminology_store.load()
            entries_by_id = {entry.id: entry for entry in entries}
            for iid in tree.get_children():
                tree.delete(iid)
            for entry in sorted(entries, key=lambda e: (e.source_lang, e.target_lang, e.source_text.casefold())):
                tree.insert(
                    "", tk.END, iid=entry.id,
                    values=("✓" if entry.enabled else "", entry.source_lang, entry.target_lang, entry.source_text, entry.target_text),
                )
            _restripe_tree(tree)
            if self._terminology_store.last_error:
                status.configure(text=self.t("terminology.import_error", error=self._terminology_store.last_error), text_color=p.status_error)
            else:
                status.configure(
                    text=self.t("terminology.entry_count", count=f"{len(entries):,}"),
                    text_color=p.text_muted,
                )

        language_options = _get_language_options(self.ui.locale)
        lang_display = [self._format_language_option(code, name) for code, name in language_options]
        lang_map = {self._format_language_option(code, name): code for code, name in language_options}
        reverse_lang = {code.lower(): display for display, code in lang_map.items()}

        def _edit_entry(existing: TerminologyEntry | None = None):
            dlg = ctk.CTkToplevel(win)
            dlg.wm_attributes("-alpha", 0)
            apply_window_icon(dlg)
            dlg.title(self.t("terminology.edit") if existing else self.t("terminology.add"))
            dlg.transient(win)
            dlg.grab_set()
            dlg.configure(fg_color=p.bg_main)
            frame = theme.make_card(dlg)
            frame.pack(fill=tk.BOTH, expand=True, padx=theme.PADDING, pady=theme.PADDING)
            frame.grid_columnconfigure(0, weight=1)

            theme.make_label(frame, self.t("terminology.source_language"), level="small").grid(row=0, column=0, sticky="w", padx=12, pady=(12, 2))
            source_initial = reverse_lang.get(existing.source_lang.lower(), "") if existing else ""
            source_var = tk.StringVar(value=source_initial)
            source_box = SearchableComboBox(frame, values=lang_display, variable=source_var)
            source_box.grid(row=1, column=0, sticky="ew", padx=12)

            theme.make_label(frame, self.t("terminology.target_language"), level="small").grid(row=2, column=0, sticky="w", padx=12, pady=(10, 2))
            target_initial = reverse_lang.get(existing.target_lang.lower(), "") if existing else ""
            target_var = tk.StringVar(value=target_initial)
            target_box = SearchableComboBox(frame, values=lang_display, variable=target_var)
            target_box.grid(row=3, column=0, sticky="ew", padx=12)

            theme.make_label(frame, self.t("terminology.source_term"), level="small").grid(row=4, column=0, sticky="w", padx=12, pady=(10, 2))
            source_entry = theme.make_entry(frame)
            source_entry.grid(row=5, column=0, sticky="ew", padx=12)
            if existing:
                source_entry.insert(0, existing.source_text)

            theme.make_label(frame, self.t("terminology.target_term"), level="small").grid(row=6, column=0, sticky="w", padx=12, pady=(10, 2))
            target_entry = theme.make_entry(frame)
            target_entry.grid(row=7, column=0, sticky="ew", padx=12)
            if existing:
                target_entry.insert(0, existing.target_text)

            enabled_var = tk.BooleanVar(value=True if existing is None else existing.enabled)
            ctk.CTkCheckBox(
                frame, text=self.t("terminology.enabled"), variable=enabled_var,
                fg_color=p.accent, hover_color=p.accent_hover, border_color=p.border,
                text_color=p.text_secondary,
            ).grid(row=8, column=0, sticky="w", padx=12, pady=(10, 2))
            error_lbl = theme.make_label(frame, "", level="tiny", text_color=p.status_error)
            error_lbl.grid(row=9, column=0, sticky="w", padx=12, pady=(2, 0))

            def _resolve_code(box, var):
                display = var.get().strip() or box.get().strip()
                return lang_map.get(display, display.split("(")[-1].rstrip(") ").strip() if "(" in display else "")

            def _save():
                src = _resolve_code(source_box, source_var)
                tgt = _resolve_code(target_box, target_var)
                source_text = source_entry.get().strip()
                target_text = target_entry.get().strip()
                if not src or not tgt or not source_text or not target_text:
                    error_lbl.configure(text=self.t("terminology.error.required"))
                    return
                entries = self._terminology_store.load()
                duplicate = next((
                    e for e in entries
                    if e.id != (existing.id if existing else "")
                    and e.source_lang.lower() == src.lower()
                    and e.target_lang.lower() == tgt.lower()
                    and e.source_text.casefold() == source_text.casefold()
                ), None)
                if duplicate is not None:
                    error_lbl.configure(text=self.t("terminology.error.duplicate"))
                    return
                try:
                    updated = TerminologyEntry.create(
                        src, tgt, source_text, target_text, enabled=enabled_var.get(),
                        entry_id=existing.id if existing else None,
                    )
                    new_entries = [e for e in entries if e.id != updated.id] + [updated]
                    self._terminology_store.save(new_entries)
                except Exception as exc:
                    error_lbl.configure(text=str(exc))
                    return
                _close_edit()
                _refresh()

            def _close_edit():
                try:
                    dlg.destroy()
                finally:
                    try:
                        if win.winfo_exists():
                            win.grab_set()
                    except Exception:
                        pass

            buttons = ctk.CTkFrame(frame, fg_color="transparent")
            buttons.grid(row=10, column=0, sticky="w", padx=12, pady=12)
            theme.make_button(buttons, self.t("settings.save"), command=_save, style="primary", height=28).pack(side=tk.LEFT, padx=(0, 6))
            theme.make_button(buttons, self.t("sidebar.cancel"), command=_close_edit, style="secondary", height=28).pack(side=tk.LEFT)
            dlg.protocol("WM_DELETE_WINDOW", _close_edit)
            self._bind_modal_keys(dlg, _close_edit, _save, source_box)
            dlg.update_idletasks()
            center_window(dlg, parent=win)
            dlg.wm_attributes("-alpha", 1)

        def _selected_entry() -> TerminologyEntry | None:
            selection = tree.selection()
            return entries_by_id.get(selection[0]) if selection else None

        def _delete_selected():
            entry = _selected_entry()
            if entry is None:
                return
            if not messagebox.askyesno(
                self.t("terminology.confirm_delete_title"), self.t("terminology.confirm_delete_body"), parent=win
            ):
                return
            self._terminology_store.save(e for e in entries_by_id.values() if e.id != entry.id)
            _refresh()

        def _import():
            path = filedialog.askopenfilename(
                parent=win, filetypes=[("Terminology", "*.csv *.json"), ("CSV", "*.csv"), ("JSON", "*.json"), ("All files", "*.*")]
            )
            if not path:
                return
            try:
                added, updated = self._terminology_store.import_file(path)
                _refresh()
                messagebox.showinfo(
                    self.t("terminology.title"), self.t("terminology.import_success", added=added, updated=updated), parent=win
                )
            except Exception as exc:
                messagebox.showerror(
                    self.t("terminology.title"), self.t("terminology.import_error", error=exc), parent=win
                )

        def _export():
            path = filedialog.asksaveasfilename(
                parent=win, defaultextension=".csv", initialfile="verbilo_terminology.csv",
                filetypes=[("CSV", "*.csv"), ("JSON", "*.json")],
            )
            if not path:
                return
            try:
                count = self._terminology_store.export_file(path)
                messagebox.showinfo(
                    self.t("terminology.title"), self.t("terminology.export_success", count=count), parent=win
                )
            except Exception as exc:
                messagebox.showerror(
                    self.t("terminology.title"), self.t("terminology.export_error", error=exc), parent=win
                )

        def _close_terms():
            win.destroy()

        actions = ctk.CTkFrame(win, fg_color="transparent")
        actions.grid(row=3, column=0, sticky="ew", padx=theme.PADDING, pady=(0, theme.PADDING))
        theme.make_button(actions, self.t("terminology.add"), command=lambda: _edit_entry(None), style="primary", height=28).pack(side=tk.LEFT, padx=(0, 6))
        theme.make_button(actions, self.t("terminology.edit"), command=lambda: _edit_entry(_selected_entry()) if _selected_entry() else None, style="secondary", height=28).pack(side=tk.LEFT, padx=(0, 6))
        theme.make_button(actions, self.t("terminology.delete"), command=_delete_selected, style="secondary", height=28).pack(side=tk.LEFT, padx=(0, 12))
        theme.make_button(actions, self.t("terminology.import"), command=_import, style="secondary", height=28).pack(side=tk.LEFT, padx=(0, 6))
        theme.make_button(actions, self.t("terminology.export"), command=_export, style="secondary", height=28).pack(side=tk.LEFT)
        theme.make_button(actions, self.t("terminology.close"), command=_close_terms, style="ghost", height=28).pack(side=tk.RIGHT)
        tree.bind("<Double-1>", lambda _event: _edit_entry(_selected_entry()) if _selected_entry() else None)
        tree.bind("<Return>", lambda _event: (_edit_entry(_selected_entry()) if _selected_entry() else None, "break")[-1])
        tree.bind("<Delete>", lambda _event: (_delete_selected(), "break")[-1])
        win.bind("<Control-n>", lambda _event: (_edit_entry(None), "break")[-1], "+")
        win.bind("<Command-n>", lambda _event: (_edit_entry(None), "break")[-1], "+")
        win.bind("<Escape>", lambda _event: (_close_terms(), "break")[-1], "+")
        self._bind_modal_keys(win, _close_terms, initial_focus=tree)
        win.protocol("WM_DELETE_WINDOW", _close_terms)
        _refresh()
        center_window(win, 900, 560, parent=self.root)
        win.wm_attributes("-alpha", 1)

    def _open_translation_memory_manager(self, path: str | None = None, *, parent=None):
        p = theme.get()
        parent_window = parent or self.root
        try:
            memory = TranslationMemory(path or default_translation_memory_path(), timeout=0.75)
        except Exception as exc:
            messagebox.showerror(
                self.t("settings.section.translation_memory"),
                self.t("settings.error.translation_memory", error=exc), parent=parent_window,
            )
            try:
                if parent is not None:
                    parent.grab_set()
            except Exception:
                pass
            return

        win = ctk.CTkToplevel(parent_window)
        win.wm_attributes("-alpha", 0)
        apply_window_icon(win)
        win.title(self.t("tm_manager.title"))
        win.transient(parent_window)
        win.grab_set()
        win.configure(fg_color=p.bg_main)
        win.geometry("920x560")
        win.minsize(760, 440)
        win.grid_columnconfigure(0, weight=1)
        win.grid_rowconfigure(1, weight=1)

        top = ctk.CTkFrame(win, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=theme.PADDING, pady=(theme.PADDING, 8))
        top.grid_columnconfigure(0, weight=1)
        search_entry = theme.make_entry(top, placeholder_text=self.t("tm_manager.search"))
        search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        card = theme.make_card(win)
        card.grid(row=1, column=0, sticky="nsew", padx=theme.PADDING, pady=(0, 8))
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(0, weight=1)
        columns = ("source", "translation", "languages", "uses")
        tree = ttk.Treeview(
            card, columns=columns, show="headings", selectmode="extended",
            takefocus=True, style="FileTable.Treeview",
        )
        _style_app_table(tree, p)
        for key, label, width, stretch in (
            ("source", self.t("tm_manager.column.source"), 300, True),
            ("translation", self.t("tm_manager.column.translation"), 300, True),
            ("languages", self.t("tm_manager.column.languages"), 100, False),
            ("uses", self.t("tm_manager.column.uses"), 70, False),
        ):
            tree.heading(key, text=label)
            tree.column(key, width=width, stretch=stretch, anchor="center" if not stretch else "w")
        scroll = ttk.Scrollbar(
            card, orient=tk.VERTICAL, command=tree.yview,
            style="Slim.Vertical.TScrollbar",
        )
        tree.configure(yscrollcommand=scroll.set)
        tree.grid(row=0, column=0, sticky="nsew", padx=(12, 0), pady=12)
        scroll.grid(row=0, column=1, sticky="ns", padx=(0, 12), pady=12)
        info = theme.make_label(win, "", level="tiny")
        info.grid(row=2, column=0, sticky="w", padx=theme.PADDING, pady=(0, 4))

        def _refresh():
            try:
                rows = memory.search(search_entry.get().strip(), limit=500)
                for iid in tree.get_children():
                    tree.delete(iid)
                for entry in rows:
                    tree.insert(
                        "", tk.END, iid=entry.key,
                        values=(entry.source_text, entry.translation, f"{entry.source_lang} → {entry.target_lang}", entry.use_count),
                    )
                _restripe_tree(tree)
                from ..utils.io import format_bytes
                stats = memory.stats()
                info.configure(
                    text=self.t(
                        "tm_manager.showing", shown=f"{len(rows):,}",
                        total=f"{stats['entries']:,}", size=format_bytes(stats["bytes"]),
                    ),
                    text_color=p.text_muted,
                )
            except Exception as exc:
                info.configure(text=self.t("settings.error.translation_memory", error=exc), text_color=p.status_error)

        def _delete():
            if self._running:
                messagebox.showwarning(
                    self.t("tm_manager.title"), self.t("settings.translation_memory.running_note"), parent=win
                )
                return
            selection = tree.selection()
            if not selection:
                return
            if not messagebox.askyesno(
                self.t("tm_manager.confirm_delete_title"), self.t("tm_manager.confirm_delete_body"), parent=win
            ):
                return
            try:
                memory.delete_many(list(selection))
                _refresh()
            except Exception as exc:
                messagebox.showerror(
                    self.t("tm_manager.title"), self.t("settings.error.translation_memory", error=exc), parent=win
                )

        theme.make_button(top, self.t("tm_manager.refresh"), command=_refresh, style="secondary", height=28).grid(row=0, column=1)
        search_entry.bind("<Return>", lambda _event: (_refresh(), "break")[-1])
        bottom = ctk.CTkFrame(win, fg_color="transparent")
        bottom.grid(row=3, column=0, sticky="ew", padx=theme.PADDING, pady=(0, theme.PADDING))
        delete_btn = theme.make_button(bottom, self.t("tm_manager.delete"), command=_delete, style="secondary", height=28)
        delete_btn.pack(side=tk.LEFT)
        if self._running:
            self._set_button_disabled(delete_btn, True)
        theme.make_button(bottom, self.t("tm_manager.close"), command=lambda: _close(), style="ghost", height=28).pack(side=tk.RIGHT)

        def _close():
            try:
                win.destroy()
            finally:
                try:
                    if parent is not None and parent.winfo_exists():
                        parent.grab_set()
                except Exception:
                    pass
        win.protocol("WM_DELETE_WINDOW", _close)
        self._bind_modal_keys(win, _close, initial_focus=search_entry)
        win.bind("<Control-f>", lambda _event: (search_entry.focus_set(), "break")[-1], "+")
        win.bind("<Command-f>", lambda _event: (search_entry.focus_set(), "break")[-1], "+")
        win.bind("<F5>", lambda _event: (_refresh(), "break")[-1], "+")
        tree.bind("<Delete>", lambda _event: (_delete(), "break")[-1], "+")
        tree.bind("<Control-a>", lambda _event: (tree.selection_set(tree.get_children()), "break")[-1], "+")
        tree.bind("<Command-a>", lambda _event: (tree.selection_set(tree.get_children()), "break")[-1], "+")
        _refresh()
        center_window(win, 920, 560, parent=parent_window)
        win.wm_attributes("-alpha", 1)

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
        dlg.title(self.t("dialog.update_available_title"))
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.configure(fg_color=p.bg_main)
        dlg.grid_columnconfigure(0, weight=1)

        card = theme.make_card(dlg)
        card.grid(row=0, column=0, sticky="nsew", padx=PAD, pady=PAD)
        card.grid_columnconfigure(0, weight=1)

        theme.make_label(
            card, self.t("dialog.update_available_heading", version=result['version']), level="subheading",
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=PAD, pady=(PAD, 4))
        theme.make_label(
            card, self.t("dialog.update_available_body"), level="small",
        ).grid(row=1, column=0, columnspan=2, sticky="w", padx=PAD, pady=(0, PAD))

        dlg_btn_frame = ctk.CTkFrame(card, fg_color="transparent")
        dlg_btn_frame.grid(row=2, column=0, columnspan=2, pady=(0, PAD))

        def _download():
            webbrowser.open(result.get("url") or RELEASES_URL)
            dlg.destroy()

        theme.make_button(dlg_btn_frame, self.t("dialog.download"), command=_download, style="primary",
                          height=32).pack(side=tk.LEFT, padx=(0, 8))
        theme.make_button(dlg_btn_frame, self.t("dialog.later"), command=dlg.destroy, style="secondary",
                          height=32).pack(side=tk.LEFT)

        dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)
        self._bind_modal_keys(dlg, dlg.destroy)
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
        win.title(self.t("dialog.local_model_manager_title"))
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
            # Use the backend's readiness scan so "downloaded" means the model
            # has the sentinel *and* all files required by the translator.
            parts = canonical_name.split("-", 1)
            if len(parts) != 2:
                return False
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
        theme.make_label(hdr, self.t("dialog.local_model_manager_heading"), level="heading").pack(side=tk.LEFT)

        # --- info note ---
        theme.make_label(
            card,
            self.t("dialog.local_model_manager_note"),
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

        _hdr_labels = ["", self.t("dialog.local_model_manager.column_source"), self.t("dialog.local_model_manager.column_target"), self.t("dialog.local_model_manager.column_size"), self.t("dialog.local_model_manager.column_progress"), ""]
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
                rw["progress_label"].configure(text=self.t("dialog.local_model_manager.progress.initial"), text_color=p.status_info)
            elif state == "error":
                rw["progress_label"].configure(text=self.t("dialog.local_model_manager.progress.error"), text_color=p.status_error)
            elif downloaded:
                rw["progress_label"].configure(text=self.t("dialog.local_model_manager.progress.complete"), text_color=p.status_success)
            else:
                rw["progress_label"].configure(text=self.t("dialog.local_model_manager.progress.idle"), text_color=p.text_muted)

            # Action frame — clear and rebuild
            for child in rw["action_frame"].winfo_children():
                child.destroy()

            if state == "downloading":
                theme.make_button(
                    rw["action_frame"], self.t("dialog.local_model_manager.cancel"),
                    command=lambda cn=cname: _cancel_download(cn),
                    style="ghost", height=22,
                ).grid(row=0, column=0)
            elif state == "error":
                theme.make_button(
                    rw["action_frame"], self.t("dialog.local_model_manager.retry"),
                    command=lambda cn=cname: _start_download(cn),
                    style="secondary", height=22,
                ).grid(row=0, column=0)
            elif downloaded:
                theme.make_button(
                    rw["action_frame"], self.t("dialog.local_model_manager.delete"),
                    command=lambda cn=cname: _delete_model(cn),
                    style="ghost", height=22,
                ).grid(row=0, column=0)
            else:
                theme.make_button(
                    rw["action_frame"], self.t("dialog.local_model_manager.download"),
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
                "had_error": False,
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
                        cum = dl_cumulative.get(cname)
                        if cum is not None:
                            cum["had_error"] = True
                    elif line.startswith("PHASE converting"):
                        rw = row_widgets.get(cname)
                        if rw and "progress_label" in rw:
                            rw["progress_label"].configure(
                                text=self.t("dialog.local_model_manager.progress.converting"), text_color=p.status_info)
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
                        cum = dl_cumulative.get(cname) or {}
                        had_error = error_received or bool(cum.get("had_error"))
                        verified = _is_downloaded(cname)
                        dl_procs.pop(cname, None)
                        dl_queues.pop(cname, None)
                        dl_cumulative.pop(cname, None)
                        dl_state[cname] = "done" if (not had_error and verified) else "error"
                        _update_row_status(cname)
                        if not had_error and verified:
                            self._refresh_language_dropdowns()
                        elif not had_error:
                            self._log(
                                f"Model download finished, but {cname!r} is not discoverable "
                                f"as a ready model under {model_dir!r}."
                            )
                    else:
                        win.after(100, _poll)  # keep polling until the thread finishes
                    return
                rc = p_obj.poll()
                if rc is None:
                    win.after(100, _poll)
                else:
                    cum = dl_cumulative.get(cname) or {}
                    had_error = bool(cum.get("had_error"))
                    verified = rc == 0 and _is_downloaded(cname)
                    dl_procs.pop(cname, None)
                    dl_queues.pop(cname, None)
                    dl_cumulative.pop(cname, None)
                    dl_state[cname] = "done" if (verified and not had_error) else "error"
                    _update_row_status(cname)
                    if verified and not had_error:
                        self._refresh_language_dropdowns()
                    elif rc == 0 and not had_error:
                        self._log(
                            f"Model download exited successfully, but {cname!r} is not "
                            f"discoverable as a ready model under {model_dir!r}."
                        )

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
                                         hf_repo=hf_repo, pair_name=cname)
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
                "--pair-name", cname,
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
            batch_frame, self.t("dialog.local_model_manager.download_selected"),
            command=_batch_download, style="primary", height=28,
        ).pack(side=tk.LEFT, padx=(0, 8))
        theme.make_button(
            batch_frame, self.t("dialog.local_model_manager.delete_selected"),
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
        win.title(self.t("dialog.about_title"))
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

        theme.make_label(brand_frame, self.t("dialog.about_heading"), level="heading").grid(
            row=0, column=1, sticky="sw",
        )
        # Small muted beta badge beside the product heading
        theme.make_label(
            brand_frame, self.t("dialog.about_byline"), level="tiny", text_color=p.text_muted,
        ).grid(row=0, column=2, sticky="sw", padx=(6, 0))
        theme.make_label(brand_frame, self.t("dialog.about_subtitle"), level="small").grid(
            row=1, column=1, sticky="nw",
        )

        # --- Version & date ---
        theme.make_label(card, self.t("dialog.about_version", version=APP_VERSION), level="body").grid(
            row=1, column=0, sticky="w", padx=PAD, pady=(0, 2),
        )
        theme.make_label(card, self.t("dialog.about_build_date", build_date=APP_BUILD_DATE), level="body").grid(
            row=2, column=0, sticky="w", padx=PAD, pady=(0, 2),
        )

        # --- Copyright ---
        theme.make_label(
            card, self.t("dialog.about_copyright"),
            level="tiny",
        ).grid(row=3, column=0, sticky="w", padx=PAD, pady=(0, PAD))

        # --- Divider ---
        theme.make_divider(card).grid(row=4, column=0, sticky="ew", padx=PAD, pady=(0, 8))

        # --- Check for updates ---
        check_frame = ctk.CTkFrame(card, fg_color="transparent")
        check_frame.grid(row=5, column=0, sticky="w", padx=PAD, pady=(0, 4))

        update_status_var = tk.StringVar(value="")
        update_status_type = tk.StringVar(value="")

        def _do_check():
            update_status_type.set("checking")
            update_status_var.set(self.t("dialog.about.update.checking"))
            self._set_button_disabled(check_btn, True)

            def _on_result(result):
                self._set_button_disabled(check_btn, False)
                if result["status"] == "update":
                    update_status_type.set("available")
                    update_status_var.set(self.t("dialog.about.update.available", version=result['version']))
                    self._show_update_dialog(result)
                elif result["status"] == "latest":
                    update_status_type.set("latest")
                    update_status_var.set(self.t("dialog.about.update.latest"))
                else:
                    update_status_type.set("error")
                    update_status_var.set(self.t("dialog.about.update.error"))

            self._run_update_check(startup=False, callback=_on_result)

        check_btn = theme.make_button(
            check_frame, self.t("dialog.about.check_for_updates"), command=_do_check, style="ghost", height=28,
        )
        check_btn.pack(side=tk.LEFT)

        update_status_lbl = theme.make_label(check_frame, "", level="small")
        update_status_lbl.pack(side=tk.LEFT, padx=(10, 0))

        def _sync_status(*_):
            val = update_status_var.get()
            stype = update_status_type.get()
            if stype == "latest":
                color = p.status_success
            elif stype == "available":
                color = p.text_secondary
            elif stype == "error":
                color = p.status_error
            else:
                color = p.text_muted
            update_status_lbl.configure(text=val, text_color=color)

        update_status_var.trace_add("write", _sync_status)

        # Surface result from an already-completed startup check
        if self._update_check_result:
            r = self._update_check_result
            if r["status"] == "latest":
                update_status_type.set("latest")
                update_status_var.set(self.t("dialog.about.update.latest"))
            elif r["status"] == "update":
                update_status_type.set("available")
                update_status_var.set(self.t("dialog.about.update.available", version=r['version']))
            elif r["status"] == "error":
                update_status_type.set("error")
                update_status_var.set(self.t("dialog.about.update.error"))

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

        theme.make_button(links_frame, self.t("dialog.about.github"), command=_open_github, style="ghost",
                          height=28).pack(side=tk.LEFT, padx=(0, theme.scale(8)))
        theme.make_button(links_frame, self.t("dialog.about.release_notes"), command=_open_releases, style="ghost",
                          height=28).pack(side=tk.LEFT, padx=(0, theme.scale(8)))
        theme.make_button(links_frame, self.t("dialog.about.report_bug"), command=_open_issue, style="ghost",
                          height=28).pack(side=tk.LEFT, padx=(0, theme.scale(8)))

        win.protocol("WM_DELETE_WINDOW", win.destroy)
        self._bind_modal_keys(win, win.destroy, initial_focus=check_btn)
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

    def _path_key(self, path: str) -> str:
        try:
            return os.path.normcase(str(Path(path).expanduser().resolve()))
        except Exception:
            return os.path.normcase(os.path.abspath(os.path.expanduser(str(path))))

    def _add_paths(self, paths, *, announce: bool = True) -> int:
        """Add files/directories from picker, startup defaults, or drag-and-drop.

        Directories use the same non-recursive supported-file discovery as the
        existing Select Folder action. Paths are normalised before deduping so
        Windows case differences and relative/absolute aliases cannot duplicate
        a row.
        """
        if getattr(self, "_running", False):
            try:
                self._log(self.t("content.file_list_locked"))
            except Exception:
                pass
            return 0
        existing = {self._path_key(path) for path in self.files}
        added = 0
        for raw in paths or ():
            if not raw:
                continue
            candidate = Path(str(raw)).expanduser()
            candidates = list_supported_files(str(candidate)) if candidate.is_dir() else [str(candidate)]
            for item in candidates:
                p = Path(item).expanduser()
                if not p.is_file() or p.suffix.lower() not in SUPPORTED_EXTS:
                    continue
                try:
                    resolved = str(p.resolve())
                except Exception:
                    resolved = str(p.absolute())
                key = self._path_key(resolved)
                if key in existing:
                    continue
                self._add_file_to_table(resolved)
                existing.add(key)
                added += 1
        if announce and added:
            try:
                self._log(self.t("content.drop_added", count=added))
            except Exception:
                pass
        return added

    def _install_drag_and_drop(self):
        try:
            return install_file_drop(
                self.root, self.root,
                lambda items: self._add_paths(items),
            )
        except Exception:
            logger.debug("Drag-and-drop setup failed", exc_info=True)
            return None

    def _add_files(self):
        init = self._initialdir_for_input()
        all_patterns = tuple(f"*{ext}" for ext in SUPPORTED_EXTS)

        def ext_to_label(ext: str) -> str:
            return f"{ext.lstrip('.').upper()} Files"

        _translate = getattr(self, "t", None)
        _all_supported = _translate("file_dialog.all_supported") if callable(_translate) else "All Supported Files"
        _all_files = _translate("file_dialog.all_files") if callable(_translate) else "All Files"
        _select_title = _translate("file_dialog.select_files") if callable(_translate) else "Select Files"
        filetypes = [(_all_supported, all_patterns)]
        seen = set()
        for ext in SUPPORTED_EXTS:
            if ext not in seen:
                seen.add(ext)
                filetypes.append((ext_to_label(ext), (f"*{ext}",)))
        filetypes.append((_all_files, ("*.*",)))

        paths = filedialog.askopenfilenames(
            title=_select_title, parent=self.root, initialdir=init, filetypes=filetypes,
        )
        if paths:
            self._add_paths(paths, announce=False)

    def _select_folder(self):
        init = self._initialdir_for_input()
        d = filedialog.askdirectory(
            title=self.t("file_dialog.select_folder"), parent=self.root, initialdir=init,
        )
        if not d:
            return
        found = list_supported_files(d)
        if not found:
            messagebox.showinfo(
                self.t("message.no_files_title"),
                self.t("message.no_files_found_body", path=d),
                parent=self.root,
            )
            return
        self._add_paths(found, announce=False)

    def _clear_files(self):
        # Freeze the job list while a worker owns its start-time snapshot.
        if getattr(self, "_running", False):
            return
        # removes selected file, or clears all if nothing selected
        selected = self.file_table.selection()
        if selected:
            self._remove_selected_files()
        else:
            self.files.clear()
            self._tree_ids.clear()
            self._file_to_iid.clear()
            self._file_status.clear()
            self._file_attempts.clear()
            self._active_run_files = ()
            self._file_start_times.clear()
            for iid in self.file_table.get_children():
                self.file_table.delete(iid)
            self.completed_files = 0
            self.total_files = 0
            self._set_progress(0.0)
            self._update_progress_label(self.t("content.ready"))
            self._sync_retry_button_visibility()
            try:
                self.log.configure(state="normal")
                self.log.delete("1.0", "end")
                self.log.configure(state="disabled")
            except Exception:
                pass
            self._sync_file_empty_state()
            self._sync_clear_button_label()

    def _select_output(self):
        init = self._initialdir_for_output()
        d = filedialog.askdirectory(
            title=self.t("file_dialog.select_output_folder"), parent=self.root, initialdir=init,
        )
        if d:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, _try_make_relative(d))

    def _schedule_language_dropdown_refresh(self) -> None:
        """Refresh language choices after the current Tk event finishes.

        Dropdown commands run from inside a popup ButtonRelease handler.  Rebuilding
        other dropdowns synchronously from that handler can leave Windows/Tk focus
        ownership attached to the popup that is about to be destroyed.  Deferring
        the refresh by one event-loop turn makes the engine/detector selection fully
        finish first.
        """

        pending = getattr(self, "_language_refresh_after_id", None)
        if pending is not None:
            try:
                self.root.after_cancel(pending)
            except Exception:
                pass

        def _apply() -> None:
            self._language_refresh_after_id = None
            if getattr(self, "_closing", False):
                return
            self._refresh_language_dropdowns()

        try:
            self._language_refresh_after_id = self.root.after_idle(_apply)
        except Exception:
            self._language_refresh_after_id = None
            _apply()

    def _on_detector_changed(self, *_) -> None:
        # Repopulate after the detector dropdown has completely finished its own
        # selection event; do not rebuild searchable entries from inside it.
        self._schedule_language_dropdown_refresh()

    def _on_engine_changed(self, *_) -> None:
        engine_key = self._engine_key_for_display(self.engine_var.get())
        self.cfg["translation_engine"] = engine_key
        save_config(self.cfg)
        self._log(f"Translation engine changed to: {engine_key!r}")
        self._update_usage_label(engine_key)

        # Same rule as the detector: let the engine popup close and focus settle
        # before source/target searchable dropdowns are repopulated.
        self._schedule_language_dropdown_refresh()

    def _on_source_lang_changed(self, *_) -> None:
        # update_values() may legitimately change source_lang_var while an engine
        # refresh is already rebuilding both lists.  Its trace must not start a
        # second target refresh inside the first one.
        if getattr(self, "_refreshing_language_dropdowns", False):
            return

        # Only local engine uses source-based target filtering.
        engine_key = self._engine_key_for_display(self.engine_var.get())
        if engine_key != "local":
            return

        from ..translators.local import list_downloaded_pairs
        model_dir = _get_local_model_dir_from_cfg(self.cfg)
        pairs = list_downloaded_pairs(model_dir)
        lang_opts = _get_language_options(self.ui.locale)

        source_code = self._source_lang_map.get(self.source_lang_var.get())

        if not source_code or source_code == "auto":
            tgt_codes = {_opus_code_to_iso(t) for _, t in pairs}
        else:
            tgt_codes = set()
            for s, t in pairs:
                if _opus_code_to_iso(s) == source_code:
                    tgt_codes.add(_opus_code_to_iso(t))

        tgt_filtered = _get_local_language_options(tgt_codes, self.ui.locale)
        tgt_display = [self._format_language_option(code, name) for code, name in tgt_filtered]
        self._lang_map = {self._format_language_option(code, name): code for code, name in tgt_filtered}
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
        """Refresh source/target choices as one non-reentrant UI transaction."""
        if getattr(self, "_refreshing_language_dropdowns", False):
            return
        self._refreshing_language_dropdowns = True
        try:
            self._refresh_language_dropdowns_impl()
        finally:
            self._refreshing_language_dropdowns = False

    def _refresh_language_dropdowns_impl(self) -> None:
        # Recompute source and target language lists based on current engine + detector
        detector = (self.detector_var.get() or "fasttext").strip().lower()
        engine_key = self._engine_key_for_display(self.engine_var.get())
        lang_opts = _get_language_options(self.ui.locale)

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

            # Source list — languages which appear as source in at least one pair.
            # Auto-detect is valid too: the local backend groups units by the
            # detected source language and selects the matching installed pair.
            src_filtered = _get_local_language_options(src_codes_set, self.ui.locale)
            src_display = [self._format_language_option(code, name) for code, name in src_filtered]
            auto_detect_label = self.t("sidebar.auto_detect")
            source_values = [auto_detect_label] + src_display
            self._source_lang_map = {auto_detect_label: "auto"}
            self._source_lang_map.update({
                self._format_language_option(code, name): code for code, name in src_filtered
            })
            self.source_lang_box.update_values(source_values)
            self._source_lang_label.configure(
                text=self.t("sidebar.source_language_count", count=len(src_display)),
            )

            # Target list — all targets reachable from any source
            tgt_filtered = _get_local_language_options(tgt_codes_set, self.ui.locale)
            tgt_display = [self._format_language_option(code, name) for code, name in tgt_filtered]
            self._lang_map = {self._format_language_option(code, name): code for code, name in tgt_filtered}
            self.target_lang_box.update_values(tgt_display)

            self._log(f"Language lists updated for engine={engine_key!r} "
                      f"(target: {len(tgt_display)}, source: {len(src_display)}, "
                      f"downloaded pairs: {len(pairs)})")
            return

        # Non-local engines: original logic
        # Target: filter by engine only (detector doesn't restrict target)
        tgt_filtered = _filter_by_engine(lang_opts, engine_key)

        # Source: intersection of engine-supported and detector-supported
        src_filtered = _filter_by_engine(_filter_by_detector(lang_opts, detector), engine_key)

        # Ollama post-filter: if a model with a known language restriction is
        # active, narrow both lists to the languages it actually supports.
        if self.cfg.get("ollama_enabled"):
            from ..translators.ollama import ollama_get_supported_lang_codes
            _ollama_codes = ollama_get_supported_lang_codes(
                self.cfg.get("ollama_model", "")
            )
            if _ollama_codes is not None:
                tgt_filtered = [(c, n) for c, n in tgt_filtered if c.lower() in _ollama_codes]
                src_filtered = [(c, n) for c, n in src_filtered if c.lower() in _ollama_codes]

        tgt_display = [self._format_language_option(code, name) for code, name in tgt_filtered]
        self._lang_map = {self._format_language_option(code, name): code for code, name in tgt_filtered}
        self.target_lang_box.update_values(tgt_display)

        src_display = [self._format_language_option(code, name) for code, name in src_filtered]
        auto_detect_label = self.t("sidebar.auto_detect")
        source_values = [auto_detect_label] + src_display
        self._source_lang_map = {auto_detect_label: "auto"}
        self._source_lang_map.update(
            {self._format_language_option(code, name): code for code, name in src_filtered}
        )
        self.source_lang_box.update_values(source_values)
        self._source_lang_label.configure(
            text=self.t("sidebar.source_language_count", count=len(src_display)),
        )
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

    def _start(self, run_files: list[str] | tuple[str, ...] | None = None):
        if self._running:
            return

        if not self.files:
            messagebox.showwarning(
                self.t("message.no_files_title"),
                self.t("message.no_files_selected_body"),
                parent=self.root,
            )
            return

        if run_files is None:
            files_to_run = self._pending_files()
        else:
            requested = set(run_files)
            files_to_run = [path for path in self.files if path in requested]

        if not files_to_run:
            if any(self._file_status.get(path) in {"error", "cancelled"} for path in self.files):
                messagebox.showinfo(
                    self.t("queue.nothing_pending_title"),
                    self.t("queue.nothing_pending_retry"),
                    parent=self.root,
                )
            else:
                messagebox.showinfo(
                    self.t("queue.nothing_pending_title"),
                    self.t("queue.nothing_pending_body"),
                    parent=self.root,
                )
            return

        # Resolve target language
        sel = self.lang_var.get()
        lang = self._lang_map.get(sel)
        if not lang:
            typed = self.target_lang_box.get()
            lang = self._lang_map.get(typed)
        if not lang:
            messagebox.showwarning(self.t("message.missing_language_title"), self.t("message.missing_language_body"))
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
                self._log(self.t("log.output_path_missing", output=output))
                return
        output = str(out_path)

        # Resolve language detector selection
        detector = (self.detector_var.get() or "fasttext").strip().lower()
        if detector not in ("fasttext", "lingua"):
            detector = "fasttext"

        # Resolve translation engine and credentials from config
        engine = self._engine_key_for_display(self.engine_var.get())
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
        ollama_config = {
            "enabled": bool(self.cfg.get("ollama_enabled", False)),
            "model": str(self.cfg.get("ollama_model", "")).strip(),
            "base_url": str(self.cfg.get("ollama_base_url", "")).strip(),
        }

        # Semantic translation pre-flight: ensure Ollama is installed and the server is running
        if ollama_config["enabled"]:
            from ..translators.ollama import (
                OllamaNotInstalledError,
                _ensure_ollama_server_no_install,
                DEFAULT_OLLAMA_BASE_URL as _DEFAULT_OLLAMA_URL,
            )
            _ollama_url = ollama_config["base_url"] or _DEFAULT_OLLAMA_URL
            try:
                _ensure_ollama_server_no_install(_ollama_url, proxies=proxies)
            except OllamaNotInstalledError:
                messagebox.showwarning(
                    self.t("message.ollama_not_ready_title"),
                    self.t("message.ollama_not_ready_body"),
                )
                return
            except RuntimeError:
                messagebox.showwarning(
                    self.t("message.ollama_not_ready_title"),
                    self.t("message.ollama_server_unreachable_body"),
                )
                return

        # Validate credentials for engines that require them
        if engine == "baidu" and (not baidu_appid or not baidu_appkey):
            messagebox.showwarning(
                self.t("message.missing_credentials_title"),
                self.t("message.baidu_missing_credentials"),
            )
            return
        if engine == "google-cloud-v3" and not google_project_id:
            messagebox.showwarning(
                self.t("message.missing_credentials_title"),
                self.t("message.google_cloud_v3_missing_credentials"),
            )
            return
        if engine == "google-cloud" and not google_api_key:
            messagebox.showwarning(
                self.t("message.missing_api_key_title"),
                self.t("message.google_cloud_missing_api_key"),
            )
            return
        if engine == "azure" and (not azure_key or not azure_region):
            messagebox.showwarning(
                self.t("message.missing_credentials_title"),
                self.t("message.azure_missing_credentials"),
            )
            return
        if engine == "deepl" and not deepl_api_key:
            messagebox.showwarning(
                self.t("message.missing_api_key_title"),
                self.t("message.deepl_missing_api_key"),
            )
            return

        # Local engine: verify required model is downloaded
        if engine == "local":
            from ..translators.local import list_downloaded_pairs
            _lm_dir = local_model_dir or _get_default_model_dir()
            _local_pairs = list_downloaded_pairs(_lm_dir)
            if not _local_pairs:
                if messagebox.askyesno(
                    self.t("message.no_models_downloaded_title"),
                    self.t("message.no_models_downloaded_body"),
                ):
                    self._open_settings()
                return
            _pair_set = {(_opus_code_to_iso(s), _opus_code_to_iso(t))
                         for s, t in _local_pairs}
            # In auto mode the local backend detects and groups each document unit by language
            # Requiring an impossible "auto → target" model here blocked that mode before translation began
            has_target_pair = any(tgt == lang for _, tgt in _pair_set)
            if (source_lang != "auto" and (source_lang, lang) not in _pair_set) or (
                source_lang == "auto" and not has_target_pair
            ):
                if messagebox.askyesno(
                    self.t("message.missing_model_title"),
                    self.t(
                        "message.missing_model_body",
                        source=("detected source language" if source_lang == "auto" else source_lang),
                        target=lang,
                    ),
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
                    engine_display = self._engine_display_name(engine)
                    if warning == "limit":
                        msg = self.t("message.usage_limit_body", usage=used_str, engine=engine_display)
                    else:
                        msg = self.t("message.usage_warn_body", usage=used_str, engine=engine_display)
                    if not messagebox.askyesno(self.t("message.usage_warning_title"), msg):
                        return
                    self._update_usage_label(engine)
            except Exception:
                pass

        terminology_snapshot = self._terminology_store.snapshot(source_lang, lang)
        if self._terminology_store.last_error:
            messagebox.showwarning(
                self.t("terminology.load_error_title"),
                self.t("terminology.load_error_body", error=self._terminology_store.last_error),
            )
            return

        tm_enabled = bool(self.cfg.get("translation_memory_enabled", False))
        tm_path = str(self.cfg.get("translation_memory_path") or "").strip()
        if tm_enabled:
            try:
                TranslationMemory(tm_path or default_translation_memory_path(), timeout=0.75).count()
            except Exception as exc:
                messagebox.showwarning(
                    self.t("settings.section.translation_memory"),
                    self.t("settings.error.translation_memory", error=exc),
                )
                return

        try:
            self._log(
                f"Starting: engine={engine!r}, source={source_lang!r}, target={lang!r}, "
                f"detector={detector!r}, terminology={len(terminology_snapshot.mapping)}, "
                f"translation_memory={tm_enabled}"
            )
            if terminology_snapshot.conflicts:
                self._log(
                    "Ignored ambiguous auto-source terminology: "
                    + ", ".join(terminology_snapshot.conflicts)
                )
        except Exception:
            pass

        # Update UI state
        self._running = True
        self._set_button_disabled(self.start_btn, True)
        self._set_button_disabled(self.cancel_btn, False)
        for button in (self.add_files_btn, self.select_folder_btn, self.clear_files_btn):
            self._set_button_disabled(button, True)

        self._active_run_files = tuple(files_to_run)
        for path in self._active_run_files:
            self._file_attempts[path] = int(self._file_attempts.get(path, 0)) + 1
            self._update_file_status(path, "pending")
        self.total_files = len(self._active_run_files)
        self.completed_files = 0
        self._file_start_times.clear()
        self._set_progress(0.0)
        self._update_progress_label(self.t("progress.starting"))
        self._current_file_name = ""
        self._last_run_report = None
        self._last_run_cancelled = False
        try:
            self._set_button_disabled(self.report_btn, True)
            self.report_btn.grid_remove()
            self._set_button_disabled(self.retry_failed_btn, True)
            self.retry_failed_btn.grid_remove()
        except Exception:
            pass

        self.worker.start(
            self._active_run_files, lang, output, None,
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
            ollama_config=ollama_config,
            terminology=terminology_snapshot.mapping,
            translation_memory_enabled=tm_enabled,
            translation_memory_path=tm_path or None,
            report_cb=lambda report: self._log_queue.put(("__report__", report)),
            terminology_conflicts=terminology_snapshot.conflicts,
            terminology_entries=terminology_snapshot.matched_entries,
            attempt_numbers={path: self._file_attempts.get(path, 1) for path in self._active_run_files},
        )

    def _pending_files(self) -> list[str]:
        return [
            path for path in self.files
            if self._file_status.get(path, "pending") == "pending"
        ]

    def _retryable_files(self) -> list[str]:
        return [
            path for path in self.files
            if self._file_status.get(path) in {"error", "cancelled"}
        ]

    def _retry_failed(self):
        if self._running:
            return
        retryable = self._retryable_files()
        if not retryable:
            self._sync_retry_button_visibility()
            return
        self._start(retryable)

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
                display_status = "retrying" if self._file_attempts.get(filepath, 1) > 1 else "started"
                self._update_file_status(filepath, display_status)
                self._current_file_name = Path(filepath).name
                self._update_progress_label(
                    self.t("progress.translating_file", name=self._current_file_name)
                )
            elif status == "progress":
                # elapsed carries the global fraction (0.0–1.0) for intra-file progress
                frac = elapsed if elapsed is not None else 0.0
                frac = max(0.0, min(1.0, frac))
                self._set_progress(frac)
                name = getattr(self, "_current_file_name", "")
                self._update_progress_label(
                    self.t("progress.translating_percent", name=name, percent=int(frac * 100))
                )
            elif status in ("finished", "error"):
                self.completed_files += 1
                self._update_file_status(filepath, status, elapsed)
                pct = self.completed_files / max(1, self.total_files)
                self._set_progress(pct)
                remaining = self.total_files - self.completed_files
                if remaining > 0:
                    self._update_progress_label(
                        self.t(
                            "progress.files_complete",
                            completed=self.completed_files,
                            total=self.total_files,
                            percent=int(pct * 100),
                        )
                    )
                else:
                    self._update_progress_label(
                        self.t(
                            "progress.files_complete",
                            completed=self.completed_files,
                            total=self.total_files,
                            percent=100,
                        )
                    )
                if self.completed_files >= self.total_files:
                    self._finish_run()
            elif status == "cancelled":
                self._update_file_status(filepath, "cancelled")
                for f in self._active_run_files:
                    iid = self._file_to_iid.get(f)
                    if iid:
                        vals = self.file_table.item(iid, "values")
                        if vals and vals[0] == self._status_text("pending"):
                            self._update_file_status(f, "cancelled")
                self._finish_run(cancelled=True)

        self.root.after(0, _update)

    def _log(self, msg: str):
        # thread-safe: safe to call from any thread
        self._log_queue.put(msg)

    def _poll_log_queue(self):
        # Processes ONE message per call so that Tk can run its widget-repaint
        # idle callbacks between insertions — each line appears immediately.
        #
        # When a message is found, reschedule via after_idle: this fires only
        # after all pending events (including widget redraws) have been processed,
        # so the text widget repaints before the next message is inserted.
        # When the queue is empty, fall back to a 50 ms timer to avoid spinning.
        try:
            msg = self._log_queue.get_nowait()
        except queue.Empty:
            self.root.after(50, self._poll_log_queue)
            return
        except Exception:
            self.root.after(50, self._poll_log_queue)
            return

        try:
            if isinstance(msg, tuple) and len(msg) == 2 and msg[0] == "__report__":
                self._handle_run_report(msg[1])
            elif msg == "__worker_done__":
                if self._running:
                    self._finish_run(cancelled=self.worker.last_run_cancelled)
            else:
                self.log.configure(state="normal")
                self.log.insert("end", str(msg) + "\n")
                self.log.see("end")
                self.log.configure(state="disabled")
        except Exception:
            pass

        self.root.after_idle(self._poll_log_queue)


    def _handle_run_report(self, report: dict):
        self._last_run_report = report if isinstance(report, dict) else None
        if self._last_run_report is None:
            return
        totals = self._last_run_report.get("totals", {}) or {}
        # The report is authoritative for terminal job states. This covers
        # cancellation before a per-file callback reached Tk and other races
        # between worker callbacks and the report queue.
        terminal_statuses = {"finished", "error", "cancelled", "skipped"}
        for item in self._last_run_report.get("files", []) or []:
            path = str(item.get("path", ""))
            status = str(item.get("status", ""))
            if path in self._file_to_iid and status in terminal_statuses:
                self._update_file_status(
                    path, status, float(item.get("elapsed_seconds", 0) or 0),
                )
                try:
                    self._file_attempts[path] = max(
                        self._file_attempts.get(path, 0), int(item.get("attempt", 1) or 1),
                    )
                except Exception:
                    pass
        warning_count = int(self._last_run_report.get("warning_count", 0) or 0)
        if warning_count:
            self._update_progress_label(
                self.t("report.completed_with_warnings", count=warning_count)
            )
        else:
            self._update_progress_label(self.t("report.completed"))
        try:
            self._log(
                "Translation report: "
                f"{totals.get('files_finished', 0)} finished, "
                f"{totals.get('files_failed', 0)} failed; "
                f"{totals.get('translated_units', 0)} translated units, "
                f"{totals.get('skipped_units', 0)} skipped, "
                f"{totals.get('tm_hits', 0)} TM hits, "
                f"{warning_count} warning(s)"
            )
        except Exception:
            pass

        self._sync_report_button_visibility()
        self._sync_retry_button_visibility()

    def _sync_retry_button_visibility(self) -> None:
        button = getattr(self, "retry_failed_btn", None)
        if button is None:
            return
        available = (not getattr(self, "_running", False)) and bool(self._retryable_files())
        try:
            if available:
                self._set_button_disabled(button, False)
                button.grid()
            else:
                self._set_button_disabled(button, True)
                button.grid_remove()
        except Exception:
            pass

    def _sync_report_button_visibility(self) -> None:
        """Show the report affordance only after a non-cancelled run is over."""
        button = getattr(self, "report_btn", None)
        if button is None:
            return
        available = (
            not getattr(self, "_running", False)
            and not getattr(self, "_last_run_cancelled", False)
            and isinstance(getattr(self, "_last_run_report", None), dict)
        )
        try:
            if available:
                self._set_button_disabled(button, False)
                button.grid()
            else:
                self._set_button_disabled(button, True)
                button.grid_remove()
        except Exception:
            pass

    def _open_translation_report(self):
        report = self._last_run_report
        if not report:
            messagebox.showinfo(self.t("report.title"), self.t("report.no_report"), parent=self.root)
            return
        p = theme.get()
        totals = report.get("totals", {}) or {}
        win = ctk.CTkToplevel(self.root)
        win.wm_attributes("-alpha", 0)
        apply_window_icon(win)
        win.title(self.t("report.title"))
        win.transient(self.root)
        win.grab_set()
        win.configure(fg_color=p.bg_main)
        win.geometry("880x600")
        win.minsize(720, 480)
        win.grid_columnconfigure(0, weight=1)
        win.grid_rowconfigure(1, weight=1)

        summary = theme.make_card(win)
        summary.grid(row=0, column=0, sticky="ew", padx=theme.PADDING, pady=(theme.PADDING, 8))
        summary.grid_columnconfigure(0, weight=1)
        theme.make_label(summary, self.t("report.title"), level="heading").grid(row=0, column=0, sticky="w", padx=12, pady=(10, 6))
        lines = [
            self.t(
                "report.summary.files",
                finished=totals.get("files_finished", 0), failed=totals.get("files_failed", 0),
                skipped=totals.get("files_skipped", 0), cancelled=totals.get("files_cancelled", 0),
            ),
            self.t(
                "report.summary.units",
                units=totals.get("units", 0), translated=totals.get("translated_units", 0),
                skipped=totals.get("skipped_units", 0),
            ),
            self.t(
                "report.summary.selective",
                target=totals.get("target_language_skips", 0), nonling=totals.get("nonlinguistic_skips", 0),
                other=totals.get("non_source_skips", 0),
            ),
            self.t(
                "report.summary.consistency",
                reuses=totals.get("consistency_reuses", 0),
                families=totals.get("consistency_families", 0),
            ),
            self.t(
                "report.summary.cache",
                hits=totals.get("persistent_cache_hits", 0),
                writes=totals.get("persistent_cache_writes", 0),
                errors=totals.get("persistent_cache_errors", 0),
            ),
            self.t(
                "report.summary.tm",
                hits=totals.get("tm_hits", 0), writes=totals.get("tm_writes", 0),
                errors=totals.get("tm_errors", 0),
            ),
            self.t(
                "report.summary.terminology",
                entries=report.get("terminology_entries", 0), mismatches=totals.get("terminology_mismatches", 0),
            ),
            self.t(
                "report.summary.quality",
                fallbacks=totals.get("fallback_items", 0), retries=totals.get("retry_items", 0),
                warnings=totals.get("constraint_warnings", 0), failed=totals.get("failed_items", 0),
            ),
            self.t(
                "report.summary.visual",
                checks=totals.get("output_validation_checks", 0),
                failures=totals.get("output_validation_failures", 0),
                retry_candidates=totals.get("layout_retry_candidates", 0),
                retry_accepted=totals.get("layout_retry_accepted", 0),
                autofit=totals.get("pptx_autofit_adjustments", 0),
                visual_warnings=(
                    totals.get("visual_overflow_warnings", 0)
                    + totals.get("visual_compression_warnings", 0)
                ),
            ),
        ]
        for row, text in enumerate(lines, start=1):
            theme.make_label(summary, text, level="small").grid(row=row, column=0, sticky="w", padx=12, pady=(0, 3))
        conflicts = report.get("terminology_conflicts", []) or []
        if conflicts:
            theme.make_label(
                summary, self.t("report.terminology_conflicts", terms=", ".join(conflicts)),
                level="tiny", text_color=p.status_warning,
            ).grid(row=len(lines) + 1, column=0, sticky="w", padx=12, pady=(3, 10))
        else:
            summary.grid_configure(pady=(theme.PADDING, 8))

        card = theme.make_card(win)
        card.grid(row=1, column=0, sticky="nsew", padx=theme.PADDING, pady=(0, 8))
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(0, weight=1)
        tree = ttk.Treeview(
            card, columns=("status", "time", "details"), show="tree headings",
            selectmode="browse", takefocus=True, style="FileTable.Treeview",
        )
        _style_app_table(tree, p)
        tree.heading("#0", text=self.t("table.file"))
        tree.heading("status", text=self.t("table.status"))
        tree.heading("time", text=self.t("table.time"))
        tree.heading("details", text=self.t("report.column.details"))
        tree.column("#0", width=250, stretch=True)
        tree.column("status", width=90, stretch=False, anchor="center")
        tree.column("time", width=90, stretch=False, anchor="center")
        tree.column("details", width=330, stretch=True)
        scroll = ttk.Scrollbar(
            card, orient=tk.VERTICAL, command=tree.yview,
            style="Slim.Vertical.TScrollbar",
        )
        tree.configure(yscrollcommand=scroll.set)
        tree.grid(row=0, column=0, sticky="nsew", padx=(12, 0), pady=12)
        scroll.grid(row=0, column=1, sticky="ns", padx=(0, 12), pady=12)
        tree.tag_configure("finished", foreground=p.status_success)
        tree.tag_configure("error", foreground=p.status_error)
        tree.tag_configure("cancelled", foreground=p.status_warning)
        tree.tag_configure("skipped", foreground=p.status_warning)
        tree.tag_configure("pending", foreground=p.status_pending)
        tree.tag_configure("started", foreground=p.status_info)
        tree.tag_configure("retrying", foreground=p.status_info)
        for item in report.get("files", []) or []:
            metrics = item.get("metrics", {}) or {}
            visual_warnings = (
                int(metrics.get("visual_overflow_warnings", 0) or 0)
                + int(metrics.get("visual_compression_warnings", 0) or 0)
            )
            attempt = max(1, int(item.get("attempt", 1) or 1))
            attempt_prefix = self.t("report.row.attempt", attempt=attempt) if attempt > 1 else ""
            details = self.t(
                "report.row.details",
                attempt=attempt_prefix,
                translated=metrics.get("translated_units", 0),
                skipped=metrics.get("skipped_units", 0),
                tm=metrics.get("tm_hits", 0),
                warnings=visual_warnings,
            )
            if item.get("error"):
                details = str(item.get("error"))
            item_status = str(item.get("status", ""))
            tree.insert(
                "", tk.END, text=Path(str(item.get("path", ""))).name,
                values=(
                    self._status_text(str(item.get("status", ""))), self._format_elapsed_time(float(item.get("elapsed_seconds", 0) or 0)), details,
                ),
                tags=(item_status,) if item_status else (),
            )
        _restripe_tree(tree)
        summary_text = "\n".join(lines + ([self.t("report.terminology_conflicts", terms=", ".join(conflicts))] if conflicts else []))

        def _copy_summary():
            try:
                win.clipboard_clear()
                win.clipboard_append(summary_text)
                copy_btn.configure(text=self.t("report.copied"))
                win.after(1200, lambda: copy_btn.configure(text=self.t("report.copy_summary")) if win.winfo_exists() else None)
            except Exception:
                pass

        def _close_report():
            win.destroy()

        bottom = ctk.CTkFrame(win, fg_color="transparent")
        bottom.grid(row=2, column=0, sticky="ew", padx=theme.PADDING, pady=(0, theme.PADDING))
        bottom.grid_columnconfigure(0, weight=1)
        copy_btn = theme.make_button(bottom, self.t("report.copy_summary"), command=_copy_summary, style="ghost", height=28)
        copy_btn.grid(row=0, column=0, sticky="w")
        theme.make_button(bottom, self.t("terminology.close"), command=_close_report, style="secondary", height=28).grid(row=0, column=1, sticky="e")
        win.protocol("WM_DELETE_WINDOW", _close_report)
        self._bind_modal_keys(win, _close_report, initial_focus=tree)
        center_window(win, 880, 600, parent=self.root)
        win.wm_attributes("-alpha", 1)

    def _finish_run(self, cancelled: bool = False):
        if not self._running:
            return
        self._running = False
        self._last_run_cancelled = bool(cancelled)
        self._set_button_disabled(self.start_btn, False)
        self._set_button_disabled(self.cancel_btn, True)
        for button in (self.add_files_btn, self.select_folder_btn, self.clear_files_btn):
            self._set_button_disabled(button, False)
        try:
            engine_key = self._engine_key_for_display(self.engine_var.get())
            self._update_usage_label(engine_key)
        except Exception:
            pass
        if cancelled:
            pct = self.completed_files / max(1, self.total_files)
            self._update_progress_label(
                self.t("progress.cancelled", percent=int(pct * 100)),
            )
        self._sync_report_button_visibility()
        self._sync_retry_button_visibility()
        self._active_run_files = ()

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
        try:
            root.deiconify()
            root.update_idletasks()
            center_window(root, theme.WINDOW_WIDTH, theme.WINDOW_HEIGHT)
        except Exception:
            pass
        root.wm_attributes("-alpha", 1)

    root.after(10, _show)
    root.mainloop()

if __name__ == "__main__":
    main()
