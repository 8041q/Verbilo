import threading
import time
import logging
import re
from pathlib import Path
from typing import Callable, Iterable, Union, Optional, Any
import tkinter as tk
import traceback
from urllib.parse import urlparse

from ..main import translate_file
from ..utils import CancelledError

SUPPORTED_EXTS = (".docx", ".pdf", ".xlsx", ".xls")


def _normalize_ollama_config(cfg: Optional[dict[str, Any]]) -> dict[str, Any]:
    from ..translators.ollama import (
        DEFAULT_OLLAMA_BASE_URL,
        DEFAULT_OLLAMA_MODEL,
        ollama_supports_non_pdf_translation,
        resolve_ollama_pdf_models,
    )

    raw = cfg or {}
    semantic_model = str(raw.get("model") or DEFAULT_OLLAMA_MODEL).strip() or DEFAULT_OLLAMA_MODEL
    resolved_models = resolve_ollama_pdf_models(
        semantic_model,
        advisor_model=str(raw.get("advisor_model") or "").strip() or None,
    )
    return {
        "enabled": bool(raw.get("enabled")),
        "model": resolved_models["semantic_model"],
        "advisor_model": resolved_models["advisor_model"],  # may be None for translation-only models
        "supports_non_pdf": ollama_supports_non_pdf_translation(resolved_models["semantic_model"]),
        "base_url": (
            str(raw.get("base_url") or DEFAULT_OLLAMA_BASE_URL).strip()
            or DEFAULT_OLLAMA_BASE_URL
        ),
    }


def _ollama_client_proxies(base_url: str, proxies: dict | None) -> dict | None:
    candidate = (base_url or "http://127.0.0.1:11434").strip() or "http://127.0.0.1:11434"
    if "://" not in candidate:
        candidate = f"http://{candidate}"
    hostname = urlparse(candidate).hostname or ""
    if hostname.lower() in {"localhost", "127.0.0.1", "::1", "0.0.0.0"}:
        return None
    return proxies


def list_supported_files(path: str) -> list[str]:
    p = Path(path)
    if p.is_file():
        return [str(p)] if p.suffix.lower() in SUPPORTED_EXTS else []
    if not p.exists():
        return []
    files = []
    for f in p.iterdir():
        if f.is_file() and f.suffix.lower() in SUPPORTED_EXTS:
            files.append(str(f))
    return files

def center_window(window, width=None, height=None, parent=None):
    # Center `window` on screen (parent=None) or over `parent`
    window.update_idletasks()

    if width is not None and height is not None:
        # logical -> physical for centering math only
        try:
            sf = window._get_window_scaling()
        except AttributeError:
            sf = 1.0
        sf = sf if sf > 0 else 1.0
        phys_w = round(width * sf)
        phys_h = round(height * sf)
    else:
        phys_w = window.winfo_width()
        phys_h = window.winfo_height()

    if parent is not None:
        parent.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width()  - phys_w) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - phys_h) // 2
    else:
        x = (window.winfo_screenwidth()  - phys_w) // 2
        y = (window.winfo_screenheight() - phys_h) // 2

    x = max(0, x)
    y = max(0, y)

    if width is not None and height is not None:
        window.geometry(f"{width}x{height}+{x}+{y}")
    else:
        # position-only: CTk._apply_geometry_scaling passes +x+y unchanged
        window.geometry(f"+{x}+{y}")


# runs translation in a background thread; call start() to begin, stop() to cancel
class Worker:

    def __init__(self):
        self._thread: Optional[threading.Thread] = None
        self._stop = threading.Event()

    @property
    def cancelled(self) -> bool:
        return self._stop.is_set()

    @property
    def alive(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def start(
        self,
        files: Iterable[str],
        target_lang: str,
        output_dir: Optional[str],
        translator_name: Optional[str],
        progress_cb: Callable[[str, str, Optional[float]], None],
        log_cb: Callable[[str], None],
        source_lang: str = "auto",
        detector: str = "fasttext",
        engine: str = "google",
        proxies: dict | None = None,
        google_api_key: str = "",
        baidu_appid: str = "",
        baidu_appkey: str = "",
        azure_key: str = "",
        azure_region: str = "",
        deepl_api_key: str = "",
        baidu_tier: str = "standard",
        google_project_id: str = "",
        google_sa_json: str = "",
        local_model_dir: str = "",
        ollama_config: Optional[dict[str, Any]] = None,
    ):
        if self._thread and self._thread.is_alive():
            raise RuntimeError("Worker already running")
        if not target_lang or not isinstance(target_lang, str) or not target_lang.strip():
            raise ValueError("target_lang must be a non-empty language code (e.g. 'en')")
        self._stop.clear()
        self._thread = threading.Thread(
            target=self._run,
            args=(
                list(files), target_lang, output_dir, translator_name,
                progress_cb, log_cb, source_lang, detector,
                engine, proxies, google_api_key, baidu_appid, baidu_appkey,
                azure_key, azure_region, deepl_api_key,
                baidu_tier, google_project_id, google_sa_json,
                local_model_dir, ollama_config,
            ),
            daemon=True,
        )
        self._thread.start()

    def stop(self):
        self._stop.set()

    def _run(self, files, target_lang, output_dir, translator_name, progress_cb, log_cb, source_lang, detector,
             engine, proxies, google_api_key, baidu_appid, baidu_appkey,
             azure_key="", azure_region="", deepl_api_key="",
             baidu_tier="standard", google_project_id="", google_sa_json="",
             local_model_dir="", ollama_config=None):
        import os

        normalized_ollama_config = _normalize_ollama_config(ollama_config)
        pdf_advisor = None
        ollama_translator = None

        def _get_ollama_translator():
            nonlocal ollama_translator
            if not normalized_ollama_config["enabled"]:
                return None
            if ollama_translator is None:
                from ..translators.ollama import OllamaSemanticTranslator, ensure_ollama_models, ollama_required_models

                ollama_base_url = normalized_ollama_config["base_url"]
                ollama_proxies = _ollama_client_proxies(ollama_base_url, proxies)
                semantic_model = normalized_ollama_config["model"]
                advisor_model = normalized_ollama_config["advisor_model"]

                ensure_ollama_models(
                    ollama_required_models(semantic_model, advisor_model=advisor_model),
                    base_url=ollama_base_url,
                    proxies=proxies,
                    status_callback=lambda message: log_cb(f"Ollama: {message}"),
                )
                ollama_translator = OllamaSemanticTranslator(
                    model=semantic_model,
                    source_lang=source_lang,
                    base_url=ollama_base_url,
                    proxies=ollama_proxies,
                )
            return ollama_translator

        def _get_pdf_ollama_components():
            nonlocal pdf_advisor
            if not normalized_ollama_config["enabled"]:
                return None, None
            translator = _get_ollama_translator()
            if translator is None:
                return None, None
            if pdf_advisor is None:
                ollama_base_url = normalized_ollama_config["base_url"]
                ollama_proxies = _ollama_client_proxies(ollama_base_url, proxies)
                advisor_model = normalized_ollama_config["advisor_model"]
                if advisor_model is None:
                    from ..advisors.null import NullAdvisor
                    pdf_advisor = NullAdvisor()
                else:
                    from ..advisors.ollama import OllamaAdvisor
                    pdf_advisor = OllamaAdvisor(
                        model=advisor_model,
                        base_url=ollama_base_url,
                        proxies=ollama_proxies,
                    )
            return pdf_advisor, translator

        # Compute file-size weights for smooth global progress
        file_sizes = []
        for f in files:
            try:
                file_sizes.append(os.path.getsize(f))
            except OSError:
                file_sizes.append(1)
        total_size = sum(file_sizes) or 1
        # cumulative_weight[i] = fraction of total work completed by files before i
        cumulative_weight = []
        cumsum = 0.0
        for sz in file_sizes:
            cumulative_weight.append(cumsum / total_size)
            cumsum += sz
        file_weight = [sz / total_size for sz in file_sizes]

        for fi, f in enumerate(files):
            if self._stop.is_set():
                log_cb("Cancelled by user")
                break
            name = Path(f).name
            t0 = time.perf_counter()

            # Build a per-file progress callback that maps (done, total)
            # within this file to a global fraction and forwards it
            base = cumulative_weight[fi]
            weight = file_weight[fi]

            def _file_progress(done: int, total: int, _base=base, _weight=weight) -> None:
                if total > 0:
                    frac = _base + _weight * (done / total)
                else:
                    frac = _base
                progress_cb(f, "progress", frac)

            try:
                progress_cb(f, "started", None)
                log_cb(f"Translating {name} ...")
                suffix = Path(f).suffix.lower()
                advisor = None
                semantic_translator = None
                primary_translator = None
                if suffix == ".pdf":
                    advisor, semantic_translator = _get_pdf_ollama_components()
                elif suffix in {".docx", ".xlsx", ".xls"} and normalized_ollama_config["supports_non_pdf"]:
                    primary_translator = _get_ollama_translator()
                out = translate_file(
                    f, target_lang, output_dir, translator_name,
                    source_lang=source_lang,
                    cancel_event=self._stop,
                    detector=detector,
                    engine=engine,
                    proxies=proxies,
                    google_api_key=google_api_key,
                    baidu_appid=baidu_appid,
                    baidu_appkey=baidu_appkey,
                    azure_key=azure_key,
                    azure_region=azure_region,
                    deepl_api_key=deepl_api_key,
                    baidu_tier=baidu_tier,
                    google_project_id=google_project_id,
                    google_sa_json=google_sa_json,
                    local_model_dir=local_model_dir,
                    progress_callback=_file_progress,
                    advisor=advisor,
                    semantic_translator=semantic_translator,
                    translator_override=primary_translator,
                )
                elapsed = time.perf_counter() - t0

                # Check cancellation right after translate_file returns
                if self._stop.is_set():
                    progress_cb(f, "cancelled", None)
                    log_cb(f"Cancelled during {name}")
                    break

                if out == "skipped-ocr":
                    progress_cb(f, "finished", elapsed)
                    log_cb(f"Skipped {name} (scanned/image PDF requiring OCR)")
                else:
                    progress_cb(f, "finished", elapsed)
                    try:
                        if out:
                            log_cb(f"Finished {name} -> {out}")
                        else:
                            log_cb(f"Finished {name}")
                    except Exception:
                        log_cb(f"Finished {name}")
            except CancelledError:
                progress_cb(f, "cancelled", None)
                log_cb(f"Cancelled during {name}")
                break
            except Exception as e:
                elapsed = time.perf_counter() - t0
                progress_cb(f, "error", elapsed)
                tb = traceback.format_exc()
                log_cb(f"Error translating {name}: {e}\n{tb}")

        # Signal that the worker loop has exited
        log_cb("__worker_done__")


class GuiLoggingHandler(logging.Handler):
    # Logging handler that forwards formatted log records to a GUI log callback.

    def __init__(self, log_cb: Callable, debug_getter: Optional[Callable] = None):
        super().__init__()
        self.log_cb = log_cb
        self._debug_getter = debug_getter or (lambda: False)

    def _sanitize_warning_text(self, text: str) -> str:
        try:
            # If the captured text contains 'UserWarning:'
            if "UserWarning:" in text:
                idx = text.find("UserWarning:")
                after = text[idx + len("UserWarning:"):]
            else:
                after = text

            # Split into lines and pick the first non-empty line
            for line in after.splitlines():
                s = line.strip()
                if s:
                    return s
            return after.strip()
        except Exception:
            return text

    def emit(self, record: logging.LogRecord) -> None:
        try:
            orig_msg = record.getMessage()

            # Skip noisy informational messages about collected cells
            if record.levelno == logging.INFO and re.search(r"collected \d+ translatable string cells", orig_msg, re.IGNORECASE):
                return

            # Determine current debug mode
            debug_enabled = False
            try:
                debug_enabled = bool(self._debug_getter())
            except Exception:
                debug_enabled = False

            if not debug_enabled and record.levelno < logging.WARNING:
                return

            # Sanitize warning messages
            sanitized = orig_msg
            if record.name == "py.warnings" or "UserWarning:" in orig_msg:
                sanitized = self._sanitize_warning_text(orig_msg)

            final = None
            if self.formatter is not None:
                try:
                    formatted = self.format(record)
                    if orig_msg:
                        final = formatted.replace(orig_msg, sanitized, 1)
                    else:
                        final = f"{formatted} {sanitized}"
                except Exception:
                    final = sanitized
            else:
                final = sanitized

            if final:
                try:
                    self.log_cb(final)
                except Exception:
                    # Swallow GUI callback errors to avoid breaking logging
                    pass
        except Exception:
            try:
                self.log_cb(f"Logging error: {record.getMessage()}")
            except Exception:
                pass
