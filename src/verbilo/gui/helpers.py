import threading
import time
import logging
import re
from pathlib import Path
from typing import Callable, Iterable, Union, Optional, Any, Mapping
import tkinter as tk
import traceback
from urllib.parse import urlparse

from ..main import translate_file
from ..progress import ProgressUpdate
from ..utils import CancelledError
from .config import redact_sensitive_text

SUPPORTED_EXTS = (".docx", ".pdf", ".xlsx", ".pptx", ".txt", ".md", ".markdown")


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

def center_window(window, width=None, height=None, parent=None, *, margin: float = 0.94):
    """Center a Tk/CustomTkinter window using its rendered size.

    Requested dimensions are also bounded to the visible virtual desktop.  This
    keeps large dialogs usable on small/high-DPI displays while preserving the
    existing parent-relative centring and negative multi-monitor coordinates.
    """
    requested_width = int(width) if width is not None else None
    requested_height = int(height) if height is not None else None
    if requested_width is not None and requested_height is not None:
        window.geometry(f"{requested_width}x{requested_height}")

    window.update_idletasks()

    # Bound an explicitly-sized dialog after CTk has applied its own scaling.
    # Scaling based on the *measured* rendered size avoids guessing the current
    # CTk/Tk DPI factor.
    if requested_width is not None and requested_height is not None:
        try:
            vw = int(window.winfo_vrootwidth())
            vh = int(window.winfo_vrootheight())
            if vw <= 1 or vh <= 1:
                raise ValueError
        except Exception:
            vw = int(window.winfo_screenwidth())
            vh = int(window.winfo_screenheight())
        max_w = max(320, int(vw * float(margin)))
        max_h = max(240, int(vh * float(margin)))
        rendered_w = max(int(window.winfo_width()), int(window.winfo_reqwidth()), 1)
        rendered_h = max(int(window.winfo_height()), int(window.winfo_reqheight()), 1)
        shrink = min(1.0, max_w / rendered_w, max_h / rendered_h)
        if shrink < 0.999:
            requested_width = max(320, int(requested_width * shrink))
            requested_height = max(240, int(requested_height * shrink))
            window.geometry(f"{requested_width}x{requested_height}")
            window.update_idletasks()

    win_w = max(int(window.winfo_width()), int(window.winfo_reqwidth()), 1)
    win_h = max(int(window.winfo_height()), int(window.winfo_reqheight()), 1)

    centered_on_parent = False
    if parent is not None:
        try:
            parent.update_idletasks()
            parent_w = max(int(parent.winfo_width()), 1)
            parent_h = max(int(parent.winfo_height()), 1)
            x = int(parent.winfo_rootx()) + (parent_w - win_w) // 2
            y = int(parent.winfo_rooty()) + (parent_h - win_h) // 2
            centered_on_parent = True
        except Exception:
            parent = None

    if parent is None:
        try:
            vx = int(window.winfo_vrootx())
            vy = int(window.winfo_vrooty())
            vw = int(window.winfo_vrootwidth())
            vh = int(window.winfo_vrootheight())
            if vw <= 1 or vh <= 1:
                raise ValueError
        except Exception:
            vx = vy = 0
            vw = int(window.winfo_screenwidth())
            vh = int(window.winfo_screenheight())
        x = vx + (vw - win_w) // 2
        y = vy + (vh - win_h) // 2

    # Preserve parent-relative negative coordinates on multi-monitor desktops.
    # For screen-centred windows clamp to the available virtual desktop so the
    # title bar and resize handles remain reachable.
    if not centered_on_parent:
        try:
            vx = int(window.winfo_vrootx())
            vy = int(window.winfo_vrooty())
            vw = int(window.winfo_vrootwidth())
            vh = int(window.winfo_vrootheight())
            if vw > 1 and vh > 1:
                x = min(max(x, vx), max(vx, vx + vw - win_w))
                y = min(max(y, vy), max(vy, vy + vh - win_h))
        except Exception:
            pass

    window.geometry(f"+{int(x)}+{int(y)}")
    window.update_idletasks()


# runs translation in a background thread; call start() to begin, stop() to cancel
class Worker:

    def __init__(self):
        self._thread: Optional[threading.Thread] = None
        self._stop = threading.Event()
        self._last_run_cancelled = False

    @property
    def cancelled(self) -> bool:
        return self._stop.is_set()

    @property
    def alive(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    @property
    def last_run_cancelled(self) -> bool:
        return bool(self._last_run_cancelled)

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
        terminology: Mapping[str, str] | None = None,
        translation_memory_enabled: bool = False,
        translation_memory_path: str | None = None,
        report_cb: Callable[[dict[str, Any]], None] | None = None,
        terminology_conflicts: Iterable[str] = (),
        terminology_entries: int = 0,
        attempt_numbers: Mapping[str, int] | None = None,
    ):
        if self._thread and self._thread.is_alive():
            raise RuntimeError("Worker already running")
        if not target_lang or not isinstance(target_lang, str) or not target_lang.strip():
            raise ValueError("target_lang must be a non-empty language code (e.g. 'en')")
        self._stop.clear()
        self._last_run_cancelled = False
        self._thread = threading.Thread(
            target=self._run,
            args=(
                list(files), target_lang, output_dir, translator_name,
                progress_cb, log_cb, source_lang, detector,
                engine, proxies, google_api_key, baidu_appid, baidu_appkey,
                azure_key, azure_region, deepl_api_key,
                baidu_tier, google_project_id, google_sa_json,
                local_model_dir, ollama_config, dict(terminology or {}),
                bool(translation_memory_enabled), translation_memory_path, report_cb,
                tuple(str(item) for item in terminology_conflicts), int(terminology_entries or 0),
                dict(attempt_numbers or {}),
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
             local_model_dir="", ollama_config=None, terminology=None,
             translation_memory_enabled=False, translation_memory_path=None, report_cb=None,
             terminology_conflicts=(), terminology_entries=0, attempt_numbers=None):
        import os

        normalized_ollama_config = _normalize_ollama_config(ollama_config)
        from ..reporting import TranslationBatchReport, TranslationFileReport
        batch_report = TranslationBatchReport(
            terminology_conflicts=tuple(terminology_conflicts or ()),
            terminology_entries=int(terminology_entries or 0),
        )
        attempt_numbers = dict(attempt_numbers or {})
        reports_by_path: dict[str, TranslationFileReport] = {}
        for path in files:
            report = TranslationFileReport(
                path=str(path),
                status="pending",
                attempt=max(1, int(attempt_numbers.get(str(path), 1) or 1)),
            )
            batch_report.files.append(report)
            reports_by_path[str(path)] = report

        pdf_advisor = None
        ollama_translator = None
        batch_cancelled = False

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

        # Batch progress is based on completed documents, not source byte size.
        # Byte weighting made image-heavy Office/PDF files look like much more
        # translation work than text-heavy smaller files.  Each current file now
        # contributes its real stage-aware document progress to one equal share
        # of the batch bar.
        file_count = max(len(files), 1)

        for fi, f in enumerate(files):
            if self._stop.is_set():
                batch_cancelled = True
                log_cb("Cancelled by user")
                break
            name = Path(f).name
            t0 = time.perf_counter()
            file_report = reports_by_path[str(f)]
            file_report.status = "started"

            def _capture_metrics(metrics, _report=file_report):
                _report.add_metrics(metrics)

            base = fi / file_count
            weight = 1.0 / file_count
            last_stage: list[str | None] = [None]
            last_forwarded: list[float] = [base]

            def _file_progress_event(
                update: ProgressUpdate,
                _base=base,
                _weight=weight,
                _name=name,
            ) -> None:
                frac = _base + _weight * update.overall_fraction
                frac = min(max(frac, 0.0), 1.0)
                stage_changed = update.stage != last_stage[0]
                # Avoid flooding Tk with one UI update per tiny text unit while
                # preserving real work-based progress. Always forward stage
                # boundaries and completion; otherwise sample at 0.2% steps.
                if stage_changed or frac >= 1.0 or frac - last_forwarded[0] >= 0.002:
                    progress_cb(f, "progress", frac)
                    last_forwarded[0] = frac
                if stage_changed:
                    last_stage[0] = update.stage
                    detail = f" — {update.detail}" if update.detail else ""
                    log_cb(f"{_name}: {update.label}{detail}")

            try:
                progress_cb(f, "started", None)
                progress_cb(f, "progress", base)
                log_cb(f"Translating {name} ...")
                suffix = Path(f).suffix.lower()
                advisor = None
                semantic_translator = None
                primary_translator = None
                if suffix == ".pdf":
                    advisor, semantic_translator = _get_pdf_ollama_components()
                elif suffix in {".docx", ".xlsx", ".pptx", ".txt", ".md", ".markdown"} and normalized_ollama_config["supports_non_pdf"]:
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
                    progress_callback=None,
                    progress_event_callback=_file_progress_event,
                    advisor=advisor,
                    semantic_translator=semantic_translator,
                    translator_override=primary_translator,
                    terminology=dict(terminology or {}),
                    translation_memory=bool(translation_memory_enabled),
                    translation_memory_path=translation_memory_path or None,
                    metrics_callback=_capture_metrics,
                )
                elapsed = time.perf_counter() - t0
                file_report.elapsed_seconds = elapsed
                file_report.output_path = None if out in (None, "skipped-ocr") else str(out)

                # Check cancellation right after translate_file returns
                if self._stop.is_set():
                    batch_cancelled = True
                    file_report.status = "cancelled"
                    progress_cb(f, "cancelled", None)
                    log_cb(f"Cancelled during {name}")
                    break

                if out == "skipped-ocr":
                    file_report.status = "skipped"
                    progress_cb(f, "finished", elapsed)
                    log_cb(f"Skipped {name} (scanned/image PDF requiring OCR)")
                else:
                    file_report.status = "finished"
                    progress_cb(f, "progress", base + weight)
                    progress_cb(f, "finished", elapsed)
                    try:
                        if out:
                            log_cb(f"Finished {name} -> {out}")
                        else:
                            log_cb(f"Finished {name}")
                    except Exception:
                        log_cb(f"Finished {name}")
            except CancelledError:
                batch_cancelled = True
                file_report.status = "cancelled"
                file_report.elapsed_seconds = time.perf_counter() - t0
                progress_cb(f, "cancelled", None)
                log_cb(f"Cancelled during {name}")
                break
            except Exception as e:
                elapsed = time.perf_counter() - t0
                file_report.status = "error"
                file_report.elapsed_seconds = elapsed
                file_report.error = str(e)
                progress_cb(f, "error", elapsed)
                tb = traceback.format_exc()
                log_cb(redact_sensitive_text(f"Error translating {name}: {e}\n{tb}"))

        if batch_cancelled or self._stop.is_set():
            for pending_report in batch_report.files:
                if pending_report.status == "pending":
                    pending_report.status = "cancelled"

        self._last_run_cancelled = bool(batch_cancelled or self._stop.is_set())

        if report_cb is not None:
            try:
                report_cb(batch_report.to_dict())
            except Exception:
                logging.debug("GUI report callback failed", exc_info=True)

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
            orig_msg = redact_sensitive_text(record.getMessage())

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
