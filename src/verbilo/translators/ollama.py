from __future__ import annotations

import json
import logging
import os
import platform
import re
import shutil
import subprocess
import threading
import time
from typing import Any, Callable, Optional
from urllib.parse import urlparse

from .cache import get_cache
from .http_session import make_session
from ..utils import CancelledError

logger = logging.getLogger(__name__)

DEFAULT_OLLAMA_MODEL = "qwen3.5:4b"
DEFAULT_OLLAMA_ADVISOR_MODEL = "qwen3.5:4b"
DEFAULT_OLLAMA_BASE_URL = "http://127.0.0.1:11434"
_LOCAL_OLLAMA_HOSTS = frozenset({"localhost", "127.0.0.1", "::1", "0.0.0.0"})
_OLLAMA_INSTALL_URL = "https://ollama.com/install.ps1"
_OLLAMA_READY_TIMEOUT = 30.0
_OLLAMA_READY_POLL_INTERVAL = 0.5
_OLLAMA_INSTALL_WAIT_TIMEOUT = 90.0
_SEMANTIC_L1_CACHE: dict[str, dict[str, dict[str, str]]] = {}
_SEMANTIC_L1_CACHE_LOCK = threading.Lock()
_SEMANTIC_BATCH_MAX_ITEMS = 8
_SEMANTIC_BATCH_MAX_CHARS = 4000
_OLLAMA_SEMANTIC_CACHE_VERSION = "v2"
_HAN_CHAR_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")

# Models that only do translation (plain-text prompt, no system message, no JSON).
# Compared case-insensitively as a prefix match.
_TRANSLATION_ONLY_MODEL_PREFIXES: frozenset[str] = frozenset({"demonbyron/hy-mt"})

# Best-effort mapping from ISO 639-1 codes to natural language names used in
# HY-MT prompt templates.  Unmapped codes fall back to the code itself.
_LANG_CODE_TO_NAME: dict[str, str] = {
    "af": "Afrikaans", "ar": "Arabic", "az": "Azerbaijani", "be": "Belarusian",
    "bg": "Bulgarian", "bn": "Bengali", "bs": "Bosnian", "ca": "Catalan",
    "cs": "Czech", "cy": "Welsh", "da": "Danish", "de": "German",
    "el": "Greek", "en": "English", "eo": "Esperanto", "es": "Spanish",
    "et": "Estonian", "eu": "Basque", "fa": "Persian", "fi": "Finnish",
    "fr": "French", "ga": "Irish", "gl": "Galician", "gu": "Gujarati",
    "he": "Hebrew", "hi": "Hindi", "hr": "Croatian", "hu": "Hungarian",
    "hy": "Armenian", "id": "Indonesian", "is": "Icelandic", "it": "Italian",
    "ja": "Japanese", "ka": "Georgian", "kk": "Kazakh", "km": "Khmer",
    "ko": "Korean", "lt": "Lithuanian", "lv": "Latvian", "mk": "Macedonian",
    "ml": "Malayalam", "mn": "Mongolian", "mr": "Marathi", "ms": "Malay",
    "mt": "Maltese", "my": "Burmese", "ne": "Nepali", "nl": "Dutch",
    "no": "Norwegian", "pa": "Punjabi", "pl": "Polish", "pt": "Portuguese",
    "ro": "Romanian", "ru": "Russian", "sk": "Slovak", "sl": "Slovenian",
    "sq": "Albanian", "sr": "Serbian", "sv": "Swedish", "sw": "Swahili",
    "ta": "Tamil", "te": "Telugu", "th": "Thai", "tl": "Filipino",
    "tr": "Turkish", "uk": "Ukrainian", "ur": "Urdu", "uz": "Uzbek",
    "vi": "Vietnamese", "xh": "Xhosa", "yo": "Yoruba", "zh": "Chinese",
    "zh-cn": "Chinese (Simplified)", "zh-tw": "Chinese (Traditional)",
    "zu": "Zulu",
}


class _OllamaInstallerBusyError(RuntimeError):
    pass


def _normalize_ollama_model_name(model: str | None, *, default: str) -> str:
    return str(model or default).strip() or default


def _is_translation_only_model(model: str) -> bool:
    """Return True if *model* is a translation-only model (e.g. HY-MT).

    Translation-only models use plain-text prompts and cannot perform
    instruction-following tasks such as block classification.
    """
    lowered = (model or "").lower()
    return any(lowered.startswith(prefix) for prefix in _TRANSLATION_ONLY_MODEL_PREFIXES)


def _has_han_text(text: str) -> bool:
    return bool(_HAN_CHAR_RE.search(text or ""))


def _uses_hymt_chinese_prompt(source_lang: str, text: str) -> bool:
    normalized = str(source_lang or "").strip().lower()
    if normalized.startswith("zh"):
        return True
    if normalized and normalized != "auto":
        return False
    return _has_han_text(text)


def resolve_ollama_pdf_models(
    model: str = DEFAULT_OLLAMA_MODEL,
    *,
    advisor_model: str | None = None,
) -> dict[str, str | None]:
    """Return the semantic and advisor model names for PDF translation.

    Returns a dict with keys ``semantic_model`` (str) and ``advisor_model``
    (str | None).  ``advisor_model`` is *None* when the selected model is
    translation-only and therefore cannot serve as an advisor.
    """
    semantic_model = _normalize_ollama_model_name(model, default=DEFAULT_OLLAMA_MODEL)
    explicit_advisor_model = str(advisor_model or "").strip()
    if explicit_advisor_model:
        resolved_advisor_model: str | None = explicit_advisor_model
    elif _is_translation_only_model(semantic_model):
        resolved_advisor_model = None
    else:
        resolved_advisor_model = semantic_model

    return {
        "semantic_model": semantic_model,
        "advisor_model": resolved_advisor_model,
    }


def ollama_pdf_required_models(
    model: str = DEFAULT_OLLAMA_MODEL,
    *,
    advisor_model: str | None = None,
) -> list[str]:
    resolved_models = resolve_ollama_pdf_models(model, advisor_model=advisor_model)
    ordered_models = [resolved_models["semantic_model"], resolved_models["advisor_model"]]
    required_models: list[str] = []
    seen: set[str] = set()
    for candidate in ordered_models:
        if candidate is None:
            continue
        lowered = candidate.lower()
        if lowered in seen:
            continue
        seen.add(lowered)
        required_models.append(candidate)
    return required_models


def _normalize_base_url(base_url: str) -> str:
    candidate = (base_url or DEFAULT_OLLAMA_BASE_URL).strip() or DEFAULT_OLLAMA_BASE_URL
    if "://" not in candidate:
        candidate = f"http://{candidate}"
    parsed = urlparse(candidate)
    netloc = parsed.netloc or parsed.path
    scheme = parsed.scheme or "http"
    if not netloc:
        return DEFAULT_OLLAMA_BASE_URL
    return f"{scheme}://{netloc}".rstrip("/")


def _is_local_ollama_url(base_url: str) -> bool:
    hostname = urlparse(_normalize_base_url(base_url)).hostname or ""
    return hostname.lower() in _LOCAL_OLLAMA_HOSTS


def _ollama_host_from_base_url(base_url: str) -> str:
    parsed = urlparse(_normalize_base_url(base_url))
    host = parsed.hostname or "127.0.0.1"
    if ":" in host and not host.startswith("["):
        host = f"[{host}]"
    port = parsed.port or 11434
    return f"{host}:{port}"


def _session_proxies_for_base_url(base_url: str, proxies: dict | None) -> dict | None:
    return None if _is_local_ollama_url(base_url) else proxies


def _is_ollama_server_reachable(
    base_url: str,
    *,
    timeout: float = 2.0,
    proxies: dict | None = None,
) -> bool:
    session = make_session(
        proxies=_session_proxies_for_base_url(base_url, proxies),
        timeout=timeout,
        retries=0,
        backoff=0.0,
    )
    try:
        response = session.get(f"{_normalize_base_url(base_url)}/api/tags")
        response.raise_for_status()
        return True
    except Exception:
        return False


def _list_ollama_models(
    base_url: str,
    *,
    timeout: float = 5.0,
    proxies: dict | None = None,
) -> set[str]:
    session = make_session(
        proxies=_session_proxies_for_base_url(base_url, proxies),
        timeout=timeout,
        retries=0,
        backoff=0.0,
    )
    response = session.get(f"{_normalize_base_url(base_url)}/api/tags")
    response.raise_for_status()
    payload = response.json()
    return {
        str(item.get("name") or "").strip().lower()
        for item in payload.get("models", [])
        if str(item.get("name") or "").strip()
    }


def _find_ollama_executable() -> str | None:
    resolved = shutil.which("ollama")
    if resolved:
        return resolved

    if platform.system().lower() != "windows":
        return None

    candidates: list[str] = []
    local_app_data = os.environ.get("LOCALAPPDATA", "")
    program_files = os.environ.get("ProgramFiles", "")
    program_files_x86 = os.environ.get("ProgramFiles(x86)", "")

    if local_app_data:
        candidates.extend(
            [
                os.path.join(local_app_data, "Programs", "Ollama", "ollama.exe"),
                os.path.join(local_app_data, "Ollama", "ollama.exe"),
            ]
        )
    if program_files:
        candidates.append(os.path.join(program_files, "Ollama", "ollama.exe"))
    if program_files_x86:
        candidates.append(os.path.join(program_files_x86, "Ollama", "ollama.exe"))

    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return candidate
    return None


def _wait_for_ollama_installation(
    base_url: str,
    *,
    timeout: float = _OLLAMA_INSTALL_WAIT_TIMEOUT,
    proxies: dict | None = None,
    status_callback: Optional[Callable[[str], None]] = None,
) -> str | None:
    if status_callback is not None:
        status_callback("Waiting for Ollama installer...")

    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        executable = _find_ollama_executable()
        if executable is not None:
            return executable
        if _is_ollama_server_reachable(base_url, timeout=2.0, proxies=proxies):
            return None
        time.sleep(_OLLAMA_READY_POLL_INTERVAL)
    return None


def _is_ollama_installer_busy_error(detail: str) -> bool:
    lowered = detail.lower()
    return "ollamasetup.exe" in lowered and "being used by another process" in lowered


def _install_ollama_windows(status_callback: Optional[Callable[[str], None]] = None) -> None:
    if status_callback is not None:
        status_callback("Installing Ollama runtime...")

    command = [
        "powershell.exe",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        f"irm {_OLLAMA_INSTALL_URL} | iex",
    ]
    try:
        completed = subprocess.run(command, capture_output=True, text=True, check=False)
    except OSError as exc:
        raise RuntimeError(f"Could not start PowerShell to install Ollama: {exc}") from exc

    if completed.returncode != 0:
        detail = completed.stderr.strip() or completed.stdout.strip() or f"exit code {completed.returncode}"
        if _is_ollama_installer_busy_error(detail):
            raise _OllamaInstallerBusyError(detail)
        raise RuntimeError(f"Ollama installation failed: {detail}")


def _build_ollama_child_env(base_url: str, proxies: dict | None) -> dict[str, str]:
    env = os.environ.copy()
    env["OLLAMA_HOST"] = _ollama_host_from_base_url(base_url)

    https_proxy = None
    if proxies:
        https_proxy = proxies.get("https") or proxies.get("http")
    if https_proxy:
        env["HTTPS_PROXY"] = https_proxy
        env["https_proxy"] = https_proxy

    env.pop("HTTP_PROXY", None)
    env.pop("http_proxy", None)
    return env


def _start_ollama_server(base_url: str, *, executable: str, proxies: dict | None = None) -> None:
    creationflags = 0
    if platform.system().lower() == "windows":
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

    subprocess.Popen(
        [executable, "serve"],
        env=_build_ollama_child_env(base_url, proxies),
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=creationflags,
        start_new_session=True,
    )


def _wait_for_ollama_server(
    base_url: str,
    *,
    timeout: float = _OLLAMA_READY_TIMEOUT,
    proxies: dict | None = None,
) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if _is_ollama_server_reachable(base_url, timeout=2.0, proxies=proxies):
            return True
        time.sleep(_OLLAMA_READY_POLL_INTERVAL)
    return False


def _ensure_ollama_server(
    base_url: str = DEFAULT_OLLAMA_BASE_URL,
    *,
    proxies: dict | None = None,
    status_callback: Optional[Callable[[str], None]] = None,
) -> str:
    root_url = _normalize_base_url(base_url)
    if _is_ollama_server_reachable(root_url, timeout=2.0, proxies=proxies):
        return root_url

    if not _is_local_ollama_url(root_url):
        raise RuntimeError(
            f"Cannot reach Ollama at {root_url}. Automatic install/start is only supported for local URLs."
        )

    executable = _find_ollama_executable()
    if executable is None:
        if platform.system().lower() != "windows":
            raise RuntimeError(
                f"Cannot reach Ollama at {root_url}, and automatic installation is only supported on Windows."
            )
        try:
            _install_ollama_windows(status_callback=status_callback)
        except _OllamaInstallerBusyError:
            if status_callback is not None:
                status_callback("Ollama installer already running; waiting for it to finish...")
        executable = _wait_for_ollama_installation(
            root_url,
            timeout=_OLLAMA_INSTALL_WAIT_TIMEOUT,
            proxies=proxies,
            status_callback=status_callback,
        )
        if executable is None and not _is_ollama_server_reachable(root_url, timeout=2.0, proxies=proxies):
            raise RuntimeError(
                "Ollama installation is still in progress or did not finish in time. Please wait a moment and try again."
            )
        if _wait_for_ollama_server(root_url, timeout=5.0, proxies=proxies):
            return root_url

    if _wait_for_ollama_server(root_url, timeout=2.0, proxies=proxies):
        return root_url

    if status_callback is not None:
        status_callback("Starting Ollama runtime...")
    try:
        _start_ollama_server(root_url, executable=executable, proxies=proxies)
    except OSError as exc:
        raise RuntimeError(f"Could not start Ollama from {executable}: {exc}") from exc

    if status_callback is not None:
        status_callback("Waiting for Ollama runtime...")
    if not _wait_for_ollama_server(root_url, timeout=_OLLAMA_READY_TIMEOUT, proxies=proxies):
        raise RuntimeError(f"Could not reach Ollama at {root_url} after starting it.")
    return root_url


def _pull_ollama_model_via_api(
    model: str,
    *,
    base_url: str,
    timeout: float,
    proxies: dict | None,
    status_callback: Optional[Callable[[str], None]],
) -> dict[str, Any]:
    model_name = (model or DEFAULT_OLLAMA_MODEL).strip() or DEFAULT_OLLAMA_MODEL
    root_url = _normalize_base_url(base_url)

    session = make_session(
        proxies=_session_proxies_for_base_url(root_url, proxies),
        timeout=timeout,
    )
    with session.post(
        f"{root_url}/api/pull",
        json={"name": model_name, "stream": True},
        stream=True,
    ) as response:
        response.raise_for_status()
        last_payload: dict[str, Any] = {"status": "started"}
        for raw_line in response.iter_lines(decode_unicode=True):
            if not raw_line:
                continue
            try:
                payload = json.loads(raw_line)
            except json.JSONDecodeError:
                logger.debug("Ignoring non-JSON Ollama pull event: %s", raw_line)
                continue
            last_payload = payload
            if status_callback is not None:
                status_callback(_format_pull_status(model_name, payload))
    return last_payload


def ensure_ollama_model(
    model: str = DEFAULT_OLLAMA_MODEL,
    *,
    base_url: str = DEFAULT_OLLAMA_BASE_URL,
    timeout: float = 600.0,
    proxies: dict | None = None,
    status_callback: Optional[Callable[[str], None]] = None,
) -> str:
    model_name = _normalize_ollama_model_name(model, default=DEFAULT_OLLAMA_MODEL)
    return ensure_ollama_models(
        [model_name],
        base_url=base_url,
        timeout=timeout,
        proxies=proxies,
        status_callback=status_callback,
    )


def ensure_ollama_models(
    models: list[str] | tuple[str, ...],
    *,
    base_url: str = DEFAULT_OLLAMA_BASE_URL,
    timeout: float = 600.0,
    proxies: dict | None = None,
    status_callback: Optional[Callable[[str], None]] = None,
) -> str:
    root_url = _ensure_ollama_server(
        base_url,
        proxies=proxies,
        status_callback=status_callback,
    )
    available_models = _list_ollama_models(root_url, proxies=proxies)
    for model_name in [_normalize_ollama_model_name(model, default=DEFAULT_OLLAMA_MODEL) for model in models]:
        if model_name.lower() in available_models:
            continue
        if status_callback is not None:
            status_callback(f"Pulling {model_name}...")
        _pull_ollama_model_via_api(
            model_name,
            base_url=root_url,
            timeout=timeout,
            proxies=proxies,
            status_callback=status_callback,
        )
        available_models.add(model_name.lower())
    return root_url


def _run_cancellable(
    fn,
    cancel_event: Optional[threading.Event],
    poll_interval: float = 0.05,
):
    result = [None]
    exc: list[BaseException | None] = [None]

    def _target() -> None:
        try:
            result[0] = fn()
        except BaseException as err:  # pragma: no cover - passthrough wrapper
            exc[0] = err

    thread = threading.Thread(target=_target, daemon=True)
    thread.start()
    while thread.is_alive():
        thread.join(timeout=poll_interval)
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError("Translation cancelled")
    if exc[0] is not None:
        raise exc[0]
    return result[0]


def _format_pull_status(model: str, payload: dict[str, Any]) -> str:
    status = str(payload.get("status", "") or "").strip()
    completed = payload.get("completed")
    total = payload.get("total")

    if status.lower() == "success":
        return f"Pulled {model}"

    if isinstance(completed, int) and isinstance(total, int) and total > 0:
        pct = max(0, min(100, round((completed / total) * 100)))
        if status:
            return f"{status} ({pct}%)"
        return f"Pulling {model} ({pct}%)"

    return status or f"Pulling {model}"


def pull_ollama_model(
    model: str = DEFAULT_OLLAMA_MODEL,
    *,
    base_url: str = DEFAULT_OLLAMA_BASE_URL,
    timeout: float = 600.0,
    proxies: dict | None = None,
    status_callback: Optional[Callable[[str], None]] = None,
) -> dict[str, Any]:
    model_name = _normalize_ollama_model_name(model, default=DEFAULT_OLLAMA_MODEL)
    result = pull_ollama_models(
        [model_name],
        base_url=base_url,
        timeout=timeout,
        proxies=proxies,
        status_callback=status_callback,
    )
    return result[model_name]


def pull_ollama_models(
    models: list[str] | tuple[str, ...],
    *,
    base_url: str = DEFAULT_OLLAMA_BASE_URL,
    timeout: float = 600.0,
    proxies: dict | None = None,
    status_callback: Optional[Callable[[str], None]] = None,
) -> dict[str, dict[str, Any]]:
    root_url = _normalize_base_url(base_url)

    if status_callback is not None:
        status_callback(f"Connecting to {root_url}")

    _ensure_ollama_server(root_url, proxies=proxies, status_callback=status_callback)
    results: dict[str, dict[str, Any]] = {}
    for model_name in [_normalize_ollama_model_name(model, default=DEFAULT_OLLAMA_MODEL) for model in models]:
        results[model_name] = _pull_ollama_model_via_api(
            model_name,
            base_url=root_url,
            timeout=timeout,
            proxies=proxies,
            status_callback=status_callback,
        )
    return results


def remove_ollama_model(
    model: str,
    *,
    base_url: str = DEFAULT_OLLAMA_BASE_URL,
    proxies: dict | None = None,
) -> None:
    """Delete a model from the local Ollama library via the DELETE /api/delete endpoint."""
    model_name = _normalize_ollama_model_name(model, default=DEFAULT_OLLAMA_MODEL)
    root_url = _normalize_base_url(base_url)
    session = make_session(proxies=proxies, timeout=30.0)
    response = session.delete(f"{root_url}/api/delete", json={"model": model_name})
    response.raise_for_status()


def check_ollama_model_available(
    model: str,
    base_url: str = DEFAULT_OLLAMA_BASE_URL,
    *,
    proxies: dict | None = None,
    timeout: float = 3.0,
) -> bool:
    """Return True if *model* is downloaded and available in the local Ollama library.

    Uses POST /api/show.  Returns False for 404 responses, server-not-reachable,
    or any other error.  Never raises.
    """
    try:
        root_url = _normalize_base_url(base_url)
        session = make_session(
            proxies=_session_proxies_for_base_url(root_url, proxies),
            timeout=timeout,
            retries=0,
            backoff=0.0,
        )
        response = session.post(f"{root_url}/api/show", json={"model": model})
        return response.status_code == 200
    except Exception:
        return False


class OllamaSemanticTranslator:
    def __init__(
        self,
        model: str = DEFAULT_OLLAMA_MODEL,
        source_lang: str = "auto",
        timeout: float = 30.0,
        base_url: str = DEFAULT_OLLAMA_BASE_URL,
        proxies: dict | None = None,
    ) -> None:
        self._model = model
        self._source_lang = source_lang or "auto"
        self._engine_name = f"ollama-semantic:{model.lower()}:{_OLLAMA_SEMANTIC_CACHE_VERSION}"
        self._base_url = _normalize_base_url(base_url)
        self._session = make_session(
            proxies=_session_proxies_for_base_url(self._base_url, proxies),
            timeout=timeout,
        )
        self._translate_all_blocks: bool = _is_translation_only_model(model)
        with _SEMANTIC_L1_CACHE_LOCK:
            self._cache = _SEMANTIC_L1_CACHE.setdefault(self._engine_name, {})

    def translate_text(self, text: str, target_lang: str) -> str | None:
        return self._translate_one(
            text=text,
            target_lang=target_lang,
            source_lang=self._source_lang,
            content_hint="body",
            strategy="semantic",
            cancel_event=None,
        )

    def translate_batch(
        self,
        texts: list[str],
        target_lang: str,
        *,
        cancel_event: Optional[threading.Event] = None,
    ) -> list[str | None]:
        blocks = [
            {
                "text": text,
                "content_hint": "body",
                "strategy": "semantic",
            }
            for text in texts
        ]
        return self.translate_blocks(
            blocks,
            target_lang,
            cancel_event=cancel_event,
        )

    def translate_blocks(
        self,
        blocks: list[dict[str, Any]],
        target_lang: str,
        *,
        cancel_event: Optional[threading.Event] = None,
    ) -> list[str | None]:
        if not blocks:
            return []

        results: list[str | None] = []
        for block in blocks:
            text = str(block.get("text", ""))
            results.append(text if not text.strip() else None)
        target_cache = self._cache.setdefault(target_lang, {})
        pending: dict[str, dict[str, Any]] = {}
        metrics = {
            "blocks": len(blocks),
            "l1_hits": 0,
            "l2_hits": 0,
            "llm_calls": 0,
            "batch_requests": 0,
            "fallback_items": 0,
        }

        for idx, block in enumerate(blocks):
            if cancel_event is not None and cancel_event.is_set():
                raise CancelledError("Translation cancelled")

            text = str(block.get("text", ""))
            if not text.strip():
                continue

            source_lang = str(block.get("source_lang", self._source_lang) or self._source_lang)
            content_hint = str(block.get("content_hint", "body") or "body")
            strategy = str(block.get("strategy", "semantic") or "semantic")
            capacity_chars = max(int(block.get("capacity_chars") or 0), 0)
            source_visible_chars = len(" ".join(text.split()))
            line_count = max(1, len(text.split("\n")))
            cache_source = self._cache_source_key(
                text=text,
                source_lang=source_lang,
                content_hint=content_hint,
                strategy=strategy,
            )

            cached = target_cache.get(cache_source)
            if cached is not None:
                if self._translation_only_failure_reason(
                    source_text=text,
                    translated_text=cached,
                    raw_translated_text=cached,
                    source_lang=source_lang,
                ) is None:
                    metrics["l1_hits"] += 1
                    results[idx] = cached
                    continue
                target_cache.pop(cache_source, None)

            item = pending.setdefault(
                cache_source,
                {
                    "indices": [],
                    "text": text,
                    "source_lang": source_lang,
                    "content_hint": content_hint,
                    "strategy": strategy,
                    "capacity_chars": capacity_chars,
                    "source_visible_chars": source_visible_chars,
                    "line_count": line_count,
                },
            )
            item["indices"].append(idx)

        if not pending:
            return results

        l2_hits = get_cache().get_batch(
            self._engine_cache_key(),
            list(pending.keys()),
            target_lang,
        )
        for cache_source, translated in l2_hits.items():
            item = pending.get(cache_source)
            if item is None:
                continue
            if self._translation_only_failure_reason(
                source_text=str(item.get("text", "")),
                translated_text=translated,
                raw_translated_text=translated,
                source_lang=str(item.get("source_lang", self._source_lang) or self._source_lang),
            ) is not None:
                continue

            metrics["l2_hits"] += 1
            target_cache[cache_source] = translated
            pending.pop(cache_source, None)
            for idx in item["indices"]:
                results[idx] = translated

        cache_pairs: list[tuple[str, str]] = []
        for cache_source, item in pending.items():
            if cancel_event is not None and cancel_event.is_set():
                raise CancelledError("Translation cancelled")

            metrics["llm_calls"] += 1
            translated = self._translate_one(
                text=item["text"],
                target_lang=target_lang,
                source_lang=item["source_lang"],
                content_hint=item["content_hint"],
                strategy=item["strategy"],
                capacity_chars=item.get("capacity_chars", 0),
                source_visible_chars=item.get("source_visible_chars", 0),
                line_count=item.get("line_count", 1),
                cancel_event=cancel_event,
            )

            if translated is None:
                metrics["fallback_items"] += len(item["indices"])
                continue

            target_cache[cache_source] = translated
            for idx in item["indices"]:
                results[idx] = translated
            cache_pairs.append((cache_source, translated))

        if cache_pairs:
            get_cache().put_batch(self._engine_cache_key(), cache_pairs, target_lang)

        logger.info(
            "Semantic translation metrics: blocks=%s l1_hits=%s l2_hits=%s llm_calls=%s batch_requests=%s fallback_items=%s",
            metrics["blocks"],
            metrics["l1_hits"],
            metrics["l2_hits"],
            metrics["llm_calls"],
            metrics["batch_requests"],
            metrics["fallback_items"],
        )

        return results

    def _iter_semantic_batches(
        self,
        pending_items: list[tuple[str, dict[str, Any]]],
    ):
        batch: list[tuple[str, dict[str, Any]]] = []
        batch_chars = 0

        for cache_source, item in pending_items:
            item_chars = len(str(item.get("text", "")))
            if batch and (
                len(batch) >= _SEMANTIC_BATCH_MAX_ITEMS
                or batch_chars + item_chars > _SEMANTIC_BATCH_MAX_CHARS
            ):
                yield batch
                batch = []
                batch_chars = 0

            batch.append((cache_source, item))
            batch_chars += item_chars

        if batch:
            yield batch

    def _translate_many(
        self,
        *,
        items: list[dict[str, Any]],
        target_lang: str,
        cancel_event: Optional[threading.Event],
    ) -> dict[int, str]:
        if not items:
            return {}

        system_prompt, user_prompt = self._build_batch_prompt(items=items, target_lang=target_lang)
        payload = _run_cancellable(
            lambda: self._request_translation_batch(system_prompt, user_prompt),
            cancel_event,
        )
        parsed = self._parse_batch_translation_payload(payload)

        translations: dict[int, str] = {}
        for idx, item in enumerate(items):
            translated_text = parsed.get(idx)
            if translated_text is None:
                continue
            translations[idx] = self._post_process_semantic_translation(
                str(item.get("text", "")),
                translated_text,
            )
        return translations

    def _translate_one(
        self,
        *,
        text: str,
        target_lang: str,
        source_lang: str,
        content_hint: str,
        strategy: str,
        capacity_chars: int = 0,
        source_visible_chars: int = 0,
        line_count: int = 1,
        cancel_event: Optional[threading.Event],
    ) -> str | None:
        if not text or not text.strip():
            return text

        translation_only = _is_translation_only_model(self._model)
        if translation_only:
            user_prompt = self._build_hymt_prompt(text=text, source_lang=source_lang, target_lang=target_lang)
            translated = _run_cancellable(
                lambda: self._request_hymt_translation(user_prompt),
                cancel_event,
            )
        else:
            system_prompt, user_prompt = self._build_prompt(
                text=text,
                source_lang=source_lang,
                target_lang=target_lang,
                content_hint=content_hint,
                strategy=strategy,
                capacity_chars=capacity_chars,
                source_visible_chars=source_visible_chars,
                line_count=line_count,
            )
            translated = _run_cancellable(
                lambda: self._request_translation(system_prompt, user_prompt),
                cancel_event,
            )

        raw_translated = str(translated or "")
        normalized = self._post_process_semantic_translation(text, raw_translated)
        failure_reason = self._translation_only_failure_reason(
            source_text=text,
            translated_text=normalized,
            raw_translated_text=raw_translated,
            source_lang=source_lang,
        )
        if failure_reason is not None:
            logger.debug(
                "HY-MT semantic block rejected reason=%s source_lang=%s source_len=%d",
                failure_reason,
                source_lang,
                len(text),
            )
            return None
        return normalized

    def _semantic_system_prompt(self) -> str:
        return (
            "You are a translation engine for layout-constrained documents. "
            "Output ONLY the translated text, nothing else — no explanations, no notes, no commentary, no alternatives. "
            "Preserve meaning, tone, and document role. Preserve meaningful line breaks. "
            "Chinese source texts are compact; their translations must be equally compact. "
            "Do not add qualifiers, articles, connectives, or words not implied by the source. "
            "A short source phrase must yield a short target phrase, never a full sentence. "
            "For labels, headings, numbers, or wording-sensitive text, stay close to the source wording. "
            "If several translations are valid, always choose the shortest natural wording that preserves the source meaning. "
            "When the payload includes capacity_chars and source_visible_chars, treat capacity_chars as the available character budget "
            "for the translation. If source_visible_chars is near or above capacity_chars, use the tightest natural wording possible "
            "without dropping required meaning."
        )

    def _build_hymt_prompt(self, *, text: str, source_lang: str, target_lang: str) -> str:
        """Build a plain-text prompt for HY-MT translation models.

        ZH source uses the Chinese instruction template; everything else uses
        the English instruction template.  No system message is used.
        """
        lang_lower = target_lang.lower()
        lang_name = _LANG_CODE_TO_NAME.get(lang_lower, target_lang)
        if _uses_hymt_chinese_prompt(source_lang, text):
            return (
                f"将以下文本翻译为{lang_name}，注意只需要输出翻译后的结果，不要额外解释：\n{text}"
            )
        return (
            f"Translate the following segment into {lang_name}, without additional explanation.\n{text}"
        )

    def _request_hymt_translation(self, user_prompt: str) -> str:
        """Send a plain-text user message to a HY-MT model and return the translated text.

        No system message is included — HY-MT has no default system prompt
        and sending an empty system message corrupts its chat template.
        """
        response = self._session.post(
            f"{self._base_url}/api/chat",
            json={
                "model": self._model,
                "stream": False,
                "think": False,
                "options": {
                    "temperature": 0.2,
                    "top_p": 0.8,
                    "top_k": 20,
                    "presence_penalty": 1.0,
                },
                "messages": [
                    {"role": "user", "content": user_prompt},
                ],
            },
        )
        response.raise_for_status()
        payload = response.json()
        content = re.sub(r"<think>.*?</think>", "", str(payload.get("message", {}).get("content", "")), flags=re.DOTALL).strip()
        if not content:
            logger.warning("HY-MT returned empty content for block; semantic fallback will be used")
        return content

    def _request_translation(self, system_prompt: str, user_prompt: str) -> str:
        response = self._session.post(
            f"{self._base_url}/api/chat",
            json={
                "model": self._model,
                "stream": False,
                "think": False,
                "options": {
                    "temperature": 0.2,
                    "top_p": 0.8,
                    "top_k": 20,
                    "presence_penalty": 1.0,
                },
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
            },
        )
        response.raise_for_status()
        payload = response.json()
        content = re.sub(r"<think>.*?</think>", "", str(payload.get("message", {}).get("content", "")), flags=re.DOTALL).strip()
        if not content:
            logger.warning("Ollama translator returned empty content for block; source text will be kept")
        return content

    def _request_translation_batch(self, system_prompt: str, user_prompt: str) -> Any:
        response = self._session.post(
            f"{self._base_url}/api/chat",
            json={
                "model": self._model,
                "stream": False,
                "think": False,
                "format": "json",
                "options": {
                    "temperature": 0.2,
                    "top_p": 0.8,
                    "top_k": 20,
                    "presence_penalty": 1.0,
                },
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
            },
        )
        response.raise_for_status()
        payload = response.json()
        content = payload.get("message", {}).get("content", "")
        if isinstance(content, str):
            return json.loads(content)
        return content

    def _build_batch_prompt(
        self,
        *,
        items: list[dict[str, Any]],
        target_lang: str,
    ) -> tuple[str, str]:
        system_prompt = (
            self._semantic_system_prompt()
            + " Translate every item in the provided array. Return JSON only using the shape "
            + '{"translations":[{"id":0,"translation":"..."}]}. '
            + "Return exactly one entry for every input id, keep ids unchanged, and preserve meaningful line breaks inside each translation string."
        )
        user_prompt = json.dumps(
            {
                "target_lang": target_lang,
                "items": [
                    {
                        "id": idx,
                        "source_lang": str(item.get("source_lang", self._source_lang) or self._source_lang),
                        "content_hint": str(item.get("content_hint", "body") or "body"),
                        "strategy": str(item.get("strategy", "semantic") or "semantic"),
                        "source_text": str(item.get("text", "")),
                    }
                    for idx, item in enumerate(items)
                ],
            },
            ensure_ascii=False,
        )
        return system_prompt, user_prompt

    def _build_prompt(
        self,
        *,
        text: str,
        source_lang: str,
        target_lang: str,
        content_hint: str,
        strategy: str,
        capacity_chars: int = 0,
        source_visible_chars: int = 0,
        line_count: int = 1,
    ) -> tuple[str, str]:
        system_prompt = self._semantic_system_prompt()
        payload: dict[str, Any] = {
            "source_lang": source_lang,
            "target_lang": target_lang,
            "content_hint": content_hint,
            "strategy": strategy,
            "source_text": text,
        }
        if capacity_chars > 0:
            payload["capacity_chars"] = capacity_chars
            payload["source_visible_chars"] = source_visible_chars
            payload["line_count"] = line_count
        user_prompt = json.dumps(payload, ensure_ascii=False)
        return system_prompt, user_prompt

    def _post_process_semantic_translation(
        self,
        source_text: str,
        translated_text: str,
    ) -> str:
        leading = source_text[: len(source_text) - len(source_text.lstrip())]
        trailing = source_text[len(source_text.rstrip()) :]
        inner = translated_text.strip()

        if inner.startswith("```"):
            lines = [line for line in inner.splitlines() if line.strip()]
            if len(lines) >= 3 and lines[0].startswith("```") and lines[-1].startswith("```"):
                inner = "\n".join(lines[1:-1]).strip()

        if len(inner) >= 2 and inner[0] == inner[-1] and inner[0] in {'"', "'"}:
            inner = inner[1:-1].strip()

        if not inner:
            return source_text
        return leading + inner + trailing

    def _translation_only_failure_reason(
        self,
        *,
        source_text: str,
        translated_text: str,
        raw_translated_text: str,
        source_lang: str,
    ) -> str | None:
        if not _is_translation_only_model(self._model):
            return None

        visible_raw = " ".join(str(raw_translated_text or "").split())
        visible_translated = " ".join(str(translated_text or "").split())
        if not visible_raw or not visible_translated:
            return "empty"

        visible_source = " ".join(str(source_text or "").split())
        if visible_source and visible_translated == visible_source and _uses_hymt_chinese_prompt(source_lang, source_text):
            return "unchanged-source"
        return None

    def _parse_batch_translation_payload(self, payload: Any) -> dict[int, str]:
        if isinstance(payload, list):
            raw_items = payload
        elif isinstance(payload, dict):
            raw_items = payload.get("translations") or payload.get("items") or payload.get("results")
        else:
            raise ValueError("Unexpected semantic batch payload type")

        if not isinstance(raw_items, list):
            raise ValueError("Semantic batch payload did not contain a translations list")

        parsed: dict[int, str] = {}
        for raw_item in raw_items:
            if not isinstance(raw_item, dict):
                continue

            raw_id = raw_item.get("id", raw_item.get("index"))
            try:
                item_id = int(raw_id)
            except (TypeError, ValueError):
                continue

            translated_text = raw_item.get("translation", raw_item.get("translated_text"))
            if translated_text is None:
                continue

            translated = str(translated_text).strip()
            if not translated:
                continue
            parsed[item_id] = translated

        return parsed

    def _cache_source_key(
        self,
        *,
        text: str,
        source_lang: str,
        content_hint: str,
        strategy: str,
    ) -> str:
        return json.dumps(
            {
                "text": text,
                "source_lang": source_lang,
                "content_hint": content_hint,
                "strategy": strategy,
            },
            ensure_ascii=False,
            sort_keys=True,
        )

    def _engine_cache_key(self) -> str:
        return self._engine_name