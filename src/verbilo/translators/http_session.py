# Resilient HTTP session factory — retry, backoff, timeout, proxy support

from __future__ import annotations

import os
import logging
from typing import Optional

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

logger = logging.getLogger(__name__)

# Default network constants
DEFAULT_TIMEOUT = 15  # seconds per request
DEFAULT_RETRIES = 3
DEFAULT_BACKOFF = 1.0  # 1s, 2s, 4s …
RETRY_STATUS_CODES = (429, 500, 502, 503, 504)


def resolve_proxies(proxies: Optional[dict] = None) -> Optional[dict]:
    # Return a proxy dict, merging explicit *proxies* with environment fallback.
    if proxies:
        return proxies
    https = os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy")
    http = os.environ.get("HTTP_PROXY") or os.environ.get("http_proxy")
    if https or http:
        result: dict[str, str] = {}
        if https:
            result["https"] = https
        if http:
            result["http"] = http
        return result
    return None


class _TimeoutAdapter(HTTPAdapter):
    # HTTPAdapter that injects a default timeout on every request

    def __init__(self, timeout: float = DEFAULT_TIMEOUT, **kwargs):
        self._timeout = timeout
        super().__init__(**kwargs)

    def send(self, request, **kwargs):
        kwargs.setdefault("timeout", self._timeout)
        return super().send(request, **kwargs)


def make_session(
    *,
    proxies: Optional[dict] = None,
    timeout: float = DEFAULT_TIMEOUT,
    retries: int = DEFAULT_RETRIES,
    backoff: float = DEFAULT_BACKOFF,
) -> requests.Session:
    # Create a :class:`requests.Session` with retry, backoff, timeout, and proxy support
    session = requests.Session()

    retry = Retry(
        total=retries,
        backoff_factor=backoff,
        status_forcelist=list(RETRY_STATUS_CODES),
        allowed_methods=["GET", "POST", "HEAD"],
        raise_on_status=False,
    )
    adapter = _TimeoutAdapter(timeout=timeout, max_retries=retry)
    session.mount("https://", adapter)
    session.mount("http://", adapter)

    resolved = resolve_proxies(proxies)
    if resolved:
        session.proxies.update(resolved)
        logger.debug("Session proxies configured: %s", list(resolved.keys()))

    return session
