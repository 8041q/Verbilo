# Resilient HTTP session factory — retry, backoff, jitter, timeout, proxy support

from __future__ import annotations

import os
import logging
import random
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

# HTTP status codes that are NOT transient — should never trigger retry or fallback
NON_TRANSIENT_STATUS_CODES = frozenset({400, 401, 403, 404, 405, 413, 415, 422})


def is_transient_error(exc: BaseException) -> bool:
    # Return True if *exc* looks like a transient network/server error
    # requests-level HTTP errors with a response attached
    if isinstance(exc, requests.exceptions.HTTPError) and exc.response is not None:
        return exc.response.status_code not in NON_TRANSIENT_STATUS_CODES
    # Connection / timeout errors are always transient
    if isinstance(exc, (requests.exceptions.ConnectionError,
                        requests.exceptions.Timeout)):
        return True
    # urllib3 retries exhausted — transient by definition
    if isinstance(exc, requests.exceptions.RetryError):
        return True
    # Programming / data errors are never transient
    if isinstance(exc, (ValueError, TypeError, KeyError, AttributeError)):
        return False
    # Default: assume transient to avoid silently dropping translations
    return True


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


class _JitteredRetry(Retry):
    # Retry subclass that adds random jitter to backoff delays and logs each retry attempt at DEBUG level

    def get_backoff_time(self) -> float:
        base = super().get_backoff_time()
        if base <= 0:
            return 0
        jittered = base * random.uniform(0.5, 1.5)
        return jittered

    def increment(self, method=None, url=None, response=None,
                  error=None, _pool=None, _stacktrace=None):
        # Log the retry attempt at DEBUG level before delegating
        retry_count = (self.total or 0)
        status = response.status if response else "N/A"
        logger.debug(
            "HTTP retry: attempt remaining=%s, status=%s, url=%s, error=%s",
            retry_count, status, url, error,
        )
        return super().increment(
            method=method, url=url, response=response,
            error=error, _pool=_pool, _stacktrace=_stacktrace,
        )


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
    # Create a :class:`requests.Session` with retry, jittered backoff, timeout,
    # and proxy support.
    session = requests.Session()

    retry = _JitteredRetry(
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
