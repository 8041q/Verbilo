from __future__ import annotations

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


class VerbiloSession(requests.Session):
    """Requests session with a real default timeout.

    ``requests.Session`` has no built-in per-session timeout.  Older Verbilo
    code merely stored ``_verbilo_timeout`` and calls without ``timeout=`` could
    therefore block indefinitely.  This subclass applies the configured value
    whenever a request does not explicitly override it.
    """

    def __init__(self, *, timeout: float | tuple[float, float] = 60.0) -> None:
        super().__init__()
        self._verbilo_timeout = timeout

    def request(self, method, url, **kwargs):  # type: ignore[override]
        kwargs.setdefault("timeout", self._verbilo_timeout)
        return super().request(method, url, **kwargs)


def make_session(*, proxies=None, timeout=60.0, retries=0, backoff=0.0):
    session = VerbiloSession(timeout=timeout)
    if proxies:
        session.proxies.update(proxies)

    retry_count = max(int(retries or 0), 0)
    if retry_count:
        # Only retry connection failures for idempotent methods.  Translation
        # POSTs are intentionally never replayed automatically because a model
        # may already be generating a response server-side.
        policy = Retry(
            total=retry_count,
            connect=retry_count,
            read=0,
            status=0,
            backoff_factor=max(float(backoff or 0.0), 0.0),
            allowed_methods=frozenset({"GET", "HEAD", "OPTIONS"}),
            raise_on_status=False,
        )
        adapter = HTTPAdapter(max_retries=policy)
        session.mount("http://", adapter)
        session.mount("https://", adapter)
    return session
