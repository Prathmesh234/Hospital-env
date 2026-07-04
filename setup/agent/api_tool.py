"""Thin client for the Hospital-env REST API (GET-only).

Agents can hit any documented endpoint (see ``/docs`` on the running server)
without crafting raw HTTP. Only ``GET`` is exposed — the API is the second
read surface next to SQL, useful for pre-joined, paginated views.
"""

from __future__ import annotations

from typing import Any

import httpx

from setup.config import get_settings

DEFAULT_TIMEOUT_S = 10.0


class APIToolError(Exception):
    """Raised when the API is unreachable or returns garbage."""


def base_url() -> str:
    settings = get_settings()
    host = settings.api_host if settings.api_host not in {"0.0.0.0", "::"} else "localhost"
    return f"http://{host}:{settings.api_port}"


def get(path: str, params: dict[str, Any] | None = None) -> tuple[int, Any]:
    """GET ``path`` (e.g. ``/patients``) and return ``(status_code, body)``.

    The body is parsed JSON when possible, raw text otherwise.
    """
    if not path.startswith("/"):
        path = "/" + path
    url = base_url() + path
    try:
        response = httpx.get(url, params=params or {}, timeout=DEFAULT_TIMEOUT_S)
    except httpx.HTTPError as exc:
        raise APIToolError(
            f"API request to {url} failed: {exc}. "
            "Is the server running? Start it with: uv run hospital-env serve"
        ) from exc
    try:
        body = response.json()
    except ValueError:
        body = response.text
    return response.status_code, body


def is_up() -> bool:
    try:
        status, _ = get("/health")
        return status == 200
    except APIToolError:
        return False
