from __future__ import annotations

import json
from urllib import error, request


class BackendAuthError(RuntimeError):
    pass


def _http_json(url: str, payload: dict) -> dict:
    data = json.dumps(payload).encode("utf-8")
    headers = {"Content-Type": "application/json"}
    req = request.Request(url, data=data, headers=headers, method="POST")
    try:
        with request.urlopen(req, timeout=30) as resp:
            text = resp.read().decode("utf-8", errors="ignore")
    except error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="ignore")
        raise BackendAuthError(f"HTTP {exc.code} from {url}: {details[:1500]}") from exc
    except error.URLError as exc:
        raise BackendAuthError(f"Network error calling {url}: {exc}") from exc
    try:
        data = json.loads(text)
        if isinstance(data, dict):
            return data
        raise ValueError("Response is not a JSON object.")
    except Exception as exc:
        raise BackendAuthError(f"Invalid JSON response from {url}: {text[:1500]}") from exc


def register(backend_base: str, email: str, password: str) -> dict:
    url = backend_base.rstrip("/") + "/v1/auth/register"
    return _http_json(url, {"email": email, "password": password})


def login(backend_base: str, email: str, password: str) -> dict:
    url = backend_base.rstrip("/") + "/v1/auth/login"
    return _http_json(url, {"email": email, "password": password})


def refresh(backend_base: str, refresh_token: str) -> dict:
    url = backend_base.rstrip("/") + "/v1/auth/refresh"
    return _http_json(url, {"refresh_token": refresh_token})
