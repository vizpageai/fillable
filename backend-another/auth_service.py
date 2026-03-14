from __future__ import annotations

import json
from urllib import error, parse, request

from config import SETTINGS


class AuthServiceError(RuntimeError):
    pass


def _http_json(url: str, payload: dict | None = None, *, form: dict | None = None) -> dict:
    data = None
    headers = {}
    if payload is not None:
        data = json.dumps(payload).encode("utf-8")
        headers["Content-Type"] = "application/json"
    if form is not None:
        data = parse.urlencode(form).encode("utf-8")
        headers["Content-Type"] = "application/x-www-form-urlencoded"
    req = request.Request(url, data=data, headers=headers, method="POST")
    try:
        with request.urlopen(req, timeout=30) as resp:
            text = resp.read().decode("utf-8", errors="ignore")
    except error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="ignore")
        raise AuthServiceError(f"HTTP {exc.code} from {url}: {details[:1500]}") from exc
    except error.URLError as exc:
        raise AuthServiceError(f"Network error calling {url}: {exc}") from exc
    try:
        data = json.loads(text)
        if isinstance(data, dict):
            return data
        raise ValueError("Response is not a JSON object.")
    except Exception as exc:
        raise AuthServiceError(f"Invalid JSON response from {url}: {text[:1500]}") from exc


def _require_api_key() -> str:
    if not SETTINGS.firebase_web_api_key:
        raise AuthServiceError("FIREBASE_WEB_API_KEY is required.")
    return SETTINGS.firebase_web_api_key


def sign_in_with_password(email: str, password: str) -> dict:
    api_key = _require_api_key()
    url = f"https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key={api_key}"
    payload = {"email": email, "password": password, "returnSecureToken": True}
    return _http_json(url, payload)


def refresh_id_token(refresh_token: str) -> dict:
    api_key = _require_api_key()
    url = f"https://securetoken.googleapis.com/v1/token?key={api_key}"
    form = {"grant_type": "refresh_token", "refresh_token": refresh_token}
    return _http_json(url, form=form)
