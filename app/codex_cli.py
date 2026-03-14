from __future__ import annotations

import json
import re
from dataclasses import dataclass
from urllib import error, request

from app.backend_auth import BackendAuthError, refresh as backend_refresh
from app.firebase_auth import is_token_expired, jwt_expiry_epoch
from app.models import AppConfig
from app.secure_store import get_firebase_refresh_token, get_user_openai_key
from app.utils import save_config

USER_OPENAI_KEY_TARGET = "Fillable.OpenAI.UserKey"
FIREBASE_REFRESH_TOKEN_TARGET = "Fillable.Firebase.RefreshToken"


@dataclass
class CodexResult:
    raw_output: str
    parsed_json: dict


class CodexCliError(RuntimeError):
    pass


class CodexCli:
    def __init__(self, config: AppConfig):
        self.config = config

    def run_json_prompt(self, prompt: str) -> CodexResult:
        if self.config.credit_balance > 0 and (self.config.firebase_id_token or "").strip():
            output = self._call_backend(prompt)
        else:
            output = self._call_openai(prompt)
        parsed = self._extract_json(output)
        if parsed is None:
            raise CodexCliError(
                "Model response did not contain valid JSON. "
                f"Raw output:\n{output.strip()[:3000]}"
            )
        return CodexResult(raw_output=output, parsed_json=parsed)

    def _call_openai(self, prompt: str) -> str:
        api_key = get_user_openai_key(USER_OPENAI_KEY_TARGET)
        if not api_key:
            raise CodexCliError(
                "OpenAI API key is missing. Add your key in Settings or via --set-user-openai-key."
            )

        base = "https://api.openai.com/v1"
        url = f"{base}/chat/completions"
        payload = {
            "model": self.config.openai_model or "gpt-4.1-mini",
            "messages": [
                {"role": "system", "content": "Respond with a single JSON object only."},
                {"role": "user", "content": prompt},
            ],
            "temperature": 0.2,
            "response_format": {"type": "json_object"},
        }
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        }
        response = self._http_json("POST", url, payload, headers)
        try:
            return str(response["choices"][0]["message"]["content"])
        except Exception as exc:
            raise CodexCliError(f"Unexpected OpenAI response shape: {response}") from exc

    def _call_backend(self, prompt: str) -> str:
        base = (self.config.backend_api_base or "").strip().rstrip("/")
        if not base:
            raise CodexCliError("Backend URL is required for credit mode.")
        token = self._get_firebase_bearer()
        if not token:
            raise CodexCliError("Firebase login is required for credit mode.")
        url = f"{base}/v1/openai-proxy"
        payload = {
            "payload": {
                "model": self.config.openai_model or "gpt-4.1-mini",
                "messages": [
                    {"role": "system", "content": "Respond with a single JSON object only."},
                    {"role": "user", "content": prompt},
                ],
                "temperature": 0.2,
                "response_format": {"type": "json_object"},
            },
        }
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {token}",
        }
        response = self._http_json("POST", url, payload, headers)
        try:
            remaining = float(response.get("credits_remaining", 0.0) or 0.0)
            self.config.credit_balance = remaining
            save_config(self.config)
        except Exception:
            pass
        try:
            return str(response["choices"][0]["message"]["content"])
        except Exception as exc:
            raise CodexCliError(f"Unexpected backend response shape: {response}") from exc

    def _get_firebase_bearer(self) -> str:
        token = (self.config.firebase_id_token or "").strip()
        exp_epoch = int(self.config.firebase_token_expiry_utc or 0)
        if token and exp_epoch == 0:
            exp_epoch = jwt_expiry_epoch(token) or 0
        if token and not is_token_expired(exp_epoch):
            return token
        refresh = (get_firebase_refresh_token(FIREBASE_REFRESH_TOKEN_TARGET) or "").strip()
        base = (self.config.backend_api_base or "").strip()
        if not refresh or not base:
            return token
        try:
            refreshed = backend_refresh(base, refresh)
        except BackendAuthError as exc:
            raise CodexCliError(str(exc)) from exc
        token = str(refreshed.get("id_token", "") or refreshed.get("idToken", "")).strip()
        refresh = str(refreshed.get("refresh_token", "") or refreshed.get("refreshToken", "")).strip()
        expires_in = int(refreshed.get("expires_in", 0) or refreshed.get("expiresIn", 0) or 0)
        if not token:
            return ""
        if expires_in:
            exp_epoch = int(__import__("time").time()) + expires_in
        else:
            exp_epoch = jwt_expiry_epoch(token) or 0
        self.config.firebase_id_token = token
        if refresh:
            from app.secure_store import set_firebase_refresh_token

            set_firebase_refresh_token(FIREBASE_REFRESH_TOKEN_TARGET, refresh)
        self.config.firebase_token_expiry_utc = exp_epoch
        save_config(self.config)
        return token

    @staticmethod
    def _http_json(method: str, url: str, payload: dict, headers: dict[str, str]) -> dict:
        body = json.dumps(payload).encode("utf-8")
        req = request.Request(url, data=body, headers=headers, method=method.upper())
        try:
            with request.urlopen(req, timeout=120) as resp:
                text = resp.read().decode("utf-8", errors="ignore")
        except error.HTTPError as exc:
            details = exc.read().decode("utf-8", errors="ignore")
            raise CodexCliError(f"HTTP {exc.code} from {url}: {details[:1500]}") from exc
        except error.URLError as exc:
            raise CodexCliError(f"Network error calling {url}: {exc}") from exc
        try:
            data = json.loads(text)
            if isinstance(data, dict):
                return data
            raise ValueError("Response is not a JSON object.")
        except Exception as exc:
            raise CodexCliError(f"Invalid JSON response from {url}: {text[:1500]}") from exc

    @staticmethod
    def _extract_json(output: str) -> dict | None:
        output = output.strip()
        try:
            data = json.loads(output)
            return data if isinstance(data, dict) else None
        except Exception:
            pass

        decoder = json.JSONDecoder()
        for idx, ch in enumerate(output):
            if ch != "{":
                continue
            try:
                data, _ = decoder.raw_decode(output[idx:])
            except Exception:
                continue
            if isinstance(data, dict):
                return data

        block_match = re.search(r"```(?:json)?\s*(\{[\s\S]*?\})\s*```", output)
        if block_match:
            try:
                data = json.loads(block_match.group(1))
                return data if isinstance(data, dict) else None
            except Exception:
                return None

        brace_match = re.search(r"(\{[\s\S]*\})", output)
        if not brace_match:
            return None
        try:
            data = json.loads(brace_match.group(1))
            return data if isinstance(data, dict) else None
        except Exception:
            return None
