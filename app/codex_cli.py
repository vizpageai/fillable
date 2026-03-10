from __future__ import annotations

import json
import re
from dataclasses import dataclass
from urllib import error, request

from app.models import AppConfig
from app.secure_store import get_secret, get_user_openai_key

USER_OPENAI_KEY_TARGET = "Fillable.OpenAI.UserKey"
SUBSCRIPTION_TOKEN_TARGET = "Fillable.Subscription.Token"


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
        mode = (self.config.ai_mode or "user_key").strip().lower()
        if mode == "app_subscription":
            output = self._call_subscription_backend(prompt)
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

        base = (self.config.openai_api_base or "https://api.openai.com/v1").rstrip("/")
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

    def _call_subscription_backend(self, prompt: str) -> str:
        token = get_secret(SUBSCRIPTION_TOKEN_TARGET)
        if not token:
            raise CodexCliError(
                "Subscription token is missing. Sign in/purchase and set token in Settings."
            )
        base = (self.config.subscription_api_base or "").strip().rstrip("/")
        if not base:
            raise CodexCliError("Subscription backend URL is required in app-subscription mode.")
        url = f"{base}/v1/openai-proxy"
        payload = {
            "model": self.config.openai_model or "gpt-4.1-mini",
            "prompt": prompt,
            "response_format": "json_object",
        }
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }
        response = self._http_json("POST", url, payload, headers)
        text = str(response.get("output_text", "")).strip()
        if text:
            return text
        raw = response.get("output")
        if isinstance(raw, str) and raw.strip():
            return raw
        raise CodexCliError(f"Subscription backend did not return output JSON text: {response}")

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
