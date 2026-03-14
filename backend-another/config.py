from __future__ import annotations

import json
import os
from dataclasses import dataclass

from dotenv import load_dotenv


load_dotenv()


def _truthy(value: str) -> bool:
    return value.strip().lower() in {"1", "true", "yes", "on"}


@dataclass(frozen=True)
class Settings:
    host: str = os.getenv("HOST", "0.0.0.0")
    port: int = int(os.getenv("PORT", "8081"))
    log_level: str = os.getenv("LOG_LEVEL", "info")

    openai_api_key: str = os.getenv("OPENAI_API_KEY", "")
    openai_base_url: str = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1").rstrip("/")

    stripe_secret_key: str = os.getenv("STRIPE_SECRET_KEY", "")
    stripe_webhook_secret: str = os.getenv("STRIPE_WEBHOOK_SECRET", "")
    stripe_price_id: str = os.getenv("STRIPE_PRICE_ID", "")
    stripe_success_url: str = os.getenv("STRIPE_SUCCESS_URL", "")
    stripe_cancel_url: str = os.getenv("STRIPE_CANCEL_URL", "")
    stripe_portal_return_url: str = os.getenv("STRIPE_PORTAL_RETURN_URL", "")

    firebase_credentials_json: str = os.getenv("FIREBASE_CREDENTIALS_JSON", "")
    firebase_credentials_path: str = os.getenv("FIREBASE_CREDENTIALS_PATH", "")
    firebase_google_credentials_path: str = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "")
    firebase_project_id: str = os.getenv("FIREBASE_PROJECT_ID", "")
    firebase_web_api_key: str = os.getenv("FIREBASE_WEB_API_KEY", "")

    require_email_verified: bool = _truthy(os.getenv("REQUIRE_EMAIL_VERIFIED", "false"))

    def firebase_credentials_dict(self) -> dict | None:
        if not self.firebase_credentials_json.strip():
            return None
        raw = self.firebase_credentials_json.strip()
        if raw.lower().endswith(".json") and os.path.exists(raw):
            try:
                return json.loads(open(raw, "r", encoding="utf-8").read())
            except Exception:
                return None
        try:
            return json.loads(raw)
        except Exception:
            return None


SETTINGS = Settings()
