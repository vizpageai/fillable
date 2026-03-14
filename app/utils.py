from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from app.models import AppConfig


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def save_json(path: Path, payload: dict[str, Any]) -> None:
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def load_config() -> AppConfig:
    config_path = AppConfig.default_path()
    if not config_path.exists():
        config = AppConfig()
        save_config(config)
        return config
    raw = load_json(config_path)
    credit_balance = float(raw.get("credit_balance", 0.0) or 0.0)
    if bool(raw.get("subscription_active", False)):
        credit_balance = max(credit_balance, 0.01)
    legacy_mode = str(raw.get("ai_mode", "")).strip().lower()
    if legacy_mode == "app_subscription":
        credit_balance = max(credit_balance, 0.01)
    model = str(raw.get("openai_model", AppConfig().openai_model)).strip() or AppConfig().openai_model
    backend_api_base = str(raw.get("backend_api_base", AppConfig().backend_api_base)).strip()
    firebase_id_token = str(raw.get("firebase_id_token", "")).strip()
    firebase_token_expiry_utc = int(raw.get("firebase_token_expiry_utc", 0) or 0)
    firebase_email = str(raw.get("firebase_email", "")).strip()
    firebase_uid = str(raw.get("firebase_uid", "")).strip()
    cmd = str(raw.get("codex_command_template", AppConfig().codex_command_template))
    return AppConfig(
        credit_balance=credit_balance,
        openai_model=model,
        backend_api_base=backend_api_base,
        firebase_id_token=firebase_id_token,
        firebase_token_expiry_utc=firebase_token_expiry_utc,
        firebase_email=firebase_email,
        firebase_uid=firebase_uid,
        codex_command_template=cmd,
    )


def save_config(config: AppConfig) -> None:
    save_json(
        AppConfig.default_path(),
        {
            "credit_balance": config.credit_balance,
            "openai_model": config.openai_model,
            "backend_api_base": config.backend_api_base,
            "firebase_id_token": config.firebase_id_token,
            "firebase_token_expiry_utc": config.firebase_token_expiry_utc,
            "firebase_email": config.firebase_email,
            "firebase_uid": config.firebase_uid,
            "codex_command_template": config.codex_command_template,
        },
    )


def sanitize_name(value: str) -> str:
    out = []
    for ch in value.upper():
        if ch.isalnum() or ch == "_":
            out.append(ch)
        else:
            out.append("_")
    cleaned = "".join(out)
    while "__" in cleaned:
        cleaned = cleaned.replace("__", "_")
    cleaned = cleaned.strip("_")
    return cleaned or "FIELD"


def truncate_text(value: str, max_len: int = 16000) -> str:
    value = value.strip()
    if len(value) <= max_len:
        return value
    return value[:max_len] + "\n\n[TRUNCATED]"
