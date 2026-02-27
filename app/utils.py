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
    legacy_defaults = {
        'codex exec --skip-git-repo-check --output-last-message "{prompt}"',
        'Get-Content -Raw "{prompt_file}" | codex exec --skip-git-repo-check --output-last-message',
        'Get-Content -Raw "{prompt_file}" | codex exec --skip-git-repo-check --output-schema "{schema_file}" --output-last-message "{output_file}"',
        'Get-Content -Raw "{prompt_file}" | codex exec --skip-git-repo-check --output-last-message "{output_file}"',
    }
    cmd = str(
        raw.get(
            "codex_command_template",
            AppConfig().codex_command_template,
        )
    )
    if cmd in legacy_defaults:
        cmd = AppConfig().codex_command_template
        save_json(config_path, {"codex_command_template": cmd})
    return AppConfig(
        codex_command_template=str(
            cmd
        )
    )


def save_config(config: AppConfig) -> None:
    save_json(
        AppConfig.default_path(),
        {"codex_command_template": config.codex_command_template},
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
