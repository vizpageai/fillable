from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


@dataclass
class Placeholder:
    name: str
    source_text: str = ""
    description: str = ""


@dataclass
class FillableTemplate:
    schema_version: int
    source_file: str
    source_type: str
    template_file: str
    created_at_utc: str
    placeholders: list[Placeholder] = field(default_factory=list)
    pdf_form_fields: list[str] = field(default_factory=list)

    @staticmethod
    def now_iso() -> str:
        return datetime.now(timezone.utc).isoformat()

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "FillableTemplate":
        placeholders = [
            Placeholder(
                name=str(item.get("name", "")).strip(),
                source_text=str(item.get("source_text", "")),
                description=str(item.get("description", "")),
            )
            for item in raw.get("placeholders", [])
            if str(item.get("name", "")).strip()
        ]
        return cls(
            schema_version=int(raw.get("schema_version", 1)),
            source_file=str(raw["source_file"]),
            source_type=str(raw["source_type"]),
            template_file=str(raw["template_file"]),
            created_at_utc=str(raw.get("created_at_utc", cls.now_iso())),
            placeholders=placeholders,
            pdf_form_fields=[str(v) for v in raw.get("pdf_form_fields", [])],
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": self.schema_version,
            "source_file": self.source_file,
            "source_type": self.source_type,
            "template_file": self.template_file,
            "created_at_utc": self.created_at_utc,
            "placeholders": [
                {
                    "name": p.name,
                    "source_text": p.source_text,
                    "description": p.description,
                }
                for p in self.placeholders
            ],
            "pdf_form_fields": self.pdf_form_fields,
        }


@dataclass
class AppConfig:
    codex_command_template: str = (
        'Get-Content -Raw "{prompt_file}" | codex exec --skip-git-repo-check --output-last-message "{output_file}"'
    )

    @classmethod
    def default_path(cls) -> Path:
        appdata = Path.home() / "AppData" / "Roaming" / "Fillable"
        appdata.mkdir(parents=True, exist_ok=True)
        return appdata / "config.json"
