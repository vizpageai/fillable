from __future__ import annotations

import json
import re
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path

from app.models import AppConfig


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
        template = self.config.codex_command_template.strip()
        if "{prompt}" not in template and "{prompt_file}" not in template:
            raise CodexCliError(
                "codex_command_template must include '{prompt}' or '{prompt_file}' placeholder"
            )

        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            suffix=".txt",
            delete=False,
        ) as temp:
            temp.write(prompt)
            temp_path = Path(temp.name)
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            suffix=".txt",
            delete=False,
        ) as out:
            output_file_path = Path(out.name)
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            suffix=".json",
            delete=False,
        ) as schema:
            schema.write('{"type":"object"}')
            schema_file_path = Path(schema.name)

        command = template
        command = command.replace("{prompt}", prompt.replace('"', '\\"'))
        command = command.replace("{prompt_file}", str(temp_path).replace('"', '\\"'))
        command = command.replace("{output_file}", str(output_file_path).replace('"', '\\"'))
        command = command.replace("{schema_file}", str(schema_file_path).replace('"', '\\"'))
        try:
            result = subprocess.run(
                ["powershell", "-NoProfile", "-Command", command],
                capture_output=True,
                text=True,
                shell=False,
                check=False,
            )
        finally:
            try:
                temp_path.unlink(missing_ok=True)
            except Exception:
                pass
        output_file_text = ""
        try:
            output_file_text = output_file_path.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            output_file_text = ""
        try:
            output_file_path.unlink(missing_ok=True)
        except Exception:
            pass
        try:
            schema_file_path.unlink(missing_ok=True)
        except Exception:
            pass
        output = output_file_text + "\n" + (result.stdout or "") + "\n" + (result.stderr or "")
        if result.returncode != 0:
            raise CodexCliError(
                f"Codex CLI exited with code {result.returncode}. Output:\n{output.strip()}"
            )

        parsed = self._extract_json(output)
        if parsed is None:
            raise CodexCliError(
                "Codex response did not contain valid JSON. "
                "Adjust your Codex command template or prompt constraints. "
                f"Raw output:\n{output.strip()[:3000]}"
            )
        return CodexResult(raw_output=output, parsed_json=parsed)

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
