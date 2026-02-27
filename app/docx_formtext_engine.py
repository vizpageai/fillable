from __future__ import annotations

import re
import tempfile
import zipfile
from pathlib import Path
from typing import Callable

from app.codex_cli import CodexCli
from app.models import Placeholder
from app.utils import sanitize_name, truncate_text

EN_SPACE = "\u2002"

_FIELD_RE = re.compile(
    r'<w:fldChar w:fldCharType="begin">.*?<w:fldChar w:fldCharType="end"/>',
    re.DOTALL,
)
_EN_SPACE_RE = re.compile(r"<w:t[^>]*>(\u2002+)</w:t>")
_TEXT_RE = re.compile(r"<w:t[^>]*>([^<]+)</w:t>")


def _is_blank_formtext(field_xml: str) -> bool:
    return "FORMTEXT" in field_xml and EN_SPACE in field_xml


def _nearby_label(xml: str, field_start: int, window: int = 1200) -> str:
    before = xml[max(0, field_start - window): field_start]
    texts = _TEXT_RE.findall(before)
    texts = [t.strip() for t in texts if t.strip() and not all(c == EN_SPACE for c in t)]
    if not texts:
        return "(no label found)"
    return " ".join(texts[-10:])


def _extract_fields(xml: str) -> list[dict]:
    fields: list[dict] = []
    idx = 0
    for m in _FIELD_RE.finditer(xml):
        block = m.group()
        if not _is_blank_formtext(block):
            continue
        fields.append(
            {
                "field_index": idx,
                "nearby_label": _nearby_label(xml, m.start()),
                "match_start": m.start(),
                "match_end": m.end(),
            }
        )
        idx += 1
    return fields


def _doc_text_sample(xml: str) -> str:
    texts = _TEXT_RE.findall(xml)
    clean = [t for t in texts if t.strip() and not all(c == EN_SPACE for c in t)]
    return truncate_text("\n".join(clean), 6000)


def _inject_labels(xml: str, fields: list[dict], labels: dict[int, str]) -> str:
    result = xml
    for field in reversed(fields):
        idx = int(field["field_index"])
        label = labels.get(idx, f"FIELD_{idx + 1}")
        token = "{{" + label + "}}"
        block = result[field["match_start"]: field["match_end"]]
        new_block = _EN_SPACE_RE.sub(
            '<w:t xml:space="preserve">' + token + "</w:t>",
            block,
            count=1,
        )
        result = result[: field["match_start"]] + new_block + result[field["match_end"]:]
    return result


def _normalize_label_pairs(fields: list[dict], labels: dict[int, str]) -> None:
    used = set(labels.values())
    for idx in range(1, len(fields)):
        prev_name = labels.get(idx - 1, "")
        name = labels.get(idx, "")
        nearby = str(fields[idx].get("nearby_label", "")).strip().lower()
        if "learning_target" not in prev_name.lower():
            continue
        if "assessment" in name.lower():
            continue
        if "learning_target" not in name.lower() and not nearby.startswith("(no label found"):
            continue

        prefix = ""
        m = re.match(r"^([A-Z]{3,4})_", prev_name)
        if m:
            prefix = m.group(1) + "_"
        base = prefix + "ASSESSMENT_EVIDENCE"
        candidate = base
        c = 2
        while candidate in used:
            candidate = f"{base}_{c}"
            c += 1
        labels[idx] = candidate
        used.add(candidate)


def generate_docx_formtext_template(
    source_path: Path,
    template_path: Path,
    codex: CodexCli,
    log: Callable[[str], None],
) -> list[Placeholder]:
    with zipfile.ZipFile(source_path, "r") as zin:
        with tempfile.TemporaryDirectory() as tmpdir:
            zin.extractall(tmpdir)
            document_xml = Path(tmpdir) / "word" / "document.xml"
            if not document_xml.exists():
                return []
            xml = document_xml.read_text(encoding="utf-8")
            fields = _extract_fields(xml)
            if not fields:
                return []

            sample = _doc_text_sample(xml)
            prompt = (
                "You label Word FORMTEXT blanks. Return JSON only with schema "
                '{"labels":[{"field_index":0,"name":"UPPER_SNAKE_CASE","description":"..."}]}. '
                "Rules: assign every field_index exactly once; names must be concise semantic keys; "
                "use day prefixes when relevant (MON_, TUE_, ...); "
                "if unknown use FIELD_<index>. "
                "Document sample:\n"
                f"{sample}\n\n"
                f"Fields:\n{[{ 'field_index': f['field_index'], 'nearby_label': f['nearby_label']} for f in fields]}"
            )
            log(f"Detected {len(fields)} FORMTEXT blanks. Calling Codex for field names...")
            result = codex.run_json_prompt(prompt)
            items = result.parsed_json.get("labels", [])

            labels: dict[int, str] = {}
            placeholders: list[Placeholder] = []
            used: set[str] = set()
            for item in items:
                try:
                    idx = int(item.get("field_index"))
                except Exception:
                    continue
                if idx < 0 or idx >= len(fields) or idx in labels:
                    continue
                name = sanitize_name(str(item.get("name", "")))
                if not name:
                    name = f"FIELD_{idx + 1}"
                if name in used:
                    c = 2
                    while f"{name}_{c}" in used:
                        c += 1
                    name = f"{name}_{c}"
                used.add(name)
                labels[idx] = name
                placeholders.append(
                    Placeholder(
                        name=name,
                        source_text=f"[BLANK_{idx}]",
                        description=str(item.get("description", "")).strip(),
                    )
                )

            for idx in range(len(fields)):
                if idx in labels:
                    continue
                name = f"FIELD_{idx + 1}"
                if name in used:
                    c = 2
                    while f"{name}_{c}" in used:
                        c += 1
                    name = f"{name}_{c}"
                used.add(name)
                labels[idx] = name
                placeholders.append(
                    Placeholder(
                        name=name,
                        source_text=f"[BLANK_{idx}]",
                        description="Auto generated field",
                    )
                )

            _normalize_label_pairs(fields, labels)
            # Rebuild placeholder list after normalization.
            placeholders = [
                Placeholder(
                    name=labels[idx],
                    source_text=f"[BLANK_{idx}]",
                    description=next(
                        (p.description for p in placeholders if p.source_text == f"[BLANK_{idx}]"),
                        "",
                    ),
                )
                for idx in range(len(fields))
            ]

            new_xml = _inject_labels(xml, fields, labels)
            document_xml.write_text(new_xml, encoding="utf-8")

            with zipfile.ZipFile(template_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for path in Path(tmpdir).rglob("*"):
                    if path.is_file():
                        arcname = path.relative_to(tmpdir).as_posix()
                        zout.write(path, arcname)
            return placeholders
