from __future__ import annotations

import csv
import re
from pathlib import Path
from typing import Callable, Any

from app.codex_cli import CodexCli
from app.docx_formtext_engine import generate_docx_formtext_template
from app.doc_handlers import (
    apply_docx_blank_placeholders,
    apply_pptx_blank_placeholders,
    collect_docx_blanks,
    collect_pptx_blanks,
    copy_file,
    detect_type,
    extract_text,
    fill_pdf_form,
    list_pdf_form_fields,
    replace_in_docx,
    replace_in_pptx,
    replace_in_text,
)
from app.models import FillableTemplate, Placeholder
from app.utils import sanitize_name, save_json, load_json, truncate_text


def default_logger(_: str) -> None:
    return


def _unique_name(base: str, used: set[str]) -> str:
    if base not in used:
        return base
    i = 2
    while f"{base}_{i}" in used:
        i += 1
    return f"{base}_{i}"


def _find_blank_by_id(blanks, blank_id: int):
    for b in blanks:
        if int(getattr(b, "blank_id", -1)) == blank_id:
            return b
    return None


def _pair_override_name(blank, blanks) -> str:
    if not hasattr(blank, "para_index"):
        return ""
    labels = [str(x).strip() for x in getattr(blank, "nearby_labels", []) if str(x).strip()]
    if not labels:
        return ""

    has_assessment = any("ASSESSMENT EVIDENCE" in x.upper() for x in labels)
    has_target = any("LEARNING TARGET" in x.upper() for x in labels)
    if not (has_assessment and has_target):
        return ""

    prev_blank = _find_blank_by_id(blanks, int(blank.blank_id) - 1)
    next_blank = _find_blank_by_id(blanks, int(blank.blank_id) + 1)
    in_pair = (
        (prev_blank is not None and hasattr(prev_blank, "para_index") and blank.para_index - prev_blank.para_index <= 4)
        or (next_blank is not None and hasattr(next_blank, "para_index") and next_blank.para_index - blank.para_index <= 4)
    )
    if not in_pair:
        return ""

    # Adjacent blank pairs under two side-by-side headers: first blank is target, second is assessment.
    if prev_blank is not None and hasattr(prev_blank, "para_index") and blank.para_index - prev_blank.para_index <= 4:
        return "ASSESSMENT_EVIDENCE"
    if next_blank is not None and hasattr(next_blank, "para_index") and next_blank.para_index - blank.para_index <= 4:
        return "LEARNING_TARGET_SUCCESS_CRITERIA"
    return ""


def _template_path_for(source: Path) -> Path:
    return source.with_name(f"{source.stem}.fillable.json")


def _template_file_for(source: Path, source_type: str, has_pdf_fields: bool = False) -> Path:
    if source_type in {"docx", "pptx", "pdf"}:
        if source_type == "pdf" and not has_pdf_fields:
            return source.with_name(f"{source.stem}.template.txt")
        return source.with_name(f"{source.stem}.template{source.suffix}")
    return source.with_name(f"{source.stem}.template.txt")


PLACEHOLDER_PATTERN = re.compile(r"\{\{\s*([A-Za-z0-9_]+)\s*\}\}")


def _resolve_template_file_path(template_json_path: Path, raw_template_file: str) -> Path:
    template_file = Path(raw_template_file)
    if not template_file.is_absolute():
        template_file = template_json_path.resolve().parent / template_file
    return template_file.resolve()


def _extract_template_placeholder_names(template_file: Path, source_type: str) -> list[str]:
    if source_type == "pdf" and template_file.suffix.lower() == ".pdf":
        return []
    text = extract_text(template_file)
    names: list[str] = []
    seen: set[str] = set()
    for match in PLACEHOLDER_PATTERN.finditer(text):
        name = sanitize_name(match.group(1))
        if name in seen:
            continue
        seen.add(name)
        names.append(name)
    return names


def _replace_template_tokens(template_file: Path, source_type: str, rename_map: dict[str, str]) -> None:
    replacements = {
        "{{" + old + "}}": "{{" + new + "}}"
        for old, new in rename_map.items()
        if old and new and old != new
    }
    if not replacements:
        return

    if source_type == "docx":
        replace_in_docx(template_file, template_file, replacements)
    elif source_type == "pptx":
        replace_in_pptx(template_file, template_file, replacements)
    else:
        replace_in_text(template_file, template_file, replacements)


def _placeholder_list_from_names(names: list[str], descriptions: dict[str, str]) -> list[Placeholder]:
    return [
        Placeholder(
            name=name,
            source_text="{{" + name + "}}",
            description=descriptions.get(name, ""),
        )
        for name in names
    ]


def sync_template_and_json(
    template_json_path: Path,
    log: Callable[[str], None] = default_logger,
) -> Path:
    template_json_path = template_json_path.resolve()
    raw = load_json(template_json_path)
    template = FillableTemplate.from_dict(raw)

    template_file = _resolve_template_file_path(template_json_path, template.template_file)
    if not template_file.exists():
        raise FileNotFoundError(f"Template file does not exist: {template_file}")

    if template.source_type == "pdf" and template.pdf_form_fields and template_file.suffix.lower() == ".pdf":
        return template_json_path

    json_mtime = template_json_path.stat().st_mtime_ns
    template_mtime = template_file.stat().st_mtime_ns

    json_names = [p.name for p in template.placeholders if p.name]
    json_descriptions = {p.name: p.description for p in template.placeholders if p.name}
    template_names = _extract_template_placeholder_names(template_file, template.source_type)

    if template_mtime > json_mtime:
        if template_names:
            template.placeholders = _placeholder_list_from_names(template_names, json_descriptions)
            save_json(template_json_path, template.to_dict())
            log(f"Synchronized JSON from template changes: {template_json_path}")
        return template_json_path

    if json_mtime > template_mtime and json_names and template_names:
        rename_map: dict[str, str] = {}

        for p in template.placeholders:
            source_token = str(p.source_text or "").strip()
            match = PLACEHOLDER_PATTERN.fullmatch(source_token)
            if not match:
                continue
            old_name = sanitize_name(match.group(1))
            new_name = sanitize_name(p.name)
            if old_name in template_names and old_name != new_name:
                rename_map[old_name] = new_name

        unresolved_old = [name for name in template_names if name not in rename_map]
        unresolved_new = [name for name in json_names if name not in rename_map.values()]
        for old_name, new_name in zip(unresolved_old, unresolved_new):
            if old_name != new_name:
                rename_map.setdefault(old_name, new_name)

        if rename_map:
            _replace_template_tokens(template_file, template.source_type, rename_map)
            log(f"Synchronized template from JSON placeholder edits: {template_file}")
            template_names = _extract_template_placeholder_names(template_file, template.source_type)

        names_for_json = template_names or json_names
        template.placeholders = _placeholder_list_from_names(names_for_json, json_descriptions)
        save_json(template_json_path, template.to_dict())

    return template_json_path


def create_template_from_user_template(
    template_file_path: Path,
    source_file_path: Path | None = None,
    log: Callable[[str], None] = default_logger,
) -> Path:
    template_file_path = template_file_path.resolve()
    source_type = detect_type(template_file_path)

    placeholders: list[Placeholder] = []
    pdf_fields: list[str] = []

    if source_type == "pdf":
        pdf_fields = list_pdf_form_fields(template_file_path)
        if not pdf_fields:
            text = extract_text(template_file_path)
            names = sorted({sanitize_name(m.group(1)) for m in PLACEHOLDER_PATTERN.finditer(text)})
            placeholders = [
                Placeholder(
                    name=name,
                    source_text="{{" + name + "}}",
                    description="Imported from user template placeholder",
                )
                for name in names
            ]
    else:
        text = extract_text(template_file_path)
        names = sorted({sanitize_name(m.group(1)) for m in PLACEHOLDER_PATTERN.finditer(text)})
        placeholders = [
            Placeholder(
                name=name,
                source_text="{{" + name + "}}",
                description="Imported from user template placeholder",
            )
            for name in names
        ]

    template = FillableTemplate(
        schema_version=1,
        source_file=str((source_file_path or template_file_path).resolve()),
        source_type=source_type,
        template_file=str(template_file_path),
        created_at_utc=FillableTemplate.now_iso(),
        placeholders=placeholders,
        pdf_form_fields=pdf_fields,
    )
    template_path = _template_path_for(template_file_path)
    save_json(template_path, template.to_dict())
    log(f"Template metadata saved: {template_path}")
    log(f"Template file linked: {template_file_path}")
    log(f"Detected {len(placeholders)} placeholder keys and {len(pdf_fields)} PDF form fields.")
    return template_path


def generate_template(
    source_path: Path,
    codex: CodexCli,
    log: Callable[[str], None] = default_logger,
) -> Path:
    source_path = source_path.resolve()
    source_type = detect_type(source_path)
    log(f"Reading source file: {source_path}")

    pdf_fields: list[str] = []
    placeholders: list[Placeholder] = []
    text = extract_text(source_path)
    used_blank_mode = False
    blank_placeholder_map: dict[int, str] = {}
    docx_blanks = []
    pptx_blanks = []

    if source_type == "docx":
        # Primary path for Word templates: parse underlying FORMTEXT fields in document.xml.
        formtext_placeholders = generate_docx_formtext_template(
            source_path,
            _template_file_for(source_path, source_type),
            codex,
            log,
        )
        if formtext_placeholders:
            placeholders = formtext_placeholders
            template_file = _template_file_for(source_path, source_type)
            template = FillableTemplate(
                schema_version=1,
                source_file=str(source_path),
                source_type=source_type,
                template_file=str(template_file),
                created_at_utc=FillableTemplate.now_iso(),
                placeholders=placeholders,
                pdf_form_fields=[],
            )
            template_path = _template_path_for(source_path)
            save_json(template_path, template.to_dict())
            log(f"Template saved: {template_path}")
            log(f"Template document saved: {template_file}")
            return template_path

    if source_type == "pdf":
        pdf_fields = list_pdf_form_fields(source_path)
    elif source_type == "docx":
        docx_blanks = collect_docx_blanks(source_path)
    elif source_type == "pptx":
        pptx_blanks = collect_pptx_blanks(source_path)

    if pdf_fields:
        placeholders = [
            Placeholder(name=sanitize_name(name), source_text=name, description="PDF form field")
            for name in pdf_fields
        ]
        log(f"Found {len(placeholders)} PDF form fields.")
    elif docx_blanks or pptx_blanks:
        used_blank_mode = True
        blanks = docx_blanks or pptx_blanks
        used_names: set[str] = set()

        # Deterministic mapping from local labels (e.g. "Learning Target: ____").
        for blank in blanks:
            pair_override = _pair_override_name(blank, blanks)
            hint = sanitize_name(pair_override or getattr(blank, "label_hint", ""))
            if not hint or hint in {"FIELD", "BLANK"}:
                continue
            hint = _unique_name(hint, used_names)
            used_names.add(hint)
            blank_placeholder_map[blank.blank_id] = "{{" + hint + "}}"
            placeholders.append(
                Placeholder(
                    name=hint,
                    source_text=f"[BLANK_{blank.blank_id}]",
                    description=f"From label hint: {getattr(blank, 'label_hint', '')}",
                )
            )

        blanks_payload = []
        for blank in blanks:
            if blank.blank_id in blank_placeholder_map:
                continue
            item = {
                "blank_id": int(blank.blank_id),
                "context": str(blank.context),
                "label_hint": str(getattr(blank, "label_hint", "")),
            }
            if hasattr(blank, "para_index"):
                item["para_index"] = int(blank.para_index)
            else:
                item["slide_index"] = int(blank.slide_index)
                item["shape_index"] = int(blank.shape_index)
            blanks_payload.append(item)

        if blanks_payload:
            prompt = (
                "You map template blanks to placeholder keys. "
                "Return JSON only with this schema: "
                "{\"blank_to_placeholder\":[{\"blank_id\":0,\"name\":\"KEY\",\"description\":\"...\"}]}. "
                "Rules: keep blank_id unchanged, max 50 items, name in UPPER_SNAKE_CASE, "
                "use stable semantic names that match the nearby label/context. "
                "Input document text:\n"
                f"{truncate_text(text)}\n\n"
                f"Unresolved blank locations:\n{blanks_payload}"
            )
            log(
                f"Found {len(blanks)} blank regions "
                f"({len(blank_placeholder_map)} resolved by labels). Calling Codex for remaining..."
            )
            result = codex.run_json_prompt(prompt)
            items = result.parsed_json.get("blank_to_placeholder", [])
            mapped_ids: set[int] = set()
            for item in items:
                try:
                    blank_id = int(item.get("blank_id"))
                except Exception:
                    continue
                if blank_id < 0 or blank_id >= len(blanks) or blank_id in mapped_ids:
                    continue
                if blank_id in blank_placeholder_map:
                    continue
                name = sanitize_name(str(item.get("name", "")))
                if not name:
                    name = f"FIELD_{blank_id + 1}"
                name = _unique_name(name, used_names)
                used_names.add(name)
                mapped_ids.add(blank_id)
                blank_placeholder_map[blank_id] = "{{" + name + "}}"
                placeholders.append(
                    Placeholder(
                        name=name,
                        source_text=f"[BLANK_{blank_id}]",
                        description=str(item.get("description", "")).strip(),
                    )
                )
        for blank in blanks:
            if blank.blank_id in blank_placeholder_map:
                continue
            name = f"FIELD_{blank.blank_id + 1}"
            name = _unique_name(name, used_names)
            used_names.add(name)
            blank_placeholder_map[blank.blank_id] = "{{" + name + "}}"
            placeholders.append(
                Placeholder(
                    name=name,
                    source_text=f"[BLANK_{blank.blank_id}]",
                    description="Auto generated blank field",
                )
            )
        log(f"Created {len(placeholders)} placeholders from blank regions.")
    else:
        prompt = (
            "You generate placeholder plans for business documents. "
            "Return JSON only with this schema: "
            "{\"placeholders\":[{\"name\":\"...\",\"source_text\":\"...\",\"description\":\"...\"}]}. "
            "Rules: max 30 placeholders, names in UPPER_SNAKE_CASE, "
            "source_text must be exact snippets from the input. "
            "Input document text:\n"
            f"{truncate_text(text)}"
        )
        log("Calling Codex CLI to generate placeholder plan...")
        result = codex.run_json_prompt(prompt)
        items = result.parsed_json.get("placeholders", [])
        for item in items:
            name = sanitize_name(str(item.get("name", "")))
            source_text = str(item.get("source_text", "")).strip()
            if not name:
                continue
            placeholders.append(
                Placeholder(
                    name=name,
                    source_text=source_text,
                    description=str(item.get("description", "")).strip(),
                )
            )
        log(f"Codex proposed {len(placeholders)} placeholders.")

    template_file = _template_file_for(source_path, source_type, has_pdf_fields=bool(pdf_fields))
    mapping = {
        p.source_text: "{{" + p.name + "}}"
        for p in placeholders
        if p.source_text
    }

    if source_type == "docx":
        if used_blank_mode:
            apply_docx_blank_placeholders(
                source_path,
                template_file,
                blank_placeholder_map,
                docx_blanks,
            )
        else:
            replace_in_docx(source_path, template_file, mapping)
    elif source_type == "pptx":
        if used_blank_mode:
            apply_pptx_blank_placeholders(
                source_path,
                template_file,
                blank_placeholder_map,
                pptx_blanks,
            )
        else:
            replace_in_pptx(source_path, template_file, mapping)
    elif source_type == "pdf":
        if pdf_fields:
            copy_file(source_path, template_file)
        else:
            preview = text
            for old, new in mapping.items():
                preview = preview.replace(old, new)
            template_file.write_text(preview, encoding="utf-8")
    else:
        replace_in_text(source_path, template_file, mapping)

    template = FillableTemplate(
        schema_version=1,
        source_file=str(source_path),
        source_type=source_type,
        template_file=str(template_file),
        created_at_utc=FillableTemplate.now_iso(),
        placeholders=placeholders,
        pdf_form_fields=pdf_fields,
    )
    template_path = _template_path_for(source_path)
    save_json(template_path, template.to_dict())
    log(f"Template saved: {template_path}")
    log(f"Template document saved: {template_file}")
    return template_path


def fill_template(
    template_json_path: Path,
    codex: CodexCli,
    context_files: list[Path] | None = None,
    extra_instructions: str = "",
    log: Callable[[str], None] = default_logger,
) -> Path:
    template_json_path = template_json_path.resolve()
    sync_template_and_json(template_json_path, log=log)
    data = load_json(template_json_path)
    template = FillableTemplate.from_dict(data)

    context_files = context_files or []
    context_chunks: list[str] = []
    for ctx in context_files:
        try:
            body = extract_text(ctx.resolve())
        except Exception as exc:
            body = f"[Could not parse {ctx}: {exc}]"
        context_chunks.append(f"Context file: {ctx}\n{truncate_text(body, 10000)}")

    placeholders = [p.name for p in template.placeholders] or [sanitize_name(x) for x in template.pdf_form_fields]
    prompt = (
        "You fill document placeholders. Return JSON only with schema "
        "{\"values\":{\"PLACEHOLDER\":\"value\"}}. "
        "Provide a value for every placeholder listed. "
        "Keep values concise and professional.\n\n"
        f"Placeholders:\n{placeholders}\n\n"
        f"Extra instructions:\n{extra_instructions.strip() or 'None'}\n\n"
        f"Template metadata:\n{data}\n\n"
        f"Additional context:\n{'\n\n'.join(context_chunks) or 'None'}"
    )
    log("Calling Codex CLI to generate filled values...")
    result = codex.run_json_prompt(prompt)
    values_raw = result.parsed_json.get("values", {})
    values = {sanitize_name(str(k)): str(v) for k, v in values_raw.items()}

    for name in placeholders:
        values.setdefault(name, "")

    template_file = Path(template.template_file)
    filled_path = template_file.with_name(f"{template_file.stem}.filled{template_file.suffix}")
    _write_filled_output(template, values, filled_path)

    log(f"Filled output saved: {filled_path}")
    return filled_path


def _build_context_chunks(context_files: list[Path] | None) -> list[str]:
    context_files = context_files or []
    context_chunks: list[str] = []
    for ctx in context_files:
        try:
            body = extract_text(ctx.resolve())
        except Exception as exc:
            body = f"[Could not parse {ctx}: {exc}]"
        context_chunks.append(f"Context file: {ctx}\n{truncate_text(body, 10000)}")
    return context_chunks


def _generate_values_with_codex(
    codex: CodexCli,
    placeholders: list[str],
    template_metadata: dict[str, Any],
    extra_instructions: str,
    context_chunks: list[str],
    record_payload: dict[str, Any] | None = None,
) -> dict[str, str]:
    prompt = (
        "You fill document placeholders. Return JSON only with schema "
        "{\"values\":{\"PLACEHOLDER\":\"value\"}}. "
        "Provide a value for every placeholder listed. "
        "Keep values concise and professional.\n\n"
        f"Placeholders:\n{placeholders}\n\n"
        f"Extra instructions:\n{extra_instructions.strip() or 'None'}\n\n"
        f"Template metadata:\n{template_metadata}\n\n"
        f"Record data:\n{record_payload or 'None'}\n\n"
        f"Additional context:\n{'\n\n'.join(context_chunks) or 'None'}"
    )
    result = codex.run_json_prompt(prompt)
    values_raw = result.parsed_json.get("values", {})
    values = {sanitize_name(str(k)): str(v) for k, v in values_raw.items()}
    for name in placeholders:
        values.setdefault(name, "")
    return values


def _match_batch_columns_with_codex(
    codex: CodexCli,
    placeholder_names: list[str],
    placeholder_descriptions: dict[str, str],
    records: list[dict[str, str]],
    log: Callable[[str], None],
) -> dict[str, str]:
    columns: list[str] = []
    seen: set[str] = set()
    for record in records:
        for key in record.keys():
            if key not in seen:
                seen.add(key)
                columns.append(key)

    normalized_column_lookup = {sanitize_name(c): c for c in columns}
    mapping: dict[str, str] = {}
    unresolved: list[str] = []
    for name in placeholder_names:
        direct = normalized_column_lookup.get(name)
        if direct:
            mapping[name] = direct
        else:
            unresolved.append(name)

    if not unresolved or not columns:
        return mapping

    sample_rows = records[:5]
    prompt = (
        "Map data columns to placeholder keys. Return JSON only with schema "
        "{\"column_to_placeholder\":[{\"column\":\"...\",\"placeholder\":\"...\",\"reason\":\"...\"}]}. "
        "Rules: only map when reasonably confident, do not invent columns/placeholders, one column per placeholder.\n\n"
        f"Placeholders:\n{[{'name': n, 'description': placeholder_descriptions.get(n, '')} for n in unresolved]}\n\n"
        f"Available columns:\n{columns}\n\n"
        f"Sample rows:\n{sample_rows}"
    )
    log("Calling Codex CLI to map batch data columns to template placeholders...")
    result = codex.run_json_prompt(prompt)
    items = result.parsed_json.get("column_to_placeholder", [])
    used_columns = set(mapping.values())
    for item in items:
        placeholder = sanitize_name(str(item.get("placeholder", "")))
        column = str(item.get("column", "")).strip()
        if not placeholder or placeholder not in unresolved:
            continue
        if column not in seen or column in used_columns:
            continue
        mapping[placeholder] = column
        used_columns.add(column)

    return mapping


def _write_filled_output(template: FillableTemplate, values: dict[str, str], output_path: Path) -> None:
    template_file = Path(template.template_file)
    source_type = template.source_type
    replacements = {"{{" + k + "}}": v for k, v in values.items()}

    if source_type == "docx":
        replace_in_docx(template_file, output_path, replacements)
    elif source_type == "pptx":
        replace_in_pptx(template_file, output_path, replacements)
    elif source_type == "pdf":
        if template.pdf_form_fields:
            field_values = {}
            for field_name in template.pdf_form_fields:
                key = sanitize_name(field_name)
                field_values[field_name] = values.get(key, "")
            fill_pdf_form(template_file, output_path, field_values)
        else:
            replace_in_text(template_file, output_path, replacements)
    else:
        replace_in_text(template_file, output_path, replacements)


def _load_batch_records(data_path: Path) -> list[dict[str, str]]:
    data_path = data_path.resolve()
    suffix = data_path.suffix.lower()
    records: list[dict[str, str]] = []

    if suffix == ".csv":
        with data_path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            for row in reader:
                records.append({str(k): "" if v is None else str(v) for k, v in row.items()})
        return records

    if suffix == ".json":
        raw = load_json(data_path)
        if isinstance(raw, list):
            source = raw
        elif isinstance(raw, dict) and isinstance(raw.get("records"), list):
            source = raw.get("records", [])
        else:
            raise ValueError("JSON batch data must be a list of objects or an object with a 'records' list.")
        for item in source:
            if not isinstance(item, dict):
                continue
            records.append({str(k): "" if v is None else str(v) for k, v in item.items()})
        return records

    raise ValueError("Batch data must be a .csv or .json file.")


def _slug_from_record(record: dict[str, str]) -> str:
    candidates = (
        "ID",
        "STUDENT_ID",
        "CUSTOMER_ID",
        "NAME",
        "FULL_NAME",
    )
    normalized = {sanitize_name(k): str(v).strip() for k, v in record.items()}
    for key in candidates:
        value = normalized.get(key, "")
        if value:
            return sanitize_name(value).lower()[:40]
    return ""


def fill_template_multiple(
    template_json_path: Path,
    codex: CodexCli,
    records_path: Path,
    context_files: list[Path] | None = None,
    extra_instructions: str = "",
    output_dir: Path | None = None,
    log: Callable[[str], None] = default_logger,
) -> list[Path]:
    template_json_path = template_json_path.resolve()
    sync_template_and_json(template_json_path, log=log)
    template_data = load_json(template_json_path)
    template = FillableTemplate.from_dict(template_data)
    records = _load_batch_records(records_path)
    if not records:
        raise ValueError("No records found in batch data file.")

    template_file = Path(template.template_file)
    batch_dir = (output_dir or template_file.parent / f"{template_file.stem}.batch").resolve()
    batch_dir.mkdir(parents=True, exist_ok=True)

    placeholder_names = [p.name for p in template.placeholders] or [sanitize_name(x) for x in template.pdf_form_fields]
    placeholder_descriptions = {p.name: p.description for p in template.placeholders}
    context_chunks = _build_context_chunks(context_files)
    column_mapping = _match_batch_columns_with_codex(
        codex=codex,
        placeholder_names=placeholder_names,
        placeholder_descriptions=placeholder_descriptions,
        records=records,
        log=log,
    )
    if column_mapping:
        log(f"Mapped {len(column_mapping)}/{len(placeholder_names)} placeholders from batch columns.")
    else:
        log("No placeholder/column mapping found from batch data.")

    outputs: list[Path] = []
    total = len(records)
    for idx, record in enumerate(records, start=1):
        mapped_values: dict[str, str] = {}
        for name in placeholder_names:
            column = column_mapping.get(name)
            mapped_values[name] = str(record.get(column, "") if column else "")
        record_payload = {
            "raw_record": record,
            "mapped_values": mapped_values,
            "column_mapping": column_mapping,
        }
        values = _generate_values_with_codex(
            codex=codex,
            placeholders=placeholder_names,
            template_metadata=template_data,
            extra_instructions=extra_instructions,
            context_chunks=context_chunks,
            record_payload=record_payload,
        )
        slug = _slug_from_record(record)
        suffix = f".{slug}" if slug else ""
        out_name = f"{template_file.stem}.filled.{idx:03d}{suffix}{template_file.suffix}"
        out_path = batch_dir / out_name
        _write_filled_output(template, values, out_path)
        outputs.append(out_path)
        log(f"Filled {idx}/{total}: {out_path}")

    return outputs
