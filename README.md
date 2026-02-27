# Fillable

Windows desktop app that uses Codex CLI to:
- Generate AI placeholder templates from `.docx`, `.pptx`, `.pdf`
- Import your own template file and build `.fillable.json` metadata
- Fill template placeholders with AI-generated content
- Add Explorer right-click actions for one-click template generation

## Features
- Right-click `.docx/.pptx/.pdf` -> `Generate AI Template (Codex)`
- Produces:
  - `yourfile.template.<ext>` (or `yourfile.template.txt` for non-form PDFs)
  - `yourfile.fillable.json` metadata template
- Blank-first behavior for Word/PowerPoint:
  - Detects blank regions like `_____`, long dashes, or large whitespace
  - Replaces blanks with placeholders (instead of replacing existing content text)
  - Falls back to snippet replacement only when no blank regions are found
- Double-click `*.fillable.json` to open Fillable UI and run fill
- Optional extra context files (`.docx/.pptx/.pdf/.txt/.md`) sent to Codex during fill
- Edit generated/imported `.fillable.json` templates in-app before filling
- Open and edit the generated `.template` document before running fill

## Requirements
- Windows 10/11
- Python 3.10+
- Codex CLI installed and available in PATH (default command uses `codex exec`)

## Setup
1. Install dependencies:
```powershell
python -m pip install -r requirements.txt
```
2. Run app:
```powershell
python run_fillable.py
```

## Configure Codex command
In the app, set `Codex command template`.
It must include `{prompt}`.

Default:
```text
Get-Content -Raw "{prompt_file}" | codex exec --skip-git-repo-check --output-last-message "{output_file}"
```

Supported placeholders in this command template: `{prompt}`, `{prompt_file}`, `{output_file}`, `{schema_file}`.

If your Codex CLI uses different flags, replace this command accordingly.

## Install right-click context menu (current user)
From project root:
```powershell
python scripts\install_context_menu.py
```

If using packaged EXE:
```powershell
python scripts\install_context_menu.py --exe "C:\path\to\Fillable.exe"
```

Remove integration:
```powershell
python scripts\uninstall_context_menu.py
```

## CLI usage
Generate template:
```powershell
python run_fillable.py --generate-template "C:\docs\input.docx"
```

Fill template:
```powershell
python run_fillable.py --fill-template "C:\docs\input.fillable.json" --context "C:\ctx\a.pdf;C:\ctx\b.docx" --instructions "Use formal tone."
```

Import an existing template file (with `{{PLACEHOLDER}}` keys or PDF form fields):
```powershell
python run_fillable.py --import-template-file "C:\docs\my_template.docx"
```

Batch fill template from CSV/JSON data:
```powershell
python run_fillable.py --fill-template "C:\docs\input.fillable.json" --batch-data "C:\docs\students.csv"
```

Optional custom output folder:
```powershell
python run_fillable.py --fill-template "C:\docs\input.fillable.json" --batch-data "C:\docs\customers.json" --batch-output-dir "C:\docs\filled_contracts"
```

Batch file notes:
- `.csv`: headers can differ from placeholder names; batch mode uses Codex to map columns to placeholders
- `.json`: either an array of objects or `{ "records": [ ... ] }`
- `--instructions` also applies in batch mode (per-record generation through Codex)

## Build standalone EXE
```powershell
powershell -ExecutionPolicy Bypass -File scripts\build_exe.ps1
```
Output:
- `dist\Fillable.exe`

## Notes and limitations
- `.docx` and `.pptx` replacements are text-based and may miss placeholders split across runs.
- PDFs:
  - If PDF has AcroForm fields, app fills those fields directly and outputs `*.filled.pdf`.
  - For non-form PDFs, app creates `*.template.txt` and outputs `*.filled.txt`.
- Codex output must include valid JSON; prompts enforce this, but model/tool changes can still require command adjustment.
