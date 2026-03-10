# Fillable

Windows desktop app that uses OpenAI to:
- Generate AI placeholder templates from `.docx`, `.pptx`, `.pdf`
- Import your own template file and build `.fillable.json` metadata
- Fill template placeholders with AI-generated content
- Add Explorer right-click actions for one-click template generation

## Features
- Right-click `.docx/.pptx/.pdf` -> `Generate AI Template (OpenAI)`
- Produces:
  - `yourfile.template.<ext>` (or `yourfile.template.txt` for non-form PDFs)
  - `yourfile.fillable.json` metadata template
- Blank-first behavior for Word/PowerPoint:
  - Detects blank regions like `_____`, long dashes, or large whitespace
  - Replaces blanks with placeholders (instead of replacing existing content text)
  - Falls back to snippet replacement only when no blank regions are found
- Double-click `*.fillable.json` to open Fillable UI and run fill
- Optional extra context files (`.docx/.pptx/.pdf/.txt/.md`) sent to OpenAI during fill
- Edit generated/imported `.fillable.json` templates in-app before filling
- Open and edit the generated `.template` document before running fill

## Requirements
- Windows 10/11
- Python 3.10+
- Internet access to call OpenAI APIs or your subscription backend

## Setup
1. Install dependencies:
```powershell
python -m pip install -r requirements.txt
```
2. Run app:
```powershell
python run_fillable.py
```

## Configure AI access
In the app Settings, choose one mode:
- `user_key`: user provides their own OpenAI API key
- `app_subscription`: app calls your backend proxy (backend uses your OpenAI API key)

Security:
- User API key is secured with Windows DPAPI in the per-user app data folder (not in `config.json`)
- Subscription token is stored in Windows Credential Manager (not in `config.json`)
- `config.json` stores only non-secret settings (mode/model/base URLs)
- Backend proxy contract: `docs/SUBSCRIPTION_BACKEND_API.md`
- First launch shows an onboarding screen asking for OpenAI API key, then opens the main app page.

## Install right-click context menu (current user)
From project root:
```powershell
python scripts\install_context_menu.py
```

If using packaged EXE:
```powershell
python scripts\install_context_menu.py --exe "C:\path\to\FillableDOC.exe"
```

Remove integration:
```powershell
python scripts\uninstall_context_menu.py
```

CLI alternatives (works with EXE too):
```powershell
python run_fillable.py --install-context-menu
python run_fillable.py --uninstall-context-menu
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
- `.csv`: headers can differ from placeholder names; batch mode uses AI to map columns to placeholders
- `.json`: either an array of objects or `{ "records": [ ... ] }`
- `--instructions` also applies in batch mode (per-record generation through AI)

Configure AI settings via CLI:
```powershell
python run_fillable.py --set-ai-mode user_key --set-openai-model gpt-4.1-mini
python run_fillable.py --set-user-openai-key "sk-..."
python run_fillable.py --set-ai-mode app_subscription --set-subscription-api-base "https://api.yourdomain.com"
python run_fillable.py --set-subscription-token "your-issued-token"
python run_fillable.py --print-config
```

## Build standalone EXE
```powershell
powershell -ExecutionPolicy Bypass -File scripts\build_exe.ps1
```
Output:
- `dist\FillableDOC.exe`

## Build installer for Microsoft Store (Win32 submission path)
This project includes an Inno Setup installer that:
- Installs `FillableDOC.exe`
- Registers right-click context menu on install
- Removes right-click context menu on uninstall

Build steps:
```powershell
powershell -ExecutionPolicy Bypass -File scripts\build_store_installer.ps1
```
Output:
- `dist\installer\FillableDOC-Setup.exe`

If `ISCC.exe` is installed at a different location:
```powershell
powershell -ExecutionPolicy Bypass -File scripts\build_store_installer.ps1 -IsccPath "C:\Path\To\ISCC.exe"
```

Store submission notes: `docs/MICROSOFT_STORE.md`

## Build MSI (WiX Toolset)
Install WiX Toolset v4:
```powershell
winget install WiXToolset.WiXToolset
```

Build MSI:
```powershell
powershell -ExecutionPolicy Bypass -File scripts\build_msi.ps1
```

If `wix` is not on PATH, pass full path:
```powershell
powershell -ExecutionPolicy Bypass -File scripts\build_msi.ps1 -WixPath "C:\Program Files\WiX Toolset v4.0\bin\wix.exe"
```

Output:
- `dist\installer\FillableDOC.msi`

## Build MSIX (GitHub Actions)
Workflow file:
- `.github/workflows/build-msix.yml`

It builds `dist/msix/*.msix` on `windows-latest`.

Optional signing secrets in GitHub Actions:
- `MSIX_CERT_BASE64`: Base64 of your `.pfx` certificate
- `MSIX_CERT_PASSWORD`: PFX password

If secrets are not set, the workflow still builds an unsigned MSIX.

## Sign release binaries (Microsoft Store)
Sign app EXE and installer outputs with your code-signing certificate:

Sign EXE + Setup EXE:
```powershell
powershell -ExecutionPolicy Bypass -File scripts\sign_release.ps1 -CertPath "C:\path\codesign.pfx" -PromptForPassword -IncludeSetupExe
```

Sign EXE + MSI:
```powershell
powershell -ExecutionPolicy Bypass -File scripts\sign_release.ps1 -CertPath "C:\path\codesign.pfx" -PromptForPassword -IncludeMsi
```

Verify signatures:
```powershell
powershell -ExecutionPolicy Bypass -File scripts\verify_signatures.ps1 -IncludeSetupExe
```

## Notes and limitations
- `.docx` and `.pptx` replacements are text-based and may miss placeholders split across runs.
- PDFs:
  - If PDF has AcroForm fields, app fills those fields directly and outputs `*.filled.pdf`.
  - For non-form PDFs, app creates `*.template.txt` and outputs `*.filled.txt`.
- Model output must include valid JSON; prompts enforce this, but model changes can still require prompt tuning.
