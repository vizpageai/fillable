from __future__ import annotations

import argparse
import sys
import winreg
from pathlib import Path


def set_value(root, subkey: str, name: str, value: str) -> None:
    key = winreg.CreateKey(root, subkey)
    winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)
    winreg.CloseKey(key)


def make_command(exe: str | None, project_root: Path) -> str:
    if exe:
        return f'"{exe}" "%1"'

    py = sys.executable
    main_py = project_root / "run_fillable.py"
    return f'"{py}" "{main_py}" "%1"'


def make_generate_command(exe: str | None, project_root: Path) -> str:
    if exe:
        return f'"{exe}" --generate-template "%1"'

    py = sys.executable
    main_py = project_root / "run_fillable.py"
    return f'"{py}" "{main_py}" --generate-template "%1"'


def make_fill_command(exe: str | None, project_root: Path) -> str:
    if exe:
        return f'"{exe}" --fill-template "%1" --prompt-instructions'

    py = sys.executable
    main_py = project_root / "run_fillable.py"
    return f'"{py}" "{main_py}" --fill-template "%1" --prompt-instructions'


def install(exe: str | None) -> None:
    project_root = Path(__file__).resolve().parents[1]
    generate_command = make_generate_command(exe, project_root)
    open_template_command = make_command(exe, project_root)
    fill_template_command = make_fill_command(exe, project_root)

    for ext in [".docx", ".pptx", ".pdf"]:
        base = fr"Software\Classes\SystemFileAssociations\{ext}\shell\FillableGenerateTemplate"
        set_value(winreg.HKEY_CURRENT_USER, base, "", "Generate AI Template (Codex)")
        set_value(winreg.HKEY_CURRENT_USER, base, "Icon", "imageres.dll,-5302")
        set_value(winreg.HKEY_CURRENT_USER, base + r"\command", "", generate_command)

    set_value(winreg.HKEY_CURRENT_USER, r"Software\Classes\.fillable.json", "", "Fillable.Template")
    set_value(winreg.HKEY_CURRENT_USER, r"Software\Classes\Fillable.Template", "", "Fillable Template")
    set_value(
        winreg.HKEY_CURRENT_USER,
        r"Software\Classes\Fillable.Template\shell\FillTemplate",
        "",
        "Fill Template (Codex)",
    )
    set_value(
        winreg.HKEY_CURRENT_USER,
        r"Software\Classes\Fillable.Template\shell\FillTemplate",
        "Icon",
        "imageres.dll,-5302",
    )
    set_value(
        winreg.HKEY_CURRENT_USER,
        r"Software\Classes\Fillable.Template\shell\FillTemplate\command",
        "",
        fill_template_command,
    )
    set_value(
        winreg.HKEY_CURRENT_USER,
        r"Software\Classes\Fillable.Template\shell\open\command",
        "",
        open_template_command,
    )

    for ext in [".docx", ".pptx", ".pdf", ".txt", ".json"]:
        base = fr"Software\Classes\SystemFileAssociations\{ext}\shell\FillableFillTemplate"
        set_value(winreg.HKEY_CURRENT_USER, base, "", "Fill Template (Codex)")
        set_value(winreg.HKEY_CURRENT_USER, base, "Icon", "imageres.dll,-5302")
        set_value(winreg.HKEY_CURRENT_USER, base + r"\command", "", fill_template_command)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--exe", default="", help="Path to built Fillable.exe")
    args = parser.parse_args()
    install(args.exe or None)
    print("Context menu installed for current user.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
