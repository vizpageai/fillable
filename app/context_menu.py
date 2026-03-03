from __future__ import annotations

import sys
import winreg
from pathlib import Path
from typing import Callable


def _set_value(root, subkey: str, name: str, value: str) -> None:
    key = winreg.CreateKey(root, subkey)
    winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)
    winreg.CloseKey(key)


def _delete_tree(root, subkey: str) -> None:
    try:
        with winreg.OpenKey(root, subkey, 0, winreg.KEY_READ | winreg.KEY_WRITE) as key:
            while True:
                try:
                    child = winreg.EnumKey(key, 0)
                except OSError:
                    break
                _delete_tree(root, subkey + "\\" + child)
    except FileNotFoundError:
        return

    try:
        winreg.DeleteKey(root, subkey)
    except FileNotFoundError:
        pass


def _project_root() -> Path:
    return Path(__file__).resolve().parents[1]


def _command_target(exe_override: str | None = None) -> str:
    if exe_override:
        return f'"{exe_override}"'
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}"'
    main_py = _project_root() / "run_fillable.py"
    return f'"{sys.executable}" "{main_py}"'


def _make_open_command(exe_override: str | None = None) -> str:
    return f"{_command_target(exe_override)} \"%1\""


def _make_generate_command(exe_override: str | None = None) -> str:
    return f"{_command_target(exe_override)} --generate-template \"%1\""


def _make_fill_command(exe_override: str | None = None, prompt_instructions: bool = True) -> str:
    command = f"{_command_target(exe_override)} --fill-template \"%1\""
    if prompt_instructions:
        command += " --prompt-instructions"
    return command


def install_context_menu(
    exe_override: str | None = None,
    prompt_instructions: bool = True,
) -> None:
    generate_command = _make_generate_command(exe_override)
    open_template_command = _make_open_command(exe_override)
    fill_template_command = _make_fill_command(exe_override, prompt_instructions=prompt_instructions)

    for ext in [".docx", ".pptx", ".pdf"]:
        base = fr"Software\Classes\SystemFileAssociations\{ext}\shell\FillableGenerateTemplate"
        _set_value(winreg.HKEY_CURRENT_USER, base, "", "Generate AI Template (Codex)")
        _set_value(winreg.HKEY_CURRENT_USER, base, "Icon", "imageres.dll,-5302")
        _set_value(winreg.HKEY_CURRENT_USER, base + r"\command", "", generate_command)

    _set_value(winreg.HKEY_CURRENT_USER, r"Software\Classes\.fillable.json", "", "Fillable.Template")
    _set_value(winreg.HKEY_CURRENT_USER, r"Software\Classes\Fillable.Template", "", "Fillable Template")
    _set_value(
        winreg.HKEY_CURRENT_USER,
        r"Software\Classes\Fillable.Template\shell\FillTemplate",
        "",
        "Fill Template (Codex)",
    )
    _set_value(
        winreg.HKEY_CURRENT_USER,
        r"Software\Classes\Fillable.Template\shell\FillTemplate",
        "Icon",
        "imageres.dll,-5302",
    )
    _set_value(
        winreg.HKEY_CURRENT_USER,
        r"Software\Classes\Fillable.Template\shell\FillTemplate\command",
        "",
        fill_template_command,
    )
    _set_value(
        winreg.HKEY_CURRENT_USER,
        r"Software\Classes\Fillable.Template\shell\open\command",
        "",
        open_template_command,
    )

    for ext in [".docx", ".pptx", ".pdf", ".txt", ".json"]:
        base = fr"Software\Classes\SystemFileAssociations\{ext}\shell\FillableFillTemplate"
        _set_value(winreg.HKEY_CURRENT_USER, base, "", "Fill Template (Codex)")
        _set_value(winreg.HKEY_CURRENT_USER, base, "Icon", "imageres.dll,-5302")
        _set_value(winreg.HKEY_CURRENT_USER, base + r"\command", "", fill_template_command)


def uninstall_context_menu() -> None:
    for ext in [".docx", ".pptx", ".pdf"]:
        _delete_tree(
            winreg.HKEY_CURRENT_USER,
            fr"Software\Classes\SystemFileAssociations\{ext}\shell\FillableGenerateTemplate",
        )

    for ext in [".docx", ".pptx", ".pdf", ".txt", ".json"]:
        _delete_tree(
            winreg.HKEY_CURRENT_USER,
            fr"Software\Classes\SystemFileAssociations\{ext}\shell\FillableFillTemplate",
        )

    _delete_tree(winreg.HKEY_CURRENT_USER, r"Software\Classes\Fillable.Template\shell\FillTemplate\command")
    _delete_tree(winreg.HKEY_CURRENT_USER, r"Software\Classes\Fillable.Template\shell\FillTemplate")
    _delete_tree(winreg.HKEY_CURRENT_USER, r"Software\Classes\Fillable.Template\shell\open\command")
    _delete_tree(winreg.HKEY_CURRENT_USER, r"Software\Classes\Fillable.Template\shell\open")
    _delete_tree(winreg.HKEY_CURRENT_USER, r"Software\Classes\Fillable.Template\shell")
    _delete_tree(winreg.HKEY_CURRENT_USER, r"Software\Classes\Fillable.Template")
    _delete_tree(winreg.HKEY_CURRENT_USER, r"Software\Classes\.fillable.json")


def ensure_context_menu_installed_once(
    marker_dir: Path,
    log: Callable[[str], None] | None = None,
) -> None:
    marker_path = marker_dir / "context_menu_installed.marker"
    if marker_path.exists():
        return
    install_context_menu()
    marker_dir.mkdir(parents=True, exist_ok=True)
    marker_path.write_text("installed", encoding="utf-8")
    if log:
        log("Context menu installed for current user.")
