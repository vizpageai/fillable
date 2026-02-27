from __future__ import annotations

import winreg


def delete_tree(root, subkey: str) -> None:
    try:
        with winreg.OpenKey(root, subkey, 0, winreg.KEY_READ | winreg.KEY_WRITE) as key:
            while True:
                try:
                    child = winreg.EnumKey(key, 0)
                except OSError:
                    break
                delete_tree(root, subkey + "\\" + child)
    except FileNotFoundError:
        return

    try:
        winreg.DeleteKey(root, subkey)
    except FileNotFoundError:
        pass


def main() -> int:
    for ext in [".docx", ".pptx", ".pdf"]:
        delete_tree(
            winreg.HKEY_CURRENT_USER,
            fr"Software\Classes\SystemFileAssociations\{ext}\shell\FillableGenerateTemplate",
        )
    for ext in [".docx", ".pptx", ".pdf", ".txt", ".json"]:
        delete_tree(
            winreg.HKEY_CURRENT_USER,
            fr"Software\Classes\SystemFileAssociations\{ext}\shell\FillableFillTemplate",
        )

    delete_tree(winreg.HKEY_CURRENT_USER, r"Software\Classes\Fillable.Template\shell\FillTemplate\command")
    delete_tree(winreg.HKEY_CURRENT_USER, r"Software\Classes\Fillable.Template\shell\FillTemplate")
    delete_tree(winreg.HKEY_CURRENT_USER, r"Software\Classes\Fillable.Template\shell\open\command")
    delete_tree(winreg.HKEY_CURRENT_USER, r"Software\Classes\Fillable.Template\shell\open")
    delete_tree(winreg.HKEY_CURRENT_USER, r"Software\Classes\Fillable.Template\shell")
    delete_tree(winreg.HKEY_CURRENT_USER, r"Software\Classes\Fillable.Template")
    delete_tree(winreg.HKEY_CURRENT_USER, r"Software\Classes\.fillable.json")

    print("Context menu removed for current user.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
