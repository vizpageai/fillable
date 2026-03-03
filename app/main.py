from __future__ import annotations

import sys
from pathlib import Path

from app.cli import build_parser, run_cli
from app.context_menu import ensure_context_menu_installed_once
from app.gui import FillableApp


def _try_install_context_menu_once() -> None:
    marker_dir = Path.home() / "AppData" / "Roaming" / "Fillable"
    try:
        ensure_context_menu_installed_once(marker_dir)
    except Exception:
        # Context menu should not block app launch.
        return


def main() -> int:
    if len(sys.argv) == 2 and not sys.argv[1].startswith("-"):
        _try_install_context_menu_once()
        raw = sys.argv[1]
        initial_generate = None
        initial_template = None
        if raw.lower().endswith(".fillable.json"):
            initial_template = raw
        else:
            initial_generate = raw
        app = FillableApp(
            initial_generate=initial_generate,
            initial_template=initial_template,
        )
        app.mainloop()
        return 0

    parser = build_parser()
    args = parser.parse_args()

    if (
        args.generate_template
        or args.import_template_file
        or args.fill_template
        or args.print_config
        or args.install_context_menu
        or args.uninstall_context_menu
    ):
        return run_cli(args)

    _try_install_context_menu_once()
    initial_generate = None
    initial_template = None
    app = FillableApp(initial_generate=initial_generate, initial_template=initial_template)
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
