from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app.context_menu import install_context_menu


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--exe", default="", help="Path to built FillableDOC.exe")
    parser.add_argument(
        "--no-prompt-instructions",
        action="store_true",
        help="Disable instruction prompt when using right-click Fill action.",
    )
    args = parser.parse_args()

    install_context_menu(
        exe_override=(args.exe.strip() or None),
        prompt_instructions=not args.no_prompt_instructions,
    )
    print("Context menu installed for current user.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
