from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app.context_menu import uninstall_context_menu


def main() -> int:
    uninstall_context_menu()
    print("Context menu removed for current user.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
