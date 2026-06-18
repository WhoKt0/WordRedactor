"""Final generation: outgoing numbers are committed after each successful PDF."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.main import run


def main() -> int:
    return run(ROOT, commit_numbers=True, skip_confirmation=False)


if __name__ == "__main__":
    raise SystemExit(main())
