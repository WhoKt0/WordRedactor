"""Двойной клик по этому файлу (или run.bat) — запуск бота."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def main() -> int:
    root = Path(__file__).resolve().parent
    venv_python = root / ".venv" / "Scripts" / "python.exe"

    if not venv_python.exists():
        print("Первый запуск: создаю .venv и ставлю зависимости...")
        subprocess.check_call([sys.executable, "-m", "venv", str(root / ".venv")])
        subprocess.check_call(
            [str(root / ".venv" / "Scripts" / "pip.exe"), "install", "-r", "requirements.txt"],
            cwd=root,
        )

    print("=== RenSer Letter Bot ===\n")
    result = subprocess.run(
        [str(venv_python), "-m", "src.main"],
        cwd=root,
    )
    input("\nНажмите Enter для выхода...")
    return result.returncode


if __name__ == "__main__":
    raise SystemExit(main())
