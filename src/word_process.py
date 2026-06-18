"""Microsoft Word process management for COM PDF conversion."""

from __future__ import annotations

import logging
import subprocess
import sys
import time
from pathlib import Path

logger = logging.getLogger(__name__)


def quit_all_word_applications(*, force_kill: bool = True) -> list[str]:
    """
    Close all running Word instances.

    Tries graceful COM Quit first, then optional taskkill for WINWORD.EXE.
    Returns human-readable status messages.
    """
    if sys.platform != "win32":
        return []

    messages: list[str] = []

    try:
        import pythoncom  # type: ignore[import-untyped]
        import win32com.client  # type: ignore[import-untyped]

        pythoncom.CoInitialize()
        try:
            closed = 0
            while closed < 20:
                try:
                    word = win32com.client.GetObject(Class="Word.Application")
                    word.DisplayAlerts = 0
                    word.Quit(SaveChanges=False)
                    closed += 1
                    time.sleep(0.2)
                except Exception:
                    break
            if closed:
                messages.append(f"Закрыто экземпляров Word через COM: {closed}")
        finally:
            pythoncom.CoUninitialize()
    except Exception as exc:
        messages.append(f"COM Quit не удался: {exc}")

    time.sleep(0.3)

    if force_kill:
        messages.extend(_taskkill_winword())

    time.sleep(0.5)
    return messages


def _taskkill_winword() -> list[str]:
    messages: list[str] = []
    try:
        result = subprocess.run(
            ["taskkill", "/IM", "WINWORD.EXE", "/F"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=15,
            check=False,
        )
        output = f"{result.stdout}\n{result.stderr}".lower()
        if result.returncode == 0:
            messages.append("Процесс WINWORD.EXE завершён через taskkill")
        elif "не найден" in output or "not found" in output:
            messages.append("Активный WINWORD.EXE не найден")
        else:
            messages.append(
                f"taskkill WINWORD.EXE: код {result.returncode}, {result.stderr.strip()}"
            )
    except Exception as exc:
        messages.append(f"taskkill WINWORD.EXE не удался: {exc}")
    return messages


def wait_for_file_unlock(path: Path, *, timeout_seconds: float = 5.0) -> bool:
    """Return True if file can be opened for append within timeout."""
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        try:
            with path.open("a+b"):
                return True
        except OSError:
            time.sleep(0.2)
    return False


def remove_file_with_retry(
    path: Path,
    *,
    max_attempts: int = 3,
    quit_word_on_failure: bool = True,
) -> None:
    """Remove a file, quitting Word between attempts if needed."""
    last_error: OSError | None = None

    for attempt in range(1, max_attempts + 1):
        try:
            if path.exists():
                path.unlink()
            return
        except OSError as exc:
            last_error = exc
            logger.warning(
                "Cannot delete %s (attempt %d/%d): %s",
                path,
                attempt,
                max_attempts,
                exc,
            )
            if quit_word_on_failure and attempt < max_attempts:
                quit_all_word_applications(force_kill=True)
                wait_for_file_unlock(path, timeout_seconds=3.0)

    if last_error is not None:
        raise last_error
