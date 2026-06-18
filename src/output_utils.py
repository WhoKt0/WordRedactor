"""Output directory helpers."""

from __future__ import annotations

import logging
import time
from pathlib import Path

from src.word_process import quit_all_word_applications

logger = logging.getLogger(__name__)


def _delete_preview_files(preview_output_dirs: list[Path]) -> list[str]:
    warnings: list[str] = []

    for directory in preview_output_dirs:
        if not directory.exists():
            continue

        for item in directory.iterdir():
            if item.name == ".gitkeep":
                continue
            if not item.is_file():
                continue
            try:
                item.unlink()
            except OSError as exc:
                warnings.append(f"Не удалось удалить {item}: {exc}")

    return warnings


def clean_preview_output(preview_output_dirs: list[Path]) -> list[str]:
    """
    Remove old files from preview output directories.

    First tries a fast delete. Only if files are locked, quits Word and retries.
    Keeps .gitkeep files. Returns list of warning messages.
    """
    warnings = _delete_preview_files(preview_output_dirs)
    if not warnings:
        return []

    message = (
        "Preview cleanup: файл заблокирован, выполняю fallback taskkill WINWORD"
    )
    print(message)
    logger.warning(message)

    for status in quit_all_word_applications(force_kill=True):
        logger.info("clean_preview_output: %s", status)

    time.sleep(0.3)
    warnings = _delete_preview_files(preview_output_dirs)
    for warning in warnings:
        logger.warning(warning)
    return warnings


def relative_project_path(project_root: Path, path: Path) -> str:
    """Return a path relative to the project root using forward slashes."""
    try:
        return path.resolve().relative_to(project_root.resolve()).as_posix()
    except ValueError:
        return path.as_posix()
