"""Plain and HTML email body composition with signature."""

from __future__ import annotations

import html
import re
from pathlib import Path

SIGNATURE_COLOR = "#1f4e79"
BODY_FONT = "Calibri, Arial, sans-serif"
BODY_FONT_SIZE = "11pt"


def load_signature_text(path: Path) -> str:
    if not path.exists():
        raise FileNotFoundError(f"Email signature text not found: {path}")
    return path.read_text(encoding="utf-8").strip()


def load_signature_html(path: Path) -> str:
    if not path.exists():
        raise FileNotFoundError(f"Email signature HTML not found: {path}")
    return path.read_text(encoding="utf-8").strip()


def compose_plain_body(main_text: str, signature_text: str) -> str:
    main = main_text.strip()
    signature = signature_text.strip()
    if not signature:
        return main
    if not main:
        return signature
    return f"{main}\n\n{signature}"


def _main_text_to_html(main_text: str) -> str:
    main = main_text.strip()
    if not main:
        return ""

    paragraphs = re.split(r"\n\s*\n", main)
    blocks: list[str] = []
    for paragraph in paragraphs:
        paragraph = paragraph.strip()
        if not paragraph:
            continue
        lines = [html.escape(line) for line in paragraph.splitlines()]
        blocks.append(
            '<p style="margin: 0 0 10px 0; font-family: '
            f"{BODY_FONT}; font-size: {BODY_FONT_SIZE}; color: #000000;\">"
            + "<br>".join(lines)
            + "</p>"
        )
    return "\n".join(blocks)


def compose_html_body(main_text: str, signature_html: str) -> str:
    main_html = _main_text_to_html(main_text)
    signature = signature_html.strip()
    parts = [
        "<!DOCTYPE html>",
        "<html>",
        '<body style="margin: 0; padding: 0;">',
        f'<div style="font-family: {BODY_FONT}; font-size: {BODY_FONT_SIZE}; color: #000000;">',
        main_html,
        "</div>",
    ]
    if signature:
        parts.append(signature)
    parts.extend(["</body>", "</html>"])
    return "\n".join(parts)
