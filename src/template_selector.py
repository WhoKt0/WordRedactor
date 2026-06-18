"""Word template selection based on final greeting name length."""

from __future__ import annotations

from pathlib import Path

TEMPLATE_DEFAULT = "Template_word.docx"
TEMPLATE_LONG = "Template_word_greeting_long.docx"
TEMPLATE_SHORT = "Template_word_greeting_short.docx"

_GREETING_PLACEHOLDER = "{{GREETING_WORD}}"


class TemplateNotFoundError(FileNotFoundError):
    """Raised when a required greeting template file is missing."""


def greeting_name_length(greeting_name_final: str) -> int:
    """Return character length of normalized greeting name (spaces inside count)."""
    return len(" ".join(greeting_name_final.split()))


def select_template_for_greeting(
    greeting_name_final: str,
    templates_dir: Path,
) -> Path:
    """
    Pick Word template by final greeting name length.

    > 20 chars -> long template (3 fewer spaces before greeting)
    < 10 chars -> short template (3 more spaces before greeting)
    else -> default template
    """
    length = greeting_name_length(greeting_name_final)

    if length > 20:
        filename = TEMPLATE_LONG
        missing_hint = (
            f"Не найден шаблон templates/{TEMPLATE_LONG} для длинного обращения.\n"
            "Создайте копию Template_word.docx и уменьшите отступ перед обращением на 3 пробела."
        )
    elif length < 10:
        filename = TEMPLATE_SHORT
        missing_hint = (
            f"Не найден шаблон templates/{TEMPLATE_SHORT} для короткого обращения.\n"
            "Создайте копию Template_word.docx и увеличьте отступ перед обращением на 3 пробела."
        )
    else:
        filename = TEMPLATE_DEFAULT
        missing_hint = (
            f"Не найден шаблон templates/{TEMPLATE_DEFAULT}.\n"
            "Положите основной Word-шаблон в папку templates/."
        )

    template_path = templates_dir / filename
    if not template_path.exists():
        raise TemplateNotFoundError(missing_hint)
    return template_path
