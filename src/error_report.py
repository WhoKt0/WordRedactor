"""Human-readable TXT error reports per generation run."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from src.models import RowStatus


@dataclass
class GenerationErrorItem:
    row_number: int
    bank_name: str
    chair_full_name: str
    greeting_name: str
    recipient_email: str
    status: str
    error_message: str
    out_number: int | None
    phase: str = "preflight"


def _build_cursor_prompt(item: GenerationErrorItem) -> str:
    if "обращение" in item.error_message.lower() or "greeting" in item.error_message.lower():
        return (
            f"Проверь строку {item.row_number} в Excel. "
            f"В поле greeting_name указано только одно слово: «{item.greeting_name}». "
            f"В chair_full_name указано: «{item.chair_full_name}». "
            "Нужно уточнить ФИО или исправить greeting_name, потому что после "
            "«Уважаемый/Уважаемая» должно быть минимум два слова."
        )
    return (
        f"Проверь строку {item.row_number} в Excel для банка «{item.bank_name}». "
        f"Ошибка: {item.error_message}"
    )


def _document_note(item: GenerationErrorItem) -> str:
    if item.status == RowStatus.VALIDATION_FAILED.value:
        return "Документ не создан (ошибка preflight)."
    if item.status == RowStatus.FAILED.value:
        return "Документ не создан."
    if item.status == RowStatus.PDF_FAILED.value:
        return "DOCX создан, но PDF не создан."
    if item.status == RowStatus.EMAIL_FAILED.value:
        return "PDF создан, но email не отправлен."
    return "Документ не создан."


def _out_number_note(item: GenerationErrorItem) -> str:
    if item.out_number is None:
        return "Исходящий номер не был использован."
    return f"Исходящий номер {item.out_number} был использован (PDF создан)."


def write_generation_error_report(
    errors: list[GenerationErrorItem],
    reports_dir: Path,
    generation_id: int,
    mode: str,
) -> Path | None:
    """Write TXT error report. Returns path if created, else None."""
    if not errors:
        return None

    reports_dir.mkdir(parents=True, exist_ok=True)
    report_path = reports_dir / f"ген_{generation_id}_ошибки.txt"

    mode_upper = mode.upper()
    lines: list[str] = [
        f"ГЕНЕРАЦИЯ №{generation_id} — ОШИБКИ",
        f"Режим: {mode_upper}",
    ]
    if mode_upper == "PREVIEW":
        lines.append("Исходящие номера не были зафиксированы.")
    else:
        lines.append("Исходящие номера фиксируются только после успешного создания PDF.")

    lines.extend(["", f"Всего ошибок: {len(errors)}", ""])

    for index, item in enumerate(errors, start=1):
        phase_label = "preflight" if item.phase == "preflight" else "генерация"
        lines.extend(
            [
                f"{index}) Строка Excel: {item.row_number}",
                f"Банк: {item.bank_name}",
                f"ФИО председателя: {item.chair_full_name}",
                f"Greeting name из Excel: {item.greeting_name}",
                f"Email: {item.recipient_email}",
                f"Этап: {phase_label}",
                f"Ошибка: {item.error_message}",
                _document_note(item),
                _out_number_note(item),
                "",
                "Prompt для Cursor:",
                _build_cursor_prompt(item),
                "",
            ]
        )

    report_path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")
    return report_path
