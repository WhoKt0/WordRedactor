"""Preflight validation of all Excel rows before generation."""

from __future__ import annotations

from dataclasses import dataclass

from src.models import BankRow, gender_placeholders
from src.name_utils import SERVICE_WORDS, GreetingNameError, build_greeting_name
from src.state_manager import GenerationState, get_next_out_number
from src.validators import validate_bank_row


@dataclass
class PreflightRowResult:
    row: BankRow
    valid: bool
    error_message: str
    safe_greeting_name: str = ""
    mr_ms: str = ""
    greeting_word: str = ""


@dataclass
class PreflightSummary:
    total: int
    valid_count: int
    invalid_count: int
    results: list[PreflightRowResult]

    @property
    def valid_rows(self) -> list[PreflightRowResult]:
        return [result for result in self.results if result.valid]

    @property
    def invalid_rows(self) -> list[PreflightRowResult]:
        return [result for result in self.results if not result.valid]


def _greeting_has_service_words(greeting_name: str) -> bool:
    return any(word.lower() in SERVICE_WORDS for word in greeting_name.split())


def validate_row_preflight(row: BankRow) -> PreflightRowResult:
    """Validate one row and prepare greeting placeholders when valid."""
    validation = validate_bank_row(row)
    if not validation.valid:
        return PreflightRowResult(
            row=row,
            valid=False,
            error_message=validation.error_message,
        )

    try:
        mr_ms, greeting_word = gender_placeholders(row.gender)
        safe_greeting_name = build_greeting_name(row.greeting_name, row.chair_full_name)
    except (ValueError, GreetingNameError) as exc:
        return PreflightRowResult(
            row=row,
            valid=False,
            error_message=str(exc),
        )

    if _greeting_has_service_words(safe_greeting_name):
        return PreflightRowResult(
            row=row,
            valid=False,
            error_message=(
                "итоговое обращение содержит служебное слово "
                f"({', '.join(sorted(SERVICE_WORDS))})"
            ),
        )

    return PreflightRowResult(
        row=row,
        valid=True,
        error_message="",
        safe_greeting_name=safe_greeting_name,
        mr_ms=mr_ms,
        greeting_word=greeting_word,
    )


def run_preflight(rows: list[BankRow]) -> PreflightSummary:
    """Validate all Excel rows before document generation."""
    results = [validate_row_preflight(row) for row in rows]
    valid_count = sum(1 for result in results if result.valid)
    return PreflightSummary(
        total=len(rows),
        valid_count=valid_count,
        invalid_count=len(rows) - valid_count,
        results=results,
    )


def _format_svhi_range(first_number: int, count: int) -> str:
    if count <= 0:
        return "—"
    if count == 1:
        return str(first_number)
    return f"{first_number}–{first_number + count - 1}"


def print_preflight_summary(
    *,
    commit_numbers: bool,
    preflight: PreflightSummary,
    state: GenerationState,
) -> None:
    """Print preflight summary to console."""
    mode_label = "FINAL / COMMIT" if commit_numbers else "PREVIEW"
    print(f"Режим: {mode_label}")
    print("Preflight validation:")
    print(f"Всего строк Excel: {preflight.total}")
    print(f"Валидных строк: {preflight.valid_count}")
    print(f"Ошибочных строк: {preflight.invalid_count}")
    print()

    first_number = get_next_out_number(state)

    if commit_numbers:
        print(f"Последний зафиксированный СВХИ: {state.last_committed_out_number}")
        print(
            "Будут использованы СВХИ: "
            f"{_format_svhi_range(first_number, preflight.valid_count)}"
        )
        print(
            "Всего будет зафиксировано номеров при успешном PDF: "
            f"{preflight.valid_count}"
        )
        print("СВХИ фиксируется только после успешного создания PDF.")
        if preflight.invalid_count:
            print("Ошибочные строки будут пропущены.")
    else:
        print(f"Первый preview-СВХИ: {first_number}")
        print("Preview-СВХИ будут показаны в документах, но НЕ будут зафиксированы.")
    print()
