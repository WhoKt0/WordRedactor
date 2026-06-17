"""Row validation before document generation."""

from __future__ import annotations

import re
from dataclasses import dataclass

from src.models import BankRow, Gender

EMAIL_PATTERN = re.compile(r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$")
INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


@dataclass
class ValidationResult:
    valid: bool
    errors: list[str]

    @property
    def error_message(self) -> str:
        return "; ".join(self.errors)


def validate_bank_row(row: BankRow) -> ValidationResult:
    """Validate a single bank row before processing."""
    errors: list[str] = []

    if not row.bank_legal_name:
        errors.append("bank_legal_name is required")

    if not row.recipient_email:
        errors.append("recipient_email is required")
    elif not is_valid_email(row.recipient_email):
        errors.append(f"invalid recipient_email: {row.recipient_email!r}")

    if row.cc_email and not is_valid_email(row.cc_email):
        errors.append(f"invalid cc_email: {row.cc_email!r}")

    if row.bcc_email and not is_valid_email(row.bcc_email):
        errors.append(f"invalid bcc_email: {row.bcc_email!r}")

    if row.gender not in (Gender.MALE.value, Gender.FEMALE.value):
        errors.append(f"gender must be 'male' or 'female', got: {row.gender!r}")

    if not row.chair_short_dative:
        errors.append("chair_short_dative is required")

    if not row.greeting_name:
        errors.append("greeting_name is required")

    if not row.pdf_filename:
        errors.append("pdf_filename is required")
    elif INVALID_FILENAME_CHARS.search(row.pdf_filename):
        errors.append("pdf_filename contains forbidden characters")
    elif not row.pdf_filename.lower().endswith(".pdf"):
        errors.append("pdf_filename must end with .pdf")

    return ValidationResult(valid=len(errors) == 0, errors=errors)


def is_valid_email(email: str) -> bool:
    return bool(EMAIL_PATTERN.match(email.strip()))
