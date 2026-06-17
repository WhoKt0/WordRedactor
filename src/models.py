"""Data models for bank rows and processing results."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any


class Gender(str, Enum):
    MALE = "male"
    FEMALE = "female"


class RowStatus(str, Enum):
    VALIDATION_FAILED = "validation_failed"
    DOCX_GENERATED = "docx_generated"
    PDF_GENERATED = "pdf_generated"
    GENERATED_NOT_SENT = "generated_not_sent"
    SENT = "sent"
    FAILED = "failed"


@dataclass
class BankRow:
    """Single row from the Excel input file."""

    row_number: int
    bank_name: str
    bank_legal_name: str
    recipient_email: str
    cc_email: str
    bcc_email: str
    chair_full_name: str
    chair_short_dative: str
    gender: str
    greeting_name: str
    pdf_filename: str
    email_bank_name: str

    @classmethod
    def from_excel_dict(cls, row_number: int, data: dict[str, Any]) -> "BankRow":
        def _str(key: str) -> str:
            val = data.get(key)
            if val is None:
                return ""
            return str(val).strip()

        return cls(
            row_number=row_number,
            bank_name=_str("bank_name"),
            bank_legal_name=_str("bank_legal_name"),
            recipient_email=_str("recipient_email"),
            cc_email=_str("cc_email"),
            bcc_email=_str("bcc_email"),
            chair_full_name=_str("chair_full_name"),
            chair_short_dative=_str("chair_short_dative"),
            gender=_str("gender").lower(),
            greeting_name=_str("greeting_name"),
            pdf_filename=_str("pdf_filename"),
            email_bank_name=_str("email_bank_name"),
        )


@dataclass
class PlaceholderContext:
    """Values for Word template placeholders."""

    out_number: int
    date_str: str
    bank_legal_name: str
    mr_ms: str
    chair_short_dative: str
    greeting_word: str
    greeting_name: str

    def as_dict(self) -> dict[str, str]:
        return {
            "{{OUT_NUMBER}}": str(self.out_number),
            "{{DATE}}": self.date_str,
            "{{BANK_LEGAL_NAME}}": self.bank_legal_name,
            "{{MR_MS}}": self.mr_ms,
            "{{CHAIR_SHORT_DATIVE}}": self.chair_short_dative,
            "{{GREETING_WORD}}": self.greeting_word,
            "{{GREETING_NAME}}": self.greeting_name,
        }


@dataclass
class StatusReportRow:
    """One row in the status report."""

    row_number: int
    bank_name: str
    bank_legal_name: str
    recipient_email: str
    out_number: int | None = None
    date: str = ""
    docx_path: str = ""
    pdf_path: str = ""
    status: str = ""
    error_message: str = ""
    sent_at: str = ""


@dataclass
class RunSummary:
    """Aggregated run statistics."""

    processed: int = 0
    sent: int = 0
    failed: int = 0
    skipped: int = 0
    dry_run: int = 0

    def to_log_dict(self) -> dict[str, int]:
        return {
            "processed": self.processed,
            "sent": self.sent,
            "failed": self.failed,
            "skipped": self.skipped,
            "dry_run": self.dry_run,
        }


def gender_placeholders(gender: str) -> tuple[str, str]:
    """Return (MR_MS, GREETING_WORD) for male/female."""
    if gender == Gender.FEMALE.value:
        return "г-же", "Уважаемая"
    if gender == Gender.MALE.value:
        return "г-ну", "Уважаемый"
    raise ValueError(f"Invalid gender: {gender!r}")
