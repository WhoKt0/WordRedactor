"""Main entry point for the commercial letter bot."""

from __future__ import annotations

import re
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path

from src.config import Settings
from src.document_generator import (
    DocumentGeneratorError,
    UnreplacedPlaceholdersError,
    generate_docx,
)
from src.email_sender import EmailSender, EmailSenderError
from src.excel_reader import ExcelReaderError, read_banks
from src.logger_setup import setup_logging
from src.models import (
    BankRow,
    PlaceholderContext,
    RowStatus,
    RunSummary,
    StatusReportRow,
    gender_placeholders,
)
from src.pdf_converter import PdfConverter, PdfConverterError
from src.status_report import write_status_report
from src.validators import validate_bank_row

logger = None

INVALID_PATH_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def _sanitize_filename(name: str, max_len: int = 80) -> str:
    cleaned = INVALID_PATH_CHARS.sub("_", name).strip(" .")
    return cleaned[:max_len] or "document"


def _build_docx_path(output_dir: Path, row: BankRow, out_number: int) -> Path:
    base = _sanitize_filename(f"{out_number}_{row.bank_name}")
    return output_dir / f"{base}.docx"


def _build_pdf_path(output_dir: Path, row: BankRow) -> Path:
    filename = row.pdf_filename.strip()
    if not filename.lower().endswith(".pdf"):
        filename += ".pdf"
    return output_dir / _sanitize_filename(filename, max_len=120)


def _load_email_template(path: Path) -> str:
    if not path.exists():
        raise FileNotFoundError(f"Email template not found: {path}")
    return path.read_text(encoding="utf-8")


def _process_row(
    row: BankRow,
    *,
    settings: Settings,
    out_number: int,
    date_str: str,
    pdf_converter: PdfConverter,
    email_sender: EmailSender,
    email_template: str,
    summary: RunSummary,
) -> tuple[StatusReportRow, bool, bool]:
    """
    Process one Excel row.

    Returns (report_row, success, should_stop).
    """
    report = StatusReportRow(
        row_number=row.row_number,
        bank_name=row.bank_name,
        bank_legal_name=row.bank_legal_name,
        recipient_email=row.recipient_email,
        date=date_str,
    )
    summary.processed += 1

    validation = validate_bank_row(row)
    if not validation.valid:
        report.status = RowStatus.VALIDATION_FAILED.value
        report.error_message = validation.error_message
        summary.skipped += 1
        logger.warning(
            "Row %d (%s): validation failed — %s",
            row.row_number,
            row.bank_name,
            validation.error_message,
        )
        return report, False, False

    report.out_number = out_number
    logger.info(
        "Processing row %d | bank=%s | email=%s | out_number=%s",
        row.row_number,
        row.bank_name,
        row.recipient_email,
        out_number,
    )

    try:
        mr_ms, greeting_word = gender_placeholders(row.gender)
    except ValueError as exc:
        report.status = RowStatus.VALIDATION_FAILED.value
        report.error_message = str(exc)
        summary.failed += 1
        return report, False, settings.app.stop_on_error

    context = PlaceholderContext(
        out_number=out_number,
        date_str=date_str,
        bank_legal_name=row.bank_legal_name,
        mr_ms=mr_ms,
        chair_short_dative=row.chair_short_dative,
        greeting_word=greeting_word,
        greeting_name=row.greeting_name,
    )

    docx_path = _build_docx_path(settings.output_docx_dir, row, out_number)
    pdf_path = _build_pdf_path(settings.output_pdf_dir, row)

    try:
        generate_docx(settings.word_template_path, docx_path, context)
        report.docx_path = str(docx_path)
        report.status = RowStatus.DOCX_GENERATED.value
        logger.info("DOCX created: %s", docx_path)

        pdf_converter.convert(docx_path, pdf_path)
        report.pdf_path = str(pdf_path)
        report.status = RowStatus.PDF_GENERATED.value
        logger.info("PDF created: %s", pdf_path)

        email_vars = {
            "{{BANK_NAME}}": row.email_bank_name or row.bank_name,
            "{{BANK_LEGAL_NAME}}": row.bank_legal_name,
            "BANK_NAME": row.email_bank_name or row.bank_name,
            "BANK_LEGAL_NAME": row.bank_legal_name,
        }
        subject = email_sender.render_template(
            settings.email.subject_template, email_vars
        )
        body = email_sender.render_template(email_template, email_vars)

        if settings.app.dry_run:
            logger.info(
                "DRY_RUN email would be sent to %s | cc=%s | bcc=%s | subject=%s | pdf=%s",
                row.recipient_email,
                row.cc_email or "-",
                row.bcc_email or "-",
                subject,
                pdf_path.name,
            )
            report.status = RowStatus.GENERATED_NOT_SENT.value
            summary.dry_run += 1
            return report, True, False

        email_sender.send(
            to_email=row.recipient_email,
            cc_email=row.cc_email,
            bcc_email=row.bcc_email,
            subject=subject,
            body=body,
            pdf_path=pdf_path,
            bank_name_for_log=row.bank_name,
        )
        report.status = RowStatus.SENT.value
        report.sent_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        summary.sent += 1
        return report, True, False

    except (DocumentGeneratorError, PdfConverterError, EmailSenderError, FileNotFoundError) as exc:
        report.status = RowStatus.FAILED.value
        report.error_message = str(exc)
        summary.failed += 1
        logger.error(
            "Row %d (%s) failed: %s",
            row.row_number,
            row.bank_name,
            exc,
        )
        return report, False, settings.app.stop_on_error
    except Exception as exc:
        report.status = RowStatus.FAILED.value
        report.error_message = str(exc)
        summary.failed += 1
        logger.error(
            "Row %d (%s) unexpected error: %s\n%s",
            row.row_number,
            row.bank_name,
            exc,
            traceback.format_exc(),
        )
        return report, False, settings.app.stop_on_error


def run(project_root: Path | None = None) -> int:
    """Run the bot. Returns process exit code."""
    global logger

    root = project_root or Path(__file__).resolve().parent.parent
    settings = Settings(root)

    settings.output_docx_dir.mkdir(parents=True, exist_ok=True)
    settings.output_pdf_dir.mkdir(parents=True, exist_ok=True)
    settings.reports_dir.mkdir(parents=True, exist_ok=True)
    (root / "logs").mkdir(parents=True, exist_ok=True)

    logger = setup_logging(settings.logging.level, settings.log_file_path)
    logger.info("=== RenSer letter bot started ===")
    logger.info("Project root: %s", root)
    logger.info("dry_run=%s | stop_on_error=%s", settings.app.dry_run, settings.app.stop_on_error)

    summary = RunSummary()
    report_rows: list[StatusReportRow] = []
    current_out_number = settings.env.start_out_number
    date_str = settings.format_date()

    try:
        banks = read_banks(settings.excel_file_path)
    except ExcelReaderError as exc:
        logger.error("Failed to read Excel: %s", exc)
        return 1

    logger.info("Excel rows to process: %d", len(banks))
    if not banks:
        logger.warning("No data rows in Excel file")
        return 0

    if not settings.word_template_path.exists():
        logger.error("Word template not found: %s", settings.word_template_path)
        return 1

    try:
        email_template = _load_email_template(settings.email_template_path)
        pdf_converter = PdfConverter(settings.pdf.converter)
        email_sender = EmailSender(settings.env)
    except (PdfConverterError, FileNotFoundError) as exc:
        logger.error("%s", exc)
        return 1

    for index, row in enumerate(banks):
        report, success, should_stop = _process_row(
            row,
            settings=settings,
            out_number=current_out_number,
            date_str=date_str,
            pdf_converter=pdf_converter,
            email_sender=email_sender,
            email_template=email_template,
            summary=summary,
        )
        report_rows.append(report)

        if success:
            current_out_number += 1
            if (
                not settings.app.dry_run
                and index < len(banks) - 1
                and settings.app.delay_between_emails_seconds > 0
            ):
                logger.info(
                    "Waiting %d seconds before next email...",
                    settings.app.delay_between_emails_seconds,
                )
                time.sleep(settings.app.delay_between_emails_seconds)
        elif should_stop:
            logger.error("Stopping due to stop_on_error=true")
            break

    report_path = write_status_report(report_rows, settings.reports_dir)
    logger.info("Status report saved: %s", report_path)
    logger.info(
        "=== Run finished | processed=%d sent=%d failed=%d skipped=%d dry_run=%d ===",
        summary.processed,
        summary.sent,
        summary.failed,
        summary.skipped,
        summary.dry_run,
    )
    return 0 if summary.failed == 0 else 1


def main() -> None:
    root = Path(__file__).resolve().parent.parent
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    sys.exit(run(root))


if __name__ == "__main__":
    main()
