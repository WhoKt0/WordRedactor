"""Main entry point for the commercial letter bot."""

from __future__ import annotations

import argparse
import re
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path

from src.config import Settings
from src.document_generator import DocumentGeneratorError, generate_docx
from src.error_report import GenerationErrorItem, write_generation_error_report
from src.excel_reader import ExcelReaderError, read_banks
from src.logger_setup import setup_logging
from src.manifest_report import ManifestRow, write_manifest_report
from src.models import (
    BankRow,
    PlaceholderContext,
    RowStatus,
    RunSummary,
    StatusReportRow,
)
from src.output_utils import clean_preview_output, relative_project_path
from src.pdf_converter import PdfConverter, PdfConverterError
from src.performance import RunPerformance
from src.preflight import PreflightRowResult, print_preflight_summary, run_preflight
from src.state_manager import (
    GenerationState,
    get_next_out_number,
    increment_generation_id,
    load_generation_state,
    mark_out_number_committed,
)
from src.template_selector import (
    TemplateNotFoundError,
    greeting_name_length,
    select_template_for_greeting,
)
from src.status_report import write_status_report

logger = None

INVALID_PATH_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

PDF_SUCCESS_STATUSES = {
    RowStatus.PDF_GENERATED.value,
    RowStatus.GENERATED_NOT_SENT.value,
    RowStatus.SENT.value,
}

FINAL_CONFIRMATION_PROMPT = (
    "Вы точно хотите создать финальные PDF и зафиксировать СВХИ "
    "после успешного PDF? Напишите YES: "
)


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


def _preflight_error_item(result: PreflightRowResult) -> GenerationErrorItem:
    return GenerationErrorItem(
        row_number=result.row.row_number,
        bank_name=result.row.bank_name,
        chair_full_name=result.row.chair_full_name,
        greeting_name=result.row.greeting_name,
        recipient_email=result.row.recipient_email,
        status=RowStatus.VALIDATION_FAILED.value,
        error_message=result.error_message,
        out_number=None,
        phase="preflight",
    )


def _generation_error_item(
    row: BankRow,
    *,
    status: str,
    error_message: str,
    out_number: int | None,
) -> GenerationErrorItem:
    return GenerationErrorItem(
        row_number=row.row_number,
        bank_name=row.bank_name,
        chair_full_name=row.chair_full_name,
        greeting_name=row.greeting_name,
        recipient_email=row.recipient_email,
        status=status,
        error_message=error_message,
        out_number=out_number,
        phase="generation",
    )


def _process_row(
    preflight_row: PreflightRowResult,
    *,
    settings: Settings,
    out_number: int,
    date_str: str,
    docx_dir: Path,
    pdf_dir: Path,
    pdf_converter: PdfConverter,
    summary: RunSummary,
    commit_numbers: bool,
    state_path: Path,
    state: GenerationState,
    project_root: Path,
    generation_id: int,
    mode: str,
    performance: RunPerformance,
) -> tuple[StatusReportRow, bool, GenerationErrorItem | None, ManifestRow | None]:
    """
    Process one preflight-validated Excel row.

    Returns (report_row, pdf_success, error_item, manifest_row).
    """
    row = preflight_row.row
    report = StatusReportRow(
        row_number=row.row_number,
        bank_name=row.bank_name,
        bank_legal_name=row.bank_legal_name,
        recipient_email=row.recipient_email,
        date=date_str,
        out_number=out_number,
    )
    summary.processed += 1

    logger.info(
        "Processing row %d | bank=%s | email=%s | out_number=%s",
        row.row_number,
        row.bank_name,
        row.recipient_email,
        out_number,
    )

    context = PlaceholderContext(
        out_number=out_number,
        date_str=date_str,
        bank_legal_name=row.bank_legal_name,
        mr_ms=preflight_row.mr_ms,
        chair_short_dative=row.chair_short_dative,
        greeting_word=preflight_row.greeting_word,
        greeting_name=preflight_row.safe_greeting_name,
    )

    docx_path = _build_docx_path(docx_dir, row, out_number)
    pdf_path = _build_pdf_path(pdf_dir, row)
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    greeting_final = preflight_row.safe_greeting_name
    templates_dir = settings.word_template_path.parent

    try:
        word_template_path = select_template_for_greeting(greeting_final, templates_dir)
    except TemplateNotFoundError as exc:
        report.status = RowStatus.FAILED.value
        report.error_message = str(exc)
        report.out_number = None
        summary.failed += 1
        logger.error(
            "Row %d (%s) template selection failed: %s",
            row.row_number,
            row.bank_name,
            exc,
        )
        return (
            report,
            False,
            _generation_error_item(
                row,
                status=report.status,
                error_message=str(exc),
                out_number=None,
            ),
            None,
        )

    template_name = word_template_path.name
    logger.info(
        "Greeting length=%d, template=%s",
        greeting_name_length(greeting_final),
        template_name,
    )

    try:
        docx_started = time.perf_counter()
        generate_docx(word_template_path, docx_path, context)
        performance.add_docx(time.perf_counter() - docx_started)
        report.docx_path = relative_project_path(project_root, docx_path)
        report.status = RowStatus.DOCX_GENERATED.value
        logger.info("DOCX created: %s", docx_path)
    except (DocumentGeneratorError, FileNotFoundError) as exc:
        report.status = RowStatus.FAILED.value
        report.error_message = f"ошибка создания DOCX: {exc}"
        report.out_number = None
        summary.failed += 1
        logger.error("Row %d (%s) DOCX failed: %s", row.row_number, row.bank_name, exc)
        return (
            report,
            False,
            _generation_error_item(
                row,
                status=report.status,
                error_message=report.error_message,
                out_number=None,
            ),
            None,
        )
    except Exception as exc:
        report.status = RowStatus.FAILED.value
        report.error_message = f"ошибка создания DOCX: {exc}"
        report.out_number = None
        summary.failed += 1
        logger.error(
            "Row %d (%s) unexpected DOCX error: %s\n%s",
            row.row_number,
            row.bank_name,
            exc,
            traceback.format_exc(),
        )
        return (
            report,
            False,
            _generation_error_item(
                row,
                status=report.status,
                error_message=report.error_message,
                out_number=None,
            ),
            None,
        )

    try:
        pdf_converter.convert(docx_path, pdf_path)
        if not pdf_path.exists() or pdf_path.stat().st_size <= 0:
            raise PdfConverterError(
                f"PDF не создан в текущем запуске: {pdf_path.resolve()}"
            )
        report.pdf_path = relative_project_path(project_root, pdf_path)
        report.status = RowStatus.PDF_GENERATED.value
        logger.info("PDF created in current run: %s", pdf_path)

        if commit_numbers:
            mark_out_number_committed(state_path, state, out_number)
    except PdfConverterError as exc:
        report.status = RowStatus.PDF_FAILED.value
        report.error_message = f"DOCX создан, но PDF не создан: {exc}"
        report.out_number = None
        summary.failed += 1
        logger.error("Row %d (%s) PDF failed: %s", row.row_number, row.bank_name, exc)
        return (
            report,
            False,
            _generation_error_item(
                row,
                status=report.status,
                error_message=report.error_message,
                out_number=None,
            ),
            None,
        )
    except Exception as exc:
        report.status = RowStatus.PDF_FAILED.value
        report.error_message = f"DOCX создан, но PDF не создан: {exc}"
        report.out_number = None
        summary.failed += 1
        logger.error(
            "Row %d (%s) unexpected PDF error: %s\n%s",
            row.row_number,
            row.bank_name,
            exc,
            traceback.format_exc(),
        )
        return (
            report,
            False,
            _generation_error_item(
                row,
                status=report.status,
                error_message=report.error_message,
                out_number=None,
            ),
            None,
        )

    manifest_row = ManifestRow(
        generation_id=generation_id,
        mode=mode,
        excel_row=row.row_number,
        bank_name=row.bank_name,
        bank_legal_name=row.bank_legal_name,
        email_bank_name=row.email_bank_name or row.bank_name,
        chair_full_name=row.chair_full_name,
        greeting_word=preflight_row.greeting_word,
        greeting_name_final=preflight_row.safe_greeting_name,
        template_used=template_name,
        recipient_email=row.recipient_email,
        cc_email=row.cc_email,
        bcc_email=row.bcc_email,
        out_number=out_number,
        letter_date=date_str,
        docx_path=report.docx_path,
        pdf_path=report.pdf_path,
        email_status="pending",
        created_at=created_at,
    )
    return report, True, None, manifest_row


def _build_status_reports_in_excel_order(
    all_rows: list[BankRow],
    preflight: list[PreflightRowResult],
    processed_reports: dict[int, StatusReportRow],
) -> list[StatusReportRow]:
    preflight_by_row = {result.row.row_number: result for result in preflight}
    reports: list[StatusReportRow] = []

    for bank_row in all_rows:
        preflight_result = preflight_by_row[bank_row.row_number]
        if not preflight_result.valid:
            reports.append(
                StatusReportRow(
                    row_number=bank_row.row_number,
                    bank_name=bank_row.bank_name,
                    bank_legal_name=bank_row.bank_legal_name,
                    recipient_email=bank_row.recipient_email,
                    status=RowStatus.VALIDATION_FAILED.value,
                    error_message=preflight_result.error_message,
                )
            )
        else:
            reports.append(processed_reports[bank_row.row_number])

    return reports


def _count_successful_pdfs(report_rows: list[StatusReportRow]) -> int:
    return sum(1 for row in report_rows if row.status in PDF_SUCCESS_STATUSES)


def _print_finish_banner(
    *,
    commit_numbers: bool,
    state: GenerationState,
    initial_committed: int,
    successful_pdfs: int,
    preflight_errors: int,
    generation_errors: int,
    report_path: Path,
    error_report_path: Path | None,
    manifest_path: Path | None,
    project_root: Path,
    performance: RunPerformance,
) -> None:
    mode_label = "FINAL / COMMIT" if commit_numbers else "PREVIEW"
    next_number = get_next_out_number(state)

    print()
    print("Генерация завершена.")
    print(f"Режим: {mode_label}")

    if commit_numbers:
        print(f"Успешно создано PDF и зафиксировано СВХИ: {successful_pdfs}")
    else:
        print(f"Успешно создано PDF: {successful_pdfs}")

    print(f"Пропущено строк из-за ошибок preflight: {preflight_errors}")
    print(f"Ошибок во время генерации: {generation_errors}")

    if commit_numbers:
        print(f"Последний зафиксированный СВХИ: {state.last_committed_out_number}")
        print(f"Следующий финальный СВХИ: {next_number}")
    else:
        print(f"Последний зафиксированный СВХИ НЕ изменён: {initial_committed}")
        print(f"Следующий финальный СВХИ остаётся: {next_number}")

    if manifest_path is not None:
        print(f"Manifest: {relative_project_path(project_root, manifest_path)}")
    else:
        print("Manifest не создан: успешных PDF нет.")

    print(f"Excel status report: {relative_project_path(project_root, report_path)}")
    if error_report_path is not None:
        print(f"Файл ошибок: {relative_project_path(project_root, error_report_path)}")
    else:
        print("Ошибок нет.")

    performance.print_summary()


def _confirm_final_generation() -> bool:
    answer = input(FINAL_CONFIRMATION_PROMPT).strip()
    if answer != "YES":
        print("Генерация отменена.")
        return False
    return True


def run(
    project_root: Path | None = None,
    *,
    commit_numbers: bool = False,
    skip_confirmation: bool = False,
) -> int:
    """Run the bot. Returns process exit code."""
    global logger

    root = project_root or Path(__file__).resolve().parent.parent
    settings = Settings(root)
    state_path = root / "state" / "generation_state.json"

    if commit_numbers:
        docx_dir = settings.output_docx_dir
        pdf_dir = settings.output_pdf_dir
        reports_dir = settings.reports_dir
        mode = "FINAL"
    else:
        docx_dir = root / "output" / "preview" / "docx"
        pdf_dir = root / "output" / "preview" / "pdf"
        reports_dir = root / "output" / "preview" / "reports"
        mode = "PREVIEW"

    docx_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir.mkdir(parents=True, exist_ok=True)
    reports_dir.mkdir(parents=True, exist_ok=True)
    (root / "logs").mkdir(parents=True, exist_ok=True)
    state_path.parent.mkdir(parents=True, exist_ok=True)

    state = load_generation_state(state_path, settings.env.start_out_number)
    initial_committed = state.last_committed_out_number

    logger = setup_logging(settings.logging.level, settings.log_file_path)
    logger.info("=== RenSer letter bot started ===")
    logger.info("Project root: %s", root)

    try:
        banks = read_banks(settings.excel_file_path)
    except ExcelReaderError as exc:
        logger.error("Failed to read Excel: %s", exc)
        print(f"Ошибка чтения Excel: {exc}")
        return 1

    if not banks:
        logger.warning("No data rows in Excel file")
        print("В Excel нет строк для обработки.")
        return 0

    if not settings.word_template_path.exists():
        logger.error("Word template not found: %s", settings.word_template_path)
        print(f"Шаблон Word не найден: {settings.word_template_path}")
        return 1

    preflight = run_preflight(banks)
    print_preflight_summary(
        commit_numbers=commit_numbers,
        preflight=preflight,
        state=state,
    )

    if commit_numbers:
        if preflight.valid_count == 0:
            print("Нет валидных строк. Финальная генерация не запущена.")
            return 1
        if not skip_confirmation and not _confirm_final_generation():
            return 0
    else:
        warnings = clean_preview_output([docx_dir, pdf_dir, reports_dir])
        for warning in warnings:
            print(f"Предупреждение: {warning}")
            logger.warning(warning)

    generation_id = increment_generation_id(state_path, state)

    logger.info(
        "mode=%s | dry_run=%s | stop_on_error=%s | generation_id=%d",
        mode,
        settings.app.dry_run,
        settings.app.stop_on_error,
        generation_id,
    )
    logger.info(
        "Preflight: total=%d valid=%d invalid=%d",
        preflight.total,
        preflight.valid_count,
        preflight.invalid_count,
    )

    pdf_converter: PdfConverter | None = None
    performance = RunPerformance()
    performance.mark_started()

    try:
        pdf_converter = PdfConverter(settings.pdf.converter, performance=performance)
    except PdfConverterError as exc:
        logger.error("%s", exc)
        print(str(exc))
        return 1

    summary = RunSummary()
    processed_reports: dict[int, StatusReportRow] = {}
    error_items: list[GenerationErrorItem] = [
        _preflight_error_item(result) for result in preflight.invalid_rows
    ]
    manifest_rows: list[ManifestRow] = []
    current_out_number = get_next_out_number(state)
    date_str = settings.format_date()
    logger.info(
        "Дата писем для этого запуска: %s (формат из config.yaml date_format; "
        "в Word-шаблоне используйте {{TODAY}} или {{DATE}})",
        date_str,
    )

    try:
        valid_rows = preflight.valid_rows
        for index, preflight_row in enumerate(valid_rows):
            report, pdf_success, error_item, manifest_row = _process_row(
                preflight_row,
                settings=settings,
                out_number=current_out_number,
                date_str=date_str,
                docx_dir=docx_dir,
                pdf_dir=pdf_dir,
                pdf_converter=pdf_converter,
                summary=summary,
                commit_numbers=commit_numbers,
                state_path=state_path,
                state=state,
                project_root=root,
                generation_id=generation_id,
                mode=mode,
                performance=performance,
            )
            processed_reports[preflight_row.row.row_number] = report
            if error_item is not None:
                error_items.append(error_item)
            if manifest_row is not None:
                manifest_rows.append(manifest_row)

            if pdf_success:
                current_out_number += 1
            elif settings.app.stop_on_error:
                logger.error("Stopping due to stop_on_error=true")
                break
    finally:
        if pdf_converter is not None:
            pdf_converter.shutdown()

    report_rows = _build_status_reports_in_excel_order(
        banks,
        preflight.results,
        processed_reports,
    )
    report_path = write_status_report(report_rows, reports_dir)
    logger.info("Status report saved: %s", report_path)

    error_report_path = write_generation_error_report(
        error_items,
        reports_dir,
        generation_id,
        mode,
    )
    if error_report_path:
        logger.info("Error report saved: %s", error_report_path)

    manifest_path = write_manifest_report(
        manifest_rows,
        reports_dir,
        generation_id,
        mode,
    )
    if manifest_path:
        logger.info("Manifest saved: %s", manifest_path)

    successful_pdfs = _count_successful_pdfs(report_rows)
    generation_errors = sum(
        1 for item in error_items if item.phase == "generation"
    )

    _print_finish_banner(
        commit_numbers=commit_numbers,
        state=state,
        initial_committed=initial_committed,
        successful_pdfs=successful_pdfs,
        preflight_errors=preflight.invalid_count,
        generation_errors=generation_errors,
        report_path=report_path,
        error_report_path=error_report_path,
        manifest_path=manifest_path,
        project_root=root,
        performance=performance,
    )

    logger.info(
        "=== Run finished | processed=%d pdf_ok=%d failed=%d skipped=%d ===",
        summary.processed,
        summary.processed - summary.failed,
        summary.failed,
        summary.skipped,
    )
    return 0 if generation_errors == 0 else 1


def _build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="RenSer letter bot")
    mode_group = parser.add_mutually_exclusive_group()
    mode_group.add_argument(
        "--preview",
        action="store_true",
        help="Preview mode (default): outgoing numbers are not committed",
    )
    mode_group.add_argument(
        "--commit-numbers",
        action="store_true",
        help="Final mode: commit outgoing numbers after each successful PDF",
    )
    parser.add_argument(
        "--yes",
        action="store_true",
        help="Skip confirmation prompt for --commit-numbers",
    )
    return parser


def main(argv: list[str] | None = None) -> None:
    root = Path(__file__).resolve().parent.parent
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))

    parser = _build_arg_parser()
    args = parser.parse_args(argv)

    if args.preview and args.commit_numbers:
        print("Нельзя одновременно указывать --preview и --commit-numbers.")
        sys.exit(2)

    commit_numbers = args.commit_numbers
    sys.exit(
        run(
            root,
            commit_numbers=commit_numbers,
            skip_confirmation=commit_numbers and args.yes,
        )
    )


if __name__ == "__main__":
    main()
