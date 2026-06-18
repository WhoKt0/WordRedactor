"""DOCX to PDF conversion via Microsoft Word COM (Windows)."""

from __future__ import annotations

import logging
import os
import sys
import time
from pathlib import Path

from src.performance import RunPerformance
from src.word_process import quit_all_word_applications

logger = logging.getLogger(__name__)

WORD_COM_UNAVAILABLE_MSG = (
    "Microsoft Word COM converter is unavailable. "
    "Install Microsoft Word or switch converter in config."
)

WD_EXPORT_FORMAT_PDF = 17
WD_DO_NOT_SAVE_CHANGES = 0


class PdfConverterError(Exception):
    """Raised when PDF conversion fails."""


class PdfConverter:
    """
    Convert DOCX files to PDF using one reusable Word.Application per run.

    Pattern: open document -> export PDF -> close document.
    Word itself stays open until shutdown().
    """

    def __init__(
        self,
        converter_name: str = "word_com",
        *,
        performance: RunPerformance | None = None,
    ) -> None:
        self.converter_name = converter_name
        self._word: object | None = None
        self._com_initialized = False
        self._conversions_done = 0
        self._performance = performance

        if converter_name != "word_com":
            raise PdfConverterError(
                f"Unsupported converter: {converter_name!r}. Only 'word_com' is supported."
            )
        if sys.platform != "win32":
            raise PdfConverterError(
                f"{WORD_COM_UNAVAILABLE_MSG} (non-Windows platform detected)"
            )

    def shutdown(self) -> None:
        """Close the reused Word instance at the end of a generation run."""
        conversions = self._conversions_done
        self._destroy_word_instance()
        logger.info(
            "PdfConverter shutdown complete | conversions_in_this_run=%d",
            conversions,
        )

    def convert(self, docx_path: Path, pdf_path: Path) -> Path:
        """Convert a DOCX file to PDF with one automatic retry on failure."""
        docx_path = docx_path.resolve()
        pdf_path = pdf_path.resolve()
        pdf_path.parent.mkdir(parents=True, exist_ok=True)

        if not docx_path.exists():
            raise PdfConverterError(
                _build_error_message(
                    summary="DOCX file not found before PDF conversion",
                    docx_path=docx_path,
                    pdf_path=pdf_path,
                    attempt=1,
                )
            )

        last_error: Exception | None = None
        last_details = ""

        for attempt in range(1, 3):
            if attempt > 1:
                logger.warning(
                    "PDF conversion retry %d/2 | docx=%s",
                    attempt,
                    docx_path,
                )
                if self._performance is not None:
                    self._performance.add_pdf_retry()
                self._destroy_word_instance()
                quit_all_word_applications(force_kill=True)
                time.sleep(1)

            convert_started = time.perf_counter()
            try:
                result = self._convert_once(docx_path, pdf_path, attempt=attempt)
                if self._performance is not None:
                    self._performance.add_pdf_success(time.perf_counter() - convert_started)
                return result
            except Exception as exc:
                last_error = exc
                last_details = _format_exception_details(exc)
                logger.error(
                    "PDF conversion attempt %d failed | docx=%s | pdf=%s | %s",
                    attempt,
                    docx_path,
                    pdf_path,
                    last_details,
                    exc_info=True,
                )

        raise PdfConverterError(
            _build_error_message(
                summary="PDF conversion failed after 2 attempts",
                docx_path=docx_path,
                pdf_path=pdf_path,
                attempt=2,
                com_details=last_details,
                original_error=last_error,
            )
        ) from last_error

    def _ensure_com(self) -> None:
        if not self._com_initialized:
            import pythoncom  # type: ignore[import-untyped]

            pythoncom.CoInitialize()
            self._com_initialized = True

    def _get_word_app(self) -> object:
        import win32com.client  # type: ignore[import-untyped]

        self._ensure_com()
        if self._word is None:
            startup_started = time.perf_counter()
            self._word = win32com.client.DispatchEx("Word.Application")
            self._word.Visible = False
            self._word.DisplayAlerts = 0
            self._word.ScreenUpdating = False
            if self._performance is not None:
                self._performance.record_word_startup(time.perf_counter() - startup_started)
            logger.info("Started reusable Word COM instance (one per generation run)")
        return self._word

    def _destroy_word_instance(self) -> None:
        word = self._word
        self._word = None
        if word is not None:
            try:
                word.Quit(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
            except Exception as exc:
                logger.warning("Failed to quit Word application: %s", exc)
        if self._com_initialized:
            import pythoncom  # type: ignore[import-untyped]

            pythoncom.CoUninitialize()
            self._com_initialized = False

    def _convert_once(self, docx_path: Path, pdf_path: Path, *, attempt: int) -> Path:
        docx_com_path = _com_path(docx_path)
        pdf_com_path = _com_path(pdf_path)

        if docx_path.stat().st_size == 0:
            raise PdfConverterError(
                _build_error_message(
                    summary="DOCX file is empty before PDF conversion",
                    docx_path=docx_path,
                    pdf_path=pdf_path,
                    attempt=attempt,
                    extra={"docx_size_bytes": 0},
                )
            )

        _remove_stale_pdf(pdf_path)

        word = self._get_word_app()
        doc = None
        word_version: str | None = None

        try:
            word_version = _safe_word_version(word)
            logger.debug(
                "PDF convert | attempt=%d | docx=%s | pdf=%s | word=%s | reused_word=%s",
                attempt,
                docx_com_path,
                pdf_com_path,
                word_version,
                self._conversions_done > 0,
            )

            doc = word.Documents.Open(
                FileName=docx_com_path,
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=False,
            )
            doc.ExportAsFixedFormat(
                OutputFileName=pdf_com_path,
                ExportFormat=WD_EXPORT_FORMAT_PDF,
                OpenAfterExport=False,
            )
        except Exception as exc:
            raise PdfConverterError(
                _build_error_message(
                    summary="PDF conversion failed during Word COM export",
                    docx_path=docx_path,
                    pdf_path=pdf_path,
                    attempt=attempt,
                    com_details=_format_exception_details(exc),
                    original_error=exc,
                    word_version=word_version,
                )
            ) from exc
        finally:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
                except Exception as exc:
                    logger.warning("Failed to close Word document: %s", exc)

        _verify_pdf_created(pdf_path, docx_path=docx_path, attempt=attempt)
        self._conversions_done += 1
        logger.info("PDF created: %s", pdf_path)
        return pdf_path


def _com_path(path: Path) -> str:
    return os.path.normpath(str(path.resolve()))


def _safe_word_version(word: object | None) -> str | None:
    if word is None:
        return None
    try:
        return str(word.Version)
    except Exception as exc:
        logger.debug("Could not read Word version: %s", exc)
        return None


def _format_exception_details(exc: Exception) -> str:
    parts = [f"type={type(exc).__name__}", f"repr={exc!r}", f"str={exc}"]

    excepinfo = getattr(exc, "excepinfo", None)
    if excepinfo:
        parts.append(f"excepinfo={excepinfo}")

    hresult = getattr(exc, "hresult", None)
    if hresult is not None:
        parts.append(f"hresult={hresult}")

    strerror = getattr(exc, "strerror", None)
    if strerror:
        parts.append(f"strerror={strerror}")

    return " | ".join(parts)


def _build_error_message(
    *,
    summary: str,
    docx_path: Path,
    pdf_path: Path,
    attempt: int,
    com_details: str = "",
    original_error: Exception | None = None,
    word_version: str | None = None,
    extra: dict[str, object] | None = None,
) -> str:
    lines = [
        summary,
        f"attempt={attempt}",
        f"docx_path={docx_path}",
        f"docx_abs={docx_path.resolve()}",
        f"docx_exists={docx_path.exists()}",
        f"docx_size_bytes={docx_path.stat().st_size if docx_path.exists() else 'n/a'}",
        f"pdf_path={pdf_path}",
        f"pdf_abs={pdf_path.resolve()}",
        f"pdf_dir_exists={pdf_path.parent.exists()}",
        f"pdf_exists_after_attempt={pdf_path.exists()}",
    ]
    if word_version:
        lines.append(f"word_version={word_version}")
    if com_details:
        lines.append(f"com_error={com_details}")
    elif original_error is not None:
        lines.append(f"com_error={_format_exception_details(original_error)}")
    if extra:
        for key, value in extra.items():
            lines.append(f"{key}={value}")
    return "\n".join(lines)


def _remove_stale_pdf(pdf_path: Path) -> None:
    if not pdf_path.exists():
        return
    try:
        pdf_path.unlink()
    except OSError as exc:
        logger.debug("Could not remove existing PDF before conversion: %s | %s", pdf_path, exc)


def _verify_pdf_created(pdf_path: Path, *, docx_path: Path, attempt: int) -> None:
    if not pdf_path.exists():
        raise PdfConverterError(
            _build_error_message(
                summary="Word COM reported success but PDF file was not created",
                docx_path=docx_path,
                pdf_path=pdf_path,
                attempt=attempt,
            )
        )

    pdf_size = pdf_path.stat().st_size
    if pdf_size <= 0:
        raise PdfConverterError(
            _build_error_message(
                summary="PDF file was created but is empty",
                docx_path=docx_path,
                pdf_path=pdf_path,
                attempt=attempt,
                extra={"pdf_size_bytes": pdf_size},
            )
        )
