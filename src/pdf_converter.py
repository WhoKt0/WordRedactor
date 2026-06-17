"""DOCX to PDF conversion via Microsoft Word COM (Windows)."""

from __future__ import annotations

import logging
import sys
from pathlib import Path

logger = logging.getLogger(__name__)

WORD_COM_UNAVAILABLE_MSG = (
    "Microsoft Word COM converter is unavailable. "
    "Install Microsoft Word or switch converter in config."
)


class PdfConverterError(Exception):
    """Raised when PDF conversion fails."""


class PdfConverter:
    """Convert DOCX files to PDF using Microsoft Word automation."""

    def __init__(self, converter_name: str = "word_com") -> None:
        self.converter_name = converter_name
        if converter_name != "word_com":
            raise PdfConverterError(
                f"Unsupported converter: {converter_name!r}. Only 'word_com' is supported."
            )
        if sys.platform != "win32":
            raise PdfConverterError(
                f"{WORD_COM_UNAVAILABLE_MSG} (non-Windows platform detected)"
            )

    def convert(self, docx_path: Path, pdf_path: Path) -> Path:
        """Convert a DOCX file to PDF."""
        docx_path = docx_path.resolve()
        pdf_path = pdf_path.resolve()
        pdf_path.parent.mkdir(parents=True, exist_ok=True)

        if not docx_path.exists():
            raise PdfConverterError(f"DOCX file not found: {docx_path}")

        try:
            import win32com.client  # type: ignore[import-untyped]
        except ImportError as exc:
            raise PdfConverterError(WORD_COM_UNAVAILABLE_MSG) from exc

        word = None
        doc = None
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0

            doc = word.Documents.Open(str(docx_path), ReadOnly=True)
            # wdFormatPDF = 17
            doc.ExportAsFixedFormat(
                OutputFileName=str(pdf_path),
                ExportFormat=17,
                OpenAfterExport=False,
                OptimizeFor=0,
                Range=0,
                Item=7,
                IncludeDocProps=True,
                KeepIRM=True,
                CreateBookmarks=0,
                DocStructureTags=True,
                BitmapMissingFonts=True,
                UseISO19005_1=False,
            )
            logger.info("PDF created: %s", pdf_path)
            return pdf_path
        except Exception as exc:
            if "Word.Application" in str(exc) or "com_error" in type(exc).__name__.lower():
                raise PdfConverterError(WORD_COM_UNAVAILABLE_MSG) from exc
            raise PdfConverterError(f"PDF conversion failed: {exc}") from exc
        finally:
            if doc is not None:
                try:
                    doc.Close(False)
                except Exception:
                    logger.debug("Failed to close Word document", exc_info=True)
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    logger.debug("Failed to quit Word application", exc_info=True)
