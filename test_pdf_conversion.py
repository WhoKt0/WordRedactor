"""Quick PDF conversion test on an existing DOCX file (no Excel / no letter generation)."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.pdf_converter import PdfConverter, PdfConverterError


def _pick_docx() -> Path | None:
    preview_dir = ROOT / "output" / "preview" / "docx"
    if not preview_dir.exists():
        return None

    docx_files = sorted(preview_dir.glob("*.docx"), key=lambda path: path.stat().st_mtime)
    return docx_files[0] if docx_files else None


def main() -> int:
    docx_path = _pick_docx()
    if docx_path is None:
        print("Не найден DOCX в output/preview/docx/. Сначала запустите python preview.py")
        return 1

    pdf_path = ROOT / "output" / "preview" / "pdf" / f"_test_{docx_path.stem}.pdf"
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    print("=== PDF conversion test ===")
    print(f"DOCX: {docx_path.resolve()}")
    print(f"DOCX exists: {docx_path.exists()}")
    if docx_path.exists():
        print(f"DOCX size: {docx_path.stat().st_size} bytes")
    print(f"PDF target: {pdf_path.resolve()}")
    print()

    converter = PdfConverter("word_com")
    try:
        result = converter.convert(docx_path, pdf_path)
    except PdfConverterError as exc:
        print("PDF conversion FAILED")
        print(str(exc))
        return 1
    finally:
        converter.shutdown()

    print("PDF conversion OK")
    print(f"PDF: {result.resolve()}")
    print(f"PDF size: {result.stat().st_size} bytes")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
