"""Preview email delivery for the latest final manifest (no SMTP send)."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.email_pipeline import EmailPipeline, EmailPipelineError
from src.logger_setup import setup_logging
from src.config import Settings


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Preview email delivery from manifest")
    parser.add_argument(
        "--manifest",
        help="Path to final manifest (default: latest in output/reports/)",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)

    settings = Settings(ROOT)
    setup_logging(settings.logging.level, settings.log_file_path)

    try:
        pipeline = EmailPipeline(ROOT, settings)
        return pipeline.run_preview(manifest_arg=args.manifest)
    except EmailPipelineError as exc:
        print(exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
