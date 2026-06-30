"""Configuration loading from config.yaml and .env."""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from typing import Any

import yaml
from dotenv import load_dotenv
from pydantic import AliasChoices, Field
from pydantic_settings import BaseSettings, SettingsConfigDict


class EnvSettings(BaseSettings):
    """Environment variables (SMTP, outgoing number, sender)."""

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    smtp_host: str = Field(default="", alias="SMTP_HOST")
    smtp_port: int = Field(default=587, alias="SMTP_PORT")
    smtp_user: str = Field(default="", alias="SMTP_USER")
    smtp_password: str = Field(default="", alias="SMTP_PASSWORD")
    from_email: str = Field(default="", alias="FROM_EMAIL")
    from_name: str = Field(default="RenSer Technologies", alias="FROM_NAME")
    start_out_number: int = Field(default=315, alias="START_OUT_NUMBER")
    test_mail: str = Field(
        default="",
        validation_alias=AliasChoices("TEST_MAIL", "Test_mail", "test_mail"),
    )


class AppConfig(BaseSettings):
    dry_run: bool = True
    stop_on_error: bool = False
    delay_between_emails_seconds: int = 15
    explicit_date: str | None = None
    date_format: str = "%d.%m.%Y"


class PathsConfig(BaseSettings):
    word_template: str = "templates/letter_template.docx"
    email_template: str = "templates/email_template.txt"
    email_signature_text: str = "templates/email_signature.txt"
    email_signature_html: str = "templates/email_signature.html"
    excel_file: str = "data/banks.xlsx"
    output_docx_dir: str = "output/docx"
    output_pdf_dir: str = "output/pdf"
    reports_dir: str = "output/reports"


class EmailConfig(BaseSettings):
    subject_template: str = (
        "Предложение по поставке банковских карт и PIN-конвертов для {{EMAIL_BANK_NAME}}"
    )


class PdfConfig(BaseSettings):
    converter: str = "word_com"


class LoggingConfig(BaseSettings):
    level: str = "INFO"


class Settings:
    """Combined application settings."""

    def __init__(self, project_root: Path, config_path: Path | None = None) -> None:
        self.project_root = project_root.resolve()
        load_dotenv(self.project_root / ".env")

        config_file = config_path or (self.project_root / "config.yaml")
        raw: dict[str, Any] = {}
        if config_file.exists():
            with config_file.open(encoding="utf-8") as fh:
                raw = yaml.safe_load(fh) or {}

        self.app = AppConfig(**(raw.get("app") or {}))
        self.paths = PathsConfig(**(raw.get("paths") or {}))
        self.email = EmailConfig(**(raw.get("email") or {}))
        self.pdf = PdfConfig(**(raw.get("pdf") or {}))
        self.logging = LoggingConfig(**(raw.get("logging") or {}))
        self.env = EnvSettings()

    def resolve(self, relative: str) -> Path:
        return (self.project_root / relative).resolve()

    @property
    def word_template_path(self) -> Path:
        return self.resolve(self.paths.word_template)

    @property
    def email_template_path(self) -> Path:
        return self.resolve(self.paths.email_template)

    @property
    def email_signature_text_path(self) -> Path:
        return self.resolve(self.paths.email_signature_text)

    @property
    def email_signature_html_path(self) -> Path:
        return self.resolve(self.paths.email_signature_html)

    @property
    def excel_file_path(self) -> Path:
        return self.resolve(self.paths.excel_file)

    @property
    def output_docx_dir(self) -> Path:
        return self.resolve(self.paths.output_docx_dir)

    @property
    def output_pdf_dir(self) -> Path:
        return self.resolve(self.paths.output_pdf_dir)

    @property
    def reports_dir(self) -> Path:
        return self.resolve(self.paths.reports_dir)

    @property
    def log_file_path(self) -> Path:
        return self.project_root / "logs" / "app.log"

    def get_letter_date(self) -> date:
        if self.app.explicit_date:
            for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
                try:
                    return datetime.strptime(self.app.explicit_date, fmt).date()
                except ValueError:
                    continue
            raise ValueError(
                f"Cannot parse explicit_date: {self.app.explicit_date!r}. "
                "Use YYYY-MM-DD or DD.MM.YYYY."
            )
        return date.today()

    def format_date(self, value: date | None = None) -> str:
        d = value or self.get_letter_date()
        return d.strftime(self.app.date_format)
