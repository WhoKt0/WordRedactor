"""SMTP email sending with PDF attachment."""

from __future__ import annotations

import logging
import re
import smtplib
import ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from pathlib import Path

from src.config import EnvSettings

logger = logging.getLogger(__name__)

PLACEHOLDER_PATTERN = re.compile(r"\{\{([A-Z_]+)\}\}")


class EmailSenderError(Exception):
    """Raised when email sending fails."""


class EmailSender:
    """Send emails via SMTP with optional CC/BCC and PDF attachment."""

    def __init__(self, env: EnvSettings) -> None:
        self.env = env

    def render_template(self, template_text: str, variables: dict[str, str]) -> str:
        def _replace(match: re.Match[str]) -> str:
            key = match.group(1)
            placeholder = f"{{{{{key}}}}}"
            if placeholder in variables:
                return variables[placeholder]
            if key in variables:
                return variables[key]
            logger.warning("Email template variable not found in data: %s", placeholder)
            return match.group(0)

        return PLACEHOLDER_PATTERN.sub(_replace, template_text)

    def send(
        self,
        *,
        to_email: str,
        cc_email: str,
        bcc_email: str,
        subject: str,
        body: str,
        pdf_path: Path,
        bank_name_for_log: str,
    ) -> None:
        if not self.env.smtp_host:
            raise EmailSenderError("SMTP_HOST is not configured")
        if not self.env.from_email:
            raise EmailSenderError("FROM_EMAIL is not configured")

        msg = MIMEMultipart()
        msg["From"] = formataddr((self.env.from_name, self.env.from_email))
        msg["To"] = to_email
        msg["Subject"] = subject

        recipients = [to_email]
        if cc_email:
            msg["Cc"] = cc_email
            recipients.extend(_split_emails(cc_email))
        if bcc_email:
            recipients.extend(_split_emails(bcc_email))

        msg.attach(MIMEText(body, "plain", "utf-8"))
        self._attach_pdf(msg, pdf_path)

        try:
            with smtplib.SMTP(self.env.smtp_host, self.env.smtp_port, timeout=60) as server:
                server.ehlo()
                if self.env.smtp_port in (587, 25):
                    server.starttls(context=ssl.create_default_context())
                    server.ehlo()
                if self.env.smtp_user and self.env.smtp_password:
                    server.login(self.env.smtp_user, self.env.smtp_password)
                server.sendmail(self.env.from_email, recipients, msg.as_string())

            logger.info(
                "Email sent to %s (bank: %s), attachment: %s",
                to_email,
                bank_name_for_log,
                pdf_path.name,
            )
        except smtplib.SMTPAuthenticationError as exc:
            raise EmailSenderError(
                "SMTP authentication failed. Check SMTP_USER/SMTP_PASSWORD. "
                "Gmail requires an app password."
            ) from exc
        except smtplib.SMTPException as exc:
            raise EmailSenderError(f"SMTP error: {exc}") from exc

    @staticmethod
    def _attach_pdf(msg: MIMEMultipart, pdf_path: Path) -> None:
        if not pdf_path.exists():
            raise EmailSenderError(f"PDF attachment not found: {pdf_path}")

        with pdf_path.open("rb") as fh:
            part = MIMEBase("application", "pdf")
            part.set_payload(fh.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f'attachment; filename="{pdf_path.name}"',
        )
        msg.attach(part)


def _split_emails(value: str) -> list[str]:
    return [e.strip() for e in value.replace(";", ",").split(",") if e.strip()]
