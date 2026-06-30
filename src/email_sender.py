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
from src.email_body import (
    compose_html_body,
    compose_plain_body,
    load_signature_html,
    load_signature_text,
)

logger = logging.getLogger(__name__)

PLACEHOLDER_PATTERN = re.compile(r"\{\{([A-Z_]+)\}\}")

SMTP_HOST_LOOKS_LIKE_EMAIL = (
    "SMTP_HOST похож на email. Укажите SMTP-сервер, например renser.kz."
)


class EmailSenderError(Exception):
    """Raised when email sending fails."""


def get_smtp_mode(smtp_port: int) -> str:
    if smtp_port == 465:
        return "SMTP_SSL"
    if smtp_port == 587:
        return "STARTTLS"
    return "PLAIN_SMTP"


def _validate_smtp_host(smtp_host: str) -> None:
    if "@" in smtp_host:
        raise EmailSenderError(SMTP_HOST_LOOKS_LIKE_EMAIL)


def _connect_smtp(env: EnvSettings) -> smtplib.SMTP:
    _validate_smtp_host(env.smtp_host)
    mode = get_smtp_mode(env.smtp_port)
    logger.info(
        "SMTP connect: host=%s port=%s mode=%s from=%s",
        env.smtp_host,
        env.smtp_port,
        mode,
        env.from_email,
    )

    if env.smtp_port == 465:
        return smtplib.SMTP_SSL(
            env.smtp_host,
            env.smtp_port,
            context=ssl.create_default_context(),
            timeout=60,
        )
    if env.smtp_port == 587:
        server = smtplib.SMTP(env.smtp_host, env.smtp_port, timeout=60)
        server.ehlo()
        server.starttls(context=ssl.create_default_context())
        server.ehlo()
        return server

    return smtplib.SMTP(env.smtp_host, env.smtp_port, timeout=60)


class EmailSender:
    """Send emails via SMTP with optional CC/BCC and PDF attachment."""

    def __init__(
        self,
        env: EnvSettings,
        *,
        signature_text_path: Path | None = None,
        signature_html_path: Path | None = None,
    ) -> None:
        self.env = env
        self.signature_text_path = signature_text_path
        self.signature_html_path = signature_html_path

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

    def _build_bodies(self, main_body: str) -> tuple[str, str]:
        signature_text = ""
        signature_html = ""
        if self.signature_text_path:
            signature_text = load_signature_text(self.signature_text_path)
        if self.signature_html_path:
            signature_html = load_signature_html(self.signature_html_path)

        plain_body = compose_plain_body(main_body, signature_text)
        html_body = compose_html_body(main_body, signature_html)
        return plain_body, html_body

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

        plain_body, html_body = self._build_bodies(body)

        msg = MIMEMultipart("mixed")
        msg["From"] = formataddr((self.env.from_name, self.env.from_email))
        msg["To"] = to_email
        msg["Subject"] = subject

        recipients = [to_email]
        if cc_email:
            msg["Cc"] = cc_email
            recipients.extend(_split_emails(cc_email))
        if bcc_email:
            recipients.extend(_split_emails(bcc_email))

        alternative = MIMEMultipart("alternative")
        alternative.attach(MIMEText(plain_body, "plain", "utf-8"))
        alternative.attach(MIMEText(html_body, "html", "utf-8"))
        msg.attach(alternative)
        self._attach_pdf(msg, pdf_path)

        server: smtplib.SMTP | None = None
        try:
            server = _connect_smtp(self.env)
            if self.env.smtp_user and self.env.smtp_password:
                server.login(self.env.smtp_user, self.env.smtp_password)
            server.sendmail(self.env.from_email, recipients, msg.as_string())

            logger.info(
                "Email sent to %s (bank: %s), attachment: %s, format=plain+html",
                to_email,
                bank_name_for_log,
                pdf_path.name,
            )
        except smtplib.SMTPAuthenticationError as exc:
            raise EmailSenderError(
                "SMTP authentication failed. Check SMTP_USER/SMTP_PASSWORD."
            ) from exc
        except smtplib.SMTPException as exc:
            raise EmailSenderError(f"SMTP error: {exc}") from exc
        finally:
            if server is not None:
                try:
                    server.quit()
                except smtplib.SMTPException:
                    logger.debug("SMTP quit failed", exc_info=True)

    @staticmethod
    def _attach_pdf(msg: MIMEMultipart, pdf_path: Path) -> None:
        if not pdf_path.exists():
            raise EmailSenderError(f"PDF attachment not found: {pdf_path}")

        filename = pdf_path.name
        with pdf_path.open("rb") as fh:
            part = MIMEBase("application", "pdf")
            part.set_payload(fh.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            "attachment",
            filename=("utf-8", "", filename),
        )
        msg.attach(part)


def _split_emails(value: str) -> list[str]:
    return [e.strip() for e in value.replace(";", ",").split(",") if e.strip()]
