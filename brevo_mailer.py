from __future__ import annotations

import html as html_module
import json
import mimetypes
import random
import re
import shutil
import smtplib
import socket
import ssl
import tempfile
import threading
import time
from dataclasses import dataclass, field
from datetime import date, datetime, time as datetime_time, timedelta
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path
from typing import Callable, Literal

import pandas as pd
from jinja2 import Environment, FileSystemLoader, StrictUndefined, Template

from config import AppConfig
from excel_reader import REPORT_COLUMNS, Recipient


ReportRow = dict[str, str]
ProgressCallback = Callable[[dict[str, object]], None]


@dataclass(frozen=True)
class MailerOptions:
    subject: str
    resume_path: Path
    template_path: Path
    daily_limit: int | None
    html_email: bool
    logo_path: Path | None = None
    schedule_time: str = ""


@dataclass
class MailerResult:
    sent: list[ReportRow] = field(default_factory=list)
    failed: list[ReportRow] = field(default_factory=list)
    skipped: list[ReportRow] = field(default_factory=list)
    stopped: bool = False


class BrevoMailer:
    def __init__(
        self,
        app_config: AppConfig,
        sent_logger,
        failed_logger,
        pause_event: threading.Event,
        stop_event: threading.Event,
    ) -> None:
        self.app_config = app_config
        self.sent_logger = sent_logger
        self.failed_logger = failed_logger
        self.pause_event = pause_event
        self.stop_event = stop_event
        self.sent_emails_path = app_config.path("sent_emails_path")
        self.reconnect_interval = app_config.get_int(
            "reconnect_check_interval_seconds", 15
        )
        self.reconnect_timeout = app_config.get_int("reconnect_timeout_seconds", 1800)
        self.daily_state_path = app_config.path("daily_state_path")

    def send_bulk(
        self,
        recipients: list[Recipient],
        initial_skipped: list[ReportRow],
        options: MailerOptions,
        progress_callback: ProgressCallback | None = None,
    ) -> MailerResult:
        self._validate_ready(options)
        result = MailerResult(skipped=list(initial_skipped))
        sent_emails = self._load_sent_emails()
        total = len(recipients) + len(initial_skipped)
        send_limit = options.daily_limit
        daily_state_date, successful_today = self._load_daily_send_state()
        started_at = time.monotonic()

        for send_index, recipient in enumerate(recipients, start=1):
            index = len(initial_skipped) + send_index
            if self.stop_event.is_set():
                result.stopped = True
                break

            self._wait_if_paused(progress_callback)

            if recipient.email in sent_emails:
                row = self._report_row(
                    recipient, "Skipped", "Already sent according to sent_emails.json"
                )
                result.skipped.append(row)
                self.sent_logger.info(
                    "%s | %s | skipped | already sent",
                    recipient.email,
                    recipient.company,
                )
                self._emit(
                    progress_callback,
                    "skipped",
                    recipient,
                    index,
                    total,
                    result,
                    started_at,
                    "Already sent",
                )
                continue

            while not self.stop_event.is_set():
                today = date.today().isoformat()
                if daily_state_date != today:
                    daily_state_date = today
                    successful_today = 0
                    self._save_daily_send_state(daily_state_date, successful_today)

                if send_limit is not None and successful_today >= send_limit:
                    target = self._next_window_start(force_tomorrow=True)
                    if not self._wait_until_datetime(
                        target,
                        recipient,
                        index,
                        total,
                        result,
                        started_at,
                        f"Daily limit {send_limit} reached. Waiting until {target:%Y-%m-%d %H:%M}.",
                        progress_callback,
                    ):
                        result.stopped = True
                        break
                    continue

                if not self._wait_for_send_window(
                    recipient,
                    index,
                    total,
                    result,
                    started_at,
                    progress_callback,
                ):
                    result.stopped = True
                    break
                break

            if result.stopped:
                break

            self._emit(
                progress_callback,
                "sending",
                recipient,
                index,
                total,
                result,
                started_at,
                "Sending",
            )

            ok, reason = self._send_with_retries(recipient, options, progress_callback)
            if ok:
                daily_state_date = date.today().isoformat()
                successful_today += 1
                self._save_daily_send_state(daily_state_date, successful_today)
                sent_emails.add(recipient.email)
                self._save_sent_emails(sent_emails)
                row = self._report_row(recipient, "Sent", "Delivered to SMTP server")
                result.sent.append(row)
                self.sent_logger.info("%s | %s | %s", recipient.email, recipient.company, row["Reason"])
                self._emit(
                    progress_callback,
                    "sent",
                    recipient,
                    index,
                    total,
                    result,
                    started_at,
                    "Sent",
                )
            else:
                row = self._report_row(recipient, "Failed", reason)
                result.failed.append(row)
                self.failed_logger.error("%s | %s | %s", recipient.email, recipient.company, reason)
                self._emit(
                    progress_callback,
                    "failed",
                    recipient,
                    index,
                    total,
                    result,
                    started_at,
                    reason,
                )

            if send_index < len(recipients):
                delay = random.uniform(
                    self.app_config.get_int("min_delay", 3),
                    self.app_config.get_int("max_delay", 7),
                )
                self._emit(
                    progress_callback,
                    "delaying",
                    recipient,
                    index,
                    total,
                    result,
                    started_at,
                    f"Waiting {delay:.1f}s",
                )
                self._sleep_with_controls(delay, progress_callback)

        self.write_reports(result)
        self._emit(
            progress_callback,
            "complete" if not result.stopped else "stopped",
            None,
            total,
            total,
            result,
            started_at,
            "Completed" if not result.stopped else "Stopped",
        )
        return result

    def render_preview(
        self,
        options: MailerOptions,
        name: str = "Reena Vijayanand",
        company: str = "AdaniConneX",
    ) -> tuple[str, str]:
        recipient = Recipient(
            name=name,
            company=company,
            email="hr@example.com",
            row_number=2,
        )
        subject = self._render_text(options.subject, recipient)
        html = self._render_template(options.template_path, recipient)
        return subject, html

    def write_reports(self, result: MailerResult) -> None:
        reports_dir = self.app_config.path("reports_dir")
        reports_dir.mkdir(parents=True, exist_ok=True)
        self._write_report(reports_dir / "sent_report.xlsx", result.sent)
        self._write_report(reports_dir / "failed_report.xlsx", result.failed)
        self._write_report(reports_dir / "skipped_report.xlsx", result.skipped)

    def export_logs(self) -> Path:
        logs_dir = self.app_config.path("logs_dir")
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        target = self.app_config.path("reports_dir") / f"logs_export_{stamp}"
        archive = shutil.make_archive(str(target), "zip", logs_dir)
        return Path(archive)

    def _send_with_retries(
        self,
        recipient: Recipient,
        options: MailerOptions,
        progress_callback: ProgressCallback | None,
    ) -> tuple[bool, str]:
        max_retries = self.app_config.get_int("retry", 3)
        last_error = "Unknown error"

        for attempt in range(1, max_retries + 1):
            if self.stop_event.is_set():
                return False, "Stopped by user"
            self._wait_if_paused(progress_callback)

            if not self._wait_for_smtp(progress_callback):
                return False, "SMTP server unavailable before send"

            try:
                message = self._build_message(recipient, options)
                self._send_message(message)
                return True, "Sent"
            except (smtplib.SMTPException, OSError, ssl.SSLError) as exc:
                last_error = f"Attempt {attempt}/{max_retries}: {exc}"
                self.failed_logger.error("%s | %s", recipient.email, last_error)
                if attempt < max_retries:
                    backoff = min(60, 5 * attempt)
                    self._sleep_with_controls(backoff, progress_callback)

        return False, last_error

    def _build_message(self, recipient: Recipient, options: MailerOptions) -> EmailMessage:
        brevo = self.app_config.brevo
        message = EmailMessage()
        message["From"] = formataddr((brevo.sender_name, brevo.sender_email))
        message["To"] = recipient.email
        message["Subject"] = self._render_text(options.subject, recipient)

        html_body = self._render_template(options.template_path, recipient)
        plain_body = _html_to_text(html_body)
        message.set_content(plain_body)

        if options.html_email:
            message.add_alternative(html_body, subtype="html")
            if options.logo_path and options.logo_path.exists():
                self._attach_inline_logo(message, options.logo_path)

        self._attach_resume(message, options.resume_path)
        return message

    def _send_message(self, message: EmailMessage) -> None:
        brevo = self.app_config.brevo
        context = ssl.create_default_context()
        with smtplib.SMTP(brevo.server, brevo.port, timeout=60) as smtp:
            smtp.ehlo()
            smtp.starttls(context=context)
            smtp.ehlo()
            smtp.login(brevo.login, brevo.key)
            smtp.send_message(message)

    def _attach_resume(self, message: EmailMessage, resume_path: Path) -> None:
        with resume_path.open("rb") as file:
            data = file.read()
        message.add_attachment(
            data,
            maintype="application",
            subtype="pdf",
            filename=resume_path.name,
        )

    def _attach_inline_logo(self, message: EmailMessage, logo_path: Path) -> None:
        content_type, _ = mimetypes.guess_type(logo_path)
        if not content_type or not content_type.startswith("image/"):
            return
        maintype, subtype = content_type.split("/", 1)
        with logo_path.open("rb") as file:
            data = file.read()
        html_part = message.get_payload()[1]
        html_part.add_related(data, maintype=maintype, subtype=subtype, cid="<company_logo>")

    def _render_template(self, template_path: Path, recipient: Recipient) -> str:
        if not template_path.exists():
            raise FileNotFoundError(f"Email template not found: {template_path}")
        environment = Environment(
            loader=FileSystemLoader(str(template_path.parent)),
            undefined=StrictUndefined,
            autoescape=True,
        )
        template = environment.get_template(template_path.name)
        return template.render(**recipient.template_context())

    def _render_text(self, text: str, recipient: Recipient) -> str:
        template = Template(text, undefined=StrictUndefined, autoescape=False)
        return template.render(**recipient.template_context())

    def _wait_for_smtp(self, progress_callback: ProgressCallback | None) -> bool:
        deadline = time.monotonic() + self.reconnect_timeout
        brevo = self.app_config.brevo
        while not self.stop_event.is_set():
            try:
                with socket.create_connection((brevo.server, brevo.port), timeout=10):
                    return True
            except OSError as exc:
                self.failed_logger.error("SMTP connectivity check failed: %s", exc)
                if time.monotonic() >= deadline:
                    return False
                if progress_callback:
                    progress_callback(
                        {
                            "event": "waiting_for_network",
                            "current_email": "",
                            "message": "Waiting for internet/SMTP reconnect",
                        }
                    )
                self._sleep_with_controls(self.reconnect_interval, progress_callback)
        return False

    def _wait_if_paused(self, progress_callback: ProgressCallback | None) -> None:
        while self.pause_event.is_set() and not self.stop_event.is_set():
            if progress_callback:
                progress_callback({"event": "paused", "message": "Paused"})
            time.sleep(0.25)

    def _wait_for_send_window(
        self,
        recipient: Recipient,
        index: int,
        total: int,
        result: MailerResult,
        started_at: float,
        progress_callback: ProgressCallback | None,
    ) -> bool:
        if not self.app_config.get_bool("send_window_enabled", True):
            return True

        while not self.stop_event.is_set():
            now = datetime.now()
            start_at, end_at = self._send_window_for(now.date())
            if start_at <= now < end_at:
                return True

            target = start_at if now < start_at else self._next_window_start()
            if not self._wait_until_datetime(
                target,
                recipient,
                index,
                total,
                result,
                started_at,
                f"Outside send window {start_at:%H:%M}-{end_at:%H:%M}. Waiting until {target:%Y-%m-%d %H:%M}.",
                progress_callback,
            ):
                return False

        return False

    def _wait_until_datetime(
        self,
        target: datetime,
        recipient: Recipient,
        index: int,
        total: int,
        result: MailerResult,
        started_at: float,
        message: str,
        progress_callback: ProgressCallback | None,
    ) -> bool:
        while not self.stop_event.is_set():
            remaining = (target - datetime.now()).total_seconds()
            if remaining <= 0:
                return True
            self._emit(
                progress_callback,
                "waiting",
                recipient,
                index,
                total,
                result,
                started_at,
                message,
            )
            self._sleep_with_controls(min(60, remaining), progress_callback)
        return False

    def _sleep_with_controls(
        self,
        seconds: float,
        progress_callback: ProgressCallback | None,
    ) -> None:
        end_at = time.monotonic() + seconds
        while time.monotonic() < end_at and not self.stop_event.is_set():
            self._wait_if_paused(progress_callback)
            time.sleep(min(0.25, max(0.0, end_at - time.monotonic())))

    def _validate_ready(self, options: MailerOptions) -> None:
        if not self.app_config.brevo.is_complete:
            raise ValueError("Brevo SMTP settings are incomplete. Update .env first.")
        if not options.resume_path.exists():
            raise FileNotFoundError(f"Resume file not found: {options.resume_path}")
        if not options.template_path.exists():
            raise FileNotFoundError(f"Email template not found: {options.template_path}")

    def _load_sent_emails(self) -> set[str]:
        if not self.sent_emails_path.exists():
            return set()
        try:
            with self.sent_emails_path.open("r", encoding="utf-8") as file:
                data = json.load(file)
            if isinstance(data, dict):
                emails = data.get("emails", [])
                if isinstance(emails, dict):
                    return {str(email).lower() for email in emails.keys()}
                return {str(email).lower() for email in emails}
        except (json.JSONDecodeError, OSError) as exc:
            self.failed_logger.error("Could not read sent_emails.json: %s", exc)
        return set()

    def _load_daily_send_state(self) -> tuple[str, int]:
        today = date.today().isoformat()
        if not self.daily_state_path.exists():
            return today, 0
        try:
            with self.daily_state_path.open("r", encoding="utf-8") as file:
                data = json.load(file)
            state_date = str(data.get("date", today))
            sent_count = int(data.get("sent_count", 0))
            return (state_date, sent_count) if state_date == today else (today, 0)
        except (json.JSONDecodeError, OSError, TypeError, ValueError) as exc:
            self.failed_logger.error("Could not read daily send state: %s", exc)
            return today, 0

    def _save_daily_send_state(self, state_date: str, sent_count: int) -> None:
        payload = {
            "date": state_date,
            "sent_count": sent_count,
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        }
        self.daily_state_path.parent.mkdir(parents=True, exist_ok=True)
        with self.daily_state_path.open("w", encoding="utf-8") as file:
            json.dump(payload, file, indent=2)
            file.write("\n")

    def _send_window_for(self, day: date) -> tuple[datetime, datetime]:
        start_time = self._parse_config_time("send_window_start", "08:00")
        end_time = self._parse_config_time("send_window_end", "12:00")
        start_at = datetime.combine(day, start_time)
        end_at = datetime.combine(day, end_time)
        if end_at <= start_at:
            end_at += timedelta(days=1)
        return start_at, end_at

    def _next_window_start(self, force_tomorrow: bool = False) -> datetime:
        now = datetime.now()
        day = now.date() + timedelta(days=1 if force_tomorrow else 0)
        start_at, end_at = self._send_window_for(day)
        if not force_tomorrow and now >= end_at:
            start_at, _ = self._send_window_for(now.date() + timedelta(days=1))
        return start_at

    def _parse_config_time(self, key: str, default: str) -> datetime_time:
        value = str(self.app_config.values.get(key, default)).strip() or default
        try:
            hour, minute = [int(part) for part in value.split(":", 1)]
            return datetime_time(hour=hour, minute=minute)
        except (TypeError, ValueError):
            fallback_hour, fallback_minute = [int(part) for part in default.split(":", 1)]
            return datetime_time(hour=fallback_hour, minute=fallback_minute)

    def _save_sent_emails(self, emails: set[str]) -> None:
        self.sent_emails_path.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "updated_at": datetime.now().isoformat(timespec="seconds"),
            "emails": sorted(emails),
        }
        with tempfile.NamedTemporaryFile(
            "w",
            encoding="utf-8",
            dir=self.sent_emails_path.parent,
            delete=False,
        ) as file:
            json.dump(payload, file, indent=2)
            file.write("\n")
            temp_name = file.name
        try:
            Path(temp_name).replace(self.sent_emails_path)
        except PermissionError:
            with self.sent_emails_path.open("w", encoding="utf-8") as file:
                json.dump(payload, file, indent=2)
                file.write("\n")

    def _report_row(self, recipient: Recipient, status: str, reason: str) -> ReportRow:
        return {
            "Name": recipient.name,
            "Email": recipient.email,
            "Company": recipient.company,
            "Timestamp": datetime.now().isoformat(timespec="seconds"),
            "Status": status,
            "Reason": reason,
        }

    def _write_report(self, path: Path, rows: list[ReportRow]) -> None:
        dataframe = pd.DataFrame(rows, columns=list(REPORT_COLUMNS))
        dataframe.to_excel(path, index=False, engine="openpyxl")

    def _emit(
        self,
        progress_callback: ProgressCallback | None,
        event: Literal[
            "sending",
            "sent",
            "failed",
            "skipped",
            "delaying",
            "waiting",
            "complete",
            "stopped",
        ],
        recipient: Recipient | None,
        index: int,
        total: int,
        result: MailerResult,
        started_at: float,
        message: str,
    ) -> None:
        if not progress_callback:
            return
        processed = len(result.sent) + len(result.failed) + len(result.skipped)
        progress_callback(
            {
                "event": event,
                "index": index,
                "total": total,
                "processed": min(processed, total),
                "current_name": recipient.name if recipient else "",
                "current_company": recipient.company if recipient else "",
                "current_email": recipient.email if recipient else "",
                "success_count": len(result.sent),
                "failed_count": len(result.failed),
                "skipped_count": len(result.skipped),
                "eta": _format_eta(processed, total, started_at),
                "message": message,
            }
        )


def _format_eta(processed: int, total: int, started_at: float) -> str:
    if processed <= 0 or total <= 0:
        return "Calculating"
    elapsed = time.monotonic() - started_at
    remaining = max(total - processed, 0)
    seconds = int((elapsed / processed) * remaining)
    minutes, seconds = divmod(seconds, 60)
    hours, minutes = divmod(minutes, 60)
    if hours:
        return f"{hours}h {minutes}m"
    if minutes:
        return f"{minutes}m {seconds}s"
    return f"{seconds}s"


def _html_to_text(html: str) -> str:
    text = re.sub(r"(?i)<br\s*/?>", "\n", html)
    text = re.sub(r"(?i)</p>", "\n\n", text)
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return html_module.unescape(text.strip())
