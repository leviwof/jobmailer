from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from dotenv import load_dotenv


BASE_DIR = Path(__file__).resolve().parent
CONFIG_FILE = BASE_DIR / "config.json"
ENV_FILE = BASE_DIR / ".env"


DEFAULT_CONFIG: dict[str, Any] = {
    "app_name": "Brevo Job Mailer",
    "subject": "Application for Backend / Full Stack Developer | Ganesh Mishra",
    "excel_path": "excel/kaamkibaatein_HR_Contact_Database.xlsx",
    "template_path": "templates/email.html",
    "resume_path": "resume/Ganesh_Mishra_Resume.pdf",
    "logo_path": "",
    "html_email": True,
    "excel_sheet": "HR Contacts",
    "daily_limit": 300,
    "send_window_enabled": True,
    "send_window_start": "08:00",
    "send_window_end": "12:00",
    "min_delay": 3,
    "max_delay": 7,
    "retry": 3,
    "sent_emails_path": "sent_emails.json",
    "daily_state_path": "daily_send_state.json",
    "reports_dir": "reports",
    "logs_dir": "logs",
    "schedule_start_time": "",
    "reconnect_check_interval_seconds": 15,
    "reconnect_timeout_seconds": 1800,
}


LEGACY_KEYS = {
    "delay_min_seconds": "min_delay",
    "delay_max_seconds": "max_delay",
    "max_retries": "retry",
}


@dataclass(frozen=True)
class BrevoSMTPConfig:
    server: str
    port: int
    sender_email: str
    sender_name: str
    login: str
    key: str

    @property
    def is_complete(self) -> bool:
        return all(
            [
                self.server,
                self.port,
                self.sender_email,
                self.sender_name,
                self.login,
                self.key,
            ]
        )


@dataclass(frozen=True)
class AppConfig:
    values: dict[str, Any]
    brevo: BrevoSMTPConfig

    def path(self, key: str) -> Path:
        value = str(self.values.get(key, "")).strip()
        if not value:
            return BASE_DIR
        path = Path(value)
        return path if path.is_absolute() else BASE_DIR / path

    def get_int(self, key: str, default: int) -> int:
        value = self.values.get(key, default)
        try:
            return int(value)
        except (TypeError, ValueError):
            return default

    def get_bool(self, key: str, default: bool = False) -> bool:
        value = self.values.get(key, default)
        if isinstance(value, bool):
            return value
        if isinstance(value, str):
            return value.strip().lower() in {"1", "true", "yes", "on"}
        return bool(value)


def ensure_directories(config: dict[str, Any] | None = None) -> None:
    values = DEFAULT_CONFIG | (config or {})
    for key in ("reports_dir", "logs_dir"):
        path = Path(values[key])
        if not path.is_absolute():
            path = BASE_DIR / path
        path.mkdir(parents=True, exist_ok=True)
    for key in ("template_path", "resume_path"):
        path = Path(values[key])
        if not path.is_absolute():
            path = BASE_DIR / path
        path.parent.mkdir(parents=True, exist_ok=True)


def load_json_config() -> dict[str, Any]:
    if not CONFIG_FILE.exists():
        return DEFAULT_CONFIG.copy()
    with CONFIG_FILE.open("r", encoding="utf-8") as file:
        loaded = json.load(file)
    return DEFAULT_CONFIG | _normalize_legacy_config(loaded)


def save_json_config(values: dict[str, Any]) -> None:
    merged = DEFAULT_CONFIG | _normalize_legacy_config(values)
    with CONFIG_FILE.open("w", encoding="utf-8") as file:
        json.dump(merged, file, indent=2)
        file.write("\n")


def load_app_config() -> AppConfig:
    load_dotenv(ENV_FILE, override=True)
    values = load_json_config()
    ensure_directories(values)

    brevo = BrevoSMTPConfig(
        server=os.getenv("BREVO_SMTP_SERVER", "smtp-relay.brevo.com").strip(),
        port=_env_int("BREVO_SMTP_PORT", 587),
        sender_email=os.getenv("SENDER_EMAIL", os.getenv("BREVO_EMAIL", "")).strip(),
        sender_name=os.getenv("SENDER_NAME", "Ganesh Mishra").strip(),
        login=os.getenv("BREVO_SMTP_LOGIN", "").strip(),
        key=os.getenv("BREVO_SMTP_KEY", "").strip(),
    )
    return AppConfig(values=values, brevo=brevo)


def _env_int(key: str, default: int) -> int:
    try:
        return int(os.getenv(key, str(default)))
    except ValueError:
        return default


def _normalize_legacy_config(values: dict[str, Any]) -> dict[str, Any]:
    normalized = dict(values)
    for old_key, new_key in LEGACY_KEYS.items():
        if old_key in normalized and new_key not in normalized:
            normalized[new_key] = normalized[old_key]
    return normalized
