from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd

from validator import is_valid_email, normalize_email


SHEET_NAME = "HR Contacts"
REQUIRED_COLUMNS = ("Name", "Company", "Email")
REPORT_COLUMNS = ("Name", "Company", "Email", "Timestamp", "Status", "Reason")


@dataclass(frozen=True)
class Recipient:
    name: str
    company: str
    email: str
    row_number: int
    title: str = ""
    category: str = ""

    @property
    def first_name(self) -> str:
        if not self.name.strip():
            return "Team"
        return self.name.strip().split()[0]

    def template_context(self) -> dict[str, str]:
        return {
            "FIRST_NAME": self.first_name,
            "NAME": self.name,
            "COMPANY": self.company or "your company",
            "EMAIL": self.email,
            "TITLE": self.title,
            "CATEGORY": self.category,
        }


def load_recipients(
    file_path: str | Path,
    sheet_name: str = SHEET_NAME,
) -> tuple[list[Recipient], list[dict[str, str]]]:
    data = _read_contacts(file_path, sheet_name)
    data = data.rename(columns={column: str(column).strip() for column in data.columns})
    missing = [column for column in REQUIRED_COLUMNS if column not in data.columns]
    if missing:
        raise ValueError(f"Missing required column(s): {', '.join(missing)}")

    recipients: list[Recipient] = []
    skipped: list[dict[str, str]] = []
    seen: set[str] = set()

    for index, row in data.iterrows():
        row_number = int(index) + 2
        name = str(row.get("Name", "")).strip()
        company = str(row.get("Company", "")).strip()
        title = str(row.get("Title", "")).strip()
        category = str(row.get("Category", "")).strip()
        email = normalize_email(row.get("Email", ""))

        if not email:
            skipped.append(
                _report_row(name, company, email, row_number, "Skipped", "Email empty")
            )
            continue

        if not is_valid_email(email):
            skipped.append(
                _report_row(
                    name,
                    company,
                    email,
                    row_number,
                    "Skipped",
                    "Invalid email address",
                )
            )
            continue

        if email in seen:
            skipped.append(
                _report_row(
                    name,
                    company,
                    email,
                    row_number,
                    "Skipped",
                    "Duplicate email in input file",
                )
            )
            continue

        seen.add(email)
        recipients.append(
            Recipient(
                name=name,
                company=company,
                email=email,
                row_number=row_number,
                title=title,
                category=category,
            )
        )

    return recipients, skipped


def load_recipients_from_report(file_path: str | Path) -> list[Recipient]:
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"Report not found: {path}")
    data = pd.read_excel(path, dtype=str, keep_default_na=False, engine="openpyxl")
    data = data.rename(columns={column: str(column).strip() for column in data.columns})
    recipients: list[Recipient] = []
    for index, row in data.iterrows():
        email = normalize_email(row.get("Email", ""))
        if not is_valid_email(email):
            continue
        recipients.append(
            Recipient(
                name=str(row.get("Name", "")).strip(),
                company=str(row.get("Company", "")).strip(),
                email=email,
                row_number=int(index) + 2,
            )
        )
    return recipients


def search_recipients(
    file_path: str | Path,
    query: str,
    sheet_name: str = SHEET_NAME,
    limit: int = 50,
) -> list[Recipient]:
    recipients, _ = load_recipients(file_path, sheet_name)
    needle = query.strip().lower()
    if not needle:
        return recipients[:limit]
    matches = [
        recipient
        for recipient in recipients
        if needle in recipient.name.lower()
        or needle in recipient.company.lower()
        or needle in recipient.email.lower()
    ]
    return matches[:limit]


def _read_contacts(file_path: str | Path, sheet_name: str) -> pd.DataFrame:
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    suffix = path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    if suffix == ".xlsx":
        try:
            return pd.read_excel(
                path,
                sheet_name=sheet_name,
                dtype=str,
                keep_default_na=False,
                engine="openpyxl",
            )
        except ValueError as exc:
            workbook = pd.ExcelFile(path, engine="openpyxl")
            available = ", ".join(workbook.sheet_names)
            raise ValueError(
                f"Sheet '{sheet_name}' was not found. Available sheets: {available}"
            ) from exc
    raise ValueError("Only .xlsx and .csv files are supported.")


def _report_row(
    name: str,
    company: str,
    email: str,
    row_number: int,
    status: str,
    reason: str,
) -> dict[str, str]:
    return {
        "Name": name,
        "Company": company,
        "Email": email,
        "Timestamp": datetime.now().isoformat(timespec="seconds"),
        "Status": status,
        "Reason": f"Row {row_number}: {reason}",
    }
