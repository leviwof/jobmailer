from __future__ import annotations

import re
from email.utils import parseaddr


EMAIL_PATTERN = re.compile(
    r"^(?=.{1,254}$)(?=.{1,64}@)[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@"
    r"(?:[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?\.)+"
    r"[A-Za-z]{2,63}$"
)


def normalize_email(email: object) -> str:
    if email is None:
        return ""
    parsed_name, parsed_email = parseaddr(str(email).strip())
    return (parsed_email or parsed_name).strip().lower()


def is_valid_email(email: object) -> bool:
    normalized = normalize_email(email)
    return bool(normalized and EMAIL_PATTERN.match(normalized))
