from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path


def setup_app_logging(logs_dir: Path) -> tuple[logging.Logger, logging.Logger]:
    logs_dir.mkdir(parents=True, exist_ok=True)

    sent_logger = logging.getLogger("job_mailer.sent")
    failed_logger = logging.getLogger("job_mailer.failed")

    sent_logger.setLevel(logging.INFO)
    failed_logger.setLevel(logging.ERROR)
    sent_logger.propagate = False
    failed_logger.propagate = False

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    _replace_handlers(
        sent_logger,
        RotatingFileHandler(
            logs_dir / "sent.log",
            maxBytes=5_000_000,
            backupCount=5,
            encoding="utf-8",
        ),
        formatter,
    )
    _replace_handlers(
        failed_logger,
        RotatingFileHandler(
            logs_dir / "failed.log",
            maxBytes=5_000_000,
            backupCount=5,
            encoding="utf-8",
        ),
        formatter,
    )

    return sent_logger, failed_logger


def _replace_handlers(
    logger: logging.Logger,
    handler: logging.Handler,
    formatter: logging.Formatter,
) -> None:
    for existing in list(logger.handlers):
        logger.removeHandler(existing)
        existing.close()
    handler.setFormatter(formatter)
    logger.addHandler(handler)
