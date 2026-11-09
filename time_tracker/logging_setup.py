"""Configuration du logging applicatif (fichier rotatif, sans PII)."""

from __future__ import annotations

import logging
import os
from logging.handlers import RotatingFileHandler

from time_tracker.paths import get_app_data_dir

_CONFIGURED = False


def setup_logging() -> logging.Logger:
    global _CONFIGURED
    logger = logging.getLogger("timetracker")
    if _CONFIGURED:
        return logger

    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    logger.propagate = False

    log_dir = get_app_data_dir()
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir, "timetracker.log")
    handler = RotatingFileHandler(
        log_path,
        maxBytes=512_000,
        backupCount=3,
        encoding="utf-8",
    )
    handler.setFormatter(
        logging.Formatter(
            fmt="%(asctime)s %(levelname)s %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
    )
    logger.addHandler(handler)
    _CONFIGURED = True
    logger.info("Logging initialisé")
    return logger


def get_logger() -> logging.Logger:
    if not _CONFIGURED:
        return setup_logging()
    return logging.getLogger("timetracker")
