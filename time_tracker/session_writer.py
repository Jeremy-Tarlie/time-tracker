"""Écriture Excel asynchrone et flush de la file pending."""

from __future__ import annotations

from collections.abc import Callable
from datetime import datetime

from time_tracker.constants import PendingSession
from time_tracker.excel import ExcelWriteError, log_session
from time_tracker.session_store import (
    clear_pending_sessions,
    pending_session_datetimes,
    save_pending_sessions,
)


def write_session_sync(
    project: str,
    session_start: datetime,
    end_time: datetime,
    notes: str,
    excel_file: str,
) -> None:
    log_session(project, session_start, end_time, notes, excel_file)


def flush_pending_sync(
    pending: list[PendingSession],
    excel_file: str,
) -> tuple[list[PendingSession], list[str]]:
    """Écrit les sessions en file jusqu'au premier échec.

    Retourne (remaining, errors).
    """
    remaining: list[PendingSession] = list(pending)
    errors: list[str] = []
    while remaining:
        session = remaining[0]
        start, end = pending_session_datetimes(session)
        try:
            log_session(
                session["project"],
                start,
                end,
                session["notes"],
                excel_file,
            )
            remaining = remaining[1:]
            save_pending_sessions(remaining)
        except ExcelWriteError as exc:
            errors.append(str(exc))
            break
    if not remaining:
        clear_pending_sessions()
    return remaining, errors


def run_in_thread(target: Callable[[], None]) -> None:
    import threading

    threading.Thread(target=target, daemon=True).start()
