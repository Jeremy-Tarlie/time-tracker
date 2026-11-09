"""Persistance des sessions actives et en attente d'écriture Excel."""

from __future__ import annotations

import json
import os
from datetime import datetime

from time_tracker.atomic_io import write_json_atomic
from time_tracker.constants import ActiveSession, PendingSession
from time_tracker.paths import ACTIVE_SESSION_FILE, PENDING_SESSIONS_FILE


def _parse_datetime(value: str) -> datetime:
    return datetime.fromisoformat(value)


def load_pending_sessions(path: str = PENDING_SESSIONS_FILE) -> list[PendingSession]:
    if not os.path.exists(path):
        return []
    try:
        with open(path, encoding="utf-8") as handle:
            raw: object = json.load(handle)
    except (OSError, json.JSONDecodeError, UnicodeDecodeError):
        return []
    if not isinstance(raw, list):
        return []
    sessions: list[PendingSession] = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        project = str(item.get("project", "")).strip()
        start = str(item.get("start", "")).strip()
        end = str(item.get("end", "")).strip()
        notes = str(item.get("notes", ""))
        if not project or not start or not end:
            continue
        try:
            _parse_datetime(start)
            _parse_datetime(end)
        except ValueError:
            continue
        sessions.append(
            {
                "project": project,
                "start": start,
                "end": end,
                "notes": notes,
            }
        )
    return sessions


def save_pending_sessions(
    sessions: list[PendingSession],
    path: str = PENDING_SESSIONS_FILE,
) -> None:
    write_json_atomic(path, sessions)


def enqueue_pending_session(
    project: str,
    start: datetime,
    end: datetime,
    notes: str,
    path: str = PENDING_SESSIONS_FILE,
) -> PendingSession:
    session: PendingSession = {
        "project": project,
        "start": start.isoformat(timespec="seconds"),
        "end": end.isoformat(timespec="seconds"),
        "notes": notes,
    }
    sessions = load_pending_sessions(path)
    sessions.append(session)
    save_pending_sessions(sessions, path)
    return session


def clear_pending_sessions(path: str = PENDING_SESSIONS_FILE) -> None:
    save_pending_sessions([], path)


def pending_session_datetimes(
    session: PendingSession,
) -> tuple[datetime, datetime]:
    return _parse_datetime(session["start"]), _parse_datetime(session["end"])


def load_active_session(path: str = ACTIVE_SESSION_FILE) -> ActiveSession | None:
    if not os.path.exists(path):
        return None
    try:
        with open(path, encoding="utf-8") as handle:
            raw: object = json.load(handle)
    except (OSError, json.JSONDecodeError, UnicodeDecodeError):
        return None
    if not isinstance(raw, dict):
        return None
    project = str(raw.get("project", "")).strip()
    start = str(raw.get("start", "")).strip()
    if not project or not start:
        return None
    try:
        _parse_datetime(start)
        accumulated = float(raw.get("accumulated_seconds", 0) or 0)
    except (TypeError, ValueError):
        return None
    paused = bool(raw.get("paused", False))
    resume_raw = raw.get("resume_at")
    resume_at: str | None
    if resume_raw is None or resume_raw == "":
        resume_at = None
    else:
        resume_at = str(resume_raw)
        try:
            _parse_datetime(resume_at)
        except ValueError:
            return None
    return {
        "project": project,
        "start": start,
        "accumulated_seconds": max(0.0, accumulated),
        "paused": paused,
        "resume_at": resume_at,
    }


def save_active_session(
    session: ActiveSession,
    path: str = ACTIVE_SESSION_FILE,
) -> None:
    write_json_atomic(path, session)


def clear_active_session(path: str = ACTIVE_SESSION_FILE) -> None:
    if os.path.exists(path):
        try:
            os.remove(path)
        except OSError:
            pass
