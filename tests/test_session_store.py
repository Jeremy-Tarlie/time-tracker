"""File d'attente et session active."""

from __future__ import annotations

import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from time_tracker.session_store import (
    clear_active_session,
    clear_pending_sessions,
    enqueue_pending_session,
    load_active_session,
    load_pending_sessions,
    pending_session_datetimes,
    save_active_session,
)


class TestSessionStore(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp())
        self.pending = str(self.tmp / "pending_sessions.json")
        self.active = str(self.tmp / "active_session.json")

    def test_enqueue_and_load_pending(self) -> None:
        start = datetime(2026, 3, 1, 10, 0, 0)
        end = datetime(2026, 3, 1, 11, 0, 0)
        enqueue_pending_session("P", start, end, "n", self.pending)
        sessions = load_pending_sessions(self.pending)
        self.assertEqual(len(sessions), 1)
        self.assertEqual(sessions[0]["project"], "P")
        s, e = pending_session_datetimes(sessions[0])
        self.assertEqual(s, start)
        self.assertEqual(e, end)
        clear_pending_sessions(self.pending)
        self.assertEqual(load_pending_sessions(self.pending), [])

    def test_active_session_roundtrip(self) -> None:
        save_active_session(
            {
                "project": "X",
                "start": "2026-03-01T10:00:00",
                "accumulated_seconds": 12.5,
                "paused": True,
                "resume_at": None,
            },
            self.active,
        )
        loaded = load_active_session(self.active)
        self.assertIsNotNone(loaded)
        assert loaded is not None
        self.assertEqual(loaded["project"], "X")
        self.assertTrue(loaded["paused"])
        clear_active_session(self.active)
        self.assertIsNone(load_active_session(self.active))


if __name__ == "__main__":
    unittest.main()
