"""Tri Récents et last_used."""

from __future__ import annotations

import unittest
from datetime import datetime

from time_tracker.config import sort_projects, touch_project_last_used
from time_tracker.constants import ProjectEntry


class TestSortRecent(unittest.TestCase):
    def test_recent_desc_uses_last_used_then_created(self) -> None:
        projects: list[ProjectEntry] = [
            {"name": "A", "created": "2026-01-01T00:00:00"},
            {"name": "B", "created": "2026-01-02T00:00:00", "last_used": "2026-02-01T10:00:00"},
            {"name": "C", "created": "2026-01-03T00:00:00", "last_used": "2026-01-15T10:00:00"},
        ]
        ordered = sort_projects(projects, "recent_desc")
        self.assertEqual([p["name"] for p in ordered], ["B", "C", "A"])

    def test_touch_project_last_used(self) -> None:
        projects: list[ProjectEntry] = [
            {"name": "A", "created": "2026-01-01T00:00:00"},
            {"name": "B", "created": "2026-01-02T00:00:00"},
        ]
        when = datetime(2026, 3, 1, 12, 0, 0)
        touch_project_last_used(projects, "B", when=when)
        self.assertEqual(projects[1]["last_used"], "2026-03-01T12:00:00")
        self.assertNotIn("last_used", projects[0])


if __name__ == "__main__":
    unittest.main()
