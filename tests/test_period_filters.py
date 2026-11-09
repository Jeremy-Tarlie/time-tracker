"""Filtres période / durée dérivée des horaires."""

from __future__ import annotations

import os
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

from time_tracker.excel import (
    ensure_excel,
    hours_by_project,
    list_sessions_for_day,
    list_sessions_in_range,
    log_session,
)


class TestPeriodFilters(unittest.TestCase):
    def test_duration_label_prefers_start_end_over_excel_cell(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "heures_projet.xlsx")
            ensure_excel(path)
            start = datetime(2026, 11, 8, 12, 13, 0)
            end = datetime(2026, 11, 8, 13, 13, 0)
            log_session("test", start, end, "", path)
            wb = load_workbook(path)
            # Simule une cellule durée Excel incohérente (25 h).
            wb["Sessions"].cell(2, 5).value = 25.0
            wb.save(path)
            wb.close()
            sessions = list_sessions_for_day(path, start)
            self.assertEqual(len(sessions), 1)
            self.assertEqual(sessions[0]["duration"], "1h 00min 00s")

    def test_hours_by_project_aggregates(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "heures_projet.xlsx")
            ensure_excel(path)
            log_session(
                "A",
                datetime(2026, 11, 9, 10, 0, 0),
                datetime(2026, 11, 9, 11, 0, 0),
                "",
                path,
            )
            log_session(
                "B",
                datetime(2026, 11, 9, 11, 0, 0),
                datetime(2026, 11, 9, 11, 30, 0),
                "",
                path,
            )
            log_session(
                "A",
                datetime(2026, 11, 9, 12, 0, 0),
                datetime(2026, 11, 9, 12, 15, 0),
                "",
                path,
            )
            rows = list_sessions_in_range(
                path,
                datetime(2026, 11, 9).date(),
                datetime(2026, 11, 9).date(),
            )
            by_project = hours_by_project(rows)
            self.assertEqual([name for name, _hours in by_project], ["A", "B"])
            self.assertAlmostEqual(by_project[0][1], 1.25, places=4)
            self.assertAlmostEqual(by_project[1][1], 0.5, places=4)

    def test_project_filter_client_side(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = str(Path(tmp) / "heures_projet.xlsx")
            ensure_excel(path)
            log_session(
                "coucou",
                datetime(2026, 11, 9, 14, 0, 0),
                datetime(2026, 11, 9, 14, 5, 0),
                "",
                path,
            )
            log_session(
                "test",
                datetime(2026, 11, 9, 15, 0, 0),
                datetime(2026, 11, 9, 15, 10, 0),
                "",
                path,
            )
            rows = list_sessions_in_range(
                path,
                datetime(2026, 11, 9).date(),
                datetime(2026, 11, 9).date(),
            )
            filtered = [row for row in rows if row["project"] == "coucou"]
            self.assertEqual(len(filtered), 1)
            self.assertEqual(filtered[0]["project"], "coucou")


if __name__ == "__main__":
    unittest.main()
