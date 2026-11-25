"""parse_datetime_input / parse_start_at_input — format partagé datetime."""

from __future__ import annotations

import unittest
from datetime import datetime

from time_tracker.dialogs import (
    DATETIME_INPUT_FORMAT,
    FutureStartError,
    parse_datetime_input,
    parse_start_at_input,
)


class TestParseDatetimeInput(unittest.TestCase):
    def test_parses_valid_format(self) -> None:
        self.assertEqual(
            parse_datetime_input("15-01-2026 09:30:00"),
            datetime(2026, 1, 15, 9, 30, 0),
        )

    def test_strips_whitespace(self) -> None:
        self.assertEqual(
            parse_datetime_input("  15-01-2026 09:30:45  "),
            datetime(2026, 1, 15, 9, 30, 45),
        )

    def test_rejects_wrong_format(self) -> None:
        with self.assertRaises(ValueError):
            parse_datetime_input("2026-01-15 09:30")

    def test_rejects_slash_date(self) -> None:
        with self.assertRaises(ValueError):
            parse_datetime_input("15/01/2026 09:30:00")

    def test_format_constant_matches_label(self) -> None:
        self.assertEqual(DATETIME_INPUT_FORMAT, "%d-%m-%Y %H:%M:%S")


class TestParseStartAtInput(unittest.TestCase):
    def test_accepts_past_or_now(self) -> None:
        now = datetime(2026, 1, 15, 10, 0, 0)
        self.assertEqual(
            parse_start_at_input("15-01-2026 09:30:00", now=now),
            datetime(2026, 1, 15, 9, 30, 0),
        )
        self.assertEqual(
            parse_start_at_input("15-01-2026 10:00:00", now=now),
            datetime(2026, 1, 15, 10, 0, 0),
        )

    def test_rejects_future(self) -> None:
        now = datetime(2026, 1, 15, 10, 0, 0)
        with self.assertRaises(FutureStartError):
            parse_start_at_input("15-01-2026 10:00:01", now=now)

    def test_rejects_wrong_format(self) -> None:
        with self.assertRaises(ValueError):
            parse_start_at_input("15/01/2026 09:30:00", now=datetime(2026, 1, 15, 10, 0))


if __name__ == "__main__":
    unittest.main()
