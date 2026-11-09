"""TimerState / format_elapsed."""

from __future__ import annotations

import unittest
from datetime import datetime
from unittest.mock import patch

from time_tracker.timer_controller import TimerState, format_elapsed


class TestFormatElapsed(unittest.TestCase):
    def test_zero(self) -> None:
        self.assertEqual(format_elapsed(0), "00:00:00")

    def test_hours_minutes_seconds(self) -> None:
        self.assertEqual(format_elapsed(3661), "01:01:01")


class TestTimerState(unittest.TestCase):
    def test_begin_and_pause(self) -> None:
        timer = TimerState()
        gen = timer.begin("P1", now=datetime(2026, 1, 1, 12, 0, 0))
        self.assertTrue(timer.tracking)
        self.assertEqual(timer.active_project, "P1")
        self.assertGreater(gen, 0)
        with (
            patch("time_tracker.timer_controller.clear_active_session"),
            patch("time_tracker.timer_controller.save_active_session"),
        ):
            timer.elapsed = 10.0
            pause_gen = timer.pause()
            self.assertTrue(timer.paused)
            self.assertFalse(timer.running)
            self.assertEqual(timer.accumulated, 10.0)
            self.assertNotEqual(pause_gen, gen)

    def test_stop_clock_returns_session(self) -> None:
        timer = TimerState()
        timer.begin("P1", now=datetime(2026, 1, 1, 12, 0, 0))
        timer.elapsed = 30.0
        timer.accumulated = 30.0
        timer.paused = True
        with patch("time_tracker.timer_controller.clear_active_session"):
            project, _start, _end, elapsed = timer.stop_clock()
        self.assertEqual(project, "P1")
        self.assertEqual(elapsed, 30.0)
        self.assertFalse(timer.tracking)


if __name__ == "__main__":
    unittest.main()
