"""TimerState / format_elapsed."""

from __future__ import annotations

import unittest
from datetime import datetime, timedelta
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

    def test_begin_with_past_now_elapses_correctly(self) -> None:
        timer = TimerState()
        past = datetime.now() - timedelta(hours=2)
        timer.begin("P1", now=past)
        elapsed = timer.tick_elapsed()
        self.assertGreaterEqual(elapsed, 2 * 3600 - 2)
        self.assertLess(elapsed, 2 * 3600 + 5)

    def test_begin_accepts_future_now(self) -> None:
        """Validation « pas de futur » est côté UI, pas du contrôleur."""
        timer = TimerState()
        future = datetime.now() + timedelta(hours=1)
        timer.begin("P1", now=future)
        self.assertEqual(timer.start_time, future)


if __name__ == "__main__":
    unittest.main()
