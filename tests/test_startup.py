"""Helpers démarrage multi-OS (sans toucher au registre réel)."""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from time_tracker.startup import (
    resolve_launch_command,
    startup_available,
    startup_label,
)


class TestStartup(unittest.TestCase):
    def test_resolve_launch_command_quotes_exe(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            exe = str(Path(tmp) / "TimeTracker.exe")
            cmd = resolve_launch_command(exe)
            self.assertTrue(cmd.startswith('"'))
            self.assertIn("TimeTracker.exe", cmd)

    def test_resolve_launch_command_dev_includes_python(self) -> None:
        with patch("time_tracker.startup.is_frozen", return_value=False):
            with patch("time_tracker.startup.sys.executable", r"C:\Python\python.exe"):
                with patch("time_tracker.startup.sys.argv", [r"C:\app\main.py"]):
                    cmd = resolve_launch_command(None)
        self.assertIn("python.exe", cmd)
        self.assertIn("main.py", cmd)

    def test_startup_label_non_empty(self) -> None:
        self.assertTrue(startup_label())

    def test_portable_disables_startup(self) -> None:
        with patch("time_tracker.startup.is_portable_mode", return_value=True):
            self.assertFalse(startup_available())

    def test_supported_platforms(self) -> None:
        with patch("time_tracker.startup.is_portable_mode", return_value=False):
            with patch.object(sys, "platform", "win32"):
                self.assertTrue(startup_available())
            with patch.object(sys, "platform", "linux"):
                self.assertTrue(startup_available())
            with patch.object(sys, "platform", "darwin"):
                self.assertTrue(startup_available())
            with patch.object(sys, "platform", "freebsd"):
                self.assertFalse(startup_available())


if __name__ == "__main__":
    unittest.main()
