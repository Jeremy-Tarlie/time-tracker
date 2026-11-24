"""Verrou d'instance unique."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from time_tracker.instance import (
    InstanceLock,
    InstanceLockError,
    acquire_instance_lock,
    ensure_single_instance,
    release_instance_lock,
)


class TestInstanceLock(unittest.TestCase):
    def tearDown(self) -> None:
        release_instance_lock()

    def test_second_lock_fails(self) -> None:
        path = str(Path(tempfile.mkdtemp()) / "timetracker.lock")
        with InstanceLock(path):
            with self.assertRaises(InstanceLockError):
                InstanceLock(path).__enter__()

    def test_write_permission_error_becomes_instance_lock_error(self) -> None:
        path = str(Path(tempfile.mkdtemp()) / "timetracker.lock")
        handle = MagicMock()
        handle.seek.return_value = 0
        handle.truncate.side_effect = PermissionError(13, "Permission denied")
        handle.close.side_effect = PermissionError(13, "Permission denied")

        with (
            patch("time_tracker.instance.open", return_value=handle),
            patch("time_tracker.instance.os.path.isfile", return_value=False),
        ):
            with self.assertRaises(InstanceLockError):
                acquire_instance_lock(path)

    def test_ensure_single_instance_false_on_permission_error(self) -> None:
        with patch(
            "time_tracker.instance.acquire_instance_lock",
            side_effect=InstanceLockError("locked"),
        ):
            self.assertFalse(ensure_single_instance())

    def test_stale_permission_lock_removed_then_acquired(self) -> None:
        root = Path(tempfile.mkdtemp())
        path = str(root / "timetracker.lock")
        Path(path).write_text("stale", encoding="utf-8")

        real_open = open
        calls = {"n": 0}

        def open_side_effect(name: str, *args: object, **kwargs: object) -> object:
            if name != path:
                return real_open(name, *args, **kwargs)
            calls["n"] += 1
            if calls["n"] == 1:
                raise PermissionError(13, "Permission denied")
            return real_open(name, *args, **kwargs)

        with patch("time_tracker.instance.open", side_effect=open_side_effect):
            handle = acquire_instance_lock(path)
            self.assertIsNotNone(handle)
            release_instance_lock()


if __name__ == "__main__":
    unittest.main()
