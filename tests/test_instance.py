"""Verrou d'instance unique."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from time_tracker.instance import InstanceLock, InstanceLockError, release_instance_lock


class TestInstanceLock(unittest.TestCase):
    def test_second_lock_fails(self) -> None:
        path = str(Path(tempfile.mkdtemp()) / "timetracker.lock")
        with InstanceLock(path):
            with self.assertRaises(InstanceLockError):
                InstanceLock(path).__enter__()
        release_instance_lock()


if __name__ == "__main__":
    unittest.main()
