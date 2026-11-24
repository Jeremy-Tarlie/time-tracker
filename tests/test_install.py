"""Remplacement robuste du binaire installé (PermissionError / hash identique)."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from time_tracker.install import _binaries_identical, replace_installed_exe


class TestReplaceInstalledExe(unittest.TestCase):
    def _tmp(self) -> Path:
        if not hasattr(self, "_tmpdir"):
            self._tmpdir = Path(tempfile.mkdtemp())
        return self._tmpdir

    def test_skips_copy_when_hashes_identical(self) -> None:
        root = self._tmp()
        source = root / "new.exe"
        dest = root / "installed.exe"
        source.write_bytes(b"same-bytes")
        dest.write_bytes(b"same-bytes")
        with patch("time_tracker.install.shutil.copy2") as copy2:
            self.assertTrue(replace_installed_exe(str(source), str(dest)))
            copy2.assert_not_called()

    def test_direct_copy_when_dest_writable(self) -> None:
        root = self._tmp()
        source = root / "new.exe"
        dest = root / "installed.exe"
        source.write_bytes(b"v2-binary")
        dest.write_bytes(b"v1-binary")
        self.assertTrue(replace_installed_exe(str(source), str(dest)))
        self.assertEqual(dest.read_bytes(), b"v2-binary")

    def test_rename_old_fallback_when_direct_copy_permission_error(self) -> None:
        root = self._tmp()
        source = root / "new.exe"
        dest = root / "installed.exe"
        source.write_bytes(b"v2-binary")
        dest.write_bytes(b"v1-binary")

        real_copy2 = __import__("shutil").copy2
        calls = {"n": 0}

        def copy2_side_effect(src: str, dst: str) -> None:
            calls["n"] += 1
            if calls["n"] == 1:
                raise PermissionError(32, "locked")
            real_copy2(src, dst)

        with patch("time_tracker.install.shutil.copy2", side_effect=copy2_side_effect):
            self.assertTrue(replace_installed_exe(str(source), str(dest)))
        self.assertEqual(dest.read_bytes(), b"v2-binary")
        self.assertFalse((root / "installed.exe.old").exists())

    def test_returns_false_when_rename_also_fails(self) -> None:
        root = self._tmp()
        source = root / "new.exe"
        dest = root / "installed.exe"
        source.write_bytes(b"v2-binary")
        dest.write_bytes(b"v1-binary")

        with (
            patch(
                "time_tracker.install.shutil.copy2",
                side_effect=PermissionError(32, "locked"),
            ),
            patch(
                "time_tracker.install.os.replace",
                side_effect=PermissionError(32, "locked"),
            ),
        ):
            self.assertFalse(replace_installed_exe(str(source), str(dest)))
        self.assertEqual(dest.read_bytes(), b"v1-binary")

    def test_binaries_identical_false_when_missing(self) -> None:
        root = self._tmp()
        source = root / "new.exe"
        source.write_bytes(b"x")
        self.assertFalse(_binaries_identical(str(source), str(root / "missing.exe")))


if __name__ == "__main__":
    unittest.main()
