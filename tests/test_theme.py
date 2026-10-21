"""Smoke test version / constantes (THEME UI non porté en package — conservé minimal)."""

from __future__ import annotations

import unittest

from time_tracker.constants import APP_VERSION


class TestConstants(unittest.TestCase):
    def test_app_version_is_semver_like(self) -> None:
        parts = APP_VERSION.split(".")
        self.assertGreaterEqual(len(parts), 2)
        for part in parts:
            self.assertTrue(part.isdigit(), f"segment non numérique: {part}")


if __name__ == "__main__":
    unittest.main()
