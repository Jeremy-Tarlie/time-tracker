"""Régression : import des modules UI (tk.Event[...] sous Python < 3.13)."""

from __future__ import annotations

import unittest


class TestImportSmoke(unittest.TestCase):
    def test_widgets_importable(self) -> None:
        import time_tracker.widgets  # noqa: F401

    def test_app_importable(self) -> None:
        import time_tracker.app  # noqa: F401


if __name__ == "__main__":
    unittest.main()
