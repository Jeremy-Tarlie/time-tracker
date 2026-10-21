"""Robustesse de load_config face à un JSON invalide."""

from __future__ import annotations

import json
import unittest
from pathlib import Path
from unittest.mock import patch

from time_tracker import config as config_mod
from time_tracker.config import load_config


class TestLoadConfig(unittest.TestCase):
    def test_corrupt_json_falls_back_to_defaults(self) -> None:
        with patch.object(config_mod, "CONFIG_FILE", str(Path(self._tmp()) / "projects.json")):
            path = Path(config_mod.CONFIG_FILE)
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text("{not json", encoding="utf-8")
            cfg = load_config()
            self.assertEqual(len(cfg["projects"]), 3)
            self.assertTrue(cfg["excel_file"].endswith("heures_projet.xlsx"))

    def test_missing_name_entries_are_skipped(self) -> None:
        with patch.object(config_mod, "CONFIG_FILE", str(Path(self._tmp()) / "projects.json")):
            path = Path(config_mod.CONFIG_FILE)
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(
                json.dumps({
                    "projects": [
                        {"created": "2026-01-01T00:00:00"},
                        {"name": "OK", "created": "2026-01-01T00:00:00"},
                    ],
                    "sort_mode": "alpha_asc",
                    "excel_file": str(path.parent / "out.xlsx"),
                }),
                encoding="utf-8",
            )
            cfg = load_config()
            self.assertEqual([p["name"] for p in cfg["projects"]], ["OK"])

    def _tmp(self) -> str:
        import tempfile

        if not hasattr(self, "_tmpdir"):
            self._tmpdir = tempfile.mkdtemp()
        return self._tmpdir


if __name__ == "__main__":
    unittest.main()
