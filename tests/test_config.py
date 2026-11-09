"""Robustesse de load_config face à un JSON invalide / écriture atomique."""

from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from time_tracker.config import load_config, load_config_detailed, save_config
from time_tracker.constants import ProjectsConfig


class TestLoadConfig(unittest.TestCase):
    def _tmp(self) -> Path:
        if not hasattr(self, "_tmpdir"):
            self._tmpdir = Path(tempfile.mkdtemp())
        return self._tmpdir

    def test_corrupt_json_quarantines_and_does_not_overwrite(self) -> None:
        path = self._tmp() / "projects.json"
        path.write_text("{not json", encoding="utf-8")
        result = load_config_detailed(str(path))
        self.assertTrue(result.was_corrupt)
        self.assertEqual(len(result.config["projects"]), 3)
        corrupt_files = list(self._tmp().glob("projects.json.corrupt.*"))
        self.assertEqual(len(corrupt_files), 1)
        self.assertFalse(path.exists())

    def test_missing_name_entries_are_skipped(self) -> None:
        path = self._tmp() / "projects.json"
        path.write_text(
            json.dumps(
                {
                    "projects": [
                        {"created": "2026-01-01T00:00:00"},
                        {"name": "OK", "created": "2026-01-01T00:00:00"},
                    ],
                    "sort_mode": "alpha_asc",
                    "excel_file": str(path.parent / "out.xlsx"),
                    "launch_at_startup": False,
                }
            ),
            encoding="utf-8",
        )
        cfg = load_config(str(path))
        self.assertEqual([p["name"] for p in cfg["projects"]], ["OK"])

    def test_empty_projects_list_preserved(self) -> None:
        path = self._tmp() / "projects.json"
        path.write_text(
            json.dumps(
                {
                    "projects": [],
                    "sort_mode": "alpha_asc",
                    "excel_file": str(path.parent / "out.xlsx"),
                    "launch_at_startup": False,
                }
            ),
            encoding="utf-8",
        )
        cfg = load_config(str(path))
        self.assertEqual(cfg["projects"], [])

    def test_save_config_atomic_roundtrip(self) -> None:
        path = self._tmp() / "projects.json"
        typed: ProjectsConfig = {
            "projects": [{"name": "A", "created": "2026-01-01T00:00:00"}],
            "sort_mode": "alpha_asc",
            "excel_file": str(path.parent / "out.xlsx"),
            "launch_at_startup": True,
            "always_on_top": False,
            "window_geometry": "320x560+10+20",
            "hotkeys": {
                "toggle": "ctrl+shift+space",
                "stop": "ctrl+shift+enter",
                "settings": "ctrl+shift+comma",
                "show": "ctrl+shift+t",
                "today": "ctrl+shift+d",
                "period": "ctrl+shift+p",
                "manual": "ctrl+shift+m",
            },
        }
        save_config(typed, str(path))
        loaded = load_config(str(path))
        self.assertEqual(loaded["projects"][0]["name"], "A")
        self.assertTrue(loaded["launch_at_startup"])
        self.assertFalse(loaded["always_on_top"])
        self.assertEqual(loaded["window_geometry"], "320x560+10+20")
        self.assertEqual(loaded["hotkeys"]["toggle"], "ctrl+shift+space")
        self.assertFalse(Path(str(path) + ".tmp").exists())

    def test_always_on_top_defaults_true_when_absent(self) -> None:
        path = self._tmp() / "projects.json"
        path.write_text(
            json.dumps(
                {
                    "projects": [{"name": "A", "created": "2026-01-01T00:00:00"}],
                    "sort_mode": "alpha_asc",
                    "excel_file": str(path.parent / "out.xlsx"),
                    "launch_at_startup": False,
                }
            ),
            encoding="utf-8",
        )
        cfg = load_config(str(path))
        self.assertTrue(cfg["always_on_top"])
        self.assertIsNone(cfg["window_geometry"])
        self.assertEqual(cfg["hotkeys"]["toggle"], "ctrl+shift+space")

    def test_custom_hotkeys_roundtrip(self) -> None:
        path = self._tmp() / "projects.json"
        path.write_text(
            json.dumps(
                {
                    "projects": [{"name": "A", "created": "2026-01-01T00:00:00"}],
                    "sort_mode": "alpha_asc",
                    "excel_file": str(path.parent / "out.xlsx"),
                    "launch_at_startup": False,
                    "hotkeys": {"toggle": "Alt+F2", "stop": "ctrl+shift+q"},
                }
            ),
            encoding="utf-8",
        )
        cfg = load_config(str(path))
        self.assertEqual(cfg["hotkeys"]["toggle"], "alt+f2")
        self.assertEqual(cfg["hotkeys"]["stop"], "ctrl+shift+q")
        self.assertEqual(cfg["hotkeys"]["settings"], "ctrl+shift+comma")

    def test_fallback_config_always_on_top_true(self) -> None:
        path = self._tmp() / "missing.json"
        result = load_config_detailed(str(path))
        self.assertTrue(result.was_missing)
        self.assertTrue(result.config["always_on_top"])

    def test_bool_string_false_falls_back_to_default(self) -> None:
        path = self._tmp() / "projects.json"
        path.write_text(
            json.dumps(
                {
                    "projects": [{"name": "A", "created": "2026-01-01T00:00:00"}],
                    "sort_mode": "alpha_asc",
                    "excel_file": str(path.parent / "out.xlsx"),
                    "launch_at_startup": "false",
                    "always_on_top": "false",
                }
            ),
            encoding="utf-8",
        )
        cfg = load_config(str(path))
        self.assertFalse(cfg["launch_at_startup"])
        self.assertTrue(cfg["always_on_top"])

    def test_bool_int_zero_one_accepted(self) -> None:
        path = self._tmp() / "projects.json"
        path.write_text(
            json.dumps(
                {
                    "projects": [{"name": "A", "created": "2026-01-01T00:00:00"}],
                    "sort_mode": "alpha_asc",
                    "excel_file": str(path.parent / "out.xlsx"),
                    "launch_at_startup": 1,
                    "always_on_top": 0,
                }
            ),
            encoding="utf-8",
        )
        cfg = load_config(str(path))
        self.assertTrue(cfg["launch_at_startup"])
        self.assertFalse(cfg["always_on_top"])


if __name__ == "__main__":
    unittest.main()
