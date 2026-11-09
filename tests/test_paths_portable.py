"""Mode portable et chemins de données."""

from __future__ import annotations

import unittest
from unittest.mock import patch

from time_tracker import paths as paths_mod


class TestPortablePaths(unittest.TestCase):
    def tearDown(self) -> None:
        paths_mod.set_portable_forced(False)
        paths_mod.refresh_data_paths()

    def test_portable_flag_forces_mode(self) -> None:
        paths_mod.set_portable_forced(True)
        self.assertTrue(paths_mod.is_portable_mode())

    def test_refresh_updates_config_file(self) -> None:
        with patch.object(paths_mod, "is_portable_mode", return_value=False):
            paths_mod.refresh_data_paths()
            self.assertTrue(paths_mod.CONFIG_FILE.endswith("projects.json"))


if __name__ == "__main__":
    unittest.main()
