"""Normalisation et conflits des raccourcis."""

from __future__ import annotations

import sys
import types
import unittest
from unittest.mock import patch

from time_tracker.hotkeys import (
    GlobalHotkeyController,
    combo_from_tk_event,
    default_hotkeys,
    find_hotkey_conflicts,
    format_hotkey_display,
    hotkey_to_pynput,
    hotkey_to_tk,
    normalize_hotkey,
    parse_hotkeys,
)


class TestHotkeys(unittest.TestCase):
    def test_normalize_hotkey_orders_mods(self) -> None:
        self.assertEqual(normalize_hotkey("Shift+Ctrl+Space"), "ctrl+shift+space")
        self.assertEqual(normalize_hotkey("  CTRL + ENTER "), "ctrl+enter")
        self.assertIsNone(normalize_hotkey("ctrl"))
        self.assertIsNone(normalize_hotkey(""))
        self.assertIsNone(normalize_hotkey("a"))
        self.assertEqual(normalize_hotkey("f5"), "f5")

    def test_parse_hotkeys_merges_defaults(self) -> None:
        parsed = parse_hotkeys({"toggle": "alt+f1", "unknown": "x"})
        self.assertEqual(parsed["toggle"], "alt+f1")
        self.assertEqual(parsed["stop"], default_hotkeys()["stop"])

    def test_parse_invalid_keeps_default(self) -> None:
        parsed = parse_hotkeys({"toggle": "ctrl"})
        self.assertEqual(parsed["toggle"], default_hotkeys()["toggle"])

    def test_conflicts(self) -> None:
        hotkeys = default_hotkeys()
        hotkeys["stop"] = hotkeys["toggle"]
        self.assertEqual(find_hotkey_conflicts(hotkeys), [hotkeys["toggle"]])

    def test_conversions(self) -> None:
        self.assertEqual(hotkey_to_tk("ctrl+shift+space"), "<Control-Shift-space>")
        self.assertEqual(hotkey_to_pynput("ctrl+shift+space"), "<ctrl>+<shift>+<space>")
        self.assertEqual(format_hotkey_display("ctrl+shift+space"), "Ctrl+Shift+Espace")

    def test_combo_from_tk_event_windows_ignores_numlock(self) -> None:
        # 0x8 = NumLock sous Windows ; Ctrl+Space avec NumLock = 0x4|0x8
        with patch("time_tracker.hotkeys.sys.platform", "win32"):
            self.assertEqual(combo_from_tk_event("space", 0x4 | 0x8), "ctrl+space")
            self.assertEqual(combo_from_tk_event("space", 0x4 | 0x20000), "ctrl+alt+space")
            self.assertIsNone(combo_from_tk_event("Escape", 0))

    def test_combo_from_tk_event_linux_alt_bit(self) -> None:
        with patch("time_tracker.hotkeys.sys.platform", "linux"):
            self.assertEqual(combo_from_tk_event("d", 0x4 | 0x8), "ctrl+alt+d")

    def test_global_start_rejects_untrusted_darwin(self) -> None:
        calls: list[str] = []

        class FakeListener:
            IS_TRUSTED = False

            def start(self) -> None:
                return None

            def stop(self) -> None:
                calls.append("stop")

        keyboard_mod = types.SimpleNamespace(
            GlobalHotKeys=lambda _mapping: FakeListener(),
            Listener=types.SimpleNamespace(IS_TRUSTED=False),
        )
        pynput_mod = types.ModuleType("pynput")
        pynput_mod.keyboard = keyboard_mod  # type: ignore[attr-defined]

        controller = GlobalHotkeyController(
            schedule=lambda cb: None,
            handlers={"toggle": lambda: None},
        )
        with (
            patch("time_tracker.hotkeys.sys.platform", "darwin"),
            patch.dict(sys.modules, {"pynput": pynput_mod}),
        ):
            error = controller.start(default_hotkeys())

        self.assertIsNotNone(error)
        assert error is not None
        self.assertIn("Permissions macOS", error)
        self.assertEqual(calls, ["stop"])


if __name__ == "__main__":
    unittest.main()
