"""Tests unitaires pour la palette THEME exportée par main."""

from __future__ import annotations

import re
import unittest

from main import THEME, THEME_REQUIRED_KEYS

_HEX_COLOR = re.compile(r"^#[0-9A-Fa-f]{6}$")


class TestTheme(unittest.TestCase):
    def test_required_keys_present_in_theme(self) -> None:
        missing = [key for key in THEME_REQUIRED_KEYS if key not in THEME]
        self.assertEqual(
            missing,
            [],
            f"Clés manquantes dans THEME: {missing}",
        )

    def test_theme_values_are_hex_colors(self) -> None:
        invalid: list[str] = []
        for key in THEME_REQUIRED_KEYS:
            value = THEME[key]
            if not isinstance(value, str) or _HEX_COLOR.fullmatch(value) is None:
                invalid.append(f"{key}={value!r}")
        self.assertEqual(
            invalid,
            [],
            f"Valeurs non hex #RRGGBB: {invalid}",
        )

    def test_palette_accent_colors_are_distinct(self) -> None:
        self.assertNotEqual(
            THEME["green"],
            THEME["green_dark"],
            "green et green_dark ne doivent pas être identiques",
        )
        self.assertNotEqual(
            THEME["green"],
            THEME["danger"],
            "green et danger ne doivent pas être identiques",
        )
        self.assertNotEqual(
            THEME["green_dark"],
            THEME["danger"],
            "green_dark et danger ne doivent pas être identiques",
        )


if __name__ == "__main__":
    unittest.main()
