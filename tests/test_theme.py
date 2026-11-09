"""Thème light + version app."""

from __future__ import annotations

import unittest

from time_tracker.constants import APP_VERSION
from time_tracker.theme import THEME, contrast_ratio


class TestConstants(unittest.TestCase):
    def test_app_version_is_semver_like(self) -> None:
        parts = APP_VERSION.split(".")
        self.assertGreaterEqual(len(parts), 2)
        for part in parts:
            self.assertTrue(part.isdigit(), f"segment non numérique: {part}")


class TestTheme(unittest.TestCase):
    def test_light_palette_tokens(self) -> None:
        self.assertEqual(THEME.bg, "#F4F6F8")
        self.assertEqual(THEME.surface, "#FFFFFF")
        self.assertEqual(THEME.accent, "#0D9488")
        self.assertEqual(THEME.danger, "#E11D48")
        # Fond clair : luminance élevée
        from time_tracker.theme import relative_luminance

        self.assertGreater(relative_luminance(THEME.bg), 0.7)

    def test_muted_on_bg_meets_aa_contrast(self) -> None:
        ratio = contrast_ratio(THEME.muted, THEME.bg)
        self.assertGreaterEqual(ratio, 4.5, f"contraste muted/bg trop faible: {ratio:.2f}")

    def test_text_on_bg_meets_aa_contrast(self) -> None:
        ratio = contrast_ratio(THEME.text, THEME.bg)
        self.assertGreaterEqual(ratio, 4.5, f"contraste text/bg trop faible: {ratio:.2f}")

    def test_accent_on_white_meets_aa_for_large_text(self) -> None:
        # Boutons : texte blanc sur accent — seuil AA large 3:1
        ratio = contrast_ratio(THEME.on_accent, THEME.accent)
        self.assertGreaterEqual(ratio, 3.0, f"contraste on_accent/accent trop faible: {ratio:.2f}")

    def test_pill_success_on_ok_bg_meets_aa_large(self) -> None:
        ratio = contrast_ratio(THEME.success, THEME.pill_ok_bg)
        self.assertGreaterEqual(ratio, 3.0, f"contraste success/pill_ok trop faible: {ratio:.2f}")

    def test_pill_warning_on_paused_bg_meets_aa_large(self) -> None:
        ratio = contrast_ratio(THEME.warning_dim, THEME.pill_paused_bg)
        self.assertGreaterEqual(
            ratio, 3.0, f"contraste warning_dim/pill_paused trop faible: {ratio:.2f}"
        )

    def test_pill_danger_on_error_bg_meets_aa_large(self) -> None:
        ratio = contrast_ratio(THEME.danger, THEME.pill_error_bg)
        self.assertGreaterEqual(ratio, 3.0, f"contraste danger/pill_error trop faible: {ratio:.2f}")


if __name__ == "__main__":
    unittest.main()
