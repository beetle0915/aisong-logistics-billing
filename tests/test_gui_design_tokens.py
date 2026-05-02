from __future__ import annotations

import sys
import unittest
from pathlib import Path

APP_DIR = Path(__file__).resolve().parents[1] / "app"
sys.path.insert(0, str(APP_DIR))

from express_app.gui.design_tokens import GUI_COLORS, GUI_FONTS, GUI_LAYOUT  # noqa: E402


class GuiDesignTokensTest(unittest.TestCase):
    def test_core_colors_match_handoff_tokens(self) -> None:
        self.assertEqual(GUI_COLORS["accent"], "#1d4ed8")
        self.assertEqual(GUI_COLORS["background"], "#f4f6fb")
        self.assertEqual(GUI_COLORS["surface"], "#ffffff")

    def test_required_color_keys_are_available_for_main_flow(self) -> None:
        required_keys = {
            "background",
            "surface",
            "surface2",
            "sidebar",
            "border",
            "border2",
            "text",
            "muted",
            "dim",
            "accent",
            "accent_light",
            "accent_bg",
            "accent_border",
            "success",
            "success_bg",
            "warning",
            "warning_bg",
            "danger",
            "danger_bg",
            "log_bg",
            "log_text",
        }

        self.assertLessEqual(required_keys, set(GUI_COLORS))

    def test_layout_sizes_match_handoff_tokens(self) -> None:
        self.assertEqual(GUI_LAYOUT["app_width"], 1060)
        self.assertEqual(GUI_LAYOUT["app_height"], 700)
        self.assertEqual(GUI_LAYOUT["sidebar_width"], 180)

    def test_font_tokens_are_tkinter_friendly(self) -> None:
        self.assertEqual(GUI_FONTS["body"], "DM Sans")
        self.assertEqual(GUI_FONTS["mono"], "JetBrains Mono")
        self.assertEqual(GUI_FONTS["title_size"], 22)
        self.assertEqual(GUI_FONTS["base_size"], 12)
        self.assertEqual(GUI_FONTS["small_size"], 11)
        self.assertEqual(GUI_FONTS["metric_size"], 16)


if __name__ == "__main__":
    unittest.main()
