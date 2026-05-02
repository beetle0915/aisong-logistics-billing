from __future__ import annotations

import sys
import unittest
from pathlib import Path

APP_DIR = Path(__file__).resolve().parents[1] / "app"
sys.path.insert(0, str(APP_DIR))

from express_app.gui.app import (  # noqa: E402
    EXPRESS_MAPPING_HELP_TEXT,
    KEYWORD_MAPPING_HELP_TEXT,
    RULE_WINDOW_TITLE,
    format_option_label,
)


class GuiOptionLabelTest(unittest.TestCase):
    def test_selected_option_uses_check_mark(self) -> None:
        label = format_option_label("生成客户历史汇总", True)

        self.assertEqual(label, "✅ 生成客户历史汇总")
        self.assertNotIn("❌", label)

    def test_unselected_option_uses_empty_box_not_cross_mark(self) -> None:
        label = format_option_label("生成客户历史汇总", False)

        self.assertEqual(label, "□ 生成客户历史汇总")
        self.assertNotIn("❌", label)

    def test_rule_config_copy_explains_core_express_mapping(self) -> None:
        self.assertEqual(RULE_WINDOW_TITLE, "快递识别与大件规则")
        self.assertIn("销售表", EXPRESS_MAPPING_HELP_TEXT)
        self.assertIn("报价表 sheet", EXPRESS_MAPPING_HELP_TEXT)
        self.assertIn("原始快递名称=标准快递名称", EXPRESS_MAPPING_HELP_TEXT)
        self.assertIn("兜底识别", KEYWORD_MAPPING_HELP_TEXT)


if __name__ == "__main__":
    unittest.main()
