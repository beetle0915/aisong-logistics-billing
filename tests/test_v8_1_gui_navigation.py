from __future__ import annotations

import sys
import unittest
from pathlib import Path

APP_DIR = Path(__file__).resolve().parents[1] / "app"
sys.path.insert(0, str(APP_DIR))

from express_app.gui.app import (  # noqa: E402
    V8_1_MAIN_NAV_ITEMS,
    V8_1_WORKFLOW_STEPS,
)


class V81GuiNavigationCopyTest(unittest.TestCase):
    def test_main_navigation_exposes_v8_1_sections(self) -> None:
        self.assertEqual(
            list(V8_1_MAIN_NAV_ITEMS),
            [
                "费用计算",
                "账户余额",
                "报价预览",
                "系统设置",
            ],
        )
        self.assertNotIn("客户档案", V8_1_MAIN_NAV_ITEMS)

    def test_workflow_steps_expose_v8_1_fee_calculation_flow(self) -> None:
        self.assertEqual(
            list(V8_1_WORKFLOW_STEPS),
            [
                "配置",
                "运行",
                "结果",
            ],
        )


if __name__ == "__main__":
    unittest.main()
