from __future__ import annotations

import sys
import unittest
from pathlib import Path

APP_DIR = Path(__file__).resolve().parents[1] / "app"
sys.path.insert(0, str(APP_DIR))

from express_app.gui import app as gui_app  # noqa: E402
from express_app.gui.app import ExpressFeeApp, V8_3_ENABLED_NAV_ITEMS  # noqa: E402


class FakeVar:
    def __init__(self, value: str = "") -> None:
        self.value = value

    def get(self) -> str:
        return self.value

    def set(self, value: str) -> None:
        self.value = value


class FakeWidget:
    def __init__(self) -> None:
        self.options: dict[str, object] = {}
        self.raised = False

    def configure(self, **kwargs: object) -> None:
        self.options.update(kwargs)

    def tkraise(self) -> None:
        self.raised = True


class V83BalancePageTest(unittest.TestCase):
    def test_balance_page_is_enabled_navigation_item(self) -> None:
        self.assertIn("账户余额", V8_3_ENABLED_NAV_ITEMS)

    def test_balance_table_columns_match_v8_3_handoff(self) -> None:
        self.assertEqual(
            list(gui_app.V8_3_BALANCE_TABLE_COLUMNS),
            ["客户", "累计消费", "累计收款", "当前余额", "最近日期", "状态", "文件路径"],
        )

    def test_show_balance_page_updates_active_nav_title_and_raises_page(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        app.active_nav_var = FakeVar("费用计算")
        app.module_title_var = FakeVar("费用计算")
        app.module_subtitle_var = FakeVar("")
        app.nav_labels = {
            "费用计算": FakeWidget(),
            "账户余额": FakeWidget(),
            "系统设置": FakeWidget(),
        }
        app.fee_page = FakeWidget()
        app.balance_page = FakeWidget()
        app.settings_page = FakeWidget()

        app._show_page("账户余额")

        self.assertEqual(app.active_nav_var.get(), "账户余额")
        self.assertEqual(app.module_title_var.get(), "账户余额")
        self.assertEqual(app.nav_labels["账户余额"].options["style"], "NavActive.TLabel")
        self.assertEqual(app.nav_labels["费用计算"].options["style"], "NavDisabled.TLabel")
        self.assertTrue(app.balance_page.raised)
        self.assertFalse(app.fee_page.raised)
        self.assertFalse(app.settings_page.raised)


if __name__ == "__main__":
    unittest.main()
