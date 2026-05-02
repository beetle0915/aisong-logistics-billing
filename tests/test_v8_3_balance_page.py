from __future__ import annotations

from datetime import date
import sys
import unittest
from pathlib import Path

APP_DIR = Path(__file__).resolve().parents[1] / "app"
sys.path.insert(0, str(APP_DIR))

from express_app.core.account_balance import CustomerBalanceRecord  # noqa: E402
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


class FakeTree:
    def __init__(self) -> None:
        self.items: dict[str, tuple[object, ...]] = {}
        self.selected: str | None = None

    def get_children(self) -> tuple[str, ...]:
        return tuple(self.items)

    def delete(self, item_id: str) -> None:
        self.items.pop(item_id, None)

    def insert(
        self,
        parent: str,
        index: object,
        iid: str,
        values: tuple[object, ...],
    ) -> str:
        self.items[iid] = values
        return iid

    def selection_set(self, item_id: str) -> None:
        self.selected = item_id


class V83BalancePageTest(unittest.TestCase):
    def test_balance_page_is_enabled_navigation_item(self) -> None:
        self.assertIn("账户余额", V8_3_ENABLED_NAV_ITEMS)

    def test_balance_table_columns_match_v8_3_handoff(self) -> None:
        self.assertEqual(
            list(gui_app.V8_3_BALANCE_TABLE_COLUMNS),
            ["客户", "累计消费", "累计收款", "当前余额", "最近日期", "状态", "文件路径"],
        )

    def test_v8_3_1_balance_page_moves_refresh_to_footer_actions(self) -> None:
        self.assertEqual(
            list(gui_app.V8_3_1_BALANCE_TOP_ACTIONS),
            [],
        )
        self.assertEqual(
            list(gui_app.V8_3_1_BALANCE_FOOTER_ACTIONS),
            [
                ("刷新数据", "Primary.TButton"),
                ("打开客户目录", "Secondary.TButton"),
                ("打开历史汇总表", "Secondary.TButton"),
            ],
        )

    def test_v8_3_1_balance_refresh_shows_all_records_without_hidden_search(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        app.balance_tree = FakeTree()
        app.balance_history_paths = {}
        app.balance_customer_dirs = {}
        app.balance_records = [
            CustomerBalanceRecord(
                customer="张三",
                total_consumed=100.0,
                total_paid=80.0,
                current_balance=-20.0,
                last_date=date(2026, 4, 7),
                history_file=Path("/tmp/张三_客户快递费历史汇总.xlsx"),
                customer_dir=Path("/tmp/张三"),
            ),
            CustomerBalanceRecord(
                customer="李四",
                total_consumed=50.0,
                total_paid=60.0,
                current_balance=10.0,
                last_date=date(2026, 4, 8),
                history_file=Path("/tmp/李四_客户快递费历史汇总.xlsx"),
                customer_dir=Path("/tmp/李四"),
            ),
        ]

        app._refresh_balance_table()

        self.assertEqual(list(app.balance_tree.items), ["balance-1", "balance-2"])
        self.assertEqual(app.balance_tree.items["balance-1"][0], "张三")
        self.assertEqual(app.balance_tree.items["balance-2"][0], "李四")
        self.assertEqual(app.balance_tree.selected, "balance-1")

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
