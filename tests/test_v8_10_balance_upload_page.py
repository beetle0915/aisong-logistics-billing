from __future__ import annotations

from datetime import date
import sys
import unittest
from pathlib import Path

APP_DIR = Path(__file__).resolve().parents[1] / "app"
sys.path.insert(0, str(APP_DIR))

from express_app.core.balance_upload import BalanceUploadResult  # noqa: E402
from express_app.core.balance_upload import BalanceUploadPreview, BalanceUploadRecord  # noqa: E402
from express_app.gui import app as gui_app  # noqa: E402
from express_app.gui.app import ExpressFeeApp  # noqa: E402
from tests.test_v8_3_balance_page import FakeTree, FakeVar, FakeWidget  # noqa: E402


class FakeButton:
    def __init__(self) -> None:
        self.options: dict[str, object] = {}

    def configure(self, **kwargs: object) -> None:
        self.options.update(kwargs)


class FakeQueue:
    def __init__(self, items: list[tuple[str, object]]) -> None:
        self.items = list(items)

    def get_nowait(self) -> tuple[str, object]:
        if not self.items:
            raise gui_app.queue.Empty
        return self.items.pop(0)


class V810BalanceUploadPageTest(unittest.TestCase):
    def test_balance_upload_navigation_is_between_bill_splitter_and_settings(self) -> None:
        self.assertEqual(
            list(gui_app.V8_10_MAIN_NAV_ITEMS),
            ["费用计算", "账户余额", "报价预览", "拆分账单", "余额上传", "系统设置"],
        )
        self.assertEqual(gui_app.V8_9_MAIN_NAV_ITEMS, gui_app.V8_10_MAIN_NAV_ITEMS)
        self.assertLess(
            gui_app.V8_10_MAIN_NAV_ITEMS.index("拆分账单"),
            gui_app.V8_10_MAIN_NAV_ITEMS.index("余额上传"),
        )
        self.assertLess(
            gui_app.V8_10_MAIN_NAV_ITEMS.index("余额上传"),
            gui_app.V8_10_MAIN_NAV_ITEMS.index("系统设置"),
        )

    def test_balance_upload_page_constants_match_simplified_handoff(self) -> None:
        self.assertEqual(
            list(gui_app.V8_10_BALANCE_UPLOAD_COLUMNS),
            ["客户", "今日快递费消费", "今日余额", "余额日期", "状态"],
        )
        self.assertEqual(
            list(gui_app.V8_10_BALANCE_UPLOAD_ACTIONS),
            ["读取数据", "确认并上传"],
        )
        self.assertIn("余额上传", gui_app.V8_10_SETTINGS_SECTIONS)

    def test_gui_module_exposes_datetime_for_balance_upload_startup_defaults(self) -> None:
        self.assertTrue(hasattr(gui_app, "datetime"))

    def test_show_balance_upload_page_updates_title_and_raises_page(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        app.active_nav_var = FakeVar("费用计算")
        app.module_title_var = FakeVar("费用计算")
        app.module_subtitle_var = FakeVar("")
        app.nav_labels = {
            "费用计算": FakeWidget(),
            "账户余额": FakeWidget(),
            "报价预览": FakeWidget(),
            "拆分账单": FakeWidget(),
            "余额上传": FakeWidget(),
            "系统设置": FakeWidget(),
        }
        app.fee_page = FakeWidget()
        app.balance_page = FakeWidget()
        app.price_preview_page = FakeWidget()
        app.bill_splitter_page = FakeWidget()
        app.balance_upload_page = FakeWidget()
        app.settings_page = FakeWidget()

        app._show_page("余额上传")

        self.assertEqual(app.active_nav_var.get(), "余额上传")
        self.assertEqual(app.module_title_var.get(), "余额上传")
        self.assertIn("上传日期", app.module_subtitle_var.get())
        self.assertTrue(app.balance_upload_page.raised)
        self.assertFalse(app.bill_splitter_page.raised)
        self.assertFalse(app.settings_page.raised)
        self.assertEqual(app.nav_labels["余额上传"].options["style"], "NavActive.TLabel")

    def test_apply_balance_upload_preview_populates_metrics_and_table(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        app.balance_upload_tree = FakeTree()
        app.balance_upload_preview = None
        app.balance_upload_customer_count_var = FakeVar()
        app.balance_upload_today_fee_var = FakeVar()
        app.balance_upload_today_balance_var = FakeVar()
        app.balance_upload_debtor_count_var = FakeVar()
        app.balance_upload_status_var = FakeVar()
        app.balance_upload_confirm_var = FakeVar(False)
        app.balance_upload_button = FakeButton()
        app.balance_upload_uploading = False
        app.settings_balance_upload_url_var = FakeVar("https://api.example.test/balance")
        app.settings_balance_upload_token_var = FakeVar("secret-token")
        preview = BalanceUploadPreview(
            upload_date=date(2026, 5, 9),
            source_dir=Path("/tmp/customers"),
            records=[
                BalanceUploadRecord("客户A", 120.0, 300.0, date(2026, 5, 9), "充足", Path("/tmp/a.xlsx")),
                BalanceUploadRecord("客户B", 0.0, -50.0, date(2026, 5, 8), "欠款", Path("/tmp/b.xlsx")),
            ],
        )

        app._apply_balance_upload_preview(preview)

        self.assertEqual(app.balance_upload_customer_count_var.get(), "2 位")
        self.assertEqual(app.balance_upload_today_fee_var.get(), "¥120.00")
        self.assertEqual(app.balance_upload_today_balance_var.get(), "¥250.00")
        self.assertEqual(app.balance_upload_debtor_count_var.get(), "1 位")
        self.assertEqual(app.balance_upload_status_var.get(), "读取完成，共 2 位客户；其中 1 位沿用最近余额日期")
        self.assertEqual(app.balance_upload_tree.items["balance-upload-1"][0], "客户A")
        self.assertEqual(app.balance_upload_tree.items["balance-upload-1"][1], "¥120.00")
        self.assertEqual(app.balance_upload_tree.items["balance-upload-2"][1], "¥0.00")
        self.assertEqual(app.balance_upload_tree.items["balance-upload-2"][2], "-¥50.00")
        self.assertEqual(app.balance_upload_tree.items["balance-upload-2"][3], "2026-05-08")
        self.assertEqual(app.balance_upload_button.options["state"], "disabled")

        app.balance_upload_confirm_var.set(True)
        app._refresh_balance_upload_action_state()

        self.assertEqual(app.balance_upload_button.options["state"], "normal")

    def test_failed_balance_upload_keeps_short_status_and_reenables_upload_button(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        app._queue = FakeQueue(
            [
                (
                    "balance_upload_done",
                    BalanceUploadResult(
                        ok=False,
                        status_code=401,
                        message='服务器返回错误：401，{"timestamp":"2026-06-01T15:30:08.236+00:00","status":401,"error":"Unauthorized"}',
                    ),
                )
            ]
        )
        app._worker = None
        app.balance_upload_uploading = True
        app.balance_upload_status_var = FakeVar()
        app.status_var = FakeVar()
        app.balance_upload_button = FakeButton()
        app.balance_upload_read_button = FakeButton()
        app.balance_upload_preview = BalanceUploadPreview(
            upload_date=date(2026, 6, 1),
            source_dir=Path("/tmp/customers"),
            records=[
                BalanceUploadRecord("客户A", 0.0, -200.0, date(2026, 5, 28), "欠款", Path("/tmp/a.xlsx")),
            ],
        )
        app.balance_upload_confirm_var = FakeVar(True)
        app.settings_balance_upload_url_var = FakeVar("https://api.example.test/balance")
        app.settings_balance_upload_token_var = FakeVar("fixed-token")
        app._append_balance_upload_log = lambda _text: None  # type: ignore[method-assign]
        app.after = lambda *_args: None  # type: ignore[method-assign]

        original_showerror = gui_app.messagebox.showerror
        try:
            gui_app.messagebox.showerror = lambda *_args, **_kwargs: None  # type: ignore[assignment]
            app._poll_queue()
        finally:
            gui_app.messagebox.showerror = original_showerror  # type: ignore[assignment]

        self.assertEqual(app.balance_upload_status_var.get(), "上传失败，请查看日志")
        self.assertEqual(app.status_var.get(), "余额上传失败")
        self.assertFalse(app.balance_upload_uploading)
        self.assertEqual(app.balance_upload_button.options["state"], "normal")
        self.assertEqual(app.balance_upload_read_button.options["state"], "normal")

    def test_balance_upload_button_does_not_stay_disabled_while_endpoint_settings_are_being_fixed(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        app.balance_upload_button = FakeButton()
        app.balance_upload_read_button = FakeButton()
        app.balance_upload_uploading = False
        app.balance_upload_preview = BalanceUploadPreview(
            upload_date=date(2026, 6, 1),
            source_dir=Path("/tmp/customers"),
            records=[
                BalanceUploadRecord("客户A", 0.0, -200.0, date(2026, 5, 28), "欠款", Path("/tmp/a.xlsx")),
            ],
        )
        app.balance_upload_confirm_var = FakeVar(True)
        app.settings_balance_upload_url_var = FakeVar("")
        app.settings_balance_upload_token_var = FakeVar("")

        app._refresh_balance_upload_action_state()

        self.assertEqual(app.balance_upload_button.options["state"], "normal")
        self.assertEqual(app.balance_upload_read_button.options["state"], "normal")


if __name__ == "__main__":
    unittest.main()
