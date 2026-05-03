from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl

APP_DIR = Path(__file__).resolve().parents[1] / "app"
sys.path.insert(0, str(APP_DIR))

from express_app.gui import app as gui_app  # noqa: E402
from express_app.gui.app import ExpressFeeApp, V8_4_ENABLED_NAV_ITEMS  # noqa: E402


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


class FakeCombobox:
    def __init__(self) -> None:
        self.configured_values: list[str] = []
        self.options: dict[str, object] = {}

    def configure(self, **kwargs: object) -> None:
        self.options.update(kwargs)
        if "values" in kwargs:
            self.configured_values = list(kwargs["values"])  # type: ignore[arg-type]


class V84PricePreviewPageTest(unittest.TestCase):
    def _write_price_file(self, price_dir: Path, name: str = "客户A-快递报价.xlsx") -> None:
        workbook = openpyxl.Workbook()
        sf = workbook.active
        sf.title = "顺丰"
        sf.append(["省份参照列", "首重费用", "续重费用"])
        sf.append(["上海市", 6.5, 0.9])
        sf.append(["广东省", 6.9, 1.6])
        st = workbook.create_sheet("申通")
        st.append(["省份参照列", "首重费用", "续重费用"])
        st.append(["上海市", 3, 1])
        workbook.save(price_dir / name)

    def _prepare_preview_app(self, price_dir: Path) -> ExpressFeeApp:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        app.price_dir_var = FakeVar(str(price_dir))
        app.price_template_customer_var = FakeVar()
        app.price_template_status_var = FakeVar()
        app.price_template_summary_var = FakeVar()
        app.price_template_catalog = None
        app.price_template_workbook = None
        app.price_template_customers = []
        app.price_template_selected_sheet = ""
        app.price_template_sheet_tabs = []
        app.price_template_rows_by_sheet = {}
        app.price_template_trees = {}
        return app

    def test_price_preview_navigation_is_enabled(self) -> None:
        self.assertIn("报价预览", V8_4_ENABLED_NAV_ITEMS)

    def test_price_template_viewer_fields_match_v8_4_design(self) -> None:
        self.assertEqual(
            list(gui_app.V8_4_PRICE_TEMPLATE_ACTIONS),
            ["同步快递报价表", "业务员", "搜索"],
        )
        self.assertEqual(
            list(gui_app.V8_4_PRICE_TEMPLATE_COLUMNS),
            ["省份", "首重费用", "续重费用", "省份", "首重费用", "续重费用"],
        )
        self.assertEqual(
            list(gui_app.V8_4_PRICE_TEMPLATE_COLUMN_IDS),
            [
                "province_left",
                "first_price_left",
                "extra_price_left",
                "province_right",
                "first_price_right",
                "extra_price_right",
            ],
        )
        self.assertNotIn("重量", gui_app.V8_4_PRICE_TEMPLATE_ACTIONS)
        self.assertEqual(len(gui_app.V8_4_PRICE_TEMPLATE_COLUMNS), 6)

    def test_price_preview_uses_isolated_ui_styles(self) -> None:
        self.assertEqual(gui_app.PRICE_PREVIEW_COMBO_STYLE, "PricePreview.TCombobox")
        self.assertEqual(gui_app.PRICE_PREVIEW_NOTEBOOK_STYLE, "PricePreview.TNotebook")
        self.assertEqual(gui_app.PRICE_PREVIEW_TAB_PADDING, (18, 10))
        self.assertEqual(gui_app.PRICE_PREVIEW_TAB_EXPAND, (0, 0, 0, 0))
        self.assertEqual(gui_app.PRICE_PREVIEW_TREE_STYLE, "PricePreview.Treeview")
        self.assertEqual(gui_app.PRICE_PREVIEW_SCROLLBAR_STYLE, "PricePreview.Vertical.TScrollbar")

    def test_sync_price_templates_populates_customer_dropdown_state(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_text:
            price_dir = Path(temp_dir_text)
            self._write_price_file(price_dir, "客户B-快递报价.xlsx")
            self._write_price_file(price_dir, "客户A-快递报价.xlsx")
            app = self._prepare_preview_app(price_dir)

            app._sync_price_templates()

            self.assertEqual(app.price_template_customers, ["客户A", "客户B"])
            self.assertEqual(app.price_template_customer_var.get(), "客户A")
            self.assertIn("已识别 2 份客户报价表", app.price_template_status_var.get())

    def test_configure_price_template_combo_updates_values_without_ttk_assumptions(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        app.price_template_customers = ["客户A", "客户B"]
        app.price_template_combo = FakeCombobox()

        app._configure_price_template_combo()

        self.assertEqual(app.price_template_combo.configured_values, ["客户A", "客户B"])

    def test_search_price_template_loads_sheet_tabs_and_rows(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_text:
            price_dir = Path(temp_dir_text)
            self._write_price_file(price_dir)
            app = self._prepare_preview_app(price_dir)

            app._sync_price_templates()
            app._search_price_template()

            self.assertEqual(app.price_template_sheet_tabs, ["顺丰", "申通"])
            self.assertEqual(app.price_template_selected_sheet, "顺丰")
            self.assertEqual(
                app.price_template_rows_by_sheet["顺丰"][0],
                ("上海市", "¥6.50", "¥0.90", "广东省", "¥6.90", "¥1.60"),
            )
            self.assertEqual(
                app.price_template_rows_by_sheet["申通"][0],
                ("上海市", "¥3.00", "¥1.00", "", "", ""),
            )
            self.assertIn("客户A", app.price_template_summary_var.get())
            self.assertIn("2 个快递公司", app.price_template_summary_var.get())

    def test_search_price_template_requires_customer_selection(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_text:
            app = self._prepare_preview_app(Path(temp_dir_text))

            app._search_price_template()

            self.assertIn("请先选择业务员", app.price_template_status_var.get())

    def test_show_price_preview_page_updates_active_nav_title_and_raises_page(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        app.active_nav_var = FakeVar("费用计算")
        app.module_title_var = FakeVar("费用计算")
        app.module_subtitle_var = FakeVar("")
        app.nav_labels = {
            "费用计算": FakeWidget(),
            "账户余额": FakeWidget(),
            "报价预览": FakeWidget(),
            "系统设置": FakeWidget(),
        }
        app.fee_page = FakeWidget()
        app.balance_page = FakeWidget()
        app.price_preview_page = FakeWidget()
        app.settings_page = FakeWidget()

        app._show_page("报价预览")

        self.assertEqual(app.active_nav_var.get(), "报价预览")
        self.assertEqual(app.module_title_var.get(), "报价预览")
        self.assertIn("报价表模板", app.module_subtitle_var.get())
        self.assertEqual(app.nav_labels["报价预览"].options["style"], "NavActive.TLabel")
        self.assertTrue(app.price_preview_page.raised)
        self.assertFalse(app.fee_page.raised)
        self.assertFalse(app.balance_page.raised)


if __name__ == "__main__":
    unittest.main()
