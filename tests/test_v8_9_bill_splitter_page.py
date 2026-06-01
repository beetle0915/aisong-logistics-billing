from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace

APP_DIR = Path(__file__).resolve().parents[1] / "app"
sys.path.insert(0, str(APP_DIR))

from express_app.gui import app as gui_app  # noqa: E402
from express_app.gui.app import ExpressFeeApp, V8_9_ENABLED_NAV_ITEMS  # noqa: E402


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


class FakeButton(FakeWidget):
    pass


class FakeFileScan(SimpleNamespace):
    def __init__(self, *, error: str = "") -> None:
        super().__init__(error=error)


class FakeCombobox:
    def __init__(self) -> None:
        self.configured_values: list[str] = []
        self.options: dict[str, object] = {}

    def configure(self, **kwargs: object) -> None:
        self.options.update(kwargs)
        if "values" in kwargs:
            self.configured_values = list(kwargs["values"])  # type: ignore[arg-type]


class FakeTree:
    def __init__(self) -> None:
        self.rows: list[tuple[object, ...]] = []
        self.items: dict[str, tuple[object, ...]] = {}
        self.selected: list[str] = []

    def get_children(self) -> list[int]:
        return list(range(len(self.rows)))

    def delete(self, *_items: object) -> None:
        self.rows.clear()
        self.items.clear()
        self.selected.clear()

    def insert(self, _parent: str, _index: str, iid: str | None = None, values: tuple[object, ...] = ()) -> str:
        self.rows.append(tuple(values))
        item_id = iid or str(len(self.rows) - 1)
        self.items[item_id] = tuple(values)
        return item_id

    def selection_set(self, item_id: str) -> None:
        self.selected = [item_id]


class FakeLogText:
    def __init__(self) -> None:
        self.text = ""
        self.state = ""
        self.seen_end = False

    def configure(self, **kwargs: object) -> None:
        if "state" in kwargs:
            self.state = str(kwargs["state"])

    def index(self, _index: str) -> str:
        return "1.0" if not self.text else "1.1"

    def insert(self, _index: str, text: str) -> None:
        self.text += text

    def delete(self, _start: str, _end: str) -> None:
        self.text = ""

    def see(self, _index: str) -> None:
        self.seen_end = True


class V89BillSplitterPageTest(unittest.TestCase):
    def test_bill_splitter_navigation_is_between_price_preview_and_settings(self) -> None:
        self.assertEqual(
            list(gui_app.V8_9_MAIN_NAV_ITEMS),
            ["费用计算", "账户余额", "报价预览", "拆分账单", "余额上传", "系统设置"],
        )
        self.assertEqual(
            list(V8_9_ENABLED_NAV_ITEMS),
            ["费用计算", "账户余额", "报价预览", "拆分账单", "余额上传", "系统设置"],
        )
        self.assertLess(
            gui_app.V8_9_MAIN_NAV_ITEMS.index("报价预览"),
            gui_app.V8_9_MAIN_NAV_ITEMS.index("拆分账单"),
        )
        self.assertLess(
            gui_app.V8_9_MAIN_NAV_ITEMS.index("拆分账单"),
            gui_app.V8_9_MAIN_NAV_ITEMS.index("系统设置"),
        )

    def test_bill_splitter_page_constants_match_v8_9_design(self) -> None:
        self.assertEqual(gui_app.V8_9_BILL_SPLIT_TARGET_SHEET_NAME, "账单明细")
        self.assertEqual(gui_app.V8_9_BILL_SPLIT_DEFAULT_FIELD, "经手人")
        self.assertEqual(gui_app.V8_9_BILL_SPLIT_OUTPUT_DIR_SUFFIX, "拆分结果")
        self.assertEqual(
            tuple(getattr(gui_app, "V8_9_BILL_SPLIT_WORKFLOW_STEPS", ())),
            (("config", "配置"), ("run", "运行"), ("results", "结果")),
        )
        self.assertEqual(
            tuple(getattr(gui_app, "V8_9_BILL_SPLIT_CONFIG_PAGE_SECTIONS", ())),
            ("账单目录", "待拆分文件"),
        )
        self.assertEqual(
            tuple(getattr(gui_app, "V8_9_BILL_SPLIT_RUN_PAGE_SECTIONS", ())),
            ("运行状态", "运行日志"),
        )
        self.assertEqual(
            tuple(getattr(gui_app, "V8_9_BILL_SPLIT_RESULTS_PAGE_SECTIONS", ())),
            ("拆分结果",),
        )
        self.assertEqual(
            list(gui_app.V8_9_BILL_SPLIT_ACTIONS),
            ["选择目录", "同步字段", "开始测试", "开始拆分", "打开输出目录", "打开选中文件"],
        )
        self.assertNotIn("开始计算", gui_app.V8_9_BILL_SPLIT_ACTIONS)
        self.assertNotIn("读取表头", gui_app.V8_9_BILL_SPLIT_ACTIONS)
        self.assertNotIn("立即拆分", gui_app.V8_9_BILL_SPLIT_ACTIONS)

    def test_bill_splitter_tree_columns_match_v8_9_design(self) -> None:
        self.assertEqual(
            list(gui_app.V8_9_BILL_SPLIT_FILE_COLUMN_IDS),
            ["name", "sheet", "rows", "field_status", "path"],
        )
        self.assertEqual(
            list(gui_app.V8_9_BILL_SPLIT_FILE_COLUMNS),
            ["文件名", "账单明细", "数据行", "字段状态", "路径"],
        )
        self.assertEqual(
            list(gui_app.V8_9_BILL_SPLIT_RESULT_COLUMN_IDS),
            ["split_value", "name", "rows", "status", "path"],
        )
        self.assertEqual(
            list(gui_app.V8_9_BILL_SPLIT_RESULT_COLUMNS),
            ["拆分值", "文件名", "行数", "状态", "路径"],
        )

    def test_show_bill_splitter_page_updates_active_nav_title_and_raises_page(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        app.active_nav_var = FakeVar("费用计算")
        app.module_title_var = FakeVar("费用计算")
        app.module_subtitle_var = FakeVar("")
        app.nav_labels = {
            "费用计算": FakeWidget(),
            "账户余额": FakeWidget(),
            "报价预览": FakeWidget(),
            "拆分账单": FakeWidget(),
            "系统设置": FakeWidget(),
        }
        app.fee_page = FakeWidget()
        app.balance_page = FakeWidget()
        app.price_preview_page = FakeWidget()
        app.bill_splitter_page = FakeWidget()
        app.settings_page = FakeWidget()
        app.active_workflow_var = FakeVar("results")
        app.active_bill_split_workflow_var = FakeVar("results")
        app.workflow_step_labels = {"results": FakeWidget()}
        app.workflow_pages = {"results": FakeWidget()}
        app.bill_split_workflow_step_labels = {
            "config": FakeWidget(),
            "run": FakeWidget(),
            "results": FakeWidget(),
        }
        app.bill_split_workflow_pages = {
            "config": FakeWidget(),
            "run": FakeWidget(),
            "results": FakeWidget(),
        }

        app._show_page("拆分账单")

        self.assertEqual(app.active_nav_var.get(), "拆分账单")
        self.assertEqual(app.module_title_var.get(), "拆分账单")
        self.assertIn("账单明细", app.module_subtitle_var.get())
        self.assertEqual(app.nav_labels["拆分账单"].options["style"], "NavActive.TLabel")
        self.assertTrue(app.bill_splitter_page.raised)
        self.assertFalse(app.price_preview_page.raised)
        self.assertFalse(app.settings_page.raised)
        self.assertEqual(app.active_bill_split_workflow_var.get(), "config")
        self.assertTrue(app.bill_split_workflow_pages["config"].raised)
        self.assertFalse(app.bill_split_workflow_pages["run"].raised)
        self.assertFalse(app.bill_split_workflow_pages["results"].raised)
        self.assertEqual(app.bill_split_workflow_step_labels["config"].options["style"], "StepActive.TLabel")
        self.assertEqual(app.bill_split_workflow_step_labels["run"].options["style"], "StepIdle.TLabel")
        self.assertEqual(app.active_workflow_var.get(), "results")
        self.assertFalse(app.workflow_pages["results"].raised)

    def test_show_bill_split_workflow_step_is_independent_from_fee_workflow(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        app.active_workflow_var = FakeVar("config")
        app.workflow_step_labels = {
            "config": FakeWidget(),
            "run": FakeWidget(),
            "results": FakeWidget(),
        }
        app.workflow_pages = {
            "config": FakeWidget(),
            "run": FakeWidget(),
            "results": FakeWidget(),
        }
        app.active_bill_split_workflow_var = FakeVar("config")
        app.bill_split_workflow_step_labels = {
            "config": FakeWidget(),
            "run": FakeWidget(),
            "results": FakeWidget(),
        }
        app.bill_split_workflow_pages = {
            "config": FakeWidget(),
            "run": FakeWidget(),
            "results": FakeWidget(),
        }

        self.assertIn("_show_bill_split_workflow_step", vars(ExpressFeeApp))
        app._show_bill_split_workflow_step("run")

        self.assertEqual(app.active_bill_split_workflow_var.get(), "run")
        self.assertEqual(app.active_workflow_var.get(), "config")
        self.assertTrue(app.bill_split_workflow_pages["run"].raised)
        self.assertFalse(app.workflow_pages["run"].raised)
        self.assertEqual(app.bill_split_workflow_step_labels["run"].options["style"], "StepActive.TLabel")
        self.assertEqual(app.workflow_step_labels["run"].options, {})

        app._show_bill_split_workflow_step("results")

        self.assertEqual(app.active_bill_split_workflow_var.get(), "results")
        self.assertEqual(app.active_workflow_var.get(), "config")
        self.assertTrue(app.bill_split_workflow_pages["results"].raised)
        self.assertFalse(app.workflow_pages["results"].raised)
        self.assertEqual(app.bill_split_workflow_step_labels["results"].options["style"], "StepActive.TLabel")
        self.assertEqual(app.workflow_step_labels["results"].options, {})

    def test_configure_bill_split_field_combo_defaults_to_handler_field(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        app.bill_split_common_headers = ["运单号码", "费用", "经手人", "店铺名称"]
        app.bill_split_field_var = FakeVar()
        app.bill_split_field_combo = FakeCombobox()

        app._configure_bill_split_field_combo()

        self.assertEqual(app.bill_split_field_combo.configured_values, app.bill_split_common_headers)
        self.assertEqual(app.bill_split_field_var.get(), "经手人")

    def test_configure_bill_split_field_combo_preserves_selected_common_field(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        app.bill_split_common_headers = ["运单号码", "省份", "费用", "经手人"]
        app.bill_split_field_var = FakeVar("省份")
        app.bill_split_field_combo = FakeCombobox()

        app._configure_bill_split_field_combo()

        self.assertEqual(app.bill_split_field_combo.configured_values, app.bill_split_common_headers)
        self.assertEqual(app.bill_split_field_var.get(), "省份")

    def test_choose_bill_split_dir_waits_for_start_test_before_scanning(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        app.bill_split_source_dir_var = FakeVar()
        app.bill_split_output_dir_var = FakeVar()
        app.bill_split_status_var = FakeVar()
        app.bill_split_summary_files_var = FakeVar("9")
        app.bill_split_summary_success_var = FakeVar("4")
        app.bill_split_summary_failed_var = FakeVar("1")
        app.bill_split_summary_outputs_var = FakeVar("4")
        app.bill_split_common_headers = ["旧字段"]
        app.bill_split_scan_result = SimpleNamespace(files=[object()], errors=[])
        app.bill_split_file_tree = FakeTree()
        app.bill_split_result_file_tree = FakeTree()
        app.bill_split_result_tree = FakeTree()
        app.bill_split_result_paths = {"old": Path("/tmp/old.xlsx")}
        app.bill_split_field_var = FakeVar("旧字段")
        app.bill_split_field_combo = FakeCombobox()
        app.bill_split_log_text = FakeLogText()
        app.bill_split_buttons = []
        app.bill_split_run_button = FakeButton()
        app.bill_split_open_output_button = FakeButton()
        app.bill_split_open_selected_button = FakeButton()

        scan_calls: list[str] = []
        original_askdirectory = gui_app.filedialog.askdirectory
        original_scan = ExpressFeeApp._scan_bill_split_dir

        def fake_askdirectory(**_kwargs: object) -> str:
            return "/tmp/5.7日拆分记录"

        def fake_scan(self: ExpressFeeApp) -> None:
            scan_calls.append("scan")

        try:
            gui_app.filedialog.askdirectory = fake_askdirectory  # type: ignore[assignment]
            ExpressFeeApp._scan_bill_split_dir = fake_scan  # type: ignore[method-assign]

            app._choose_bill_split_dir()
        finally:
            gui_app.filedialog.askdirectory = original_askdirectory  # type: ignore[assignment]
            ExpressFeeApp._scan_bill_split_dir = original_scan  # type: ignore[method-assign]

        self.assertEqual(scan_calls, [])
        self.assertEqual(app.bill_split_source_dir_var.get(), "/tmp/5.7日拆分记录")
        self.assertEqual(app.bill_split_output_dir_var.get(), "/tmp/5.7日拆分记录拆分结果【拆分字段：经手人】")
        self.assertEqual(app.bill_split_common_headers, [])
        self.assertIsNone(app.bill_split_scan_result)
        self.assertEqual(app.bill_split_file_tree.rows, [])
        self.assertEqual(app.bill_split_result_file_tree.rows, [])
        self.assertEqual(app.bill_split_result_tree.rows, [])
        self.assertEqual(app.bill_split_run_button.options["state"], gui_app.tk.DISABLED)
        self.assertIn("等待开始测试", app.bill_split_status_var.get())
        self.assertIn("已选择账单目录", app.bill_split_log_text.text)

    def test_apply_bill_split_scan_result_updates_file_list_and_status(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        app.bill_split_output_dir_var = FakeVar()
        app.bill_split_source_dir_var = FakeVar("/tmp/5.7日拆分记录")
        app.bill_split_field_var = FakeVar("省份")
        app.bill_split_status_var = FakeVar()
        app.bill_split_summary_files_var = FakeVar()
        app.bill_split_summary_success_var = FakeVar()
        app.bill_split_summary_failed_var = FakeVar()
        app.bill_split_summary_outputs_var = FakeVar()
        app.bill_split_file_tree = FakeTree()
        app.bill_split_common_headers = []
        app.bill_split_field_combo = FakeCombobox()
        app.bill_split_log_text = FakeLogText()
        app.bill_split_scan_result = None
        app.bill_split_buttons = []
        app.bill_split_run_button = None
        app.bill_split_open_output_button = None

        scan = SimpleNamespace(
            output_dir=Path("/tmp/5.7日拆分记录拆分结果【拆分字段：省份】"),
            common_headers=["运单号码", "省份", "费用", "经手人"],
            files=[
                SimpleNamespace(
                    path=Path("/tmp/5.7日拆分记录/陈开5月.xlsx"),
                    sheet_found=True,
                    total_rows=18,
                    has_split_field=True,
                    headers=["运单号码", "省份", "费用", "经手人"],
                    error="",
                )
            ],
            logs=["已忽略非账单明细 sheet。"],
            errors=[],
        )

        app._apply_bill_split_scan_result(scan)

        self.assertEqual(app.bill_split_output_dir_var.get(), "/tmp/5.7日拆分记录拆分结果【拆分字段：省份】")
        self.assertEqual(app.bill_split_summary_files_var.get(), "1")
        self.assertEqual(app.bill_split_field_var.get(), "省份")
        self.assertIn("已通过按【省份】拆分字段的测试", app.bill_split_status_var.get())
        self.assertIn("开始拆分", app.bill_split_status_var.get())
        self.assertEqual(
            app.bill_split_file_tree.rows,
            [("陈开5月.xlsx", "已找到", "18", "可拆分", "/tmp/5.7日拆分记录/陈开5月.xlsx")],
        )

    def test_bill_split_start_test_scan_moves_to_run_step_and_enables_split(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        app.active_bill_split_workflow_var = FakeVar("config")
        app.bill_split_workflow_step_labels = {
            "config": FakeWidget(),
            "run": FakeWidget(),
            "results": FakeWidget(),
        }
        app.bill_split_workflow_pages = {
            "config": FakeWidget(),
            "run": FakeWidget(),
            "results": FakeWidget(),
        }
        app.bill_split_output_dir_var = FakeVar()
        app.bill_split_source_dir_var = FakeVar("/tmp/账单")
        app.bill_split_field_var = FakeVar("经手人")
        app.bill_split_status_var = FakeVar()
        app.bill_split_summary_files_var = FakeVar()
        app.bill_split_summary_success_var = FakeVar()
        app.bill_split_summary_failed_var = FakeVar()
        app.bill_split_summary_outputs_var = FakeVar()
        app.bill_split_file_tree = FakeTree()
        app.bill_split_common_headers = []
        app.bill_split_field_combo = FakeCombobox()
        app.bill_split_log_text = FakeLogText()
        app.bill_split_scan_result = None
        app.bill_split_buttons = []
        app.bill_split_run_button = FakeButton()
        app.bill_split_open_output_button = FakeButton()

        scan = SimpleNamespace(
            output_dir=Path("/tmp/账单拆分结果"),
            common_headers=["运单号码", "经手人", "费用"],
            files=[
                SimpleNamespace(
                    path=Path("/tmp/账单/A.xlsx"),
                    sheet_name="账单明细",
                    headers=["运单号码", "经手人", "费用"],
                    total_rows=3,
                    error="",
                )
            ],
            logs=["预检通过：1 个文件。"],
            errors=[],
        )

        app._apply_bill_split_scan_result(scan)

        self.assertEqual(app.active_bill_split_workflow_var.get(), "run")
        self.assertTrue(app.bill_split_workflow_pages["run"].raised)
        self.assertEqual(app.bill_split_run_button.options["state"], gui_app.tk.NORMAL)

    def test_sync_bill_split_fields_updates_headers_without_leaving_config_step(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_text:
            source_dir = Path(temp_dir_text) / "账单"
            source_dir.mkdir()
            app = ExpressFeeApp.__new__(ExpressFeeApp)
            app.active_bill_split_workflow_var = FakeVar("config")
            app.bill_split_workflow_step_labels = {
                "config": FakeWidget(),
                "run": FakeWidget(),
                "results": FakeWidget(),
            }
            app.bill_split_workflow_pages = {
                "config": FakeWidget(),
                "run": FakeWidget(),
                "results": FakeWidget(),
            }
            app.bill_split_output_dir_var = FakeVar()
            app.bill_split_source_dir_var = FakeVar(str(source_dir))
            app.bill_split_field_var = FakeVar("省份")
            app.bill_split_status_var = FakeVar()
            app.bill_split_summary_files_var = FakeVar()
            app.bill_split_summary_success_var = FakeVar()
            app.bill_split_summary_failed_var = FakeVar()
            app.bill_split_summary_outputs_var = FakeVar()
            app.bill_split_file_tree = FakeTree()
            app.bill_split_common_headers = []
            app.bill_split_field_combo = FakeCombobox()
            app.bill_split_log_text = FakeLogText()
            app.bill_split_scan_result = None
            app.bill_split_result_paths = {}
            app.bill_split_buttons = []
            app.bill_split_run_button = FakeButton()
            app.bill_split_open_output_button = FakeButton()

            scan = SimpleNamespace(
                output_dir=source_dir.parent / "账单拆分结果",
                common_headers=["运单号码", "省份", "经手人", "费用"],
                files=[
                    SimpleNamespace(
                        path=source_dir / "A.xlsx",
                        sheet_name="账单明细",
                        headers=["运单号码", "省份", "经手人", "费用"],
                        total_rows=3,
                        error="",
                    )
                ],
                logs=["同步字段完成：1 个文件。"],
                errors=[],
            )
            calls: list[tuple[Path, str]] = []
            original_scan = gui_app.scan_bill_split_directory

            def fake_scan_bill_split_directory(path: Path, split_field: str) -> object:
                calls.append((path, split_field))
                return scan

            try:
                gui_app.scan_bill_split_directory = fake_scan_bill_split_directory  # type: ignore[assignment]

                app._sync_bill_split_fields()
            finally:
                gui_app.scan_bill_split_directory = original_scan  # type: ignore[assignment]

            self.assertEqual(calls, [(source_dir, "省份")])
            self.assertEqual(app.active_bill_split_workflow_var.get(), "config")
            self.assertTrue(app.bill_split_workflow_pages["config"].raised)
            self.assertFalse(app.bill_split_workflow_pages["run"].raised)
            self.assertEqual(app.bill_split_field_var.get(), "省份")
            self.assertEqual(app.bill_split_field_combo.configured_values, ["运单号码", "省份", "经手人", "费用"])
            self.assertEqual(app.bill_split_run_button.options["state"], gui_app.tk.NORMAL)
            self.assertIn("已同步共同字段：4 个字段", app.bill_split_status_var.get())
            self.assertIn("同步字段完成", app.bill_split_log_text.text)

    def test_apply_bill_split_scan_result_shows_failed_file_without_blocking_valid_files(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        app.bill_split_output_dir_var = FakeVar()
        app.bill_split_source_dir_var = FakeVar("/tmp/账单")
        app.bill_split_field_var = FakeVar()
        app.bill_split_status_var = FakeVar()
        app.bill_split_summary_files_var = FakeVar()
        app.bill_split_summary_success_var = FakeVar()
        app.bill_split_summary_failed_var = FakeVar()
        app.bill_split_summary_outputs_var = FakeVar()
        app.bill_split_file_tree = FakeTree()
        app.bill_split_common_headers = []
        app.bill_split_field_combo = FakeCombobox()
        app.bill_split_log_text = FakeLogText()
        app.bill_split_scan_result = None
        app.bill_split_buttons = []
        app.bill_split_run_button = None
        app.bill_split_open_output_button = None

        scan = SimpleNamespace(
            output_dir=Path("/tmp/账单拆分结果"),
            common_headers=["运单号码", "经手人", "费用"],
            files=[
                SimpleNamespace(
                    path=Path("/tmp/账单/A.xlsx"),
                    sheet_name="账单明细",
                    headers=["运单号码", "经手人", "费用"],
                    total_rows=1,
                    error="",
                ),
                SimpleNamespace(
                    path=Path("/tmp/账单/坏文件.xlsx"),
                    sheet_name="账单明细",
                    headers=[],
                    total_rows=0,
                    error="读取失败：File is not a zip file",
                ),
            ],
            logs=["跳过读取失败的文件：坏文件.xlsx，File is not a zip file"],
            errors=[],
        )

        app._apply_bill_split_scan_result(scan)

        self.assertEqual(
            app.bill_split_file_tree.rows,
            [
                ("A.xlsx", "已找到", "1", "可拆分", "/tmp/账单/A.xlsx"),
                (
                    "坏文件.xlsx",
                    "读取失败",
                    "0",
                    "读取失败：File is not a zip file",
                    "/tmp/账单/坏文件.xlsx",
                ),
            ],
        )
        self.assertEqual(app.bill_split_summary_failed_var.get(), "1")

    def test_bill_split_results_page_shows_split_outputs(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        app.active_bill_split_workflow_var = FakeVar("run")
        app.bill_split_workflow_step_labels = {
            "config": FakeWidget(),
            "run": FakeWidget(),
            "results": FakeWidget(),
        }
        app.bill_split_workflow_pages = {
            "config": FakeWidget(),
            "run": FakeWidget(),
            "results": FakeWidget(),
        }
        app.bill_split_output_dir_var = FakeVar("/tmp/账单拆分结果")
        app.bill_split_source_dir_var = FakeVar("/tmp/账单")
        app.bill_split_field_var = FakeVar("经手人")
        app.bill_split_status_var = FakeVar()
        app.bill_split_summary_files_var = FakeVar("2")
        app.bill_split_summary_success_var = FakeVar("0")
        app.bill_split_summary_failed_var = FakeVar("0")
        app.bill_split_summary_outputs_var = FakeVar("0")
        app.bill_split_file_tree = FakeTree()
        app.bill_split_file_tree.insert(
            "",
            gui_app.tk.END,
            values=("A.xlsx", "已找到", "3", "可拆分", "/tmp/账单/A.xlsx"),
        )
        app.bill_split_file_tree.insert(
            "",
            gui_app.tk.END,
            values=("B.xlsx", "已找到", "4", "可拆分", "/tmp/账单/B.xlsx"),
        )
        app.bill_split_result_tree = FakeTree()
        app.bill_split_result_paths = {}
        app.bill_split_field_combo = FakeCombobox()
        app.bill_split_log_text = FakeLogText()
        app.bill_split_buttons = []
        app.bill_split_run_button = FakeButton()
        app.bill_split_open_output_button = FakeButton()
        app.bill_split_open_selected_button = FakeButton()
        app.bill_split_common_headers = ["运单号码", "经手人", "费用"]
        app.bill_split_scan_result = SimpleNamespace(
            files=[FakeFileScan(), FakeFileScan(error="读取失败：File is not a zip file")],
            errors=[],
        )

        result = SimpleNamespace(
            output_dir=Path("/tmp/账单拆分结果"),
            common_headers=["运单号码", "经手人", "费用"],
            output_paths=[
                Path("/tmp/账单拆分结果/经手人_张三.xlsx"),
                Path("/tmp/账单拆分结果/经手人_李四.xlsx"),
            ],
            outputs=[
                SimpleNamespace(
                    split_value="张三/A",
                    output_path=Path("/tmp/账单拆分结果/经手人_张三_A.xlsx"),
                    row_count=6,
                ),
                SimpleNamespace(
                    split_value="李四",
                    output_path=Path("/tmp/账单拆分结果/经手人_李四.xlsx"),
                    row_count=1,
                ),
            ],
        )

        app._apply_bill_split_result(result)

        self.assertEqual(app.active_bill_split_workflow_var.get(), "results")
        self.assertTrue(app.bill_split_workflow_pages["results"].raised)
        self.assertEqual(app.bill_split_summary_files_var.get(), "2")
        self.assertEqual(app.bill_split_summary_success_var.get(), "1")
        self.assertEqual(app.bill_split_summary_failed_var.get(), "1")
        self.assertEqual(app.bill_split_summary_outputs_var.get(), "2")
        self.assertEqual(
            app.bill_split_result_tree.rows,
            [
                ("张三/A", "经手人_张三_A.xlsx", "6", "成功", "/tmp/账单拆分结果/经手人_张三_A.xlsx"),
                ("李四", "经手人_李四.xlsx", "1", "成功", "/tmp/账单拆分结果/经手人_李四.xlsx"),
            ],
        )

    def test_bill_split_log_is_independent_from_fee_calculation_log(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        app.log_text = FakeLogText()
        app.bill_split_log_text = FakeLogText()

        app._append_bill_split_log("开始拆分")

        self.assertEqual(app.log_text.text, "")
        self.assertIn("开始拆分", app.bill_split_log_text.text)
        self.assertTrue(app.bill_split_log_text.seen_end)


if __name__ == "__main__":
    unittest.main()
