from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path
from tkinter import messagebox

APP_DIR = Path(__file__).resolve().parents[1] / "app"
sys.path.insert(0, str(APP_DIR))

from express_app.gui.app import (  # noqa: E402
    SIDEBAR_BOTTOM_ACTIONS,
    WORKBENCH_LABEL_FONT_SIZE,
    V8_2_1_AUTO_WORKFLOW_TRANSITIONS,
    V8_2_1_CONFIG_PAGE_SECTIONS,
    V8_2_1_CONFIG_GENERATION_OPTIONS,
    V8_6_CONFIG_PAGE_ACTIONS,
    V8_6_RUN_PAGE_ACTIONS,
    V8_2_1_CONFIG_SYSTEM_DIRECTORY_LABELS,
    V8_2_1_WORKFLOW_STEPS,
    ExpressFeeApp,
)
from express_app.core.models import (  # noqa: E402
    ExpressFeeBatchJobConfig,
    ExpressFeePreflightFileResult,
    ExpressFeePreflightResult,
)


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


class FakeBoolVar(FakeVar):
    def __init__(self, value: bool = False) -> None:
        self.value = value

    def get(self) -> bool:
        return bool(self.value)

    def set(self, value: bool) -> None:
        self.value = value


def write_file(path: Path, text: str = "content") -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")
    return path


class V821WorkflowTest(unittest.TestCase):
    def test_workflow_steps_match_head_design(self) -> None:
        self.assertEqual(
            list(V8_2_1_WORKFLOW_STEPS),
            [
                ("config", "配置"),
                ("run", "运行"),
                ("results", "结果"),
            ],
        )
        self.assertEqual(
            V8_2_1_AUTO_WORKFLOW_TRANSITIONS,
            {"on_start": "run", "on_done": "results"},
        )

    def test_config_page_keeps_business_inputs_and_moves_directories_to_summary(self) -> None:
        self.assertEqual(
            list(V8_2_1_CONFIG_PAGE_SECTIONS),
            ["销售出库单", "当前系统设置", "生成选项"],
        )
        self.assertEqual(
            list(V8_2_1_CONFIG_SYSTEM_DIRECTORY_LABELS),
            ["报价表目录", "总结果目录", "客户明细目录"],
        )
        self.assertEqual(
            list(V8_2_1_CONFIG_GENERATION_OPTIONS),
            ["生成客户每日明细", "生成客户历史汇总"],
        )
        self.assertNotIn("刷新全部客户历史汇总", V8_2_1_CONFIG_GENERATION_OPTIONS)

    def test_sidebar_no_longer_has_global_run_action(self) -> None:
        self.assertNotIn("开始计算", V8_2_1_CONFIG_PAGE_SECTIONS)
        self.assertEqual(list(SIDEBAR_BOTTOM_ACTIONS), ["系统设置", "打开客户目录"])
        self.assertNotIn("打开输出目录", SIDEBAR_BOTTOM_ACTIONS)
        self.assertGreaterEqual(WORKBENCH_LABEL_FONT_SIZE, 12)

    def test_fee_config_page_no_longer_has_settings_shortcut_actions(self) -> None:
        self.assertEqual(list(V8_6_CONFIG_PAGE_ACTIONS), ["开始测试"])
        self.assertEqual(list(V8_6_RUN_PAGE_ACTIONS), ["开始计算"])
        self.assertNotIn("去系统设置", V8_6_CONFIG_PAGE_ACTIONS)
        self.assertNotIn("去系统设置修改", V8_6_CONFIG_PAGE_ACTIONS)

    def test_show_workflow_step_updates_active_state_and_raises_page(self) -> None:
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

        app._show_workflow_step("run")

        self.assertEqual(app.active_workflow_var.get(), "run")
        self.assertEqual(app.workflow_step_labels["run"].options["style"], "StepActive.TLabel")
        self.assertEqual(
            app.workflow_step_labels["config"].options["style"],
            "StepIdle.TLabel",
        )
        self.assertTrue(app.workflow_pages["run"].raised)
        self.assertFalse(app.workflow_pages["config"].raised)

    def test_set_run_buttons_state_updates_all_run_buttons(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        first = FakeWidget()
        second = FakeWidget()
        app.run_buttons = [first, second]

        app._set_run_buttons_state("disabled")

        self.assertEqual(first.options["state"], "disabled")
        self.assertEqual(second.options["state"], "disabled")

    def test_preflight_success_unlocks_run_step_without_starting_calculation(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        config = ExpressFeeBatchJobConfig(
            sales_files=[Path("/tmp/sales.xlsx")],
            price_dir=Path("/tmp/prices"),
            output_dir=Path("/tmp/output"),
            split_dir=Path("/tmp/split"),
        )
        result = ExpressFeePreflightResult(sales_files=[Path("/tmp/sales.xlsx")], price_dir=Path("/tmp/prices"))
        result.file_results = [
            ExpressFeePreflightFileResult(
                sales_file=Path("/tmp/sales.xlsx"),
                total_rows=1,
                success_rows=1,
            )
        ]
        result.logs = ["运行前测试通过：1 个销售表，共 0 行可计算。"]
        app._build_config = lambda require_output_access=False: config  # type: ignore[method-assign]
        app._save_current_config = lambda: None  # type: ignore[method-assign]
        app._clear_log = lambda: setattr(app, "log_cleared", True)  # type: ignore[method-assign]
        app._append_log = lambda message: app.logged.append(message)  # type: ignore[method-assign]
        app._show_page = lambda page: setattr(app, "shown_page", page)  # type: ignore[method-assign]
        app._show_workflow_step = lambda step: setattr(app, "shown_step", step)  # type: ignore[method-assign]
        app._set_preflight_buttons_state = lambda state: setattr(app, "preflight_button_state", state)  # type: ignore[method-assign]
        app._set_run_buttons_state = lambda state: setattr(app, "run_button_state", state)  # type: ignore[method-assign]
        app.status_var = FakeVar("就绪")
        app._worker = None
        app._last_preflight_result = None
        app._last_preflight_config_signature = None
        app.logged = []

        app._apply_preflight_result(config, result)

        self.assertEqual(app.status_var.get(), "测试通过，等待开始计算")
        self.assertEqual(app.shown_step, "run")
        self.assertEqual(app.run_button_state, "normal")
        self.assertIs(app._last_preflight_result, result)
        self.assertIsNone(app._worker)

    def test_preflight_failure_keeps_user_on_run_step_and_locks_calculation(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        config = object()
        result = ExpressFeePreflightResult(
            sales_files=[Path("/tmp/sales.xlsx")],
            price_dir=Path("/tmp/prices"),
            errors=["重量原值：abc"],
            logs=["运行前测试未通过：1 个问题需要处理。", "重量原值：abc"],
        )
        app._append_log = lambda message: app.logged.append(message)  # type: ignore[method-assign]
        app._show_workflow_step = lambda step: setattr(app, "shown_step", step)  # type: ignore[method-assign]
        app._set_run_buttons_state = lambda state: setattr(app, "run_button_state", state)  # type: ignore[method-assign]
        app.status_var = FakeVar("就绪")
        app._last_preflight_result = None
        app._last_preflight_config_signature = None
        app.logged = []

        app._apply_preflight_result(config, result)

        self.assertEqual(app.status_var.get(), "测试未通过")
        self.assertEqual(app.shown_step, "run")
        self.assertEqual(app.run_button_state, "disabled")
        self.assertIsNone(app._last_preflight_result)
        self.assertIn("重量原值：abc", "\n".join(app.logged))

    def test_start_preflight_shows_run_step_so_user_can_watch_test_logs(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        config = ExpressFeeBatchJobConfig(
            sales_files=[Path("/tmp/sales.xlsx")],
            price_dir=Path("/tmp/prices"),
            output_dir=Path("/tmp/output"),
            split_dir=Path("/tmp/split"),
        )
        app._worker = None
        app._build_config = lambda require_output_access=False: config  # type: ignore[method-assign]
        app._clear_log = lambda: None  # type: ignore[method-assign]
        app._clear_results = lambda: None  # type: ignore[method-assign]
        app._reset_summary = lambda: None  # type: ignore[method-assign]
        app._append_log = lambda message: app.logged.append(message)  # type: ignore[method-assign]
        app._show_page = lambda page: setattr(app, "shown_page", page)  # type: ignore[method-assign]
        app._show_workflow_step = lambda step: setattr(app, "shown_step", step)  # type: ignore[method-assign]
        app._set_preflight_buttons_state = lambda state: setattr(app, "preflight_state", state)  # type: ignore[method-assign]
        app._set_run_buttons_state = lambda state: setattr(app, "run_state", state)  # type: ignore[method-assign]
        app.after = lambda *_args, **_kwargs: None  # type: ignore[method-assign]
        app.status_var = FakeVar("就绪")
        app.logged = []
        started: list[bool] = []

        class FakeThread:
            def __init__(self, *_args, **_kwargs) -> None:
                pass

            def start(self) -> None:
                started.append(True)

        import express_app.gui.app as gui_app

        original_thread = gui_app.threading.Thread
        try:
            gui_app.threading.Thread = FakeThread  # type: ignore[assignment]
            app._start_preflight()
        finally:
            gui_app.threading.Thread = original_thread  # type: ignore[assignment]

        self.assertEqual(app.shown_step, "run")
        self.assertEqual(app.status_var.get(), "测试中")
        self.assertEqual(app.run_state, "disabled")
        self.assertTrue(started)

    def test_start_preflight_config_error_still_shows_run_step_with_log(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        app._worker = None
        app._build_config = lambda require_output_access=False: None  # type: ignore[method-assign]
        app._clear_log = lambda: None  # type: ignore[method-assign]
        app._clear_results = lambda: None  # type: ignore[method-assign]
        app._reset_summary = lambda: None  # type: ignore[method-assign]
        app._append_log = lambda message: app.logged.append(message)  # type: ignore[method-assign]
        app._show_page = lambda page: setattr(app, "shown_page", page)  # type: ignore[method-assign]
        app._show_workflow_step = lambda step: setattr(app, "shown_step", step)  # type: ignore[method-assign]
        app._set_preflight_buttons_state = lambda state: setattr(app, "preflight_state", state)  # type: ignore[method-assign]
        app._set_run_buttons_state = lambda state: setattr(app, "run_state", state)  # type: ignore[method-assign]
        app.status_var = FakeVar("就绪")
        app.logged = []

        app._start_preflight()

        self.assertEqual(app.shown_step, "run")
        self.assertEqual(app.status_var.get(), "测试未通过")
        self.assertEqual(app.run_state, "disabled")
        self.assertIn("请先补全配置", "\n".join(app.logged))

    def test_preflight_signature_changes_when_sales_file_content_changes(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_text:
            temp_dir = Path(temp_dir_text)
            sales_file = write_file(temp_dir / "sales.xlsx", "before")
            config = ExpressFeeBatchJobConfig(
                sales_files=[sales_file],
                price_dir=temp_dir / "prices",
                output_dir=temp_dir / "output",
                split_dir=temp_dir / "split",
            )
            app = ExpressFeeApp.__new__(ExpressFeeApp)

            before = app._config_signature(config)
            sales_file.write_text("after", encoding="utf-8")
            after = app._config_signature(config)

        self.assertNotEqual(before, after)

    def test_preflight_signature_changes_when_price_file_content_changes(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_text:
            temp_dir = Path(temp_dir_text)
            price_dir = temp_dir / "prices"
            price_file = write_file(price_dir / "客户A-快递报价.xlsx", "before")
            config = ExpressFeeBatchJobConfig(
                sales_files=[temp_dir / "sales.xlsx"],
                price_dir=price_dir,
                output_dir=temp_dir / "output",
                split_dir=temp_dir / "split",
            )
            app = ExpressFeeApp.__new__(ExpressFeeApp)

            before = app._config_signature(config)
            price_file.write_text("after price table", encoding="utf-8")
            after = app._config_signature(config)

        self.assertNotEqual(before, after)

    def test_preflight_signature_changes_when_price_file_is_added(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_text:
            temp_dir = Path(temp_dir_text)
            price_dir = temp_dir / "prices"
            price_dir.mkdir()
            write_file(price_dir / "客户A-快递报价.xlsx", "before")
            config = ExpressFeeBatchJobConfig(
                sales_files=[temp_dir / "sales.xlsx"],
                price_dir=price_dir,
                output_dir=temp_dir / "output",
                split_dir=temp_dir / "split",
            )
            app = ExpressFeeApp.__new__(ExpressFeeApp)

            before = app._config_signature(config)
            write_file(price_dir / "客户B-快递报价.xlsx", "new price table")
            after = app._config_signature(config)

        self.assertNotEqual(before, after)

    def test_price_directory_access_accepts_salesman_version_folders(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_text:
            temp_dir = Path(temp_dir_text)
            price_dir = temp_dir / "prices"
            write_file(price_dir / "张三" / "20260503张三-快递报价.xlsx")
            app = ExpressFeeApp.__new__(ExpressFeeApp)

            price_files = app._find_price_workbooks(price_dir)

        self.assertEqual([path.name for path in price_files], ["20260503张三-快递报价.xlsx"])

    def test_preflight_signature_changes_when_versioned_price_file_changes(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_text:
            temp_dir = Path(temp_dir_text)
            price_dir = temp_dir / "prices"
            price_file = write_file(price_dir / "张三" / "20260503张三-快递报价.xlsx", "before")
            config = ExpressFeeBatchJobConfig(
                sales_files=[temp_dir / "sales.xlsx"],
                price_dir=price_dir,
                output_dir=temp_dir / "output",
                split_dir=temp_dir / "split",
            )
            app = ExpressFeeApp.__new__(ExpressFeeApp)

            before = app._config_signature(config)
            price_file.write_text("after price table", encoding="utf-8")
            after = app._config_signature(config)

        self.assertNotEqual(before, after)

    def test_hidden_refresh_all_option_is_forced_off_in_job_config(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_text:
            temp_dir = Path(temp_dir_text)
            sales_file = temp_dir / "sales.xlsx"
            sales_file.touch()
            price_dir = temp_dir / "prices"
            output_dir = temp_dir / "output"
            split_dir = temp_dir / "split"
            price_dir.mkdir()
            output_dir.mkdir()
            split_dir.mkdir()

            app = ExpressFeeApp.__new__(ExpressFeeApp)
            app.sales_files = [sales_file]
            app.price_dir_var = FakeVar(str(price_dir))
            app.output_dir_var = FakeVar(str(output_dir))
            app.split_dir_var = FakeVar(str(split_dir))
            app.split_var = FakeBoolVar(True)
            app.history_var = FakeBoolVar(True)
            app.refresh_all_var = FakeBoolVar(True)
            app.rule_config = None
            app._ensure_price_dir_access = lambda path: path  # type: ignore[method-assign]
            app._ensure_writable_dir_access = lambda path, *_args, **_kwargs: path  # type: ignore[method-assign]

            config = app._build_config()

        self.assertFalse(config.refresh_all_customers)

    def test_history_generation_requires_customer_detail_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_text:
            temp_dir = Path(temp_dir_text)
            sales_file = temp_dir / "sales.xlsx"
            sales_file.touch()
            price_dir = temp_dir / "prices"
            price_dir.mkdir()
            (price_dir / "客户A-快递报价.xlsx").touch()

            app = ExpressFeeApp.__new__(ExpressFeeApp)
            app.sales_files = [sales_file]
            app.price_dir_var = FakeVar(str(price_dir))
            app.output_dir_var = FakeVar(str(temp_dir / "output"))
            app.split_dir_var = FakeVar("")
            app.split_var = FakeBoolVar(False)
            app.history_var = FakeBoolVar(True)
            app.rule_config = None
            app._ensure_price_dir_access = lambda path: path  # type: ignore[method-assign]
            app._ensure_writable_dir_access = lambda path, *_args, **_kwargs: path  # type: ignore[method-assign]
            app.errors: list[tuple[str, str]] = []

            original_showerror = messagebox.showerror
            try:
                messagebox.showerror = lambda title, message: app.errors.append((title, message))  # type: ignore[method-assign]
                config = app._build_config(require_output_access=False)
            finally:
                messagebox.showerror = original_showerror  # type: ignore[method-assign]

        self.assertIsNone(config)
        self.assertIn("请选择客户每日明细目录", app.errors[0][1])


if __name__ == "__main__":
    unittest.main()
