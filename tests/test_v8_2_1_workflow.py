from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

APP_DIR = Path(__file__).resolve().parents[1] / "app"
sys.path.insert(0, str(APP_DIR))

from express_app.gui.app import (  # noqa: E402
    V8_2_1_AUTO_WORKFLOW_TRANSITIONS,
    V8_2_1_CONFIG_PAGE_SECTIONS,
    V8_2_1_CONFIG_GENERATION_OPTIONS,
    V8_2_1_CONFIG_SYSTEM_DIRECTORY_LABELS,
    V8_2_1_WORKFLOW_STEPS,
    ExpressFeeApp,
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


if __name__ == "__main__":
    unittest.main()
