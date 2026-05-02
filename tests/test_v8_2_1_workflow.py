from __future__ import annotations

import sys
import unittest
from pathlib import Path

APP_DIR = Path(__file__).resolve().parents[1] / "app"
sys.path.insert(0, str(APP_DIR))

from express_app.gui.app import (  # noqa: E402
    V8_2_1_AUTO_WORKFLOW_TRANSITIONS,
    V8_2_1_CONFIG_PAGE_SECTIONS,
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


if __name__ == "__main__":
    unittest.main()
