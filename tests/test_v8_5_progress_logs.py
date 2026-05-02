from __future__ import annotations

import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

import openpyxl

APP_DIR = Path(__file__).resolve().parents[1] / "app"
sys.path.insert(0, str(APP_DIR))

from express_app.core.calculator import build_default_rule_config, run_express_fee_batch_job  # noqa: E402
from express_app.core.models import ExpressFeeBatchJobConfig  # noqa: E402
from express_app.gui import app as gui_app  # noqa: E402
from express_app.gui.app import ExpressFeeApp  # noqa: E402


class FakeLogText:
    def __init__(self) -> None:
        self.text = ""
        self.options: dict[str, object] = {}
        self.seen_end = False

    def configure(self, **kwargs: object) -> None:
        self.options.update(kwargs)

    def index(self, value: str) -> str:
        return "1.0" if not self.text else "2.0"

    def insert(self, index: str, text: str) -> None:
        self.text += text

    def see(self, index: str) -> None:
        self.seen_end = True


class V85ProgressLogsTest(unittest.TestCase):
    def _write_sales_file(self, path: Path, row_count: int = 3) -> None:
        workbook = openpyxl.Workbook()
        ws = workbook.active
        ws.append(["出库日期", "业务员", "快递公司", "省", "重量"])
        for index in range(row_count):
            ws.append([date(2026, 4, 7), "客户A", "顺丰", "广东", 1 + index])
        workbook.save(path)

    def _write_price_file(self, price_dir: Path) -> None:
        workbook = openpyxl.Workbook()
        ws = workbook.active
        ws.title = "顺丰"
        ws.append(["省份参照列", "首重费用", "续重费用"])
        ws.append(["广东", 10, 2])
        workbook.save(price_dir / "客户A-快递报价.xlsx")

    def test_batch_job_emits_stage_progress_messages(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_text:
            temp_dir = Path(temp_dir_text)
            sales_file = temp_dir / "销售出库单.xlsx"
            price_dir = temp_dir / "prices"
            price_dir.mkdir()
            self._write_sales_file(sales_file, row_count=3)
            self._write_price_file(price_dir)
            progress_messages: list[str] = []

            run_express_fee_batch_job(
                ExpressFeeBatchJobConfig(
                    sales_files=[sales_file],
                    price_dir=price_dir,
                    output_dir=temp_dir / "output",
                    split_dir=temp_dir / "split",
                    split_customer_daily_files=True,
                    generate_customer_history=True,
                    rule_config=build_default_rule_config(),
                ),
                progress_callback=progress_messages.append,
            )

            progress_text = "\n".join(progress_messages)
            self.assertIn("准备运行：共 1 个销售表", progress_text)
            self.assertIn("正在读取报价表", progress_text)
            self.assertIn("销售表 1/1", progress_text)
            self.assertIn("计算快递费：已处理 3/3 行", progress_text)
            self.assertIn("客户每日明细：已完成", progress_text)
            self.assertIn("客户历史汇总：已完成", progress_text)
            self.assertNotIn("->", progress_text)

    def test_poll_queue_appends_log_events_without_waiting_for_done(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        app._queue = __import__("queue").Queue()
        app._worker = None
        app.log_text = FakeLogText()

        app._queue.put(("log", "计算快递费：已处理 1/3 行"))

        app._poll_queue()

        self.assertIn("计算快递费：已处理 1/3 行", app.log_text.text)
        self.assertTrue(app.log_text.seen_end)

    def test_done_queue_event_appends_compact_summary_without_file_paths(self) -> None:
        from express_app.core.models import ExpressFeeBatchJobResult, ExpressFeeJobResult

        original_showwarning = gui_app.messagebox.showwarning
        original_showinfo = gui_app.messagebox.showinfo
        gui_app.messagebox.showwarning = lambda *args, **kwargs: None
        gui_app.messagebox.showinfo = lambda *args, **kwargs: None
        self.addCleanup(lambda: setattr(gui_app.messagebox, "showwarning", original_showwarning))
        self.addCleanup(lambda: setattr(gui_app.messagebox, "showinfo", original_showinfo))

        app = ExpressFeeApp.__new__(ExpressFeeApp)
        app._queue = __import__("queue").Queue()
        app._worker = None
        app.log_text = FakeLogText()
        app.status_var = type("Var", (), {"set": lambda self, value: None})()
        app._last_result = None
        app._populate_result_table = lambda result: None
        app._update_summary = lambda result: None
        app._set_run_buttons_state = lambda state: None
        app._show_workflow_step = lambda step: None

        result = ExpressFeeBatchJobResult(
            sales_files=[Path("/tmp/销售出库单.xlsx")],
            price_dir=Path("/tmp/prices"),
            output_dir=Path("/tmp/output"),
            split_dir=Path("/tmp/split"),
            job_results=[
                ExpressFeeJobResult(
                    sales_file=Path("/tmp/销售出库单.xlsx"),
                    price_dir=Path("/tmp/prices"),
                    output_path=Path("/tmp/output/result.xlsx"),
                    split_dir=Path("/tmp/split"),
                    total_rows=3,
                    success_rows=2,
                    failed_rows=1,
                    processing_errors=["第 3 行：重量不是数字：abc"],
                )
            ],
            logs=["冗长日志 -> /tmp/output/result.xlsx"],
        )
        app._queue.put(("done", result))

        app._poll_queue()

        self.assertIn("批量任务完成", app.log_text.text)
        self.assertIn("总成功行数：2 条", app.log_text.text)
        self.assertIn("第 3 行：重量不是数字：abc", app.log_text.text)
        self.assertNotIn("/tmp/output/result.xlsx", app.log_text.text)


if __name__ == "__main__":
    unittest.main()
