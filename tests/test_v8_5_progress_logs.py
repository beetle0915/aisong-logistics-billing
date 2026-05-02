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


class V851ProgressLogsTest(unittest.TestCase):
    def _write_sales_file(
        self,
        path: Path,
        row_count: int = 3,
        customer: str = "客户A",
    ) -> None:
        workbook = openpyxl.Workbook()
        ws = workbook.active
        ws.append(["出库日期", "业务员", "快递公司", "省", "重量"])
        for index in range(row_count):
            ws.append([date(2026, 4, 7), customer, "顺丰", "广东", 1 + index])
        workbook.save(path)

    def _write_price_file(self, price_dir: Path, customer: str = "客户A") -> None:
        workbook = openpyxl.Workbook()
        ws = workbook.active
        ws.title = "顺丰"
        ws.append(["省份参照列", "首重费用", "续重费用"])
        ws.append(["广东", 10, 2])
        workbook.save(price_dir / f"{customer}-快递报价.xlsx")

    def test_batch_job_emits_customer_readable_stage_summaries(self) -> None:
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
            self.assertIn("阶段 1/5：准备数据", progress_text)
            self.assertIn("里程碑：报价读取完成", progress_text)
            self.assertIn("阶段 2/5：计算快递费", progress_text)
            self.assertIn("进度：快递费计算已处理 3/3 行，成功 3 行，异常 0 行", progress_text)
            self.assertIn("阶段 3/5：生成客户每日明细", progress_text)
            self.assertIn("里程碑：客户每日明细生成完成", progress_text)
            self.assertIn("阶段 4/5：刷新客户历史汇总", progress_text)
            self.assertIn("进度：客户历史汇总已刷新", progress_text)
            self.assertIn("阶段 5/5：完成", progress_text)
            self.assertNotIn("->", progress_text)
            self.assertNotIn(str(temp_dir), progress_text)
            self.assertNotIn("报价目录：", progress_text)
            stage_numbers = [
                int(message.split("阶段 ", 1)[1].split("/5", 1)[0])
                for message in progress_messages
                if message.startswith("阶段 ")
            ]
            self.assertEqual(stage_numbers, sorted(stage_numbers))

    def test_batch_job_finishes_all_sales_files_before_customer_detail_stage(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_text:
            temp_dir = Path(temp_dir_text)
            first_sales_file = temp_dir / "销售出库单_1.xlsx"
            second_sales_file = temp_dir / "销售出库单_2.xlsx"
            price_dir = temp_dir / "prices"
            price_dir.mkdir()
            self._write_sales_file(first_sales_file, row_count=2, customer="客户A")
            self._write_sales_file(second_sales_file, row_count=2, customer="客户A")
            self._write_price_file(price_dir)
            progress_messages: list[str] = []

            run_express_fee_batch_job(
                ExpressFeeBatchJobConfig(
                    sales_files=[first_sales_file, second_sales_file],
                    price_dir=price_dir,
                    output_dir=temp_dir / "output",
                    split_dir=temp_dir / "split",
                    split_customer_daily_files=True,
                    generate_customer_history=True,
                    rule_config=build_default_rule_config(),
                ),
                progress_callback=progress_messages.append,
            )

            first_detail_stage_index = progress_messages.index("阶段 3/5：生成客户每日明细")
            second_sales_done_index = next(
                index
                for index, message in enumerate(progress_messages)
                if message.startswith("里程碑：销售表 2/2 计算完成")
            )

            self.assertLess(second_sales_done_index, first_detail_stage_index)

    def test_poll_queue_appends_log_events_without_waiting_for_done(self) -> None:
        app = ExpressFeeApp.__new__(ExpressFeeApp)
        app._queue = __import__("queue").Queue()
        app._worker = None
        app.log_text = FakeLogText()

        app._queue.put(("log", "进度：快递费计算已处理 1/3 行，成功 1 行，异常 0 行"))

        app._poll_queue()

        self.assertIn("进度：快递费计算已处理 1/3 行", app.log_text.text)
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

    def test_done_queue_event_removes_local_paths_from_error_summary(self) -> None:
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

        temp_root = Path("/tmp/v851-private-path")
        result = ExpressFeeBatchJobResult(
            sales_files=[temp_root / "sales.xlsx"],
            price_dir=temp_root / "prices",
            output_dir=temp_root / "output",
            split_dir=temp_root / "split",
            job_results=[
                ExpressFeeJobResult(
                    sales_file=temp_root / "sales.xlsx",
                    price_dir=temp_root / "prices",
                    output_path=temp_root / "output" / "result.xlsx",
                    split_dir=temp_root / "split",
                    processing_errors=[
                        "[报价目录错误]\n"
                        "阶段：读取报价表\n"
                        f"位置：{temp_root / 'prices'}\n"
                        f"原因：报价目录中没有可读取的 .xlsx 文件：{temp_root / 'prices'}\n"
                        "建议：请重新选择报价目录。"
                    ],
                    split_errors=[f"客户目录不存在：{temp_root / 'split' / '客户A'}"],
                )
            ],
            history_errors=[f"客户A：客户拆分目录不存在：{temp_root / 'split'}"],
        )
        app._queue.put(("done", result))

        app._poll_queue()

        self.assertIn("需要处理的问题：", app.log_text.text)
        self.assertIn("[报价目录错误]", app.log_text.text)
        self.assertNotIn(str(temp_root), app.log_text.text)
        self.assertNotIn("/tmp/", app.log_text.text)


if __name__ == "__main__":
    unittest.main()
