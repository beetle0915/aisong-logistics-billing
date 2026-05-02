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


class ErrorLogTest(unittest.TestCase):
    def _write_sales_file(
        self,
        path: Path,
        headers: list[str] | None = None,
        row: list[object] | None = None,
    ) -> None:
        workbook = openpyxl.Workbook()
        ws = workbook.active
        ws.append(headers or ["出库日期", "业务员", "快递公司", "省", "重量"])
        ws.append(row or [date(2026, 4, 7), "客户A", "顺丰", "广东", 2.5])
        workbook.save(path)

    def _write_price_file(
        self,
        price_dir: Path,
        salesman: str = "客户A",
        sheet_name: str = "顺丰",
        province: str = "广东",
    ) -> Path:
        workbook = openpyxl.Workbook()
        ws = workbook.active
        ws.title = sheet_name
        ws.append(["省份参照列", "首重费用", "续重费用"])
        ws.append([province, 10, 2])
        output_path = price_dir / f"{salesman}-快递报价.xlsx"
        workbook.save(output_path)
        return output_path

    def _run_batch(self, temp_dir: Path, sales_file: Path, price_dir: Path):
        return run_express_fee_batch_job(
            ExpressFeeBatchJobConfig(
                sales_files=[sales_file],
                price_dir=price_dir,
                output_dir=temp_dir / "output",
                split_dir=temp_dir / "split",
                split_customer_daily_files=False,
                generate_customer_history=False,
                rule_config=build_default_rule_config(),
            )
        )

    def test_empty_price_directory_log_suggests_reselecting_price_dir(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_text:
            temp_dir = Path(temp_dir_text)
            sales_file = temp_dir / "销售出库单.xlsx"
            price_dir = temp_dir / "empty-prices"
            price_dir.mkdir()
            self._write_sales_file(sales_file)

            result = self._run_batch(temp_dir, sales_file, price_dir)

            logs = "\n".join(result.logs)
            self.assertFalse(result.ok)
            self.assertIn("[报价表错误]", logs)
            self.assertIn("目录中没有找到可用的 .xlsx 报价文件", logs)
            self.assertIn("建议：请重新选择报价表目录", logs)

    def test_missing_sales_columns_log_shows_missing_and_detected_headers(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_text:
            temp_dir = Path(temp_dir_text)
            sales_file = temp_dir / "销售出库单.xlsx"
            price_dir = temp_dir / "prices"
            price_dir.mkdir()
            self._write_price_file(price_dir)
            self._write_sales_file(
                sales_file,
                headers=["出库日期", "业务员", "快递公司", "省"],
                row=[date(2026, 4, 7), "客户A", "顺丰", "广东"],
            )

            result = self._run_batch(temp_dir, sales_file, price_dir)

            logs = "\n".join(result.logs)
            self.assertFalse(result.ok)
            self.assertIn("[销售表结构错误]", logs)
            self.assertIn("缺少必要列：重量", logs)
            self.assertIn("当前识别到的列：出库日期、业务员、快递公司、省", logs)
            self.assertIn("建议：请恢复销售出库单表头", logs)

    def test_invalid_weight_log_shows_row_value_and_suggestion(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_text:
            temp_dir = Path(temp_dir_text)
            sales_file = temp_dir / "销售出库单.xlsx"
            price_dir = temp_dir / "prices"
            price_dir.mkdir()
            self._write_price_file(price_dir)
            self._write_sales_file(
                sales_file,
                row=[date(2026, 4, 7), "客户A", "顺丰", "广东", "abc"],
            )

            result = self._run_batch(temp_dir, sales_file, price_dir)

            logs = "\n".join(result.logs)
            self.assertFalse(result.ok)
            self.assertIn("[行级计算错误]", logs)
            self.assertIn("行号：2", logs)
            self.assertIn("重量原值：abc", logs)
            self.assertIn("原因：重量不是数字", logs)
            self.assertIn("建议：请将重量改为数字", logs)

    def test_missing_price_log_shows_context_and_suggested_sheet(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_text:
            temp_dir = Path(temp_dir_text)
            sales_file = temp_dir / "销售出库单.xlsx"
            price_dir = temp_dir / "prices"
            price_dir.mkdir()
            self._write_price_file(price_dir, province="浙江")
            self._write_sales_file(
                sales_file,
                row=[date(2026, 4, 7), "客户A", "顺丰", "广东", 2.5],
            )

            result = self._run_batch(temp_dir, sales_file, price_dir)

            logs = "\n".join(result.logs)
            self.assertFalse(result.ok)
            self.assertIn("[报价表错误]", logs)
            self.assertIn("业务员：客户A", logs)
            self.assertIn("计费模板：顺丰", logs)
            self.assertIn("省：广东", logs)
            self.assertIn("重量：2.5", logs)
            self.assertIn("建议：请检查 客户A 的报价表中是否存在 sheet「顺丰」", logs)

    def test_missing_large_piece_template_log_names_required_sheet(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_text:
            temp_dir = Path(temp_dir_text)
            sales_file = temp_dir / "销售出库单.xlsx"
            price_dir = temp_dir / "prices"
            price_dir.mkdir()
            self._write_price_file(price_dir)
            self._write_sales_file(
                sales_file,
                row=[date(2026, 4, 7), "客户A", "顺丰", "广东", 20],
            )

            result = self._run_batch(temp_dir, sales_file, price_dir)

            logs = "\n".join(result.logs)
            self.assertFalse(result.ok)
            self.assertIn("[报价表错误]", logs)
            self.assertIn("缺少大件报价模板：顺丰_大件", logs)
            self.assertIn("建议：请在 客户A 的报价表中新增 sheet「顺丰_大件」", logs)


if __name__ == "__main__":
    unittest.main()
