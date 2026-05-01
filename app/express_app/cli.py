"""Command-line entry point for the modular express fee app."""

from __future__ import annotations

import argparse
from pathlib import Path

from .core.calculator import (
    DEFAULT_PRICE_DIR,
    DEFAULT_OUTPUT_DIR,
    DEFAULT_SPLIT_DIR,
    detect_default_sales_file,
)
from .core.models import ExpressFeeBatchJobConfig, ExpressFeeJobConfig
from .core.calculator import run_express_fee_batch_job, run_express_fee_job


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="计算销售出库单中的快递费用")
    parser.add_argument(
        "--sales-file",
        type=Path,
        nargs="+",
        default=None,
        help="销售出库单路径，可一次传入多个；默认优先读取 express 根目录下的原始销售表",
    )
    parser.add_argument(
        "--price-dir",
        type=Path,
        default=DEFAULT_PRICE_DIR,
        help=f"快递报价表目录，默认：{DEFAULT_PRICE_DIR}",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="单文件输出路径；批量运行时请使用 --output-dir",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=DEFAULT_OUTPUT_DIR,
        help=f"批量输出目录，默认：{DEFAULT_OUTPUT_DIR}",
    )
    parser.add_argument(
        "--split-dir",
        type=Path,
        default=DEFAULT_SPLIT_DIR,
        help=f"客户每日快递费明细输出目录，默认：{DEFAULT_SPLIT_DIR}",
    )
    parser.add_argument(
        "--no-split",
        action="store_true",
        help="只生成总结果文件，不生成客户每日拆分文件",
    )
    parser.add_argument(
        "--no-customer-summary",
        action="store_true",
        help="不生成客户历史汇总表",
    )
    parser.add_argument(
        "--refresh-all-customers",
        action="store_true",
        help="刷新客户每日快递费明细目录下所有客户的历史汇总表",
    )
    parser.add_argument(
        "--round-digits",
        type=int,
        default=2,
        help="快递费用保留的小数位；传 -1 表示不四舍五入，默认：2",
    )
    return parser


def main() -> int:
    parser = build_arg_parser()
    args = parser.parse_args()
    sales_files = args.sales_file or [detect_default_sales_file()]
    round_digits = None if args.round_digits < 0 else args.round_digits
    if len(sales_files) > 1 and args.output is not None:
        parser.error("批量运行多个销售表时不能使用 --output，请改用 --output-dir。")

    if len(sales_files) == 1 and args.output is not None:
        config = ExpressFeeJobConfig(
            sales_file=sales_files[0],
            price_dir=args.price_dir,
            output_path=args.output,
            split_dir=args.split_dir,
            round_digits=round_digits,
            split_customer_daily_files=not args.no_split,
            generate_customer_history=not args.no_customer_summary,
            refresh_all_customers=args.refresh_all_customers,
        )
        result = run_express_fee_job(config)
    else:
        config = ExpressFeeBatchJobConfig(
            sales_files=sales_files,
            price_dir=args.price_dir,
            output_dir=args.output_dir,
            split_dir=args.split_dir,
            round_digits=round_digits,
            split_customer_daily_files=not args.no_split,
            generate_customer_history=not args.no_customer_summary,
            refresh_all_customers=args.refresh_all_customers,
        )
        result = run_express_fee_batch_job(config)

    print("\n".join(result.logs))
    return 0 if result.ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
