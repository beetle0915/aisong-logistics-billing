#!/usr/bin/env python3
"""Calculate express shipping fees for a sales outbound workbook.

The script reads a sales detail workbook and a directory of standardized
price workbooks named like "小张-快递报价.xlsx". It parses the original
express company name into "快递公司（标准版）", chooses the normal or large-piece
price template, then appends calculation result columns in the sales sheet:

    快递费用, 首重费用, 续重费用, 续重重量, 快递公司（标准版）

Pricing key:
    业务员 + 计费模板 + 省
"""

from __future__ import annotations

import argparse
import math
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .models import (
    ExpressCompanyKeywordRule,
    ExpressFeeBatchJobConfig,
    ExpressFeeBatchJobResult,
    ExpressFeeJobConfig,
    ExpressFeeJobResult,
    ExpressFeeRuleConfig,
)


DEFAULT_ROOT = Path("/Users/beetle/Desktop/express")
DEFAULT_SALES_FILE_CANDIDATES = [
    DEFAULT_ROOT / "【5.1】_销售出库单.xlsx",
    DEFAULT_ROOT / "原始数据表" / "【5.1】_销售出库单.xlsx",
]
DEFAULT_PRICE_DIR = DEFAULT_ROOT / "快递报价表"
DEFAULT_OUTPUT_DIR = DEFAULT_ROOT / "输出结果"
DEFAULT_SPLIT_DIR = DEFAULT_ROOT / "客户每日快递费明细"
OUTPUT_VERSION_SUFFIX = "v7_0_2"

STANDARD_EXPRESS_COLUMN = "快递公司（标准版）"
RAW_EXPRESS_COLUMN = "快递公司"
SHIPPING_DATE_COLUMN = "出库日期"

REQUIRED_SALES_COLUMNS = [SHIPPING_DATE_COLUMN, "业务员", RAW_EXPRESS_COLUMN, "省", "重量"]
RESULT_COLUMNS = ["快递费用", "首重费用", "续重费用", "续重重量", STANDARD_EXPRESS_COLUMN]
SUMMARY_SHEET_NAME = "快递费汇总"
DETAIL_SHEET_NAME = "快递明细"
CUSTOMER_HISTORY_SUMMARY_FILE = "客户快递费历史汇总.xlsx"
CUSTOMER_HISTORY_SHEET = "历史汇总"
CUSTOMER_EXPRESS_SUMMARY_SHEET = "按快递公司汇总"
CUSTOMER_DETAIL_INDEX_SHEET = "明细索引"
DAILY_DETAIL_FILE_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}_.+_快递费明细\.xlsx$")
PRICE_COLUMNS = {
    "province": "省份参照列",
    "first_price": "首重费用",
    "extra_price": "续重费用",
}
DEFAULT_EXPRESS_COMPANY_EXACT_MAP = {
    "顺丰速运新3": "顺丰",
    "顺丰速运开": "顺丰",
    "顺丰陈开专属": "顺丰",
    "申通E物流": "申通",
    "德邦快递盖子大360小泡货特惠_4": "德邦",
    "艾松仓库大件德邦": "德邦",
}
DEFAULT_EXPRESS_COMPANY_KEYWORD_RULES = [
    ExpressCompanyKeywordRule(keyword="顺丰", standard_name="顺丰"),
    ExpressCompanyKeywordRule(keyword="申通", standard_name="申通"),
    ExpressCompanyKeywordRule(keyword="德邦", standard_name="德邦"),
]
DEFAULT_LARGE_PIECE_COMPANIES = {"顺丰", "德邦"}
DEFAULT_LARGE_PIECE_THRESHOLD_KG = 20
DEFAULT_LARGE_PIECE_SUFFIX = "_大件"


def build_default_rule_config() -> ExpressFeeRuleConfig:
    return ExpressFeeRuleConfig(
        exact_company_map=dict(DEFAULT_EXPRESS_COMPANY_EXACT_MAP),
        keyword_company_rules=list(DEFAULT_EXPRESS_COMPANY_KEYWORD_RULES),
        large_piece_companies=set(DEFAULT_LARGE_PIECE_COMPANIES),
        large_piece_threshold_kg=DEFAULT_LARGE_PIECE_THRESHOLD_KG,
        large_piece_suffix=DEFAULT_LARGE_PIECE_SUFFIX,
    )


def normalize_rule_config(rule_config: ExpressFeeRuleConfig | None) -> ExpressFeeRuleConfig:
    if rule_config is None:
        return build_default_rule_config()

    default = build_default_rule_config()
    threshold = rule_config.large_piece_threshold_kg
    if threshold <= 0:
        threshold = default.large_piece_threshold_kg

    suffix = normalize_text(rule_config.large_piece_suffix) or default.large_piece_suffix
    return ExpressFeeRuleConfig(
        exact_company_map={
            normalize_text(raw): normalize_text(standard)
            for raw, standard in rule_config.exact_company_map.items()
            if normalize_text(raw) and normalize_text(standard)
        },
        keyword_company_rules=[
            ExpressCompanyKeywordRule(
                keyword=normalize_text(item.keyword),
                standard_name=normalize_text(item.standard_name),
            )
            for item in rule_config.keyword_company_rules
            if normalize_text(item.keyword) and normalize_text(item.standard_name)
        ],
        large_piece_companies={
            normalize_text(company)
            for company in rule_config.large_piece_companies
            if normalize_text(company)
        },
        large_piece_threshold_kg=threshold,
        large_piece_suffix=suffix,
    )


@dataclass(frozen=True)
class Price:
    first_price: float
    extra_price: float


@dataclass
class ProcessingSummary:
    total_rows: int = 0
    success_rows: int = 0
    failed_rows: int = 0
    errors: list[str] | None = None

    def __post_init__(self) -> None:
        if self.errors is None:
            self.errors = []

    def add_error(self, row_number: int, message: str) -> None:
        self.failed_rows += 1
        self.errors.append(f"第 {row_number} 行：{message}")


@dataclass
class SplitFile:
    customer: str
    shipping_date: str
    row_count: int
    output_path: Path


@dataclass
class SplitSummary:
    generated_files: list[SplitFile] | None = None
    skipped_rows: int = 0
    errors: list[str] | None = None

    def __post_init__(self) -> None:
        if self.generated_files is None:
            self.generated_files = []
        if self.errors is None:
            self.errors = []

    def add_error(self, row_number: int, message: str) -> None:
        self.skipped_rows += 1
        self.errors.append(f"第 {row_number} 行：{message}")


@dataclass
class DailyCustomerSummary:
    customer: str
    shipping_date: str
    row_count: int
    total_weight: float
    total_fee: float
    sf_count: int
    st_count: int
    db_count: int
    large_count: int
    express_summary: dict[str, dict[str, float]]
    file_name: str
    file_path: Path
    updated_at: str


@dataclass
class CustomerHistorySummary:
    customer: str
    day_count: int
    row_count: int
    total_fee: float
    output_path: Path | None = None
    errors: list[str] | None = None

    def __post_init__(self) -> None:
        if self.errors is None:
            self.errors = []


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def parse_number(value: Any, field_name: str) -> float:
    if value is None or value == "":
        raise ValueError(f"{field_name}为空")
    if isinstance(value, bool):
        raise ValueError(f"{field_name}不是数字：{value}")
    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{field_name}不是数字：{value}") from exc
    if math.isnan(number) or math.isinf(number):
        raise ValueError(f"{field_name}不是有效数字：{value}")
    return number


def parse_shipping_date(value: Any) -> str:
    if value is None or value == "":
        raise ValueError("出库日期为空")
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()

    text = normalize_text(value)
    if not text:
        raise ValueError("出库日期为空")

    normalized = text.replace("/", "-")
    date_part = normalized.split()[0]
    for fmt in ("%Y-%m-%d", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(date_part, fmt).date().isoformat()
        except ValueError:
            pass

    try:
        return datetime.fromisoformat(normalized).date().isoformat()
    except ValueError as exc:
        raise ValueError(f"出库日期无法解析：{value}") from exc


def sanitize_filename(value: Any) -> str:
    text = normalize_text(value)
    if not text:
        return "未命名客户"
    text = re.sub(r'[\\/:\*\?"<>\|]', "_", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text or "未命名客户"


def detect_default_sales_file() -> Path:
    for path in DEFAULT_SALES_FILE_CANDIDATES:
        try:
            exists = path.exists()
        except OSError:
            exists = False
        if exists:
            return path
    return DEFAULT_SALES_FILE_CANDIDATES[0]


def parse_salesman_from_filename(path: Path) -> str:
    stem = path.stem
    marker = "-快递报价"
    if marker in stem:
        return stem.split(marker, 1)[0].strip()
    return stem.split("-", 1)[0].strip()


def find_header_indexes(ws: openpyxl.worksheet.worksheet.Worksheet) -> dict[str, int]:
    headers: dict[str, int] = {}
    for column in range(1, ws.max_column + 1):
        header = normalize_text(ws.cell(row=1, column=column).value)
        if header:
            headers[header] = column
    return headers


def require_columns(headers: dict[str, int], required: list[str], source_name: str) -> None:
    missing = [column for column in required if column not in headers]
    if missing:
        joined = "、".join(missing)
        raise ValueError(f"{source_name} 缺少必要列：{joined}")


def parse_standard_express_company(
    raw_company: Any,
    available_companies: set[str],
    rule_config: ExpressFeeRuleConfig,
) -> str:
    raw_name = normalize_text(raw_company)
    if not raw_name:
        raise ValueError("快递公司为空")

    if raw_name in available_companies:
        return raw_name

    mapped_name = rule_config.exact_company_map.get(raw_name)
    if mapped_name is not None:
        if mapped_name not in available_companies:
            raise ValueError(f"标准快递公司不在报价表中：{mapped_name}")
        return mapped_name

    matches = {
        item.standard_name
        for item in rule_config.keyword_company_rules
        if item.keyword in raw_name and item.standard_name in available_companies
    }
    if len(matches) == 1:
        return matches.pop()
    if len(matches) > 1:
        joined = "、".join(sorted(matches))
        raise ValueError(f"快递公司匹配到多个标准名称：{raw_name} -> {joined}")

    raise ValueError(f"快递公司无法解析为标准名称：{raw_name}")


def resolve_price_sheet_name(
    express_company: str,
    weight: float,
    rule_config: ExpressFeeRuleConfig,
) -> str:
    if (
        express_company in rule_config.large_piece_companies
        and weight >= rule_config.large_piece_threshold_kg
    ):
        return f"{express_company}{rule_config.large_piece_suffix}"
    return express_company


def load_price_tables(price_dir: Path) -> dict[tuple[str, str, str], Price]:
    if not price_dir.exists():
        raise FileNotFoundError(f"报价目录不存在：{price_dir}")

    price_map: dict[tuple[str, str, str], Price] = {}
    price_files = sorted(
        path
        for path in price_dir.glob("*.xlsx")
        if not path.name.startswith("~$") and path.is_file()
    )
    if not price_files:
        raise FileNotFoundError(
            f"报价目录下没有 .xlsx 文件：{price_dir}。"
            "如果文件实际存在，请在桌面软件中重新选择报价表目录，"
            "让 macOS 授权当前应用访问该目录。"
        )

    for price_file in price_files:
        salesman = parse_salesman_from_filename(price_file)
        if not salesman:
            raise ValueError(f"无法从报价文件名解析业务员：{price_file.name}")

        workbook = openpyxl.load_workbook(price_file, data_only=True)
        for ws in workbook.worksheets:
            express_company = normalize_text(ws.title)
            headers = find_header_indexes(ws)
            require_columns(
                headers,
                [
                    PRICE_COLUMNS["province"],
                    PRICE_COLUMNS["first_price"],
                    PRICE_COLUMNS["extra_price"],
                ],
                f"{price_file.name}/{ws.title}",
            )

            province_col = headers[PRICE_COLUMNS["province"]]
            first_col = headers[PRICE_COLUMNS["first_price"]]
            extra_col = headers[PRICE_COLUMNS["extra_price"]]

            for row in range(2, ws.max_row + 1):
                province = normalize_text(ws.cell(row=row, column=province_col).value)
                if not province:
                    continue

                first_price = parse_number(
                    ws.cell(row=row, column=first_col).value,
                    f"{price_file.name}/{ws.title} 第 {row} 行首重费用",
                )
                extra_price = parse_number(
                    ws.cell(row=row, column=extra_col).value,
                    f"{price_file.name}/{ws.title} 第 {row} 行续重费用",
                )

                key = (salesman, express_company, province)
                if key in price_map:
                    raise ValueError(
                        "报价重复："
                        f"业务员={salesman}，快递={express_company}，省={province}"
                    )
                price_map[key] = Price(first_price=first_price, extra_price=extra_price)

    return price_map


def calculate_extra_weight(weight: float) -> int:
    if weight <= 0:
        raise ValueError(f"重量必须大于0：{weight}")
    if weight <= 1:
        return 0
    return math.ceil(weight - 1)


def calculate_fee(weight: float, price: Price, round_digits: int | None) -> tuple[float, int]:
    extra_weight = calculate_extra_weight(weight)
    fee = price.first_price + price.extra_price * extra_weight
    if round_digits is not None:
        fee = round(fee, round_digits)
    return fee, extra_weight


def resolve_output_path(sales_file: Path, output_path: Path | None) -> Path:
    if output_path is not None:
        return output_path

    output_name = f"{sales_file.stem}_快递费计算结果{sales_file.suffix}"
    return DEFAULT_OUTPUT_DIR / output_name


def resolve_batch_output_paths(
    sales_files: list[Path],
    output_dir: Path,
) -> dict[Path, Path]:
    """Return deterministic, non-conflicting output paths for a batch run."""

    output_paths: dict[Path, Path] = {}
    used_paths: set[Path] = set()
    for sales_file in sales_files:
        resolved_sales_file = sales_file.expanduser().resolve()
        base_name = f"{resolved_sales_file.stem}_快递费计算结果_{OUTPUT_VERSION_SUFFIX}"
        suffix = resolved_sales_file.suffix or ".xlsx"
        candidate = (output_dir / f"{base_name}{suffix}").resolve()
        counter = 2
        while candidate in used_paths:
            candidate = (output_dir / f"{base_name}_{counter}{suffix}").resolve()
            counter += 1

        output_paths[resolved_sales_file] = candidate
        used_paths.add(candidate)

    return output_paths


def delete_columns_by_header(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    column_names: list[str],
) -> None:
    target_names = set(column_names)
    delete_indexes = sorted(
        [
            column
            for column in range(1, ws.max_column + 1)
            if normalize_text(ws.cell(row=1, column=column).value) in target_names
        ],
        reverse=True,
    )
    for column_index in delete_indexes:
        ws.delete_cols(column_index)


def append_result_columns(
    ws: openpyxl.worksheet.worksheet.Worksheet,
) -> dict[str, int]:
    result_indexes: dict[str, int] = {}
    next_column = ws.max_column + 1
    for column_name in RESULT_COLUMNS:
        ws.cell(row=1, column=next_column).value = column_name
        result_indexes[column_name] = next_column
        next_column += 1

    return result_indexes


def process_sales_workbook(
    sales_file: Path,
    price_map: dict[tuple[str, str, str], Price],
    available_standard_companies: set[str],
    output_path: Path,
    round_digits: int | None,
    rule_config: ExpressFeeRuleConfig,
) -> ProcessingSummary:
    workbook = openpyxl.load_workbook(sales_file)
    ws = workbook.active

    delete_columns_by_header(ws, RESULT_COLUMNS)
    sales_headers = find_header_indexes(ws)
    require_columns(sales_headers, REQUIRED_SALES_COLUMNS, sales_file.name)
    result_columns = append_result_columns(ws)

    summary = ProcessingSummary(total_rows=max(ws.max_row - 1, 0))

    salesman_col = sales_headers["业务员"]
    raw_express_col = sales_headers[RAW_EXPRESS_COLUMN]
    province_col = sales_headers["省"]
    weight_col = sales_headers["重量"]

    for row in range(2, ws.max_row + 1):
        for column_name in RESULT_COLUMNS:
            ws.cell(row=row, column=result_columns[column_name]).value = None

        try:
            salesman = normalize_text(ws.cell(row=row, column=salesman_col).value)
            raw_express_company = ws.cell(row=row, column=raw_express_col).value
            express_company = parse_standard_express_company(
                raw_express_company,
                available_standard_companies,
                rule_config,
            )
            province = normalize_text(ws.cell(row=row, column=province_col).value)
            weight = parse_number(ws.cell(row=row, column=weight_col).value, "重量")
            price_sheet_name = resolve_price_sheet_name(
                express_company,
                weight,
                rule_config,
            )

            if not salesman:
                raise ValueError("业务员为空")
            if not province:
                raise ValueError("省为空")

            ws.cell(row=row, column=result_columns[STANDARD_EXPRESS_COLUMN]).value = (
                express_company
            )

            key = (salesman, price_sheet_name, province)
            price = price_map.get(key)
            if price is None:
                raise ValueError(
                    "找不到报价："
                    f"业务员={salesman}，计费模板={price_sheet_name}，省={province}"
                )

            fee, extra_weight = calculate_fee(weight, price, round_digits)

            ws.cell(row=row, column=result_columns["快递费用"]).value = fee
            ws.cell(row=row, column=result_columns["首重费用"]).value = price.first_price
            ws.cell(row=row, column=result_columns["续重费用"]).value = price.extra_price
            ws.cell(row=row, column=result_columns["续重重量"]).value = extra_weight
            summary.success_rows += 1
        except ValueError as exc:
            summary.add_error(row, str(exc))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    return summary


def number_or_zero(value: Any) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def format_number(value: float) -> float | int:
    if float(value).is_integer():
        return int(value)
    return round(value, 2)


def autosize_columns(ws: openpyxl.worksheet.worksheet.Worksheet) -> None:
    for column_index, column_cells in enumerate(ws.columns, start=1):
        column_letter = get_column_letter(column_index)
        max_length = 0
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, len(value))
        ws.column_dimensions[column_letter].width = min(max(max_length + 2, 10), 36)


def style_header_row(ws: openpyxl.worksheet.worksheet.Worksheet, row_number: int = 1) -> None:
    fill = PatternFill("solid", fgColor="D9EAF7")
    font = Font(bold=True)
    for cell in ws[row_number]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")


def build_summary_sheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    customer: str,
    shipping_date: str,
    headers: list[str],
    rows: list[list[Any]],
    rule_config: ExpressFeeRuleConfig,
) -> None:
    header_index = {header: index for index, header in enumerate(headers)}
    fee_index = header_index["快递费用"]
    weight_index = header_index["重量"]
    express_index = header_index[STANDARD_EXPRESS_COLUMN]

    total_rows = len(rows)
    total_fee = sum(number_or_zero(row[fee_index]) for row in rows)
    total_weight = sum(number_or_zero(row[weight_index]) for row in rows)
    large_count = sum(
        1
        for row in rows
        if normalize_text(row[express_index]) in rule_config.large_piece_companies
        and number_or_zero(row[weight_index]) >= rule_config.large_piece_threshold_kg
    )

    express_summary: dict[str, dict[str, float]] = {}
    for row in rows:
        express_company = normalize_text(row[express_index]) or "未识别"
        bucket = express_summary.setdefault(
            express_company,
            {"count": 0, "weight": 0.0, "fee": 0.0},
        )
        bucket["count"] += 1
        bucket["weight"] += number_or_zero(row[weight_index])
        bucket["fee"] += number_or_zero(row[fee_index])

    ws["A1"] = "客户每日快递费汇总"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells("A1:D1")

    summary_rows = [
        ("客户名称", customer),
        ("日期", shipping_date),
        ("总单数", total_rows),
        ("总快递费用", format_number(total_fee)),
        ("总重量", format_number(total_weight)),
        ("顺丰单数", int(express_summary.get("顺丰", {}).get("count", 0))),
        ("申通单数", int(express_summary.get("申通", {}).get("count", 0))),
        ("德邦单数", int(express_summary.get("德邦", {}).get("count", 0))),
        ("大件单数", large_count),
    ]
    for row_number, (label, value) in enumerate(summary_rows, start=3):
        ws.cell(row=row_number, column=1).value = label
        ws.cell(row=row_number, column=2).value = value

    table_start = 14
    ws.cell(row=table_start, column=1).value = STANDARD_EXPRESS_COLUMN
    ws.cell(row=table_start, column=2).value = "单数"
    ws.cell(row=table_start, column=3).value = "总重量"
    ws.cell(row=table_start, column=4).value = "总快递费用"
    style_header_row(ws, table_start)

    for offset, express_company in enumerate(sorted(express_summary), start=1):
        bucket = express_summary[express_company]
        row_number = table_start + offset
        ws.cell(row=row_number, column=1).value = express_company
        ws.cell(row=row_number, column=2).value = int(bucket["count"])
        ws.cell(row=row_number, column=3).value = format_number(bucket["weight"])
        ws.cell(row=row_number, column=4).value = format_number(bucket["fee"])

    for row_number in range(3, 12):
        ws.cell(row=row_number, column=1).font = Font(bold=True)
    autosize_columns(ws)


def build_detail_sheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    headers: list[str],
    rows: list[list[Any]],
) -> None:
    ws.append(headers)
    for row in rows:
        ws.append(row)
    style_header_row(ws)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    autosize_columns(ws)


def split_customer_daily_files(
    result_workbook_path: Path,
    split_dir: Path,
    rule_config: ExpressFeeRuleConfig,
) -> SplitSummary:
    workbook = openpyxl.load_workbook(result_workbook_path, data_only=True)
    ws = workbook.active
    headers = [ws.cell(row=1, column=column).value for column in range(1, ws.max_column + 1)]
    header_index = {header: index for index, header in enumerate(headers)}
    require_columns(
        {header: index + 1 for index, header in enumerate(headers) if header},
        [SHIPPING_DATE_COLUMN, "业务员", "快递费用", "重量", STANDARD_EXPRESS_COLUMN],
        result_workbook_path.name,
    )

    summary = SplitSummary()
    groups: dict[tuple[str, str], list[list[Any]]] = {}

    date_index = header_index[SHIPPING_DATE_COLUMN]
    customer_index = header_index["业务员"]
    for row_number in range(2, ws.max_row + 1):
        values = [ws.cell(row=row_number, column=column).value for column in range(1, ws.max_column + 1)]
        try:
            shipping_date = parse_shipping_date(values[date_index])
            customer = normalize_text(values[customer_index])
            if not customer:
                raise ValueError("业务员为空")
            groups.setdefault((shipping_date, customer), []).append(values)
        except ValueError as exc:
            summary.add_error(row_number, str(exc))

    split_dir.mkdir(parents=True, exist_ok=True)
    for (shipping_date, customer), rows in sorted(groups.items()):
        safe_customer = sanitize_filename(customer)
        customer_dir = split_dir / safe_customer
        customer_dir.mkdir(parents=True, exist_ok=True)
        output_path = customer_dir / f"{shipping_date}_{safe_customer}_快递费明细.xlsx"

        customer_workbook = openpyxl.Workbook()
        summary_sheet = customer_workbook.active
        summary_sheet.title = SUMMARY_SHEET_NAME
        build_summary_sheet(summary_sheet, customer, shipping_date, headers, rows, rule_config)

        detail_sheet = customer_workbook.create_sheet(DETAIL_SHEET_NAME)
        build_detail_sheet(detail_sheet, headers, rows)

        customer_workbook.save(output_path)
        summary.generated_files.append(
            SplitFile(
                customer=customer,
                shipping_date=shipping_date,
                row_count=len(rows),
                output_path=output_path,
            )
        )

    return summary


def is_daily_detail_file(path: Path) -> bool:
    return (
        path.is_file()
        and path.suffix == ".xlsx"
        and not path.name.startswith("~$")
        and path.name != CUSTOMER_HISTORY_SUMMARY_FILE
        and DAILY_DETAIL_FILE_PATTERN.match(path.name) is not None
    )


def read_daily_customer_summary(
    path: Path,
    rule_config: ExpressFeeRuleConfig,
) -> DailyCustomerSummary:
    workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if DETAIL_SHEET_NAME not in workbook.sheetnames:
        raise ValueError(f"缺少 {DETAIL_SHEET_NAME} sheet")

    ws = workbook[DETAIL_SHEET_NAME]
    row_iter = ws.iter_rows(values_only=True)
    try:
        headers = list(next(row_iter))
    except StopIteration as exc:
        raise ValueError("快递明细为空") from exc
    header_map = {header: index for index, header in enumerate(headers) if header}
    required = [SHIPPING_DATE_COLUMN, "业务员", "重量", "快递费用", STANDARD_EXPRESS_COLUMN]
    missing = [column for column in required if column not in header_map]
    if missing:
        raise ValueError(f"快递明细缺少必要列：{'、'.join(missing)}")

    dates: set[str] = set()
    customers: set[str] = set()
    row_count = 0
    total_weight = 0.0
    total_fee = 0.0
    large_count = 0
    express_summary: dict[str, dict[str, float]] = {}

    for row_number, row_values_tuple in enumerate(row_iter, start=2):
        row_values = list(row_values_tuple)
        if all(value in (None, "") for value in row_values):
            continue

        shipping_date = parse_shipping_date(row_values[header_map[SHIPPING_DATE_COLUMN]])
        customer = normalize_text(row_values[header_map["业务员"]])
        if not customer:
            raise ValueError(f"第 {row_number} 行业务员为空")

        express_company = normalize_text(row_values[header_map[STANDARD_EXPRESS_COLUMN]]) or "未识别"
        weight = number_or_zero(row_values[header_map["重量"]])
        fee = number_or_zero(row_values[header_map["快递费用"]])

        dates.add(shipping_date)
        customers.add(customer)
        row_count += 1
        total_weight += weight
        total_fee += fee
        if (
            express_company in rule_config.large_piece_companies
            and weight >= rule_config.large_piece_threshold_kg
        ):
            large_count += 1

        bucket = express_summary.setdefault(
            express_company,
            {"count": 0, "weight": 0.0, "fee": 0.0},
        )
        bucket["count"] += 1
        bucket["weight"] += weight
        bucket["fee"] += fee

    if row_count == 0:
        raise ValueError("快递明细为空")
    if len(dates) != 1:
        raise ValueError(f"快递明细包含多个日期：{'、'.join(sorted(dates))}")
    if len(customers) != 1:
        raise ValueError(f"快递明细包含多个客户：{'、'.join(sorted(customers))}")

    updated_at = datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
    return DailyCustomerSummary(
        customer=next(iter(customers)),
        shipping_date=next(iter(dates)),
        row_count=row_count,
        total_weight=total_weight,
        total_fee=total_fee,
        sf_count=int(express_summary.get("顺丰", {}).get("count", 0)),
        st_count=int(express_summary.get("申通", {}).get("count", 0)),
        db_count=int(express_summary.get("德邦", {}).get("count", 0)),
        large_count=large_count,
        express_summary=express_summary,
        file_name=path.name,
        file_path=path,
        updated_at=updated_at,
    )


def apply_sheet_basics(ws: openpyxl.worksheet.worksheet.Worksheet) -> None:
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    autosize_columns(ws)


def write_customer_history_sheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    daily_summaries: list[DailyCustomerSummary],
) -> None:
    headers = [
        "日期",
        "单数",
        "总重量",
        "总快递费用",
        "累计快递费用",
        "顺丰单数",
        "申通单数",
        "德邦单数",
        "大件单数",
        "明细文件",
    ]
    ws.append(headers)
    style_history_header(ws)

    fills = [
        PatternFill("solid", fgColor="EAF4FF"),
        PatternFill("solid", fgColor="EAF7EA"),
    ]
    cumulative_fee = 0.0
    previous_date = None
    color_index = 0
    for daily in daily_summaries:
        if previous_date is not None and daily.shipping_date != previous_date:
            color_index = 1 - color_index
        previous_date = daily.shipping_date
        cumulative_fee += daily.total_fee

        ws.append(
            [
                daily.shipping_date,
                daily.row_count,
                format_number(daily.total_weight),
                round(daily.total_fee, 2),
                round(cumulative_fee, 2),
                daily.sf_count,
                daily.st_count,
                daily.db_count,
                daily.large_count,
                daily.file_name,
            ]
        )
        row_number = ws.max_row
        for cell in ws[row_number]:
            cell.fill = fills[color_index]
            cell.alignment = Alignment(vertical="center")
        for column in (4, 5):
            ws.cell(row=row_number, column=column).number_format = "0.00"
        ws.cell(row=row_number, column=5).fill = PatternFill("solid", fgColor="FFF2CC")
        ws.cell(row=row_number, column=5).font = Font(bold=True)

    apply_sheet_basics(ws)


def style_history_header(ws: openpyxl.worksheet.worksheet.Worksheet) -> None:
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")


def write_customer_express_summary_sheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    daily_summaries: list[DailyCustomerSummary],
) -> None:
    headers = [STANDARD_EXPRESS_COLUMN, "总单数", "总重量", "总快递费用", "费用占比"]
    ws.append(headers)
    style_history_header(ws)

    express_summary: dict[str, dict[str, float]] = {}
    for daily in daily_summaries:
        for express_company, values in daily.express_summary.items():
            bucket = express_summary.setdefault(
                express_company,
                {"count": 0, "weight": 0.0, "fee": 0.0},
            )
            bucket["count"] += values["count"]
            bucket["weight"] += values["weight"]
            bucket["fee"] += values["fee"]

    total_fee = sum(values["fee"] for values in express_summary.values())
    for express_company in sorted(express_summary):
        values = express_summary[express_company]
        ratio = values["fee"] / total_fee if total_fee else 0
        ws.append(
            [
                express_company,
                int(values["count"]),
                format_number(values["weight"]),
                round(values["fee"], 2),
                ratio,
            ]
        )
        row_number = ws.max_row
        ws.cell(row=row_number, column=4).number_format = "0.00"
        ws.cell(row=row_number, column=5).number_format = "0.00%"

    apply_sheet_basics(ws)


def write_customer_detail_index_sheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    daily_summaries: list[DailyCustomerSummary],
) -> None:
    headers = ["日期", "明细文件名", "文件路径", "单数", "总快递费用", "最后更新时间"]
    ws.append(headers)
    style_history_header(ws)
    for daily in daily_summaries:
        ws.append(
            [
                daily.shipping_date,
                daily.file_name,
                str(daily.file_path),
                daily.row_count,
                round(daily.total_fee, 2),
                daily.updated_at,
            ]
        )
        ws.cell(row=ws.max_row, column=5).number_format = "0.00"

    apply_sheet_basics(ws)


def build_customer_history_summary(
    customer_dir: Path,
    rule_config: ExpressFeeRuleConfig,
) -> CustomerHistorySummary:
    customer = customer_dir.name
    output_path = customer_dir / CUSTOMER_HISTORY_SUMMARY_FILE
    errors: list[str] = []
    daily_summaries: list[DailyCustomerSummary] = []

    for detail_file in sorted(customer_dir.iterdir()):
        if not is_daily_detail_file(detail_file):
            continue
        try:
            daily_summaries.append(read_daily_customer_summary(detail_file, rule_config))
        except ValueError as exc:
            errors.append(f"{detail_file.name}：{exc}")

    daily_summaries.sort(key=lambda item: item.shipping_date)
    if not daily_summaries:
        return CustomerHistorySummary(
            customer=customer,
            day_count=0,
            row_count=0,
            total_fee=0.0,
            output_path=None,
            errors=errors or ["没有可用的每日明细文件"],
        )

    workbook = openpyxl.Workbook()
    history_sheet = workbook.active
    history_sheet.title = CUSTOMER_HISTORY_SHEET
    write_customer_history_sheet(history_sheet, daily_summaries)

    express_sheet = workbook.create_sheet(CUSTOMER_EXPRESS_SUMMARY_SHEET)
    write_customer_express_summary_sheet(express_sheet, daily_summaries)

    index_sheet = workbook.create_sheet(CUSTOMER_DETAIL_INDEX_SHEET)
    write_customer_detail_index_sheet(index_sheet, daily_summaries)

    workbook.save(output_path)
    return CustomerHistorySummary(
        customer=customer,
        day_count=len(daily_summaries),
        row_count=sum(item.row_count for item in daily_summaries),
        total_fee=sum(item.total_fee for item in daily_summaries),
        output_path=output_path,
        errors=errors,
    )


def refresh_customer_history_summaries(
    split_dir: Path,
    customers: set[str] | None = None,
    rule_config: ExpressFeeRuleConfig | None = None,
) -> list[CustomerHistorySummary]:
    resolved_rule_config = normalize_rule_config(rule_config)
    if not split_dir.exists():
        return [
            CustomerHistorySummary(
                customer=split_dir.name,
                day_count=0,
                row_count=0,
                total_fee=0.0,
                output_path=None,
                errors=[f"客户拆分目录不存在：{split_dir}"],
            )
        ]

    if customers is None:
        customer_dirs = [
            path
            for path in sorted(split_dir.iterdir())
            if path.is_dir() and not path.name.startswith(".")
        ]
    else:
        customer_dirs = [
            split_dir / sanitize_filename(customer)
            for customer in sorted(customers)
        ]

    summaries: list[CustomerHistorySummary] = []
    for customer_dir in customer_dirs:
        if not customer_dir.exists():
            summaries.append(
                CustomerHistorySummary(
                    customer=customer_dir.name,
                    day_count=0,
                    row_count=0,
                    total_fee=0.0,
                    output_path=None,
                    errors=[f"客户目录不存在：{customer_dir}"],
                )
            )
            continue
        summaries.append(build_customer_history_summary(customer_dir, resolved_rule_config))
    return summaries


def run_express_fee_job(config: ExpressFeeJobConfig) -> ExpressFeeJobResult:
    """Run one complete express fee job and return structured results.

    This is the stable core entry point for both the V6 CLI and the future
    desktop GUI. It keeps all file generation behavior from V5, but avoids
    printing directly so callers can decide how to display logs.
    """

    sales_file = config.sales_file.expanduser().resolve()
    price_dir = config.price_dir.expanduser().resolve()
    output_path = resolve_output_path(
        sales_file,
        config.output_path.expanduser().resolve()
        if config.output_path is not None
        else None,
    )
    split_dir = (
        config.split_dir.expanduser().resolve()
        if config.split_dir is not None
        else DEFAULT_SPLIT_DIR
    )
    rule_config = normalize_rule_config(config.rule_config)

    logs = [
        f"销售表：{sales_file}",
        f"报价目录：{price_dir}",
        f"输出文件：{output_path}",
    ]
    if config.split_customer_daily_files:
        logs.append(f"客户拆分目录：{split_dir}")

    price_map = load_price_tables(price_dir)
    available_standard_companies = {
        sheet_name
        for _, sheet_name, _ in price_map
        if not sheet_name.endswith(rule_config.large_piece_suffix)
    }
    processing_summary = process_sales_workbook(
        sales_file,
        price_map,
        available_standard_companies,
        output_path,
        config.round_digits,
        rule_config,
    )

    logs.extend(
        [
            "",
            f"共处理：{processing_summary.total_rows} 条",
            f"成功计算：{processing_summary.success_rows} 条",
            f"失败：{processing_summary.failed_rows} 条",
        ]
    )
    if processing_summary.errors:
        logs.append("")
        logs.append("失败明细：")
        logs.extend(processing_summary.errors)

    split_files: list[Path] = []
    split_errors: list[str] = []
    touched_customers: set[str] = set()
    if config.split_customer_daily_files:
        split_summary = split_customer_daily_files(output_path, split_dir, rule_config)
        logs.append("")
        logs.append("客户拆分：")
        logs.append(f"生成客户文件：{len(split_summary.generated_files)} 个")
        logs.append(f"跳过行数：{split_summary.skipped_rows} 条")
        for item in split_summary.generated_files:
            touched_customers.add(item.customer)
            split_files.append(item.output_path)
            logs.append(
                f"{item.customer} {item.shipping_date}："
                f"{item.row_count} 条 -> {item.output_path}"
            )
        if split_summary.errors:
            logs.append("")
            logs.append("拆分失败明细：")
            logs.extend(split_summary.errors)
            split_errors.extend(split_summary.errors)

    history_files: list[Path] = []
    history_errors: list[str] = []
    if config.generate_customer_history:
        if config.refresh_all_customers:
            refresh_customers = None
        elif touched_customers:
            refresh_customers = touched_customers
        else:
            refresh_customers = None if not config.split_customer_daily_files else set()

        if refresh_customers is None or refresh_customers:
            history_summaries = refresh_customer_history_summaries(
                split_dir,
                refresh_customers,
                rule_config,
            )
            logs.append("")
            logs.append("客户历史汇总：")
            for item in history_summaries:
                if item.output_path is None:
                    logs.append(f"{item.customer}：未生成")
                else:
                    history_files.append(item.output_path)
                    logs.append(
                        f"{item.customer}：{item.day_count} 天，"
                        f"{item.row_count} 单，累计费用 {round(item.total_fee, 2)} "
                        f"-> {item.output_path}"
                    )
                if item.errors:
                    for error in item.errors:
                        message = f"{item.customer}：{error}"
                        history_errors.append(message)
                        logs.append(f"  警告：{error}")

    return ExpressFeeJobResult(
        sales_file=sales_file,
        price_dir=price_dir,
        output_path=output_path,
        split_dir=split_dir,
        total_rows=processing_summary.total_rows,
        success_rows=processing_summary.success_rows,
        failed_rows=processing_summary.failed_rows,
        processing_errors=processing_summary.errors or [],
        split_files=split_files,
        split_errors=split_errors,
        history_files=history_files,
        history_errors=history_errors,
        logs=logs,
    )


def run_express_fee_batch_job(
    config: ExpressFeeBatchJobConfig,
) -> ExpressFeeBatchJobResult:
    """Run multiple sales workbooks and refresh customer history once.

    Each sales workbook is processed independently so a bad file does not
    prevent later files from being calculated. Customer history is generated
    after the whole batch, which avoids refreshing the same customer workbook
    repeatedly when a batch contains multiple dates.
    """

    sales_files = [path.expanduser().resolve() for path in config.sales_files]
    price_dir = config.price_dir.expanduser().resolve()
    output_dir = (
        config.output_dir.expanduser().resolve()
        if config.output_dir is not None
        else DEFAULT_OUTPUT_DIR
    )
    split_dir = (
        config.split_dir.expanduser().resolve()
        if config.split_dir is not None
        else DEFAULT_SPLIT_DIR
    )
    rule_config = normalize_rule_config(config.rule_config)
    output_paths = resolve_batch_output_paths(sales_files, output_dir)

    logs = [
        f"批量任务：{len(sales_files)} 个销售表",
        f"报价目录：{price_dir}",
        f"输出目录：{output_dir}",
    ]
    if config.split_customer_daily_files:
        logs.append(f"客户拆分目录：{split_dir}")

    job_results: list[ExpressFeeJobResult] = []
    touched_customers: set[str] = set()
    for index, sales_file in enumerate(sales_files, start=1):
        logs.extend(["", f"========== 第 {index}/{len(sales_files)} 个文件 =========="])
        output_path = output_paths[sales_file]
        job_config = ExpressFeeJobConfig(
            sales_file=sales_file,
            price_dir=price_dir,
            output_path=output_path,
            split_dir=split_dir,
            round_digits=config.round_digits,
            split_customer_daily_files=config.split_customer_daily_files,
            generate_customer_history=False,
            refresh_all_customers=False,
            rule_config=rule_config,
        )

        try:
            result = run_express_fee_job(job_config)
        except Exception as exc:  # Keep a batch moving if one workbook is bad.
            message = f"运行失败：{exc}"
            result = ExpressFeeJobResult(
                sales_file=sales_file,
                price_dir=price_dir,
                output_path=output_path,
                split_dir=split_dir,
                processing_errors=[message],
                logs=[
                    f"销售表：{sales_file}",
                    f"报价目录：{price_dir}",
                    f"输出文件：{output_path}",
                    message,
                ],
            )

        job_results.append(result)
        logs.extend(result.logs)
        for split_file in result.split_files:
            touched_customers.add(split_file.parent.name)

    history_files: list[Path] = []
    history_errors: list[str] = []
    if config.generate_customer_history:
        if config.refresh_all_customers:
            refresh_customers = None
        elif touched_customers:
            refresh_customers = touched_customers
        else:
            refresh_customers = None if not config.split_customer_daily_files else set()

        if refresh_customers is None or refresh_customers:
            history_summaries = refresh_customer_history_summaries(
                split_dir,
                refresh_customers,
                rule_config,
            )
            logs.append("")
            logs.append("批量客户历史汇总：")
            for item in history_summaries:
                if item.output_path is None:
                    logs.append(f"{item.customer}：未生成")
                else:
                    history_files.append(item.output_path)
                    logs.append(
                        f"{item.customer}：{item.day_count} 天，"
                        f"{item.row_count} 单，累计费用 {round(item.total_fee, 2)} "
                        f"-> {item.output_path}"
                    )
                if item.errors:
                    for error in item.errors:
                        message = f"{item.customer}：{error}"
                        history_errors.append(message)
                        logs.append(f"  警告：{error}")

    logs.extend(
        [
            "",
            "批量任务完成：",
            f"销售表数量：{len(job_results)} 个",
            f"成功文件数：{sum(1 for result in job_results if result.ok)} 个",
            f"错误文件数：{sum(1 for result in job_results if not result.ok)} 个",
            f"总处理行数：{sum(result.total_rows for result in job_results)} 条",
            f"总成功行数：{sum(result.success_rows for result in job_results)} 条",
            f"总失败行数：{sum(result.failed_rows for result in job_results)} 条",
            f"总结果文件：{len([result for result in job_results if result.output_path.exists()])} 个",
        ]
    )

    return ExpressFeeBatchJobResult(
        sales_files=sales_files,
        price_dir=price_dir,
        output_dir=output_dir,
        split_dir=split_dir,
        job_results=job_results,
        history_files=history_files,
        history_errors=history_errors,
        logs=logs,
    )


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="计算销售出库单中的快递费用")
    parser.add_argument(
        "--sales-file",
        type=Path,
        default=detect_default_sales_file(),
        help="销售出库单路径，默认优先读取 express 根目录下的原始销售表",
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
        help="输出文件路径，默认写入 express/输出结果 目录",
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

    sales_file = args.sales_file.expanduser().resolve()
    price_dir = args.price_dir.expanduser().resolve()
    output_path = resolve_output_path(
        sales_file,
        args.output.expanduser().resolve() if args.output is not None else None,
    )
    split_dir = args.split_dir.expanduser().resolve()
    round_digits = None if args.round_digits < 0 else args.round_digits
    rule_config = build_default_rule_config()

    print(f"销售表：{sales_file}")
    print(f"报价目录：{price_dir}")
    print(f"输出文件：{output_path}")
    if not args.no_split:
        print(f"客户拆分目录：{split_dir}")

    price_map = load_price_tables(price_dir)
    available_standard_companies = {
        sheet_name
        for _, sheet_name, _ in price_map
        if not sheet_name.endswith(rule_config.large_piece_suffix)
    }
    summary = process_sales_workbook(
        sales_file,
        price_map,
        available_standard_companies,
        output_path,
        round_digits,
        rule_config,
    )

    print()
    print(f"共处理：{summary.total_rows} 条")
    print(f"成功计算：{summary.success_rows} 条")
    print(f"失败：{summary.failed_rows} 条")

    if summary.errors:
        print()
        print("失败明细：")
        for error in summary.errors:
            print(error)

    split_summary = None
    touched_customers: set[str] = set()
    if not args.no_split:
        split_summary = split_customer_daily_files(output_path, split_dir, rule_config)
        print()
        print("客户拆分：")
        print(f"生成客户文件：{len(split_summary.generated_files)} 个")
        print(f"跳过行数：{split_summary.skipped_rows} 条")
        for item in split_summary.generated_files:
            touched_customers.add(item.customer)
            print(
                f"{item.customer} {item.shipping_date}："
                f"{item.row_count} 条 -> {item.output_path}"
            )
        if split_summary.errors:
            print()
            print("拆分失败明细：")
            for error in split_summary.errors:
                print(error)

    history_summaries: list[CustomerHistorySummary] = []
    if not args.no_customer_summary:
        if args.refresh_all_customers:
            refresh_customers = None
        elif touched_customers:
            refresh_customers = touched_customers
        else:
            refresh_customers = None if args.no_split else set()

        if refresh_customers is None or refresh_customers:
            history_summaries = refresh_customer_history_summaries(
                split_dir,
                refresh_customers,
                rule_config,
            )
            print()
            print("客户历史汇总：")
            for item in history_summaries:
                if item.output_path is None:
                    print(f"{item.customer}：未生成")
                else:
                    print(
                        f"{item.customer}：{item.day_count} 天，"
                        f"{item.row_count} 单，累计费用 {round(item.total_fee, 2)} "
                        f"-> {item.output_path}"
                    )
                if item.errors:
                    for error in item.errors:
                        print(f"  警告：{error}")

    if split_summary and split_summary.skipped_rows > 0:
        return 1
    return 0 if summary.failed_rows == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
