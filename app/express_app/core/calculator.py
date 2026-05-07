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
import copy
import math
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Callable

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from ..version import OUTPUT_VERSION_SUFFIX
from .models import (
    ExpressCompanyKeywordRule,
    ExpressFeeBatchJobConfig,
    ExpressFeeBatchJobResult,
    ExpressFeeJobConfig,
    ExpressFeeJobResult,
    ExpressFeePreflightFileResult,
    ExpressFeePreflightResult,
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

STANDARD_EXPRESS_COLUMN = "快递公司（标准版）"
RAW_EXPRESS_COLUMN = "快递公司"
SHIPPING_DATE_COLUMN = "出库日期"
SHIPPED_STATUS_COLUMN = "已发货"

REQUIRED_SALES_COLUMNS = [SHIPPING_DATE_COLUMN, "业务员", RAW_EXPRESS_COLUMN, "省", "重量"]
PRICE_VERSION_COLUMN = "报价版本"
PRICE_EFFECTIVE_DATE_COLUMN = "报价生效日期"
RESULT_COLUMNS = [
    "快递费用",
    "首重费用",
    "续重费用",
    "续重重量",
    STANDARD_EXPRESS_COLUMN,
    PRICE_VERSION_COLUMN,
    PRICE_EFFECTIVE_DATE_COLUMN,
]
SUMMARY_SHEET_NAME = "快递费汇总"
DETAIL_SHEET_NAME = "快递明细"
CUSTOMER_HISTORY_SUMMARY_FILE = "客户快递费历史汇总.xlsx"
CUSTOMER_HISTORY_SHEET = "历史汇总"
CUSTOMER_HISTORY_DETAIL_SHEET = "快递明细"
CUSTOMER_PAYMENT_SHEET = "收款记录"
CUSTOMER_PAYMENT_HEADERS = ["支付时间", "支付方式", "收款人", "已付金额", "收款时间"]
CUSTOMER_PAYMENT_MIN_ROWS = 20
CUSTOMER_HISTORY_DETAIL_SYSTEM_HEADERS = [
    "历史记录Key",
    "首次导入时间",
    "最后更新时间",
    "来源明细文件",
    "来源行号",
]
DAILY_DETAIL_FILE_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}_.+_快递费明细\.xlsx$")
PRICE_COLUMNS = {
    "province": "省份参照列",
    "first_price": "首重费用",
    "extra_price": "续重费用",
}
PRICE_TEMPLATE_REQUIRED_COLUMNS = [
    PRICE_COLUMNS["province"],
    PRICE_COLUMNS["first_price"],
    PRICE_COLUMNS["extra_price"],
]
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
INTEGER_ROUNDING_EXPRESS_COMPANIES = {"德邦"}
PRICE_VERSION_FILE_PATTERN = re.compile(r"^(\d{8})_?(.+)-快递报价.*\.xlsx$")
ProgressCallback = Callable[[str], None]
PROGRESS_ROW_INTERVAL = 200
PROGRESS_GROUP_INTERVAL = 20


def emit_progress(progress_callback: ProgressCallback | None, message: str) -> None:
    if progress_callback is not None:
        progress_callback(message)


def should_emit_progress(current: int, total: int, interval: int) -> bool:
    return current == 1 or current == total or current % interval == 0


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
    version: str = ""
    effective_date: str = ""
    price_file: Path | None = None
    sheet_name: str = ""
    row_number: int = 0


@dataclass(frozen=True)
class PriceWorkbookRef:
    salesman: str
    price_file: Path
    version: str = ""
    effective_date: str = ""
    is_versioned: bool = False

    @property
    def effective_sort_date(self) -> date:
        if not self.effective_date:
            return date.min
        return date.fromisoformat(self.effective_date)


@dataclass
class VersionedPriceCatalog:
    versions_by_salesman: dict[str, list[PriceWorkbookRef]]
    price_maps_by_file: dict[Path, dict[tuple[str, str], Price]]
    templates_by_file: dict[Path, set[str]]
    standard_companies: set[str]
    selection_cache: dict[tuple[str, str], PriceWorkbookRef | None] = field(default_factory=dict)

    @property
    def salesmen(self) -> set[str]:
        return set(self.versions_by_salesman)


@dataclass(frozen=True)
class PreflightPriceCatalog:
    standard_companies: set[str]
    templates_by_salesman: dict[str, set[str]]

    @property
    def salesmen(self) -> set[str]:
        return set(self.templates_by_salesman)


@dataclass(frozen=True)
class PriceTemplateRow:
    row_number: int
    province: str
    first_price: Any
    extra_price: Any


@dataclass(frozen=True)
class PriceTemplateSheet:
    sheet_name: str
    headers: list[str]
    rows: list[PriceTemplateRow]
    errors: list[str]

    @property
    def missing_columns(self) -> list[str]:
        return [column for column in PRICE_TEMPLATE_REQUIRED_COLUMNS if column not in self.headers]


@dataclass(frozen=True)
class PriceTemplateWorkbook:
    customer: str
    price_file: Path
    sheets: list[PriceTemplateSheet]
    errors: list[str]


@dataclass(frozen=True)
class PriceTemplateSummary:
    customer: str
    price_file: Path
    sheet_names: list[str]
    status: str
    errors: list[str]
    version: str = ""
    effective_date: str = ""

    @property
    def sheet_count(self) -> int:
        return len(self.sheet_names)


@dataclass(frozen=True)
class PriceTemplateCatalog:
    summaries: list[PriceTemplateSummary]
    errors: list[str]

    @property
    def customer_names(self) -> list[str]:
        names: list[str] = []
        for summary in self.summaries:
            if summary.customer not in names:
                names.append(summary.customer)
        return names


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
        if message.startswith("["):
            self.errors.append(message)
        else:
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


@dataclass
class HistoricalDetailRow:
    visible_values: list[Any]
    record_key: str
    first_imported_at: str
    last_updated_at: str
    source_file: str
    source_row: int


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


def parse_shipping_datetime(value: Any) -> datetime:
    if value is None or value == "":
        raise ValueError("出库日期为空")
    if isinstance(value, datetime):
        return value.replace(microsecond=0)
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())

    text = normalize_text(value)
    if not text:
        raise ValueError("出库日期为空")

    normalized = text.replace("/", "-")
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(normalized, fmt).replace(microsecond=0)
        except ValueError:
            pass

    try:
        return datetime.fromisoformat(normalized).replace(microsecond=0)
    except ValueError as exc:
        raise ValueError(f"出库日期无法解析：{value}") from exc


def build_historical_detail_key(
    outbound_number: Any,
    shipping_time: Any,
    shipped_status: Any = None,
) -> str:
    outbound_text = normalize_text(outbound_number)
    if not outbound_text:
        raise ValueError("出库单号为空")
    shipping_datetime = parse_shipping_datetime(shipping_time)
    shipped_text = normalize_text(shipped_status)
    return f"{outbound_text}|{shipping_datetime:%Y-%m-%d %H:%M:%S}|{shipped_text}"


def sanitize_filename(value: Any) -> str:
    text = normalize_text(value)
    if not text:
        return "未命名客户"
    text = re.sub(r'[\\/:\*\?"<>\|]', "_", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text or "未命名客户"


def customer_history_summary_file_name(customer: str) -> str:
    safe_customer = sanitize_filename(customer)
    return f"{safe_customer}_{CUSTOMER_HISTORY_SUMMARY_FILE}"


def is_customer_history_summary_file(path: Path) -> bool:
    return path.name == CUSTOMER_HISTORY_SUMMARY_FILE or path.name.endswith(
        f"_{CUSTOMER_HISTORY_SUMMARY_FILE}"
    )


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


def parse_price_version_filename(path: Path, salesman: str) -> tuple[str, str]:
    match = PRICE_VERSION_FILE_PATTERN.match(path.name)
    if match is None:
        raise ValueError(
            format_error_block(
                "报价表错误",
                f"报价文件名不符合版本格式：{path.name}",
                f"请将文件命名为类似「20260503{salesman}-快递报价.xlsx」。",
                stage="识别报价版本",
                location=str(path),
            )
        )

    version = match.group(1)
    file_salesman = normalize_text(match.group(2))
    try:
        effective_date = datetime.strptime(version, "%Y%m%d").date().isoformat()
    except ValueError as exc:
        raise ValueError(
            format_error_block(
                "报价表错误",
                f"报价文件名前 8 位不是有效日期：{version}",
                "请使用 YYYYMMDD 格式，例如 20260503。",
                stage="识别报价版本",
                location=str(path),
            )
        ) from exc

    if file_salesman != salesman:
        raise ValueError(
            format_error_block(
                "报价表错误",
                "报价文件名中的业务员和文件夹名不一致。",
                "请保持文件夹名和报价文件名中的业务员一致。",
                stage="识别报价版本",
                location=str(path),
                context={"文件夹业务员": salesman, "文件名业务员": file_salesman},
            )
        )
    return version, effective_date


def discover_price_workbook_refs(price_dir: Path) -> list[PriceWorkbookRef]:
    if not price_dir.exists():
        raise FileNotFoundError(
            format_error_block(
                "报价表错误",
                f"报价目录不存在：{price_dir}",
                "请重新选择报价表目录，确认目录存在且当前用户有读取权限。",
                stage="读取报价表",
                location=str(price_dir),
            )
        )

    refs: list[PriceWorkbookRef] = []
    for path in sorted(price_dir.glob("*.xlsx")):
        if not path.name.startswith("~$") and path.is_file():
            salesman = parse_salesman_from_filename(path)
            if salesman:
                refs.append(PriceWorkbookRef(salesman=salesman, price_file=path))

    for salesman_dir in sorted(path for path in price_dir.iterdir() if path.is_dir() and not path.name.startswith(".")):
        salesman = normalize_text(salesman_dir.name)
        for path in sorted(salesman_dir.glob("*.xlsx")):
            if path.name.startswith("~$") or not path.is_file():
                continue
            version, effective_date = parse_price_version_filename(path, salesman)
            refs.append(
                PriceWorkbookRef(
                    salesman=salesman,
                    price_file=path,
                    version=version,
                    effective_date=effective_date,
                    is_versioned=True,
                )
            )

    versioned_salesmen = {ref.salesman for ref in refs if ref.is_versioned}
    refs = [
        ref
        for ref in refs
        if ref.is_versioned or ref.salesman not in versioned_salesmen
    ]

    if not refs:
        raise FileNotFoundError(
            format_error_block(
                "报价表错误",
                "目录中没有找到可用的 .xlsx 报价文件。",
                "请重新选择报价表目录，确认选择的是包含业务员报价 Excel 的目录。",
                stage="读取报价表",
                location=str(price_dir),
                context={"报价目录": price_dir},
            )
        )

    seen_versions: dict[tuple[str, str], Path] = {}
    for ref in refs:
        if not ref.is_versioned:
            continue
        key = (ref.salesman, ref.effective_date)
        previous = seen_versions.get(key)
        if previous is not None:
            raise ValueError(
                format_error_block(
                    "报价表错误",
                    f"业务员「{ref.salesman}」存在重复报价版本：{ref.effective_date}",
                    "请同一天只保留一个业务员报价版本文件。",
                    stage="识别报价版本",
                    location=str(ref.price_file.parent),
                    context={"重复文件": f"{previous.name}、{ref.price_file.name}"},
                )
            )
        seen_versions[key] = ref.price_file

    return refs


def find_header_indexes(ws: openpyxl.worksheet.worksheet.Worksheet) -> dict[str, int]:
    headers: dict[str, int] = {}
    for column in range(1, ws.max_column + 1):
        header = normalize_text(ws.cell(row=1, column=column).value)
        if header:
            headers[header] = column
    return headers


def format_log_value(value: Any) -> str:
    text = normalize_text(value)
    return text or "空"


def format_error_block(
    category: str,
    reason: str,
    suggestion: str,
    *,
    stage: str | None = None,
    location: str | None = None,
    context: dict[str, Any] | None = None,
    system_error: Any | None = None,
) -> str:
    lines = [f"[{category}]"]
    if stage:
        lines.append(f"阶段：{stage}")
    if location:
        lines.append(f"位置：{location}")
    for key, value in (context or {}).items():
        lines.append(f"{key}：{format_log_value(value)}")
    lines.append(f"原因：{reason}")
    lines.append(f"建议：{suggestion}")
    if system_error:
        lines.append(f"系统返回：{system_error}")
    return "\n".join(lines)


def format_missing_columns_error(
    headers: dict[str, int],
    missing: list[str],
    source_name: str,
    *,
    category: str,
    stage: str,
    suggestion: str,
) -> str:
    detected_headers = "、".join(headers) if headers else "未识别到任何表头"
    return format_error_block(
        category,
        f"缺少必要列：{'、'.join(missing)}",
        suggestion,
        stage=stage,
        location=source_name,
        context={"当前识别到的列": detected_headers},
    )


def require_columns(
    headers: dict[str, int],
    required: list[str],
    source_name: str,
    *,
    category: str = "表结构错误",
    stage: str = "检查表头",
    suggestion: str = "请检查表头是否完整，并确认选择的是正确的 Excel 文件。",
) -> None:
    missing = [column for column in required if column not in headers]
    if missing:
        raise ValueError(
            format_missing_columns_error(
                headers,
                missing,
                source_name,
                category=category,
                stage=stage,
                suggestion=suggestion,
            )
        )


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


def load_price_workbook_prices(ref: PriceWorkbookRef) -> tuple[dict[tuple[str, str], Price], set[str]]:
    price_map: dict[tuple[str, str], Price] = {}
    templates: set[str] = set()
    workbook = openpyxl.load_workbook(ref.price_file, data_only=True, read_only=True)
    try:
        for ws in workbook.worksheets:
            express_company = normalize_text(ws.title)
            headers = find_header_indexes(ws)
            require_columns(
                headers,
                PRICE_TEMPLATE_REQUIRED_COLUMNS,
                f"{ref.price_file.name}/{ws.title}",
                category="报价表错误",
                stage="读取报价表",
                suggestion="请检查报价表 sheet 表头，必须包含「省份参照列」「首重费用」「续重费用」。",
            )
            templates.add(express_company)
            province_col = headers[PRICE_COLUMNS["province"]]
            first_col = headers[PRICE_COLUMNS["first_price"]]
            extra_col = headers[PRICE_COLUMNS["extra_price"]]

            for row in range(2, ws.max_row + 1):
                province = normalize_text(ws.cell(row=row, column=province_col).value)
                if not province:
                    continue

                first_price = parse_number(
                    ws.cell(row=row, column=first_col).value,
                    f"{ref.price_file.name}/{ws.title} 第 {row} 行首重费用",
                )
                extra_price = parse_number(
                    ws.cell(row=row, column=extra_col).value,
                    f"{ref.price_file.name}/{ws.title} 第 {row} 行续重费用",
                )

                key = (express_company, province)
                if key in price_map:
                    raise ValueError(
                        "报价重复："
                        f"业务员={ref.salesman}，快递={express_company}，省={province}"
                    )
                price_map[key] = Price(
                    first_price=first_price,
                    extra_price=extra_price,
                    version=ref.version,
                    effective_date=ref.effective_date,
                    price_file=ref.price_file,
                    sheet_name=express_company,
                    row_number=row,
                )
    finally:
        workbook.close()
    return price_map, templates


def load_price_workbook_templates(ref: PriceWorkbookRef) -> set[str]:
    templates: set[str] = set()
    workbook = openpyxl.load_workbook(ref.price_file, data_only=True, read_only=True)
    try:
        for ws in workbook.worksheets:
            express_company = normalize_text(ws.title)
            headers = find_header_indexes(ws)
            require_columns(
                headers,
                PRICE_TEMPLATE_REQUIRED_COLUMNS,
                f"{ref.price_file.name}/{ws.title}",
                category="报价表错误",
                stage="运行前验证",
                suggestion="请检查报价表 sheet 表头，必须包含「省份参照列」「首重费用」「续重费用」。",
            )
            templates.add(express_company)
    finally:
        workbook.close()
    return templates


def select_price_workbook_ref(
    catalog: VersionedPriceCatalog,
    salesman: str,
    shipping_date: str,
) -> PriceWorkbookRef | None:
    cache_key = (salesman, shipping_date)
    if cache_key in catalog.selection_cache:
        return catalog.selection_cache[cache_key]

    versions = catalog.versions_by_salesman.get(salesman, [])
    selected: PriceWorkbookRef | None = None
    current_date = date.fromisoformat(shipping_date)
    for ref in versions:
        if ref.is_versioned and ref.effective_sort_date > current_date:
            continue
        selected = ref
    catalog.selection_cache[cache_key] = selected
    return selected


def selected_price_workbook_refs_for_context(
    versions_by_salesman: dict[str, list[PriceWorkbookRef]],
    sales_context: dict[str, set[str]],
) -> set[Path]:
    catalog = VersionedPriceCatalog(
        versions_by_salesman=versions_by_salesman,
        price_maps_by_file={},
        templates_by_file={},
        standard_companies=set(),
    )
    selected_files: set[Path] = set()
    for salesman, shipping_dates in sales_context.items():
        for shipping_date in shipping_dates:
            selected = select_price_workbook_ref(catalog, salesman, shipping_date)
            if selected is not None:
                selected_files.add(selected.price_file)
    return selected_files


def build_versioned_price_catalog(
    price_dir: Path,
    sales_context: dict[str, set[str]] | None = None,
    rule_config: ExpressFeeRuleConfig | None = None,
) -> VersionedPriceCatalog:
    resolved_rule_config = normalize_rule_config(rule_config)
    refs = discover_price_workbook_refs(price_dir)
    if sales_context is not None:
        refs = [ref for ref in refs if ref.salesman in sales_context]
    versions_by_salesman: dict[str, list[PriceWorkbookRef]] = {}
    for ref in refs:
        versions_by_salesman.setdefault(ref.salesman, []).append(ref)
    for versions in versions_by_salesman.values():
        versions.sort(key=lambda item: (item.effective_sort_date, item.version, item.price_file.name))

    selected_files: set[Path] = set()
    if sales_context is None:
        selected_files = {ref.price_file for versions in versions_by_salesman.values() for ref in versions}
    else:
        selected_files = selected_price_workbook_refs_for_context(versions_by_salesman, sales_context)

    price_maps_by_file: dict[Path, dict[tuple[str, str], Price]] = {}
    templates_by_file: dict[Path, set[str]] = {}
    standard_companies: set[str] = set()
    ref_by_file = {ref.price_file: ref for versions in versions_by_salesman.values() for ref in versions}
    for price_file in sorted(selected_files, key=lambda path: str(path)):
        price_map, templates = load_price_workbook_prices(ref_by_file[price_file])
        price_maps_by_file[price_file] = price_map
        templates_by_file[price_file] = templates
        standard_companies.update(templates)

    return VersionedPriceCatalog(
        versions_by_salesman=versions_by_salesman,
        price_maps_by_file=price_maps_by_file,
        templates_by_file=templates_by_file,
        standard_companies={
            name
            for name in standard_companies
            if not name.endswith(resolved_rule_config.large_piece_suffix)
        },
    )


def build_preflight_price_catalog(
    price_dir: Path,
    sales_context: dict[str, set[str]],
    rule_config: ExpressFeeRuleConfig,
) -> VersionedPriceCatalog:
    refs = discover_price_workbook_refs(price_dir)
    refs = [ref for ref in refs if ref.salesman in sales_context]
    versions_by_salesman: dict[str, list[PriceWorkbookRef]] = {}
    for ref in refs:
        versions_by_salesman.setdefault(ref.salesman, []).append(ref)
    for versions in versions_by_salesman.values():
        versions.sort(key=lambda item: (item.effective_sort_date, item.version, item.price_file.name))

    catalog = VersionedPriceCatalog(
        versions_by_salesman=versions_by_salesman,
        price_maps_by_file={},
        templates_by_file={},
        standard_companies=set(),
    )
    selected_files = selected_price_workbook_refs_for_context(versions_by_salesman, sales_context)

    ref_by_file = {ref.price_file: ref for versions in versions_by_salesman.values() for ref in versions}
    for price_file in selected_files:
        templates = load_price_workbook_templates(ref_by_file[price_file])
        catalog.templates_by_file[price_file] = templates
        catalog.standard_companies.update(
            name for name in templates if not name.endswith(rule_config.large_piece_suffix)
        )
    catalog.standard_companies.update(
        standard
        for standard in rule_config.exact_company_map.values()
        if standard
    )
    catalog.standard_companies.update(
        item.standard_name
        for item in rule_config.keyword_company_rules
        if item.standard_name
    )
    return catalog


def load_price_tables(price_dir: Path) -> dict[tuple[str, str, str], Price]:
    catalog = build_versioned_price_catalog(price_dir)
    price_map: dict[tuple[str, str, str], Price] = {}
    for versions in catalog.versions_by_salesman.values():
        for ref in versions:
            if ref.price_file not in catalog.price_maps_by_file:
                continue
            for (express_company, province), price in catalog.price_maps_by_file[ref.price_file].items():
                key = (ref.salesman, express_company, province)
                if key in price_map:
                    raise ValueError(
                        "报价重复："
                        f"业务员={ref.salesman}，快递={express_company}，省={province}"
                    )
                price_map[key] = price
    return price_map


def load_preflight_price_catalog(price_dir: Path, rule_config: ExpressFeeRuleConfig) -> PreflightPriceCatalog:
    """Read the minimum price workbook metadata needed for input validation."""
    refs = discover_price_workbook_refs(price_dir)
    templates_by_salesman: dict[str, set[str]] = {}
    standard_companies: set[str] = set()
    for ref in refs:
        templates = load_price_workbook_templates(ref)
        templates_by_salesman.setdefault(ref.salesman, set()).update(templates)
        standard_companies.update(
            name for name in templates if not name.endswith(rule_config.large_piece_suffix)
        )
    standard_companies.update(
        standard
        for standard in rule_config.exact_company_map.values()
        if standard
    )
    standard_companies.update(
        item.standard_name
        for item in rule_config.keyword_company_rules
        if item.standard_name
    )
    return PreflightPriceCatalog(standard_companies=standard_companies, templates_by_salesman=templates_by_salesman)


def find_price_files_by_salesman(price_dir: Path) -> dict[str, Path]:
    if not price_dir.exists():
        raise FileNotFoundError(
            format_error_block(
                "报价表错误",
                f"报价目录不存在：{price_dir}",
                "请重新选择报价表目录，确认目录存在且当前用户有读取权限。",
                stage="读取报价表",
                location=str(price_dir),
            )
        )
    refs = [
        ref
        for ref in discover_price_workbook_refs(price_dir)
        if not ref.is_versioned
    ]
    if not refs:
        raise FileNotFoundError(
            format_error_block(
                "报价表错误",
                "目录中没有找到可用的 .xlsx 报价文件。",
                "请重新选择报价表目录，确认选择的是包含业务员报价 Excel 的目录。",
                stage="读取报价表",
                location=str(price_dir),
                context={"报价目录": price_dir},
            )
        )
    return {ref.salesman: ref.price_file for ref in refs}


def scan_price_template_catalog(price_dir: Path) -> PriceTemplateCatalog:
    refs = discover_price_workbook_refs(price_dir)
    summaries: list[PriceTemplateSummary] = []
    errors: list[str] = []
    for ref in refs:
        try:
            workbook = openpyxl.load_workbook(ref.price_file, data_only=True, read_only=True)
            sheet_names = list(workbook.sheetnames)
        except OSError as exc:
            sheet_names = []
            message = format_error_block(
                "报价表错误",
                "报价文件无法打开。",
                "请确认文件是有效的 .xlsx，且没有被 Excel 独占锁定。",
                stage="同步报价目录",
                location=str(ref.price_file),
                system_error=exc,
            )
            errors.append(message)
        status = "正常" if sheet_names else "读取失败"
        summaries.append(
            PriceTemplateSummary(
                customer=ref.salesman,
                price_file=ref.price_file,
                sheet_names=sheet_names,
                status=status,
                errors=[] if status == "正常" else [errors[-1]],
                version=ref.version,
                effective_date=ref.effective_date,
            )
        )

    summaries.sort(key=lambda item: (item.customer, item.effective_date, item.version, item.price_file.name))
    if not summaries:
        raise FileNotFoundError(
            format_error_block(
                "报价表错误",
                "目录中没有找到可用的客户报价文件。",
                "请确认报价文件名类似「客户A-快递报价.xlsx」。",
                stage="同步报价目录",
                location=str(price_dir),
            )
        )
    return PriceTemplateCatalog(summaries=summaries, errors=errors)


def load_price_template_workbook(
    price_file: Path,
    *,
    customer: str | None = None,
) -> PriceTemplateWorkbook:
    resolved_customer = customer or parse_salesman_from_filename(price_file)
    try:
        workbook = openpyxl.load_workbook(price_file, data_only=True, read_only=True)
    except OSError as exc:
        raise ValueError(
            format_error_block(
                "报价表错误",
                "报价文件无法打开。",
                "请确认文件是有效的 .xlsx，且没有被 Excel 独占锁定。",
                stage="读取客户报价",
                location=str(price_file),
                system_error=exc,
            )
        ) from exc

    sheets = [load_price_template_sheet(ws) for ws in workbook.worksheets]
    return PriceTemplateWorkbook(
        customer=resolved_customer,
        price_file=price_file,
        sheets=sheets,
        errors=[error for sheet in sheets for error in sheet.errors],
    )


def load_price_template_sheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
) -> PriceTemplateSheet:
    headers = list(find_header_indexes(ws))
    missing = [column for column in PRICE_TEMPLATE_REQUIRED_COLUMNS if column not in headers]
    if missing:
        return PriceTemplateSheet(
            sheet_name=normalize_text(ws.title),
            headers=headers,
            rows=[],
            errors=[f"{ws.title} 缺少必要列：{'、'.join(missing)}。请检查报价表表头。"],
        )

    header_indexes = find_header_indexes(ws)
    province_col = header_indexes[PRICE_COLUMNS["province"]]
    first_col = header_indexes[PRICE_COLUMNS["first_price"]]
    extra_col = header_indexes[PRICE_COLUMNS["extra_price"]]
    rows: list[PriceTemplateRow] = []
    for row_number in range(2, ws.max_row + 1):
        province = normalize_text(ws.cell(row=row_number, column=province_col).value)
        first_price = ws.cell(row=row_number, column=first_col).value
        extra_price = ws.cell(row=row_number, column=extra_col).value
        if not province and first_price in (None, "") and extra_price in (None, ""):
            continue
        rows.append(
            PriceTemplateRow(
                row_number=row_number,
                province=province,
                first_price=first_price,
                extra_price=extra_price,
            )
        )

    return PriceTemplateSheet(
        sheet_name=normalize_text(ws.title),
        headers=headers,
        rows=rows,
        errors=[],
    )


def calculate_extra_weight(weight: float) -> int:
    if weight <= 0:
        raise ValueError(f"重量必须大于0：{weight}")
    if weight <= 1:
        return 0
    return math.ceil(weight - 1)


def round_half_up_to_integer(value: float) -> int:
    return int(Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def calculate_fee(
    weight: float,
    price: Price,
    round_digits: int | None,
    express_company: str,
) -> tuple[float | int, int]:
    extra_weight = calculate_extra_weight(weight)
    fee = price.first_price + price.extra_price * extra_weight
    if express_company in INTEGER_ROUNDING_EXPRESS_COMPANIES:
        fee = round_half_up_to_integer(fee)
    elif round_digits is not None:
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


def collect_sales_price_context(sales_file: Path) -> dict[str, set[str]]:
    workbook = openpyxl.load_workbook(sales_file, read_only=True, data_only=True)
    try:
        ws = workbook.active
        headers = find_header_indexes(ws)
        require_columns(
            headers,
            [SHIPPING_DATE_COLUMN, "业务员"],
            str(sales_file),
            category="销售表结构错误",
            stage="读取销售出库单",
            suggestion="请确认销售表包含「出库日期」和「业务员」列。",
        )
        salesman_col = headers["业务员"] - 1
        date_col = headers[SHIPPING_DATE_COLUMN] - 1
        context: dict[str, set[str]] = {}
        for values in ws.iter_rows(min_row=2, values_only=True):
            salesman = normalize_text(values[salesman_col])
            if not salesman:
                continue
            try:
                shipping_date = parse_shipping_date(values[date_col])
            except ValueError:
                continue
            context.setdefault(salesman, set()).add(shipping_date)
        return context
    finally:
        workbook.close()


def merge_sales_price_contexts(contexts: list[dict[str, set[str]]]) -> dict[str, set[str]]:
    merged: dict[str, set[str]] = {}
    for context in contexts:
        for salesman, dates in context.items():
            merged.setdefault(salesman, set()).update(dates)
    return merged


def merge_visible_headers(existing_headers: list[str] | None, incoming_headers: list[str]) -> list[str]:
    merged = list(existing_headers or [])
    for header in incoming_headers:
        if header and header not in merged:
            merged.append(header)
    return merged


def align_row_values_to_headers(
    row_values: list[Any],
    source_headers: list[str],
    target_headers: list[str],
) -> list[Any]:
    source_map = {header: index for index, header in enumerate(source_headers) if header}
    return [
        row_values[source_map[header]]
        if header in source_map and source_map[header] < len(row_values)
        else None
        for header in target_headers
    ]


def build_available_standard_companies_from_catalog(
    price_catalog: VersionedPriceCatalog,
    rule_config: ExpressFeeRuleConfig,
) -> set[str]:
    standard_companies = {
        sheet_name
        for templates in price_catalog.templates_by_file.values()
        for sheet_name in templates
        if not sheet_name.endswith(rule_config.large_piece_suffix)
    }
    standard_companies.update(
        standard
        for standard in rule_config.exact_company_map.values()
        if standard
    )
    standard_companies.update(
        item.standard_name
        for item in rule_config.keyword_company_rules
        if item.standard_name
    )
    return standard_companies


def process_sales_workbook(
    sales_file: Path,
    price_map: dict[tuple[str, str, str], Price],
    available_standard_companies: set[str],
    output_path: Path,
    round_digits: int | None,
    rule_config: ExpressFeeRuleConfig,
    progress_callback: ProgressCallback | None = None,
    price_catalog: VersionedPriceCatalog | None = None,
) -> ProcessingSummary:
    workbook = openpyxl.load_workbook(sales_file)
    ws = workbook.active

    delete_columns_by_header(ws, RESULT_COLUMNS)
    sales_headers = find_header_indexes(ws)
    require_columns(
        sales_headers,
        REQUIRED_SALES_COLUMNS,
        str(sales_file),
        category="销售表结构错误",
        stage="读取销售出库单",
        suggestion="请恢复销售出库单表头，尤其是「出库日期」「业务员」「快递公司」「省」「重量」。",
    )
    result_columns = append_result_columns(ws)

    summary = ProcessingSummary(total_rows=max(ws.max_row - 1, 0))
    emit_progress(
        progress_callback,
        f"里程碑：开始计算快递费，共 {summary.total_rows} 行",
    )

    salesman_col = sales_headers["业务员"]
    raw_express_col = sales_headers[RAW_EXPRESS_COLUMN]
    province_col = sales_headers["省"]
    weight_col = sales_headers["重量"]
    shipping_date_col = sales_headers[SHIPPING_DATE_COLUMN]

    for row in range(2, ws.max_row + 1):
        for column_name in RESULT_COLUMNS:
            ws.cell(row=row, column=result_columns[column_name]).value = None

        salesman = ""
        raw_express_company = None
        express_company = ""
        province = ""
        weight_value = None
        weight: float | None = None
        price_sheet_name = ""
        shipping_date = ""
        try:
            salesman = normalize_text(ws.cell(row=row, column=salesman_col).value)
            shipping_date = parse_shipping_date(ws.cell(row=row, column=shipping_date_col).value)
            raw_express_company = ws.cell(row=row, column=raw_express_col).value
            express_company = parse_standard_express_company(
                raw_express_company,
                available_standard_companies,
                rule_config,
            )
            province = normalize_text(ws.cell(row=row, column=province_col).value)
            weight_value = ws.cell(row=row, column=weight_col).value
            weight = parse_number(weight_value, "重量")
            if weight <= 0:
                raise ValueError(f"重量必须大于0：{weight}")
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

            selected_ref: PriceWorkbookRef | None = None
            if price_catalog is not None:
                selected_ref = select_price_workbook_ref(price_catalog, salesman, shipping_date)
                if selected_ref is None:
                    raise ValueError(f"没有可用业务员报价版本：{salesman} / {shipping_date}")
                price = price_catalog.price_maps_by_file.get(selected_ref.price_file, {}).get(
                    (price_sheet_name, province)
                )
            else:
                key = (salesman, price_sheet_name, province)
                price = price_map.get(key)
            if price is None:
                if (
                    price_sheet_name.endswith(rule_config.large_piece_suffix)
                    and (
                        (
                            price_catalog is not None
                            and selected_ref is not None
                            and (express_company, province)
                            in price_catalog.price_maps_by_file.get(selected_ref.price_file, {})
                        )
                        or (
                            price_catalog is None
                            and (salesman, express_company, province) in price_map
                        )
                    )
                ):
                    reason = f"缺少大件报价模板：{price_sheet_name}"
                    suggestion = (
                        f"请在 {salesman} 的报价表中新增 sheet「{price_sheet_name}」，"
                        f"并填写「{province}」的首重和续重价格。"
                    )
                else:
                    reason = "找不到对应报价。"
                    suggestion = (
                        f"请检查 {salesman} 的报价表中是否存在 sheet「{price_sheet_name}」，"
                        f"并确认其中有「{province}」的首重和续重价格。"
                    )
                raise ValueError(
                    format_error_block(
                        "报价表错误",
                        reason,
                        suggestion,
                        stage="匹配报价",
                        location=f"销售表 第 {row} 行",
                        context={
                            "业务员": salesman,
                            "原始快递公司": raw_express_company,
                            "标准快递公司": express_company,
                            "计费模板": price_sheet_name,
                            "省": province,
                            "重量": weight,
                            "报价版本": selected_ref.version if selected_ref else "",
                            "报价生效日期": selected_ref.effective_date if selected_ref else "",
                        },
                    )
                )

            fee, extra_weight = calculate_fee(
                weight,
                price,
                round_digits,
                express_company,
            )

            ws.cell(row=row, column=result_columns["快递费用"]).value = fee
            ws.cell(row=row, column=result_columns["首重费用"]).value = price.first_price
            ws.cell(row=row, column=result_columns["续重费用"]).value = price.extra_price
            ws.cell(row=row, column=result_columns["续重重量"]).value = extra_weight
            ws.cell(row=row, column=result_columns[PRICE_VERSION_COLUMN]).value = price.version or None
            ws.cell(row=row, column=result_columns[PRICE_EFFECTIVE_DATE_COLUMN]).value = (
                price.effective_date or None
            )
            summary.success_rows += 1
        except ValueError as exc:
            message = str(exc)
            if not message.startswith("["):
                reason = message
                suggestion = "请检查该行的业务员、快递公司、省和重量是否填写完整且格式正确。"
                context = {
                    "行号": row,
                    "业务员": salesman,
                    "原始快递公司": raw_express_company,
                    "标准快递公司": express_company or "未识别",
                    "计费模板": price_sheet_name or "未确定",
                    "省": province,
                    "重量原值": weight_value,
                }
                if reason.startswith("重量不是数字"):
                    suggestion = "请将重量改为数字，例如 1、2.5、20。"
                elif reason.startswith("重量为空"):
                    suggestion = "请补充重量，重量必须是数字，例如 1、2.5、20。"
                elif reason.startswith("重量必须大于0"):
                    suggestion = "请将重量改为大于 0 的数字。"
                elif "快递公司" in reason:
                    suggestion = "请检查快递公司名称，或在规则配置中补充快递公司映射。"
                elif reason == "业务员为空":
                    suggestion = "请补充业务员，系统会用业务员匹配对应报价表。"
                elif reason == "省为空":
                    suggestion = "请补充省份，省份需要与报价表中的省份名称一致。"

                message = format_error_block(
                    "行级计算错误",
                    reason,
                    suggestion,
                    stage="计算快递费",
                    location=f"销售表 第 {row} 行",
                    context=context,
                )
            summary.add_error(row, message)

        processed_rows = row - 1
        if should_emit_progress(processed_rows, summary.total_rows, PROGRESS_ROW_INTERVAL):
            emit_progress(
                progress_callback,
                "进度：快递费计算"
                f"已处理 {processed_rows}/{summary.total_rows} 行，"
                f"成功 {summary.success_rows} 行，异常 {summary.failed_rows} 行",
            )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    emit_progress(
        progress_callback,
        "里程碑：快递费计算完成，"
        f"共 {summary.total_rows} 行，成功 {summary.success_rows} 行，"
        f"异常 {summary.failed_rows} 行",
    )
    return summary


def validate_sales_workbook_for_calculation(
    sales_file: Path,
    price_catalog: PreflightPriceCatalog | VersionedPriceCatalog,
    rule_config: ExpressFeeRuleConfig,
    require_history_detail_key: bool,
    progress_callback: ProgressCallback | None = None,
) -> ExpressFeePreflightFileResult:
    """Validate one sales workbook without writing any output files."""

    workbook = openpyxl.load_workbook(sales_file, read_only=True, data_only=True)
    try:
        ws = workbook.active
        sales_headers = find_header_indexes(ws)
        required_columns = list(REQUIRED_SALES_COLUMNS)
        if require_history_detail_key:
            required_columns.append("出库单号")
        require_columns(
            sales_headers,
            required_columns,
            str(sales_file),
            category="销售表结构错误",
            stage="运行前验证",
            suggestion="请恢复销售出库单表头，尤其是「出库日期」「业务员」「快递公司」「省」「重量」。",
        )

        total_rows = max(ws.max_row - 1, 0)
        result = ExpressFeePreflightFileResult(sales_file=sales_file, total_rows=total_rows)
        emit_progress(progress_callback, f"验证：开始检查 {sales_file.name}，共 {total_rows} 行")

        salesman_col = sales_headers["业务员"] - 1
        raw_express_col = sales_headers[RAW_EXPRESS_COLUMN] - 1
        province_col = sales_headers["省"] - 1
        weight_col = sales_headers["重量"] - 1
        shipping_date_col = sales_headers[SHIPPING_DATE_COLUMN] - 1
        outbound_col = sales_headers.get("出库单号")
        outbound_index = outbound_col - 1 if outbound_col is not None else None

        for row, values in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            salesman = ""
            raw_express_company = None
            express_company = ""
            province = ""
            weight_value = None
            shipping_date_value = None
            outbound_value = None
            try:
                salesman = normalize_text(values[salesman_col])
                if not salesman:
                    raise ValueError("业务员为空")
                if salesman not in price_catalog.salesmen:
                    raise ValueError(f"业务员没有价格表：{salesman}")

                shipping_date_value = values[shipping_date_col]
                if require_history_detail_key and outbound_index is not None:
                    outbound_value = values[outbound_index]
                    shipping_date = build_historical_detail_key(outbound_value, shipping_date_value).split("|", 1)[1]
                    shipping_date = shipping_date.split(" ", 1)[0]
                else:
                    shipping_date = parse_shipping_date(shipping_date_value)

                selected_ref: PriceWorkbookRef | None = None
                available_templates: set[str]
                if isinstance(price_catalog, VersionedPriceCatalog):
                    selected_ref = select_price_workbook_ref(price_catalog, salesman, shipping_date)
                    if selected_ref is None:
                        raise ValueError(f"没有可用业务员报价版本：{salesman} / {shipping_date}")
                    available_templates = price_catalog.templates_by_file.get(selected_ref.price_file, set())
                else:
                    available_templates = price_catalog.templates_by_salesman[salesman]

                raw_express_company = values[raw_express_col]
                express_company = parse_standard_express_company(
                    raw_express_company,
                    price_catalog.standard_companies,
                    rule_config,
                )
                province = normalize_text(values[province_col])
                weight_value = values[weight_col]
                weight = parse_number(weight_value, "重量")
                if weight <= 0:
                    raise ValueError(f"重量必须大于0：{weight}")

                price_sheet_name = resolve_price_sheet_name(express_company, weight, rule_config)
                if price_sheet_name not in available_templates:
                    raise ValueError(f"业务员报价表缺少计费模板：{salesman}/{price_sheet_name}")
                if not province:
                    raise ValueError("省为空")

                result.success_rows += 1
            except ValueError as exc:
                message = str(exc)
                if not message.startswith("["):
                    reason = message
                    suggestion = "请检查该行的业务员、快递公司、省、重量和出库日期是否填写完整且格式正确。"
                    context = {
                        "行号": row,
                        "业务员": salesman,
                        "原始快递公司": raw_express_company,
                        "标准快递公司": express_company or "未识别",
                        "省": province,
                        "重量原值": weight_value,
                        "出库日期原值": shipping_date_value,
                    }
                    if require_history_detail_key:
                        context["出库单号原值"] = outbound_value
                    if reason.startswith("重量不是数字"):
                        suggestion = "请将重量改为数字，例如 1、2.5、20。"
                    elif reason.startswith("重量为空"):
                        suggestion = "请补充重量，重量必须是数字，例如 1、2.5、20。"
                    elif reason.startswith("重量必须大于0"):
                        suggestion = "请将重量改为大于 0 的数字。"
                    elif reason.startswith("业务员没有价格表"):
                        suggestion = (
                            "请在报价目录下新增该业务员文件夹，并放入类似"
                            "「20260503业务员-快递报价.xlsx」的报价版本文件。"
                        )
                    elif reason.startswith("没有可用业务员报价版本"):
                        suggestion = "请在该业务员文件夹中补充出库日期当天或更早日期的报价版本文件。"
                    elif reason.startswith("业务员报价表缺少计费模板"):
                        suggestion = "请在该业务员报价表中补充对应快递公司 sheet，或检查快递公司映射是否正确。"
                    elif reason.startswith("出库日期"):
                        suggestion = "请补充正确的出库日期，例如 2026-04-07。"
                    elif reason == "出库单号为空":
                        suggestion = "请补充出库单号，客户历史明细需要用它和出库日期去重。"
                    elif "快递公司" in reason:
                        suggestion = "请检查快递公司名称，或在规则配置中补充快递公司映射。"
                    elif reason == "业务员为空":
                        suggestion = "请补充业务员，系统会用业务员匹配对应报价表。"
                    elif reason == "省为空":
                        suggestion = "请补充省份，省份需要与报价表中的省份名称一致。"

                    message = format_error_block(
                        "运行前验证错误",
                        reason,
                        suggestion,
                        stage="运行前验证",
                        location=f"销售表 第 {row} 行",
                        context=context,
                    )
                result.failed_rows += 1
                result.errors.append(message)

            processed_rows = row - 1
            if should_emit_progress(processed_rows, result.total_rows, PROGRESS_ROW_INTERVAL):
                emit_progress(
                    progress_callback,
                    "验证："
                    f"{sales_file.name} 已检查 {processed_rows}/{result.total_rows} 行，"
                    f"通过 {result.success_rows} 行，问题 {result.failed_rows} 行",
                )

        return result
    finally:
        workbook.close()


def validate_express_fee_batch_job(
    config: ExpressFeeBatchJobConfig,
    progress_callback: ProgressCallback | None = None,
) -> ExpressFeePreflightResult:
    """Validate batch inputs before the GUI writes calculation outputs."""

    sales_files = [path.expanduser().resolve() for path in config.sales_files]
    price_dir = config.price_dir.expanduser().resolve()
    rule_config = normalize_rule_config(config.rule_config)
    result = ExpressFeePreflightResult(sales_files=sales_files, price_dir=price_dir)

    emit_progress(progress_callback, f"验证：开始运行前测试，共 {len(sales_files)} 个销售表")
    try:
        sales_context = merge_sales_price_contexts(
            [collect_sales_price_context(sales_file) for sales_file in sales_files]
        )
        price_catalog = build_preflight_price_catalog(price_dir, sales_context, rule_config)
    except Exception as exc:
        message = str(exc)
        if not message.startswith("["):
            message = format_error_block(
                "运行前验证错误",
                message,
                "请检查报价目录和报价 Excel 文件后重新测试。",
                stage="运行前验证",
                location=str(price_dir),
                system_error=exc.__class__.__name__,
            )
        result.errors.append(message)
        result.logs.extend(result.errors)
        return result

    emit_progress(
        progress_callback,
        "验证：报价目录检查完成，"
        f"识别 {len(price_catalog.salesmen)} 个业务员报价表，"
        f"{len(price_catalog.standard_companies)} 个快递公司模板",
    )

    for index, sales_file in enumerate(sales_files, start=1):
        emit_progress(
            progress_callback,
            f"验证：正在检查销售表 {index}/{len(sales_files)}：{sales_file.name}",
        )
        try:
            file_result = validate_sales_workbook_for_calculation(
                sales_file,
                price_catalog,
                rule_config,
                require_history_detail_key=(
                    config.split_customer_daily_files or config.generate_customer_history
                ),
                progress_callback=progress_callback,
            )
        except Exception as exc:
            message = str(exc)
            if not message.startswith("["):
                message = format_error_block(
                    "运行前验证错误",
                    message,
                    "请检查该销售出库单是否完整、可读取，表头是否符合要求。",
                    stage="运行前验证",
                    location=str(sales_file),
                    system_error=exc.__class__.__name__,
                )
            file_result = ExpressFeePreflightFileResult(
                sales_file=sales_file,
                failed_rows=1,
                errors=[message],
            )
        result.file_results.append(file_result)

    result.errors.extend(
        error
        for file_result in result.file_results
        for error in file_result.errors
    )
    if result.ok:
        result.logs.append(
            f"运行前测试通过：{len(result.file_results)} 个销售表，"
            f"共 {result.total_rows} 行可计算。"
        )
        emit_progress(progress_callback, result.logs[-1])
    else:
        result.logs.append(
            f"运行前测试未通过：{len(result.file_results)} 个销售表，"
            f"共 {result.failed_rows} 个问题需要处理。"
        )
        result.logs.extend(result.errors)
        emit_progress(progress_callback, result.logs[-(len(result.errors) + 1)])

    return result


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
    progress_callback: ProgressCallback | None = None,
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
    total_rows = max(ws.max_row - 1, 0)
    emit_progress(progress_callback, f"里程碑：开始整理客户每日明细，共 {total_rows} 行")
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
    group_items = sorted(groups.items())
    total_groups = len(group_items)
    emit_progress(progress_callback, f"里程碑：准备生成客户每日明细，共 {total_groups} 份")
    for group_index, ((shipping_date, customer), rows) in enumerate(group_items, start=1):
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
        if should_emit_progress(group_index, total_groups, PROGRESS_GROUP_INTERVAL):
            emit_progress(
                progress_callback,
                f"进度：客户每日明细已生成 {group_index}/{total_groups} 份",
            )

    emit_progress(
        progress_callback,
        "里程碑：客户每日明细生成完成，"
        f"共 {len(summary.generated_files)} 份，异常 {summary.skipped_rows} 行",
    )

    return summary


def is_daily_detail_file(path: Path) -> bool:
    return (
        path.is_file()
        and path.suffix == ".xlsx"
        and not path.name.startswith("~$")
        and not is_customer_history_summary_file(path)
        and DAILY_DETAIL_FILE_PATTERN.match(path.name) is not None
    )


def read_daily_customer_summary(
    path: Path,
    rule_config: ExpressFeeRuleConfig,
) -> DailyCustomerSummary:
    workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
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
    finally:
        workbook.close()


def apply_sheet_basics(ws: openpyxl.worksheet.worksheet.Worksheet) -> None:
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    autosize_columns(ws)


def apply_table_border(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    max_row: int | None = None,
    max_column: int | None = None,
) -> None:
    border = Border(
        left=Side(style="thin", color="B7B7B7"),
        right=Side(style="thin", color="B7B7B7"),
        top=Side(style="thin", color="B7B7B7"),
        bottom=Side(style="thin", color="B7B7B7"),
    )
    last_row = max_row or ws.max_row
    last_column = max_column or ws.max_column
    for row in ws.iter_rows(
        min_row=1,
        max_row=last_row,
        min_col=1,
        max_col=last_column,
    ):
        for cell in row:
            cell.border = border


def write_customer_history_sheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    daily_summaries: list[DailyCustomerSummary],
) -> None:
    headers = [
        "日期",
        "单数",
        "总重量",
        "今日快递总消费",
        "累计快递费用",
        "今日收款",
        "累计收款",
        "当前余额",
        "顺丰单数",
        "申通单数",
        "德邦单数",
        "大件单数",
    ]
    ws.append(headers)
    style_history_header(ws)
    ws.row_dimensions[1].height = 30

    fills = [
        PatternFill("solid", fgColor="EAF4FF"),
        PatternFill("solid", fgColor="EAF7EA"),
    ]
    highlight_fill = PatternFill("solid", fgColor="FFF2CC")
    expense_fill = PatternFill("solid", fgColor="FCE4D6")
    income_fill = PatternFill("solid", fgColor="E2F0D9")
    cumulative_fee = 0.0
    previous_date = None
    color_index = 0
    for daily in daily_summaries:
        if previous_date is not None and daily.shipping_date != previous_date:
            color_index = 1 - color_index
        previous_date = daily.shipping_date
        cumulative_fee += daily.total_fee

        row_number = ws.max_row + 1
        ws.append(
            [
                datetime.fromisoformat(daily.shipping_date).date(),
                daily.row_count,
                format_number(daily.total_weight),
                round(daily.total_fee, 2),
                round(cumulative_fee, 2),
                (
                    f'=SUMIFS(\'{CUSTOMER_PAYMENT_SHEET}\'!$D:$D,'
                    f'\'{CUSTOMER_PAYMENT_SHEET}\'!$A:$A,">="&A{row_number},'
                    f'\'{CUSTOMER_PAYMENT_SHEET}\'!$A:$A,"<"&A{row_number}+1)'
                ),
                (
                    f'=SUMIFS(\'{CUSTOMER_PAYMENT_SHEET}\'!$D:$D,'
                    f'\'{CUSTOMER_PAYMENT_SHEET}\'!$A:$A,"<"&A{row_number}+1,'
                    f'\'{CUSTOMER_PAYMENT_SHEET}\'!$A:$A,"<>")'
                ),
                f"=G{row_number}-E{row_number}",
                daily.sf_count,
                daily.st_count,
                daily.db_count,
                daily.large_count,
            ]
        )
        ws.row_dimensions[row_number].height = 28
        for cell in ws[row_number]:
            cell.fill = fills[color_index]
            cell.alignment = Alignment(vertical="center")
        ws.cell(row=row_number, column=1).number_format = "yyyy-mm-dd"
        for column in (4, 5, 6, 7, 8):
            ws.cell(row=row_number, column=column).number_format = "0.00"
        ws.cell(row=row_number, column=4).fill = expense_fill
        ws.cell(row=row_number, column=6).fill = income_fill
        for column in (4, 5, 6, 7, 8):
            ws.cell(row=row_number, column=column).font = Font(bold=True)
        for column in (5, 7, 8):
            ws.cell(row=row_number, column=column).fill = highlight_fill

    apply_sheet_basics(ws)
    history_widths = {
        "A": 14,
        "B": 10,
        "C": 12,
        "D": 20,
        "E": 18,
        "F": 16,
        "G": 16,
        "H": 16,
        "I": 12,
        "J": 12,
        "K": 12,
        "L": 12,
    }
    for column_letter, width in history_widths.items():
        ws.column_dimensions[column_letter].width = width
    apply_table_border(ws, max_column=len(headers))


def style_history_header(ws: openpyxl.worksheet.worksheet.Worksheet) -> None:
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")


def copy_worksheet_contents(
    source: openpyxl.worksheet.worksheet.Worksheet,
    target: openpyxl.worksheet.worksheet.Worksheet,
) -> None:
    for row in source.iter_rows():
        for source_cell in row:
            target_cell = target.cell(row=source_cell.row, column=source_cell.column)
            target_cell.value = source_cell.value
            if source_cell.has_style:
                target_cell.font = copy.copy(source_cell.font)
                target_cell.fill = copy.copy(source_cell.fill)
                target_cell.border = copy.copy(source_cell.border)
                target_cell.alignment = copy.copy(source_cell.alignment)
                target_cell.number_format = source_cell.number_format
                target_cell.protection = copy.copy(source_cell.protection)
            if source_cell.hyperlink:
                target_cell._hyperlink = copy.copy(source_cell.hyperlink)
            if source_cell.comment:
                target_cell.comment = copy.copy(source_cell.comment)

    for column_letter, dimension in source.column_dimensions.items():
        target.column_dimensions[column_letter].width = dimension.width
    for row_index, dimension in source.row_dimensions.items():
        target.row_dimensions[row_index].height = dimension.height
    if source.freeze_panes:
        target.freeze_panes = source.freeze_panes
    if source.auto_filter.ref:
        target.auto_filter.ref = source.auto_filter.ref


def find_existing_customer_history_workbook(paths: list[Path]) -> Path | None:
    for path in paths:
        if path.exists():
            return path
    return None


def read_existing_history_detail_rows(
    workbook_path: Path | None,
) -> tuple[list[str] | None, dict[str, HistoricalDetailRow]]:
    if workbook_path is None:
        return None, {}

    workbook = openpyxl.load_workbook(workbook_path, data_only=False, read_only=True)
    try:
        if CUSTOMER_HISTORY_DETAIL_SHEET not in workbook.sheetnames:
            return None, {}

        ws = workbook[CUSTOMER_HISTORY_DETAIL_SHEET]
        rows = ws.iter_rows(values_only=True)
        try:
            headers = [normalize_text(value) for value in next(rows)]
        except StopIteration:
            return None, {}

        system_start_index = None
        for index, header in enumerate(headers):
            if header == CUSTOMER_HISTORY_DETAIL_SYSTEM_HEADERS[0]:
                system_start_index = index
                break
        if system_start_index is None:
            visible_headers = [header for header in headers if header]
            system_start_index = len(visible_headers)
        else:
            visible_headers = headers[:system_start_index]

        key_index = system_start_index
        first_imported_index = system_start_index + 1
        last_updated_index = system_start_index + 2
        source_file_index = system_start_index + 3
        source_row_index = system_start_index + 4
        header_map = {header: index for index, header in enumerate(visible_headers) if header}

        detail_rows: dict[str, HistoricalDetailRow] = {}
        for row_values_tuple in rows:
            row_values = list(row_values_tuple)
            if all(value in (None, "") for value in row_values):
                continue
            record_key = normalize_text(row_values[key_index] if key_index < len(row_values) else None)
            if "出库单号" in header_map and SHIPPING_DATE_COLUMN in header_map:
                shipped_status_index = header_map.get(SHIPPED_STATUS_COLUMN)
                shipped_status = (
                    row_values[shipped_status_index]
                    if shipped_status_index is not None and shipped_status_index < len(row_values)
                    else None
                )
                try:
                    record_key = build_historical_detail_key(
                        row_values[header_map["出库单号"]],
                        row_values[header_map[SHIPPING_DATE_COLUMN]],
                        shipped_status,
                    )
                except ValueError:
                    if not record_key:
                        continue
            elif not record_key:
                continue

            detail_rows[record_key] = HistoricalDetailRow(
                visible_values=row_values[:system_start_index],
                record_key=record_key,
                first_imported_at=normalize_text(
                    row_values[first_imported_index]
                    if first_imported_index < len(row_values)
                    else None
                ),
                last_updated_at=normalize_text(
                    row_values[last_updated_index] if last_updated_index < len(row_values) else None
                ),
                source_file=normalize_text(
                    row_values[source_file_index] if source_file_index < len(row_values) else None
                ),
                source_row=int(number_or_zero(row_values[source_row_index]))
                if source_row_index < len(row_values)
                else 0,
            )

        return visible_headers or None, detail_rows
    finally:
        workbook.close()


def load_daily_detail_rows_for_history(
    customer_dir: Path,
    existing_workbook_path: Path | None,
) -> tuple[list[str], list[HistoricalDetailRow], list[str]]:
    visible_headers, detail_rows = read_existing_history_detail_rows(existing_workbook_path)
    errors: list[str] = []
    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for detail_file in sorted(customer_dir.iterdir()):
        if not is_daily_detail_file(detail_file):
            continue
        workbook = openpyxl.load_workbook(detail_file, data_only=True, read_only=True)
        try:
            if DETAIL_SHEET_NAME not in workbook.sheetnames:
                errors.append(f"{detail_file.name}：缺少 {DETAIL_SHEET_NAME} sheet")
                continue
            ws = workbook[DETAIL_SHEET_NAME]
            row_iter = ws.iter_rows(values_only=True)
            try:
                headers = list(next(row_iter))
            except StopIteration:
                errors.append(f"{detail_file.name}：快递明细为空")
                continue

            normalized_headers = [normalize_text(header) for header in headers]
            header_map = {header: index for index, header in enumerate(normalized_headers) if header}
            missing = [
                column
                for column in ("出库单号", SHIPPING_DATE_COLUMN)
                if column not in header_map
            ]
            if missing:
                errors.append(f"{detail_file.name}：快递明细缺少必要列：{'、'.join(missing)}")
                continue

            if visible_headers is None:
                visible_headers = normalized_headers
            else:
                merged_headers = merge_visible_headers(visible_headers, normalized_headers)
                if merged_headers != visible_headers:
                    detail_rows = {
                        key: HistoricalDetailRow(
                            visible_values=align_row_values_to_headers(
                                detail_row.visible_values,
                                visible_headers,
                                merged_headers,
                            ),
                            record_key=detail_row.record_key,
                            first_imported_at=detail_row.first_imported_at,
                            last_updated_at=detail_row.last_updated_at,
                            source_file=detail_row.source_file,
                            source_row=detail_row.source_row,
                        )
                        for key, detail_row in detail_rows.items()
                    }
                    visible_headers = merged_headers

            for row_number, row_values_tuple in enumerate(row_iter, start=2):
                row_values = list(row_values_tuple)
                if all(value in (None, "") for value in row_values):
                    continue
                try:
                    shipped_status_index = header_map.get(SHIPPED_STATUS_COLUMN)
                    shipped_status = (
                        row_values[shipped_status_index]
                        if shipped_status_index is not None and shipped_status_index < len(row_values)
                        else None
                    )
                    record_key = build_historical_detail_key(
                        row_values[header_map["出库单号"]],
                        row_values[header_map[SHIPPING_DATE_COLUMN]],
                        shipped_status,
                    )
                except ValueError as exc:
                    errors.append(f"{detail_file.name} 第 {row_number} 行：{exc}")
                    continue

                previous = detail_rows.get(record_key)
                first_imported_at = previous.first_imported_at if previous else now_text
                detail_rows[record_key] = HistoricalDetailRow(
                    visible_values=align_row_values_to_headers(
                        row_values,
                        normalized_headers,
                        visible_headers,
                    ),
                    record_key=record_key,
                    first_imported_at=first_imported_at or now_text,
                    last_updated_at=now_text,
                    source_file=detail_file.name,
                    source_row=row_number,
                )
        finally:
            workbook.close()

    return visible_headers or [], list(detail_rows.values()), errors


def write_customer_history_detail_sheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    visible_headers: list[str],
    detail_rows: list[HistoricalDetailRow],
) -> None:
    headers = visible_headers + CUSTOMER_HISTORY_DETAIL_SYSTEM_HEADERS
    ws.append(headers)
    style_history_header(ws)
    ws.row_dimensions[1].height = 28

    visible_column_count = len(visible_headers)
    for detail_row in detail_rows:
        values = list(detail_row.visible_values)
        if len(values) < visible_column_count:
            values.extend([None] * (visible_column_count - len(values)))
        ws.append(
            values[:visible_column_count]
            + [
                detail_row.record_key,
                detail_row.first_imported_at,
                detail_row.last_updated_at,
                detail_row.source_file,
                detail_row.source_row,
            ]
        )
        ws.row_dimensions[ws.max_row].height = 24

    style_header_row(ws)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    autosize_columns(ws)
    apply_table_border(ws)

    for column_index in range(visible_column_count + 1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(column_index)].hidden = True


def write_customer_payment_sheet(ws: openpyxl.worksheet.worksheet.Worksheet) -> None:
    is_empty_sheet = ws.max_row == 1 and ws.max_column == 1 and ws["A1"].value is None
    if is_empty_sheet:
        for column, header in enumerate(CUSTOMER_PAYMENT_HEADERS, start=1):
            ws.cell(row=1, column=column).value = header
    elif [
        ws.cell(row=1, column=column).value
        for column in range(1, len(CUSTOMER_PAYMENT_HEADERS) + 1)
    ] != CUSTOMER_PAYMENT_HEADERS:
        ws.insert_rows(1)
        for column, header in enumerate(CUSTOMER_PAYMENT_HEADERS, start=1):
            ws.cell(row=1, column=column).value = header

    style_history_header(ws)
    max_row = max(ws.max_row, CUSTOMER_PAYMENT_MIN_ROWS)
    for row_number in range(2, max_row + 1):
        ws.cell(row=row_number, column=1).number_format = "yyyy-mm-dd"
        ws.cell(row=row_number, column=4).number_format = "0.00"
        ws.cell(row=row_number, column=5).number_format = "yyyy-mm-dd hh:mm"
        ws.row_dimensions[row_number].height = 24
    apply_sheet_basics(ws)
    payment_widths = {
        "A": 14,
        "B": 14,
        "C": 14,
        "D": 14,
        "E": 20,
    }
    for column_letter, width in payment_widths.items():
        ws.column_dimensions[column_letter].width = width
    ws.row_dimensions[1].height = 28
    apply_table_border(ws, max_row=max_row, max_column=len(CUSTOMER_PAYMENT_HEADERS))


def add_or_preserve_customer_payment_sheet(
    workbook: openpyxl.Workbook,
    existing_summary_paths: list[Path],
) -> None:
    payment_sheet = workbook.create_sheet(CUSTOMER_PAYMENT_SHEET)
    for existing_summary_path in existing_summary_paths:
        if not existing_summary_path.exists():
            continue
        existing_workbook = openpyxl.load_workbook(existing_summary_path, data_only=False)
        if CUSTOMER_PAYMENT_SHEET in existing_workbook.sheetnames:
            copy_worksheet_contents(existing_workbook[CUSTOMER_PAYMENT_SHEET], payment_sheet)
            break
    write_customer_payment_sheet(payment_sheet)


def build_customer_history_summary(
    customer_dir: Path,
    rule_config: ExpressFeeRuleConfig,
    progress_callback: ProgressCallback | None = None,
) -> CustomerHistorySummary:
    customer = customer_dir.name
    output_path = customer_dir / customer_history_summary_file_name(customer)
    legacy_output_path = customer_dir / CUSTOMER_HISTORY_SUMMARY_FILE
    existing_workbook_path = find_existing_customer_history_workbook([output_path, legacy_output_path])
    errors: list[str] = []
    daily_summaries: list[DailyCustomerSummary] = []

    detail_files = [
        detail_file
        for detail_file in sorted(customer_dir.iterdir())
        if is_daily_detail_file(detail_file)
    ]
    for detail_index, detail_file in enumerate(detail_files, start=1):
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
    workbook.calculation.calcMode = "auto"
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    history_sheet = workbook.active
    history_sheet.title = CUSTOMER_HISTORY_SHEET
    write_customer_history_sheet(history_sheet, daily_summaries)

    visible_headers, detail_rows, detail_errors = load_daily_detail_rows_for_history(
        customer_dir,
        existing_workbook_path,
    )
    errors.extend(detail_errors)
    detail_sheet = workbook.create_sheet(CUSTOMER_HISTORY_DETAIL_SHEET)
    write_customer_history_detail_sheet(detail_sheet, visible_headers, detail_rows)

    add_or_preserve_customer_payment_sheet(workbook, [output_path, legacy_output_path])

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
    progress_callback: ProgressCallback | None = None,
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
    total_customers = len(customer_dirs)
    emit_progress(progress_callback, f"里程碑：开始刷新客户历史汇总，共 {total_customers} 位客户")
    for customer_index, customer_dir in enumerate(customer_dirs, start=1):
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
        summaries.append(
            build_customer_history_summary(
                customer_dir,
                resolved_rule_config,
                progress_callback,
            )
        )
        if should_emit_progress(customer_index, total_customers, PROGRESS_GROUP_INTERVAL):
            emit_progress(
                progress_callback,
                f"进度：客户历史汇总已刷新 {customer_index}/{total_customers} 位客户",
            )
    emit_progress(progress_callback, f"里程碑：客户历史汇总刷新完成，共 {total_customers} 位客户")
    return summaries


def run_express_fee_job(
    config: ExpressFeeJobConfig,
    progress_callback: ProgressCallback | None = None,
    suppress_initial_stage_headers: bool = False,
) -> ExpressFeeJobResult:
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

    if not suppress_initial_stage_headers:
        emit_progress(progress_callback, "阶段 1/5：准备数据，正在读取报价")
    sales_context = collect_sales_price_context(sales_file)
    price_catalog = build_versioned_price_catalog(price_dir, sales_context, rule_config)
    price_map: dict[tuple[str, str, str], Price] = {}
    loaded_price_count = sum(len(price_map_by_file) for price_map_by_file in price_catalog.price_maps_by_file.values())
    emit_progress(progress_callback, f"里程碑：报价读取完成，共 {loaded_price_count} 条报价")
    available_standard_companies = build_available_standard_companies_from_catalog(
        price_catalog,
        rule_config,
    )
    if not suppress_initial_stage_headers:
        emit_progress(progress_callback, f"阶段 2/5：计算快递费，销售表：{sales_file.name}")
    processing_summary = process_sales_workbook(
        sales_file,
        price_map,
        available_standard_companies,
        output_path,
        config.round_digits,
        rule_config,
        progress_callback,
        price_catalog,
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
        emit_progress(progress_callback, "阶段 3/5：生成客户每日明细")
        split_summary = split_customer_daily_files(
            output_path,
            split_dir,
            rule_config,
            progress_callback,
        )
        logs.append("")
        logs.append("客户拆分：")
        logs.append(f"生成客户文件：{len(split_summary.generated_files)} 个")
        logs.append(f"跳过行数：{split_summary.skipped_rows} 条")
        for item in split_summary.generated_files:
            touched_customers.add(item.customer)
            split_files.append(item.output_path)
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
            emit_progress(progress_callback, "阶段 4/5：刷新客户历史汇总")
            history_summaries = refresh_customer_history_summaries(
                split_dir,
                refresh_customers,
                rule_config,
                progress_callback,
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
                        f"{item.row_count} 单，累计费用 {round(item.total_fee, 2)}"
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
    progress_callback: ProgressCallback | None = None,
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
    emit_progress(progress_callback, f"阶段 1/5：准备数据，共 {len(sales_files)} 个销售表")
    price_catalog: VersionedPriceCatalog | None = None
    try:
        sales_context = merge_sales_price_contexts(
            [collect_sales_price_context(sales_file) for sales_file in sales_files]
        )
        price_catalog = build_versioned_price_catalog(price_dir, sales_context, rule_config)
        loaded_price_count = sum(
            len(price_map_by_file)
            for price_map_by_file in price_catalog.price_maps_by_file.values()
        )
        emit_progress(progress_callback, f"里程碑：报价读取完成，共 {loaded_price_count} 条报价")
    except Exception as exc:
        message = str(exc)
        if not message.startswith("["):
            message = format_error_block(
                "报价表错误",
                str(exc),
                "请检查报价目录、业务员报价文件夹和报价版本文件。",
                stage="读取报价表",
                location=str(price_dir),
                system_error=exc.__class__.__name__,
            )
        job_result = ExpressFeeJobResult(
            sales_file=sales_files[0] if sales_files else Path(""),
            price_dir=price_dir,
            output_path=next(iter(output_paths.values()), output_dir / "未生成.xlsx"),
            split_dir=split_dir,
            processing_errors=[message],
            logs=[message],
        )
        return ExpressFeeBatchJobResult(
            sales_files=sales_files,
            price_dir=price_dir,
            output_dir=output_dir,
            split_dir=split_dir,
            job_results=[job_result],
            logs=logs + [message],
        )

    job_results: list[ExpressFeeJobResult] = []
    touched_customers: set[str] = set()
    emit_progress(progress_callback, "阶段 2/5：计算快递费")
    for index, sales_file in enumerate(sales_files, start=1):
        logs.extend(["", f"========== 第 {index}/{len(sales_files)} 个文件 =========="])
        emit_progress(
            progress_callback,
            f"进度：正在处理销售表 {index}/{len(sales_files)}：{sales_file.name}",
        )
        output_path = output_paths[sales_file]
        job_config = ExpressFeeJobConfig(
            sales_file=sales_file,
            price_dir=price_dir,
            output_path=output_path,
            split_dir=split_dir,
            round_digits=config.round_digits,
            split_customer_daily_files=False,
            generate_customer_history=False,
            refresh_all_customers=False,
            rule_config=rule_config,
        )

        try:
            output_path.parent.mkdir(parents=True, exist_ok=True)
            available_standard_companies = build_available_standard_companies_from_catalog(
                price_catalog,
                rule_config,
            )
            processing_summary = process_sales_workbook(
                sales_file,
                {},
                available_standard_companies,
                output_path,
                config.round_digits,
                rule_config,
                progress_callback,
                price_catalog,
            )
            result_logs = [
                f"销售表：{sales_file}",
                f"报价目录：{price_dir}",
                f"输出文件：{output_path}",
                "",
                f"共处理：{processing_summary.total_rows} 条",
                f"成功计算：{processing_summary.success_rows} 条",
                f"失败：{processing_summary.failed_rows} 条",
            ]
            if processing_summary.errors:
                result_logs.append("")
                result_logs.append("失败明细：")
                result_logs.extend(processing_summary.errors)
            result = ExpressFeeJobResult(
                sales_file=sales_file,
                price_dir=price_dir,
                output_path=output_path,
                split_dir=split_dir,
                total_rows=processing_summary.total_rows,
                success_rows=processing_summary.success_rows,
                failed_rows=processing_summary.failed_rows,
                processing_errors=processing_summary.errors or [],
                logs=result_logs,
            )
        except Exception as exc:  # Keep a batch moving if one workbook is bad.
            message = str(exc)
            if not message.startswith("["):
                message = format_error_block(
                    "未知错误",
                    str(exc),
                    "请复制完整运行日志给开发者排查。",
                    stage="处理销售表",
                    location=str(sales_file),
                    system_error=exc.__class__.__name__,
                )
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
            emit_progress(progress_callback, f"里程碑：销售表 {index}/{len(sales_files)} 处理异常，请查看最终问题明细")

        job_results.append(result)
        logs.extend(result.logs)
        emit_progress(
            progress_callback,
            f"里程碑：销售表 {index}/{len(sales_files)} 计算完成，"
            f"成功 {result.success_rows} 行，异常 {result.failed_rows} 行",
        )

    if config.split_customer_daily_files:
        emit_progress(progress_callback, "阶段 3/5：生成客户每日明细")
        logs.append("")
        logs.append("批量客户拆分：")
        for index, result in enumerate(job_results, start=1):
            if not result.output_path.exists():
                logs.append(f"{result.sales_file.name}：未生成客户每日明细")
                continue
            split_summary = split_customer_daily_files(
                result.output_path,
                split_dir,
                rule_config,
                progress_callback,
            )
            logs.append(
                f"{result.sales_file.name}：生成客户文件 {len(split_summary.generated_files)} 个，"
                f"跳过 {split_summary.skipped_rows} 行"
            )
            for item in split_summary.generated_files:
                touched_customers.add(item.customer)
                result.split_files.append(item.output_path)
            if split_summary.errors:
                result.split_errors.extend(split_summary.errors)
                logs.append(f"{result.sales_file.name} 拆分失败明细：")
                logs.extend(split_summary.errors)
        emit_progress(
            progress_callback,
            f"里程碑：批量客户每日明细生成完成，已处理 {len(job_results)} 个销售表",
        )

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
            emit_progress(progress_callback, "阶段 4/5：刷新客户历史汇总")
            history_summaries = refresh_customer_history_summaries(
                split_dir,
                refresh_customers,
                rule_config,
                progress_callback,
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
                        f"{item.row_count} 单，累计费用 {round(item.total_fee, 2)}"
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
    emit_progress(
        progress_callback,
        "阶段 5/5：完成，"
        f"销售表 {len(job_results)}/{len(sales_files)}，"
        f"成功 {sum(item.success_rows for item in job_results)} 行，"
        f"异常 {sum(item.failed_rows for item in job_results)} 行；"
        "生成文件请到「结果」页查看",
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
