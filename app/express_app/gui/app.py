"""Minimal Tkinter GUI for the express fee calculator."""

from __future__ import annotations

import queue
import re
import os
import subprocess
import sys
import threading
import tkinter as tk
import uuid
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from express_app.core.account_balance import (
    AccountBalanceDashboard,
    CustomerBalanceRecord,
    collect_account_balance_dashboard,
)
from express_app.core.bill_splitter import (
    BillSplitResult,
    BillSplitScanResult,
    BILL_DETAIL_SHEET_NAME,
    DEFAULT_SPLIT_FIELD,
    build_bill_split_output_dir,
    scan_bill_split_directory,
    split_bills_by_field,
)
from express_app.core.balance_upload import (
    BalanceUploadPreview,
    collect_balance_upload_preview,
    build_balance_upload_payload,
    upload_balance_payload,
)
from express_app.core.license_key import normalize_license_key, verify_license_key
from express_app.core.calculator import (
    DEFAULT_OUTPUT_DIR,
    DEFAULT_PRICE_DIR,
    DEFAULT_SPLIT_DIR,
    build_default_rule_config,
    load_price_template_workbook,
    scan_price_template_catalog,
    validate_express_fee_batch_job,
)
from express_app.core.models import (
    ExpressCompanyKeywordRule,
    ExpressFeeBatchJobConfig,
    ExpressFeeBatchJobResult,
    ExpressFeePreflightResult,
    ExpressFeeRuleConfig,
)
from express_app.core import run_express_fee_batch_job
from express_app.gui.config_store import GuiConfig, load_gui_config, save_gui_config
from express_app.gui.design_tokens import GUI_COLORS, GUI_FONTS, GUI_LAYOUT
from express_app.version import APP_DISPLAY_NAME, APP_VERSION, APP_VERSION_LABEL


APP_TITLE = f"{APP_DISPLAY_NAME} {APP_VERSION_LABEL}"
LICENSE_GATE_TITLE = "输入启动密钥"
LICENSE_GATE_PROMPT = "请输入 5 分钟内有效的启动密钥"
STARTUP_LOCK_OVERLAY_COLOR = "#E8EEF6"
OUTPUT_VERSION_LABEL = APP_VERSION_LABEL
V8_1_MAIN_NAV_ITEMS = (
    "费用计算",
    "账户余额",
    "报价预览",
    "系统设置",
)
V8_1_WORKFLOW_STEPS = ("配置", "运行", "结果")
V8_2_ENABLED_NAV_ITEMS = ("费用计算", "系统设置")
V8_3_ENABLED_NAV_ITEMS = ("费用计算", "账户余额", "系统设置")
V8_4_ENABLED_NAV_ITEMS = ("费用计算", "账户余额", "报价预览", "系统设置")
V8_10_MAIN_NAV_ITEMS = ("费用计算", "账户余额", "报价预览", "拆分账单", "余额上传", "系统设置")
V8_10_ENABLED_NAV_ITEMS = V8_10_MAIN_NAV_ITEMS
V8_9_MAIN_NAV_ITEMS = V8_10_MAIN_NAV_ITEMS
V8_9_ENABLED_NAV_ITEMS = V8_10_ENABLED_NAV_ITEMS
V8_10_SETTINGS_SECTIONS = ("目录配置", "精准映射", "关键词映射", "大件规则", "余额上传")
V8_2_SETTINGS_SECTIONS = V8_10_SETTINGS_SECTIONS
V8_2_1_WORKFLOW_STEPS = (
    ("config", "配置"),
    ("run", "运行"),
    ("results", "结果"),
)
V8_2_1_AUTO_WORKFLOW_TRANSITIONS = {"on_start": "run", "on_done": "results"}
V8_2_1_CONFIG_PAGE_SECTIONS = ("销售出库单", "当前系统设置", "生成选项")
V8_2_1_CONFIG_SYSTEM_DIRECTORY_LABELS = ("报价表目录", "总结果目录", "客户明细目录")
V8_2_1_CONFIG_GENERATION_OPTIONS = ("生成客户每日明细", "生成客户历史汇总")
V8_3_BALANCE_TABLE_COLUMNS = (
    "客户",
    "累计消费",
    "累计收款",
    "累计异常扣款",
    "当前余额",
    "最近日期",
    "状态",
    "文件路径",
)
V8_3_1_BALANCE_TOP_ACTIONS = ()
V8_3_1_BALANCE_FOOTER_ACTIONS = (
    ("刷新数据", "Primary.TButton"),
    ("打开客户目录", "Secondary.TButton"),
    ("打开历史汇总表", "Secondary.TButton"),
)
V8_9_3_BALANCE_METRIC_COLUMNS = 3
WORKBENCH_LABEL_FONT_SIZE = 12
SIDEBAR_BOTTOM_ACTIONS = ("系统设置", "打开客户目录")
V8_6_CONFIG_PAGE_ACTIONS = ("开始测试",)
V8_6_RUN_PAGE_ACTIONS = ("开始计算",)
V8_4_PRICE_TEMPLATE_ACTIONS = ("同步快递报价表", "业务员", "搜索", "打开快递价格表")
V8_4_PRICE_TEMPLATE_COLUMN_IDS = (
    "province_left",
    "first_price_left",
    "extra_price_left",
    "province_right",
    "first_price_right",
    "extra_price_right",
)
V8_4_PRICE_TEMPLATE_COLUMNS = ("省份", "首重费用", "续重费用", "省份", "首重费用", "续重费用")
V8_9_BILL_SPLIT_WORKFLOW_STEPS = (
    ("config", "配置"),
    ("run", "运行"),
    ("results", "结果"),
)
V8_9_BILL_SPLIT_CONFIG_PAGE_SECTIONS = ("账单目录", "待拆分文件")
V8_9_BILL_SPLIT_RUN_PAGE_SECTIONS = ("运行状态", "运行日志")
V8_9_BILL_SPLIT_RESULTS_PAGE_SECTIONS = ("拆分结果",)
V8_9_BILL_SPLIT_PAGE_SECTIONS = ("配置", "运行", "结果")
V8_9_BILL_SPLIT_ACTIONS = ("选择目录", "同步字段", "开始测试", "开始拆分", "打开输出目录", "打开选中文件")
V8_9_BILL_SPLIT_CONFIG_PAGE_ACTIONS = ("选择目录", "同步字段", "开始测试")
V8_9_BILL_SPLIT_RUN_PAGE_ACTIONS = ("开始拆分",)
V8_9_BILL_SPLIT_RESULTS_PAGE_ACTIONS = ("打开输出目录", "打开选中文件")
V8_9_BILL_SPLIT_AUTO_WORKFLOW_TRANSITIONS = {
    "on_test_start": "run",
    "on_test_done": "run",
    "on_split_start": "run",
    "on_split_done": "results",
}
V8_9_BILL_SPLIT_DEFAULT_FIELD = DEFAULT_SPLIT_FIELD
V8_9_BILL_SPLIT_TARGET_SHEET_NAME = BILL_DETAIL_SHEET_NAME
V8_9_BILL_SPLIT_OUTPUT_DIR_SUFFIX = "拆分结果"
V8_9_BILL_SPLIT_FILE_COLUMN_IDS = ("name", "sheet", "rows", "field_status", "path")
V8_9_BILL_SPLIT_FILE_COLUMNS = ("文件名", "账单明细", "数据行", "字段状态", "路径")
V8_9_BILL_SPLIT_RESULT_COLUMN_IDS = ("split_value", "name", "rows", "status", "path")
V8_9_BILL_SPLIT_RESULT_COLUMNS = ("拆分值", "文件名", "行数", "状态", "路径")
V8_10_BALANCE_UPLOAD_COLUMN_IDS = ("customer", "today_fee", "today_balance", "balance_date", "status")
V8_10_BALANCE_UPLOAD_COLUMNS = ("客户", "今日快递费消费", "今日余额", "余额日期", "状态")
V8_10_BALANCE_UPLOAD_ACTIONS = ("读取数据", "确认并上传")
SETTINGS_TOP_TAB_STYLE = "SettingsTop.TNotebook"
PRICE_PREVIEW_COMBO_STYLE = "PricePreview.TCombobox"
PRICE_PREVIEW_NOTEBOOK_STYLE = "PricePreview.TNotebook"
PRICE_PREVIEW_TAB_PADDING = (18, 10)
PRICE_PREVIEW_TAB_EXPAND = (0, 0, 0, 0)
PRICE_PREVIEW_TREE_STYLE = "PricePreview.Treeview"
PRICE_PREVIEW_SCROLLBAR_STYLE = "PricePreview.Vertical.TScrollbar"
PRICE_PREVIEW_TREE_COLUMN_WIDTHS = {
    "province_left": 120,
    "first_price_left": 120,
    "extra_price_left": 120,
    "province_right": 120,
    "first_price_right": 120,
    "extra_price_right": 120,
}
PRICE_PREVIEW_TREE_CELL_ANCHOR = "center"
PRICE_PREVIEW_GROUP_DIVIDER_WIDTH = 2
PRICE_PREVIEW_GROUP_DIVIDER_COLOR = GUI_COLORS["border2"]
OPTION_SELECTED_PREFIX = "✅"
OPTION_UNSELECTED_PREFIX = "□"
RULE_WINDOW_TITLE = "快递识别与大件规则"
EXPRESS_MAPPING_HELP_TEXT = (
    "把销售表里的原始快递名称，对应到报价表 sheet 名。格式：原始快递名称=标准快递名称，"
    "例如 顺丰速运新3=顺丰。"
)
KEYWORD_MAPPING_HELP_TEXT = (
    "用于兜底识别：当原始快递名称包含关键词时，自动识别为标准快递名称。"
    "格式：关键词=标准快递名称，例如 顺丰=顺丰。"
)
LARGE_RULE_HELP_TEXT = (
    "达到重量阈值后，系统会使用 标准快递名称+模板后缀 的报价 sheet，"
    "例如 顺丰_大件、德邦_大件。"
)

COLORS = {
    "background": GUI_COLORS["background"],
    "surface": GUI_COLORS["surface"],
    "surface_alt": GUI_COLORS["surface2"],
    "border": GUI_COLORS["border"],
    "border2": GUI_COLORS["border2"],
    "primary": GUI_COLORS["accent"],
    "primary_dark": GUI_COLORS["accent"],
    "accent": GUI_COLORS["warning"],
    "accent_light": GUI_COLORS["accent_light"],
    "accent_bg": GUI_COLORS["accent_bg"],
    "accent_border": GUI_COLORS["accent_border"],
    "success": GUI_COLORS["success"],
    "success_bg": GUI_COLORS["success_bg"],
    "warning": GUI_COLORS["warning"],
    "warning_bg": GUI_COLORS["warning_bg"],
    "danger": GUI_COLORS["danger"],
    "danger_bg": GUI_COLORS["danger_bg"],
    "text": GUI_COLORS["text"],
    "muted": GUI_COLORS["muted"],
    "dim": GUI_COLORS["dim"],
    "sidebar": GUI_COLORS["sidebar"],
    "log_bg": GUI_COLORS["log_bg"],
    "log_text": GUI_COLORS["log_text"],
}


def format_option_label(label: str, selected: bool) -> str:
    prefix = OPTION_SELECTED_PREFIX if selected else OPTION_UNSELECTED_PREFIX
    return f"{prefix} {label}"


def parse_company_list(text: str) -> set[str]:
    normalized = text.replace("，", ",").replace("、", ",").replace(";", ",")
    return {item.strip() for item in normalized.split(",") if item.strip()}


def iter_mapping_lines(text: str):
    for line_number, line in enumerate(text.splitlines(), start=1):
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        yield line_number, line


def parse_mapping_line(line: str, source_name: str, line_number: int) -> tuple[str, str]:
    separator = "=>" if "=>" in line else "="
    if separator not in line:
        raise ValueError(f"{source_name} 第 {line_number} 行缺少 =。")
    raw_name, standard_name = [part.strip() for part in line.split(separator, 1)]
    if not raw_name or not standard_name:
        raise ValueError(f"{source_name} 第 {line_number} 行不能有空值。")
    return raw_name, standard_name


def parse_mapping_text(text: str, source_name: str) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for line_number, line in iter_mapping_lines(text):
        raw_name, standard_name = parse_mapping_line(line, source_name, line_number)
        if raw_name in mapping:
            raise ValueError(f"{source_name} 第 {line_number} 行重复：{raw_name}")
        mapping[raw_name] = standard_name
    if not mapping:
        raise ValueError(f"{source_name}不能为空。")
    return mapping


def parse_ordered_mapping_text(text: str, source_name: str) -> list[tuple[str, str]]:
    rules: list[tuple[str, str]] = []
    seen_keywords: set[str] = set()
    for line_number, line in iter_mapping_lines(text):
        keyword, standard_name = parse_mapping_line(line, source_name, line_number)
        if keyword in seen_keywords:
            raise ValueError(f"{source_name} 第 {line_number} 行重复：{keyword}")
        seen_keywords.add(keyword)
        rules.append((keyword, standard_name))
    if not rules:
        raise ValueError(f"{source_name}不能为空。")
    return rules


def format_exact_mapping_text(rule_config: ExpressFeeRuleConfig) -> str:
    return "\n".join(
        f"{raw_name}={standard_name}"
        for raw_name, standard_name in sorted(rule_config.exact_company_map.items())
    )


def format_keyword_mapping_text(rule_config: ExpressFeeRuleConfig) -> str:
    return "\n".join(
        f"{item.keyword}={item.standard_name}" for item in rule_config.keyword_company_rules
    )


def format_large_piece_companies(rule_config: ExpressFeeRuleConfig) -> str:
    return "、".join(sorted(rule_config.large_piece_companies))


def format_super_large_piece_companies(rule_config: ExpressFeeRuleConfig) -> str:
    return "、".join(sorted(rule_config.super_large_piece_companies))


def build_rule_config_from_text_fields(
    exact_mapping_text: str,
    keyword_mapping_text: str,
    large_companies_text: str,
    threshold_text: str,
    suffix_text: str,
    super_large_companies_text: str,
    super_large_threshold_text: str,
    super_large_suffix_text: str,
) -> ExpressFeeRuleConfig:
    try:
        threshold = float(threshold_text.strip())
    except ValueError as exc:
        raise ValueError("重量阈值必须是数字。") from exc
    if threshold <= 0:
        raise ValueError("重量阈值必须大于 0。")

    suffix = suffix_text.strip()
    if not suffix:
        raise ValueError("模板后缀不能为空。")

    try:
        super_large_threshold = float(super_large_threshold_text.strip())
    except ValueError as exc:
        raise ValueError("超大件重量阈值必须是数字。") from exc
    if super_large_threshold <= 0:
        raise ValueError("超大件重量阈值必须大于 0。")
    if super_large_threshold <= threshold:
        raise ValueError("超大件重量阈值必须大于大件重量阈值。")

    super_large_suffix = super_large_suffix_text.strip()
    if not super_large_suffix:
        raise ValueError("超大件模板后缀不能为空。")

    large_piece_companies = parse_company_list(large_companies_text)
    if not large_piece_companies:
        raise ValueError("大件快递至少需要填写一个标准快递公司。")

    super_large_piece_companies = parse_company_list(super_large_companies_text)
    if not super_large_piece_companies:
        raise ValueError("超大件快递至少需要填写一个标准快递公司。")

    return ExpressFeeRuleConfig(
        exact_company_map=parse_mapping_text(exact_mapping_text, "快递公司精准映射"),
        keyword_company_rules=[
            ExpressCompanyKeywordRule(keyword=keyword, standard_name=standard)
            for keyword, standard in parse_ordered_mapping_text(
                keyword_mapping_text,
                "快递公司关键词映射",
            )
        ],
        large_piece_companies=large_piece_companies,
        large_piece_threshold_kg=threshold,
        large_piece_suffix=suffix,
        super_large_piece_companies=super_large_piece_companies,
        super_large_piece_threshold_kg=super_large_threshold,
        super_large_piece_suffix=super_large_suffix,
    )


class RuleConfigWindow(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Tk,
        rule_config: ExpressFeeRuleConfig,
        on_save,
    ) -> None:
        super().__init__(parent)
        self.title(RULE_WINDOW_TITLE)
        self.geometry("840x720")
        self.minsize(760, 640)
        self.transient(parent)
        self.grab_set()

        self.on_save = on_save
        self.large_companies_var = tk.StringVar(
            value="、".join(sorted(rule_config.large_piece_companies))
        )
        self.threshold_var = tk.StringVar(value=str(rule_config.large_piece_threshold_kg))
        self.suffix_var = tk.StringVar(value=rule_config.large_piece_suffix)
        self.super_large_companies_var = tk.StringVar(
            value="、".join(sorted(rule_config.super_large_piece_companies))
        )
        self.super_large_threshold_var = tk.StringVar(
            value=str(rule_config.super_large_piece_threshold_kg)
        )
        self.super_large_suffix_var = tk.StringVar(value=rule_config.super_large_piece_suffix)

        self._build_ui()
        self._load_rule_config(rule_config)
        self.focus()

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=14)
        root.pack(fill=tk.BOTH, expand=True)
        root.columnconfigure(0, weight=1)

        title = ttk.Label(
            root,
            text="核心识别规则：销售表快递名称 → 报价表 sheet",
            font=("Helvetica Neue", 13, "bold"),
        )
        title.grid(row=0, column=0, sticky="w", pady=(0, 6))
        intro = ttk.Label(
            root,
            text="优先维护这里。系统先按精准映射识别，再用关键词映射兜底，识别结果必须能对应报价表里的 sheet 名。",
            wraplength=780,
            foreground=COLORS["muted"],
        )
        intro.grid(row=1, column=0, sticky="ew", pady=(0, 10))

        exact_frame = ttk.LabelFrame(root, text="1. 精准映射（优先匹配）", padding=10)
        exact_frame.grid(row=2, column=0, sticky="nsew", pady=(0, 10))
        exact_frame.rowconfigure(0, weight=1)
        exact_frame.columnconfigure(0, weight=1)
        ttk.Label(
            exact_frame,
            text=EXPRESS_MAPPING_HELP_TEXT,
            wraplength=760,
            foreground=COLORS["muted"],
        ).grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        self.exact_text = tk.Text(exact_frame, height=8, wrap=tk.NONE)
        self.exact_text.grid(row=1, column=0, sticky="nsew")
        exact_scroll = ttk.Scrollbar(
            exact_frame,
            orient=tk.VERTICAL,
            command=self.exact_text.yview,
        )
        exact_scroll.grid(row=1, column=1, sticky="ns")
        self.exact_text.configure(yscrollcommand=exact_scroll.set)

        keyword_frame = ttk.LabelFrame(root, text="2. 关键词映射（兜底识别）", padding=10)
        keyword_frame.grid(row=3, column=0, sticky="nsew", pady=(0, 10))
        keyword_frame.rowconfigure(0, weight=1)
        keyword_frame.columnconfigure(0, weight=1)
        ttk.Label(
            keyword_frame,
            text=KEYWORD_MAPPING_HELP_TEXT,
            wraplength=760,
            foreground=COLORS["muted"],
        ).grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        self.keyword_text = tk.Text(keyword_frame, height=7, wrap=tk.NONE)
        self.keyword_text.grid(row=1, column=0, sticky="nsew")
        keyword_scroll = ttk.Scrollbar(
            keyword_frame,
            orient=tk.VERTICAL,
            command=self.keyword_text.yview,
        )
        keyword_scroll.grid(row=1, column=1, sticky="ns")
        self.keyword_text.configure(yscrollcommand=keyword_scroll.set)

        large_frame = ttk.LabelFrame(root, text="高级计费规则：大件 / 超大件模板", padding=10)
        large_frame.grid(row=4, column=0, sticky="ew", pady=(0, 10))
        large_frame.columnconfigure(1, weight=1)
        ttk.Label(
            large_frame,
            text=LARGE_RULE_HELP_TEXT,
            wraplength=760,
            foreground=COLORS["muted"],
        ).grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        ttk.Label(large_frame, text="大件快递").grid(row=1, column=0, sticky="w")
        ttk.Entry(large_frame, textvariable=self.large_companies_var).grid(
            row=1,
            column=1,
            sticky="ew",
            padx=(10, 0),
        )
        ttk.Label(large_frame, text="大件首重重量").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(large_frame, textvariable=self.threshold_var, width=12).grid(
            row=2,
            column=1,
            sticky="w",
            padx=(10, 0),
            pady=(8, 0),
        )
        ttk.Label(large_frame, text="大件模板后缀").grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(large_frame, textvariable=self.suffix_var, width=18).grid(
            row=3,
            column=1,
            sticky="w",
            padx=(10, 0),
            pady=(8, 0),
        )
        ttk.Label(large_frame, text="超大件快递").grid(row=4, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(large_frame, textvariable=self.super_large_companies_var).grid(
            row=4,
            column=1,
            sticky="ew",
            padx=(10, 0),
            pady=(8, 0),
        )
        ttk.Label(large_frame, text="超大件首重重量").grid(row=5, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(large_frame, textvariable=self.super_large_threshold_var, width=12).grid(
            row=5,
            column=1,
            sticky="w",
            padx=(10, 0),
            pady=(8, 0),
        )
        ttk.Label(large_frame, text="超大件模板后缀").grid(row=6, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(large_frame, textvariable=self.super_large_suffix_var, width=18).grid(
            row=6,
            column=1,
            sticky="w",
            padx=(10, 0),
            pady=(8, 0),
        )

        actions = ttk.Frame(root)
        actions.grid(row=5, column=0, sticky="ew")
        ttk.Button(actions, text="恢复默认规则", command=self._restore_defaults).pack(
            side=tk.LEFT
        )
        ttk.Button(actions, text="取消", command=self.destroy).pack(side=tk.RIGHT)
        ttk.Button(actions, text="保存", command=self._save).pack(
            side=tk.RIGHT,
            padx=(0, 10),
        )

    def _load_rule_config(self, rule_config: ExpressFeeRuleConfig) -> None:
        self.large_companies_var.set(format_large_piece_companies(rule_config))
        self.threshold_var.set(str(rule_config.large_piece_threshold_kg))
        self.suffix_var.set(rule_config.large_piece_suffix)
        self.super_large_companies_var.set(format_super_large_piece_companies(rule_config))
        self.super_large_threshold_var.set(str(rule_config.super_large_piece_threshold_kg))
        self.super_large_suffix_var.set(rule_config.super_large_piece_suffix)
        self.exact_text.delete("1.0", tk.END)
        self.exact_text.insert(tk.END, format_exact_mapping_text(rule_config))
        self.keyword_text.delete("1.0", tk.END)
        self.keyword_text.insert(tk.END, format_keyword_mapping_text(rule_config))

    def _restore_defaults(self) -> None:
        self._load_rule_config(build_default_rule_config())

    def _save(self) -> None:
        try:
            rule_config = self._build_rule_config()
        except ValueError as exc:
            messagebox.showerror("规则错误", str(exc), parent=self)
            return

        self.on_save(rule_config)
        messagebox.showinfo("已保存", "规则配置已保存。", parent=self)
        self.destroy()

    def _build_rule_config(self) -> ExpressFeeRuleConfig:
        return build_rule_config_from_text_fields(
            exact_mapping_text=self.exact_text.get("1.0", tk.END),
            keyword_mapping_text=self.keyword_text.get("1.0", tk.END),
            large_companies_text=self.large_companies_var.get(),
            threshold_text=self.threshold_var.get(),
            suffix_text=self.suffix_var.get(),
            super_large_companies_text=self.super_large_companies_var.get(),
            super_large_threshold_text=self.super_large_threshold_var.get(),
            super_large_suffix_text=self.super_large_suffix_var.get(),
        )

    def _parse_company_list(self, text: str) -> set[str]:
        return parse_company_list(text)

    def _parse_mapping_line(self, line: str, source_name: str, line_number: int) -> tuple[str, str]:
        return parse_mapping_line(line, source_name, line_number)

    def _iter_mapping_lines(self, text: str):
        return iter_mapping_lines(text)

    def _parse_mapping_text(self, text: str, source_name: str) -> dict[str, str]:
        return parse_mapping_text(text, source_name)

    def _parse_ordered_mapping_text(self, text: str, source_name: str) -> list[tuple[str, str]]:
        return parse_ordered_mapping_text(text, source_name)


class LicenseKeyDialog(tk.Toplevel):
    def __init__(self, parent: tk.Tk) -> None:
        super().__init__(parent)
        self.title(LICENSE_GATE_TITLE)
        self.resizable(False, False)
        self.configure(bg=COLORS["surface"])
        self.result = False

        self.transient(parent)
        self.grab_set()

        root = ttk.Frame(self, style="Surface.TFrame", padding=(24, 22, 24, 18))
        root.grid(row=0, column=0, sticky="nsew")
        root.columnconfigure(0, weight=1)

        ttk.Label(
            root,
            text=APP_DISPLAY_NAME,
            style="CardTitle.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            root,
            text=LICENSE_GATE_PROMPT,
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(6, 14))

        self.key_var = tk.StringVar()
        self.entry = ttk.Entry(root, textvariable=self.key_var, width=24, justify="center")
        self.entry.grid(row=2, column=0, sticky="ew")
        self.entry.bind("<Return>", lambda _event: self._verify())

        self.error_var = tk.StringVar(value="")
        ttk.Label(
            root,
            textvariable=self.error_var,
            style="Error.TLabel",
        ).grid(row=3, column=0, sticky="w", pady=(8, 0))

        button_row = ttk.Frame(root, style="Surface.TFrame")
        button_row.grid(row=4, column=0, sticky="e", pady=(18, 0))
        ttk.Button(
            button_row,
            text="取消",
            style="Secondary.TButton",
            command=self._cancel,
        ).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(
            button_row,
            text="验证",
            style="Primary.TButton",
            command=self._verify,
        ).grid(row=0, column=1)

        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.after(50, self.entry.focus_set)
        self.update_idletasks()
        self.lift()
        self.focus_force()
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        width = self.winfo_width()
        height = self.winfo_height()
        x = parent_x + max((parent_width - width) // 2, 0)
        y = parent_y + max((parent_height - height) // 2, 0)
        self.geometry(f"+{x}+{y}")

    def _verify(self) -> None:
        key = normalize_license_key(self.key_var.get())
        if not key:
            self.error_var.set("密钥格式不正确，应为 XXXX-XXXX。")
            return
        if not verify_license_key(key):
            self.error_var.set("密钥无效或已过期，请重新复制最新密钥。")
            return
        self.result = True
        self.destroy()

    def _cancel(self) -> None:
        self.result = False
        self.destroy()


class StartupLockOverlay(tk.Frame):
    def __init__(self, parent: tk.Tk) -> None:
        super().__init__(parent, bg=STARTUP_LOCK_OVERLAY_COLOR, highlightthickness=0)
        self.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.lift()

        panel = tk.Frame(self, bg=STARTUP_LOCK_OVERLAY_COLOR)
        panel.place(relx=0.5, rely=0.46, anchor="center")
        tk.Label(
            panel,
            text=APP_DISPLAY_NAME,
            bg=STARTUP_LOCK_OVERLAY_COLOR,
            fg=COLORS["primary_dark"],
            font=("Helvetica Neue", 22, "bold"),
        ).pack()
        tk.Label(
            panel,
            text="主界面已锁定，请输入启动密钥",
            bg=STARTUP_LOCK_OVERLAY_COLOR,
            fg=COLORS["muted"],
            font=("Helvetica Neue", 12),
        ).pack(pady=(10, 0))


def should_skip_license_gate(environ: dict[str, str] | None = None) -> bool:
    env = os.environ if environ is None else environ
    return env.get("EXPRESS_APP_SELF_CHECK") == "1" or env.get("EXPRESS_APP_DISABLE_LICENSE_GATE") == "1"


def request_startup_license(parent: tk.Tk) -> bool:
    dialog = LicenseKeyDialog(parent)
    parent.wait_window(dialog)
    return dialog.result


def create_startup_lock_overlay(app: tk.Tk, overlay_factory=StartupLockOverlay):
    overlay = overlay_factory(app)
    setattr(app, "startup_lock_overlay", overlay)
    return overlay


def remove_startup_lock_overlay(app: tk.Tk) -> None:
    overlay = getattr(app, "startup_lock_overlay", None)
    if overlay is not None:
        overlay.destroy()
    setattr(app, "startup_lock_overlay", None)


def run_startup_license_gate(
    app: tk.Tk,
    *,
    environ: dict[str, str] | None = None,
    create_overlay=create_startup_lock_overlay,
    remove_overlay=None,
    request_license=request_startup_license,
) -> bool:
    if should_skip_license_gate(environ):
        return True
    app.update_idletasks()
    overlay = create_overlay(app)
    if remove_overlay is None:
        remove_overlay = lambda _overlay: remove_startup_lock_overlay(app)
    if request_license(app):
        remove_overlay(overlay)
        return True
    remove_overlay(overlay)
    app.destroy()
    return False


class ExpressFeeApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry(f"{GUI_LAYOUT['app_width']}x{GUI_LAYOUT['app_height']}")
        self.minsize(GUI_LAYOUT["app_width"], GUI_LAYOUT["app_height"])
        self.configure(bg=COLORS["background"])

        self._queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self._worker: threading.Thread | None = None
        self._last_result: ExpressFeeBatchJobResult | None = None
        self._last_preflight_result: ExpressFeePreflightResult | None = None
        self._last_preflight_config_signature: tuple[object, ...] | None = None
        self.result_paths: dict[str, Path] = {}

        gui_config = load_gui_config()
        self.sales_files: list[Path] = list(gui_config.sales_files)
        self.sales_dir = gui_config.sales_dir
        self.sales_file_var = tk.StringVar(value=self._format_sales_files_display())
        self.price_dir_var = tk.StringVar(value=str(gui_config.price_dir))
        self.output_dir_var = tk.StringVar(value=str(gui_config.output_dir))
        self.split_dir_var = tk.StringVar(value=str(gui_config.split_dir))
        self.split_var = tk.BooleanVar(value=gui_config.split_customer_daily_files)
        self.history_var = tk.BooleanVar(value=gui_config.generate_customer_history)
        self.refresh_all_var = tk.BooleanVar(value=False)
        self.split_option_label_var = tk.StringVar()
        self.history_option_label_var = tk.StringVar()
        self.refresh_all_option_label_var = tk.StringVar()
        self.rule_config = gui_config.rule_config or build_default_rule_config()
        self.status_var = tk.StringVar(value="就绪")
        self.summary_sales_files_var = tk.StringVar(value="0")
        self.summary_success_var = tk.StringVar(value="0")
        self.summary_failed_var = tk.StringVar(value="0")
        self.summary_outputs_var = tk.StringVar(value="0")
        self.active_nav_var = tk.StringVar(value="费用计算")
        self.active_workflow_var = tk.StringVar(value="config")
        self.module_title_var = tk.StringVar(value="费用计算")
        self.module_subtitle_var = tk.StringVar(
            value="按配置、运行、结果三个步骤完成快递费用计算和客户明细生成。"
        )
        self.nav_labels: dict[str, ttk.Label] = {}
        self.workflow_step_labels: dict[str, ttk.Label] = {}
        self.workflow_pages: dict[str, ttk.Frame] = {}
        self.active_bill_split_workflow_var = tk.StringVar(value="config")
        self.bill_split_workflow_step_labels: dict[str, ttk.Label] = {}
        self.bill_split_workflow_pages: dict[str, ttk.Frame] = {}
        self.preflight_buttons: list[ttk.Button] = []
        self.run_buttons: list[ttk.Button] = []
        self.content_container: ttk.Frame | None = None
        self.fee_page: ttk.Frame | None = None
        self.balance_page: ttk.Frame | None = None
        self.price_preview_page: ttk.Frame | None = None
        self.bill_splitter_page: ttk.Frame | None = None
        self.settings_page: ttk.Frame | None = None
        self.balance_tree: ttk.Treeview | None = None
        self.balance_status_var = tk.StringVar(value="等待刷新")
        self.balance_total_consumed_var = tk.StringVar(value="¥0.00")
        self.balance_total_paid_var = tk.StringVar(value="¥0.00")
        self.balance_total_abnormal_deducted_var = tk.StringVar(value="¥0.00")
        self.balance_available_balance_var = tk.StringVar(value="¥0.00")
        self.balance_debt_total_var = tk.StringVar(value="¥0.00")
        self.balance_debtor_count_var = tk.StringVar(value="0 位")
        self.balance_records: list[CustomerBalanceRecord] = []
        self.balance_history_paths: dict[str, Path] = {}
        self.balance_customer_dirs: dict[str, Path] = {}
        self.price_template_customer_var = tk.StringVar()
        self.price_template_status_var = tk.StringVar(value="点击同步快递报价表后选择业务员。")
        self.price_template_summary_var = tk.StringVar(value="选择业务员后，客户报价会按快递公司分标签展示。")
        self.price_template_catalog = None
        self.price_template_workbook = None
        self.price_template_current_file: Path | None = None
        self.price_template_customers: list[str] = []
        self.price_template_selected_sheet = ""
        self.price_template_sheet_tabs: list[str] = []
        self.price_template_rows_by_sheet: dict[str, list[tuple[str, ...]]] = {}
        self.price_template_record_counts_by_sheet: dict[str, int] = {}
        self.price_template_combo: ttk.Combobox | None = None
        self.price_template_notebook: ttk.Notebook | None = None
        self.price_template_trees: dict[str, ttk.Treeview] = {}
        self.bill_split_source_dir_var = tk.StringVar()
        self.bill_split_output_dir_var = tk.StringVar()
        self.bill_split_field_var = tk.StringVar(value=V8_9_BILL_SPLIT_DEFAULT_FIELD)
        self.bill_split_status_var = tk.StringVar(value="等待选择目录")
        self.bill_split_summary_files_var = tk.StringVar(value="0")
        self.bill_split_summary_success_var = tk.StringVar(value="0")
        self.bill_split_summary_failed_var = tk.StringVar(value="0")
        self.bill_split_summary_outputs_var = tk.StringVar(value="0")
        self.bill_split_common_headers: list[str] = []
        self.bill_split_scan_result: BillSplitScanResult | None = None
        self.bill_split_result_paths: dict[str, Path] = {}
        self.bill_split_buttons: list[ttk.Button] = []
        self.bill_split_config_buttons: list[ttk.Button] = []
        self.bill_split_run_buttons: list[ttk.Button] = []
        self.bill_split_sync_button: ttk.Button | None = None
        self.bill_split_test_button: ttk.Button | None = None
        self.bill_split_run_button: ttk.Button | None = None
        self.bill_split_open_output_button: ttk.Button | None = None
        self.bill_split_open_selected_button: ttk.Button | None = None
        self.bill_split_field_combo: ttk.Combobox | None = None
        self.bill_split_file_tree: ttk.Treeview | None = None
        self.bill_split_result_file_tree: ttk.Treeview | None = None
        self.bill_split_result_tree: ttk.Treeview | None = None
        self.bill_split_log_text: scrolledtext.ScrolledText | None = None
        self.active_settings_section_var = tk.StringVar(value=V8_2_SETTINGS_SECTIONS[0])
        self.settings_section_notebook: ttk.Notebook | None = None
        self.settings_section_pages: dict[str, ttk.Frame] = {}
        self.settings_exact_text: tk.Text | None = None
        self.settings_keyword_text: tk.Text | None = None
        self.balance_upload_page: ttk.Frame | None = None
        self.balance_upload_tree: ttk.Treeview | None = None
        self.balance_upload_log_text: scrolledtext.ScrolledText | None = None
        self.balance_upload_preview: BalanceUploadPreview | None = None
        self.balance_upload_uploading = False
        self.balance_upload_date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.balance_upload_confirm_var = tk.BooleanVar(value=False)
        self.balance_upload_status_var = tk.StringVar(value="等待读取数据")
        self.balance_upload_customer_count_var = tk.StringVar(value="0 位")
        self.balance_upload_today_fee_var = tk.StringVar(value="¥0.00")
        self.balance_upload_today_balance_var = tk.StringVar(value="¥0.00")
        self.balance_upload_debtor_count_var = tk.StringVar(value="0 位")
        self.balance_upload_button: ttk.Button | None = None
        self.balance_upload_read_button: ttk.Button | None = None
        self.settings_large_companies_var = tk.StringVar()
        self.settings_threshold_var = tk.StringVar()
        self.settings_suffix_var = tk.StringVar()
        self.settings_super_large_companies_var = tk.StringVar()
        self.settings_super_large_threshold_var = tk.StringVar()
        self.settings_super_large_suffix_var = tk.StringVar()
        self.settings_balance_upload_url_var = tk.StringVar(value=gui_config.balance_upload_url)
        self.settings_balance_upload_token_var = tk.StringVar(value=gui_config.balance_upload_token)

        self._configure_styles()
        self._refresh_option_labels()
        self._sync_rule_settings_vars()
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        font_family = "Helvetica Neue"
        mono_family = "Menlo"
        style.configure("App.TFrame", background=COLORS["background"])
        style.configure("Surface.TFrame", background=COLORS["surface"])
        style.configure("Shell.TFrame", background=COLORS["background"])
        style.configure("Sidebar.TFrame", background=COLORS["sidebar"])
        style.configure("Content.TFrame", background=COLORS["background"])
        style.configure("Header.TFrame", background=COLORS["surface"])
        style.configure(
            "Title.TLabel",
            background=COLORS["surface"],
            foreground=COLORS["text"],
            font=(font_family, 18, "bold"),
        )
        style.configure(
            "HeaderMeta.TLabel",
            background=COLORS["surface"],
            foreground=COLORS["muted"],
            font=(mono_family, 10),
        )
        style.configure(
            "Muted.TLabel",
            background=COLORS["surface"],
            foreground=COLORS["muted"],
            font=(font_family, 11),
        )
        style.configure(
            "Error.TLabel",
            background=COLORS["surface"],
            foreground=COLORS["danger"],
            font=(font_family, 10),
        )
        style.configure(
            "Status.TLabel",
            background=COLORS["success_bg"],
            foreground=COLORS["success"],
            padding=(10, 4),
            font=(font_family, 11, "bold"),
        )
        style.configure(
            "ModuleTitle.TLabel",
            background=COLORS["background"],
            foreground=COLORS["text"],
            font=(font_family, 16, "bold"),
        )
        style.configure(
            "ModuleSub.TLabel",
            background=COLORS["background"],
            foreground=COLORS["muted"],
            font=(font_family, 11),
        )
        style.configure(
            "NavActive.TLabel",
            background=COLORS["accent_bg"],
            foreground=COLORS["primary"],
            padding=(12, 10),
            font=(font_family, 12, "bold"),
        )
        style.configure(
            "NavDisabled.TLabel",
            background=COLORS["sidebar"],
            foreground=COLORS["muted"],
            padding=(12, 10),
            font=(font_family, 12),
        )
        style.configure(
            "StepActive.TLabel",
            background=COLORS["accent_bg"],
            foreground=COLORS["primary"],
            padding=(10, 6),
            font=(font_family, 11, "bold"),
        )
        style.configure(
            "StepIdle.TLabel",
            background=COLORS["surface_alt"],
            foreground=COLORS["muted"],
            padding=(10, 6),
            font=(font_family, 11),
        )
        style.configure(
            "Panel.TLabelframe",
            background=COLORS["surface"],
            bordercolor=COLORS["border"],
            relief=tk.SOLID,
        )
        style.configure(
            "Panel.TLabelframe.Label",
            background=COLORS["surface"],
            foreground=COLORS["primary_dark"],
            font=(font_family, 12, "bold"),
        )
        style.configure(
            "Field.TLabel",
            background=COLORS["surface"],
            foreground=COLORS["text"],
            font=(font_family, 11),
        )
        style.configure(
            "Path.TEntry",
            fieldbackground=COLORS["surface_alt"],
            foreground=COLORS["text"],
            insertcolor=COLORS["text"],
            bordercolor=COLORS["border"],
            font=(mono_family, 11),
        )
        style.configure(
            "MetricLabel.TLabel",
            background=COLORS["surface"],
            foreground=COLORS["muted"],
            font=(font_family, 10),
            anchor="center",
        )
        style.configure(
            "MetricValue.TLabel",
            background=COLORS["surface"],
            foreground=COLORS["primary_dark"],
            font=(mono_family, 16, "bold"),
            anchor="center",
        )
        style.configure(
            "Primary.TButton",
            background=COLORS["primary"],
            foreground="#FFFFFF",
            font=(font_family, 11, "bold"),
            padding=(14, 8),
        )
        style.map(
            "Primary.TButton",
            background=[("active", COLORS["accent_light"]), ("disabled", COLORS["dim"])],
            foreground=[("disabled", "#E2E8F0")],
        )
        style.configure(
            "Secondary.TButton",
            background=COLORS["surface"],
            foreground=COLORS["muted"],
            font=(font_family, 11),
            padding=(10, 7),
            bordercolor=COLORS["border"],
        )
        style.map("Secondary.TButton", background=[("active", COLORS["surface_alt"])])
        style.configure(
            "App.TCheckbutton",
            background=COLORS["surface"],
            foreground=COLORS["text"],
            font=(font_family, 11),
            indicatoron=False,
            padding=(10, 7),
        )
        style.map(
            "App.TCheckbutton",
            background=[("selected", COLORS["accent_bg"]), ("active", COLORS["surface_alt"])],
            foreground=[("selected", COLORS["primary_dark"])],
        )
        style.configure(
            "Treeview",
            background=COLORS["surface"],
            fieldbackground=COLORS["surface"],
            foreground=COLORS["text"],
            rowheight=28,
            bordercolor=COLORS["border"],
            font=(font_family, 10),
        )
        style.configure(
            "Treeview.Heading",
            background=COLORS["surface_alt"],
            foreground=COLORS["primary_dark"],
            font=(font_family, 10, "bold"),
            relief=tk.FLAT,
        )
        style.map(
            "Treeview",
            background=[("selected", COLORS["primary"])],
            foreground=[("selected", "#FFFFFF")],
        )
        style.configure(
            SETTINGS_TOP_TAB_STYLE,
            background=COLORS["background"],
            borderwidth=0,
            tabmargins=(0, 0, 0, 10),
        )
        style.configure(
            f"{SETTINGS_TOP_TAB_STYLE}.Tab",
            background=COLORS["surface_alt"],
            foreground=COLORS["muted"],
            padding=(18, 9),
            font=(font_family, 11, "bold"),
        )
        style.map(
            f"{SETTINGS_TOP_TAB_STYLE}.Tab",
            background=[
                ("selected", COLORS["primary"]),
                ("active", COLORS["accent_bg"]),
            ],
            foreground=[
                ("selected", "#FFFFFF"),
                ("active", COLORS["primary_dark"]),
            ],
            expand=[("selected", (0, 0, 0, 0))],
        )
        style.configure(
            PRICE_PREVIEW_COMBO_STYLE,
            fieldbackground="#FFFFFF",
            background=COLORS["surface_alt"],
            foreground=COLORS["text"],
            arrowcolor=COLORS["primary"],
            bordercolor=COLORS["accent_light"],
            lightcolor=COLORS["accent_light"],
            darkcolor=COLORS["border"],
            padding=(10, 7),
            font=(font_family, 11),
        )
        style.map(
            PRICE_PREVIEW_COMBO_STYLE,
            fieldbackground=[("readonly", "#FFFFFF"), ("focus", "#FFFFFF")],
            selectbackground=[("readonly", COLORS["accent_bg"])],
            selectforeground=[("readonly", COLORS["primary_dark"])],
            bordercolor=[("focus", COLORS["primary"]), ("hover", COLORS["accent_light"])],
        )
        style.configure(
            PRICE_PREVIEW_NOTEBOOK_STYLE,
            background=COLORS["surface"],
            borderwidth=0,
            tabmargins=(0, 0, 0, 8),
        )
        style.configure(
            f"{PRICE_PREVIEW_NOTEBOOK_STYLE}.Tab",
            background=COLORS["surface_alt"],
            foreground=COLORS["muted"],
            padding=PRICE_PREVIEW_TAB_PADDING,
            font=(font_family, 11, "bold"),
        )
        style.map(
            f"{PRICE_PREVIEW_NOTEBOOK_STYLE}.Tab",
            background=[
                ("selected", COLORS["primary"]),
                ("active", COLORS["accent_bg"]),
            ],
            foreground=[
                ("selected", "#FFFFFF"),
                ("active", COLORS["primary_dark"]),
            ],
            expand=[("selected", PRICE_PREVIEW_TAB_EXPAND)],
        )
        style.configure(
            PRICE_PREVIEW_TREE_STYLE,
            background="#FFFFFF",
            fieldbackground="#FFFFFF",
            foreground=COLORS["text"],
            rowheight=34,
            bordercolor=COLORS["border"],
            font=(font_family, 11),
        )
        style.configure(
            "PricePreviewDivider.TFrame",
            background=PRICE_PREVIEW_GROUP_DIVIDER_COLOR,
        )
        style.configure(
            f"{PRICE_PREVIEW_TREE_STYLE}.Heading",
            background=COLORS["accent_bg"],
            foreground=COLORS["primary_dark"],
            font=(font_family, 11, "bold"),
            relief=tk.FLAT,
            padding=(10, 8),
        )
        style.map(
            PRICE_PREVIEW_TREE_STYLE,
            background=[("selected", COLORS["primary"])],
            foreground=[("selected", "#FFFFFF")],
        )
        style.configure(
            PRICE_PREVIEW_SCROLLBAR_STYLE,
            background=COLORS["accent_bg"],
            troughcolor=COLORS["surface_alt"],
            bordercolor=COLORS["surface"],
            arrowcolor=COLORS["primary"],
            relief=tk.FLAT,
            width=14,
        )
        style.map(
            PRICE_PREVIEW_SCROLLBAR_STYLE,
            background=[("active", COLORS["accent_light"])],
            arrowcolor=[("active", "#FFFFFF")],
        )

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=0, style="Shell.TFrame")
        root.pack(fill=tk.BOTH, expand=True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)

        header = ttk.Frame(root, padding=(20, 10), style="Header.TFrame")
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(2, weight=1)
        self._traffic_lights(header).grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 14))
        ttk.Label(header, text=APP_DISPLAY_NAME, style="Title.TLabel").grid(
            row=0,
            column=1,
            sticky="w",
        )
        ttk.Label(header, text=OUTPUT_VERSION_LABEL, style="HeaderMeta.TLabel").grid(
            row=1,
            column=1,
            sticky="w",
            pady=(3, 0),
        )
        ttk.Label(
            header,
            textvariable=self.module_title_var,
            style="ModuleSub.TLabel",
            background=COLORS["surface"],
        ).grid(row=0, column=2, sticky="w", padx=(24, 0))
        ttk.Label(header, textvariable=self.status_var, style="Status.TLabel").grid(
            row=0,
            column=3,
            rowspan=2,
            sticky="e",
        )

        body = ttk.Frame(root, style="App.TFrame")
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        sidebar = ttk.Frame(body, padding=(10, 14), style="Sidebar.TFrame")
        sidebar.grid(row=0, column=0, sticky="ns")
        sidebar.configure(width=GUI_LAYOUT["sidebar_width"])
        sidebar.grid_propagate(False)
        ttk.Label(
            sidebar,
            text="工作台",
            background=COLORS["sidebar"],
            foreground=COLORS["muted"],
            font=("Helvetica Neue", WORKBENCH_LABEL_FONT_SIZE, "bold"),
        ).pack(anchor="w", padx=8, pady=(0, 8))
        for index, item in enumerate(V8_9_MAIN_NAV_ITEMS):
            enabled = item in V8_9_ENABLED_NAV_ITEMS
            label = ttk.Label(
                sidebar,
                text=item if enabled else f"{item}  后续",
                style="NavActive.TLabel" if item == self.active_nav_var.get() else "NavDisabled.TLabel",
                cursor="hand2" if enabled else "arrow",
            )
            label.pack(
                fill=tk.X,
                pady=(0, 4),
            )
            self.nav_labels[item] = label
            if enabled:
                label.bind("<Button-1>", lambda _event, nav_item=item: self._show_page(nav_item))
        ttk.Frame(sidebar, style="Sidebar.TFrame").pack(fill=tk.BOTH, expand=True)
        ttk.Button(
            sidebar,
            text="系统设置",
            command=self._open_rule_config,
            style="Secondary.TButton",
        ).pack(fill=tk.X, pady=(0, 8))
        ttk.Button(
            sidebar,
            text="打开客户目录",
            command=self._open_split_dir,
            style="Secondary.TButton",
        ).pack(fill=tk.X)

        content = ttk.Frame(body, padding=16, style="Content.TFrame")
        content.grid(row=0, column=1, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(2, weight=1)

        ttk.Label(content, textvariable=self.module_title_var, style="ModuleTitle.TLabel").grid(
            row=0,
            column=0,
            sticky="w",
        )
        ttk.Label(
            content,
            textvariable=self.module_subtitle_var,
            style="ModuleSub.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(3, 12))

        self.content_container = ttk.Frame(content, style="Content.TFrame")
        self.content_container.grid(row=2, column=0, sticky="nsew")
        self.content_container.columnconfigure(0, weight=1)
        self.content_container.rowconfigure(0, weight=1)

        self.fee_page = ttk.Frame(self.content_container, style="Content.TFrame")
        self.fee_page.grid(row=0, column=0, sticky="nsew")
        self.fee_page.columnconfigure(0, weight=1)
        self.fee_page.rowconfigure(2, weight=1)
        self._build_fee_calculation_page(self.fee_page)

        self.balance_page = ttk.Frame(self.content_container, style="Content.TFrame")
        self.balance_page.grid(row=0, column=0, sticky="nsew")
        self.balance_page.columnconfigure(0, weight=1)
        self.balance_page.rowconfigure(1, weight=1)
        self._build_balance_page(self.balance_page)

        self.price_preview_page = ttk.Frame(self.content_container, style="Content.TFrame")
        self.price_preview_page.grid(row=0, column=0, sticky="nsew")
        self.price_preview_page.columnconfigure(0, weight=1)
        self.price_preview_page.rowconfigure(1, weight=1)
        self._build_price_preview_page(self.price_preview_page)

        self.bill_splitter_page = ttk.Frame(self.content_container, style="Content.TFrame")
        self.bill_splitter_page.grid(row=0, column=0, sticky="nsew")
        self.bill_splitter_page.columnconfigure(0, weight=1)
        self.bill_splitter_page.rowconfigure(3, weight=1)
        self._build_bill_splitter_page(self.bill_splitter_page)

        self.balance_upload_page = ttk.Frame(self.content_container, style="Content.TFrame")
        self.balance_upload_page.grid(row=0, column=0, sticky="nsew")
        self.balance_upload_page.columnconfigure(0, weight=1)
        self.balance_upload_page.rowconfigure(2, weight=1)
        self._build_balance_upload_page(self.balance_upload_page)

        self.settings_page = ttk.Frame(self.content_container, style="Content.TFrame")
        self.settings_page.grid(row=0, column=0, sticky="nsew")
        self.settings_page.columnconfigure(0, weight=1)
        self.settings_page.rowconfigure(0, weight=1)
        self._build_settings_page(self.settings_page)

        self._show_page("费用计算")

    def _build_fee_calculation_page(self, content: ttk.Frame) -> None:
        content.rowconfigure(1, weight=1)
        content.columnconfigure(0, weight=1)

        steps = ttk.Frame(content, style="Content.TFrame")
        steps.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        self.workflow_step_labels.clear()
        for index, (step_id, step_label) in enumerate(V8_2_1_WORKFLOW_STEPS, start=1):
            style_name = "StepActive.TLabel" if step_id == "config" else "StepIdle.TLabel"
            label = ttk.Label(
                steps,
                text=f"{index}. {step_label}",
                style=style_name,
                cursor="hand2",
            )
            label.pack(
                side=tk.LEFT,
                padx=(0, 8),
            )
            label.bind(
                "<Button-1>",
                lambda _event, target_step=step_id: self._show_workflow_step(target_step),
            )
            self.workflow_step_labels[step_id] = label

        page_container = ttk.Frame(content, style="Content.TFrame")
        page_container.grid(row=1, column=0, sticky="nsew")
        page_container.columnconfigure(0, weight=1)
        page_container.rowconfigure(0, weight=1)
        self.workflow_pages.clear()

        config_page = ttk.Frame(page_container, style="Content.TFrame")
        config_page.grid(row=0, column=0, sticky="nsew")
        self._build_fee_config_page(config_page)
        self.workflow_pages["config"] = config_page

        run_page = ttk.Frame(page_container, style="Content.TFrame")
        run_page.grid(row=0, column=0, sticky="nsew")
        self._build_fee_run_page(run_page)
        self.workflow_pages["run"] = run_page

        results_page = ttk.Frame(page_container, style="Content.TFrame")
        results_page.grid(row=0, column=0, sticky="nsew")
        self._build_fee_results_page(results_page)
        self.workflow_pages["results"] = results_page

        self._show_workflow_step("config")

    def _build_fee_config_page(self, content: ttk.Frame) -> None:
        content.columnconfigure(0, weight=1)
        content.rowconfigure(1, weight=1)

        input_frame = ttk.LabelFrame(
            content,
            text="销售出库单",
            padding=14,
            style="Panel.TLabelframe",
        )
        input_frame.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        input_frame.columnconfigure(1, weight=1)

        self._path_row(
            input_frame,
            0,
            "销售出库单",
            self.sales_file_var,
            self._choose_sales_file,
            button_text="多选",
            entry_state="readonly",
        )

        settings_frame = ttk.LabelFrame(
            content,
            text="当前系统设置",
            padding=14,
            style="Panel.TLabelframe",
        )
        settings_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 12))
        settings_frame.columnconfigure(1, weight=1)
        self._directory_summary_row(settings_frame, 0, "报价表目录", self.price_dir_var)
        self._directory_summary_row(settings_frame, 1, "总结果目录", self.output_dir_var)
        self._directory_summary_row(settings_frame, 2, "客户明细目录", self.split_dir_var)

        options_frame = ttk.LabelFrame(
            content,
            text="生成选项",
            padding=14,
            style="Panel.TLabelframe",
        )
        options_frame.grid(row=2, column=0, sticky="ew")
        options_frame.columnconfigure(0, weight=1)

        ttk.Checkbutton(
            options_frame,
            textvariable=self.split_option_label_var,
            variable=self.split_var,
            command=self._sync_option_state,
            style="App.TCheckbutton",
        ).grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Checkbutton(
            options_frame,
            textvariable=self.history_option_label_var,
            variable=self.history_var,
            command=self._sync_option_state,
            style="App.TCheckbutton",
        ).grid(row=1, column=0, sticky="ew", pady=(0, 8))

        config_actions = ttk.Frame(content, style="Content.TFrame")
        config_actions.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        test_button = ttk.Button(
            config_actions,
            text="开始测试",
            command=self._start_preflight,
            style="Primary.TButton",
        )
        test_button.pack(side=tk.RIGHT)
        self.preflight_buttons.append(test_button)

    def _build_fee_run_page(self, content: ttk.Frame) -> None:
        content.columnconfigure(0, weight=1)
        content.rowconfigure(2, weight=1)

        status_frame = ttk.LabelFrame(
            content,
            text="运行状态",
            padding=14,
            style="Panel.TLabelframe",
        )
        status_frame.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        status_frame.columnconfigure(1, weight=1)
        ttk.Label(status_frame, text="当前状态", style="Field.TLabel").grid(
            row=0,
            column=0,
            sticky="w",
        )
        ttk.Label(status_frame, textvariable=self.status_var, style="Status.TLabel").grid(
            row=0,
            column=1,
            sticky="w",
            padx=(12, 0),
        )
        run_button = ttk.Button(
            status_frame,
            text="开始计算",
            command=self._start_job,
            style="Primary.TButton",
            state=tk.DISABLED,
        )
        run_button.grid(row=0, column=2, sticky="e")
        self.run_buttons.append(run_button)

        summary_frame = ttk.Frame(content, style="Content.TFrame")
        summary_frame.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        for column in range(4):
            summary_frame.columnconfigure(column, weight=1)
        self._metric(summary_frame, 0, "销售表", self.summary_sales_files_var)
        self._metric(summary_frame, 1, "成功", self.summary_success_var)
        self._metric(summary_frame, 2, "失败", self.summary_failed_var)
        self._metric(summary_frame, 3, "生成文件", self.summary_outputs_var)

        log_frame = ttk.LabelFrame(content, text="运行日志", padding=8, style="Panel.TLabelframe")
        log_frame.grid(row=2, column=0, sticky="nsew")
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=13)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        self.log_text.configure(
            state=tk.DISABLED,
            bg=COLORS["log_bg"],
            fg=COLORS["log_text"],
            insertbackground=COLORS["log_text"],
            relief=tk.FLAT,
            borderwidth=0,
            font=("Menlo", 10),
        )

    def _build_fee_results_page(self, content: ttk.Frame) -> None:
        content.columnconfigure(0, weight=1)
        content.rowconfigure(1, weight=1)

        summary_frame = ttk.Frame(content, style="Content.TFrame")
        summary_frame.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        for column in range(4):
            summary_frame.columnconfigure(column, weight=1)
        self._metric(summary_frame, 0, "销售表", self.summary_sales_files_var)
        self._metric(summary_frame, 1, "成功", self.summary_success_var)
        self._metric(summary_frame, 2, "失败", self.summary_failed_var)
        self._metric(summary_frame, 3, "生成文件", self.summary_outputs_var)

        result_frame = ttk.LabelFrame(
            content,
            text="生成结果",
            padding=8,
            style="Panel.TLabelframe",
        )
        result_frame.grid(row=1, column=0, sticky="nsew")
        result_frame.rowconfigure(0, weight=1)
        result_frame.columnconfigure(0, weight=1)

        self.result_tree = ttk.Treeview(
            result_frame,
            columns=("type", "name", "path"),
            show="headings",
            height=7,
            selectmode="browse",
        )
        self.result_tree.heading("type", text="类型")
        self.result_tree.heading("name", text="文件名")
        self.result_tree.heading("path", text="路径")
        self.result_tree.column("type", width=130, minwidth=100, stretch=False)
        self.result_tree.column("name", width=260, minwidth=180, stretch=False)
        self.result_tree.column("path", width=560, minwidth=260, stretch=True)
        self.result_tree.grid(row=0, column=0, sticky="nsew")
        self.result_tree.bind("<Double-1>", lambda _event: self._open_selected_result())

        result_scroll = ttk.Scrollbar(
            result_frame,
            orient=tk.VERTICAL,
            command=self.result_tree.yview,
        )
        result_scroll.grid(row=0, column=1, sticky="ns")
        self.result_tree.configure(yscrollcommand=result_scroll.set)

        result_actions = ttk.Frame(result_frame, style="Surface.TFrame")
        result_actions.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(
            result_actions,
            text="打开选中文件",
            command=self._open_selected_result,
            style="Secondary.TButton",
        ).pack(
            side=tk.LEFT
        )
        ttk.Button(
            result_actions,
            text="打开所在目录",
            command=self._open_selected_result_dir,
            style="Secondary.TButton",
        ).pack(
            side=tk.LEFT, padx=(10, 0)
        )
        ttk.Button(
            result_actions,
            text="打开输出目录",
            command=self._open_output_dir,
            style="Secondary.TButton",
        ).pack(
            side=tk.LEFT, padx=(10, 0)
        )
        rerun_button = ttk.Button(
            result_actions,
            text="重新测试",
            command=self._start_preflight,
            style="Primary.TButton",
        )
        rerun_button.pack(side=tk.RIGHT)
        self.preflight_buttons.append(rerun_button)

    def _build_settings_page(self, content: ttk.Frame) -> None:
        content.rowconfigure(0, weight=1)
        content.columnconfigure(0, weight=1)

        section_tabs = ttk.Notebook(content, style=SETTINGS_TOP_TAB_STYLE)
        section_tabs.grid(row=0, column=0, sticky="nsew")
        self.settings_section_notebook = section_tabs
        self.settings_section_pages = {}
        for section in V8_2_SETTINGS_SECTIONS:
            tab_frame = ttk.Frame(section_tabs, style="Content.TFrame")
            tab_frame.columnconfigure(0, weight=1)
            section_tabs.add(tab_frame, text=section)
            self.settings_section_pages[section] = tab_frame
        section_tabs.bind(
            "<<NotebookTabChanged>>",
            lambda _event: self._show_settings_section(
                section_tabs.tab(section_tabs.select(), "text")
            ),
        )

        dirs_frame = ttk.LabelFrame(
            self.settings_section_pages["目录配置"],
            text="目录配置",
            padding=14,
            style="Panel.TLabelframe",
        )
        dirs_frame.grid(row=0, column=0, sticky="ew")
        dirs_frame.columnconfigure(1, weight=1)
        self._path_row(dirs_frame, 0, "报价表目录", self.price_dir_var, self._choose_price_dir)
        self._path_row(dirs_frame, 1, "总结果目录", self.output_dir_var, self._choose_output_dir)
        self._path_row(dirs_frame, 2, "客户明细目录", self.split_dir_var, self._choose_split_dir)

        exact_frame = ttk.LabelFrame(
            self.settings_section_pages["精准映射"],
            text="精准映射",
            padding=12,
            style="Panel.TLabelframe",
        )
        exact_frame.grid(row=0, column=0, sticky="nsew")
        self.settings_section_pages["精准映射"].rowconfigure(0, weight=1)
        exact_frame.rowconfigure(1, weight=1)
        exact_frame.columnconfigure(0, weight=1)
        ttk.Label(
            exact_frame,
            text=EXPRESS_MAPPING_HELP_TEXT,
            wraplength=360,
            foreground=COLORS["muted"],
            background=COLORS["surface"],
        ).grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.settings_exact_text = tk.Text(exact_frame, height=8, wrap=tk.NONE)
        self.settings_exact_text.grid(row=1, column=0, sticky="nsew")

        keyword_frame = ttk.LabelFrame(
            self.settings_section_pages["关键词映射"],
            text="关键词映射",
            padding=12,
            style="Panel.TLabelframe",
        )
        keyword_frame.grid(row=0, column=0, sticky="nsew")
        self.settings_section_pages["关键词映射"].rowconfigure(0, weight=1)
        keyword_frame.rowconfigure(1, weight=1)
        keyword_frame.columnconfigure(0, weight=1)
        ttk.Label(
            keyword_frame,
            text=KEYWORD_MAPPING_HELP_TEXT,
            wraplength=360,
            foreground=COLORS["muted"],
            background=COLORS["surface"],
        ).grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.settings_keyword_text = tk.Text(keyword_frame, height=8, wrap=tk.NONE)
        self.settings_keyword_text.grid(row=1, column=0, sticky="nsew")

        large_frame = ttk.LabelFrame(
            self.settings_section_pages["大件规则"],
            text="大件规则",
            padding=14,
            style="Panel.TLabelframe",
        )
        large_frame.grid(row=0, column=0, sticky="ew")
        large_frame.columnconfigure(1, weight=1)
        ttk.Label(large_frame, text="大件快递", style="Field.TLabel").grid(
            row=0,
            column=0,
            sticky="w",
            pady=5,
        )
        ttk.Entry(large_frame, textvariable=self.settings_large_companies_var).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(10, 8),
            pady=5,
        )
        ttk.Label(large_frame, text="大件首重重量", style="Field.TLabel").grid(
            row=1,
            column=0,
            sticky="w",
            pady=5,
        )
        ttk.Entry(large_frame, textvariable=self.settings_threshold_var, width=12).grid(
            row=1,
            column=1,
            sticky="w",
            padx=(10, 8),
            pady=5,
        )
        ttk.Label(large_frame, text="大件模板后缀", style="Field.TLabel").grid(
            row=2,
            column=0,
            sticky="w",
            pady=5,
        )
        ttk.Entry(large_frame, textvariable=self.settings_suffix_var, width=16).grid(
            row=2,
            column=1,
            sticky="w",
            padx=(10, 8),
            pady=5,
        )
        ttk.Label(large_frame, text="超大件快递", style="Field.TLabel").grid(
            row=3,
            column=0,
            sticky="w",
            pady=5,
        )
        ttk.Entry(large_frame, textvariable=self.settings_super_large_companies_var).grid(
            row=3,
            column=1,
            sticky="ew",
            padx=(10, 8),
            pady=5,
        )
        ttk.Label(large_frame, text="超大件首重重量", style="Field.TLabel").grid(
            row=4,
            column=0,
            sticky="w",
            pady=5,
        )
        ttk.Entry(large_frame, textvariable=self.settings_super_large_threshold_var, width=12).grid(
            row=4,
            column=1,
            sticky="w",
            padx=(10, 8),
            pady=5,
        )
        ttk.Label(large_frame, text="超大件模板后缀", style="Field.TLabel").grid(
            row=5,
            column=0,
            sticky="w",
            pady=5,
        )
        ttk.Entry(large_frame, textvariable=self.settings_super_large_suffix_var, width=16).grid(
            row=5,
            column=1,
            sticky="w",
            padx=(10, 8),
            pady=5,
        )

        upload_frame = ttk.LabelFrame(
            self.settings_section_pages["余额上传"],
            text="余额上传接口",
            padding=14,
            style="Panel.TLabelframe",
        )
        upload_frame.grid(row=0, column=0, sticky="ew")
        upload_frame.columnconfigure(1, weight=1)
        ttk.Label(upload_frame, text="上传接口地址", style="Field.TLabel").grid(
            row=0,
            column=0,
            sticky="w",
            pady=5,
        )
        ttk.Entry(upload_frame, textvariable=self.settings_balance_upload_url_var).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(10, 8),
            pady=5,
        )
        ttk.Label(upload_frame, text="上传密钥", style="Field.TLabel").grid(
            row=1,
            column=0,
            sticky="w",
            pady=5,
        )
        ttk.Entry(
            upload_frame,
            textvariable=self.settings_balance_upload_token_var,
            show="*",
        ).grid(row=1, column=1, sticky="ew", padx=(10, 8), pady=5)

        actions = ttk.Frame(content, style="Content.TFrame")
        actions.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        ttk.Button(
            actions,
            text="恢复默认规则",
            command=self._restore_default_settings_rules,
            style="Secondary.TButton",
        ).pack(side=tk.LEFT)
        ttk.Button(
            actions,
            text="保存系统设置",
            command=self._save_settings_page,
            style="Primary.TButton",
        ).pack(side=tk.RIGHT)

        self._load_rule_settings_text()
        self._show_settings_section(V8_2_SETTINGS_SECTIONS[0])

    def _build_balance_page(self, content: ttk.Frame) -> None:
        content.columnconfigure(0, weight=1)
        content.rowconfigure(1, weight=1)

        summary_frame = ttk.Frame(content, style="Content.TFrame")
        summary_frame.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        for column in range(3):
            summary_frame.columnconfigure(column, weight=1)
        self._metric(
            summary_frame,
            0,
            "累计消费总额",
            self.balance_total_consumed_var,
            columns_per_row=V8_9_3_BALANCE_METRIC_COLUMNS,
        )
        self._metric(
            summary_frame,
            1,
            "累计收款总额",
            self.balance_total_paid_var,
            columns_per_row=V8_9_3_BALANCE_METRIC_COLUMNS,
        )
        self._metric(
            summary_frame,
            2,
            "累计异常扣款",
            self.balance_total_abnormal_deducted_var,
            columns_per_row=V8_9_3_BALANCE_METRIC_COLUMNS,
        )
        self._metric(
            summary_frame,
            3,
            "可用余额合计",
            self.balance_available_balance_var,
            columns_per_row=V8_9_3_BALANCE_METRIC_COLUMNS,
        )
        self._metric(
            summary_frame,
            4,
            "欠款金额合计",
            self.balance_debt_total_var,
            columns_per_row=V8_9_3_BALANCE_METRIC_COLUMNS,
        )
        self._metric(
            summary_frame,
            5,
            "欠款客户",
            self.balance_debtor_count_var,
            columns_per_row=V8_9_3_BALANCE_METRIC_COLUMNS,
        )

        table_frame = ttk.LabelFrame(
            content,
            text="客户余额表",
            padding=8,
            style="Panel.TLabelframe",
        )
        table_frame.grid(row=1, column=0, sticky="nsew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.balance_tree = ttk.Treeview(
            table_frame,
            columns=("customer", "consumed", "paid", "deducted", "balance", "last_date", "status", "path"),
            show="headings",
            height=12,
            selectmode="browse",
        )
        for column_id, label in zip(self.balance_tree["columns"], V8_3_BALANCE_TABLE_COLUMNS):
            self.balance_tree.heading(column_id, text=label)
        self.balance_tree.column("customer", width=110, minwidth=90, stretch=False)
        self.balance_tree.column("consumed", width=120, minwidth=100, stretch=False)
        self.balance_tree.column("paid", width=120, minwidth=100, stretch=False)
        self.balance_tree.column("deducted", width=120, minwidth=110, stretch=False)
        self.balance_tree.column("balance", width=120, minwidth=100, stretch=False)
        self.balance_tree.column("last_date", width=110, minwidth=90, stretch=False)
        self.balance_tree.column("status", width=80, minwidth=70, stretch=False)
        self.balance_tree.column("path", width=420, minwidth=260, stretch=True)
        self.balance_tree.grid(row=0, column=0, sticky="nsew")
        self.balance_tree.bind("<Double-1>", lambda _event: self._open_selected_balance_history())

        table_scroll = ttk.Scrollbar(
            table_frame,
            orient=tk.VERTICAL,
            command=self.balance_tree.yview,
        )
        table_scroll.grid(row=0, column=1, sticky="ns")
        self.balance_tree.configure(yscrollcommand=table_scroll.set)

        actions = ttk.Frame(table_frame, style="Surface.TFrame")
        actions.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Label(
            actions,
            textvariable=self.balance_status_var,
            background=COLORS["surface"],
            foreground=COLORS["muted"],
        ).pack(side=tk.LEFT)
        action_buttons = ttk.Frame(actions, style="Surface.TFrame")
        action_buttons.pack(side=tk.RIGHT)
        ttk.Button(
            action_buttons,
            text="刷新数据",
            command=self._load_account_balance_dashboard,
            style="Primary.TButton",
        ).pack(side=tk.LEFT)
        ttk.Button(
            action_buttons,
            text="打开客户目录",
            command=self._open_selected_balance_customer_dir,
            style="Secondary.TButton",
        ).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(
            action_buttons,
            text="打开历史汇总表",
            command=self._open_selected_balance_history,
            style="Secondary.TButton",
        ).pack(side=tk.LEFT, padx=(8, 0))

    def _build_balance_upload_page(self, content: ttk.Frame) -> None:
        content.columnconfigure(0, weight=1)
        content.rowconfigure(2, weight=1)

        config_frame = ttk.LabelFrame(
            content,
            text="上传确认",
            padding=12,
            style="Panel.TLabelframe",
        )
        config_frame.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        config_frame.columnconfigure(1, weight=1)
        ttk.Label(config_frame, text="上传日期", style="Field.TLabel").grid(
            row=0,
            column=0,
            sticky="w",
            pady=5,
        )
        ttk.Entry(config_frame, textvariable=self.balance_upload_date_var, width=16).grid(
            row=0,
            column=1,
            sticky="w",
            padx=(10, 8),
            pady=5,
        )
        ttk.Checkbutton(
            config_frame,
            text="已确认以上日期为本次要上传的数据日期",
            variable=self.balance_upload_confirm_var,
            command=self._refresh_balance_upload_action_state,
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        summary_frame = ttk.Frame(content, style="Content.TFrame")
        summary_frame.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        for column in range(4):
            summary_frame.columnconfigure(column, weight=1)
        self._metric(summary_frame, 0, "待上传客户", self.balance_upload_customer_count_var, columns_per_row=4)
        self._metric(summary_frame, 1, "今日快递费消费", self.balance_upload_today_fee_var, columns_per_row=4)
        self._metric(summary_frame, 2, "今日余额合计", self.balance_upload_today_balance_var, columns_per_row=4)
        self._metric(summary_frame, 3, "欠款客户", self.balance_upload_debtor_count_var, columns_per_row=4)

        table_frame = ttk.LabelFrame(
            content,
            text="上传预览",
            padding=8,
            style="Panel.TLabelframe",
        )
        table_frame.grid(row=2, column=0, sticky="nsew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.balance_upload_tree = ttk.Treeview(
            table_frame,
            columns=V8_10_BALANCE_UPLOAD_COLUMN_IDS,
            show="headings",
            height=12,
            selectmode="browse",
        )
        for column_id, label in zip(V8_10_BALANCE_UPLOAD_COLUMN_IDS, V8_10_BALANCE_UPLOAD_COLUMNS):
            self.balance_upload_tree.heading(column_id, text=label)
        self.balance_upload_tree.column("customer", width=150, minwidth=110, stretch=False)
        self.balance_upload_tree.column("today_fee", width=150, minwidth=120, stretch=False)
        self.balance_upload_tree.column("today_balance", width=150, minwidth=120, stretch=False)
        self.balance_upload_tree.column("balance_date", width=120, minwidth=100, stretch=False)
        self.balance_upload_tree.column("status", width=90, minwidth=70, stretch=True)
        self.balance_upload_tree.grid(row=0, column=0, sticky="nsew")

        table_scroll = ttk.Scrollbar(
            table_frame,
            orient=tk.VERTICAL,
            command=self.balance_upload_tree.yview,
        )
        table_scroll.grid(row=0, column=1, sticky="ns")
        self.balance_upload_tree.configure(yscrollcommand=table_scroll.set)

        self.balance_upload_log_text = scrolledtext.ScrolledText(table_frame, height=4, wrap=tk.WORD)
        self.balance_upload_log_text.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self.balance_upload_log_text.configure(state=tk.DISABLED)

        actions = ttk.Frame(table_frame, style="Surface.TFrame")
        actions.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Label(
            actions,
            textvariable=self.balance_upload_status_var,
            background=COLORS["surface"],
            foreground=COLORS["muted"],
        ).pack(side=tk.LEFT)
        action_buttons = ttk.Frame(actions, style="Surface.TFrame")
        action_buttons.pack(side=tk.RIGHT)
        self.balance_upload_read_button = ttk.Button(
            action_buttons,
            text="读取数据",
            command=self._load_balance_upload_preview,
            style="Secondary.TButton",
        )
        self.balance_upload_read_button.pack(side=tk.LEFT)
        self.balance_upload_button = ttk.Button(
            action_buttons,
            text="确认并上传",
            command=self._start_balance_upload,
            style="Primary.TButton",
        )
        self.balance_upload_button.pack(side=tk.LEFT, padx=(8, 0))
        self._refresh_balance_upload_action_state()

    def _build_price_preview_page(self, content: ttk.Frame) -> None:
        content.columnconfigure(0, weight=1)
        content.rowconfigure(1, weight=1)

        control_frame = ttk.LabelFrame(
            content,
            text="客户报价表",
            padding=14,
            style="Panel.TLabelframe",
        )
        control_frame.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        control_frame.columnconfigure(1, weight=1)
        control_frame.columnconfigure(3, weight=1)
        self._directory_summary_row(control_frame, 0, "报价表目录", self.price_dir_var)
        ttk.Label(control_frame, text="业务员", style="Field.TLabel").grid(
            row=1,
            column=0,
            sticky="w",
            pady=5,
        )
        self.price_template_combo = ttk.Combobox(
            control_frame,
            textvariable=self.price_template_customer_var,
            values=self.price_template_customers,
            state="readonly",
            style=PRICE_PREVIEW_COMBO_STYLE,
        )
        self.price_template_combo.grid(row=1, column=1, sticky="ew", padx=(10, 8), pady=6)
        ttk.Button(
            control_frame,
            text="搜索",
            command=self._search_price_template,
            style="Primary.TButton",
        ).grid(row=1, column=2, sticky="e", padx=(0, 8), pady=5)
        ttk.Button(
            control_frame,
            text="同步快递报价表",
            command=self._sync_price_templates,
            style="Secondary.TButton",
        ).grid(row=1, column=3, sticky="e", pady=5)

        viewer_frame = ttk.LabelFrame(
            content,
            text="当前客户快递价格信息",
            padding=14,
            style="Panel.TLabelframe",
        )
        viewer_frame.grid(row=1, column=0, sticky="nsew")
        viewer_frame.rowconfigure(1, weight=1)
        viewer_frame.columnconfigure(0, weight=1)
        viewer_frame.columnconfigure(1, weight=0)
        ttk.Label(
            viewer_frame,
            textvariable=self.price_template_summary_var,
            background=COLORS["surface"],
            foreground=COLORS["muted"],
            wraplength=760,
        ).grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(
            viewer_frame,
            text="打开快递价格表",
            command=self._open_current_price_template,
            style="Secondary.TButton",
        ).grid(row=0, column=1, sticky="e", padx=(12, 0), pady=(0, 8))
        self.price_template_notebook = ttk.Notebook(
            viewer_frame,
            style=PRICE_PREVIEW_NOTEBOOK_STYLE,
        )
        self.price_template_notebook.grid(row=1, column=0, columnspan=2, sticky="nsew")
        self.price_template_notebook.bind(
            "<<NotebookTabChanged>>",
            lambda _event: self._on_price_template_tab_changed(),
        )
        ttk.Label(
            viewer_frame,
            textvariable=self.price_template_status_var,
            background=COLORS["surface"],
            foreground=COLORS["muted"],
            wraplength=760,
        ).grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))

    def _build_bill_splitter_page(self, content: ttk.Frame) -> None:
        content.columnconfigure(0, weight=1)
        content.rowconfigure(1, weight=1)

        steps = ttk.Frame(content, style="Content.TFrame")
        steps.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        self.bill_split_workflow_step_labels.clear()
        for index, (step_id, step_label) in enumerate(V8_9_BILL_SPLIT_WORKFLOW_STEPS, start=1):
            style_name = "StepActive.TLabel" if step_id == "config" else "StepIdle.TLabel"
            label = ttk.Label(
                steps,
                text=f"{index}. {step_label}",
                style=style_name,
                cursor="hand2",
            )
            label.pack(side=tk.LEFT, padx=(0, 8))
            label.bind(
                "<Button-1>",
                lambda _event, target_step=step_id: self._show_bill_split_workflow_step(target_step),
            )
            self.bill_split_workflow_step_labels[step_id] = label

        page_container = ttk.Frame(content, style="Content.TFrame")
        page_container.grid(row=1, column=0, sticky="nsew")
        page_container.columnconfigure(0, weight=1)
        page_container.rowconfigure(0, weight=1)
        self.bill_split_workflow_pages.clear()

        config_page = ttk.Frame(page_container, style="Content.TFrame")
        config_page.grid(row=0, column=0, sticky="nsew")
        self._build_bill_split_config_page(config_page)
        self.bill_split_workflow_pages["config"] = config_page

        run_page = ttk.Frame(page_container, style="Content.TFrame")
        run_page.grid(row=0, column=0, sticky="nsew")
        self._build_bill_split_run_page(run_page)
        self.bill_split_workflow_pages["run"] = run_page

        results_page = ttk.Frame(page_container, style="Content.TFrame")
        results_page.grid(row=0, column=0, sticky="nsew")
        self._build_bill_split_results_page(results_page)
        self.bill_split_workflow_pages["results"] = results_page

        self._show_bill_split_workflow_step("config")

    def _build_bill_split_config_page(self, content: ttk.Frame) -> None:
        content.columnconfigure(0, weight=1)
        content.rowconfigure(1, weight=1)

        control_frame = ttk.LabelFrame(
            content,
            text="账单目录",
            padding=14,
            style="Panel.TLabelframe",
        )
        control_frame.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        control_frame.columnconfigure(1, weight=1)
        control_frame.columnconfigure(3, weight=1)

        ttk.Label(control_frame, text="源目录", style="Field.TLabel").grid(
            row=0,
            column=0,
            sticky="w",
            pady=5,
        )
        ttk.Entry(
            control_frame,
            textvariable=self.bill_split_source_dir_var,
            state="readonly",
        ).grid(row=0, column=1, columnspan=3, sticky="ew", padx=(10, 8), pady=5)
        choose_button = ttk.Button(
            control_frame,
            text="选择目录",
            command=self._choose_bill_split_dir,
            style="Secondary.TButton",
        )
        choose_button.grid(row=0, column=4, pady=5)
        self.bill_split_buttons.append(choose_button)
        self.bill_split_config_buttons.append(choose_button)

        ttk.Label(control_frame, text="输出目录", style="Field.TLabel").grid(
            row=1,
            column=0,
            sticky="w",
            pady=5,
        )
        ttk.Label(
            control_frame,
            textvariable=self.bill_split_output_dir_var,
            background=COLORS["surface"],
            foreground=COLORS["muted"],
            font=("Menlo", 10),
            wraplength=620,
        ).grid(row=1, column=1, columnspan=4, sticky="ew", padx=(10, 0), pady=5)

        ttk.Label(control_frame, text="共同字段", style="Field.TLabel").grid(
            row=2,
            column=0,
            sticky="w",
            pady=5,
        )
        self.bill_split_field_combo = ttk.Combobox(
            control_frame,
            textvariable=self.bill_split_field_var,
            values=self.bill_split_common_headers,
            state="readonly",
            style=PRICE_PREVIEW_COMBO_STYLE,
        )
        self.bill_split_field_combo.grid(row=2, column=1, sticky="ew", padx=(10, 8), pady=5)
        sync_button = ttk.Button(
            control_frame,
            text="同步字段",
            command=self._sync_bill_split_fields,
            style="Secondary.TButton",
            state=tk.DISABLED,
        )
        sync_button.grid(row=2, column=2, sticky="e", pady=5, padx=(0, 8))
        self.bill_split_sync_button = sync_button
        self.bill_split_buttons.append(sync_button)
        self.bill_split_config_buttons.append(sync_button)

        scan_button = ttk.Button(
            control_frame,
            text="开始测试",
            command=self._scan_bill_split_dir,
            style="Primary.TButton",
            state=tk.DISABLED,
        )
        scan_button.grid(row=2, column=3, columnspan=2, sticky="e", pady=5)
        self.bill_split_test_button = scan_button
        self.bill_split_buttons.append(scan_button)
        self.bill_split_config_buttons.append(scan_button)

        file_frame = ttk.LabelFrame(
            content,
            text="待拆分文件",
            padding=8,
            style="Panel.TLabelframe",
        )
        file_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 12))
        file_frame.rowconfigure(0, weight=1)
        file_frame.columnconfigure(0, weight=1)
        self.bill_split_file_tree = self._build_bill_split_file_tree(file_frame)
        self.bill_split_file_tree.grid(row=0, column=0, sticky="nsew")
        file_scroll = ttk.Scrollbar(file_frame, orient=tk.VERTICAL, command=self.bill_split_file_tree.yview)
        file_scroll.grid(row=0, column=1, sticky="ns")
        self.bill_split_file_tree.configure(yscrollcommand=file_scroll.set)

    def _build_bill_split_run_page(self, content: ttk.Frame) -> None:
        content.columnconfigure(0, weight=1)
        content.rowconfigure(1, weight=1)

        status_frame = ttk.LabelFrame(
            content,
            text="运行状态",
            padding=14,
            style="Panel.TLabelframe",
        )
        status_frame.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        status_frame.columnconfigure(1, weight=1)
        ttk.Label(status_frame, text="当前状态", style="Field.TLabel").grid(
            row=0,
            column=0,
            sticky="w",
        )
        ttk.Label(status_frame, textvariable=self.bill_split_status_var, style="Status.TLabel").grid(
            row=0,
            column=1,
            sticky="w",
            padx=(12, 0),
        )
        run_button = ttk.Button(
            status_frame,
            text="开始拆分",
            command=self._start_bill_split,
            style="Primary.TButton",
            state=tk.DISABLED,
        )
        run_button.grid(row=0, column=2, sticky="e")
        self.bill_split_run_button = run_button
        self.bill_split_run_buttons.append(run_button)

        log_frame = ttk.LabelFrame(content, text="运行日志", padding=8, style="Panel.TLabelframe")
        log_frame.grid(row=1, column=0, sticky="nsew")
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        self.bill_split_log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=13)
        self.bill_split_log_text.grid(row=0, column=0, sticky="nsew")
        self.bill_split_log_text.configure(
            state=tk.DISABLED,
            bg=COLORS["log_bg"],
            fg=COLORS["log_text"],
            insertbackground=COLORS["log_text"],
            relief=tk.FLAT,
            borderwidth=0,
            font=("Menlo", 10),
        )

    def _build_bill_split_results_page(self, content: ttk.Frame) -> None:
        content.columnconfigure(0, weight=1)
        content.rowconfigure(0, weight=1)

        result_frame = ttk.LabelFrame(
            content,
            text="拆分结果",
            padding=8,
            style="Panel.TLabelframe",
        )
        result_frame.grid(row=0, column=0, sticky="nsew")
        result_frame.rowconfigure(0, weight=1)
        result_frame.columnconfigure(0, weight=1)
        self.bill_split_result_tree = ttk.Treeview(
            result_frame,
            columns=V8_9_BILL_SPLIT_RESULT_COLUMN_IDS,
            show="headings",
            height=6,
            selectmode="browse",
        )
        for column_id, label in zip(V8_9_BILL_SPLIT_RESULT_COLUMN_IDS, V8_9_BILL_SPLIT_RESULT_COLUMNS):
            self.bill_split_result_tree.heading(column_id, text=label)
        self.bill_split_result_tree.column("split_value", width=130, minwidth=110, stretch=False)
        self.bill_split_result_tree.column("name", width=260, minwidth=180, stretch=False)
        self.bill_split_result_tree.column("rows", width=80, minwidth=70, stretch=False)
        self.bill_split_result_tree.column("status", width=90, minwidth=80, stretch=False)
        self.bill_split_result_tree.column("path", width=520, minwidth=260, stretch=True)
        self.bill_split_result_tree.grid(row=0, column=0, sticky="nsew")
        self.bill_split_result_tree.bind("<Double-1>", lambda _event: self._open_selected_bill_split_result())
        result_scroll = ttk.Scrollbar(
            result_frame,
            orient=tk.VERTICAL,
            command=self.bill_split_result_tree.yview,
        )
        result_scroll.grid(row=0, column=1, sticky="ns")
        self.bill_split_result_tree.configure(yscrollcommand=result_scroll.set)

        result_actions = ttk.Frame(result_frame, style="Surface.TFrame")
        result_actions.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Label(
            result_actions,
            textvariable=self.bill_split_status_var,
            background=COLORS["surface"],
            foreground=COLORS["muted"],
        ).pack(side=tk.LEFT)
        self.bill_split_open_output_button = ttk.Button(
            result_actions,
            text="打开输出目录",
            command=self._open_bill_split_output_dir,
            style="Secondary.TButton",
            state=tk.DISABLED,
        )
        self.bill_split_open_output_button.pack(side=tk.RIGHT, padx=(8, 0))
        self.bill_split_open_selected_button = ttk.Button(
            result_actions,
            text="打开选中文件",
            command=self._open_selected_bill_split_result,
            style="Secondary.TButton",
            state=tk.DISABLED,
        )
        self.bill_split_open_selected_button.pack(side=tk.RIGHT)

    def _build_bill_split_file_tree(self, parent: ttk.Frame) -> ttk.Treeview:
        tree = ttk.Treeview(
            parent,
            columns=V8_9_BILL_SPLIT_FILE_COLUMN_IDS,
            show="headings",
            height=6,
            selectmode="browse",
        )
        for column_id, label in zip(V8_9_BILL_SPLIT_FILE_COLUMN_IDS, V8_9_BILL_SPLIT_FILE_COLUMNS):
            tree.heading(column_id, text=label)
        tree.column("name", width=220, minwidth=160, stretch=False)
        tree.column("sheet", width=90, minwidth=80, stretch=False)
        tree.column("rows", width=80, minwidth=70, stretch=False)
        tree.column("field_status", width=140, minwidth=120, stretch=False)
        tree.column("path", width=480, minwidth=260, stretch=True)
        return tree

    def _show_page(self, nav_item: str) -> None:
        if nav_item not in V8_9_ENABLED_NAV_ITEMS:
            return
        self.active_nav_var.set(nav_item)
        for item, label in self.nav_labels.items():
            label.configure(style="NavActive.TLabel" if item == nav_item else "NavDisabled.TLabel")
        if nav_item == "系统设置":
            self.module_title_var.set("系统设置")
            self.module_subtitle_var.set("维护目录配置、快递公司映射、关键词映射和大件模板规则。")
            if self.settings_page is not None:
                self.settings_page.tkraise()
        elif nav_item == "账户余额":
            self.module_title_var.set("账户余额")
            self.module_subtitle_var.set("从客户历史汇总表读取消费、收款和余额，快速识别欠款客户。")
            if self.balance_page is not None:
                self.balance_page.tkraise()
        elif nav_item == "报价预览":
            self.module_title_var.set("报价预览")
            self.module_subtitle_var.set("按客户查看报价表模板，快速核对各快递公司、各省份的发货价格。")
            if self.price_preview_page is not None:
                self.price_preview_page.tkraise()
        elif nav_item == "拆分账单":
            self.module_title_var.set("拆分账单")
            self.module_subtitle_var.set("按账单明细 sheet 拆分 Excel，支持按经手人或共同字段生成独立账单文件。")
            if self.bill_splitter_page is not None:
                self.bill_splitter_page.tkraise()
            self._show_bill_split_workflow_step("config")
        elif nav_item == "余额上传":
            self.module_title_var.set("余额上传")
            self.module_subtitle_var.set("确认上传日期后，读取客户今日快递费消费和今日余额并上传。")
            if self.balance_upload_page is not None:
                self.balance_upload_page.tkraise()
        else:
            self.module_title_var.set("费用计算")
            self.module_subtitle_var.set("按配置、运行、结果三个步骤完成快递费用计算和客户明细生成。")
            if self.fee_page is not None:
                self.fee_page.tkraise()

    def _show_workflow_step(self, step_id: str) -> None:
        if step_id not in self.workflow_pages:
            return
        self.active_workflow_var.set(step_id)
        for item, label in self.workflow_step_labels.items():
            label.configure(style="StepActive.TLabel" if item == step_id else "StepIdle.TLabel")
        self.workflow_pages[step_id].tkraise()

    def _show_bill_split_workflow_step(self, step_id: str) -> None:
        pages = self.__dict__.get("bill_split_workflow_pages", {})
        if step_id not in pages:
            return
        self.active_bill_split_workflow_var.set(step_id)
        for item, label in self.__dict__.get("bill_split_workflow_step_labels", {}).items():
            label.configure(style="StepActive.TLabel" if item == step_id else "StepIdle.TLabel")
        pages[step_id].tkraise()

    def _show_settings_section(self, section: str) -> None:
        if section not in self.settings_section_pages:
            return
        self.active_settings_section_var.set(section)

    def _set_run_buttons_state(self, state: str) -> None:
        for button in self.run_buttons:
            button.configure(state=state)

    def _set_preflight_buttons_state(self, state: str) -> None:
        for button in self.__dict__.get("preflight_buttons", []):
            button.configure(state=state)

    def _traffic_lights(self, parent: ttk.Frame) -> ttk.Frame:
        frame = ttk.Frame(parent, style="Header.TFrame")
        for color in ("#ff5f57", "#febc2e", "#28c840"):
            dot = tk.Canvas(
                frame,
                width=12,
                height=12,
                highlightthickness=0,
                bg=COLORS["surface"],
            )
            dot.create_oval(1, 1, 11, 11, fill=color, outline=color)
            dot.pack(side=tk.LEFT, padx=(0, 6))
        return frame

    def _path_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        command,
        button_text: str = "选择",
        entry_state: str = "normal",
    ) -> None:
        ttk.Label(parent, text=label, style="Field.TLabel").grid(
            row=row,
            column=0,
            sticky="w",
            pady=5,
        )
        entry = ttk.Entry(parent, textvariable=variable, state=entry_state)
        entry.grid(row=row, column=1, sticky="ew", padx=(10, 8), pady=5)
        ttk.Button(
            parent,
            text=button_text,
            command=command,
            style="Secondary.TButton",
        ).grid(row=row, column=2, pady=5)

    def _directory_summary_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
    ) -> None:
        ttk.Label(parent, text=label, style="Field.TLabel").grid(
            row=row,
            column=0,
            sticky="w",
            pady=5,
        )
        value = ttk.Label(
            parent,
            textvariable=variable,
            background=COLORS["surface"],
            foreground=COLORS["muted"],
            font=("Menlo", 10),
            wraplength=620,
        )
        value.grid(row=row, column=1, sticky="ew", padx=(10, 0), pady=5)

    def _metric(
        self,
        parent: ttk.Frame,
        column: int,
        label: str,
        variable: tk.StringVar,
        *,
        columns_per_row: int | None = None,
    ) -> None:
        frame = ttk.Frame(parent, padding=(12, 10), style="Surface.TFrame")
        row_index, column_index = _metric_grid_position(column, columns_per_row)
        frame.grid(
            row=row_index,
            column=column_index,
            sticky="ew",
            padx=(0 if column_index == 0 else 8, 0),
            pady=(0 if row_index == 0 else 8, 0),
        )
        frame.columnconfigure(0, weight=1)
        value = ttk.Label(frame, textvariable=variable, style="MetricValue.TLabel")
        value.grid(row=0, column=0, sticky="ew")
        ttk.Label(frame, text=label, style="MetricLabel.TLabel").grid(
            row=1,
            column=0,
            sticky="ew",
            pady=(3, 0),
        )

    def _choose_sales_file(self) -> None:
        initial_dir = self.sales_dir
        if self.sales_files:
            initial_dir = self.sales_files[0].parent
        paths = filedialog.askopenfilenames(
            title="选择销售出库单，可多选",
            initialdir=str(initial_dir),
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if not paths:
            return
        self.sales_files = [Path(path) for path in paths]
        self.sales_dir = self.sales_files[0].parent
        self.sales_file_var.set(self._format_sales_files_display())
        self._save_current_config()

    def _choose_price_dir(self) -> None:
        path = filedialog.askdirectory(
            title="选择报价表目录",
            initialdir=self.price_dir_var.get() or str(DEFAULT_PRICE_DIR),
        )
        if path:
            self.price_dir_var.set(path)
            self._save_current_config()

    def _choose_output_dir(self) -> None:
        path = filedialog.askdirectory(
            title="选择总结果目录",
            initialdir=self.output_dir_var.get() or str(DEFAULT_OUTPUT_DIR),
        )
        if path:
            self.output_dir_var.set(path)
            self._save_current_config()

    def _choose_split_dir(self) -> None:
        path = filedialog.askdirectory(
            title="选择客户每日明细目录",
            initialdir=self.split_dir_var.get() or str(DEFAULT_SPLIT_DIR),
        )
        if path:
            self.split_dir_var.set(path)
            self._save_current_config()

    def _sync_option_state(self) -> None:
        self.refresh_all_var.set(False)
        self._refresh_option_labels()
        self._save_current_config()

    def _refresh_option_labels(self) -> None:
        self.split_option_label_var.set(
            format_option_label("生成客户每日明细", self.split_var.get())
        )
        self.history_option_label_var.set(
            format_option_label("生成客户历史汇总", self.history_var.get())
        )
        self.refresh_all_var.set(False)

    def _sync_rule_settings_vars(self) -> None:
        self.settings_large_companies_var.set(format_large_piece_companies(self.rule_config))
        self.settings_threshold_var.set(str(self.rule_config.large_piece_threshold_kg))
        self.settings_suffix_var.set(self.rule_config.large_piece_suffix)
        self.settings_super_large_companies_var.set(
            format_super_large_piece_companies(self.rule_config)
        )
        self.settings_super_large_threshold_var.set(
            str(self.rule_config.super_large_piece_threshold_kg)
        )
        self.settings_super_large_suffix_var.set(self.rule_config.super_large_piece_suffix)

    def _load_rule_settings_text(self) -> None:
        self._sync_rule_settings_vars()
        if self.settings_exact_text is not None:
            self.settings_exact_text.delete("1.0", tk.END)
            self.settings_exact_text.insert(tk.END, format_exact_mapping_text(self.rule_config))
        if self.settings_keyword_text is not None:
            self.settings_keyword_text.delete("1.0", tk.END)
            self.settings_keyword_text.insert(tk.END, format_keyword_mapping_text(self.rule_config))

    def _build_settings_rule_config(self) -> ExpressFeeRuleConfig:
        exact_text = self.settings_exact_text.get("1.0", tk.END) if self.settings_exact_text else ""
        keyword_text = (
            self.settings_keyword_text.get("1.0", tk.END) if self.settings_keyword_text else ""
        )
        return build_rule_config_from_text_fields(
            exact_mapping_text=exact_text,
            keyword_mapping_text=keyword_text,
            large_companies_text=self.settings_large_companies_var.get(),
            threshold_text=self.settings_threshold_var.get(),
            suffix_text=self.settings_suffix_var.get(),
            super_large_companies_text=self.settings_super_large_companies_var.get(),
            super_large_threshold_text=self.settings_super_large_threshold_var.get(),
            super_large_suffix_text=self.settings_super_large_suffix_var.get(),
        )

    def _restore_default_settings_rules(self) -> None:
        self.rule_config = build_default_rule_config()
        self._load_rule_settings_text()
        self._save_current_config()
        messagebox.showinfo("已恢复", "系统设置中的规则已恢复为默认值。")

    def _save_settings_page(self) -> None:
        try:
            rule_config = self._build_settings_rule_config()
        except ValueError as exc:
            messagebox.showerror("规则错误", str(exc))
            return

        self.rule_config = rule_config
        self._sync_rule_settings_vars()
        self._save_current_config()
        messagebox.showinfo("已保存", "系统设置已保存。")

    def _format_sales_files_display(self) -> str:
        if not self.sales_files:
            return ""
        if len(self.sales_files) == 1:
            return str(self.sales_files[0])
        names = "；".join(path.name for path in self.sales_files)
        return f"已选择 {len(self.sales_files)} 个文件：{names}"

    def _sync_price_templates(self) -> None:
        try:
            catalog = scan_price_template_catalog(Path(self.price_dir_var.get()).expanduser())
        except (OSError, ValueError) as exc:
            self.price_template_catalog = None
            self.price_template_customers = []
            self.price_template_customer_var.set("")
            self._configure_price_template_combo()
            self._clear_price_template_view()
            self.price_template_status_var.set(str(exc))
            return

        self.price_template_catalog = catalog
        self.price_template_customers = catalog.customer_names
        if self.price_template_customers and self.price_template_customer_var.get() not in self.price_template_customers:
            self.price_template_customer_var.set(self.price_template_customers[0])
        self._configure_price_template_combo()
        self.price_template_status_var.set(
            f"同步完成：已识别 {len(self.price_template_customers)} 份客户报价表。"
        )

    def _search_price_template(self) -> None:
        customer = self.price_template_customer_var.get().strip()
        if not customer:
            self.price_template_status_var.set("请先选择业务员。")
            return
        if self.price_template_catalog is None:
            try:
                self._sync_price_templates()
            except (OSError, ValueError) as exc:
                self.price_template_status_var.set(str(exc))
                return
        summary = self._find_price_template_summary(customer)
        if summary is None:
            self._clear_price_template_view()
            self.price_template_status_var.set("找不到对应的报价表，请点击“同步快递报价表”后重试。")
            return

        try:
            workbook = load_price_template_workbook(summary.price_file, customer=summary.customer)
        except ValueError as exc:
            self._clear_price_template_view()
            self.price_template_status_var.set(str(exc))
            return

        self._apply_price_template_workbook(workbook)

    def _find_price_template_summary(self, customer: str):
        if self.price_template_catalog is None:
            return None
        for summary in self.price_template_catalog.summaries:
            if summary.customer == customer:
                return summary
        return None

    def _open_current_price_template(self) -> None:
        path = self.price_template_current_file
        if path is None:
            messagebox.showinfo("未选择报价表", "请先选择业务员并点击搜索。")
            return
        if not path.exists():
            messagebox.showwarning("文件不存在", f"快递价格表不存在：\n{path}")
            return
        self._open_existing_path(path)

    def _configure_price_template_combo(self) -> None:
        combo = self.__dict__.get("price_template_combo")
        if combo is not None:
            combo.configure(values=self.price_template_customers)

    def _clear_price_template_view(self) -> None:
        self.price_template_workbook = None
        self.price_template_current_file = None
        self.price_template_sheet_tabs = []
        self.price_template_rows_by_sheet = {}
        self.price_template_record_counts_by_sheet = {}
        self.price_template_selected_sheet = ""
        self.price_template_summary_var.set("选择业务员后，客户报价会按快递公司分标签展示。")
        notebook = self.__dict__.get("price_template_notebook")
        if notebook is not None:
            for tab_id in notebook.tabs():
                notebook.forget(tab_id)
        self.price_template_trees.clear()

    def _apply_price_template_workbook(self, workbook) -> None:
        self._clear_price_template_view()
        self.price_template_workbook = workbook
        self.price_template_current_file = workbook.price_file
        self.price_template_sheet_tabs = [sheet.sheet_name for sheet in workbook.sheets]
        self.price_template_rows_by_sheet = {
            sheet.sheet_name: self._pair_price_template_rows(sheet.rows) for sheet in workbook.sheets
        }
        self.price_template_record_counts_by_sheet = {
            sheet.sheet_name: len(sheet.rows) for sheet in workbook.sheets
        }
        total_rows = sum(self.price_template_record_counts_by_sheet.values())
        self.price_template_summary_var.set(
            f"客户：{workbook.customer}    报价文件：{workbook.price_file.name}    "
            f"{len(workbook.sheets)} 个快递公司    省份记录：{total_rows} 条"
        )

        notebook = self.__dict__.get("price_template_notebook")
        if notebook is not None:
            for sheet in workbook.sheets:
                tab_frame = ttk.Frame(notebook, style="Surface.TFrame")
                tab_frame.rowconfigure(0, weight=1)
                tab_frame.columnconfigure(0, weight=1)
                tab_frame.columnconfigure(2, weight=1)
                tree = ttk.Treeview(
                    tab_frame,
                    columns=V8_4_PRICE_TEMPLATE_COLUMN_IDS[:3],
                    show="headings",
                    height=12,
                    style=PRICE_PREVIEW_TREE_STYLE,
                )
                for column_id, label in zip(tree["columns"], V8_4_PRICE_TEMPLATE_COLUMNS):
                    tree.heading(column_id, text=label)
                self._configure_price_preview_tree_columns(tree, V8_4_PRICE_TEMPLATE_COLUMN_IDS[:3])
                tree.grid(row=0, column=0, sticky="nsew")
                divider = ttk.Frame(
                    tab_frame,
                    width=PRICE_PREVIEW_GROUP_DIVIDER_WIDTH,
                    style="PricePreviewDivider.TFrame",
                )
                divider.grid(row=0, column=1, sticky="ns", padx=8)
                right_tree = ttk.Treeview(
                    tab_frame,
                    columns=V8_4_PRICE_TEMPLATE_COLUMN_IDS[3:],
                    show="headings",
                    height=12,
                    style=PRICE_PREVIEW_TREE_STYLE,
                )
                for column_id, label in zip(
                    right_tree["columns"],
                    V8_4_PRICE_TEMPLATE_COLUMNS[3:],
                ):
                    right_tree.heading(column_id, text=label)
                self._configure_price_preview_tree_columns(
                    right_tree,
                    V8_4_PRICE_TEMPLATE_COLUMN_IDS[3:],
                )
                right_tree.grid(row=0, column=2, sticky="nsew")
                tree.bind(
                    "<MouseWheel>",
                    lambda event, left=tree, right=right_tree: self._mousewheel_price_preview_trees(
                        event,
                        left,
                        right,
                    ),
                )
                right_tree.bind(
                    "<MouseWheel>",
                    lambda event, left=tree, right=right_tree: self._mousewheel_price_preview_trees(
                        event,
                        left,
                        right,
                    ),
                )
                scrollbar = ttk.Scrollbar(
                    tab_frame,
                    orient=tk.VERTICAL,
                    command=lambda *args, left=tree, right=right_tree: self._yview_price_preview_trees(
                        left,
                        right,
                        *args,
                    ),
                    style=PRICE_PREVIEW_SCROLLBAR_STYLE,
                )
                scrollbar.grid(row=0, column=3, sticky="ns", padx=(8, 0))
                tree.configure(
                    yscrollcommand=lambda first, last, linked=right_tree, bar=scrollbar: self._sync_price_preview_tree_scroll(
                        linked,
                        bar,
                        first,
                        last,
                    )
                )
                right_tree.configure(yscrollcommand=scrollbar.set)
                for values in self.price_template_rows_by_sheet[sheet.sheet_name]:
                    tree.insert("", tk.END, values=values[:3])
                    right_tree.insert("", tk.END, values=values[3:])
                self.price_template_trees[sheet.sheet_name] = tree
                notebook.add(tab_frame, text=sheet.sheet_name)

        if self.price_template_sheet_tabs:
            self.price_template_selected_sheet = self.price_template_sheet_tabs[0]
            self.price_template_status_var.set(
                f"已加载 {workbook.customer} 报价：{len(workbook.sheets)} 个快递公司，{total_rows} 条省份价格。"
            )
        else:
            self.price_template_status_var.set(f"{workbook.customer} 的报价表没有可用 sheet。")

    def _on_price_template_tab_changed(self) -> None:
        if self.price_template_notebook is None or not self.price_template_notebook.tabs():
            return
        sheet_name = self.price_template_notebook.tab(self.price_template_notebook.select(), "text")
        self.price_template_selected_sheet = sheet_name
        customer = self.price_template_workbook.customer if self.price_template_workbook else ""
        row_count = self.price_template_record_counts_by_sheet.get(sheet_name, 0)
        self.price_template_status_var.set(f"当前查看：{customer} / {sheet_name}，{row_count} 条省份价格。")

    def _configure_price_preview_tree_columns(self, tree, column_ids: tuple[str, ...]) -> None:
        for column_id in column_ids:
            tree.column(
                column_id,
                width=PRICE_PREVIEW_TREE_COLUMN_WIDTHS[column_id],
                minwidth=PRICE_PREVIEW_TREE_COLUMN_WIDTHS[column_id],
                stretch=False,
                anchor=PRICE_PREVIEW_TREE_CELL_ANCHOR,
            )

    def _yview_price_preview_trees(self, left_tree, right_tree, *args) -> None:
        left_tree.yview(*args)
        right_tree.yview(*args)

    def _sync_price_preview_tree_scroll(self, linked_tree, scrollbar, first: str, last: str) -> None:
        linked_tree.yview_moveto(first)
        scrollbar.set(first, last)

    def _mousewheel_price_preview_trees(self, event, left_tree, right_tree) -> str:
        units = -1 if event.delta > 0 else 1
        left_tree.yview_scroll(units, "units")
        right_tree.yview_scroll(units, "units")
        return "break"

    def _choose_bill_split_dir(self) -> None:
        initial_dir = self.bill_split_source_dir_var.get().strip() or str(Path.home())
        path = filedialog.askdirectory(
            title="选择要拆分的账单目录",
            initialdir=initial_dir,
        )
        if not path:
            return
        self.bill_split_source_dir_var.set(path)
        self._clear_bill_split_log()
        self._clear_bill_split_results()
        self._clear_bill_split_file_trees()
        self.bill_split_scan_result = None
        self.bill_split_common_headers = []
        self.bill_split_field_var.set("")
        source_dir = Path(path).expanduser()
        output_dir = build_bill_split_output_dir(
            source_dir,
            self.bill_split_field_var.get().strip() or V8_9_BILL_SPLIT_DEFAULT_FIELD,
        )
        self.bill_split_output_dir_var.set(str(output_dir))
        self.bill_split_summary_files_var.set("0")
        self.bill_split_summary_success_var.set("0")
        self.bill_split_summary_failed_var.set("0")
        self.bill_split_summary_outputs_var.set("0")
        self.bill_split_status_var.set("已选择账单目录，等待开始测试")
        self._configure_bill_split_field_combo()
        self._append_bill_split_log(f"已选择账单目录：{path}\n点击开始测试读取共同表头。")
        self._update_bill_split_action_state()

    def _scan_bill_split_dir(self) -> None:
        scan = self._read_bill_split_scan("开始测试")
        if scan is None:
            return
        self._show_bill_split_workflow_step(V8_9_BILL_SPLIT_AUTO_WORKFLOW_TRANSITIONS["on_test_start"])
        self._apply_bill_split_scan_result(scan)

    def _sync_bill_split_fields(self) -> None:
        scan = self._read_bill_split_scan("同步字段")
        if scan is None:
            return
        self._apply_bill_split_scan_result(scan, workflow_step="config")
        self._append_bill_split_log("同步字段完成")

    def _read_bill_split_scan(self, action_label: str) -> BillSplitScanResult | None:
        source_dir = self._get_bill_split_source_dir()
        if source_dir is None:
            return None
        self._clear_bill_split_results()
        self.bill_split_status_var.set(f"{action_label}：读取共同表头中")
        self._append_bill_split_log(f"{action_label}：{source_dir}")
        try:
            scan = scan_bill_split_directory(source_dir, self.bill_split_field_var.get().strip() or V8_9_BILL_SPLIT_DEFAULT_FIELD)
        except Exception as exc:  # GUI boundary: show scan failure.
            self.bill_split_status_var.set("读取表头失败")
            self._append_bill_split_log(f"读取表头失败：{exc}")
            messagebox.showerror("读取表头失败", str(exc))
            return None
        return scan

    def _apply_bill_split_scan_result(self, scan: BillSplitScanResult, workflow_step: str | None = None) -> None:
        self.bill_split_scan_result = scan
        self.bill_split_common_headers = list(scan.common_headers)
        self.bill_split_output_dir_var.set(str(scan.output_dir))
        self.bill_split_summary_files_var.set(str(len(scan.files)))
        self.bill_split_summary_success_var.set("0")
        failed_file_count = sum(1 for file_scan in scan.files if file_scan.error)
        self.bill_split_summary_failed_var.set(str(len(scan.errors) + failed_file_count))
        self.bill_split_summary_outputs_var.set("0")
        self._configure_bill_split_field_combo()
        self._refresh_bill_split_file_tree(scan)
        if scan.logs:
            self._append_bill_split_log("\n".join(scan.logs))
        if scan.errors:
            self.bill_split_status_var.set("读取完成，但存在问题，请查看拆分日志")
            self._append_bill_split_log("\n".join(scan.errors))
        elif scan.files:
            if workflow_step == "config":
                self.bill_split_status_var.set(f"已同步共同字段：{len(scan.common_headers)} 个字段")
            else:
                split_field = self.bill_split_field_var.get().strip() or V8_9_BILL_SPLIT_DEFAULT_FIELD
                self.bill_split_status_var.set(f"已通过按【{split_field}】拆分字段的测试，可进行【开始拆分】")
        else:
            self.bill_split_status_var.set("当前目录没有可拆分的 Excel 文件")
        self._update_bill_split_action_state()
        self._show_bill_split_workflow_step(workflow_step or V8_9_BILL_SPLIT_AUTO_WORKFLOW_TRANSITIONS["on_test_done"])

    def _configure_bill_split_field_combo(self) -> None:
        combo = self.__dict__.get("bill_split_field_combo")
        if combo is not None:
            combo.configure(values=self.bill_split_common_headers)
        current_field = self.bill_split_field_var.get().strip()
        if current_field in self.bill_split_common_headers:
            return
        if V8_9_BILL_SPLIT_DEFAULT_FIELD in self.bill_split_common_headers:
            self.bill_split_field_var.set(V8_9_BILL_SPLIT_DEFAULT_FIELD)
        else:
            self.bill_split_field_var.set(self.bill_split_common_headers[0] if self.bill_split_common_headers else "")

    def _clear_bill_split_file_trees(self) -> None:
        for tree_name in ("bill_split_file_tree", "bill_split_result_file_tree"):
            tree = self.__dict__.get(tree_name)
            if tree is None:
                continue
            for item_id in tree.get_children():
                tree.delete(item_id)

    def _refresh_bill_split_file_tree(self, scan: BillSplitScanResult) -> None:
        for tree_name in ("bill_split_file_tree", "bill_split_result_file_tree"):
            self._populate_bill_split_file_tree(self.__dict__.get(tree_name), scan)

    def _populate_bill_split_file_tree(self, tree, scan: BillSplitScanResult) -> None:
        if tree is None:
            return
        for item_id in tree.get_children():
            tree.delete(item_id)
        selected_field = self.bill_split_field_var.get().strip()
        for file_scan in scan.files:
            if file_scan.error:
                if file_scan.error.startswith("读取失败"):
                    sheet_status = "读取失败"
                elif "缺少" in file_scan.error:
                    sheet_status = "缺失"
                else:
                    sheet_status = "已跳过"
                field_status = file_scan.error
            else:
                sheet_status = "已找到"
                has_field = selected_field in file_scan.headers if selected_field else False
                field_status = "可拆分" if has_field else f"缺少{selected_field or '字段'}"
            tree.insert(
                "",
                tk.END,
                values=(
                    file_scan.path.name,
                    sheet_status,
                    str(file_scan.total_rows),
                    field_status,
                    str(file_scan.path),
                ),
            )

    def _start_bill_split(self) -> None:
        if self._worker and self._worker.is_alive():
            messagebox.showinfo("正在运行", "当前任务还在运行，请稍等。")
            return
        source_dir = self._get_bill_split_source_dir()
        if source_dir is None:
            return
        split_field = self.bill_split_field_var.get().strip()
        if not split_field:
            messagebox.showwarning("字段为空", "请选择拆分字段。")
            return
        if self.bill_split_common_headers and split_field not in self.bill_split_common_headers:
            messagebox.showwarning("字段缺失", f"所选字段不在共同表头中：{split_field}")
            return

        self._clear_bill_split_results()
        self.bill_split_summary_success_var.set("0")
        self.bill_split_summary_failed_var.set("0")
        self.bill_split_summary_outputs_var.set("0")
        self.bill_split_status_var.set("拆分中")
        self.status_var.set("拆分账单中")
        self._append_bill_split_log(f"开始拆分：{source_dir}，字段：{split_field}")
        self._show_bill_split_workflow_step(V8_9_BILL_SPLIT_AUTO_WORKFLOW_TRANSITIONS["on_split_start"])
        self._set_bill_split_controls_state(tk.DISABLED)

        self._worker = threading.Thread(
            target=self._run_bill_split_worker,
            args=(source_dir, split_field),
            daemon=True,
        )
        self._worker.start()
        self.after(100, self._poll_queue)

    def _run_bill_split_worker(self, source_dir: Path, split_field: str) -> None:
        try:
            result = split_bills_by_field(source_dir, split_field)
        except Exception as exc:  # GUI boundary: show unexpected errors to user.
            self._queue.put(("bill_split_error", exc))
        else:
            self._queue.put(("bill_split_done", result))

    def _apply_bill_split_result(self, result: BillSplitResult) -> None:
        self.bill_split_output_dir_var.set(str(result.output_dir))
        successful_files, failed_files = self._bill_split_source_file_counts()
        self.bill_split_summary_success_var.set(str(successful_files))
        self.bill_split_summary_failed_var.set(str(failed_files))
        self.bill_split_summary_outputs_var.set(str(len(result.output_paths)))
        self._populate_bill_split_result_table(result)
        self.bill_split_status_var.set(f"拆分完成：生成 {len(result.output_paths)} 个文件")
        self._append_bill_split_log(self._format_bill_split_completion_summary(result))
        self._update_bill_split_action_state()
        self._show_bill_split_workflow_step(V8_9_BILL_SPLIT_AUTO_WORKFLOW_TRANSITIONS["on_split_done"])

    def _bill_split_source_file_counts(self) -> tuple[int, int]:
        scan = self.__dict__.get("bill_split_scan_result")
        if scan is None:
            return (0, 0)
        files = getattr(scan, "files", [])
        failed_files = sum(1 for file_scan in files if getattr(file_scan, "error", ""))
        return (len(files) - failed_files, failed_files + len(getattr(scan, "errors", [])))

    def _populate_bill_split_result_table(self, result: BillSplitResult) -> None:
        self._clear_bill_split_results()
        for output in result.outputs:
            path = output.output_path
            item_id = f"bill-split-result-{len(self.bill_split_result_paths) + 1}"
            self.bill_split_result_paths[item_id] = path
            self.bill_split_result_tree.insert(
                "",
                tk.END,
                iid=item_id,
                values=(output.split_value, path.name, str(output.row_count), "成功", str(path)),
            )
        if self.bill_split_result_tree is not None:
            first_item = self.bill_split_result_tree.get_children()
            if first_item:
                self.bill_split_result_tree.selection_set(first_item[0])
        if self.bill_split_open_selected_button is not None and result.output_paths:
            self.bill_split_open_selected_button.configure(state=tk.NORMAL)

    def _format_bill_split_completion_summary(self, result: BillSplitResult) -> str:
        lines = [
            "拆分完成：",
            f"输出目录：{result.output_dir}",
            f"共同字段：{len(result.common_headers)} 个",
            f"生成文件：{len(result.output_paths)} 个",
        ]
        lines.extend(f"生成：{path.name}" for path in result.output_paths)
        return "\n".join(lines)

    def _clear_bill_split_results(self) -> None:
        self.bill_split_result_paths.clear()
        tree = self.__dict__.get("bill_split_result_tree")
        if tree is not None:
            for item_id in tree.get_children():
                tree.delete(item_id)
        if self.__dict__.get("bill_split_open_selected_button") is not None:
            self.bill_split_open_selected_button.configure(state=tk.DISABLED)

    def _get_bill_split_source_dir(self) -> Path | None:
        source_dir_text = self.bill_split_source_dir_var.get().strip()
        if not source_dir_text:
            messagebox.showwarning("路径错误", "请选择账单目录。")
            return None
        source_dir = Path(source_dir_text).expanduser()
        if not source_dir.exists() or not source_dir.is_dir():
            messagebox.showwarning("目录无效", f"账单目录不存在或不可读取：\n{source_dir}")
            return None
        return source_dir

    def _set_bill_split_controls_state(self, state: str) -> None:
        for button in self.bill_split_buttons:
            button.configure(state=state)
        if self.bill_split_run_button is not None:
            self.bill_split_run_button.configure(state=state)
        if self.bill_split_field_combo is not None:
            self.bill_split_field_combo.configure(state="disabled" if state == tk.DISABLED else "readonly")

    def _update_bill_split_action_state(self) -> None:
        has_source_dir = bool(self.bill_split_source_dir_var.get().strip())
        has_files = bool(self.bill_split_scan_result and self.bill_split_scan_result.files)
        has_field = bool(self.bill_split_field_var.get().strip())
        has_errors = bool(self.bill_split_scan_result and self.bill_split_scan_result.errors)
        for button in self.bill_split_buttons:
            button.configure(state=tk.NORMAL if has_source_dir else tk.DISABLED)
        if self.bill_split_run_button is not None:
            self.bill_split_run_button.configure(
                state=tk.NORMAL if has_source_dir and has_files and has_field and not has_errors else tk.DISABLED
            )
        if self.bill_split_open_output_button is not None:
            self.bill_split_open_output_button.configure(
                state=tk.NORMAL if self.bill_split_output_dir_var.get().strip() else tk.DISABLED
            )
        if self.bill_split_field_combo is not None:
            self.bill_split_field_combo.configure(state="readonly" if self.bill_split_common_headers else "disabled")

    def _append_bill_split_log(self, text: str) -> None:
        log_text = self.bill_split_log_text
        if log_text is None:
            return
        log_text.configure(state=tk.NORMAL)
        if log_text.index("end-1c") != "1.0":
            log_text.insert(tk.END, "\n")
        log_text.insert(tk.END, text)
        log_text.see(tk.END)
        log_text.configure(state=tk.DISABLED)

    def _clear_bill_split_log(self) -> None:
        log_text = self.bill_split_log_text
        if log_text is None:
            return
        log_text.configure(state=tk.NORMAL)
        log_text.delete("1.0", tk.END)
        log_text.configure(state=tk.DISABLED)

    def _open_bill_split_output_dir(self) -> None:
        value = self.bill_split_output_dir_var.get().strip()
        if not value:
            messagebox.showinfo("未选择输出目录", "请先选择账单目录。")
            return
        path = Path(value).expanduser()
        if not path.exists():
            messagebox.showwarning("目录不存在", f"输出目录尚未生成：\n{path}")
            return
        self._open_existing_path(path)

    def _get_selected_bill_split_result_path(self) -> Path | None:
        if self.bill_split_result_tree is None:
            return None
        selection = self.bill_split_result_tree.selection()
        if not selection:
            messagebox.showinfo("未选择文件", "请先在拆分结果中选择一个文件。")
            return None
        return self.bill_split_result_paths.get(selection[0])

    def _open_selected_bill_split_result(self) -> None:
        path = self._get_selected_bill_split_result_path()
        if path is None:
            return
        if not path.exists():
            messagebox.showwarning("文件不存在", f"文件不存在：\n{path}")
            return
        self._open_existing_path(path)

    def _pair_price_template_rows(self, rows) -> list[tuple[str, str, str, str, str, str]]:
        display_rows = [self._format_price_template_display_row(row) for row in rows]
        split_index = (len(display_rows) + 1) // 2
        left_rows = display_rows[:split_index]
        right_rows = display_rows[split_index:]
        paired_rows: list[tuple[str, str, str, str, str, str]] = []
        for index, left in enumerate(left_rows):
            right = right_rows[index] if index < len(right_rows) else ("", "", "")
            paired_rows.append(left + right)
        return paired_rows

    def _format_price_template_display_row(self, row) -> tuple[str, str, str]:
        return (
            row.province,
            self._format_template_price(row.first_price),
            self._format_template_price(row.extra_price),
        )

    def _format_template_price(self, value: object) -> str:
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return self._format_currency(float(value))
        text = "" if value is None else str(value)
        return text

    def _start_job(self) -> None:
        if self._worker and self._worker.is_alive():
            self._show_workflow_step("run")
            messagebox.showinfo("正在运行", "当前任务还在运行，请稍等。")
            return

        config = self._build_config(require_output_access=True)
        if config is None:
            return
        if not self._preflight_allows_run(config):
            self._set_run_buttons_state(tk.DISABLED)
            self._show_page("费用计算")
            self._show_workflow_step("config")
            messagebox.showwarning("请先开始测试", "请先点击“开始测试”，测试通过后再开始计算。")
            return
        self._save_current_config()

        self._clear_log()
        self._clear_results()
        self._reset_summary()
        self._append_log("开始运行...")
        self.status_var.set("运行中")
        self._show_page("费用计算")
        self._show_workflow_step(V8_2_1_AUTO_WORKFLOW_TRANSITIONS["on_start"])
        self._set_preflight_buttons_state(tk.DISABLED)
        self._set_run_buttons_state(tk.DISABLED)
        self._last_result = None

        self._worker = threading.Thread(
            target=self._run_job_worker,
            args=(config,),
            daemon=True,
        )
        self._worker.start()
        self.after(100, self._poll_queue)

    def _start_preflight(self) -> None:
        if self._worker and self._worker.is_alive():
            self._show_workflow_step("run")
            messagebox.showinfo("正在运行", "当前任务还在运行，请稍等。")
            return

        self._clear_log()
        self._clear_results()
        self._reset_summary()
        self._last_result = None
        self._last_preflight_result = None
        self._last_preflight_config_signature = None
        self.status_var.set("测试中")
        self._show_page("费用计算")
        self._show_workflow_step("run")
        self._set_preflight_buttons_state(tk.DISABLED)
        self._set_run_buttons_state(tk.DISABLED)
        self._append_log("开始运行前测试：只检查目录、表头和关键字段，不生成任何结果文件。")

        config = self._build_config(require_output_access=False)
        if config is None:
            self.status_var.set("测试未通过")
            self._set_preflight_buttons_state(tk.NORMAL)
            self._set_run_buttons_state(tk.DISABLED)
            self._append_log("运行前测试未通过：请先补全配置，再重新点击“开始测试”。")
            return

        self._worker = threading.Thread(
            target=self._run_preflight_worker,
            args=(config,),
            daemon=True,
        )
        self._worker.start()
        self.after(100, self._poll_queue)

    def _apply_preflight_result(
        self,
        config: ExpressFeeBatchJobConfig,
        result: ExpressFeePreflightResult,
    ) -> None:
        logs = result.logs or ["运行前测试没有返回详细信息。"]
        self._append_log("\n".join(self._format_customer_error_line(line) for line in logs))
        if result.ok:
            self._last_preflight_result = result
            self._last_preflight_config_signature = self._config_signature(config)
            self.status_var.set("测试通过，等待开始计算")
            self._set_run_buttons_state(tk.NORMAL)
            self._show_workflow_step("run")
            return

        self._last_preflight_result = None
        self._last_preflight_config_signature = None
        self.status_var.set("测试未通过")
        self._set_run_buttons_state(tk.DISABLED)
        self._show_workflow_step("run")

    def _preflight_allows_run(self, config: ExpressFeeBatchJobConfig) -> bool:
        return (
            self._last_preflight_result is not None
            and self._last_preflight_result.ok
            and self._last_preflight_config_signature == self._config_signature(config)
        )

    def _config_signature(self, config: ExpressFeeBatchJobConfig) -> tuple[object, ...]:
        rule_config = config.rule_config or build_default_rule_config()
        return (
            tuple(self._file_signature(path) for path in config.sales_files),
            self._price_dir_signature(config.price_dir),
            str(config.price_dir.expanduser().resolve()),
            str(config.output_dir.expanduser().resolve()) if config.output_dir else "",
            str(config.split_dir.expanduser().resolve()) if config.split_dir else "",
            config.round_digits,
            config.split_customer_daily_files,
            config.generate_customer_history,
            config.refresh_all_customers,
            tuple(sorted(rule_config.exact_company_map.items())),
            tuple((item.keyword, item.standard_name) for item in rule_config.keyword_company_rules),
            tuple(sorted(rule_config.large_piece_companies)),
            rule_config.large_piece_threshold_kg,
            rule_config.large_piece_suffix,
        )

    def _file_signature(self, path: Path) -> tuple[str, int | None, int | None]:
        resolved_path = path.expanduser().resolve()
        try:
            stat = resolved_path.stat()
        except OSError:
            return (str(resolved_path), None, None)
        return (str(resolved_path), stat.st_mtime_ns, stat.st_size)

    def _price_dir_signature(self, price_dir: Path) -> tuple[tuple[str, int | None, int | None], ...]:
        resolved_dir = price_dir.expanduser().resolve()
        try:
            price_files = self._find_price_workbooks(resolved_dir)
        except OSError:
            return ((str(resolved_dir), None, None),)
        return tuple(self._file_signature(path) for path in price_files)

    def _build_config(self, require_output_access: bool = True) -> ExpressFeeBatchJobConfig | None:
        sales_files = [path.expanduser() for path in self.sales_files]
        price_dir_text = self.price_dir_var.get().strip()
        output_dir_text = self.output_dir_var.get().strip()
        split_dir_text = self.split_dir_var.get().strip()
        price_dir = Path(price_dir_text).expanduser()
        output_dir = Path(output_dir_text).expanduser()
        split_dir = Path(split_dir_text).expanduser()

        if not sales_files:
            messagebox.showerror("路径错误", "请选择至少一个销售出库单。")
            return None
        missing_files = [path for path in sales_files if not path.exists()]
        if missing_files:
            joined = "\n".join(str(path) for path in missing_files[:5])
            messagebox.showerror("路径错误", f"销售出库单不存在：\n{joined}")
            return None
        if not price_dir_text:
            messagebox.showerror("路径错误", "请选择报价表目录。")
            return None
        if not price_dir.exists():
            messagebox.showerror("路径错误", f"报价表目录不存在：\n{price_dir}")
            return None
        price_dir = self._ensure_price_dir_access(price_dir)
        if price_dir is None:
            return None
        if not output_dir_text:
            messagebox.showerror("路径错误", "请选择总结果目录。")
            return None
        if require_output_access:
            output_dir = self._ensure_writable_dir_access(
                output_dir,
                "总结果目录",
                "重新选择总结果目录",
                self.output_dir_var,
            )
            if output_dir is None:
                return None
        if (self.split_var.get() or self.history_var.get()) and not split_dir_text:
            messagebox.showerror("路径错误", "请选择客户每日明细目录。")
            return None
        if require_output_access and (self.split_var.get() or self.history_var.get()) and split_dir_text:
            split_dir = self._ensure_writable_dir_access(
                split_dir,
                "客户每日明细目录",
                "重新选择客户每日明细目录",
                self.split_dir_var,
                check_existing_subdirs=True,
            )
            if split_dir is None:
                return None

        return ExpressFeeBatchJobConfig(
            sales_files=sales_files,
            price_dir=price_dir,
            output_dir=output_dir,
            split_dir=split_dir,
            split_customer_daily_files=self.split_var.get(),
            generate_customer_history=self.history_var.get(),
            refresh_all_customers=False,
            rule_config=self.rule_config,
        )

    def _ensure_price_dir_access(self, price_dir: Path) -> Path | None:
        try:
            price_files = self._find_price_workbooks(price_dir)
        except OSError as exc:
            return self._ask_reselect_price_dir(
                price_dir,
                f"当前报价目录无法读取：\n{price_dir}\n\n系统返回：{exc}",
            )

        if price_files:
            return price_dir

        return self._ask_reselect_price_dir(
            price_dir,
            "当前报价目录没有读取到任何 .xlsx 报价文件：\n"
            f"{price_dir}\n\n"
            "如果你确认文件存在，通常是 macOS 还没有给当前应用授权访问该目录。"
        )

    def _ask_reselect_price_dir(self, current_dir: Path, message: str) -> Path | None:
        should_select = messagebox.askyesno(
            "报价目录需要重新选择",
            f"{message}\n\n是否现在重新选择报价表目录？",
        )
        if not should_select:
            return None

        try:
            initial_dir = current_dir if current_dir.exists() else DEFAULT_PRICE_DIR
        except OSError:
            initial_dir = DEFAULT_PRICE_DIR

        selected = filedialog.askdirectory(
            title="重新选择报价表目录",
            initialdir=str(initial_dir),
        )
        if not selected:
            return None

        new_price_dir = Path(selected).expanduser()
        try:
            price_files = self._find_price_workbooks(new_price_dir)
        except OSError as exc:
            messagebox.showerror(
                "报价目录无法读取",
                f"重新选择后仍然无法读取：\n{new_price_dir}\n\n系统返回：{exc}",
            )
            return None

        if not price_files:
            messagebox.showerror(
                "报价文件不存在",
                f"重新选择后仍然没有读取到 .xlsx 报价文件：\n{new_price_dir}",
            )
            return None

        self.price_dir_var.set(str(new_price_dir))
        self._save_current_config()
        return new_price_dir

    def _find_price_workbooks(self, price_dir: Path) -> list[Path]:
        return sorted(
            path
            for path in price_dir.rglob("*.xlsx")
            if not path.name.startswith("~$") and path.is_file()
        )

    def _ensure_writable_dir_access(
        self,
        directory: Path,
        label: str,
        dialog_title: str,
        variable: tk.StringVar,
        check_existing_subdirs: bool = False,
    ) -> Path | None:
        try:
            self._assert_writable_dir(directory, check_existing_subdirs)
        except OSError as exc:
            return self._ask_reselect_writable_dir(
                directory,
                label,
                dialog_title,
                variable,
                f"当前{label}无法写入：\n{directory}\n\n系统返回：{exc}",
                check_existing_subdirs,
            )
        return directory

    def _ask_reselect_writable_dir(
        self,
        current_dir: Path,
        label: str,
        dialog_title: str,
        variable: tk.StringVar,
        message: str,
        check_existing_subdirs: bool,
    ) -> Path | None:
        should_select = messagebox.askyesno(
            f"{label}需要重新选择",
            f"{message}\n\n是否现在重新选择？",
        )
        if not should_select:
            return None

        try:
            initial_dir = current_dir if current_dir.exists() else DEFAULT_OUTPUT_DIR
        except OSError:
            initial_dir = DEFAULT_OUTPUT_DIR

        selected = filedialog.askdirectory(
            title=dialog_title,
            initialdir=str(initial_dir),
        )
        if not selected:
            return None

        new_dir = Path(selected).expanduser()
        try:
            self._assert_writable_dir(new_dir, check_existing_subdirs)
        except OSError as exc:
            messagebox.showerror(
                f"{label}无法写入",
                f"重新选择后仍然无法写入：\n{new_dir}\n\n系统返回：{exc}",
            )
            return None

        variable.set(str(new_dir))
        self._save_current_config()
        return new_dir

    def _assert_writable_dir(
        self,
        directory: Path,
        check_existing_subdirs: bool = False,
    ) -> None:
        directory.mkdir(parents=True, exist_ok=True)
        self._write_access_probe(directory)
        if check_existing_subdirs:
            for child in directory.iterdir():
                if child.is_dir() and not child.name.startswith("."):
                    self._write_access_probe(child)

    def _write_access_probe(self, directory: Path) -> None:
        probe_path = directory / f".express_app_access_test_{uuid.uuid4().hex}"
        probe_path.write_text("ok", encoding="utf-8")
        probe_path.unlink()

    def _build_gui_config_from_state(self) -> GuiConfig:
        sales_dir = self.sales_dir
        if self.sales_files:
            sales_dir = self.sales_files[0].parent
        return GuiConfig(
            sales_files=[path.expanduser() for path in self.sales_files],
            sales_dir=sales_dir.expanduser(),
            price_dir=Path(self.price_dir_var.get()).expanduser(),
            output_dir=Path(self.output_dir_var.get()).expanduser(),
            split_dir=Path(self.split_dir_var.get()).expanduser(),
            split_customer_daily_files=self.split_var.get(),
            generate_customer_history=self.history_var.get(),
            refresh_all_customers=False,
            rule_config=self.rule_config,
        )

    def _clear_results(self) -> None:
        self.result_paths.clear()
        for item_id in self.result_tree.get_children():
            self.result_tree.delete(item_id)

    def _reset_summary(self) -> None:
        self.summary_sales_files_var.set("0")
        self.summary_success_var.set("0")
        self.summary_failed_var.set("0")
        self.summary_outputs_var.set("0")

    def _update_summary(self, result: ExpressFeeBatchJobResult) -> None:
        self.summary_sales_files_var.set(str(len(result.job_results)))
        self.summary_success_var.set(str(result.success_rows))
        self.summary_failed_var.set(str(result.failed_rows))
        self.summary_outputs_var.set(str(len(self.result_paths)))

    def _load_account_balance_dashboard(self) -> None:
        split_dir = Path(self.split_dir_var.get()).expanduser()
        dashboard = collect_account_balance_dashboard(split_dir)
        self._apply_account_balance_dashboard(dashboard)

    def _apply_account_balance_dashboard(self, dashboard: AccountBalanceDashboard) -> None:
        self.balance_records = dashboard.records
        self.balance_total_consumed_var.set(self._format_currency(dashboard.total_consumed))
        self.balance_total_paid_var.set(self._format_currency(dashboard.total_paid))
        self.balance_total_abnormal_deducted_var.set(
            self._format_currency(dashboard.total_abnormal_deducted)
        )
        self.balance_available_balance_var.set(self._format_currency(dashboard.available_balance_total))
        self.balance_debt_total_var.set(self._format_currency(dashboard.debt_total))
        self.balance_debtor_count_var.set(f"{dashboard.debtor_count} 位")
        if dashboard.errors:
            self.balance_status_var.set(f"读取完成，{len(dashboard.errors)} 个客户存在问题")
        else:
            self.balance_status_var.set(f"读取完成，共 {dashboard.customer_count} 位客户")
        self._refresh_balance_table()

    def _refresh_balance_table(self) -> None:
        if self.balance_tree is None:
            return
        self.balance_history_paths.clear()
        self.balance_customer_dirs.clear()
        for item_id in self.balance_tree.get_children():
            self.balance_tree.delete(item_id)

        for index, record in enumerate(self.balance_records, start=1):
            item_id = f"balance-{index}"
            self.balance_history_paths[item_id] = record.history_file
            self.balance_customer_dirs[item_id] = record.customer_dir
            self.balance_tree.insert(
                "",
                tk.END,
                iid=item_id,
                values=(
                    record.customer,
                    self._format_currency(record.total_consumed),
                    self._format_currency(record.total_paid),
                    self._format_currency(record.total_abnormal_deducted),
                    self._format_currency(record.current_balance),
                    record.last_date.isoformat() if record.last_date else "",
                    record.status,
                    str(record.history_file),
                ),
            )
        first_item = self.balance_tree.get_children()
        if first_item:
            self.balance_tree.selection_set(first_item[0])

    def _get_selected_balance_item_id(self) -> str | None:
        if self.balance_tree is None:
            return None
        selection = self.balance_tree.selection()
        if not selection:
            messagebox.showinfo("未选择客户", "请先在客户余额表中选择一个客户。")
            return None
        return selection[0]

    def _open_selected_balance_history(self) -> None:
        item_id = self._get_selected_balance_item_id()
        if item_id is None:
            return
        path = self.balance_history_paths.get(item_id)
        if path is None or not path.exists():
            messagebox.showwarning("文件不存在", f"客户历史汇总表不存在：\n{path}")
            return
        self._open_existing_path(path)

    def _open_selected_balance_customer_dir(self) -> None:
        item_id = self._get_selected_balance_item_id()
        if item_id is None:
            return
        path = self.balance_customer_dirs.get(item_id)
        if path is None or not path.exists():
            messagebox.showwarning("目录不存在", f"客户目录不存在：\n{path}")
            return
        self._open_existing_path(path)

    def _load_balance_upload_preview(self) -> None:
        upload_date = self._parse_balance_upload_date()
        if upload_date is None:
            return
        self.balance_upload_confirm_var.set(False)
        self._clear_balance_upload_log()
        self._append_balance_upload_log(f"读取上传日期：{upload_date.isoformat()}")
        preview = collect_balance_upload_preview(Path(self.split_dir_var.get()).expanduser(), upload_date)
        self._apply_balance_upload_preview(preview)
        for error in preview.errors:
            self._append_balance_upload_log(error)

    def _parse_balance_upload_date(self) -> date | None:
        text = self.balance_upload_date_var.get().strip()
        for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                pass
        messagebox.showwarning("日期格式错误", "上传日期请填写为 YYYY-MM-DD，例如 2026-05-09。")
        return None

    def _apply_balance_upload_preview(self, preview: BalanceUploadPreview) -> None:
        self.balance_upload_preview = preview
        self.balance_upload_customer_count_var.set(f"{preview.customer_count} 位")
        self.balance_upload_today_fee_var.set(self._format_currency(preview.total_today_fee))
        self.balance_upload_today_balance_var.set(self._format_currency(preview.total_today_balance))
        self.balance_upload_debtor_count_var.set(f"{preview.debtor_count} 位")
        if preview.errors:
            self.balance_upload_status_var.set(f"读取完成，{len(preview.errors)} 个客户存在问题")
        elif preview.carried_forward_count:
            self.balance_upload_status_var.set(
                f"读取完成，共 {preview.customer_count} 位客户；其中 {preview.carried_forward_count} 位沿用最近余额日期"
            )
        else:
            self.balance_upload_status_var.set(f"读取完成，共 {preview.customer_count} 位客户")
        self._refresh_balance_upload_table()
        self._refresh_balance_upload_action_state()

    def _refresh_balance_upload_table(self) -> None:
        tree = self.balance_upload_tree
        if tree is None:
            return
        for item_id in tree.get_children():
            tree.delete(item_id)
        preview = self.balance_upload_preview
        if preview is None:
            return
        for index, record in enumerate(preview.records, start=1):
            tree.insert(
                "",
                tk.END,
                iid=f"balance-upload-{index}",
                values=(
                    record.customer,
                    self._format_currency(record.today_fee),
                    self._format_currency(record.today_balance),
                    record.balance_date.isoformat(),
                    record.status,
                ),
            )
        first_item = tree.get_children()
        if first_item:
            tree.selection_set(first_item[0])

    def _refresh_balance_upload_action_state(self) -> None:
        preview = self.__dict__.get("balance_upload_preview")
        has_records = bool(preview and preview.records)
        has_confirmed_date = bool(self.balance_upload_confirm_var.get())
        has_endpoint = bool(self.settings_balance_upload_url_var.get().strip())
        has_token = bool(self.settings_balance_upload_token_var.get().strip())
        is_uploading = bool(self.__dict__.get("balance_upload_uploading", False))
        state = tk.NORMAL if has_records and has_confirmed_date and has_endpoint and has_token and not is_uploading else tk.DISABLED
        if self.balance_upload_button is not None:
            self.balance_upload_button.configure(state=state)
        if self.__dict__.get("balance_upload_read_button") is not None:
            self.balance_upload_read_button.configure(state=tk.DISABLED if is_uploading else tk.NORMAL)

    def _start_balance_upload(self) -> None:
        if self._worker and self._worker.is_alive():
            messagebox.showinfo("正在运行", "当前任务还在运行，请稍等。")
            return
        preview = self.balance_upload_preview
        if preview is None or not preview.records:
            messagebox.showwarning("没有可上传数据", "请先读取要上传的数据。")
            return
        upload_date = self._parse_balance_upload_date()
        if upload_date is None:
            return
        if upload_date != preview.upload_date:
            messagebox.showwarning("日期不一致", "当前填写的上传日期与预览日期不一致，请重新读取数据。")
            return
        if not self.balance_upload_confirm_var.get():
            messagebox.showwarning("请确认日期", "请先确认本次上传的数据日期。")
            return

        url = self.settings_balance_upload_url_var.get().strip()
        token = self.settings_balance_upload_token_var.get().strip()
        if not url or not token:
            messagebox.showwarning("上传配置缺失", "请先在系统设置的余额上传中配置接口地址和上传密钥。")
            return

        self._save_current_config()
        self.balance_upload_uploading = True
        self.balance_upload_status_var.set("上传中")
        self._append_balance_upload_log(f"开始上传：{preview.upload_date.isoformat()}，{preview.customer_count} 位客户")
        self._refresh_balance_upload_action_state()
        self._worker = threading.Thread(
            target=self._run_balance_upload_worker,
            args=(preview, url, token),
            daemon=True,
        )
        self._worker.start()
        self.after(100, self._poll_queue)

    def _run_balance_upload_worker(self, preview: BalanceUploadPreview, url: str, token: str) -> None:
        try:
            payload = build_balance_upload_payload(preview, app_version=APP_VERSION)
            result = upload_balance_payload(url, token, payload)
        except Exception as exc:  # GUI boundary: show unexpected errors to user.
            self._queue.put(("balance_upload_error", exc))
        else:
            self._queue.put(("balance_upload_done", result))

    def _append_balance_upload_log(self, text: str) -> None:
        log_text = self.balance_upload_log_text
        if log_text is None:
            return
        log_text.configure(state=tk.NORMAL)
        if log_text.index("end-1c") != "1.0":
            log_text.insert(tk.END, "\n")
        log_text.insert(tk.END, text)
        log_text.see(tk.END)
        log_text.configure(state=tk.DISABLED)

    def _clear_balance_upload_log(self) -> None:
        log_text = self.balance_upload_log_text
        if log_text is None:
            return
        log_text.configure(state=tk.NORMAL)
        log_text.delete("1.0", tk.END)
        log_text.configure(state=tk.DISABLED)

    def _format_currency(self, value: float) -> str:
        sign = "-" if value < 0 else ""
        return f"{sign}¥{abs(value):,.2f}"

    def _populate_result_table(self, result: ExpressFeeBatchJobResult) -> None:
        self._clear_results()
        for job_result in result.job_results:
            if job_result.output_path:
                self._add_result_row("总结果", job_result.output_path)
            for split_file in job_result.split_files:
                self._add_result_row("客户每日明细", split_file)
            for history_file in job_result.history_files:
                self._add_result_row("客户历史汇总", history_file)
        for history_file in result.history_files:
            self._add_result_row("客户历史汇总", history_file)

        first_item = self.result_tree.get_children()
        if first_item:
            self.result_tree.selection_set(first_item[0])

    def _add_result_row(self, category: str, path: Path) -> None:
        item_id = f"result-{len(self.result_paths) + 1}"
        self.result_paths[item_id] = path
        self.result_tree.insert(
            "",
            tk.END,
            iid=item_id,
            values=(category, path.name, str(path)),
        )

    def _get_selected_result_path(self) -> Path | None:
        selection = self.result_tree.selection()
        if not selection:
            messagebox.showinfo("未选择文件", "请先在生成结果列表中选择一个文件。")
            return None
        return self.result_paths.get(selection[0])

    def _open_selected_result(self) -> None:
        path = self._get_selected_result_path()
        if path is None:
            return
        if not path.exists():
            messagebox.showwarning("文件不存在", f"文件不存在：\n{path}")
            return
        self._open_existing_path(path)

    def _open_selected_result_dir(self) -> None:
        path = self._get_selected_result_path()
        if path is None:
            return
        directory = path.parent
        if not directory.exists():
            messagebox.showwarning("目录不存在", f"目录不存在：\n{directory}")
            return
        self._open_existing_path(directory)

    def _open_rule_config(self) -> None:
        self._show_page("系统设置")

    def _apply_rule_config(self, rule_config: ExpressFeeRuleConfig) -> None:
        self.rule_config = rule_config
        self._load_rule_settings_text()
        self._save_current_config()

    def _save_current_config(self) -> None:
        try:
            save_gui_config(self._build_gui_config_from_state())
        except OSError:
            pass

    def _on_close(self) -> None:
        self._save_current_config()
        self.destroy()

    def _run_job_worker(self, config: ExpressFeeBatchJobConfig) -> None:
        try:
            result = run_express_fee_batch_job(
                config,
                progress_callback=lambda message: self._queue.put(("log", message)),
            )
        except Exception as exc:  # GUI boundary: show unexpected errors to user.
            self._queue.put(("error", exc))
        else:
            self._queue.put(("done", result))

    def _run_preflight_worker(self, config: ExpressFeeBatchJobConfig) -> None:
        try:
            result = validate_express_fee_batch_job(
                config,
                progress_callback=lambda message: self._queue.put(
                    ("log", self._format_customer_error_line(message))
                ),
            )
        except Exception as exc:  # GUI boundary: show unexpected errors to user.
            self._queue.put(("preflight_error", exc))
        else:
            self._queue.put(("preflight_done", (config, result)))

    def _poll_queue(self) -> None:
        handled_terminal_event = False
        while True:
            try:
                kind, payload = self._queue.get_nowait()
            except queue.Empty:
                break

            if kind == "log":
                self._append_log(str(payload))
            elif kind == "preflight_done":
                handled_terminal_event = True
                config, result = payload
                assert isinstance(config, ExpressFeeBatchJobConfig)
                assert isinstance(result, ExpressFeePreflightResult)
                self._apply_preflight_result(config, result)
                self._set_preflight_buttons_state(tk.NORMAL)
            elif kind == "preflight_error":
                handled_terminal_event = True
                self._append_log(f"运行前测试失败：{payload}")
                self.status_var.set("测试失败")
                self._last_preflight_result = None
                self._last_preflight_config_signature = None
                self._set_preflight_buttons_state(tk.NORMAL)
                self._set_run_buttons_state(tk.DISABLED)
                self._show_workflow_step("run")
                messagebox.showerror("运行前测试失败", str(payload))
            elif kind == "done":
                handled_terminal_event = True
                result = payload
                assert isinstance(result, ExpressFeeBatchJobResult)
                self._last_result = result
                self._append_log(self._format_run_completion_summary(result))
                self._populate_result_table(result)
                self._update_summary(result)
                self.status_var.set("完成" if result.ok else "完成，有错误")
                self._set_preflight_buttons_state(tk.NORMAL)
                self._set_run_buttons_state(tk.NORMAL)
                self._show_workflow_step(V8_2_1_AUTO_WORKFLOW_TRANSITIONS["on_done"])
                if result.ok:
                    messagebox.showinfo("运行完成", "快递费计算已完成。")
                else:
                    messagebox.showwarning("运行完成", "任务已完成，但存在错误，请查看运行日志。")
            elif kind == "error":
                handled_terminal_event = True
                self._append_log(f"运行失败：{payload}")
                self.status_var.set("失败")
                self._set_preflight_buttons_state(tk.NORMAL)
                self._set_run_buttons_state(tk.NORMAL)
                self._show_workflow_step("run")
                messagebox.showerror("运行失败", str(payload))
            elif kind == "bill_split_done":
                handled_terminal_event = True
                result = payload
                assert isinstance(result, BillSplitResult)
                self._apply_bill_split_result(result)
                self.status_var.set("拆分完成")
                self._set_bill_split_controls_state(tk.NORMAL)
                self._update_bill_split_action_state()
                messagebox.showinfo("拆分完成", "账单拆分已完成。")
            elif kind == "bill_split_error":
                handled_terminal_event = True
                self._append_bill_split_log(f"拆分失败：{payload}")
                self.bill_split_status_var.set("拆分失败，请查看拆分日志")
                self.status_var.set("拆分失败")
                self.bill_split_summary_failed_var.set("1")
                self._set_bill_split_controls_state(tk.NORMAL)
                self._update_bill_split_action_state()
                messagebox.showerror("拆分失败", str(payload))
            elif kind == "balance_upload_done":
                handled_terminal_event = True
                self.balance_upload_uploading = False
                result = payload
                self._append_balance_upload_log(result.message)
                self.balance_upload_status_var.set(result.message)
                self.status_var.set("余额上传完成" if result.ok else "余额上传失败")
                self._refresh_balance_upload_action_state()
                if result.ok:
                    messagebox.showinfo("上传完成", result.message)
                else:
                    messagebox.showerror("上传失败", result.message)
            elif kind == "balance_upload_error":
                handled_terminal_event = True
                self.balance_upload_uploading = False
                self._append_balance_upload_log(f"上传失败：{payload}")
                self.balance_upload_status_var.set("上传失败，请查看日志")
                self.status_var.set("余额上传失败")
                self._refresh_balance_upload_action_state()
                messagebox.showerror("上传失败", str(payload))

        if not handled_terminal_event and self._worker and self._worker.is_alive():
            self.after(100, self._poll_queue)

    def _format_run_completion_summary(self, result: ExpressFeeBatchJobResult) -> str:
        lines = [
            "批量任务完成：",
            f"销售表数量：{len(result.job_results)} 个",
            f"成功文件数：{sum(1 for item in result.job_results if item.ok)} 个",
            f"错误文件数：{sum(1 for item in result.job_results if not item.ok)} 个",
            f"总处理行数：{result.total_rows} 条",
            f"总成功行数：{result.success_rows} 条",
            f"总失败行数：{result.failed_rows} 条",
            f"生成结果文件：{len([item for item in result.job_results if item.output_path.exists()])} 个",
        ]
        error_lines: list[str] = []
        for job_result in result.job_results:
            error_lines.extend(job_result.processing_errors)
            error_lines.extend(job_result.split_errors)
            error_lines.extend(job_result.history_errors)
        error_lines.extend(result.history_errors)
        if error_lines:
            lines.extend(["", "需要处理的问题："])
            lines.extend(self._format_customer_error_line(line) for line in error_lines)
        return "\n".join(lines)

    def _format_customer_error_line(self, line: str) -> str:
        home_path = str(Path.home())
        sanitized = line.replace(home_path, "~")
        sanitized = re.sub(r"(/private)?/var/folders/[^\s，。\n]+", "本地临时目录", sanitized)
        sanitized = re.sub(r"/tmp/[^\s，。\n]+", "本地临时目录", sanitized)
        sanitized = re.sub(r"/Users/[^\s，。\n]+", "本地目录", sanitized)
        return sanitized

    def _append_log(self, text: str) -> None:
        self.log_text.configure(state=tk.NORMAL)
        if self.log_text.index("end-1c") != "1.0":
            self.log_text.insert(tk.END, "\n")
        self.log_text.insert(tk.END, text)
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _clear_log(self) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _open_output_dir(self) -> None:
        path = self._path_from_entry(self.output_dir_var, "总结果目录")
        if path is None:
            return
        self._open_path(path)

    def _open_split_dir(self) -> None:
        path = self._path_from_entry(self.split_dir_var, "客户每日明细目录")
        if path is None:
            return
        self._open_path(path)

    def _path_from_entry(self, variable: tk.StringVar, label: str) -> Path | None:
        value = variable.get().strip()
        if not value:
            messagebox.showwarning("路径为空", f"请先设置{label}。")
            return None
        return Path(value).expanduser().resolve()

    def _open_existing_path(self, path: Path) -> None:
        try:
            if sys.platform == "win32":
                import os

                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.run(["open", str(path)], check=False)
            else:
                subprocess.run(["xdg-open", str(path)], check=False)
        except OSError as exc:
            messagebox.showerror("无法打开", str(exc))

    def _open_path(self, path: Path) -> None:
        path.mkdir(parents=True, exist_ok=True)
        self._open_existing_path(path)


def _metric_grid_position(index: int, columns_per_row: int | None = None) -> tuple[int, int]:
    if columns_per_row is None:
        return 0, index
    return index // columns_per_row, index % columns_per_row


def main() -> None:
    app = ExpressFeeApp()
    if not run_startup_license_gate(app):
        return
    app.mainloop()


if __name__ == "__main__":
    main()
