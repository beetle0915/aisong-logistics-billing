"""Minimal Tkinter GUI for the express fee calculator."""

from __future__ import annotations

import queue
import subprocess
import threading
import tkinter as tk
import uuid
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from express_app.core.calculator import (
    DEFAULT_OUTPUT_DIR,
    DEFAULT_PRICE_DIR,
    DEFAULT_SPLIT_DIR,
    build_default_rule_config,
)
from express_app.core.models import (
    ExpressCompanyKeywordRule,
    ExpressFeeBatchJobConfig,
    ExpressFeeBatchJobResult,
    ExpressFeeRuleConfig,
)
from express_app.core import run_express_fee_batch_job
from express_app.gui.config_store import GuiConfig, load_gui_config, save_gui_config


APP_TITLE = "艾松物流计费系统 V7.0.1"
OUTPUT_VERSION_LABEL = "V7.0.1"

COLORS = {
    "background": "#F8FAFC",
    "surface": "#FFFFFF",
    "surface_alt": "#EFF6FF",
    "border": "#CBD5E1",
    "primary": "#1E40AF",
    "primary_dark": "#1E3A8A",
    "accent": "#F59E0B",
    "success": "#16A34A",
    "danger": "#DC2626",
    "text": "#0F172A",
    "muted": "#475569",
}


class RuleConfigWindow(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Tk,
        rule_config: ExpressFeeRuleConfig,
        on_save,
    ) -> None:
        super().__init__(parent)
        self.title("规则配置")
        self.geometry("760x640")
        self.minsize(680, 560)
        self.transient(parent)
        self.grab_set()

        self.on_save = on_save
        self.large_companies_var = tk.StringVar(
            value="、".join(sorted(rule_config.large_piece_companies))
        )
        self.threshold_var = tk.StringVar(value=str(rule_config.large_piece_threshold_kg))
        self.suffix_var = tk.StringVar(value=rule_config.large_piece_suffix)

        self._build_ui()
        self._load_rule_config(rule_config)
        self.focus()

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=14)
        root.pack(fill=tk.BOTH, expand=True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)
        root.rowconfigure(3, weight=1)

        large_frame = ttk.LabelFrame(root, text="大件规则", padding=10)
        large_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        large_frame.columnconfigure(1, weight=1)
        ttk.Label(large_frame, text="大件快递").grid(row=0, column=0, sticky="w")
        ttk.Entry(large_frame, textvariable=self.large_companies_var).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(10, 0),
        )
        ttk.Label(large_frame, text="重量阈值").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(large_frame, textvariable=self.threshold_var, width=12).grid(
            row=1,
            column=1,
            sticky="w",
            padx=(10, 0),
            pady=(8, 0),
        )
        ttk.Label(large_frame, text="模板后缀").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(large_frame, textvariable=self.suffix_var, width=18).grid(
            row=2,
            column=1,
            sticky="w",
            padx=(10, 0),
            pady=(8, 0),
        )

        exact_frame = ttk.LabelFrame(root, text="快递公司精确映射", padding=10)
        exact_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        exact_frame.rowconfigure(0, weight=1)
        exact_frame.columnconfigure(0, weight=1)
        self.exact_text = tk.Text(exact_frame, height=8, wrap=tk.NONE)
        self.exact_text.grid(row=0, column=0, sticky="nsew")
        exact_scroll = ttk.Scrollbar(
            exact_frame,
            orient=tk.VERTICAL,
            command=self.exact_text.yview,
        )
        exact_scroll.grid(row=0, column=1, sticky="ns")
        self.exact_text.configure(yscrollcommand=exact_scroll.set)

        keyword_frame = ttk.LabelFrame(root, text="快递公司关键词映射", padding=10)
        keyword_frame.grid(row=3, column=0, sticky="nsew", pady=(0, 10))
        keyword_frame.rowconfigure(0, weight=1)
        keyword_frame.columnconfigure(0, weight=1)
        self.keyword_text = tk.Text(keyword_frame, height=7, wrap=tk.NONE)
        self.keyword_text.grid(row=0, column=0, sticky="nsew")
        keyword_scroll = ttk.Scrollbar(
            keyword_frame,
            orient=tk.VERTICAL,
            command=self.keyword_text.yview,
        )
        keyword_scroll.grid(row=0, column=1, sticky="ns")
        self.keyword_text.configure(yscrollcommand=keyword_scroll.set)

        actions = ttk.Frame(root)
        actions.grid(row=4, column=0, sticky="ew")
        ttk.Button(actions, text="恢复默认规则", command=self._restore_defaults).pack(
            side=tk.LEFT
        )
        ttk.Button(actions, text="取消", command=self.destroy).pack(side=tk.RIGHT)
        ttk.Button(actions, text="保存", command=self._save).pack(
            side=tk.RIGHT,
            padx=(0, 10),
        )

    def _load_rule_config(self, rule_config: ExpressFeeRuleConfig) -> None:
        self.large_companies_var.set("、".join(sorted(rule_config.large_piece_companies)))
        self.threshold_var.set(str(rule_config.large_piece_threshold_kg))
        self.suffix_var.set(rule_config.large_piece_suffix)
        self.exact_text.delete("1.0", tk.END)
        self.exact_text.insert(
            tk.END,
            "\n".join(
                f"{raw_name}={standard_name}"
                for raw_name, standard_name in sorted(rule_config.exact_company_map.items())
            ),
        )
        self.keyword_text.delete("1.0", tk.END)
        self.keyword_text.insert(
            tk.END,
            "\n".join(
                f"{item.keyword}={item.standard_name}"
                for item in rule_config.keyword_company_rules
            ),
        )

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
        threshold_text = self.threshold_var.get().strip()
        try:
            threshold = float(threshold_text)
        except ValueError as exc:
            raise ValueError("重量阈值必须是数字。") from exc
        if threshold <= 0:
            raise ValueError("重量阈值必须大于 0。")

        suffix = self.suffix_var.get().strip()
        if not suffix:
            raise ValueError("模板后缀不能为空。")

        large_piece_companies = self._parse_company_list(self.large_companies_var.get())
        if not large_piece_companies:
            raise ValueError("大件快递至少需要填写一个标准快递公司。")

        return ExpressFeeRuleConfig(
            exact_company_map=self._parse_mapping_text(
                self.exact_text.get("1.0", tk.END),
                "快递公司精确映射",
            ),
            keyword_company_rules=[
                ExpressCompanyKeywordRule(keyword=keyword, standard_name=standard)
                for keyword, standard in self._parse_ordered_mapping_text(
                    self.keyword_text.get("1.0", tk.END),
                    "快递公司关键词映射",
                )
            ],
            large_piece_companies=large_piece_companies,
            large_piece_threshold_kg=threshold,
            large_piece_suffix=suffix,
        )

    def _parse_company_list(self, text: str) -> set[str]:
        normalized = text.replace("，", ",").replace("、", ",").replace(";", ",")
        return {item.strip() for item in normalized.split(",") if item.strip()}

    def _parse_mapping_line(self, line: str, source_name: str, line_number: int) -> tuple[str, str]:
        separator = "=>" if "=>" in line else "="
        if separator not in line:
            raise ValueError(f"{source_name} 第 {line_number} 行缺少 =。")
        raw_name, standard_name = [part.strip() for part in line.split(separator, 1)]
        if not raw_name or not standard_name:
            raise ValueError(f"{source_name} 第 {line_number} 行不能有空值。")
        return raw_name, standard_name

    def _iter_mapping_lines(self, text: str):
        for line_number, line in enumerate(text.splitlines(), start=1):
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            yield line_number, line

    def _parse_mapping_text(self, text: str, source_name: str) -> dict[str, str]:
        mapping: dict[str, str] = {}
        for line_number, line in self._iter_mapping_lines(text):
            raw_name, standard_name = self._parse_mapping_line(line, source_name, line_number)
            if raw_name in mapping:
                raise ValueError(f"{source_name} 第 {line_number} 行重复：{raw_name}")
            mapping[raw_name] = standard_name
        if not mapping:
            raise ValueError(f"{source_name}不能为空。")
        return mapping

    def _parse_ordered_mapping_text(self, text: str, source_name: str) -> list[tuple[str, str]]:
        rules: list[tuple[str, str]] = []
        seen_keywords: set[str] = set()
        for line_number, line in self._iter_mapping_lines(text):
            keyword, standard_name = self._parse_mapping_line(line, source_name, line_number)
            if keyword in seen_keywords:
                raise ValueError(f"{source_name} 第 {line_number} 行重复：{keyword}")
            seen_keywords.add(keyword)
            rules.append((keyword, standard_name))
        if not rules:
            raise ValueError(f"{source_name}不能为空。")
        return rules


class ExpressFeeApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1160x780")
        self.minsize(1040, 720)
        self.configure(bg=COLORS["background"])

        self._queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self._worker: threading.Thread | None = None
        self._last_result: ExpressFeeBatchJobResult | None = None
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
        self.refresh_all_var = tk.BooleanVar(value=gui_config.refresh_all_customers)
        self.rule_config = gui_config.rule_config or build_default_rule_config()
        self.status_var = tk.StringVar(value="就绪")
        self.summary_sales_files_var = tk.StringVar(value="0")
        self.summary_success_var = tk.StringVar(value="0")
        self.summary_failed_var = tk.StringVar(value="0")
        self.summary_outputs_var = tk.StringVar(value="0")

        self._configure_styles()
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        font_family = "Helvetica Neue"
        style.configure("App.TFrame", background=COLORS["background"])
        style.configure("Surface.TFrame", background=COLORS["surface"])
        style.configure("Header.TFrame", background=COLORS["primary"])
        style.configure(
            "Title.TLabel",
            background=COLORS["primary"],
            foreground="#FFFFFF",
            font=(font_family, 20, "bold"),
        )
        style.configure(
            "HeaderMeta.TLabel",
            background=COLORS["primary"],
            foreground="#DBEAFE",
            font=(font_family, 11),
        )
        style.configure(
            "Status.TLabel",
            background=COLORS["accent"],
            foreground="#111827",
            padding=(12, 5),
            font=(font_family, 11, "bold"),
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
            "MetricLabel.TLabel",
            background=COLORS["surface_alt"],
            foreground=COLORS["muted"],
            font=(font_family, 10),
            anchor="center",
        )
        style.configure(
            "MetricValue.TLabel",
            background=COLORS["surface_alt"],
            foreground=COLORS["primary_dark"],
            font=(font_family, 18, "bold"),
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
            background=[("active", COLORS["primary_dark"]), ("disabled", "#94A3B8")],
            foreground=[("disabled", "#E2E8F0")],
        )
        style.configure(
            "Secondary.TButton",
            background="#E2E8F0",
            foreground=COLORS["text"],
            font=(font_family, 11),
            padding=(12, 8),
        )
        style.map("Secondary.TButton", background=[("active", "#CBD5E1")])
        style.configure(
            "App.TCheckbutton",
            background=COLORS["surface"],
            foreground=COLORS["text"],
            font=(font_family, 11),
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

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=16, style="App.TFrame")
        root.pack(fill=tk.BOTH, expand=True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(4, weight=1)

        header = ttk.Frame(root, padding=(18, 14), style="Header.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="艾松物流计费系统", style="Title.TLabel").grid(
            row=0,
            column=0,
            sticky="w",
        )
        ttk.Label(header, text=OUTPUT_VERSION_LABEL, style="HeaderMeta.TLabel").grid(
            row=1,
            column=0,
            sticky="w",
            pady=(3, 0),
        )
        ttk.Label(header, textvariable=self.status_var, style="Status.TLabel").grid(
            row=0,
            column=1,
            rowspan=2,
            sticky="e",
        )

        input_frame = ttk.LabelFrame(
            root,
            text="输入设置",
            padding=12,
            style="Panel.TLabelframe",
        )
        input_frame.grid(row=1, column=0, sticky="ew", pady=(0, 12))
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
        self._path_row(input_frame, 1, "报价表目录", self.price_dir_var, self._choose_price_dir)
        self._path_row(input_frame, 2, "总结果目录", self.output_dir_var, self._choose_output_dir)
        self._path_row(input_frame, 3, "客户明细目录", self.split_dir_var, self._choose_split_dir)

        options = ttk.Frame(input_frame, style="Surface.TFrame")
        options.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        ttk.Checkbutton(
            options,
            text="生成客户每日明细",
            variable=self.split_var,
            command=self._sync_option_state,
            style="App.TCheckbutton",
        ).pack(side=tk.LEFT, padx=(0, 18))
        ttk.Checkbutton(
            options,
            text="生成客户历史汇总",
            variable=self.history_var,
            command=self._save_current_config,
            style="App.TCheckbutton",
        ).pack(side=tk.LEFT, padx=(0, 18))
        ttk.Checkbutton(
            options,
            text="刷新全部客户历史汇总",
            variable=self.refresh_all_var,
            command=self._save_current_config,
            style="App.TCheckbutton",
        ).pack(side=tk.LEFT)

        summary_frame = ttk.Frame(root, style="App.TFrame")
        summary_frame.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        for column in range(4):
            summary_frame.columnconfigure(column, weight=1)
        self._metric(summary_frame, 0, "销售表", self.summary_sales_files_var)
        self._metric(summary_frame, 1, "成功", self.summary_success_var)
        self._metric(summary_frame, 2, "失败", self.summary_failed_var)
        self._metric(summary_frame, 3, "生成文件", self.summary_outputs_var)

        workspace = ttk.Frame(root, style="App.TFrame")
        workspace.grid(row=4, column=0, sticky="nsew", pady=(0, 12))
        workspace.columnconfigure(0, weight=2)
        workspace.columnconfigure(1, weight=3)
        workspace.rowconfigure(0, weight=1)

        log_frame = ttk.LabelFrame(workspace, text="运行日志", padding=8, style="Panel.TLabelframe")
        log_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=13)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        self.log_text.configure(
            state=tk.DISABLED,
            bg=COLORS["surface"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
            relief=tk.FLAT,
            borderwidth=0,
            font=("Menlo", 10),
        )

        result_frame = ttk.LabelFrame(workspace, text="生成结果", padding=8, style="Panel.TLabelframe")
        result_frame.grid(row=0, column=1, sticky="nsew")
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

        actions = ttk.Frame(root, style="App.TFrame")
        actions.grid(row=5, column=0, sticky="ew")
        self.run_button = ttk.Button(
            actions,
            text="开始运行",
            command=self._start_job,
            style="Primary.TButton",
        )
        self.run_button.pack(side=tk.LEFT)
        ttk.Button(
            actions,
            text="规则配置",
            command=self._open_rule_config,
            style="Secondary.TButton",
        ).pack(
            side=tk.LEFT, padx=(10, 0)
        )
        ttk.Button(
            actions,
            text="打开选中文件",
            command=self._open_selected_result,
            style="Secondary.TButton",
        ).pack(
            side=tk.LEFT
        )
        ttk.Button(
            actions,
            text="打开所在目录",
            command=self._open_selected_result_dir,
            style="Secondary.TButton",
        ).pack(
            side=tk.LEFT, padx=(10, 0)
        )
        ttk.Button(
            actions,
            text="打开输出目录",
            command=self._open_output_dir,
            style="Secondary.TButton",
        ).pack(
            side=tk.LEFT, padx=(10, 0)
        )
        ttk.Button(
            actions,
            text="打开客户目录",
            command=self._open_split_dir,
            style="Secondary.TButton",
        ).pack(
            side=tk.LEFT, padx=(10, 0)
        )

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

    def _metric(self, parent: ttk.Frame, column: int, label: str, variable: tk.StringVar) -> None:
        frame = ttk.Frame(parent, padding=(12, 10), style="Surface.TFrame")
        frame.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 8, 0))
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
        if not self.split_var.get():
            self.refresh_all_var.set(False)
        self._save_current_config()

    def _format_sales_files_display(self) -> str:
        if not self.sales_files:
            return ""
        if len(self.sales_files) == 1:
            return str(self.sales_files[0])
        names = "；".join(path.name for path in self.sales_files)
        return f"已选择 {len(self.sales_files)} 个文件：{names}"

    def _start_job(self) -> None:
        if self._worker and self._worker.is_alive():
            messagebox.showinfo("正在运行", "当前任务还在运行，请稍等。")
            return

        config = self._build_config()
        if config is None:
            return
        self._save_current_config()

        self._clear_log()
        self._clear_results()
        self._reset_summary()
        self._append_log("开始运行...")
        self.status_var.set("运行中")
        self.run_button.configure(state=tk.DISABLED)
        self._last_result = None

        self._worker = threading.Thread(
            target=self._run_job_worker,
            args=(config,),
            daemon=True,
        )
        self._worker.start()
        self.after(100, self._poll_queue)

    def _build_config(self) -> ExpressFeeBatchJobConfig | None:
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
        output_dir = self._ensure_writable_dir_access(
            output_dir,
            "总结果目录",
            "重新选择总结果目录",
            self.output_dir_var,
        )
        if output_dir is None:
            return None
        if self.split_var.get() and not split_dir_text:
            messagebox.showerror("路径错误", "请选择客户每日明细目录。")
            return None
        if (self.split_var.get() or self.history_var.get()) and split_dir_text:
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
            refresh_all_customers=self.refresh_all_var.get(),
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
            for path in price_dir.glob("*.xlsx")
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
            refresh_all_customers=self.refresh_all_var.get(),
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
        RuleConfigWindow(self, self.rule_config, self._apply_rule_config)

    def _apply_rule_config(self, rule_config: ExpressFeeRuleConfig) -> None:
        self.rule_config = rule_config
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
            result = run_express_fee_batch_job(config)
        except Exception as exc:  # GUI boundary: show unexpected errors to user.
            self._queue.put(("error", exc))
        else:
            self._queue.put(("done", result))

    def _poll_queue(self) -> None:
        try:
            kind, payload = self._queue.get_nowait()
        except queue.Empty:
            if self._worker and self._worker.is_alive():
                self.after(100, self._poll_queue)
            return

        if kind == "done":
            result = payload
            assert isinstance(result, ExpressFeeBatchJobResult)
            self._last_result = result
            self._append_log("\n".join(result.logs))
            self._populate_result_table(result)
            self._update_summary(result)
            self.status_var.set("完成" if result.ok else "完成，有错误")
            self.run_button.configure(state=tk.NORMAL)
            if result.ok:
                messagebox.showinfo("运行完成", "快递费计算已完成。")
            else:
                messagebox.showwarning("运行完成", "任务已完成，但存在错误，请查看运行日志。")
        elif kind == "error":
            self._append_log(f"运行失败：{payload}")
            self.status_var.set("失败")
            self.run_button.configure(state=tk.NORMAL)
            messagebox.showerror("运行失败", str(payload))

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
        path = Path(self.output_dir_var.get()).expanduser()
        self._open_path(path)

    def _open_split_dir(self) -> None:
        path = Path(self.split_dir_var.get()).expanduser()
        self._open_path(path)

    def _open_existing_path(self, path: Path) -> None:
        try:
            subprocess.run(["open", str(path)], check=False)
        except OSError as exc:
            messagebox.showerror("无法打开", str(exc))

    def _open_path(self, path: Path) -> None:
        path.mkdir(parents=True, exist_ok=True)
        self._open_existing_path(path)


def main() -> None:
    app = ExpressFeeApp()
    app.mainloop()


if __name__ == "__main__":
    main()
