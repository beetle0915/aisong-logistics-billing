# Bill Splitter Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a low-intrusion `拆分账单` page and core Excel splitter that splits only the `账单明细` sheet by `经手人` or another common column.

**Architecture:** Put Excel scanning and splitting in a new pure core module, then wire a separate GUI page into the existing Tkinter workbench. The GUI mirrors the `费用计算` page style, with its own file list, result list, and log panel, while existing pages remain untouched except for navigation constants and page routing.

**Tech Stack:** Python 3, Tkinter/ttk, openpyxl, unittest, existing macOS packaging script.

---

## File Structure

- Create `app/express_app/core/bill_splitter.py`: scan directories, compute common headers, filter rows, write split workbooks.
- Modify `app/express_app/gui/app.py`: add `拆分账单` navigation/page, variables, UI builders, scan/split actions, and page-specific logs.
- Modify `app/express_app/version.py`: bump app version after implementation.
- Create `tests/test_bill_splitter.py`: core scan/split behavior.
- Create `tests/test_v8_9_bill_splitter_page.py`: navigation, page constants, state and GUI routing behavior.
- Modify `tests/test_branding_packaging.py`: version metadata.
- Create `docs/PRD_艾松运费管家_V8_9_0_拆分账单.md`: product requirements.

## Task 1: Core Splitter

**Files:**
- Create: `app/express_app/core/bill_splitter.py`
- Create: `tests/test_bill_splitter.py`

- [ ] **Step 1: Write failing tests for scanning**

Create test workbooks with:

- one workbook containing `账单明细`, `Sheet2`, `账单总览`;
- one workbook containing only `Sheet2`;
- one `~$临时.xlsx`.

Assert the scanner:

```python
scan = scan_bill_split_directory(input_dir)
self.assertEqual(len(scan.files), 1)
self.assertEqual(scan.files[0].sheet_name, "账单明细")
self.assertIn("经手人", scan.common_headers)
self.assertIn("缺少 sheet：账单明细", scan.skipped_files[0].reason)
```

- [ ] **Step 2: Run the scan tests and confirm red**

Run:

```bash
python -m unittest tests.test_bill_splitter
```

Expected: import failure because `bill_splitter.py` does not exist.

- [ ] **Step 3: Implement scan dataclasses and scanner**

Add dataclasses:

```python
@dataclass(frozen=True)
class BillSplitFileScan:
    path: Path
    file_name: str
    sheet_name: str
    headers: list[str]
    data_row_count: int
    valid_row_count: int
    status: str = "可拆分"
    reason: str = ""

@dataclass(frozen=True)
class BillSplitSkippedFile:
    path: Path
    file_name: str
    reason: str

@dataclass(frozen=True)
class BillSplitScanResult:
    input_dir: Path
    output_dir: Path
    files: list[BillSplitFileScan]
    skipped_files: list[BillSplitSkippedFile]
    common_headers: list[str]
```

Implement:

```python
DETAIL_SHEET_NAME = "账单明细"
DEFAULT_SPLIT_COLUMN = "经手人"
SUMMARY_MARKERS = {"合 计", "合计", "小计", "总额", "合计："}
INVALID_SPLIT_VALUES = {"", "#N/A", "N/A", "无"}
```

Scanner must use `read_only=True`, `data_only=True`, inspect only `账单明细`, compute common headers in first valid file order, and not trust `ws.max_row`.

- [ ] **Step 4: Write failing tests for output split**

Create two workbooks where `经手人=艾松电器` appears in both source files and one workbook has a `合 计` row where `经手人` column contains `159985.00`.

Assert:

```python
result = split_bills_by_column(input_dir, "经手人")
self.assertTrue((output_dir / "经手人_艾松电器.xlsx").exists())
self.assertFalse((output_dir / "经手人_159985.00.xlsx").exists())
```

Load `经手人_艾松电器.xlsx` and assert it has two source-file sheets, each with `来源文件` and `来源行号`.

- [ ] **Step 5: Implement splitting and writers**

Implement:

```python
@dataclass(frozen=True)
class BillSplitOutputRecord:
    split_value: str
    output_path: Path
    sheet_count: int
    row_count: int

@dataclass(frozen=True)
class BillSplitResult:
    input_dir: Path
    output_dir: Path
    split_column: str
    outputs: list[BillSplitOutputRecord]
    skipped_rows: int
    skipped_files: list[BillSplitSkippedFile]
```

Core behavior:

- output dir `input_dir.parent / f"{input_dir.name}拆分结果"`;
- create dir with `exist_ok=True`;
- overwrite same output file;
- write headers + source rows + `来源文件`/`来源行号`;
- freeze first row, bold header, auto filter, bounded column widths;
- clean file and sheet names.

- [ ] **Step 6: Run core tests**

Run:

```bash
python -m unittest tests.test_bill_splitter
```

Expected: all core splitter tests pass.

## Task 2: UI Design Constants And Navigation

**Files:**
- Modify: `app/express_app/gui/app.py`
- Create: `tests/test_v8_9_bill_splitter_page.py`

- [ ] **Step 1: Write failing navigation tests**

Assert:

```python
self.assertEqual(
    list(gui_app.V8_1_MAIN_NAV_ITEMS),
    ["费用计算", "账户余额", "报价预览", "拆分账单", "系统设置"],
)
self.assertIn("拆分账单", gui_app.V8_9_ENABLED_NAV_ITEMS)
```

Assert `_show_page("拆分账单")` sets title/subtitle and raises `bill_splitter_page`.

- [ ] **Step 2: Run and confirm red**

Run:

```bash
python -m unittest tests.test_v8_9_bill_splitter_page
```

Expected: fail because constants/page route do not exist.

- [ ] **Step 3: Add navigation constants and route**

In `app.py`, add:

```python
V8_9_ENABLED_NAV_ITEMS = ("费用计算", "账户余额", "报价预览", "拆分账单", "系统设置")
V8_9_BILL_SPLITTER_FILE_COLUMNS = ("状态", "文件名", "Sheet", "数据行数", "说明")
V8_9_BILL_SPLITTER_RESULT_COLUMNS = ("拆分值", "输出文件", "Sheet 数", "记录数")
```

Use `V8_9_ENABLED_NAV_ITEMS` where navigation checks enabled pages.

- [ ] **Step 4: Add bill splitter page frame and routing**

Create `self.bill_splitter_page`, call `_build_bill_splitter_page`, and handle `_show_page("拆分账单")`:

```python
self.module_title_var.set("拆分账单")
self.module_subtitle_var.set("按账单明细中的共同字段拆分 Excel，并按字段值生成独立工作簿。")
self.bill_splitter_page.tkraise()
```

- [ ] **Step 5: Run UI navigation tests**

Run:

```bash
python -m unittest tests.test_v8_9_bill_splitter_page tests.test_v8_1_gui_navigation tests.test_v8_4_price_preview_page
```

Expected: all pass.

## Task 3: GUI Page Integration

**Files:**
- Modify: `app/express_app/gui/app.py`
- Modify: `tests/test_v8_9_bill_splitter_page.py`

- [ ] **Step 1: Write failing tests for app state helpers**

Instantiate `ExpressFeeApp.__new__`, set fake variables/widgets, and assert:

- `_apply_bill_split_scan_result(scan)` fills combobox values and file tree rows;
- default split column becomes `经手人`;
- `_append_bill_split_log("...")` writes to the splitter log variable/widget without touching fee logs.

- [ ] **Step 2: Build page UI**

Add `_build_bill_splitter_page(content)` with:

- directory `LabelFrame`;
- path row and buttons: `选择目录`, `重新扫描`, `打开目录`;
- split config row: `拆分字段` combobox, `立即拆分`, `打开结果目录`;
- file Treeview;
- result Treeview;
- log Text or Listbox area with `清空日志`.

Use existing styles: `Panel.TLabelframe`, `Field.TLabel`, `Primary.TButton`, `Secondary.TButton`.

- [ ] **Step 3: Add scan/split actions**

Import:

```python
from express_app.core.bill_splitter import scan_bill_split_directory, split_bills_by_column
```

Add actions:

- `_choose_bill_splitter_dir`;
- `_scan_bill_splitter_dir`;
- `_run_bill_splitter`;
- `_open_bill_splitter_dir`;
- `_open_bill_splitter_output_dir`;
- `_clear_bill_splitter_log`;
- `_apply_bill_split_scan_result`;
- `_apply_bill_split_result`;
- `_append_bill_split_log`.

Use worker threads for split execution if following existing long-running patterns; scan can be synchronous for first version if fast enough, but must catch exceptions and log them.

- [ ] **Step 4: Run GUI tests**

Run:

```bash
python -m unittest tests.test_v8_9_bill_splitter_page
```

Expected: pass.

## Task 4: Version, Docs, And Regression

**Files:**
- Modify: `app/express_app/version.py`
- Modify: `tests/test_branding_packaging.py`

- [ ] **Step 1: Update version test first**

Set expectations to:

```python
self.assertEqual(APP_VERSION, "8.9.0")
self.assertEqual(APP_VERSION_LABEL, "V8.9.0")
self.assertEqual(OUTPUT_VERSION_SUFFIX, "v8_9_0")
```

- [ ] **Step 2: Run and confirm red**

Run:

```bash
python -m unittest tests.test_branding_packaging
```

Expected: fail while version remains `8.8.3`.

- [ ] **Step 3: Bump app version**

Change `APP_VERSION = "8.9.0"`.

- [ ] **Step 4: Full regression**

Run:

```bash
python -m unittest discover tests
```

Expected: all tests pass.

## Task 5: Local Packaging

**Files:**
- Use existing `app/build_macos_app.py`

- [ ] **Step 1: Run macOS build**

Run from `app/`:

```bash
/Users/beetle/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -B build_macos_app.py
```

Expected: `.app` generated locally.

- [ ] **Step 2: Run self-check**

Run:

```bash
EXPRESS_APP_SELF_CHECK=1 ./艾松运费管家.app/Contents/MacOS/ExpressFeeCalculator
```

Expected: exit 0 and no GUI hang.

- [ ] **Step 3: Report local app path**

Return app path for user inspection. Do not push to GitHub.

## Self-Review

- Spec coverage: navigation, directory scan, `账单明细` only, `经手人` split, output naming, overwrite behavior, logs, tests, and local packaging are covered.
- Placeholder scan: no placeholders remain.
- Type consistency: core dataclass names and GUI method names are consistent across tasks.
