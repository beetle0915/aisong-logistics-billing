# V8.8.3 Abnormal Deductions Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add an `异常扣款记录` sheet to customer history workbooks, sync daily and cumulative abnormal deductions into `历史汇总`, and deduct those amounts from workbook and account-balance current balances for V8.8.3.

**Architecture:** Reuse the existing `收款记录` preservation pattern in `calculator.py`. Add constants and a generic preservation helper for manual sheets, update `write_customer_history_sheet()` formulas, then update workbook tests and version metadata.

**Tech Stack:** Python, openpyxl, unittest, existing Tkinter desktop app packaging metadata.

---

## File Structure

- Modify `app/express_app/core/calculator.py`: add abnormal deduction sheet constants, formulas, sheet writer, and preservation call.
- Modify `app/express_app/version.py`: bump `APP_VERSION` to `8.8.3`.
- Modify `app/express_app/core/account_balance.py`: include `异常扣款记录` amounts when calculating account-balance current balance.
- Modify `tests/test_customer_history_workbook.py`: assert sheet order, headers, formulas, formatting, and preservation.
- Modify `tests/test_v8_3_account_balance.py`: assert account balance deducts abnormal deductions and old files remain compatible.
- Modify `tests/test_branding_packaging.py`: assert version metadata.
- Create `docs/PRD_艾松运费管家_V8_8_3_异常扣款记录.md`: product requirements and acceptance criteria.

## Task 1: Tests For Workbook Structure And Formulas

**Files:**
- Modify: `tests/test_customer_history_workbook.py`

- [ ] **Step 1: Add failing assertions for the new sheet order and history headers**

Update `test_history_workbook_adds_payment_sheet_and_balance_formulas` so the generated workbook expects:

```python
self.assertEqual(
    workbook.sheetnames,
    [CUSTOMER_HISTORY_SHEET, HISTORY_DETAIL_SHEET, PAYMENT_SHEET, ABNORMAL_DEDUCTION_SHEET],
)
```

and the `历史汇总` headers are:

```python
[
    "日期",
    "单数",
    "总重量",
    "今日快递总消费",
    "累计快递费用",
    "今日收款",
    "累计收款",
    "今日异常扣款",
    "累计异常扣款",
    "当前余额",
    "顺丰单数",
    "申通单数",
    "德邦单数",
    "大件单数",
]
```

- [ ] **Step 2: Add failing assertions for abnormal deduction formulas**

In the same test, assert:

```python
self.assertEqual(
    history_sheet["H2"].value,
    '=SUMIFS(\'异常扣款记录\'!$B:$B,\'异常扣款记录\'!$A:$A,">="&A2,'
    '\'异常扣款记录\'!$A:$A,"<"&A2+1)',
)
self.assertEqual(
    history_sheet["I2"].value,
    '=SUMIFS(\'异常扣款记录\'!$B:$B,\'异常扣款记录\'!$A:$A,"<"&A2+1,'
    '\'异常扣款记录\'!$A:$A,"<>")',
)
self.assertEqual(history_sheet["J2"].value, "=G2-E2-I2")
```

- [ ] **Step 3: Add failing assertions for abnormal deduction sheet headers and formatting**

Add:

```python
deduction_sheet = workbook[ABNORMAL_DEDUCTION_SHEET]
self.assertEqual(
    [deduction_sheet.cell(row=1, column=column).value for column in range(1, 4)],
    ["扣款时间", "额度", "备注"],
)
self.assert_cell_has_thin_border(deduction_sheet["A1"])
self.assert_cell_has_thin_border(deduction_sheet["C20"])
self.assertEqual(deduction_sheet["A2"].number_format, "yyyy-mm-dd")
self.assertEqual(deduction_sheet["B2"].number_format, "0.00")
```

- [ ] **Step 4: Run the targeted test and confirm it fails**

Run:

```bash
/Users/beetle/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -B -m unittest tests.test_customer_history_workbook.CustomerHistoryWorkbookTest.test_history_workbook_adds_payment_sheet_and_balance_formulas
```

Expected: FAIL because `异常扣款记录` and new formula columns do not exist yet.

## Task 2: Implement Workbook Support

**Files:**
- Modify: `app/express_app/core/calculator.py`

- [ ] **Step 1: Add constants**

Add near the payment constants:

```python
CUSTOMER_ABNORMAL_DEDUCTION_SHEET = "异常扣款记录"
CUSTOMER_ABNORMAL_DEDUCTION_HEADERS = ["扣款时间", "额度", "备注"]
CUSTOMER_ABNORMAL_DEDUCTION_MIN_ROWS = 20
```

- [ ] **Step 2: Update `write_customer_history_sheet()` headers and formulas**

Set the header order to include `今日异常扣款` and `累计异常扣款` after `累计收款`. For each daily row, append:

```python
(
    f'=SUMIFS(\'{CUSTOMER_ABNORMAL_DEDUCTION_SHEET}\'!$B:$B,'
    f'\'{CUSTOMER_ABNORMAL_DEDUCTION_SHEET}\'!$A:$A,">="&A{row_number},'
    f'\'{CUSTOMER_ABNORMAL_DEDUCTION_SHEET}\'!$A:$A,"<"&A{row_number}+1)'
),
(
    f'=SUMIFS(\'{CUSTOMER_ABNORMAL_DEDUCTION_SHEET}\'!$B:$B,'
    f'\'{CUSTOMER_ABNORMAL_DEDUCTION_SHEET}\'!$A:$A,"<"&A{row_number}+1,'
    f'\'{CUSTOMER_ABNORMAL_DEDUCTION_SHEET}\'!$A:$A,"<>")'
),
f"=G{row_number}-E{row_number}-I{row_number}",
```

Update numeric formatting and fill ranges so columns 4 through 10 have money formatting, with income fill on column 6, deduction fill on column 8, and highlight fill on cumulative/ending balance columns 5, 7, 9, and 10.

- [ ] **Step 3: Add `write_customer_abnormal_deduction_sheet()`**

Implement a writer mirroring `write_customer_payment_sheet()` with three columns:

```python
def write_customer_abnormal_deduction_sheet(ws: openpyxl.worksheet.worksheet.Worksheet) -> None:
    is_empty_sheet = ws.max_row == 1 and ws.max_column == 1 and ws["A1"].value is None
    if is_empty_sheet:
        for column, header in enumerate(CUSTOMER_ABNORMAL_DEDUCTION_HEADERS, start=1):
            ws.cell(row=1, column=column).value = header
    elif [
        ws.cell(row=1, column=column).value
        for column in range(1, len(CUSTOMER_ABNORMAL_DEDUCTION_HEADERS) + 1)
    ] != CUSTOMER_ABNORMAL_DEDUCTION_HEADERS:
        ws.insert_rows(1)
        for column, header in enumerate(CUSTOMER_ABNORMAL_DEDUCTION_HEADERS, start=1):
            ws.cell(row=1, column=column).value = header

    style_history_header(ws)
    max_row = max(ws.max_row, CUSTOMER_ABNORMAL_DEDUCTION_MIN_ROWS)
    for row_number in range(2, max_row + 1):
        ws.cell(row=row_number, column=1).number_format = "yyyy-mm-dd"
        ws.cell(row=row_number, column=2).number_format = "0.00"
        ws.row_dimensions[row_number].height = 24
    apply_sheet_basics(ws)
    widths = {"A": 14, "B": 14, "C": 28}
    for column_letter, width in widths.items():
        ws.column_dimensions[column_letter].width = width
    ws.row_dimensions[1].height = 28
    apply_table_border(
        ws,
        max_row=max_row,
        max_column=len(CUSTOMER_ABNORMAL_DEDUCTION_HEADERS),
    )
```

- [ ] **Step 4: Add a preservation helper and call it after payments**

Add a helper that creates a sheet, copies an existing sheet if present, then applies the writer:

```python
def add_or_preserve_manual_customer_sheet(
    workbook: openpyxl.Workbook,
    existing_summary_paths: list[Path],
    sheet_name: str,
    writer: Callable[[openpyxl.worksheet.worksheet.Worksheet], None],
) -> None:
    target_sheet = workbook.create_sheet(sheet_name)
    for existing_summary_path in existing_summary_paths:
        if not existing_summary_path.exists():
            continue
        existing_workbook = openpyxl.load_workbook(existing_summary_path, data_only=False)
        try:
            if sheet_name in existing_workbook.sheetnames:
                copy_worksheet_contents(existing_workbook[sheet_name], target_sheet)
                break
        finally:
            existing_workbook.close()
    writer(target_sheet)
```

Then use it for `收款记录` and `异常扣款记录`, preserving the final sheet order.

- [ ] **Step 5: Run workbook tests**

Run:

```bash
/Users/beetle/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -B -m unittest tests.test_customer_history_workbook
```

Expected: PASS.

## Task 3: Preservation Test

**Files:**
- Modify: `tests/test_customer_history_workbook.py`

- [ ] **Step 1: Add failing preservation test**

Add a test that creates an existing customer history workbook with `异常扣款记录`, writes:

```python
["扣款时间", "额度", "备注"]
[date(2026, 4, 2), 88.8, "异常调整"]
```

After `build_customer_history_summary(...)`, assert the refreshed workbook still has the row:

```python
self.assert_cell_date(deduction_sheet["A2"].value, date(2026, 4, 2))
self.assertEqual(deduction_sheet["B2"].value, 88.8)
self.assertEqual(deduction_sheet["C2"].value, "异常调整")
```

- [ ] **Step 2: Run the test and confirm it fails if implementation is incomplete**

Run:

```bash
/Users/beetle/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -B -m unittest tests.test_customer_history_workbook.CustomerHistoryWorkbookTest.test_abnormal_deduction_records_are_preserved_when_history_is_refreshed
```

Expected before implementation: FAIL because the sheet is not preserved. Expected after Task 2: PASS.

## Task 4: Version Bump And Verification

**Files:**
- Modify: `app/express_app/version.py`
- Modify: `tests/test_branding_packaging.py`

- [ ] **Step 1: Update version test first**

Change expected values in `tests/test_branding_packaging.py`:

```python
self.assertEqual(APP_VERSION, "8.8.3")
self.assertEqual(APP_VERSION_LABEL, "V8.8.3")
self.assertEqual(OUTPUT_VERSION_SUFFIX, "v8_8_3")
```

- [ ] **Step 2: Run the version test and confirm failure**

Run:

```bash
/Users/beetle/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -B -m unittest tests.test_branding_packaging
```

Expected: FAIL while `APP_VERSION` is still `8.8.2`.

- [ ] **Step 3: Update version metadata**

Change `app/express_app/version.py`:

```python
APP_VERSION = "8.8.3"
```

- [ ] **Step 4: Run final targeted verification**

Run:

```bash
/Users/beetle/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -B -m unittest tests.test_customer_history_workbook tests.test_branding_packaging
```

Expected: all tests pass.

## Task 5: Account Balance Deduction

**Files:**
- Modify: `tests/test_v8_3_account_balance.py`
- Modify: `app/express_app/core/account_balance.py`

- [ ] **Step 1: Add failing account balance test**

Add `test_account_balance_deducts_abnormal_deduction_records` with one customer history workbook:

```python
self._write_history_workbook(
    customer_dir,
    "客户A",
    [(date(2026, 4, 2), 200.0)],
    [(date(2026, 4, 2), 500.0)],
    [(date(2026, 4, 2), 80.0, "破损扣款")],
)
```

Assert:

```python
self.assertAlmostEqual(record.total_consumed, 200.0)
self.assertAlmostEqual(record.total_paid, 500.0)
self.assertAlmostEqual(record.current_balance, 220.0)
self.assertAlmostEqual(dashboard.total_balance, 220.0)
```

- [ ] **Step 2: Run the test and confirm failure**

Run:

```bash
python -m unittest tests.test_v8_3_account_balance.AccountBalanceDashboardTest.test_account_balance_deducts_abnormal_deduction_records
```

Expected: FAIL with `300.0 != 220.0`, proving the old account-balance logic ignores abnormal deductions.

- [ ] **Step 3: Implement account-balance deduction**

In `account_balance.py`, import `CUSTOMER_ABNORMAL_DEDUCTION_SHEET`, add `_sum_abnormal_deduction_records()`, and calculate:

```python
total_abnormal_deducted = _sum_abnormal_deduction_records(workbook)
current_balance=round(total_paid - total_consumed - total_abnormal_deducted, 2)
```

Missing `异常扣款记录` must return `0.0` for old workbook compatibility.

- [ ] **Step 4: Run account balance and workbook tests**

Run:

```bash
python -m unittest tests.test_v8_3_account_balance tests.test_customer_history_workbook tests.test_branding_packaging
```

Expected: all tests pass.

## Self-Review

- Spec coverage: workbook sheet creation, formulas, preservation, account-balance deduction, order, and version bump are covered.
- Placeholder scan: no placeholders remain.
- Type consistency: helper signatures use existing `Path`, `openpyxl`, and `Callable` patterns; `Callable` must be imported from `collections.abc` or `typing` if not already available.
