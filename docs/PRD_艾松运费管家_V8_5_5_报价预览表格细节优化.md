# 艾松运费管家 V8.5.5 报价预览表格细节优化 PRD

## 背景

V8.5.4 已经把报价预览表格改为左右两组 `省份 / 首重费用 / 续重费用`。继续查看时发现，列宽和对齐还需要更规整，左右两组之间也需要更清晰的视觉分隔。

## 目标

1. 版本升级为 `V8.5.5`。
2. `报价预览 -> 当前客户快递价格信息` 中，省份、首重费用、续重费用列等宽。
3. 表格中的实际数据居中展示。
4. 左三列和右三列之间增加一条实线分隔。

## 非目标

- 不修改报价读取逻辑。
- 不修改 6 列字段结构。
- 不修改计费、拆分和历史汇总逻辑。

## 验收标准

1. `APP_VERSION` 为 `8.5.5`，输出后缀为 `v8_5_5`。
2. 报价预览 6 个价格列宽度一致。
3. 报价预览价格数据居中显示。
4. 左右两组三列之间有可见实线分隔。

## 测试计划

- `python -m unittest tests.test_v8_4_price_preview_page tests.test_branding_packaging`
- `python -m unittest discover tests`
- `python -m compileall -q app tests`
- `git diff --check`
- `python3 -B app/build_macos_app.py`
- macOS `.app` 自检
