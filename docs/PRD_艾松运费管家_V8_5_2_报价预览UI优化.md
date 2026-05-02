# 艾松运费管家 V8.5.2 报价预览 UI 优化 PRD

## 背景

V8.4 已经新增 `报价预览` 页面，用户可以按业务员查看客户报价表中的快递公司 sheet 和省份价格。当前页面功能可用，但业务员下拉框、快递公司标签和价格表滚动条仍偏基础控件风格。左侧导航中还保留了不可用的 `客户档案` 占位入口，容易让用户误以为该模块可用或即将使用。

V8.5.2 先做小范围 UI 优化，不进行全局 UI 重构，不引入新的 CustomTkinter 打包依赖。

## 目标

1. 版本号升级为 `V8.5.2`。
2. 左侧导航移除 `客户档案` 占位入口。
3. `报价预览` 页面增加专属 UI 样式：
   - 业务员下拉框更清晰。
   - 快递公司标签更突出当前选中状态。
   - 价格表行高、表头、选中态和滚动条更适合阅读。
4. 不影响 `费用计算`、`账户余额`、`系统设置` 现有布局。
5. 不引入 CustomTkinter 运行依赖，避免本次小改动牵动 macOS/Windows 打包链路。

## 设计说明

- 使用 `ttk.Style` 新增报价预览专属样式，避免改动全局 `Treeview` 影响结果页和账户余额页。
- 保留现有 `ttk.Notebook + ttk.Treeview` 结构，因为报价 sheet 数量不固定，当前结构对多快递公司更稳定。
- 将 CustomTkinter 作为后续全局 UI 升级参考，不在 V8.5.2 中强制接入。

## 验收标准

1. 主导航不再显示 `客户档案`。
2. `报价预览` 仍可同步报价表、选择业务员、搜索并展示快递公司价格表。
3. 报价预览控件使用专属样式常量，不影响其他页面表格。
4. 本地测试和打包自检通过。

## 测试计划

- `python -m unittest tests.test_v8_1_gui_navigation tests.test_branding_packaging tests.test_v8_4_price_preview_page`
- `python -m unittest discover tests`
- `python -m compileall -q app tests`
- `git diff --check`
