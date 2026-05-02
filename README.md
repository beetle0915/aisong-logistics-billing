# 艾松物流计费系统

艾松物流计费系统是一个本地桌面工具，用于按销售出库单、业务员、快递公司、重量和报价表自动计算快递费用，并生成总结果、客户每日明细和客户历史汇总。

## 功能

- 批量选择销售出库单。
- 读取业务员对应的快递报价表。
- 自动解析 `快递公司（标准版）`。
- 支持顺丰、德邦 20kg 以上大件模板。
- 生成总结果 Excel。
- 按业务员拆分客户每日明细。
- 生成客户历史汇总表，文件名带业务员名称，含收款记录和余额公式。
- 提供 Tkinter 桌面 GUI。
- 支持 macOS `.app` 打包。
- 支持 GitHub Actions 打包 Windows `.exe`。
- GitHub Actions 会在每次 push 后自动打包 Windows 版本，artifact 名称带版本号。

## 目录

```text
app/
  express_app/              # 应用源码
  build_macos_app.py        # macOS .app 打包脚本
  build_windows_exe.py      # Windows PyInstaller 打包脚本
  windows_entry.py          # Windows GUI 入口
  requirements-windows.txt  # Windows 打包依赖
docs/
  PRD_艾松物流计费系统_V7_0_1_品牌升级.md
  PRD_艾松物流计费系统_V7_0_2_打开目录修复.md
  PRD_艾松物流计费系统_V7_0_4_德邦费用四舍五入.md
  PRD_艾松物流计费系统_V7_0_5_收款记录和余额公式.md
  说明_当前快递费用计算逻辑.md
.github/workflows/
  build-windows.yml
```

## 本地运行

进入 `app` 目录：

```bash
cd app
python -m express_app
```

## macOS 打包

```bash
cd app
python -B build_macos_app.py
```

生成：

```text
艾松物流计费系统.app
```

## Windows 打包

推荐使用 GitHub Actions 打包。

上传到 GitHub 后，进入：

```text
Actions -> Build Windows EXE -> Run workflow
```

任务完成后下载 artifact：

```text
aisong-logistics-billing-v7.0.5-windows
```

其中包含：

```text
艾松物流计费系统.exe
```

## 不要上传的数据

以下内容包含业务数据或生成结果，不应上传 GitHub：

```text
原始数据表/
快递报价表/
输出结果/
客户每日快递费明细/
*.xlsx
```
