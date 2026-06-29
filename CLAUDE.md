# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

ExcelBridge — 可视化 Excel 数据批量匹配迁移工具。用点击选列替代 VLOOKUP 公式：从多个分散的 Excel 文件中按关键字段匹配并汇总数据到目标模板。基于 CustomTkinter 构建，目标平台 Windows 7/10/11。

## 运行与构建

```bash
# 开发运行
python main.py

# 打包为 Windows exe（Wine + Windows Python 3.8.10）
WINEPREFIX=~/.wine_py38 wine "C:\\Program Files\\Python38\\python.exe" -m PyInstaller ExcelBridge.spec

# 依赖安装（xlrd 版本锁定）
pip install -r requirements.txt
```

## 版本号

版本号唯一来源是 `_version.py`。标题栏、关于对话框均读取 `__version__`。构建时自动附加 git commit 短哈希和日期。修改版本只需改 `_version.py` 一行。

## 架构

| 文件 | 职责 |
|------|------|
| `main.py` | CustomTkinter GUI，三栏布局。左侧：文件夹选择+文件列表+规则存取。中间：源文件/目标模板上下分屏预览（ttk.Treeview）。右侧：操作指引（红色闪烁提示）+映射规则（2×2网格）+执行按钮+内嵌日志 |
| `_version.py` | 版本号唯一来源 |
| `rule_engine.py` | 核心匹配引擎：`ColumnMapping` 数据类、`preview_matches()` 扫描匹配、`build_target_index()` 目标索引、列标签解析、表头行检测 |
| `excel_reader.py` | Excel 读取：`.xlsx`（openpyxl）和 `.xls`（xlrd 1.2.0），隐藏行列过滤，ZIP 头检测伪装 xlsx |
| `excel_writer.py` | openpyxl 写入目标文件，按匹配键定位行写入数据，保留模板样式 |

默认窗口 1200×700，最小 1200×700（三栏：280px / stretch / 320px）。

## 匹配引擎核心逻辑（V1.3 重构）

### 用户操作流程（6 步）
1. 选择源文件夹和目标模板
2. 点击文件列表加载源文件
3. **强制点击行号 ○ 标记表头行**（未标记时操作指引卡片红色闪烁）
4. 点击数据格标记复制列（蓝）/ 匹配键（橙）— 标签从表头行自动提取
5. 在目标模板标记匹配键（绿）/ 粘贴列（紫）
6. 预览匹配结果 → 确认写入

### 标签驱动的跨文件列定位

用户标记列时，系统从表头行提取列标题作为标签（如"证件号"）。对每个源文件：
1. `_find_header_rows()` — 用标签在文件中定位表头行（标签优先，关键词兜底）
2. `_resolve_column()` — 在表头行中按标签定位列，支持**前缀匹配**（"证件号"→"证件号码" ✓，"证件号"→"签证证件号码" ✗）
3. 标签找不到 → 返回 -1 → 文件被跳过并提示"[列不匹配]"

### 目标文件处理

- 加载时自动检测表头行（关键词命中数最多的行）
- `build_target_index()` 构建索引时跳过表头行，防止表头值被写入匹配
- 用户标记目标列时，标签从目标首行（默认）提取

### 关键函数说明

| 函数 | 文件 | 说明 |
|------|------|------|
| `_find_header_rows()` | rule_engine.py | 标签驱动表头行检测，关键词仅兜底 |
| `_resolve_column()` | rule_engine.py | 精确+前缀匹配定位列，找不到返回 -1 |
| `build_target_index()` | rule_engine.py | 构建目标键→行号索引，跳过表头行 |
| `_is_header_key()` | rule_engine.py | 精确匹配 header_keys，不再使用 `_looks_like_header()` |
| `_get_col_header()` | main.py | 从表头行提取列标签，支持源/目标 |
| `_start_blink()` / `_stop_blink()` | main.py | 操作指引卡片红色闪烁（600ms 间隔） |

## Python 版本约束

- **3.8.10** — 最后支持 Win7 的版本。不能用 `match` 语句（3.10+），不能用 f-string 内同名引号
- **xlrd**：锁定 `1.2.0`，2.0+ 不再支持 `.xls`
- **openpyxl**：`data_only=True`，不能传 `read_only=True`（否则隐藏检测失效）
- **单元格值规范化**：`excel_reader.py` 和 `excel_writer.py` 各自有 `_normalize_cell()`，必须保持一致（浮点转整数字符串）

## PyInstaller 打包

- `ExcelBridge.spec` 已配置 hiddenimports（含 `_version`）、datas（`_icons`）、icon（`icon.ico`）
- `__file__` 在打包后指向 `%TEMP%\_MEIXXXX\`，`_config_dir()` 已处理此差异
- `console=False`，不显示命令行窗口

## 色彩系统

四色映射：🔵 蓝（复制列 `#1A56DB`）、🟠 橙（匹配键 `#D97706`）、🟢 绿（目标键 `#00544C`）、🟣 紫（粘贴至 `#8126D1`）

## 常见陷阱

- 字体缩放不生效 → `_font_var` 必须在 `_build()` 开头创建
- xls 隐藏行过滤无效 → 必须 `xlrd.open_workbook(filepath, formatting_info=True)`
- 按钮圆角断裂 → `border_width=0`
- Windows 7 高 DPI → `SetProcessDpiAwareness(2)` + `ctk.set_widget_scaling()`
- 表格列数固定 → `_make_table()` 需显式传 `col_count`
- 表头标记未设置 → `_src_header_row = -1`，用户必须点击行号 ○
