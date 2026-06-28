# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

ExcelBridge — 可视化 Excel 数据批量匹配迁移工具。用 4 次点击替代 VLOOKUP 公式：从多个分散的 Excel 文件中按关键字段匹配并汇总数据到目标模板。基于 CustomTkinter 构建，目标平台 Windows 7/10/11。

## 运行与构建

```bash
# 开发运行
python main.py

# 打包为 Windows exe（必须在 Wine + Windows Python 3.8.10 中执行）
WINEPREFIX=~/.wine_py38 wine "C:\\Program Files\\Python38\\python.exe" -m PyInstaller ExcelBridge.spec

# 依赖安装（注意 xlrd 版本锁定）
pip install -r requirements.txt
```

## 架构

4 个模块，依赖关系：`main.py` → `rule_engine.py` → `excel_reader.py` / `excel_writer.py`

| 文件 | 职责 |
|------|------|
| `main.py` | CustomTkinter GUI 应用，三栏布局：左侧栏（文件夹选择+文件列表+规则存取）、中间面板（源文件/目标模板上下分屏预览）、右侧栏（操作指引+映射规则+执行按钮+内嵌日志）。自定义标题栏（Windows overrideredirect），边缘拖拽调整大小 |
| `excel_reader.py` | 读取 `.xlsx`（openpyxl）和 `.xls`（xlrd 1.2.0），支持隐藏行列过滤。检测 ZIP 头签名处理伪装成 `.xls` 的实际 xlsx 文件 |
| `excel_writer.py` | 用 openpyxl 写入目标文件，按匹配键定位行并写入数据，保留模板样式，输出文件名加时间戳后缀 |
| `rule_engine.py` | 核心逻辑：`ColumnMapping` 定义四列映射规则（源复制列/源匹配键/目标匹配键/目标粘贴列），`preview_matches()` 扫描所有源文件生成预览结果，智能跳过表头行，`save_rule()`/`load_rule()` 存取 JSON 规则 |

## 关键约束

- **Python 版本**：3.8.10 — 最后支持 Win7 的版本。不能用 `match` 语句（3.10+），不能用 f-string 内同名引号
- **xlrd 版本**：锁定 `1.2.0`，2.0+ 不再支持 `.xls` 格式
- **openpyxl**：读取时传 `data_only=True` 但**不能**传 `read_only=True`，否则 `row_dimensions`/`column_dimensions` 为空，隐藏检测失效
- **单元格值规范化**：`excel_reader.py` 和 `excel_writer.py` 各自有 `_normalize_cell()`，必须保持一致（纯数字浮点转整数字符串），否则键匹配会悄无声息失败
- **PyInstaller 打包**：`__file__` 在打包后指向临时目录 `%TEMP%\_MEIXXXX\`，持久化路径需用 `sys.executable` 所在目录。`_config_dir()` 已处理此差异

## 色彩系统

四色映射视觉语言（类 `C` 中定义）：

| 颜色 | 色值 | 语义 |
|------|------|------|
| 🔵 蓝色 | `#3B82F6` / `#EFF6FF` | 源文件「要复制的值」|
| 🟠 橙色 | `#F59E0B` / `#FFFBEB` | 源文件「匹配键」|
| 🟢 绿色 | `#10B981` / `#ECFDF5` | 目标「匹配键」|
| 🟣 紫色 | `#8B5CF6` / `#F5F3FF` | 目标「粘贴至」|

主品牌蓝 `#2563EB`，主背景 `#F7F9FC`，卡片白 `#FFFFFF`。

## 文件结构

```
excel-to/
├── main.py              # GUI 主入口 (1575行)
├── excel_reader.py      # Excel 读取 + 文件扫描
├── excel_writer.py      # Excel 写入引擎
├── rule_engine.py       # 规则引擎 + 预览匹配
├── requirements.txt     # Python 依赖
├── DESIGN_BRIEF.md      # 设计概要
├── LESSONS.md           # 开发避坑手册（必读）
├── _icons/              # PNG 图标（19个，含 emoji 渲染图标）
├── icon.ico / icon.png  # 应用图标
└── .excelbridge_config.json  # 用户配置（字号、跳过隐藏）
```

## 常见陷阱速查

- 字体缩放不生效 → `_font_var` 必须在 `_build()` 开头创建，不能在 `_build_center()` 中才创建
- xls 隐藏行过滤无效 → 必须 `xlrd.open_workbook(filepath, formatting_info=True)`
- 按钮圆角断裂 → `border_width` 和 `corner_radius` 冲突，设 `border_width=0`
- 配置不保存 → 检查是否用了 `__file__` 而非 `sys.executable` 做配置目录
- Windows 7 高 DPI 字体极小 → 需要 `SetProcessDpiAwareness(2)` + `ctk.set_widget_scaling(_dpi_scale)`
- 表格列数固定 → `_make_table()` 需要显式传 `col_count` 参数

更多细节见 `LESSONS.md`。
