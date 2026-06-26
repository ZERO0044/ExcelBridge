# ExcelBridge 开发避坑手册

从零到 V1.0 交付，记录所有踩过的坑，便于日后维护和新项目参考。

---

## 一、环境与兼容性

### 1.1 Python 版本选型
- **坑**：最新 Python 3.14 不支持 Windows 7，需要选择 Python 3.8.10（最后支持 Win7 的版本）
- **表现**：Win7 上 pyinstaller 打出的 exe 直接闪退或无响应
- **解法**：`Wine + Python 3.8.10` 交叉编译，Wine 环境独立于系统 Python（`~/.wine_py38`）
- **教训**：先确认目标平台最低要求，再选技术栈。Win7 是 2026 年仍在流通的系统

### 1.2 PyInstaller 交叉编译
- **坑**：Linux 上 `pyinstaller` 生成 Linux ELF，不是 Windows PE
- **解法**：必须用 Wine 里的 Windows Python 跑 PyInstaller
  ```bash
  WINEPREFIX=~/.wine_py38 wine "C:\\Program Files\\Python38\\python.exe" -m PyInstaller ExcelBridge.spec
  ```
- **教训**：PyInstaller 不是交叉编译器，打包目标平台 = 运行 PyInstaller 的平台

### 1.3 pip 依赖版本锁定
- **坑**：xlrd 2.0+ 不再支持 `.xls` 格式（仅支持 `.xlsx`）
- **解法**：`requirements.txt` 中锁定 `xlrd==1.2.0`
- **教训**：第三方库的大版本更新可能砍掉核心功能，必须锁定版本

### 1.4 Windows 7 DPI 缩放
- **坑**：Win7 高 DPI 屏幕上字体和控件极小
- **解法**：启动时声明 DPI 感知 + 获取系统缩放比 + `ctk.set_widget_scaling()`
  ```python
  ctypes.windll.shcore.SetProcessDpiAwareness(2)  # 声明支持
  _dpi_scale = GetDeviceCaps(hdc, LOGPIXELSY) / 96  # 获取实际缩放
  ctk.set_widget_scaling(_dpi_scale)  # 应用缩放
  ```
- **教训**：DPI 感知需要两步——声明 + 实际缩放，只声明不够

### 1.5 emoji 字体跨平台
- **坑**：Linux 上 emoji 渲染正常，Windows 7 上全部变成方框（无 emoji 字体）
- **解法**：将 emoji 渲染为 PNG 文件（Pillow + Noto Color Emoji 字体），存入 `_icons/` 目录
- **教训**：永远不要依赖系统自带的 emoji 字体


## 二、openpyxl 的坑

### 2.1 `read_only=True` 导致隐藏检测失效
- **坑**：`load_workbook(filepath, data_only=True, read_only=True)` 在只读模式下不加载 `row_dimensions` 和 `column_dimensions`
- **表现**：`ws.row_dimensions` 永远是空字典，隐藏行过滤完全不起作用
- **解法**：去掉 `read_only=True`，改为 `load_workbook(filepath, data_only=True)`
- **教训**：`read_only=True` 是流式模式，只适合纯数据读取，不适合需要元数据的场景

### 2.2 键值类型不一致导致匹配失败
- **坑**：`excel_reader.py` 对纯数字 `12345.0` 规范化为 `"12345"`，但 `excel_writer.py` 用 `str(cell.value)` 得到 `"12345.0"`
- **表现**：证件号是纯数字时，源和目标永远匹配不上，数据悄无声息丢失
- **解法**：两处使用相同的 `_normalize_cell()` 函数，统一处理 `int(value)` 逻辑
- **教训**：跨模块的字符串规范化必须共用同一函数，不能各自写各自

### 2.3 openpyxl 按扩展名拒绝文件
- **坑**：很多 `.xls` 文件实际是 xlsx 格式（ZIP 压缩），但 openpyxl 按扩展名拒绝打开
- **表现**：`openpyxl.utils.exceptions.InvalidFileException: openpyxl does not support the old .xls file format`
- **解法**：检测文件头 2 字节——`PK` = ZIP/xlsx → 临时复制为 `.xlsx` 再用 openpyxl 打开
- **教训**：不要相信扩展名，文件头签名才是真相


## 三、xlrd 的坑

### 3.1 `formatting_info=True` 是隐藏检测的前提
- **坑**：`xlrd.open_workbook(filepath)` 不传 `formatting_info=True` 时，ROW 记录完全不解析
- **表现**：`rowinfo_map` 永远为空，`.xls` 文件跳过隐藏完全失效
- **解法**：`xlrd.open_workbook(filepath, formatting_info=True)` —— 必须传！
- **教训**：xlrd 默认不解析格式记录，需要显式开启

### 3.2 xlwt 不设隐藏标志
- **坑**：用 `xlwt` 生成的 `.xls` 测试文件，`row.hidden = True` 不生效——ROW 记录的 flags 字段为 0x0000
- **表现**：拿 xlwt 生成的测试文件验证隐藏检测功能，结果永远是"没有隐藏行"，误导排查方向
- **解法**：用真实的 WPS/Excel 创建测试文件，或手动修正 BIFF 二进制中的 flags 位
- **教训**：测试数据必须用目标用户的真实软件生成，不能用同生态的库自产自测

### 3.3 xlrd 的键类型
- **坑**：xlrd 单元格类型判断用 `.ctype` 而非 `.type`，数值用 `XL_CELL_NUMBER` (2)
- **教训**：读文档比猜 API 快

### 3.4 OLE 文件结构
- **坑**：BIFF `.xls` 是 OLE 复合文档，数据跨扇区分散。原始字节扫描可能找不到 ROW 记录
- **解法**：用 xlrd 的 `formatting_info=True` 就够了，不需要自己解析 OLE


## 四、CustomTkinter 的坑

### 4.1 ttk.Treeview 与 CTk 渲染引擎不兼容
- **坑**：CustomTkinter 没有原生表格组件，只能用 ttk.Treeview。两者渲染管线完全不同
- **表现**：Treeview 永远直角、原生滚动条、颜色微妙偏差，与周围的圆角 CTk 控件格格不入
- **现状**：无完美解法。CTkTable 无虚拟滚动（大数据卡死），tksheet 样式系统独立（需要桥接层）
- **教训**：选 GUI 框架时先看有没有你需要的核心组件类型

### 4.2 ttk.Style 配置必须在 root 创建后
- **坑**：在 `CTk()` 实例化之前配置 `ttk.Style()` 会不生效或被覆盖
- **解法**：`_setup_styles()` 在 `__init__` 的 `super().__init__()` 之后调用

### 4.3 CTkOptionMenu 的 `command` 回调
- **坑**：`command` 在用户选择时触发，但 `variable.set()` 也可能触发（取决于版本），导致 `_build()` 被递归调用
- **解法**：在 `_build()` 中先设置 `_font_var`，再在 `_build_center()` 中判断 `hasattr` 避免重建

### 4.4 `_on_font_change` 触发 `_build()` 重建的时序问题
- **坑**：`_build()` 依次调用 `_build_left()` → `_build_center()` → `_build_right()`，但 `_font_var` 在 `_build_center()` 中才创建
- **表现**：`_build_left()` 和部分 `_build_right()` 读不到正确的字号的当前值，只能读到过期的 `_saved_font_size`
- **解法**：在 `_build()` 开头立即创建 `_font_var`，确保三个 `_build_*` 一致

### 4.5 无虚拟滚动的大数据表格
- **坑**：`preview_dialog.py` 用 CTkScrollableFrame + CTkFrame 逐行绘制，每行一个 widget
- **表现**：几百行就开始卡，上千行直接卡死
- **教训**：Python GUI 中每行一个 widget 只适合 <100 行

### 4.6 CTkTextbox 不支持富文本标签
- **坑**：CTkTextbox 没有直接的 `tag_config(foreground=...)` 像 tkinter Text 那样自由
- **实际**：CTkTextbox 继承自 tkinter Text，可以使用标准的 `tag_config`/`tag_add`（文档不明确但确实可用）
- **解法**：直接用 tkinter Text tag API

### 4.7 `iconbitmap` 与 `iconphoto` 的跨平台差异
- **坑**：Windows 用 `iconbitmap(.ico)`，Linux 用 `iconphoto(.png)`，但是 2048px 的 png 在 X11 上会报错
- **解法**：PIL thumbnail 缩小到 64x64
- **教训**：图标尺寸别太大，64px 足够


## 五、Python 3.8 兼容性

### 5.1 f-string 中不能有同名引号
- **坑**：`f"{val.split(" ",1)[0]}"` — Python 3.8 不支持 f-string 内的双引号
- **解法**：改用单引号 `f"{val.split(' ',1)[0]}"`
- **教训**：Python 3.12+ 放宽了此限制，但 3.8 不行

### 5.2 不能使用 `match` 语句 (3.10+)
- 项目中未踩，但需注意

### 5.3 不能使用海象运算符 `:=` (3.8+)
- 项目中未踩，但需注意


## 六、Git 与配置持久化

### 6.1 PyInstaller 中 `__file__` 指向临时目录
- **坑**：`os.path.dirname(os.path.abspath(__file__))` 在 PyInstaller 打包后指向 `%TEMP%\_MEIXXXX\`，程序关闭即删除
- **表现**：Windows 7 上修改字号/跳过隐藏后，下次打开恢复默认——配置文件被写进了临时目录
- **解法**：区分场景，打包后用 `sys.executable` 所在目录
  ```python
  def _config_dir(self):
      if getattr(sys, 'frozen', False):
          return os.path.dirname(os.path.abspath(sys.executable))
      return os.path.dirname(os.path.abspath(__file__))
  ```
- **教训**：所有持久化路径都要考虑 PyInstaller 的临时目录陷阱

### 6.2 Token 安全存储
- **坑**：`git push` 时 token 不应该写在 remote URL 里（明码留在 `.git/config`）
- **解法**：`git config credential.helper store` + `~/.git-credentials` (权限 600)


## 七、测试经验

### 7.1 不能用同生态库自产自测
- **坑**：用 xlwt 生成 .xls → 用 xlrd 读取测试隐藏行 → 全部失败（xlwt 不写 hidden flags）
- **教训**：测试数据必须来自真实用户场景（WPS/Excel 手动创建）

### 7.2 文件签名检测比扩展名可靠
- **坑**：`.xls` 文件可能是 xlsx 伪装（ZIP 头 PK），也可能是真 BIFF（OLE 头 D0CF11E0）
- **解法**：读前 4 字节判断

### 7.3 Wine 环境测试的局限性
- **坑**：Wine 的 DPI 报告不一定等于真实 Windows 的 DPI
- **教训**：DPI 相关功能必须在真机上验证

### 7.4 空 rowinfo_map 不代表没有隐藏行
- **坑**：`rowinfo_map = {}` 可能是文件确实没有隐藏行，也可能是 xlrd 没解析（没传 formatting_info）
- **教训**：先排除库配置问题，再下"没有隐藏行"的结论


## 八、架构决策回顾

| 决策 | 选型 | 为什么 |
|------|------|--------|
| GUI 框架 | CustomTkinter | 比 tkinter 原生好看，比 PyQt 轻量（Win7 兼容） |
| 表格组件 | ttk.Treeview | 唯一支持虚拟滚动的方案 |
| xlsx 读写 | openpyxl | 保留样式，隐藏检测完整 |
| xls 读取 | xlrd 1.2.0 | 最后支持 .xls 的版本 |
| 打包 | PyInstaller | 单文件 exe，12MB 合理 |
| 图标 | PNG 图片 | 避免依赖系统 emoji 字体 |
| 配置存储 | JSON 文件 | 简单，无需数据库 |

---

## 九、常见错误速查

| 错误信息 | 原因 | 修复 |
|----------|------|------|
| emoji 变方框 | Win7 无 emoji 字体 | 用 PNG 替代 |
| 字体缩放部分生效 | `_font_var` 创建在 `_build_center()` | 提前到 `_build()` 开头 |
| xls 跳隐藏不生效 | 没传 `formatting_info=True` | 加参数 |
| 纯数字键匹配失败 | reader/writer 规范化不一致 | 共用 `_normalize_cell` |
| 配置不保存 | `__file__` 指向 tmp | 用 `sys.executable` |
| exe 字体极小 | Win7 DPI 没缩放 | `ctk.set_widget_scaling()` |
| 列数固定 10 列 | `_make_table` 重建时丢参数 | 传 `col_count` 参数 |
| 按钮圆角断裂 | `border_width` + `corner_radius` 冲突 | 去掉 border_width |
