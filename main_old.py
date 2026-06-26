#!/usr/bin/env python3
"""
Excel 单元格批量迁移工具
CustomTkinter 现代 UI · VSCode/Notion 风格布局
Windows 7 兼容
"""

import os
import sys
import threading

# ═══ Windows 7 DPI 兼容补丁 ═══
import ctypes
if sys.platform == 'win32':
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinter as tk

from excel_grid import ExcelGrid, apply_light_style
from excel_reader import read_excel, find_files, get_relative_path
from rule_engine import (
    ColumnMapping, preview_matches, save_rule, load_rule,
)
from excel_writer import process_and_write
from preview_dialog import PreviewDialog

# ── 配色 (现代亮色主题) ──
ACCENT_BLUE    = '#2563EB'
ACCENT_GREEN   = '#16A34A'
ACCENT_ORANGE  = '#EA580C'
ACCENT_YELLOW  = '#CA8A04'
DANGER_RED     = '#DC2626'
TEXT_PRIMARY   = '#1E1E1E'
TEXT_SECONDARY = '#6B7280'
TEXT_MUTED     = '#9CA3AF'
BG_MAIN        = '#F3F4F6'
BG_SIDEBAR     = '#FFFFFF'
BG_CARD        = '#FFFFFF'
BG_INPUT       = '#F9FAFB'
BORDER         = '#D1D5DB'
STEP_ACTIVE    = '#2563EB'
STEP_DONE      = '#16A34A'
STEP_PENDING   = '#D1D5DB'


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Excel 单元格批量迁移工具")
        self.geometry("1280x820")
        self.minsize(1000, 650)

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        apply_light_style()

        self.configure(fg_color=BG_MAIN)

        # 状态
        self.source_dir = ""
        self.target_path = ""
        self.source_files = []
        self.current_source_file = ""
        self.mapping = ColumnMapping()
        self.target_sheet = ""
        self.source_sheet = ""
        self._source_data_cache = {}
        self._src_click_count = 0
        self._tgt_click_count = 0
        self._file_buttons = {}

        self._build_ui()
        self._log("欢迎使用 Excel 单元格批量迁移工具")

    # ═══ UI 构建 ═══

    def _build_ui(self):
        # 主窗口使用 grid 布局，确保底部栏始终可见
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # ── 顶部步骤条 ──
        self.step_bar = ctk.CTkFrame(self, fg_color=BG_SIDEBAR, height=48,
                                      corner_radius=0, border_width=0,
                                      border_color=BORDER)
        self.step_bar.grid(row=0, column=0, sticky='ew')
        self.step_bar.grid_propagate(False)

        step_container = ctk.CTkFrame(self.step_bar, fg_color='transparent')
        step_container.pack(expand=True, pady=8)

        self._step_circles = []
        self._step_labels = []
        steps = [("1", "选择文件"), ("2", "定义映射"), ("3", "预览执行")]
        for i, (num, text) in enumerate(steps):
            cf = ctk.CTkFrame(step_container, fg_color='transparent')
            cf.pack(side='left')

            circle = ctk.CTkFrame(cf, fg_color=STEP_PENDING,
                                   width=28, height=28, corner_radius=14)
            circle.pack(side='left', padx=(0, 6))
            circle.pack_propagate(False)
            cl = ctk.CTkLabel(circle, text=num, font=('sans-serif', 12, 'bold'),
                              text_color='white')
            cl.place(relx=0.5, rely=0.5, anchor='center')

            tl = ctk.CTkLabel(cf, text=text, font=('sans-serif', 11),
                              text_color=TEXT_SECONDARY)
            tl.pack(side='left')

            self._step_circles.append(circle)
            self._step_labels.append(tl)

            if i < len(steps) - 1:
                sep = ctk.CTkFrame(step_container, fg_color=BORDER,
                                    width=50, height=2)
                sep.pack(side='left', padx=6)

        self._update_steps(0)

        # ── 主体 (grid row 1, expand) ──
        body = ctk.CTkFrame(self, fg_color='transparent')
        body.grid(row=1, column=0, sticky='nsew', padx=8, pady=(0, 4))

        # 侧边栏
        self._build_sidebar(body)

        # 分隔线
        ctk.CTkFrame(body, fg_color=BORDER, width=1).pack(
            side='left', fill='y', padx=(0, 8))

        # 右侧内容区
        content = ctk.CTkFrame(body, fg_color='transparent')
        content.pack(side='left', fill='both', expand=True)
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(0, weight=55)
        content.grid_rowconfigure(1, weight=45)

        self._build_source_card(content)
        self._build_target_card(content)

        # ── 底部操作栏 (grid row 2, 始终可见) ──
        self._build_bottom_bar()

    def _update_steps(self, active):
        for i, (circle, label) in enumerate(zip(self._step_circles, self._step_labels)):
            if i < active:
                circle.configure(fg_color=STEP_DONE)
                label.configure(text_color=ACCENT_GREEN)
            elif i == active:
                circle.configure(fg_color=STEP_ACTIVE)
                label.configure(text_color=ACCENT_BLUE)
            else:
                circle.configure(fg_color=STEP_PENDING)
                label.configure(text_color=TEXT_SECONDARY)

    def _build_sidebar(self, parent):
        sidebar = ctk.CTkFrame(parent, fg_color='transparent', width=260)
        sidebar.pack(side='left', fill='both')
        sidebar.pack_propagate(False)

        # ── 文件夹选择 ──
        ctk.CTkLabel(sidebar, text="📂 源文件夹",
                     font=('sans-serif', 12, 'bold'),
                     text_color=TEXT_PRIMARY).pack(anchor='w', pady=(0, 4))

        dir_row = ctk.CTkFrame(sidebar, fg_color='transparent')
        dir_row.pack(fill='x', pady=(0, 6))
        self.source_dir_var = ctk.StringVar()
        ctk.CTkEntry(dir_row, textvariable=self.source_dir_var, height=28,
                     fg_color=BG_INPUT, border_color=BORDER,
                     corner_radius=6).pack(side='left', fill='x', expand=True,
                                            padx=(0, 4))
        ctk.CTkButton(dir_row, text="...", width=30, height=28,
                      fg_color=BG_INPUT, text_color=TEXT_PRIMARY,
                      border_width=1, border_color=BORDER, corner_radius=6,
                      hover_color='#E5E7EB',
                      command=self._on_select_source_dir).pack(side='right')

        ctk.CTkLabel(sidebar, text="📋 目标模板",
                     font=('sans-serif', 12, 'bold'),
                     text_color=TEXT_PRIMARY).pack(anchor='w', pady=(8, 4))

        tgt_row = ctk.CTkFrame(sidebar, fg_color='transparent')
        tgt_row.pack(fill='x', pady=(0, 6))
        self.target_path_var = ctk.StringVar()
        ctk.CTkEntry(tgt_row, textvariable=self.target_path_var, height=28,
                     fg_color=BG_INPUT, border_color=BORDER,
                     corner_radius=6).pack(side='left', fill='x', expand=True,
                                            padx=(0, 4))
        ctk.CTkButton(tgt_row, text="...", width=30, height=28,
                      fg_color=BG_INPUT, text_color=TEXT_PRIMARY,
                      border_width=1, border_color=BORDER, corner_radius=6,
                      hover_color='#E5E7EB',
                      command=self._on_select_target).pack(side='right')

        # ── 文件列表 ──
        list_header = ctk.CTkFrame(sidebar, fg_color='transparent')
        list_header.pack(fill='x', pady=(12, 4))
        ctk.CTkLabel(list_header, text="📄 源文件",
                     font=('sans-serif', 12, 'bold'),
                     text_color=TEXT_PRIMARY).pack(side='left')
        ctk.CTkButton(list_header, text="🔄", width=26, height=22,
                      fg_color='transparent', text_color=TEXT_SECONDARY,
                      hover_color='#E5E7EB',
                      command=self._on_rescan).pack(side='right')

        self.file_scroll = ctk.CTkScrollableFrame(sidebar,
                                                    fg_color='#F9FAFB',
                                                    corner_radius=8,
                                                    border_width=1,
                                                    border_color=BORDER,
                                                    height=180)
        self.file_scroll.pack(fill='x', pady=(0, 12))

        # ── 映射状态卡 ──
        ctk.CTkLabel(sidebar, text="📐 映射规则",
                     font=('sans-serif', 12, 'bold'),
                     text_color=TEXT_PRIMARY).pack(anchor='w', pady=(0, 4))

        self.map_card = ctk.CTkFrame(sidebar, fg_color='#F9FAFB',
                                      corner_radius=8, border_width=1,
                                      border_color=BORDER)
        self.map_card.pack(fill='x', pady=(0, 8))

        self.map_text_var = ctk.StringVar(value="尚未定义映射")
        self.map_label = ctk.CTkLabel(self.map_card,
                                       textvariable=self.map_text_var,
                                       font=('monospace', 10),
                                       text_color=TEXT_SECONDARY,
                                       justify='left', anchor='w')
        self.map_label.pack(fill='x', padx=10, pady=8)

        self._update_mapping_display()

        # 操作按钮
        btn_row = ctk.CTkFrame(sidebar, fg_color='transparent')
        btn_row.pack(fill='x', pady=(0, 4))
        ctk.CTkButton(btn_row, text="清除", width=80, height=28,
                      fg_color='transparent', text_color=DANGER_RED,
                      border_width=1, border_color=BORDER, corner_radius=6,
                      hover_color='#FEE2E2',
                      command=self._on_clear_mapping).pack(side='left')
        ctk.CTkButton(btn_row, text="保存规则", height=28,
                      fg_color='transparent', text_color=TEXT_SECONDARY,
                      border_width=1, border_color=BORDER, corner_radius=6,
                      hover_color='#E5E7EB',
                      command=self._on_save_rule).pack(side='right')
        ctk.CTkButton(sidebar, text="📂 加载规则", height=28,
                      fg_color='transparent', text_color=TEXT_SECONDARY,
                      border_width=1, border_color=BORDER, corner_radius=6,
                      hover_color='#E5E7EB',
                      command=self._on_load_rule).pack(fill='x')

    def _build_source_card(self, parent):
        card = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=10,
                             border_width=1, border_color=BORDER)
        card.grid(row=0, column=0, sticky='nsew', pady=(0, 4))
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(1, weight=0)
        card.grid_rowconfigure(2, weight=1)

        # 标题栏
        title_bar = ctk.CTkFrame(card, fg_color='transparent')
        title_bar.grid(row=0, column=0, sticky='ew', padx=12, pady=(10, 0))
        ctk.CTkLabel(title_bar, text="源文件预览",
                     font=('sans-serif', 12, 'bold'),
                     text_color=TEXT_PRIMARY).pack(side='left')
        ctk.CTkLabel(title_bar, text="点击单元格进行选择",
                     font=('sans-serif', 10), text_color=TEXT_MUTED).pack(side='right')

        # 控制栏
        ctrl = ctk.CTkFrame(card, fg_color='transparent')
        ctrl.grid(row=1, column=0, sticky='ew', padx=12, pady=(4, 4))
        ctk.CTkLabel(ctrl, text="Sheet:", font=('sans-serif', 10),
                     text_color=TEXT_SECONDARY).pack(side='left', padx=(0, 6))
        self.src_sheet_var = ctk.StringVar()
        self.src_sheet_combo = ctk.CTkComboBox(ctrl, width=130, height=24,
                                                variable=self.src_sheet_var,
                                                font=('sans-serif', 10),
                                                fg_color=BG_INPUT,
                                                border_color=BORDER,
                                                button_color=ACCENT_BLUE,
                                                dropdown_fg_color=BG_CARD,
                                                dropdown_text_color=TEXT_PRIMARY,
                                                state='readonly',
                                                command=self._on_source_sheet_change)
        self.src_sheet_combo.pack(side='left')

        self.src_status_badge = ctk.CTkFrame(ctrl, fg_color='#E5E7EB',
                                               corner_radius=4, height=20)
        self.src_status_badge.pack(side='right')
        self.src_status_var = ctk.StringVar(value="等待选择")
        ctk.CTkLabel(self.src_status_badge, textvariable=self.src_status_var,
                     font=('sans-serif', 9), text_color=TEXT_SECONDARY).pack(
            padx=8)

        # Grid
        gf = ctk.CTkFrame(card, fg_color=BG_MAIN, corner_radius=6)
        gf.grid(row=2, column=0, sticky='nsew', padx=8, pady=(0, 8))
        self.src_grid = ExcelGrid(gf)
        self.src_grid.pack(fill='both', expand=True)
        self.src_grid.on_cell_click = self._on_source_cell_click

    def _build_target_card(self, parent):
        card = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=10,
                             border_width=1, border_color=BORDER)
        card.grid(row=1, column=0, sticky='nsew', pady=(4, 0))
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(1, weight=0)
        card.grid_rowconfigure(2, weight=1)

        title_bar = ctk.CTkFrame(card, fg_color='transparent')
        title_bar.grid(row=0, column=0, sticky='ew', padx=12, pady=(10, 0))
        ctk.CTkLabel(title_bar, text="目标模板预览",
                     font=('sans-serif', 12, 'bold'),
                     text_color=TEXT_PRIMARY).pack(side='left')
        ctk.CTkLabel(title_bar, text="点击单元格进行选择",
                     font=('sans-serif', 10), text_color=TEXT_MUTED).pack(side='right')

        ctrl = ctk.CTkFrame(card, fg_color='transparent')
        ctrl.grid(row=1, column=0, sticky='ew', padx=12, pady=(4, 4))
        ctk.CTkLabel(ctrl, text="Sheet:", font=('sans-serif', 10),
                     text_color=TEXT_SECONDARY).pack(side='left', padx=(0, 6))
        self.tgt_sheet_var = ctk.StringVar()
        self.tgt_sheet_combo = ctk.CTkComboBox(ctrl, width=130, height=24,
                                                variable=self.tgt_sheet_var,
                                                font=('sans-serif', 10),
                                                fg_color=BG_INPUT,
                                                border_color=BORDER,
                                                button_color=ACCENT_BLUE,
                                                dropdown_fg_color=BG_CARD,
                                                dropdown_text_color=TEXT_PRIMARY,
                                                state='readonly',
                                                command=self._on_target_sheet_change)
        self.tgt_sheet_combo.pack(side='left')

        self.tgt_status_badge = ctk.CTkFrame(ctrl, fg_color='#E5E7EB',
                                               corner_radius=4, height=20)
        self.tgt_status_badge.pack(side='right')
        self.tgt_status_var = ctk.StringVar(value="等待选择")
        ctk.CTkLabel(self.tgt_status_badge, textvariable=self.tgt_status_var,
                     font=('sans-serif', 9), text_color=TEXT_SECONDARY).pack(
            padx=8)

        gf = ctk.CTkFrame(card, fg_color=BG_MAIN, corner_radius=6)
        gf.grid(row=2, column=0, sticky='nsew', padx=8, pady=(0, 8))
        self.tgt_grid = ExcelGrid(gf)
        self.tgt_grid.pack(fill='both', expand=True)
        self.tgt_grid.on_cell_click = self._on_target_cell_click

    def _build_bottom_bar(self):
        """底部操作栏 — 始终可见，使用 grid 固定在第2行"""
        bottom = ctk.CTkFrame(self, fg_color=BG_SIDEBAR, corner_radius=0,
                               border_width=1, border_color=BORDER)
        bottom.grid(row=2, column=0, sticky='ew')

        # 按钮行
        inner = ctk.CTkFrame(bottom, fg_color='transparent')
        inner.pack(fill='x', padx=12, pady=(8, 4))

        ctk.CTkButton(inner, text="🔍 预览匹配结果", height=34,
                      fg_color=ACCENT_BLUE, text_color='white',
                      hover_color='#1D4ED8', corner_radius=8, width=150,
                      font=('sans-serif', 12),
                      command=self._on_preview).pack(side='left', padx=(0, 8))

        ctk.CTkButton(inner, text="▶ 开始批量处理", height=34,
                      fg_color=ACCENT_GREEN, text_color='white',
                      hover_color='#15803D', corner_radius=8, width=150,
                      font=('sans-serif', 12),
                      command=self._on_process).pack(side='left')

        # 日志按钮 + 状态
        self.progress_var = ctk.StringVar(value="就绪")
        ctk.CTkLabel(inner, textvariable=self.progress_var,
                     font=('sans-serif', 10),
                     text_color=TEXT_MUTED).pack(side='right', padx=(8, 0))
        ctk.CTkButton(inner, text="📋 日志", height=28, width=70,
                      fg_color='transparent', text_color=TEXT_SECONDARY,
                      border_width=1, border_color=BORDER, corner_radius=6,
                      font=('sans-serif', 10),
                      hover_color='#E5E7EB',
                      command=self._on_show_log).pack(side='right')

        # 分隔
        ctk.CTkFrame(bottom, fg_color=BORDER, height=1).pack(
            fill='x', padx=12)

        # 嵌入式日志（精简）
        self.log_text = ctk.CTkTextbox(bottom, height=60,
                                        fg_color=BG_SIDEBAR,
                                        text_color=TEXT_SECONDARY,
                                        font=('monospace', 9),
                                        border_width=0,
                                        corner_radius=0)
        self.log_text.pack(fill='x', padx=12, pady=(4, 8))

        sys.stdout = _RedirectText(self.log_text)

    # ═══ 日志 ═══

    def _log(self, msg):
        self.log_text.insert('end', msg + '\n')
        self.log_text.see('end')

    def _on_show_log(self):
        """弹出日志查看窗口"""
        win = ctk.CTkToplevel(self)
        win.title("日志输出")
        win.geometry("700x450")
        win.after(100, win.lift)
        win.grab_set()

        # 工具栏
        tb = ctk.CTkFrame(win, fg_color='transparent')
        tb.pack(fill='x', padx=12, pady=(12, 4))
        ctk.CTkLabel(tb, text="📋 日志输出",
                     font=('sans-serif', 14, 'bold'),
                     text_color=TEXT_PRIMARY).pack(side='left')

        def copy_log():
            win.clipboard_clear()
            win.clipboard_append(self.log_text.get('1.0', 'end-1c'))
            win.update()

        ctk.CTkButton(tb, text="📋 复制全部", width=100, height=28,
                      fg_color='transparent', text_color=TEXT_SECONDARY,
                      border_width=1, border_color=BORDER, corner_radius=6,
                      font=('sans-serif', 10),
                      hover_color='#E5E7EB',
                      command=copy_log).pack(side='right')

        ctk.CTkButton(tb, text="🗑 清空", width=70, height=28,
                      fg_color='transparent', text_color=DANGER_RED,
                      border_width=1, border_color=BORDER, corner_radius=6,
                      font=('sans-serif', 10),
                      hover_color='#FEE2E2',
                      command=lambda: (self.log_text.delete('1.0', 'end'),
                                        win.destroy())).pack(side='right', padx=(0, 8))

        # 日志内容（可选择复制）
        log_view = ctk.CTkTextbox(win, fg_color=BG_SIDEBAR,
                                   text_color=TEXT_PRIMARY,
                                   font=('monospace', 10),
                                   border_width=1, border_color=BORDER,
                                   corner_radius=8)
        log_view.pack(fill='both', expand=True, padx=12, pady=(4, 12))
        log_view.insert('1.0', self.log_text.get('1.0', 'end-1c'))

    # ═══ 文件操作 ═══

    def _on_select_source_dir(self):
        path = filedialog.askdirectory(title="选择源文件夹")
        if not path:
            return
        self.source_dir = path
        self.source_dir_var.set(path)
        self._scan_source_files()

    def _on_rescan(self):
        if self.source_dir:
            self._scan_source_files()

    def _scan_source_files(self):
        self._log("扫描源文件夹: {}".format(self.source_dir))
        try:
            self.source_files = find_files(self.source_dir, recursive=True)
            if self.target_path:
                target_abs = os.path.abspath(self.target_path)
                self.source_files = [f for f in self.source_files
                                     if os.path.abspath(f) != target_abs]
        except Exception as e:
            self._log("扫描失败: {}".format(e))
            self.source_files = []
        self._log("找到 {} 个文件".format(len(self.source_files)))
        self._refresh_file_list()
        self._update_steps(0 if not self.target_path else 1)

    def _refresh_file_list(self):
        for w in self.file_scroll.winfo_children():
            w.destroy()
        self._file_buttons.clear()
        for fpath in self.source_files:
            rel = get_relative_path(fpath, self.source_dir)
            btn = ctk.CTkButton(
                self.file_scroll, text=rel, height=28, anchor='w',
                fg_color='transparent', text_color=TEXT_PRIMARY,
                hover_color='#DBEAFE', corner_radius=4,
                font=('sans-serif', 10),
                command=lambda p=fpath: self._on_file_click(p)
            )
            btn.pack(fill='x', pady=1)
            self._file_buttons[fpath] = btn

    def _on_file_click(self, filepath):
        for p, btn in self._file_buttons.items():
            if p == filepath:
                btn.configure(fg_color=ACCENT_BLUE, text_color='white')
            else:
                btn.configure(fg_color='transparent', text_color=TEXT_PRIMARY)
        self._load_source_file(filepath)

    def _load_source_file(self, filepath):
        self.current_source_file = filepath
        rel = get_relative_path(filepath, self.source_dir)
        self._log("加载: {}".format(rel))
        try:
            data = read_excel(filepath)
            self._source_data_cache[filepath] = data
            sheets = list(data.keys())
            self.src_sheet_combo.configure(values=sheets)
            if sheets:
                self.src_sheet_combo.set(sheets[0])
                self.source_sheet = sheets[0]
                self.src_grid.load_data(data[sheets[0]], sheets[0])
            self._update_steps(1)
        except Exception as e:
            self._log("加载失败: {}".format(e))

    def _on_source_sheet_change(self, choice):
        if not choice or not self.current_source_file:
            return
        data = self._source_data_cache.get(self.current_source_file, {})
        if choice in data:
            self.source_sheet = choice
            self.src_grid.load_data(data[choice], choice)
            self._reset_src_status()

    def _on_select_target(self):
        path = filedialog.askopenfilename(
            title="选择目标模板文件",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if path:
            self.target_path = path
            self.target_path_var.set(path)
            self._load_target_file(path)

    def _load_target_file(self, path):
        self._log("加载目标模板: {}".format(path))
        try:
            data = read_excel(path)
            sheets = list(data.keys())
            self.tgt_sheet_combo.configure(values=sheets)
            if sheets:
                self.tgt_sheet_combo.set(sheets[0])
                self.target_sheet = sheets[0]
                self.tgt_grid.load_data(data[sheets[0]], sheets[0])
            self.mapping.target_key_col = -1
            self.mapping.target_key_label = ''
            self.mapping.target_dest_col = -1
            self.mapping.target_dest_label = ''
            self._tgt_click_count = 0
            self.tgt_grid.clear_highlights()
            self._reset_tgt_status()
            self._update_mapping_display()
            if self.source_dir:
                self._scan_source_files()
            self._update_steps(1)
        except Exception as e:
            self._log("加载失败: {}".format(e))

    def _on_target_sheet_change(self, choice):
        if not choice or not self.target_path:
            return
        try:
            data = read_excel(self.target_path)
            if choice in data:
                self.target_sheet = choice
                self.tgt_grid.load_data(data[choice], choice)
                self._reset_tgt_status()
        except Exception as e:
            self._log("切换失败: {}".format(e))

    # ═══ 单元格点击 ═══

    def _on_source_cell_click(self, row_idx, col_idx):
        value = self.src_grid.get_cell_value(row_idx, col_idx)
        col_letter = chr(ord('A') + col_idx) if col_idx < 26 else '?'

        if self._src_click_count == 0:
            self.mapping.source_value_col = col_idx
            self.mapping.source_value_label = value[:20]
            self.src_grid.clear_highlights()
            self.src_grid.highlight_cell(row_idx, col_idx, 'source_value')
            self.src_status_var.set("🔵 复制列: {} | \"{}\"".format(col_letter, value[:15]))
            self.src_status_badge.configure(fg_color='#DBEAFE')
            self._log("✅ 源文件: 复制列 → {} (\"{}\")".format(col_letter, value[:20]))
            self._src_click_count = 1
        elif self._src_click_count == 1:
            self.mapping.source_key_col = col_idx
            self.mapping.source_key_label = value[:20]
            self.src_grid.highlight_cell(row_idx, col_idx, 'source_key')
            self.src_status_var.set("🟠 匹配键: {} | \"{}\"".format(col_letter, value[:15]))
            self.src_status_badge.configure(fg_color='#FFEDD5')
            self._log("✅ 源文件: 匹配键 → {} (\"{}\") | 源映射完成".format(col_letter, value[:20]))
            self._src_click_count = 2
            self._update_steps(2)
        else:
            self.src_grid.clear_highlights()
            self.mapping.source_value_col = col_idx
            self.mapping.source_value_label = value[:20]
            self.mapping.source_key_col = -1
            self.mapping.source_key_label = ''
            self.src_grid.highlight_cell(row_idx, col_idx, 'source_value')
            self.src_status_var.set("🔵 复制列: {} | \"{}\"".format(col_letter, value[:15]))
            self.src_status_badge.configure(fg_color='#DBEAFE')
            self._src_click_count = 1
        self._update_mapping_display()

    def _on_target_cell_click(self, row_idx, col_idx):
        value = self.tgt_grid.get_cell_value(row_idx, col_idx)
        col_letter = chr(ord('A') + col_idx) if col_idx < 26 else '?'

        if self._tgt_click_count == 0:
            self.mapping.target_key_col = col_idx
            self.mapping.target_key_label = value[:20]
            self.tgt_grid.clear_highlights()
            self.tgt_grid.highlight_cell(row_idx, col_idx, 'target_key')
            self.tgt_status_var.set("🟢 匹配键: {} | \"{}\"".format(col_letter, value[:15]))
            self.tgt_status_badge.configure(fg_color='#DCFCE7')
            self._log("✅ 目标: 匹配键 → {} (\"{}\")".format(col_letter, value[:20]))
            self._tgt_click_count = 1
        elif self._tgt_click_count == 1:
            self.mapping.target_dest_col = col_idx
            self.mapping.target_dest_label = value[:20]
            self.tgt_grid.highlight_cell(row_idx, col_idx, 'target_dest')
            self.tgt_status_var.set("🟡 粘贴到: {} | \"{}\"".format(col_letter, value[:15]))
            self.tgt_status_badge.configure(fg_color='#FEF9C3')
            self._log("✅ 目标: 粘贴到 → {} | 目标映射完成".format(col_letter))
            self._tgt_click_count = 2
            self._update_steps(2)
        else:
            self.tgt_grid.clear_highlights()
            self.mapping.target_key_col = col_idx
            self.mapping.target_key_label = value[:20]
            self.mapping.target_dest_col = -1
            self.mapping.target_dest_label = ''
            self.tgt_grid.highlight_cell(row_idx, col_idx, 'target_key')
            self.tgt_status_var.set("🟢 匹配键: {} | \"{}\"".format(col_letter, value[:15]))
            self.tgt_status_badge.configure(fg_color='#DCFCE7')
            self._tgt_click_count = 1
        self._update_mapping_display()

    def _reset_src_status(self):
        self.src_status_var.set("等待选择")
        self.src_status_badge.configure(fg_color='#E5E7EB')

    def _reset_tgt_status(self):
        self.tgt_status_var.set("等待选择")
        self.tgt_status_badge.configure(fg_color='#E5E7EB')

    # ═══ 映射管理 ═══

    def _update_mapping_display(self):
        lines = ["─ 源文件 ─"]
        if self.mapping.source_value_col >= 0:
            lines.append("  复制: {}列".format(chr(65 + self.mapping.source_value_col)))
        else:
            lines.append("  复制: —")
        if self.mapping.source_key_col >= 0:
            lines.append("  匹配: {}列".format(chr(65 + self.mapping.source_key_col)))
        else:
            lines.append("  匹配: —")
        lines.append("─ 目标模板 ─")
        if self.mapping.target_key_col >= 0:
            lines.append("  匹配: {}列".format(chr(65 + self.mapping.target_key_col)))
        else:
            lines.append("  匹配: —")
        if self.mapping.target_dest_col >= 0:
            lines.append("  粘贴: {}列".format(chr(65 + self.mapping.target_dest_col)))
        else:
            lines.append("  粘贴: —")
        if self.mapping.is_ready():
            lines.append("")
            lines.append("✅ 映射完整，可预览/处理")
            self.map_label.configure(text_color=ACCENT_GREEN)
        else:
            self.map_label.configure(text_color=TEXT_SECONDARY)
        self.map_text_var.set('\n'.join(lines))

    def _on_clear_mapping(self):
        self.mapping = ColumnMapping()
        self._src_click_count = 0
        self._tgt_click_count = 0
        self.src_grid.clear_highlights()
        self.tgt_grid.clear_highlights()
        self._reset_src_status()
        self._reset_tgt_status()
        self._update_mapping_display()
        self._update_steps(1)
        self._log("映射已清除")

    # ═══ 预览 & 执行 ═══

    def _on_preview(self):
        if not self.mapping.is_ready():
            messagebox.showwarning("提示", "请先完成映射定义")
            return
        if not self.source_dir or not self.target_path:
            messagebox.showwarning("提示", "请先选择源文件夹和目标模板")
            return
        self._log("生成预览...")
        try:
            result = preview_matches(self.source_dir, self.target_path,
                                     self.mapping,
                                     source_sheet=self.source_sheet or None,
                                     target_sheet=self.target_sheet or None)
        except Exception as e:
            self._log("预览失败: {}".format(e))
            return
        self._log("预览: {} 条匹配, {} 条有数据".format(
            result.matched_count + result.no_data_count, result.matched_count))
        PreviewDialog(self, result, self.mapping, self._on_preview_confirmed)

    def _on_preview_confirmed(self, entries):
        self._pending_entries = entries
        if entries:
            self._do_process()

    def _on_process(self):
        if not self.mapping.is_ready():
            messagebox.showwarning("提示", "请先完成映射定义")
            return
        self._on_preview()

    def _do_process(self):
        entries = getattr(self, '_pending_entries', None)
        if not entries:
            return
        selected = [e for e in entries if e.selected and e.has_data]
        if not selected:
            self._log("没有选中数据")
            return
        self._log("开始批量处理 {} 条...".format(len(selected)))
        self.progress_var.set("处理中...")

        def run():
            try:
                output = process_and_write(
                    self.target_path, self.mapping, entries,
                    target_sheet=self.target_sheet or None,
                    progress_callback=self._log)
                self._log("✅ 完成! {}".format(output))
                self.progress_var.set("完成")
                self.after(0, lambda: messagebox.showinfo("完成", "处理完成!\n\n{}".format(output)))
            except Exception as e:
                self._log("❌ 失败: {}".format(e))
                self.progress_var.set("失败")
                self.after(0, lambda: messagebox.showerror("错误", "处理失败:\n{}".format(e)))

        threading.Thread(target=run, daemon=True).start()

    # ═══ 规则存取 ═══

    def _on_save_rule(self):
        if not self.mapping.is_source_ready():
            messagebox.showwarning("提示", "请至少完成源文件映射")
            return
        path = filedialog.asksaveasfilename(
            title="保存映射规则", defaultextension=".json",
            filetypes=[("JSON 文件", "*.json")])
        if path:
            try:
                save_rule(path, self.mapping, self.source_dir, self.target_path)
                self._log("规则已保存: {}".format(path))
            except Exception as e:
                self._log("保存失败: {}".format(e))

    def _on_load_rule(self):
        path = filedialog.askopenfilename(
            title="加载映射规则",
            filetypes=[("JSON 文件", "*.json")])
        if not path:
            return
        try:
            mapping, source_dir, target_path = load_rule(path)
            self.mapping = mapping
            self._src_click_count = 2
            self._tgt_click_count = 2
            self._update_mapping_display()
            self._update_steps(2)
            self._log("规则已加载: {}".format(path))
            if source_dir and os.path.isdir(source_dir):
                self.source_dir = source_dir
                self.source_dir_var.set(source_dir)
                self._scan_source_files()
            if target_path and os.path.isfile(target_path):
                self.target_path = target_path
                self.target_path_var.set(target_path)
                self._load_target_file(target_path)
        except Exception as e:
            self._log("加载失败: {}".format(e))


class _RedirectText:
    def __init__(self, widget):
        self.widget = widget
    def write(self, s):
        self.widget.insert('end', s)
        self.widget.see('end')
        self.widget.update_idletasks()
    def flush(self):
        pass


def main():
    app = App()
    app.mainloop()


if __name__ == '__main__':
    main()
