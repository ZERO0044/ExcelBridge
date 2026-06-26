#!/usr/bin/env python3
"""
ExcelBridge — Figma UI3 风格界面
三栏布局 · 上下分屏预览 · 按钮即时反馈
"""

import os, sys, struct, zlib
import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, filedialog
from typing import Optional

from excel_reader import find_files, read_excel, get_relative_path


# ═══ 简约 PNG 图标 ═══

def _create_png(w, h, pixels):
    def chunk(ct, d):
        c = ct + d
        return struct.pack('>I', len(d)) + c + struct.pack('>I', zlib.crc32(c) & 0xFFFFFFFF)
    raw = b''
    for y in range(h):
        raw += b'\x00'
        for x in range(w):
            raw += struct.pack('BBBB', *pixels[y * w + x])
    ihdr = struct.pack('>IIBBBBB', w, h, 8, 6, 0, 0, 0)
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', ihdr) + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b'')

def _solid_png(hex_color, size=20, r=5):
    R, G, B = int(hex_color[1:3], 16), int(hex_color[3:5], 16), int(hex_color[5:7], 16)
    px = []
    for y in range(size):
        for x in range(size):
            inside = True
            if r > 0:
                for cx, cy in [(r-1,r-1), (size-r,r-1), (r-1,size-r), (size-r,size-r)]:
                    if (x-cx)**2 + (y-cy)**2 > r**2:
                        if (x < r and y < r and cx == r-1): inside = False
                        if (x >= size-r and y < r and cx == size-r): inside = False
                        if (x < r and y >= size-r and cx == r-1): inside = False
                        if (x >= size-r and y >= size-r and cx == size-r): inside = False
            px.append((R, G, B, 255) if inside else (0, 0, 0, 0))
    return _create_png(size, size, px)


# ═══ 色彩 ═══

class C:
    PRIMARY       = '#2563EB'
    PRIMARY_HOVER = '#1D4ED8'
    PRIMARY_BG    = '#EFF6FF'
    COPY_BLUE     = '#3B82F6'
    MATCH_ORANGE  = '#F59E0B'
    TARGET_GREEN  = '#10B981'
    PASTE_PURPLE  = '#8B5CF6'
    COPY_BG       = '#EFF6FF'
    MATCH_BG      = '#FFFBEB'
    TARGET_BG     = '#ECFDF5'
    PASTE_BG      = '#F5F3FF'
    TEXT_TITLE    = '#0F172A'
    TEXT_BODY     = '#334155'
    TEXT_HELP     = '#64748B'
    TEXT_MUTED    = '#94A3B8'
    BG_MAIN       = '#F7F9FC'
    BG_SIDEBAR_L  = '#F7F9FC'
    BG_SIDEBAR_R  = '#F8FAFC'
    BG_CARD       = '#FFFFFF'
    BG_WHITE      = '#FFFFFF'
    BORDER        = '#E2E8F0'
    SUCCESS       = '#10B981'
    ERROR         = '#EF4444'
    TBL_HDR_BG    = '#F1F5F9'
    TBL_HDR_FG    = '#1E293B'
    TBL_EVEN      = '#FFFFFF'
    TBL_ODD       = '#FAFAFA'


# ═══ 主应用 ═══

class ExcelBridgeApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("ExcelBridge — 数据批量匹配迁移工具")
        self.geometry("1250x780")
        self.minsize(1000, 620)
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self.configure(fg_color=C.BG_MAIN)

        # 图标
        self._icons_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '_icons')
        os.makedirs(self._icons_dir, exist_ok=True)
        self._icons = {}
        self._gen_icons()

        # 状态
        self._step = 0
        self._map_copy: Optional[int] = None
        self._map_match: Optional[int] = None
        self._map_tkey: Optional[int] = None
        self._map_paste: Optional[int] = None
        self._map_copy_val: str = ""
        self._map_match_val: str = ""
        self._map_tkey_val: str = ""
        self._map_paste_val: str = ""
        self._active_table = 'source'
        self._log_lines = []  # 收集日志

        ExcelBridgeApp._setup_styles()  # 全局样式（只执行一次）
        self._build()
        self._sync_ui()

    # ── 图标 ──

    def _gen_icons(self):
        for name, color in {
            'dot_blue': C.COPY_BLUE, 'dot_orange': C.MATCH_ORANGE,
            'dot_green': C.TARGET_GREEN, 'dot_purple': C.PASTE_PURPLE,
            'primary': C.PRIMARY, 'success': C.SUCCESS,
        }.items():
            path = os.path.join(self._icons_dir, f'{name}.png')
            if not os.path.exists(path):
                with open(path, 'wb') as f:
                    f.write(_solid_png(color, 20, 5))
            try:
                from PIL import Image
                img = Image.open(path)
                self._icons[name] = ctk.CTkImage(img, img, size=(18, 18))
            except ImportError:
                self._icons[name] = None

    # ── 布局入口 ──

    def _build(self):
        self.grid_columnconfigure(0, minsize=250)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, minsize=270)
        self.grid_rowconfigure(0, weight=1)

        self._build_left()
        self._build_center()
        self._build_right()

    def _cf(self, p, **kw):
        return ctk.CTkFrame(p, fg_color='transparent', **kw)

    # ── 左侧栏 ──

    def _build_left(self):
        L = ctk.CTkFrame(self, fg_color=C.BG_SIDEBAR_L, corner_radius=0)
        L.grid(row=0, column=0, sticky='nswe')
        ctk.CTkFrame(L, fg_color=C.BORDER, width=1).place(relx=1, rely=0, relheight=1, anchor='ne')
        L.grid_rowconfigure(2, weight=1)
        L.grid_columnconfigure(0, weight=1)

        top = self._cf(L)
        top.grid(row=0, column=0, sticky='ew', padx=10, pady=(12, 4))

        # 源文件夹（紧凑单行）
        ctk.CTkLabel(top, text="📂 源文件夹", font=('sans-serif', 10, 'bold'),
                     text_color=C.TEXT_TITLE).grid(row=0, column=0, sticky='w')
        self._src_entry = ctk.CTkEntry(top, height=28, placeholder_text="选择文件夹...",
                                        fg_color=C.BG_WHITE, border_color=C.BORDER,
                                        corner_radius=6, text_color=C.TEXT_BODY,
                                        font=('sans-serif', 10))
        self._src_entry.grid(row=1, column=0, sticky='ew', pady=(2, 2), columnspan=2)
        r1 = self._cf(top)
        r1.grid(row=2, column=0, sticky='ew')
        self._btn(r1, "浏览...", 56, 24, C.PRIMARY_BG, C.PRIMARY, 'DBEAFE',
                  lambda: self._browse_folder()).grid(row=0, column=0, padx=(0, 4))
        self._btn(r1, "🔄 扫描", 56, 24, 'transparent', C.TEXT_HELP, 'E5E7EB',
                  lambda: self._act("重新扫描")).grid(row=0, column=1)

        # 目标模板
        ctk.CTkLabel(top, text="📋 目标模板", font=('sans-serif', 10, 'bold'),
                     text_color=C.TEXT_TITLE).grid(row=3, column=0, sticky='w', pady=(10, 0))
        self._tgt_entry = ctk.CTkEntry(top, height=28, placeholder_text="选择模板...",
                                        fg_color=C.BG_WHITE, border_color=C.BORDER,
                                        corner_radius=6, text_color=C.TEXT_BODY,
                                        font=('sans-serif', 10))
        self._tgt_entry.grid(row=4, column=0, sticky='ew', pady=(2, 2), columnspan=2)
        r2 = self._cf(top)
        r2.grid(row=5, column=0, sticky='ew')
        self._btn(r2, "浏览...", 56, 24, C.PRIMARY_BG, C.PRIMARY, 'DBEAFE',
                  lambda: self._browse_target()).grid(row=0, column=0, padx=(0, 4))
        self._btn(r2, "📂 加载", 56, 24, 'transparent', C.TEXT_HELP, 'E5E7EB',
                  lambda: self._act("加载目标模板")).grid(row=0, column=1)

        # 文件列表标题
        ctk.CTkLabel(L, text="📄 源文件列表（点击加载）",
                     font=('sans-serif', 10, 'bold'),
                     text_color=C.TEXT_BODY).grid(
            row=1, column=0, sticky='w', padx=12, pady=(12, 4))

        # 文件列表（动态扫描）
        self._file_frame = ctk.CTkScrollableFrame(
            L, fg_color=C.BG_WHITE, corner_radius=8,
            scrollbar_button_color=C.BORDER,
            scrollbar_button_hover_color=C.TEXT_MUTED)
        self._file_frame.grid(row=2, column=0, sticky='nswe', padx=10, pady=(8, 6))
        self._file_buttons = {}  # path -> CTkButton

        # 底部按钮
        bb = self._cf(L)
        bb.grid(row=3, column=0, sticky='ew', padx=10, pady=(0, 12))
        bb.grid_columnconfigure(0, weight=1)
        self._btn(bb, "💾 保存规则", None, 28, 'transparent', C.TEXT_HELP, 'E5E7EB',
                  lambda: self._act("保存规则")).grid(row=0, column=0, sticky='ew', pady=(0, 3))
        self._btn(bb, "📂 加载规则", None, 28, '#F3F4F6', C.TEXT_HELP, 'E5E7EB',
                  lambda: self._act("加载规则")).grid(row=1, column=0, sticky='ew')

    def _browse_folder(self):
        path = filedialog.askdirectory(title="选择源文件夹")
        if path:
            self._source_dir = path
            self._src_entry.delete(0, 'end')
            self._src_entry.insert(0, path)
            self._scan_files(path)
            self._act(f"选择源文件夹: {path}")

    def _scan_files(self, source_dir):
        """扫描目录下的 xlsx/xls 文件并填充列表"""
        for w in self._file_frame.winfo_children():
            w.destroy()
        self._file_buttons.clear()

        try:
            files = find_files(source_dir, recursive=True)
            # 排除目标文件
            if hasattr(self, '_tgt_path') and self._tgt_path:
                target_abs = os.path.abspath(self._tgt_path)
                files = [f for f in files if os.path.abspath(f) != target_abs]
        except Exception:
            files = []

        self._source_files_list = files  # 保存用于批量处理

        if not files:
            ctk.CTkLabel(self._file_frame, text="未找到 xlsx/xls 文件",
                         font=('sans-serif', 10),
                         text_color=C.TEXT_MUTED).pack(pady=12)
            return

        for fpath in files:
            rel = get_relative_path(fpath, source_dir)
            btn = ctk.CTkButton(
                self._file_frame, text=f"📄 {rel}", height=28, anchor='w',
                fg_color='transparent', text_color=C.TEXT_BODY,
                hover_color='#DBEAFE', corner_radius=6,
                font=('sans-serif', 10),
                command=lambda p=fpath: self._on_file_select(p))
            btn.pack(fill='x', pady=1)
            self._file_buttons[fpath] = btn

    def _on_file_select(self, filepath):
        """选中文件 → 加载到源文件预览"""
        self._current_src_path = filepath
        for p, btn in self._file_buttons.items():
            if p == filepath:
                btn.configure(fg_color=C.PRIMARY_BG, text_color=C.PRIMARY)
            else:
                btn.configure(fg_color='transparent', text_color=C.TEXT_BODY)

        try:
            data = read_excel(filepath)
            sheet = list(data.keys())[0]
            rows = data[sheet]
            self._src_tree.delete(*self._src_tree.get_children())
            for i, row in enumerate(rows):
                tag = 'even.X' if i % 2 == 0 else 'odd.X'
                vals = [str(c) if c is not None else '' for c in row]
                self._src_tree.insert('', 'end', iid=str(i), text=str(i+1),
                                      values=vals[:10] if len(vals) > 10 else vals + ['']*(10-len(vals)),
                                      tags=(tag,))
            if hasattr(self, '_src_info_label') and self._src_info_label is not None:
                try:
                    ncols = len(rows[0]) if rows else 0
                    self._src_info_label.configure(text=f"共 {ncols} 列 · {len(rows)} 行")
                except Exception:
                    pass
            self._act(f"选择文件: {os.path.basename(filepath)}")
            self._help_var.set("👉 在上方源文件预览中，点击「要复制的数据列」（蓝色标记）和「匹配依据列」（橙色标记）。")
            self._src_status.configure(text="已加载文件")
            self._src_badge.configure(fg_color=C.PRIMARY_BG)
        except Exception as e:
            print(f"加载文件失败: {e}")

    def _browse_target(self):
        path = filedialog.askopenfilename(title="选择目标模板",
                                           filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")])
        if path:
            self._tgt_entry.delete(0, 'end')
            self._tgt_entry.insert(0, path)
            self._tgt_path = path
            # 加载目标模板到预览
            try:
                data = read_excel(path)
                sheet = list(data.keys())[0]
                rows = data[sheet]
                self._tgt_tree.delete(*self._tgt_tree.get_children())
                for i, row in enumerate(rows):
                    tag = 'even.X' if i % 2 == 0 else 'odd.X'
                    vals = [str(c) if c is not None else '' for c in row]
                    self._tgt_tree.insert('', 'end', iid=str(i), text=str(i+1),
                                          values=vals[:10] if len(vals) > 10 else vals + ['']*(10-len(vals)),
                                          tags=(tag,))
                if hasattr(self, '_tgt_info_label') and self._tgt_info_label is not None:
                    self._tgt_info_label.configure(text=f"共 {len(rows[0]) if rows else 0} 列 · {len(rows)} 行")
                self._act(f"选择目标模板: {path}")
            except Exception as e:
                print(f"加载目标失败: {e}")

    # ── 中间面板 ──

    def _build_center(self):
        C_ = ctk.CTkFrame(self, fg_color=C.BG_WHITE, corner_radius=0)
        C_.grid(row=0, column=1, sticky='nswe')
        C_.grid_rowconfigure(1, weight=50)  # 源文件预览
        C_.grid_rowconfigure(2, weight=50)  # 目标模板预览（等高）
        C_.grid_columnconfigure(0, weight=1)

        # 步骤条
        self._step_widgets = []
        sb = self._cf(C_)
        sb.grid(row=0, column=0, sticky='ew', padx=20, pady=(14, 6))

        steps = ["① 选源文件", "② 标记字段", "③ 匹配目标", "④ 执行写入"]
        for i, label in enumerate(steps):
            f = self._cf(sb)
            f.grid(row=0, column=i, padx=(0, 12 if i < 3 else 0))
            dot = ctk.CTkFrame(f, fg_color=C.BORDER, width=24, height=24, corner_radius=12)
            dot.grid(row=0, column=0, padx=(0, 6))
            dot.grid_propagate(False)
            dl = ctk.CTkLabel(dot, text=str(i+1), font=('sans-serif', 10, 'bold'), text_color='white')
            dl.place(relx=0.5, rely=0.5, anchor='center')
            tl = ctk.CTkLabel(f, text=label, font=('sans-serif', 10), text_color=C.TEXT_HELP)
            tl.grid(row=0, column=1)
            if i < 3:
                ctk.CTkFrame(f, fg_color=C.BORDER, width=24, height=2).grid(row=0, column=2, padx=(6, 0))
            self._step_widgets.append((f, dot, dl, tl))

        # ── 上下分屏：源文件(上) + 目标模板(下) ──
        # 源文件预览卡片
        self._src_card = self._build_preview_card(C_, 0, "📄 源文件预览",
                                                    C.COPY_BLUE, C.COPY_BG,
                                                    C.MATCH_ORANGE, C.MATCH_BG)
        # 目标模板预览卡片
        self._tgt_card = self._build_preview_card(C_, 1, "📋 目标模板预览",
                                                    C.TARGET_GREEN, C.TARGET_BG,
                                                    C.PASTE_PURPLE, C.PASTE_BG)

    def _build_preview_card(self, parent, row, title, color1, bg1, color2, bg2):
        """构建单个预览卡片（源或目标）"""
        card = ctk.CTkFrame(parent, fg_color=C.BG_CARD, corner_radius=8)
        card.grid(row=row + 1, column=0, sticky='nswe', padx=20,
                  pady=(0 if row == 0 else 8, 8 if row == 1 else 0))
        card.grid_rowconfigure(3, weight=1)
        card.grid_columnconfigure(0, weight=1)

        # 标题栏
        hdr = self._cf(card)
        hdr.grid(row=0, column=0, sticky='ew', padx=12, pady=(8, 2))
        ctk.CTkLabel(hdr, text=title, font=('sans-serif', 11, 'bold'),
                     text_color=C.TEXT_TITLE).pack(side='left')

        # 状态徽章
        badge = ctk.CTkFrame(hdr, fg_color='#E5E7EB', corner_radius=6, height=20)
        badge.pack(side='right')
        status_label = ctk.CTkLabel(badge, text="等待选择", font=('sans-serif', 9),
                                     text_color=C.TEXT_HELP)
        status_label.pack(padx=8)

        # 图例
        leg = self._cf(card)
        leg.grid(row=1, column=0, sticky='ew', padx=12, pady=(0, 2))
        for i, (text, c) in enumerate([
            ("复制列" if row == 0 else "匹配键", color1),
            ("匹配键" if row == 0 else "粘贴至", color2),
        ]):
            ld = ctk.CTkFrame(leg, fg_color=bg1 if i == 0 else bg2,
                               corner_radius=6, height=20)
            ld.pack(side='left', padx=(0, 8))
            dot = ctk.CTkFrame(ld, fg_color=c, width=8, height=8, corner_radius=6)
            dot.pack(side='left', padx=(6, 4), pady=4)
            ctk.CTkLabel(ld, text=text, font=('sans-serif', 9),
                         text_color=c).pack(side='left', padx=(0, 6))

        # 列信息
        info_label = ctk.CTkLabel(leg, text="—", font=('sans-serif', 9),
                     text_color=C.TEXT_MUTED).pack(side='right')

        # Sheet 选择器
        sh = self._cf(card)
        sh.grid(row=2, column=0, sticky='ew', padx=12, pady=(0, 4))
        ctk.CTkLabel(sh, text="📊 Sheet:", font=('sans-serif', 9),
                     text_color=C.TEXT_HELP).pack(side='left', padx=(0, 4))
        sv = ctk.StringVar(value="Sheet1")
        ctk.CTkComboBox(sh, width=110, height=24, variable=sv,
                        values=["Sheet1", "Sheet2"], font=('sans-serif', 9),
                        fg_color=C.BG_WHITE, border_color=C.BORDER,
                        button_color=C.PRIMARY, state='readonly',
                        command=lambda v: self._act(f"切换Sheet: {v}")).pack(side='left')

        # 表格
        tf = ctk.CTkFrame(card, fg_color='transparent', corner_radius=6)
        tf.grid(row=3, column=0, sticky='nswe', padx=8, pady=(0, 8))
        tf.grid_rowconfigure(0, weight=1)
        tf.grid_columnconfigure(0, weight=1)

        tree = self._make_table(tf, row == 0)
        # 存引用
        if row == 0:
            self._src_tree = tree
            self._src_status = status_label
            self._src_badge = badge
            self._src_info_label = info_label
        else:
            self._tgt_tree = tree
            self._tgt_status = status_label
            self._tgt_badge = badge
            self._tgt_info_label = info_label

        return card

    @staticmethod
    def _setup_styles():
        """全局表格和滚动条样式（只执行一次）"""
        style = ttk.Style()
        if 'clam' in style.theme_names():
            style.theme_use('clam')
        style.configure('TScrollbar', background=C.BORDER,
                        troughcolor=C.BG_WHITE, arrowcolor=C.TEXT_MUTED,
                        borderwidth=0, arrowsize=12)
        style.map('TScrollbar', background=[('active', C.TEXT_MUTED)])
        style.configure('X.Treeview', background=C.BG_WHITE, foreground=C.TEXT_BODY,
                        fieldbackground=C.BG_WHITE, borderwidth=0, rowheight=28,
                        font=('sans-serif', 9))
        style.configure('X.Treeview.Heading', background=C.TBL_HDR_BG,
                        foreground=C.TBL_HDR_FG, borderwidth=0, relief='flat',
                        font=('sans-serif', 9, 'bold'), padding=(6, 5))
        style.map('X.Treeview.Heading', background=[('active', '#E2E8F0')])
        style.configure('even.X', background=C.TBL_EVEN)
        style.configure('odd.X', background=C.TBL_ODD)
        style.configure('hl1.X', background=C.COPY_BG)
        style.configure('hl2.X', background=C.MATCH_BG)
        style.map('X.Treeview', background=[('selected', '#BFDBFE')],
                  foreground=[('selected', '#1E40AF')])

    def _make_table(self, parent, is_source):
        """创建美化的 Treeview（样式已由 _setup_styles 初始化）"""
        cols = tuple(chr(65 + i) for i in range(10))
        tree = ttk.Treeview(parent, columns=cols, show='tree headings',
                             style='X.Treeview', selectmode='browse')
        tree.grid(row=0, column=0, sticky='nswe')

        vsb = ttk.Scrollbar(parent, orient='vertical', command=tree.yview,
                            )
        vsb.grid(row=0, column=1, sticky='ns')
        hsb = ttk.Scrollbar(parent, orient='horizontal', command=tree.xview,
                            )
        hsb.grid(row=1, column=0, sticky='ew')
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.column('#0', width=40, minwidth=40, stretch=False, anchor='center')
        tree.heading('#0', text='')
        for c in cols:
            tree.column(c, width=85, minwidth=55, stretch=True, anchor='w')
            tree.heading(c, text=c)

        # 根据表格类型设置高亮颜色
        hl1_color = C.COPY_BG if is_source else C.TARGET_BG
        hl2_color = C.MATCH_BG if is_source else C.PASTE_BG
        style = ttk.Style()
        style.configure('hl1.X', background=hl1_color)
        style.configure('hl2.X', background=hl2_color)

        tree.bind('<Button-1>', lambda e, t=tree, s=is_source: self._on_cell(e, t, s))
        return tree

    # ── 右侧栏 ──

    def _build_right(self):
        R = ctk.CTkFrame(self, fg_color=C.BG_SIDEBAR_R, corner_radius=0)
        R.grid(row=0, column=2, sticky='nswe')
        ctk.CTkFrame(R, fg_color=C.BORDER, width=1).place(relx=0, rely=0, relheight=1, anchor='nw')
        R.grid_columnconfigure(0, weight=1)
        R.grid_rowconfigure(3, weight=1)  # 按钮+日志区自动扩展

        # 操作指引卡片
        gcard = ctk.CTkFrame(R, fg_color=C.BG_WHITE, corner_radius=8)
        gcard.grid(row=0, column=0, sticky='ew', padx=14, pady=(16, 8))
        ctk.CTkLabel(gcard, text="💡 操作指引", font=('sans-serif', 11, 'bold'),
                     text_color=C.TEXT_TITLE).grid(row=0, column=0, sticky='w',
                                                    padx=12, pady=(10, 2))
        self._help_var = ctk.StringVar(value="👈 请先在左侧选择源文件夹和目标模板文件")
        ctk.CTkLabel(gcard, textvariable=self._help_var, font=('sans-serif', 10),
                     text_color=C.TEXT_HELP, justify='left', anchor='w',
                     wraplength=220).grid(row=1, column=0, sticky='w',
                                           padx=12, pady=(2, 10))

        # 映射规则卡片
        rcard = ctk.CTkFrame(R, fg_color=C.BG_WHITE, corner_radius=8)
        rcard.grid(row=1, column=0, sticky='ew', padx=14, pady=(0, 8))
        ctk.CTkLabel(rcard, text="📐 映射规则", font=('sans-serif', 11, 'bold'),
                     text_color=C.TEXT_TITLE).grid(row=0, column=0, sticky='w',
                                                    padx=12, pady=(10, 6))

        rules = [
            ("🔵 复制列", '_map_copy', C.COPY_BLUE),
            ("🟠 源匹配键", '_map_match', C.MATCH_ORANGE),
            ("🟢 目标匹配键", '_map_tkey', C.TARGET_GREEN),
            ("🟣 粘贴至", '_map_paste', C.PASTE_PURPLE),
        ]
        self._rule_rows = []
        for i, (label, attr, color) in enumerate(rules):
            rf = self._cf(rcard)
            rf.grid(row=i + 1, column=0, sticky='ew', padx=12, pady=2)
            dot = ctk.CTkFrame(rf, fg_color=C.BORDER, width=8, height=8, corner_radius=6)
            dot.grid(row=0, column=0, padx=(0, 6))
            dot.grid_propagate(False)
            vl = ctk.CTkLabel(rf, text=f"{label}  — 待设置", font=('sans-serif', 10),
                              text_color=C.TEXT_MUTED, anchor='w')
            vl.grid(row=0, column=1, sticky='w')
            self._rule_rows.append((dot, vl, attr, color))

        self._btn(rcard, "清除映射", None, 26, 'transparent', C.ERROR, 'FEE2E2',
                  lambda: self._act("清除映射")).grid(row=5, column=0, sticky='w',
                                                       padx=(12, 0), pady=(4, 10))

        # 执行按钮 + 内嵌日志
        acard = ctk.CTkFrame(R, fg_color='transparent')
        acard.grid(row=3, column=0, sticky='nsew', padx=14, pady=(0, 12))
        acard.grid_columnconfigure(0, weight=1)
        acard.grid_rowconfigure(4, weight=1)  # 日志区自动填充

        self._btn(acard, "🔍 预览匹配结果", None, 38, C.PRIMARY_BG, C.PRIMARY, '#DBEAFE',
                  lambda: self._act("预览匹配结果")).grid(
            row=0, column=0, sticky='ew', pady=(0, 8))

        self._btn(acard, "🚀 确认并写入", None, 45, C.PRIMARY, 'white', C.PRIMARY_HOVER,
                  lambda: self._open_preview_window(), bold=True).grid(
            row=1, column=0, sticky='ew', pady=(0, 6))

        self._status_var = ctk.StringVar(value="就绪")
        ctk.CTkLabel(acard, textvariable=self._status_var, font=('sans-serif', 10),
                     text_color=C.TEXT_MUTED).grid(row=2, column=0, pady=(0, 4))

        # 内嵌迷你日志
        self._mini_log = ctk.CTkTextbox(acard,
                                         fg_color=C.BG_WHITE,
                                         text_color=C.TEXT_HELP,
                                         font=('monospace', 9),
                                         border_width=0, corner_radius=6, wrap='word')
        self._mini_log.grid(row=4, column=0, sticky='nsew', pady=(4, 0))
        self._mini_log.insert('1.0', '就绪 — 点击「📋 详细日志」查看完整记录')
        self._mini_log.configure(state='disabled')

        # 日志按钮
        ctk.CTkButton(acard, text="📋 详细日志", height=26,
                      fg_color='transparent', text_color=C.TEXT_HELP,
                      hover_color='#E5E7EB', corner_radius=6,
                      font=('sans-serif', 10),
                      command=self._open_log_window).grid(
            row=5, column=0, sticky='ew', pady=(4, 0))

    # ── 按钮工厂（统一圆角消除锯齿）──

    def _btn(self, parent, text, width, height, fg, text_color, hover,
             command, anchor='center', border=0, border_color=None, bold=False):
        kwargs = dict(
            text=text, height=height, fg_color=fg, text_color=text_color,
            hover_color=f'#{hover.lstrip("#")}', corner_radius=8 if height >= 34 else 6,
            font=('sans-serif', 12 if height >= 34 else 10,
                  'bold' if bold else 'normal'),
            command=command,
            border_width=0,  # 不用边框，避免圆角断裂
        )
        if width: kwargs['width'] = width
        if anchor != 'center': kwargs['anchor'] = anchor
        btn = ctk.CTkButton(parent, **kwargs)
        return btn

    # ── 行为（即时反馈）──

    def _act(self, msg):
        """统一操作入口：打印 + 记录日志 + 更新 mini 日志"""
        import datetime
        ts = datetime.datetime.now().strftime('%H:%M:%S')
        line = f"[{ts}] {msg}"
        self._log_lines.append(line)
        print(f"执行：{msg}")
        # 更新右下角迷你日志
        try:
            self._mini_log.configure(state='normal')
            self._mini_log.insert('end', line + '\n')
            self._mini_log.see('end')
            self._mini_log.configure(state='disabled')
        except Exception:
            pass

        # 模拟步骤推进
        if "选择源文件夹" in msg:
            self._step = max(self._step, 1)
        elif "加载目标模板" in msg:
            self._step = max(self._step, 1)
        elif "选择文件" in msg:
            self._step = max(self._step, 1)
            self._help_var.set("👉 在上方源文件预览中，点击「要复制的数据列」（蓝色标记）和「匹配依据列」（橙色标记）。")
            self._src_status.configure(text="已加载文件")
            self._src_badge.configure(fg_color=C.PRIMARY_BG)
        elif "预览匹配结果" in msg:
            self._open_preview_window()
        elif "开始批量写入" in msg:
            self._step = 4
            self._status_var.set("写入中...")
            self.after(1000, lambda: self._status_var.set("写入完成 ✅"))
            self._help_var.set("🎉 写入完成！请点击「📋 查看日志」查看详细结果。")
        elif "打开日志窗口" in msg:
            self._open_log_window()

        self._sync_ui()

    def _on_cell(self, event, tree, is_source):
        """点击表格单元格 → 模拟映射标记"""
        region = tree.identify_region(event.x, event.y)
        row = tree.identify_row(event.y)
        col_str = tree.identify_column(event.x)
        if not row or not col_str or region not in ('cell', 'tree'):
            return

        row_idx = int(row)
        col_idx = int(col_str.replace('#', '')) - 1
        if col_idx < 0: col_idx = 0

        # 获取单元格值
        vals = tree.item(row, 'values')
        cell_val = str(vals[col_idx]).strip() if col_idx < len(vals) else ''

        # 清除旧高亮
        for iid in tree.get_children():
            tree.item(iid, tags=('even.X' if int(iid) % 2 == 0 else 'odd.X',))
        tree.selection_remove(tree.selection())

        # 选中行 + 标记高亮
        tree.selection_add(row)
        if is_source:
            if self._map_copy is None:
                self._map_copy = col_idx
                self._map_copy_val = cell_val
                tree.item(row, tags=('hl1.X',))
                self._src_status.configure(text=f"🔵 复制列: {chr(65+col_idx)} \"{cell_val[:15]}\"")
                self._src_badge.configure(fg_color=C.COPY_BG)
                self._help_var.set("✅ 已标记复制列！现在请点击「匹配依据列」（橙色标记）。")
            elif self._map_match is None and col_idx != self._map_copy:
                self._map_match = col_idx
                self._map_match_val = cell_val
                tree.item(row, tags=('hl2.X',))
                self._src_status.configure(text=f"🟠 匹配键: {chr(65+col_idx)} \"{cell_val[:15]}\"")
                self._src_badge.configure(fg_color=C.MATCH_BG)
                self._help_var.set("🎯 源文件标记完成！现在在下方目标模板中点击「匹配键列」（绿色）和「粘贴列」（紫色）。")
                self._step = 2
            else:
                self._map_copy = col_idx
                self._map_copy_val = cell_val
                self._map_match = None
                self._map_match_val = ''
                tree.item(row, tags=('hl1.X',))
                self._src_status.configure(text=f"🔵 复制列: {chr(65+col_idx)} \"{cell_val[:15]}\"")
                self._src_badge.configure(fg_color=C.COPY_BG)
                self._help_var.set("👉 请点击「匹配依据列」（橙色标记）。")
                self._step = 1
        else:
            if self._map_tkey is None:
                self._map_tkey = col_idx
                self._map_tkey_val = cell_val
                tree.item(row, tags=('hl1.X',))
                self._tgt_status.configure(text=f"🟢 匹配键: {chr(65+col_idx)} \"{cell_val[:15]}\"")
                self._tgt_badge.configure(fg_color=C.TARGET_BG)
                self._help_var.set("✅ 目标匹配键已标记！现在请点击「要粘贴到的列」（紫色标记）。")
            elif self._map_paste is None and col_idx != self._map_tkey:
                self._map_paste = col_idx
                self._map_paste_val = cell_val
                tree.item(row, tags=('hl2.X',))
                self._tgt_status.configure(text=f"🟣 粘贴至: {chr(65+col_idx)} \"{cell_val[:15]}\"")
                self._tgt_badge.configure(fg_color=C.PASTE_BG)
                self._help_var.set("🎉 映射完成！请点击右侧「预览匹配结果」查看数据。")
                self._step = 3
            else:
                self._map_tkey = col_idx
                self._map_tkey_val = cell_val
                self._map_paste = None
                self._map_paste_val = ''
                tree.item(row, tags=('hl1.X',))
                self._tgt_status.configure(text=f"🟢 匹配键: {chr(65+col_idx)} \"{cell_val[:15]}\"")
                self._tgt_badge.configure(fg_color=C.TARGET_BG)
                self._help_var.set("👉 请点击「要粘贴到的列」（紫色标记）。")
                self._step = 2

        print(f"执行：点击{'源' if is_source else '目标'}表格 {chr(65+col_idx)}{row_idx+1}")
        self._sync_ui()

    # ── 同步 ──

    def _sync_ui(self):
        """更新步骤条 + 规则卡片"""
        # 步骤条
        for i, (f, dot, dl, tl) in enumerate(self._step_widgets):
            if i < self._step:
                dot.configure(fg_color=C.SUCCESS)
                dl.configure(text='✔')
                tl.configure(text_color=C.SUCCESS, font=('sans-serif', 10, 'bold'))
            elif i == self._step:
                dot.configure(fg_color=C.PRIMARY)
                dl.configure(text=str(i+1))
                tl.configure(text_color=C.PRIMARY, font=('sans-serif', 10, 'bold'))
            else:
                dot.configure(fg_color=C.BORDER)
                dl.configure(text=str(i+1))
                tl.configure(text_color=C.TEXT_HELP, font=('sans-serif', 10))

        # 规则卡片
        mapping_vals = {
            '_map_copy': (self._map_copy, self._map_copy_val),
            '_map_match': (self._map_match, self._map_match_val),
            '_map_tkey': (self._map_tkey, self._map_tkey_val),
            '_map_paste': (self._map_paste, self._map_paste_val),
        }
        for dot, vl, attr, color in self._rule_rows:
            col_idx, cell_val = mapping_vals[attr]
            if col_idx is not None:
                dot.configure(fg_color=color)
                label_text = f"{vl.cget('text').split('  ')[0]}  {chr(65+col_idx)}列"
                if cell_val:
                    label_text += f" \"{cell_val[:12]}\""
                vl.configure(text=label_text, text_color=color)
            else:
                dot.configure(fg_color=C.BORDER)
                vl.configure(text=f"{vl.cget('text').split('  ')[0]}  — 待设置", text_color=C.TEXT_MUTED)

    def _open_preview_window(self):
        """弹出预览匹配结果窗口 — 使用真实加载的数据"""
        win = ctk.CTkToplevel(self)
        win.title("预览匹配结果")
        win.geometry("700x480")
        win.after(100, win.lift)
        win.grab_set()

        hb = self._cf(win)
        hb.pack(fill='x', padx=16, pady=(16, 8))
        ctk.CTkLabel(hb, text="预览确认 — 以下数据将写入目标文件",
                     font=('sans-serif', 14, 'bold'),
                     text_color=C.TEXT_TITLE).pack(side='left')

        # 检查映射是否完整
        if None in (self._map_copy, self._map_match, self._map_tkey, self._map_paste):
            empty = ctk.CTkFrame(win, fg_color='#F9FAFB', corner_radius=8)
            empty.pack(fill='both', expand=True, padx=16, pady=(0, 12))
            ctk.CTkLabel(empty, text="📋 尚未完成映射",
                         font=('sans-serif', 16, 'bold'),
                         text_color=C.TEXT_MUTED).pack(expand=True)
            ctk.CTkLabel(empty, text="请先标记源文件的复制列和匹配键，以及目标文件的匹配键和粘贴列。",
                         font=('sans-serif', 11),
                         text_color=C.TEXT_HELP).pack(expand=True, pady=(0, 20))
            ctk.CTkButton(win, text="关闭", width=80, height=34,
                          fg_color=C.PRIMARY, text_color='white',
                          hover_color=C.PRIMARY_HOVER, corner_radius=8,
                          command=win.destroy).pack(pady=(0, 16))
            return

        # 使用 rule_engine 的智能表头检测
        from rule_engine import preview_matches, ColumnMapping
        mapping_obj = ColumnMapping(
            source_value_col=self._map_copy,
            source_key_col=self._map_match,
            target_key_col=self._map_tkey,
            target_dest_col=self._map_paste,
        )
        source_dir = getattr(self, '_source_dir', '') or self._src_entry.get().strip()
        target_path = getattr(self, '_tgt_path', '') or self._tgt_entry.get().strip()

        entries = []
        preview_error = None
        if not source_dir:
            preview_error = "请先选择源文件夹（点击左侧「浏览...」选择文件夹）"
        elif not target_path:
            preview_error = "请先选择目标模板文件（点击左侧「浏览...」选择文件）"
        elif not os.path.isdir(source_dir):
            preview_error = f"源文件夹不存在: {source_dir}"
        elif not os.path.isfile(target_path):
            preview_error = f"目标文件不存在: {target_path}"
        else:
            try:
                result = preview_matches(source_dir, target_path, mapping_obj)
                entries = [(e.key, e.value, e.source_file) for e in result.entries if e.has_data]
            except Exception as e:
                preview_error = f"预览失败: {e}"

        # 显示真实数据
        stats = f"共 {len(entries)} 条"
        ctk.CTkLabel(hb, text=stats, font=('sans-serif', 11),
                     text_color=C.TEXT_HELP).pack(side='right')

        if not entries:
            empty = ctk.CTkFrame(win, fg_color='#F9FAFB', corner_radius=8)
            empty.pack(fill='both', expand=True, padx=16, pady=(0, 12))
            if preview_error:
                ctk.CTkLabel(empty, text="⚠️ " + preview_error,
                             font=('sans-serif', 13, 'bold'),
                             text_color=C.ERROR, wraplength=600).pack(expand=True, padx=20)
            else:
                ctk.CTkLabel(empty, text="📋 未找到匹配数据",
                             font=('sans-serif', 14, 'bold'),
                             text_color=C.TEXT_MUTED).pack(expand=True)
                ctk.CTkLabel(empty, text="源文件夹和目标文件中没有匹配的数据。",
                             font=('sans-serif', 11),
                             text_color=C.TEXT_HELP).pack(expand=True, pady=(0, 10))
                ctk.CTkLabel(empty,
                             text=f"源: {source_dir}\n目标: {target_path}",
                             font=('sans-serif', 9),
                             text_color=C.TEXT_MUTED).pack(expand=True, pady=(0, 10))
            ctk.CTkButton(win, text="关闭", width=80, height=34,
                          fg_color=C.PRIMARY, text_color='white',
                          hover_color=C.PRIMARY_HOVER, corner_radius=8,
                          command=win.destroy).pack(pady=(0, 16))
            return

        sf = ctk.CTkScrollableFrame(win, fg_color='#F9FAFB', corner_radius=8)
        sf.pack(fill='both', expand=True, padx=16, pady=(0, 12))

        ch = ctk.CTkFrame(sf, fg_color=C.TBL_HDR_BG, height=28, corner_radius=6)
        ch.pack(fill='x')
        ch.pack_propagate(False)
        for w, t in [(30, ''), (90, '匹配键'), (130, '来源文件'), (280, '要写入的值')]:
            ctk.CTkLabel(ch, text=t, width=w, font=('sans-serif', 9, 'bold'),
                         text_color=C.TEXT_HELP).pack(side='left', padx=(4, 0))

        self._preview_entries = []
        for i, (key, val, fname) in enumerate(entries):
            row = ctk.CTkFrame(sf, fg_color=C.BG_WHITE if i % 2 == 0 else '#FAFAFA',
                               height=30, corner_radius=6)
            row.pack(fill='x')
            row.pack_propagate(False)
            cb_var = ctk.BooleanVar(value=True)
            cb = ctk.CTkCheckBox(row, text='', variable=cb_var,
                                  checkbox_width=16, checkbox_height=16,
                                  border_width=1, corner_radius=6, width=24)
            cb.pack(side='left', padx=(4, 0))
            ctk.CTkLabel(row, text=key, width=90, font=('sans-serif', 10),
                         text_color=C.TEXT_BODY).pack(side='left', padx=(4, 0))
            ctk.CTkLabel(row, text=fname, width=130, font=('sans-serif', 10),
                         text_color=C.TEXT_HELP).pack(side='left', padx=(4, 0))
            ctk.CTkLabel(row, text=val, font=('sans-serif', 10),
                         text_color=C.TEXT_BODY).pack(side='left', fill='x', expand=True, padx=(4, 0))
            self._preview_entries.append((key, val, fname, cb_var))

        bb = self._cf(win)
        bb.pack(fill='x', padx=16, pady=(0, 16))
        ctk.CTkButton(bb, text="取消", width=80, height=34,
                      fg_color='#F3F4F6', text_color=C.TEXT_HELP,
                      corner_radius=8,
                      command=win.destroy).pack(side='right', padx=(8, 0))
        ctk.CTkButton(bb, text="✅ 确认并写入", width=150, height=34,
                      fg_color=C.PRIMARY, text_color='white',
                      hover_color=C.PRIMARY_HOVER, corner_radius=8,
                      command=lambda: self._do_write(win)).pack(side='right')

        self._step = max(self._step, 3)
        self._status_var.set("预览完成，请确认")

    def _do_write(self, win):
        """执行写入（后台线程，不阻塞界面）"""
        win.destroy()
        target = getattr(self, '_tgt_path', '')
        if not target or not getattr(self, '_preview_entries', []):
            self._status_var.set("无数据可写入")
            return

        from excel_writer import process_and_write
        from rule_engine import ColumnMapping, MatchEntry

        mapping = ColumnMapping(
            source_value_col=self._map_copy,
            source_key_col=self._map_match,
            target_key_col=self._map_tkey,
            target_dest_col=self._map_paste,
        )
        entries = []
        for key, val, fname, cb_var in self._preview_entries:
            if cb_var.get():
                e = MatchEntry(key=key, value=val, source_file=fname, source_sheet='',
                               source_row=0, target_row=0, has_data=True, selected=True)
                entries.append(e)

        if not entries:
            self._status_var.set("无选中数据")
            return

        self._status_var.set("写入中...")
        self.progress_var = ctk.StringVar(value="写入中...")

        import threading
        def run():
            try:
                output = process_and_write(target, mapping, entries,
                                            progress_callback=lambda m: self.after(0, lambda: self._act(m)))
                self.after(0, lambda: self._on_write_done(output))
            except Exception as e:
                self.after(0, lambda: self._on_write_error(str(e)))

        threading.Thread(target=run, daemon=True).start()

    def _on_write_done(self, output):
        self._act(f"写入完成: {output}")
        self._status_var.set("写入完成 ✅")
        self._step = 4
        self._sync_ui()

    def _on_write_error(self, err):
        self._act(f"写入失败: {err}")
        self._status_var.set("写入失败 ❌")

    def _open_log_window(self):
        """弹出日志查看窗口"""
        win = ctk.CTkToplevel(self)
        win.title("日志输出")
        win.geometry("650x400")
        win.after(100, win.lift)
        win.grab_set()

        tb = self._cf(win)
        tb.pack(fill='x', padx=16, pady=(16, 8))
        ctk.CTkLabel(tb, text="📋 日志输出",
                     font=('sans-serif', 14, 'bold'),
                     text_color=C.TEXT_TITLE).pack(side='left')

        def copy_all():
            win.clipboard_clear()
            win.clipboard_append(log_box.get('1.0', 'end-1c'))

        ctk.CTkButton(tb, text="📋 复制全部", width=100, height=28,
                      fg_color='#F3F4F6', text_color=C.TEXT_HELP,
                      corner_radius=6,
                      command=copy_all).pack(side='right')

        log_box = ctk.CTkTextbox(win, fg_color='#F9FAFB', text_color=C.TEXT_BODY,
                                  font=('monospace', 10), border_width=0,
                                  corner_radius=8)
        log_box.pack(fill='both', expand=True, padx=16, pady=(4, 16))

        if self._log_lines:
            log_box.insert('1.0', '\n'.join(self._log_lines))
        else:
            log_box.insert('1.0', '(暂无日志)')
        log_box.configure(state='disabled')
        self._status_var.set("已打开日志窗口")

    # ── 表格初始为空，选择文件后才加载数据 ──


def main():
    app = ExcelBridgeApp()
    app.mainloop()

if __name__ == '__main__':
    main()
