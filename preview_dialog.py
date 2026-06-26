"""
预览确认对话框 — CustomTkinter 现代风格
"""

import customtkinter as ctk
from typing import List, Callable
from rule_engine import PreviewResult, ColumnMapping, MatchEntry

# 配色 (light)
ACCENT_BLUE   = '#2563EB'
ACCENT_GREEN  = '#16A34A'
DANGER_RED    = '#DC2626'
TEXT_PRIMARY  = '#1F2937'
TEXT_SECONDARY = '#6B7280'
TEXT_MUTED    = '#9CA3AF'
BG_CARD       = '#FFFFFF'
BG_SIDEBAR    = '#F9FAFB'
BORDER        = '#D1D5DB'


class PreviewDialog(ctk.CTkToplevel):
    def __init__(self, parent, result: PreviewResult, mapping: ColumnMapping,
                 callback: Callable[[List[MatchEntry]], None]):
        super().__init__(parent)
        self.result = result
        self.mapping = mapping
        self.callback = callback
        self._check_buttons = {}

        self.title("预览确认 — 将要写入的数据")
        self.geometry("750x520")
        self.minsize(580, 350)
        self.after(100, self.lift)
        self.grab_set()

        self._build_ui()
        self._populate()
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _build_ui(self):
        main = ctk.CTkFrame(self, fg_color='transparent')
        main.pack(fill='both', expand=True, padx=16, pady=16)

        # 标题
        header = ctk.CTkFrame(main, fg_color='transparent')
        header.pack(fill='x', pady=(0, 12))
        ctk.CTkLabel(header, text="预览确认",
                     font=('sans-serif', 16, 'bold'),
                     text_color=TEXT_PRIMARY).pack(side='left')
        stats = "共 {} 条 · {} 条有数据 · {} 条无数据".format(
            len(self.result.entries), self.result.matched_count,
            self.result.no_data_count)
        ctk.CTkLabel(header, text=stats, font=('sans-serif', 11),
                     text_color=TEXT_SECONDARY).pack(side='right')

        # 列头
        ch = ctk.CTkFrame(main, fg_color=BG_SIDEBAR, corner_radius=6, height=30)
        ch.pack(fill='x', pady=(0, 4))
        ch.pack_propagate(False)
        ctk.CTkLabel(ch, text="", width=30).pack(side='left')
        ctk.CTkLabel(ch, text="匹配键", width=80,
                     font=('sans-serif', 10, 'bold'),
                     text_color=TEXT_SECONDARY).pack(side='left', padx=(0, 8))
        ctk.CTkLabel(ch, text="来源文件", width=140,
                     font=('sans-serif', 10, 'bold'),
                     text_color=TEXT_SECONDARY).pack(side='left', padx=(0, 8))
        ctk.CTkLabel(ch, text="要写入的值",
                     font=('sans-serif', 10, 'bold'),
                     text_color=TEXT_SECONDARY).pack(
            side='left', fill='x', expand=True, padx=(0, 8))
        ctk.CTkLabel(ch, text="目标行", width=55,
                     font=('sans-serif', 10, 'bold'),
                     text_color=TEXT_SECONDARY).pack(side='right')

        # 滚动列表
        self.scroll_frame = ctk.CTkScrollableFrame(
            main, fg_color=BG_SIDEBAR, corner_radius=8,
            border_width=1, border_color=BORDER)
        self.scroll_frame.pack(fill='both', expand=True, pady=(0, 12))

        # 过滤栏
        fb = ctk.CTkFrame(main, fg_color='transparent')
        fb.pack(fill='x', pady=(0, 12))
        self.filter_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(fb, text="仅显示有数据的行", variable=self.filter_var,
                        checkbox_width=18, checkbox_height=18,
                        border_width=1, corner_radius=4,
                        command=self._apply_filter).pack(side='left')
        ctk.CTkButton(fb, text="取消全选", width=80, height=28,
                      fg_color='transparent', text_color=TEXT_SECONDARY,
                      border_width=1, border_color=BORDER, corner_radius=6,
                      font=('sans-serif', 11),
                      command=self._deselect_all).pack(side='right', padx=(4, 0))
        ctk.CTkButton(fb, text="全选", width=60, height=28,
                      fg_color='transparent', text_color=TEXT_SECONDARY,
                      border_width=1, border_color=BORDER, corner_radius=6,
                      font=('sans-serif', 11),
                      command=self._select_all).pack(side='right', padx=(4, 0))

        # 底部按钮
        bb = ctk.CTkFrame(main, fg_color='transparent')
        bb.pack(fill='x')
        ctk.CTkButton(bb, text="取消", width=80, height=36,
                      fg_color='transparent', text_color=TEXT_SECONDARY,
                      border_width=1, border_color=BORDER, corner_radius=8,
                      command=self._on_cancel).pack(side='right', padx=(8, 0))
        n = sum(1 for e in self.result.entries if e.selected and e.has_data)
        ctk.CTkButton(bb, text="确认并写入 ({} 条)".format(n), width=160, height=36,
                      fg_color=ACCENT_BLUE, text_color='white',
                      hover_color='#1D4ED8', corner_radius=8,
                      font=('sans-serif', 13, 'bold'),
                      command=self._on_confirm).pack(side='right')

    def _populate(self):
        for w in self.scroll_frame.winfo_children():
            w.destroy()
        self._check_buttons.clear()
        filtered = self.filter_var.get()
        for i, entry in enumerate(self.result.entries):
            if filtered and not entry.has_data:
                continue
            self._create_entry_row(i, entry)

    def _create_entry_row(self, idx, entry):
        row = ctk.CTkFrame(self.scroll_frame, fg_color='transparent', height=32)
        row.pack(fill='x', pady=1)
        row.pack_propagate(False)

        var = ctk.BooleanVar(value=entry.selected)
        cb = ctk.CTkCheckBox(row, text='', variable=var,
                              checkbox_width=18, checkbox_height=18,
                              border_width=1, corner_radius=4, width=24,
                              command=lambda e=entry, v=var: setattr(
                                  e, 'selected', v.get()))
        cb.pack(side='left', padx=(4, 0))
        self._check_buttons[idx] = (cb, var)

        kc = TEXT_PRIMARY if entry.has_data else TEXT_MUTED
        ctk.CTkLabel(row, text=entry.key, width=80, anchor='w',
                     font=('sans-serif', 11), text_color=kc).pack(
            side='left', padx=(0, 8))
        ctk.CTkLabel(row, text=entry.source_file, width=140, anchor='w',
                     font=('sans-serif', 11),
                     text_color=TEXT_SECONDARY).pack(side='left', padx=(0, 8))
        vc = TEXT_PRIMARY if entry.has_data else TEXT_MUTED
        dv = entry.value if entry.has_data else '(无数据)'
        ctk.CTkLabel(row, text=dv, anchor='w', font=('sans-serif', 11),
                     text_color=vc).pack(side='left', fill='x', expand=True,
                                          padx=(0, 8))
        tt = str(entry.target_row) if entry.target_row > 0 else '—'
        tc = ACCENT_GREEN if entry.target_row > 0 else DANGER_RED
        ctk.CTkLabel(row, text=tt, width=55, anchor='center',
                     font=('sans-serif', 11), text_color=tc).pack(side='right')

    def _apply_filter(self):
        self._populate()

    def _select_all(self):
        for e in self.result.entries:
            if e.has_data:
                e.selected = True
        self._populate()

    def _deselect_all(self):
        for e in self.result.entries:
            e.selected = False
        self._populate()

    def _on_confirm(self):
        self.callback(self.result.entries)
        self.destroy()

    def _on_cancel(self):
        self.callback(None)
        self.destroy()
