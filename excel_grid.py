"""
Excel 表格网格组件
基于 ttk.Treeview，支持点击选单元格、多色高亮
"""

import tkinter as tk
from tkinter import ttk
from typing import List, Any, Optional, Tuple, Callable


# ── 亮色主题配色 ──
COLORS = {
    'bg_main':       '#F3F4F6',
    'bg_even':       '#FFFFFF',
    'bg_odd':        '#F9FAFB',
    'bg_selected':   '#DBEAFE',
    'fg_primary':    '#1F2937',
    'fg_secondary':  '#6B7280',
    'fg_selected':   '#1E40AF',
    'border':        '#D1D5DB',
    'header_bg':     '#E5E7EB',
    'header_fg':     '#374151',
}

# 高亮标签
HIGHLIGHT_TAGS = {
    'source_value': {'bg': '#DBEAFE', 'fg': '#1E40AF'},   # 蓝
    'source_key':   {'bg': '#FFEDD5', 'fg': '#9A3412'},   # 橙
    'target_key':   {'bg': '#DCFCE7', 'fg': '#166534'},   # 绿
    'target_dest':  {'bg': '#FEF9C3', 'fg': '#854D0E'},   # 黄
}


def _col_letter(col_idx: int) -> str:
    result = ''
    n = col_idx
    while n >= 0:
        result = chr(ord('A') + (n % 26)) + result
        n = n // 26 - 1
    return result


def apply_light_style():
    """配置 ttk 亮色主题样式"""
    style = ttk.Style()
    if 'clam' in style.theme_names():
        style.theme_use('clam')

    style.configure('Light.Treeview',
                    background=COLORS['bg_even'],
                    foreground=COLORS['fg_primary'],
                    fieldbackground=COLORS['bg_even'],
                    borderwidth=1,
                    rowheight=26)
    style.configure('Light.Treeview.Heading',
                    background=COLORS['header_bg'],
                    foreground=COLORS['header_fg'],
                    borderwidth=1,
                    relief='flat',
                    font=('sans-serif', 9, 'bold'))
    style.map('Light.Treeview.Heading',
              background=[('active', '#D1D5DB')])
    style.map('Light.Treeview',
              background=[('selected', COLORS['bg_selected'])],
              foreground=[('selected', COLORS['fg_selected'])])

    # 行样式
    style.configure('even.Treeview', background=COLORS['bg_even'])
    style.configure('odd.Treeview', background=COLORS['bg_odd'])

    # 高亮标签
    for tag_name, tag_colors in HIGHLIGHT_TAGS.items():
        style_name = tag_name + '.Treeview'
        style.configure(style_name,
                        background=tag_colors['bg'],
                        foreground=tag_colors['fg'])

    return style


class ExcelGrid(ttk.Frame):
    """Excel 表格网格组件。"""

    def __init__(self, parent, max_display_rows=200, max_display_cols=20, **kwargs):
        super().__init__(parent, **kwargs)
        self.max_display_rows = max_display_rows
        self.max_display_cols = max_display_cols
        self._data: List[List[str]] = []
        self._highlighted: List[Tuple[int, int, str]] = []
        self._row_count = 0
        self._col_count = 0
        self.on_cell_click: Optional[Callable[[int, int], None]] = None
        self._last_click_info: Optional[Tuple[int, int, str]] = None

        self._build_ui()

    def _build_ui(self):
        self.v_scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL)
        self.v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.h_scrollbar = ttk.Scrollbar(self, orient=tk.HORIZONTAL)
        self.h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree = ttk.Treeview(
            self,
            yscrollcommand=self.v_scrollbar.set,
            xscrollcommand=self.h_scrollbar.set,
            selectmode='browse',
            show='tree headings',
            style='Light.Treeview',
        )
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.v_scrollbar.config(command=self.tree.yview)
        self.h_scrollbar.config(command=self.tree.xview)

        self.tree.bind('<Button-1>', self._on_click)
        self._context_menu = tk.Menu(self, tearoff=0)
        self._context_menu.add_command(label="清除高亮",
                                        command=self.clear_highlights)
        self.tree.bind('<Button-3>', self._on_right_click)

    def load_data(self, rows: List[List[Any]], sheet_name: str = ""):
        self._data = rows
        self._row_count = min(len(rows), self.max_display_rows)

        if rows:
            max_cols = max((len(r) for r in rows[:self._row_count]))
            while max_cols > 0:
                has_data = any(
                    len(r) > max_cols - 1 and str(r[max_cols - 1]).strip()
                    for r in rows[:self._row_count]
                )
                if has_data:
                    break
                max_cols -= 1
            self._col_count = min(max(max_cols, 1), self.max_display_cols)
        else:
            self._col_count = 0

        self._highlighted.clear()
        self.tree.delete(*self.tree.get_children())
        self.tree['columns'] = []

        if not rows or self._col_count == 0:
            return

        col_ids = [_col_letter(i) for i in range(self._col_count)]
        self.tree['columns'] = col_ids

        self.tree.column('#0', width=48, minwidth=40, stretch=False,
                         anchor='center')
        self.tree.heading('#0', text='')

        for col_id in col_ids:
            self.tree.column(col_id, width=88, minwidth=60, stretch=True,
                             anchor='w')
            self.tree.heading(col_id, text=col_id)

        for row_idx in range(self._row_count):
            row_data = rows[row_idx]
            values = [str(row_data[c]) if c < len(row_data) and row_data[c] is not None else ''
                      for c in range(self._col_count)]
            tag = 'even.Treeview' if row_idx % 2 == 0 else 'odd.Treeview'
            self.tree.insert('', tk.END, iid=str(row_idx),
                             text=str(row_idx + 1), values=values, tags=(tag,))

    def _on_click(self, event):
        row_iid = self.tree.identify_row(event.y)
        if not row_iid:
            return
        try:
            row_idx = int(row_iid)
        except ValueError:
            return
        if row_idx < 0 or row_idx >= self._row_count:
            return

        col_str = self.tree.identify_column(event.x)
        region = self.tree.identify_region(event.x, event.y)
        col_idx = -1

        if col_str.startswith('#'):
            col_num = int(col_str.replace('#', ''))
            col_idx = col_num - 1
        else:
            for i in range(self._col_count):
                if _col_letter(i) == col_str:
                    col_idx = i
                    break
        if col_idx < 0 and region == 'tree':
            col_idx = 0
        if col_idx < 0 or col_idx >= self._col_count:
            return

        cv = ''
        if col_idx < len(self._data[row_idx]):
            cv = str(self._data[row_idx][col_idx]).strip()
        self._last_click_info = (row_idx, col_idx, cv)
        if self.on_cell_click:
            self.on_cell_click(row_idx, col_idx)

    def _on_right_click(self, event):
        try:
            self._context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self._context_menu.grab_release()

    def highlight_cell(self, row_idx, col_idx, tag):
        if tag not in HIGHLIGHT_TAGS or row_idx >= self._row_count:
            return
        self._highlighted.append((row_idx, col_idx, tag))
        try:
            iid = str(row_idx)
            item = self.tree.item(iid)
            tags = list(item.get('tags', ()))
            tags = [t for t in tags if t not in HIGHLIGHT_TAGS
                    and t not in ('even.Treeview', 'odd.Treeview')]
            tags.append(tag + '.Treeview')
            self.tree.item(iid, tags=tuple(tags))
            self.tree.see(iid)
        except Exception:
            pass

    def clear_highlights(self):
        for row_idx, _col_idx, tag in self._highlighted:
            try:
                iid = str(row_idx)
                item = self.tree.item(iid)
                tags = list(item.get('tags', ()))
                tags = [t for t in tags if tag + '.Treeview' not in t]
                if not any(t.startswith(tn) for t in tags
                          for tn in HIGHLIGHT_TAGS):
                    base = 'even.Treeview' if row_idx % 2 == 0 else 'odd.Treeview'
                    tags = [t for t in tags
                            if t not in ('even.Treeview', 'odd.Treeview')]
                    tags.append(base)
                self.tree.item(iid, tags=tuple(tags))
            except Exception:
                pass
        self._highlighted.clear()

    def get_cell_value(self, row_idx, col_idx):
        if 0 <= row_idx < len(self._data) and 0 <= col_idx < len(self._data[row_idx]):
            val = self._data[row_idx][col_idx]
            return str(val) if val is not None else ''
        return ''

    def scroll_to_row(self, row_idx):
        if 0 <= row_idx < self._row_count:
            self.tree.see(str(row_idx))
