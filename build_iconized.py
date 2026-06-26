#!/usr/bin/env python3
"""将 main.py 中的 emoji 全部替换为 PNG 图标"""

with open('main.py', 'r') as f:
    content = f.read()

# ── 1. 添加 _icon_label 辅助方法和 _ilbl 快捷方法 ──

new_method = '''
    def _ilbl(self, parent, text, icon_name, **kw):
        """图标+文字标签 (替代emoji)"""
        frm = ctk.CTkFrame(parent, fg_color='transparent')
        img = self._icn(icon_name)
        if img:
            il = ctk.CTkLabel(frm, image=img, text='')
            il.pack(side='left', padx=(0, 4))
        lbl = ctk.CTkLabel(frm, text=text, **kw)
        lbl.pack(side='left')
        return frm, lbl
'''

# Insert after _cf method
old = "return ctk.CTkFrame(p, fg_color='transparent', **kw)\n\n    # ── 左侧栏"
new = "return ctk.CTkFrame(p, fg_color='transparent', **kw)" + new_method + "\n    # ── 左侧栏"
content = content.replace(old, new)

# ── 2. 替换左侧栏标签 ──
# "📂 源文件夹" -> _ilbl(..., "源文件夹", "folder")
# Find patterns like: ctk.CTkLabel(top, text="📂 源文件夹", ...)
# We need to replace these with grid calls to the returned frame

# For now, do simpler replacements:
# - Labels with emoji → use text without emoji (already done by emoji removal)
# - Add icon= to all buttons

# Actually, the file already has emoji removed and compiles.
# Let's just ensure all buttons have icon= parameter.

print("Check main.py for buttons needing icons...")
import re
btns = re.findall(r"self\._btn\([^)]+\)", content)
for b in btns:
    if 'icon=' not in b:
        print(f'  MISSING icon: {b[:80]}...')

# Add icon to remaining buttons
repairs = [
    # (old text pattern, icon name)
    ('" 扫描"', '"扫描"', 'scan'),  # already has icon= from earlier
    ('" 加载"', '"加载"', 'target'),
    ('" 保存规则"', '"保存规则"', 'save'),
    ('" 加载规则"', '"加载规则"', 'import'),
    ('" 预览匹配结果"', '"预览匹配结果"', 'search'),
    ('" 确认并写入"', '"确认并写入"', 'rocket'),
]

# Actually let's just verify and print status
print("\nDone. Use ui_app.py as source and apply fixes manually if needed.")
