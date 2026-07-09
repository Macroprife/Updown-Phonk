"""
档案处理工具 — 统一启动菜单
双击运行，所有子工具在新窗口中独立运行，互不干扰。
"""
import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk

# ═══════════════════════════════════════════
# 依赖检查
# ═══════════════════════════════════════════
REQUIRED_PACKAGES = {
    'pandas':    'pip install pandas',
    'openpyxl':  'pip install openpyxl',
    'xlrd':      'pip install xlrd',
    'PIL':       'pip install Pillow',
    'fitz':      'pip install PyMuPDF',
    'PyPDF2':    'pip install PyPDF2',
}


def _check_dependencies():
    """返回缺失的包列表 [(模块名, 安装命令), ...]。"""
    missing = []
    for mod, install_cmd in REQUIRED_PACKAGES.items():
        try:
            __import__(mod)
        except ImportError:
            missing.append((mod, install_cmd))
    return missing

# ═══════════════════════════════════════════
# 工具注册表（只需维护这一个列表）
# ═══════════════════════════════════════════
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 每个分类的强调色
CAT_ACCENT = {
    "📊 目录与数据": "#58a6ff",   # 蓝
    "🖼️ JPG 处理":   "#f778ba",   # 粉
    "📄 PDF 处理":    "#3fb950",   # 绿
}

TOOLS = [
    ("📊 目录与数据", "合并表格",
     "扫描文件夹中所有 Excel，提取卷内目录并生成汇总统计表",
     "目录类/合并表格（New）.py", "cli"),

    ("📊 目录与数据", "生成成品表",
     "将合并表格输出转为案卷级 + 文件级成品表（仿劳动桥模板）",
     "目录类/表格处理转换成品表.py", "cli"),

    ("📊 目录与数据", "统计 PDF 与图片",
     "分别统计图片和 PDF 文件的数量、大小，生成 Excel 报告",
     "目录类/统计PDF与图片.py", "gui"),

    ("🖼️ JPG 处理", "图片转 PDF",
     "将每个子文件夹中的图片批量合并为 PDF",
     "JPG类/转PDF.py", "cli"),

    ("🖼️ JPG 处理", "图片分割",
     "根据 Excel 配置按页数将图片批量拆分到不同目录",
     "JPG类/Split(增加错误输出）.py", "gui"),

    ("📄 PDF 处理", "PDF 转图片",
     "批量将 PDF 转换为 JPG/PNG 图片",
     "PDF类/转JPG.py", "cli"),

    ("📄 PDF 处理", "PDF 删页",
     "批量删除 PDF 的第一页或首尾页",
     "PDF类/PDF删页.py", "gui"),

    ("📄 PDF 处理", "PDF 分割",
     "根据 Excel 配置将 PDF 按页数拆分为多个文件",
     "PDF类/Split(未测试).py", "gui"),

    ("📄 PDF 处理", "复制不同名文件",
     "比较两个文件夹，复制互不匹配的 PDF 文件到新目录",
     "PDF类/复制不同名文件.py", "cli"),

    ("📄 PDF 处理", "PDF 层级迁移",
     "去掉 PDF 文件路径中的倒数第二级目录",
     "PDF类/迁移.py", "cli"),

    ("📄 PDF 处理", "扫描空文件夹",
     "递归扫描并列出所有空文件夹",
     "PDF类/扫描空文件夹.py", "cli"),

    ("📄 PDF 处理", "提取文件名",
     "批量提取文件名到 Excel（含规范化清理）",
     "PDF类/提取文件名.py", "cli"),

    ("📄 PDF 处理", "统计 PDF 总页数",
     "递归统计文件夹中所有 PDF 的总页数",
     "PDF类/统计总数.py", "cli"),
]


# ═══════════════════════════════════════════
# 配色系统
# ═══════════════════════════════════════════
C = {
    "bg":         "#0d1117",
    "card_bg":    "#161b22",
    "card_hover": "#1c2333",
    "card_brd":   "#21262d",
    "sep":        "#21262d",
    "title":      "#e6edf3",
    "subtitle":   "#8b949e",
    "text":       "#c9d1d9",
    "muted":      "#484f58",
    "btn_bg":     "#1f6feb",
    "btn_hover":  "#388bfd",
    "btn_text":   "#ffffff",
    "cli_tag_bg": "#21262d",
    "cli_tag_fg": "#8b949e",
    "gui_tag_bg": "#1a3a2a",
    "gui_tag_fg": "#3fb950",
}


class LauncherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("档案处理工具集")
        self.root.geometry("720x660")
        self.root.configure(bg=C["bg"])
        self.root.minsize(580, 500)

        self.python_exe = self._find_console_python()
        self._cards = []  # 用于悬浮动画

        style = ttk.Style()
        style.theme_use("clam")

        self.build_ui()
        self._animate_title()

    # ── UI 构建 ──

    def build_ui(self):
        outer = tk.Frame(self.root, bg=C["bg"])
        outer.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)

        # ── 顶栏 ──
        header = tk.Frame(outer, bg=C["bg"], height=130)
        header.pack(fill=tk.X)
        header.pack_propagate(False)

        # Logo 装饰线
        self._accent_line = tk.Canvas(header, bg=C["bg"], height=3,
                                       highlightthickness=0)
        self._accent_line.pack(fill=tk.X, side=tk.TOP)

        self._title_label = tk.Label(
            header, text="📁  档案处理工具集",
            font=("微软雅黑", 24, "bold"),
            fg=C["title"], bg=C["bg"],
        )
        self._title_label.pack(pady=(24, 4))

        tk.Label(header, text="选择一个工具，在新窗口中独立运行",
                 font=("微软雅黑", 10), fg=C["subtitle"], bg=C["bg"]
                 ).pack(pady=(0, 0))

        # 工具计数
        total = sum(1 for t in TOOLS)
        cli_count = sum(1 for t in TOOLS if t[4] == "cli")
        gui_count = sum(1 for t in TOOLS if t[4] == "gui")
        tk.Label(header,
                 text=f"共 {total} 个工具  ·  🖥 CLI ×{cli_count}  ·  🪟 GUI ×{gui_count}",
                 font=("微软雅黑", 8), fg=C["muted"], bg=C["bg"]
                 ).pack(pady=(2, 0))

        # 分隔线
        sep_line = tk.Frame(outer, height=1, bg=C["sep"])
        sep_line.pack(fill=tk.X)

        # ── 内容区 ──
        content = tk.Frame(outer, bg=C["bg"])
        content.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)

        canvas = tk.Canvas(content, bg=C["bg"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(content, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=C["bg"])

        scroll_frame.bind("<Configure>",
                          lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 鼠标滚轮
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # 按分类分组
        from collections import OrderedDict
        groups = OrderedDict()
        for cat, name, desc, path, kind in TOOLS:
            groups.setdefault(cat, []).append((name, desc, path, kind))

        row = 0
        self._cards = []

        for cat, items in groups.items():
            accent = CAT_ACCENT.get(cat, C["btn_bg"])

            # 分类标题 — 左边加一条色带
            cat_frame = tk.Frame(scroll_frame, bg=C["bg"])
            cat_frame.grid(row=row, column=0, sticky="ew",
                           pady=(20 if row > 0 else 8, 4), padx=16)

            # 色带
            strip = tk.Frame(cat_frame, bg=accent, width=3, height=20)
            strip.pack(side=tk.LEFT, padx=(0, 8))

            tk.Label(cat_frame,
                     text=f"{cat}  ({len(items)})",
                     font=("微软雅黑", 12, "bold"),
                     fg=C["title"], bg=C["bg"], anchor="w"
                     ).pack(side=tk.LEFT)
            row += 1

            # 工具卡片
            for name, desc, path, kind in items:
                card = self._make_card(scroll_frame, name, desc, path, kind, accent)
                card.grid(row=row, column=0, sticky="ew", pady=3, padx=16)
                self._cards.append(card)
                row += 1

        # 底部留白
        tk.Frame(scroll_frame, bg=C["bg"], height=24).grid(row=row, column=0)
        row += 1

        # 底部状态栏
        footer = tk.Frame(outer, bg=C["card_bg"], height=28)
        footer.pack(fill=tk.X, side=tk.BOTTOM)
        footer.pack_propagate(False)
        tk.Label(footer, text=f"▸ {BASE_DIR}",
                 font=("Consolas", 8), fg=C["muted"], bg=C["card_bg"]
                 ).pack(side=tk.LEFT, padx=14, pady=4)

    def _make_card(self, parent, name, desc, path, kind, accent):
        """创建一张工具卡片 Frame。"""
        card = tk.Frame(parent, bg=C["card_bg"], bd=0,
                        highlightbackground=C["card_brd"],
                        highlightthickness=1,
                        padx=14, pady=10)
        card.columnconfigure(0, weight=1)

        # 顶行：名称 + 标签
        top = tk.Frame(card, bg=C["card_bg"])
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(0, weight=1)

        tk.Label(top, text=name,
                 font=("微软雅黑", 11, "bold"),
                 fg=C["title"], bg=C["card_bg"], anchor="w"
                 ).grid(row=0, column=0, sticky="w")

        # 类型标签
        if kind == "cli":
            tag_bg, tag_fg, tag_text = C["cli_tag_bg"], C["cli_tag_fg"], "CLI"
        else:
            tag_bg, tag_fg, tag_text = C["gui_tag_bg"], C["gui_tag_fg"], "GUI"
        tag = tk.Frame(top, bg=tag_bg, padx=8, pady=1, bd=0)
        tag.grid(row=0, column=1, sticky="e", padx=(8, 0))
        tk.Label(tag, text=tag_text,
                 font=("Consolas", 8, "bold"),
                 fg=tag_fg, bg=tag_bg
                 ).pack()

        # 描述
        tk.Label(card, text=desc,
                 font=("微软雅黑", 9), fg=C["subtitle"], bg=C["card_bg"],
                 anchor="w", wraplength=520, justify="left"
                 ).grid(row=1, column=0, sticky="w", pady=(4, 8))

        # 启动按钮
        btn = tk.Button(card, text="▶  启动",
                        font=("微软雅黑", 9, "bold"),
                        bg=C["btn_bg"], fg=C["btn_text"],
                        activebackground=C["btn_hover"],
                        activeforeground=C["btn_text"],
                        cursor="hand2", bd=0, padx=20, pady=5,
                        relief="flat",
                        command=lambda p=path, k=kind, n=name: self.launch(p, k, n))
        btn.grid(row=2, column=0, sticky="w")

        # 悬浮效果
        for widget in [card, top, btn]:
            widget.bind("<Enter>",
                        lambda e, c=card, a=accent: self._on_card_enter(c, a))
            widget.bind("<Leave>",
                        lambda e, c=card: self._on_card_leave(c))
        # 递归绑定子元素
        for child in card.winfo_children():
            if isinstance(child, tk.Frame):
                for sub in child.winfo_children():
                    sub.bind("<Enter>",
                             lambda e, c=card, a=accent: self._on_card_enter(c, a))
                    sub.bind("<Leave>",
                             lambda e, c=card: self._on_card_leave(c))
            child.bind("<Enter>",
                       lambda e, c=card, a=accent: self._on_card_enter(c, a))
            child.bind("<Leave>",
                       lambda e, c=card: self._on_card_leave(c))

        return card

    def _on_card_enter(self, card, accent):
        card.configure(bg=C["card_hover"])
        for child in card.winfo_children():
            self._bg_recurse(child, C["card_hover"])
        card.configure(highlightbackground=accent, highlightthickness=1)

    def _on_card_leave(self, card):
        card.configure(bg=C["card_bg"])
        for child in card.winfo_children():
            self._bg_recurse(child, C["card_bg"])
        card.configure(highlightbackground=C["card_brd"], highlightthickness=1)

    def _bg_recurse(self, widget, color):
        """递归改背景色，跳过按钮。"""
        try:
            if not isinstance(widget, tk.Button):
                widget.configure(bg=color)
        except Exception:
            pass
        for child in widget.winfo_children():
            self._bg_recurse(child, color)

    # ── 标题微弱呼吸动画 ──

    def _animate_title(self):
        import math
        t = getattr(self, "_anim_tick", 0) + 0.04
        self._anim_tick = t
        # 色带从左到右移动
        w = self._accent_line.winfo_width()
        if w > 2:
            self._accent_line.delete("all")
            # 渐变条
            for i in range(3):
                x = (w / 3) * i + (math.sin(t + i * 2.1) * w * 0.15)
                self._accent_line.create_rectangle(
                    x, 0, x + w / 3, 3,
                    fill=["#58a6ff", "#f778ba", "#3fb950"][i],
                    outline="",
                )
        self._title_anim = self.root.after(40, self._animate_title)

    # ── 辅助方法 ──

    @staticmethod
    def _find_console_python():
        """返回 python.exe（非 pythonw.exe），供 CLI 子进程使用。"""
        if sys.platform != 'win32':
            return sys.executable
        py_exe = sys.executable.replace('pythonw.exe', 'python.exe')
        if os.path.exists(py_exe):
            return py_exe
        return 'python'

    def _monitor_subprocess(self, proc, tool_name):
        """监控 GUI 子进程：2 秒内异常退出则弹窗显示错误。"""
        def _check():
            ret = proc.poll()
            if ret is not None and ret != 0:
                from tkinter import messagebox
                stderr_text = ""
                if proc.stderr:
                    stderr_text = proc.stderr.read()
                msg = f"「{tool_name}」启动失败 (退出码 {ret})"
                if stderr_text.strip():
                    msg += f"\n\n{stderr_text.strip()[:500]}"
                messagebox.showerror("启动失败", msg)
        self.root.after(2000, _check)

    def launch(self, rel_path, kind, tool_name=""):
        script = os.path.join(BASE_DIR, rel_path)
        if not os.path.exists(script):
            from tkinter import messagebox
            messagebox.showerror("错误", f"找不到脚本：\n{script}")
            return

        if kind == "gui":
            proc = subprocess.Popen(
                [sys.executable, script],
                cwd=os.path.dirname(script),
                stderr=subprocess.PIPE, text=True,
            )
            self._monitor_subprocess(proc, tool_name)
        else:
            subprocess.Popen(
                ['cmd', '/k', self.python_exe, script],
                cwd=BASE_DIR,
                creationflags=subprocess.CREATE_NEW_CONSOLE
            )


def main():
    missing = _check_dependencies()
    if missing:
        from tkinter import messagebox
        lines = ["以下 Python 包未安装，请先安装后重试：", ""]
        for _mod, cmd in missing:
            lines.append(f"  {cmd}")
        lines.append("")
        lines.append("或一次性安装全部：")
        lines.append(f"  pip install -r {os.path.join(BASE_DIR, 'requirements.txt')}")
        messagebox.showerror("缺少依赖", "\n".join(lines))
        return

    root = tk.Tk()
    try:
        root.iconbitmap(default="")
    except Exception:
        pass
    LauncherApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
