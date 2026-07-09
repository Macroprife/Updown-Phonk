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

TOOLS = [
    # (分类, 名称, 描述, 相对路径, 类型: "cli" 或 "gui")
    # ── 目录 / 数据 ──
    ("📊 目录与数据", "合并表格",
     "扫描文件夹中所有 Excel，提取卷内目录并生成汇总统计表",
     "目录类/合并表格（New）.py", "cli"),

    ("📊 目录与数据", "生成成品表",
     "将合并表格输出转为案卷级 + 文件级成品表（仿劳动桥模板）",
     "目录类/表格处理转换成品表.py", "cli"),

    ("📊 目录与数据", "统计 PDF 与图片",
     "分别统计图片和 PDF 文件的数量、大小，生成 Excel 报告",
     "目录类/统计PDF与图片.py", "gui"),

    # ── JPG 处理 ──
    ("🖼️ JPG 处理", "图片转 PDF",
     "将每个子文件夹中的图片批量合并为 PDF",
     "JPG类/转PDF.py", "cli"),

    ("🖼️ JPG 处理", "图片分割",
     "根据 Excel 配置按页数将图片批量拆分到不同目录",
     "JPG类/Split(增加错误输出）.py", "gui"),

    # ── PDF 处理 ──
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


class LauncherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("档案处理工具集")
        self.root.geometry("680x620")
        self.root.resizable(True, True)
        self.root.configure(bg="#f5f6fa")

        # 找到有控制台的 python.exe
        # （双击运行时 sys.executable 是 pythonw.exe，CLI 工具需要 python.exe）
        self.python_exe = self._find_console_python()

        style = ttk.Style()
        style.theme_use("clam")

        self.build_ui()

    def build_ui(self):
        main = tk.Frame(self.root, bg="#f5f6fa", padx=24, pady=20)
        main.pack(fill=tk.BOTH, expand=True)

        # 标题
        tk.Label(main, text="📁 档案处理工具集",
                 font=("微软雅黑", 20, "bold"),
                 fg="#2c3e50", bg="#f5f6fa").pack(pady=(0, 6))

        tk.Label(main, text="选择一个工具，在新窗口中独立运行",
                 font=("微软雅黑", 10), fg="#7f8c8d", bg="#f5f6fa").pack(pady=(0, 18))

        # 按分类分组
        from collections import OrderedDict
        groups = OrderedDict()
        for cat, name, desc, path, kind in TOOLS:
            groups.setdefault(cat, []).append((name, desc, path, kind))

        # 画布 + 滚动（工具多了不溢出）
        canvas = tk.Canvas(main, bg="#f5f6fa", highlightthickness=0)
        scrollbar = ttk.Scrollbar(main, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg="#f5f6fa")

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

        row = 0
        for cat, items in groups.items():
            # 分类标题
            tk.Label(scroll_frame, text=cat,
                     font=("微软雅黑", 13, "bold"),
                     fg="#34495e", bg="#f5f6fa", anchor="w").grid(
                row=row, column=0, sticky="w", pady=(18 if row > 0 else 0, 4), padx=4)
            row += 1

            # 分隔线
            sep = tk.Frame(scroll_frame, height=1, bg="#dcdde1")
            sep.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(0, 8))
            row += 1

            for name, desc, path, kind in items:
                card = tk.Frame(scroll_frame, bg="white", relief="groove", bd=1,
                                padx=14, pady=10)
                card.grid(row=row, column=0, sticky="ew", pady=4, padx=4)
                card.columnconfigure(0, weight=1)

                tk.Label(card, text=name,
                         font=("微软雅黑", 11, "bold"),
                         fg="#2c3e50", bg="white", anchor="w").grid(
                    row=0, column=0, sticky="w")

                tk.Label(card, text=desc,
                         font=("微软雅黑", 9), fg="#7f8c8d", bg="white",
                         anchor="w", wraplength=500, justify="left").grid(
                    row=1, column=0, sticky="w", pady=(2, 6))

                btn = tk.Button(card, text="▶ 启动",
                                font=("微软雅黑", 9, "bold"),
                                bg="#3498db", fg="white",
                                activebackground="#2980b9",
                                cursor="hand2", bd=0, padx=16, pady=4,
                                command=lambda p=path, k=kind, n=name: self.launch(p, k, n))
                btn.grid(row=2, column=0, sticky="w")

                row += 1

        # 底部
        tk.Label(main, text=f"工具目录：{BASE_DIR}",
                 font=("微软雅黑", 8), fg="#bdc3c7", bg="#f5f6fa").pack(side=tk.BOTTOM, pady=(10, 0))

    # ── 辅助方法 ──

    @staticmethod
    def _find_console_python():
        """返回 python.exe（非 pythonw.exe），供 CLI 子进程使用。"""
        if sys.platform != 'win32':
            return sys.executable
        # 常见情况：pythonw.exe → python.exe 在同目录下
        py_exe = sys.executable.replace('pythonw.exe', 'python.exe')
        if os.path.exists(py_exe):
            return py_exe
        # 备选：PATH 中的 python
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
            # Windows: 新控制台窗口。必须用 python.exe（有控制台），
            # 因为双击运行时 sys.executable 是 pythonw.exe（无控制台），
            # 用 pythonw.exe 会导致输出不可见、input() 无效。
            subprocess.Popen(
                ['cmd', '/k', self.python_exe, script],
                cwd=BASE_DIR,
                creationflags=subprocess.CREATE_NEW_CONSOLE
            )


def main():
    # ── 启动前检查依赖 ──
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
