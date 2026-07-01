"""
Updown-Phonk GUI 基类
=====================
抽取 6 个 tkinter GUI 程序中的重复代码，提供：
  - 窗口居中、样式统一
  - 文件夹/文件选择对话框
  - 带颜色标签的日志输出
  - 进度条 + 按钮繁忙态管理
  - 常用 UI 组件工厂方法

使用方式：
    class MyApp(BaseApp):
        def __init__(self, root):
            super().__init__(root, title="我的工具", geometry="600x500")
            self.build_ui()

        def build_ui(self):
            # 用提供的辅助方法搭建界面
            ...
"""

import os
import tkinter as tk
from tkinter import filedialog, ttk
from datetime import datetime


class BaseApp:
    """GUI 基类 —— 提供所有 tkinter 程序公用的基础设施。

    子类只需在 build_ui() 中组合界面组件，不必重复写浏览、日志、进度条。
    """

    # ── 默认样式常量（子类可覆盖） ──────────────────────────
    FONT_DEFAULT = ("楷体", 10)
    FONT_TITLE   = ("楷体", 12, "bold")
    FONT_LOG     = ("Consolas", 9)
    BG_COLOR     = "#F0F0F0"
    THEME        = "clam"

    # ── 生命周期 ────────────────────────────────────────────

    def __init__(self, root, *, title="Updown-Phonk", geometry="650x500",
                 resizable=(False, False)):
        self.root = root
        self.root.title(title)
        self.root.geometry(geometry)
        self.root.resizable(*resizable)

        # 缓存常用控件，子类可直接访问
        self.log_text    = None   # Text 日志框
        self.progress    = None   # ttk.Progressbar
        self.btn_primary = None   # 主要的"开始"按钮

        self._setup_styles()
        self.center_window()

    # ── 窗口工具 ────────────────────────────────────────────

    def center_window(self):
        """使窗口在屏幕居中"""
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        x = (self.root.winfo_screenwidth()  - w) // 2
        y = (self.root.winfo_screenheight() - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _setup_styles(self):
        """全局 ttk 样式 —— 子类可 override 添加更多样式"""
        style = ttk.Style()
        style.theme_use(self.THEME)

        style.configure("Title.TLabel",   font=self.FONT_TITLE)
        style.configure("Default.TLabel", font=self.FONT_DEFAULT)
        style.configure("Default.TButton", font=self.FONT_DEFAULT)

    # ── 控件工厂 ────────────────────────────────────────────

    def add_title(self, parent, text, row=0, column=0, **grid_kw):
        """大标题"""
        label = ttk.Label(parent, text=text, style="Title.TLabel")
        label.grid(row=row, column=column, **grid_kw)
        return label

    def add_label(self, parent, text, **grid_kw):
        """普通标签"""
        label = ttk.Label(parent, text=text, style="Default.TLabel")
        label.grid(**grid_kw)
        return label

    def add_log_area(self, parent, row, column, rowspan=1, columnspan=1,
                     height=10, **grid_kw):
        """带滚动条的日志文本框，返回 (Text, Scrollbar)"""
        frame = ttk.Frame(parent)
        frame.grid(row=row, column=column, rowspan=rowspan,
                   columnspan=columnspan, sticky=(tk.W, tk.E, tk.N, tk.S), **grid_kw)

        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        text = tk.Text(frame, height=height, wrap=tk.WORD,
                       font=self.FONT_LOG, padx=4, pady=4)
        text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar.config(command=text.yview)
        text.config(yscrollcommand=scrollbar.set)

        # 预注册颜色标签
        text.tag_config("success", foreground="green")
        text.tag_config("error",   foreground="red")
        text.tag_config("warning", foreground="orange")
        text.tag_config("info",    foreground="blue")

        self.log_text = text
        return text, scrollbar

    def add_progress_bar(self, parent, row, column, **grid_kw):
        """不确定模式进度条"""
        bar = ttk.Progressbar(parent, mode="indeterminate", length=400)
        bar.grid(row=row, column=column, sticky=(tk.W, tk.E), **grid_kw)
        self.progress = bar
        return bar

    def add_browse_folder_row(self, parent, label_text, var,
                              row=0, column=0, entry_width=50):
        """一行：标签 + Entry + 浏览按钮"""
        frame = ttk.Frame(parent)
        frame.grid(row=row, column=column, sticky=(tk.W, tk.E), pady=3)

        ttk.Label(frame, text=label_text, style="Default.TLabel",
                  width=12, anchor=tk.W).pack(side=tk.LEFT)

        entry = ttk.Entry(frame, textvariable=var, font=self.FONT_DEFAULT,
                          width=entry_width)
        entry.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)

        btn = ttk.Button(frame, text="浏览...",
                         command=lambda: self._browse_folder(var))
        btn.pack(side=tk.LEFT)

        return frame, entry, btn

    def add_browse_file_row(self, parent, label_text, var,
                            filetypes=None, row=0, column=0, entry_width=50):
        """一行：标签 + Entry + 浏览文件按钮"""
        if filetypes is None:
            filetypes = [("所有文件", "*.*")]

        frame = ttk.Frame(parent)
        frame.grid(row=row, column=column, sticky=(tk.W, tk.E), pady=3)

        ttk.Label(frame, text=label_text, style="Default.TLabel",
                  width=12, anchor=tk.W).pack(side=tk.LEFT)

        entry = ttk.Entry(frame, textvariable=var, font=self.FONT_DEFAULT,
                          width=entry_width)
        entry.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)

        btn = ttk.Button(frame, text="浏览...",
                         command=lambda: self._browse_file(var, filetypes))
        btn.pack(side=tk.LEFT)

        return frame, entry, btn

    def add_action_button(self, parent, text="开始处理", command=None,
                          row=0, column=0, **grid_kw):
        """主要动作按钮，自动关联 btn_primary"""
        btn = ttk.Button(parent, text=text, command=command,
                         style="Default.TButton", width=16)
        btn.grid(row=row, column=column, **grid_kw)
        self.btn_primary = btn
        return btn

    # ── 日志方法 ────────────────────────────────────────────

    def log(self, message, tag=None):
        """在日志区追加一行（带时间戳）"""
        if self.log_text is None:
            return
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{ts}] {message}\n", tag)
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def clear_log(self):
        """清空日志"""
        if self.log_text:
            self.log_text.delete(1.0, tk.END)

    # ── 繁忙态管理 ──────────────────────────────────────────

    def set_busy(self, busy=True):
        """切换主按钮 / 进度条状态"""
        if self.btn_primary:
            self.btn_primary.config(state=tk.DISABLED if busy else tk.NORMAL,
                                    text="处理中..." if busy else None)
        if self.progress:
            if busy:
                self.progress.start()
            else:
                self.progress.stop()

    # ── 内部方法 ────────────────────────────────────────────

    def _browse_folder(self, var):
        folder = filedialog.askdirectory(title="选择文件夹")
        if folder:
            var.set(folder)

    def _browse_file(self, var, filetypes):
        path = filedialog.askopenfilename(title="选择文件", filetypes=filetypes)
        if path:
            var.set(path)

    # ── 子类入口 ────────────────────────────────────────────

    def build_ui(self):
        """子类 override 此方法搭建界面"""
        raise NotImplementedError
