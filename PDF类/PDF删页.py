"""
PDF 批量删除页工具
==================
支持：仅删除首页 / 删除首尾页
GUI 基于 BaseApp 基类
"""
import os
import threading
import tkinter as tk
from tkinter import messagebox, ttk
from PyPDF2 import PdfReader, PdfWriter

from updown.gui import BaseApp


class PDFProcessorApp(BaseApp):
    MODE_DESC = {
        "remove_first":      "仅删除第一页（保留第2页到最后一页）",
        "remove_first_last": "删除第一页和最后一页（保留中间所有页）",
    }

    def __init__(self, root):
        super().__init__(root, title="📚 PDF 批量删除页面工具", geometry="650x500")
        self.source_folder = tk.StringVar()
        self.target_folder = tk.StringVar()
        self.processing_mode = tk.StringVar(value="remove_first")
        self.build_ui()

    # ── 界面搭建 ────────────────────────────────────────────

    def build_ui(self):
        mf = ttk.Frame(self.root, padding="20")
        mf.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        mf.columnconfigure(1, weight=1)

        # 标题
        self.add_title(mf, "📚 PDF 批量删除页面工具",
                       row=0, column=0, columnspan=3, pady=(0, 20))

        # 源文件夹
        self.add_browse_folder_row(mf, "源文件夹：", self.source_folder,
                                   row=1, column=0, columnspan=3)

        # 目标文件夹
        self.add_browse_folder_row(mf, "目标文件夹：", self.target_folder,
                                   row=2, column=0, columnspan=3)

        # ── 处理模式选择 ──
        ttk.Label(mf, text="处理模式：", style="Default.TLabel",
                  width=12, anchor=tk.W).grid(row=3, column=0, sticky=tk.W, pady=(20, 5))

        mode_frame = ttk.LabelFrame(mf, text="处理模式", padding="10")
        mode_frame.grid(row=3, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(15, 5))

        ttk.Radiobutton(mode_frame,
                        text="🗑️ 仅删除第一页（保留第2页到最后一页）",
                        variable=self.processing_mode,
                        value="remove_first").grid(row=0, column=0, sticky=tk.W, pady=2)

        ttk.Radiobutton(mode_frame,
                        text="🗑️🗑️ 删除第一页和最后一页（保留中间所有页）",
                        variable=self.processing_mode,
                        value="remove_first_last").grid(row=1, column=0, sticky=tk.W, pady=2)

        ttk.Label(mode_frame,
                  text="💡 页数 ≤2 的 PDF 将被跳过",
                  foreground="gray").grid(row=2, column=0, sticky=tk.W, pady=(10, 0))

        # 开始按钮
        self.add_action_button(mf, text="🚀 开始处理", command=self.start_processing,
                               row=4, column=0, columnspan=3, pady=(20, 10))

        # 进度条
        self.add_progress_bar(mf, row=5, column=0, columnspan=3, pady=(5, 10))

        # 日志
        self.add_log_area(mf, row=6, column=0, columnspan=3, height=12,
                          pady=(10, 0))
        mf.rowconfigure(6, weight=1)

    # ── 业务逻辑 ────────────────────────────────────────────

    def start_processing(self):
        source = self.source_folder.get()
        target = self.target_folder.get()
        if not source:
            return messagebox.showerror("错误", "请选择源文件夹！")
        if not os.path.exists(source):
            return messagebox.showerror("错误", "源文件夹不存在！")
        if not target:
            return messagebox.showerror("错误", "请选择目标文件夹！")

        self.clear_log()
        self.set_busy(True)
        threading.Thread(target=self._run, args=(source, target), daemon=True).start()

    def _run(self, source, target):
        try:
            self._process_pdfs(source, target)
        except Exception as e:
            self.log(f"❌ 处理异常: {e}", "error")
        finally:
            self.root.after(0, lambda: self.set_busy(False))
            self.log("✨ 处理完成！", "success")

    def _process_pdfs(self, source, target):
        mode = self.processing_mode.get()
        os.makedirs(target, exist_ok=True)

        is_first_last = (mode == "remove_first_last")
        mode_name = "删除首尾页" if is_first_last else "删除第一页"
        min_pages = 2 if is_first_last else 1

        self.log(f"📁 源文件夹: {source}", "info")
        self.log(f"📁 目标文件夹: {target}", "info")
        self.log(f"🔧 模式: {mode_name}", "info")
        self.log("-" * 50, "info")

        success = skip = errors = 0
        total = 0

        for fname in os.listdir(source):
            if not fname.lower().endswith(".pdf"):
                continue
            total += 1
            in_path = os.path.join(source, fname)
            out_path = os.path.join(target, fname)

            self.log(f"\n📄 {fname}", "info")
            ok = self._remove_pages(in_path, out_path, is_first_last, min_pages)
            if ok:
                success += 1
            else:
                try:
                    r = PdfReader(in_path)
                    if len(r.pages) <= min_pages:
                        skip += 1
                    else:
                        errors += 1
                except Exception:
                    errors += 1

        self.log("\n" + "=" * 50, "info")
        self.log(f"📊 总 PDF: {total}  |  ✅ {success}  |  ⚠️跳过 {skip}  |  ❌失败 {errors}")

    def _remove_pages(self, in_path, out_path, first_last, min_pages):
        try:
            reader = PdfReader(in_path)
            n = len(reader.pages)
            if n <= min_pages:
                self.log(f"⚠️ 跳过：{os.path.basename(in_path)} (≤{min_pages}页)", "warning")
                return False
            writer = PdfWriter()
            end = (n - 1) if first_last else n
            for i in range(1, end):
                writer.add_page(reader.pages[i])
            with open(out_path, "wb") as f:
                writer.write(f)
            kept = n - (2 if first_last else 1)
            self.log(f"✅ {os.path.basename(in_path)} ({n}→{kept}页)", "success")
            return True
        except Exception as e:
            self.log(f"❌ {os.path.basename(in_path)}: {e}", "error")
            return False


def main():
    root = tk.Tk()
    PDFProcessorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
