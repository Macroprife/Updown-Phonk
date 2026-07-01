"""
PDF 批量分割工具（传送门）
========================
按 Excel 配置将 PDF 按页范围拆分为多个子 PDF。
GUI 基于 BaseApp 基类，保留「传送门」视觉主题。
"""
import os
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import PyPDF2

from updown.gui import BaseApp


class PDFSplitterApp(BaseApp):
    REQUIRED_COLS = ["总文件名", "分件文件名", "起始页", "总页数", "每份文件页数"]

    def __init__(self, root):
        super().__init__(root, title="传送门 - PDF批量分割工具",
                         geometry="600x500")
        self.root.configure(bg="#F0F8FF")

        self.excel_path   = tk.StringVar()
        self.pdf_dir      = tk.StringVar()
        self.output_dir   = tk.StringVar()
        self.sheet_name   = tk.StringVar(value="抽象")

        self._btn_run = None   # tk.Button（保留红底风格）
        self.build_ui()

    # ── UI ──────────────────────────────────────────────────

    def build_ui(self):
        # 主框架（保留原蓝色调）
        mf = tk.Frame(self.root, bg="#E8F4FD", relief="ridge", bd=2)
        mf.place(relx=0.5, rely=0.5, anchor="center", width=520, height=420)

        def lbl(text, **kw):
            return tk.Label(mf, text=text, bg="#E8F4FD",
                            font=("正楷", 10), **kw)

        def browse_btn(cmd):
            return tk.Button(mf, text="浏览", command=cmd,
                             bg="#4A90D9", fg="white",
                             font=("正楷", 9), cursor="hand2", width=6)

        # 标题
        tk.Label(mf, text="传送门", font=("正楷", 28, "bold"),
                 bg="#E8F4FD", fg="#2C5F8A").pack(pady=(20, 5))
        tk.Label(mf, text="PDF批量分割工具", font=("正楷", 12),
                 bg="#E8F4FD", fg="#4A90D9").pack(pady=(0, 20))

        # 路径行
        def path_row(label, var, btn):
            f = tk.Frame(mf, bg="#E8F4FD")
            f.pack(pady=6, padx=40, fill="x")
            lbl(label, width=12, anchor="w").pack(side="left")
            tk.Entry(f, textvariable=var, font=("正楷", 10),
                     width=30).pack(side="left", padx=(0, 5))
            btn.pack(side="left")

        path_row("Excel文件:", self.excel_path,
                 browse_btn(lambda: self._browse("excel")))
        path_row("PDF文件夹:", self.pdf_dir,
                 browse_btn(lambda: self._browse("pdf")))
        path_row("输出文件夹:", self.output_dir,
                 browse_btn(lambda: self._browse("out")))

        # 工作表名
        sf = tk.Frame(mf, bg="#E8F4FD")
        sf.pack(pady=6, padx=40, fill="x")
        lbl("工作表名:", width=12, anchor="w").pack(side="left")
        tk.Entry(sf, textvariable=self.sheet_name,
                 font=("正楷", 10), width=30).pack(side="left", padx=(0, 5))

        # 进度条
        self.progress = ttk.Progressbar(mf, length=400, mode="determinate")
        self.progress.pack(pady=10)

        # 状态标签
        self._status = tk.Label(mf, text="就绪", bg="#E8F4FD",
                                fg="#666", font=("正楷", 9))
        self._status.pack(pady=2)

        # 主按钮（保留红底传送风格）
        self._btn_run = tk.Button(mf, text="开启传送",
                                  command=self.run_split,
                                  bg="#FF6B6B", fg="white",
                                  font=("正楷", 14, "bold"),
                                  width=12, height=2, cursor="hand2")
        self._btn_run.pack(pady=12)

    def _browse(self, kind):
        if kind == "excel":
            p = tk.filedialog.askopenfilename(
                title="选择Excel文件",
                filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")])
            if p:
                self.excel_path.set(p)
        else:
            d = tk.filedialog.askdirectory(
                title="选择PDF文件夹" if kind == "pdf" else "选择输出文件夹")
            if d:
                (self.pdf_dir if kind == "pdf" else self.output_dir).set(d)

    def _status_set(self, msg, color="#666"):
        self._status.config(text=msg, fg=color)
        self.root.update()

    # ── 业务 ────────────────────────────────────────────────

    def run_split(self):
        excel = self.excel_path.get()
        pdfd  = self.pdf_dir.get()
        outd  = self.output_dir.get()

        for v, n in [(excel, "Excel"), (pdfd, "PDF文件夹"), (outd, "输出文件夹")]:
            if not v:
                return messagebox.showerror("错误", f"请选择{n}！")

        self._btn_run.config(state="disabled", bg="#ccc", text="传送中...")
        self._status_set("传送门开启中...", "#4A90D9")

        try:
            df = pd.read_excel(excel, sheet_name=self.sheet_name.get().strip() or "抽象")
            missing = [c for c in self.REQUIRED_COLS if c not in df.columns]
            if missing:
                raise Exception(f"Excel缺少必需列: {missing}")

            groups = df.groupby("总文件名")
            total = len(groups)
            self.progress["maximum"] = total
            ok = errs = 0
            errors = []

            for idx, (name, gdf) in enumerate(groups):
                self.progress["value"] = idx + 1
                self._status_set(f"传送中: {name} ({idx+1}/{total})")
                try:
                    self._split_one(name, gdf, pdfd, outd)
                    ok += 1
                except Exception as e:
                    errors.append(f"{name}: {e}")
                    errs += 1

            msg = f"传送完成！\n成功: {ok}  失败: {errs}"
            if errors:
                msg += "\n\n错误:\n" + "\n".join(errors[:20])
                if len(errors) > 20:
                    msg += f"\n…还有 {len(errors)-20} 个"
            messagebox.showinfo("传送结果", msg)
            self._status_set("传送完成", "#28A745")

        except Exception as e:
            messagebox.showerror("传送失败", str(e))
            self._status_set("传送失败", "#DC3545")
        finally:
            self._btn_run.config(state="normal", bg="#FF6B6B", text="开启传送")
            self.progress["value"] = 0

    def _split_one(self, total_name, gdf, pdf_dir, out_dir):
        # 找 PDF
        pdf_path = None
        for f in os.listdir(pdf_dir):
            if f.lower().endswith(".pdf") and total_name in f:
                pdf_path = os.path.join(pdf_dir, f)
                break
        if not pdf_path:
            raise Exception(f"找不到 {total_name} 对应的PDF")

        sub = os.path.join(out_dir, total_name)
        os.makedirs(sub, exist_ok=True)

        with open(pdf_path, "rb") as fh:
            reader = PyPDF2.PdfReader(fh)
            n = len(reader.pages)

            expected = int(gdf.iloc[0]["总页数"])
            if n != expected:
                self._status_set(f"⚠ {total_name} 页数不符({n}≠{expected})", "#FF8C00")

            gdf = gdf.sort_values("起始页")
            for i, (_, row) in enumerate(gdf.iterrows()):
                start = int(row["起始页"]) - 1
                part = str(row["分件文件名"]).strip()
                exp  = int(row["每份文件页数"])

                end = n  # 最后一份
                if i < len(gdf) - 1:
                    end = int(gdf.iloc[i + 1]["起始页"]) - 1

                writer = PyPDF2.PdfWriter()
                for p in range(start, end):
                    writer.add_page(reader.pages[p])

                actual = end - start
                if actual != exp:
                    self._status_set(f"⚠ {part} 页数({actual}≠{exp})", "#FF8C00")

                out = os.path.join(sub, f"{part}.pdf")
                with open(out, "wb") as of:
                    writer.write(of)


def main():
    root = tk.Tk()
    PDFSplitterApp(root)
    root.mainloop()


if __name__ == "__main__":
    try:
        import pandas
        import PyPDF2
    except ImportError:
        print("请安装依赖: pip install pandas PyPDF2 openpyxl")
        exit(1)
    main()
