import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PyPDF2 import PdfReader, PdfWriter


class PDFProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 批量处理工具")
        self.root.geometry("650x500")
        self.root.resizable(False, False)
        
        # 变量定义 - 必须放在 create_widgets() 之前
        self.source_folder = tk.StringVar()
        self.target_folder = tk.StringVar()
        self.processing_mode = tk.StringVar(value="remove_first")  # 默认删除第一页
        
        # 设置样式
        self.setup_styles()
        
        # 创建界面
        self.create_widgets()
        
    def setup_styles(self):
        """设置界面样式"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # 配置各种样式
        style.configure('Title.TLabel', font=('楷体', 16, 'bold'))
        style.configure('Heading.TLabel', font=('楷体', 11, 'bold'))
        style.configure('Info.TLabel', font=('楷体', 9))
        style.configure('Success.TLabel', font=('楷体', 9), foreground='green')
        style.configure('Error.TLabel', font=('楷体', 9), foreground='red')
        style.configure('Custom.TButton', font=('楷体', 10))
        
    def create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 标题
        title_label = ttk.Label(main_frame, text="📚 PDF 批量删除页面工具", style='Title.TLabel')
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # ==================== 源文件夹选择 ====================
        row_idx = 1
        ttk.Label(main_frame, text="1. 选择源文件夹：", style='Heading.TLabel').grid(
            row=row_idx, column=0, sticky=tk.W, pady=(5, 5))
        
        source_frame = ttk.Frame(main_frame)
        source_frame.grid(row=row_idx, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 5))
        
        self.source_entry = ttk.Entry(source_frame, textvariable=self.source_folder, width=50)
        self.source_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(source_frame, text="浏览...", command=self.browse_source, 
                  style='Custom.TButton').pack(side=tk.LEFT)
        
        # ==================== 目标文件夹选择 ====================
        row_idx += 1
        ttk.Label(main_frame, text="2. 选择目标文件夹：", style='Heading.TLabel').grid(
            row=row_idx, column=0, sticky=tk.W, pady=(5, 5))
        
        target_frame = ttk.Frame(main_frame)
        target_frame.grid(row=row_idx, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 5))
        
        self.target_entry = ttk.Entry(target_frame, textvariable=self.target_folder, width=50)
        self.target_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(target_frame, text="浏览...", command=self.browse_target, 
                  style='Custom.TButton').pack(side=tk.LEFT)
        
        # ==================== 处理模式选择 ====================
        row_idx += 1
        ttk.Label(main_frame, text="3. 选择处理模式：", style='Heading.TLabel').grid(
            row=row_idx, column=0, sticky=tk.W, pady=(20, 5))
        
        mode_frame = ttk.LabelFrame(main_frame, text="处理模式", padding="10")
        mode_frame.grid(row=row_idx, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(15, 5))
        
        # 单选按钮
        ttk.Radiobutton(mode_frame, text="🗑️ 仅删除第一页（保留第2页到最后一页）", 
                       variable=self.processing_mode, value="remove_first").grid(
                           row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        ttk.Radiobutton(mode_frame, text="🗑️🗑️ 删除第一页和最后一页（保留中间所有页）", 
                       variable=self.processing_mode, value="remove_first_last").grid(
                           row=1, column=0, sticky=tk.W)
        
        # 模式说明标签
        mode_info = ttk.Label(mode_frame, 
                            text="💡 提示：删除首尾页时，页数≤2的PDF将被跳过",
                            style='Info.TLabel',
                            foreground='gray')
        mode_info.grid(row=2, column=0, sticky=tk.W, pady=(10, 0))
        
        # ==================== 执行按钮 ====================
        row_idx += 1
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=row_idx, column=0, columnspan=3, pady=(20, 10))
        
        self.process_btn = ttk.Button(button_frame, text="🚀 开始处理", 
                                      command=self.start_processing, 
                                      style='Custom.TButton',
                                      width=20)
        self.process_btn.pack()
        
        # ==================== 进度条 ====================
        row_idx += 1
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=row_idx, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 10))
        
        # ==================== 日志输出区域 ====================
        row_idx += 1
        ttk.Label(main_frame, text="处理日志：", style='Heading.TLabel').grid(
            row=row_idx, column=0, sticky=tk.W, pady=(10, 5))
        
        row_idx += 1
        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=row_idx, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 滚动条
        scrollbar = ttk.Scrollbar(log_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 文本框
        self.log_text = tk.Text(log_frame, height=12, width=75, 
                               wrap=tk.WORD, font=('Consolas', 9))
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar.config(command=self.log_text.yview)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # 配置颜色标签
        self.log_text.tag_config('success', foreground='green')
        self.log_text.tag_config('error', foreground='red')
        self.log_text.tag_config('warning', foreground='orange')
        self.log_text.tag_config('info', foreground='blue')
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(row_idx, weight=1)
        
    def browse_source(self):
        """浏览源文件夹"""
        folder = filedialog.askdirectory(title="选择包含PDF文件的源文件夹")
        if folder:
            self.source_folder.set(folder)
            
    def browse_target(self):
        """浏览目标文件夹"""
        folder = filedialog.askdirectory(title="选择保存处理后PDF的目标文件夹")
        if folder:
            self.target_folder.set(folder)
            
    def log_message(self, message, tag=None):
        """在日志区域显示消息"""
        self.log_text.insert(tk.END, message + "\n", tag)
        self.log_text.see(tk.END)
        self.root.update()
        
    def clear_log(self):
        """清空日志"""
        self.log_text.delete(1.0, tk.END)
        
    def remove_first_page(self, input_pdf, output_pdf):
        """删除PDF的第一页"""
        try:
            reader = PdfReader(input_pdf)
            writer = PdfWriter()
            total_pages = len(reader.pages)
            
            if total_pages <= 1:
                self.log_message(f"⚠️ 跳过：{os.path.basename(input_pdf)} (只有{total_pages}页，删除后无内容)", 'warning')
                return False
            
            for page_num in range(1, total_pages):
                writer.add_page(reader.pages[page_num])
                
            with open(output_pdf, "wb") as out_file:
                writer.write(out_file)
                
            self.log_message(f"✅ 成功：{os.path.basename(input_pdf)} (原{total_pages}页 → 新{total_pages-1}页)", 'success')
            return True
            
        except Exception as e:
            self.log_message(f"❌ 错误：{os.path.basename(input_pdf)} - {str(e)}", 'error')
            return False
            
    def remove_first_last_page(self, input_pdf, output_pdf):
        """删除PDF的第一页和最后一页"""
        try:
            reader = PdfReader(input_pdf)
            writer = PdfWriter()
            total_pages = len(reader.pages)
            
            if total_pages <= 2:
                self.log_message(f"⚠️ 跳过：{os.path.basename(input_pdf)} (页数 {total_pages} ≤ 2，删除首尾后无内容)", 'warning')
                return False
            
            for page_num in range(1, total_pages - 1):
                writer.add_page(reader.pages[page_num])
                
            with open(output_pdf, "wb") as out_file:
                writer.write(out_file)
                
            self.log_message(f"✅ 成功：{os.path.basename(input_pdf)} (原{total_pages}页 → 新{total_pages-2}页)", 'success')
            return True
            
        except Exception as e:
            self.log_message(f"❌ 错误：{os.path.basename(input_pdf)} - {str(e)}", 'error')
            return False
            
    def process_pdfs(self):
        """在后台线程中处理PDF文件"""
        source = self.source_folder.get()
        target = self.target_folder.get()
        mode = self.processing_mode.get()
        
        # 创建目标文件夹
        os.makedirs(target, exist_ok=True)
        
        # 统计信息
        total_files = 0
        success_count = 0
        skip_count = 0
        error_count = 0
        
        # 根据模式选择处理函数
        if mode == "remove_first":
            process_func = self.remove_first_page
            min_pages = 1
            mode_name = "删除第一页"
        else:
            process_func = self.remove_first_last_page
            min_pages = 2
            mode_name = "删除首尾页"
            
        self.log_message(f"\n📁 源文件夹: {source}", 'info')
        self.log_message(f"📁 目标文件夹: {target}", 'info')
        self.log_message(f"🔧 处理模式: {mode_name}", 'info')
        self.log_message("-" * 50, 'info')
        
        # 遍历所有PDF文件
        for filename in os.listdir(source):
            if not filename.lower().endswith(".pdf"):
                continue
                
            total_files += 1
            input_path = os.path.join(source, filename)
            output_path = os.path.join(target, filename)
            
            self.log_message(f"\n📄 处理: {filename}", 'info')
            success = process_func(input_path, output_path)
            
            if success:
                success_count += 1
            elif success is False:
                # 检查是否是因为页数不足而跳过
                try:
                    reader = PdfReader(input_path)
                    if len(reader.pages) <= min_pages:
                        skip_count += 1
                    else:
                        error_count += 1
                except:
                    error_count += 1
                    
        # 输出统计信息
        self.log_message("\n" + "=" * 50, 'info')
        self.log_message("📊 处理完成统计：", 'info')
        self.log_message(f"   总PDF文件数: {total_files}", 'info')
        self.log_message(f"   ✅ 成功处理: {success_count}", 'success')
        self.log_message(f"   ⚠️  跳过（页数≤{min_pages}）: {skip_count}", 'warning')
        self.log_message(f"   ❌ 处理失败: {error_count}", 'error')
        self.log_message(f"📁 输出位置: {target}", 'info')
        self.log_message("=" * 50, 'info')
        
    def processing_thread(self):
        """处理线程"""
        try:
            self.process_pdfs()
        except Exception as e:
            self.log_message(f"\n❌ 处理过程中发生错误: {str(e)}", 'error')
        finally:
            # 恢复界面
            self.progress.stop()
            self.process_btn.config(state='normal')
            self.log_message("\n✨ 处理完成！", 'success')
            
    def start_processing(self):
        """开始处理"""
        # 验证输入
        source = self.source_folder.get()
        target = self.target_folder.get()
        
        if not source:
            messagebox.showerror("错误", "请选择源文件夹！")
            return
            
        if not os.path.exists(source):
            messagebox.showerror("错误", "源文件夹不存在！")
            return
            
        if not target:
            messagebox.showerror("错误", "请选择目标文件夹！")
            return
            
        # 清空日志
        self.clear_log()
        
        # 禁用按钮，显示进度条
        self.process_btn.config(state='disabled')
        self.progress.start()
        
        # 在新线程中处理
        thread = threading.Thread(target=self.processing_thread, daemon=True)
        thread.start()


def main():
    root = tk.Tk()
    app = PDFProcessorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
