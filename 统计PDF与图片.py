import os
import sys
import pandas as pd
import PyPDF2
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

# 支持的图片格式
IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.tif', 
                    '.webp', '.svg', '.ico', '.heic', '.heif', '.raw', '.cr2'}

class MediaStatisticsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Correct")
        self.root.geometry("750x600")
        
        # 设置窗口居中
        self.center_window()
        
        # 设置样式
        self.setup_styles()
        
        # 创建界面
        self.create_widgets()
        
    def center_window(self):
        """使窗口居中显示"""
        self.root.update_idletasks()
        width = 750
        height = 600
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
    def setup_styles(self):
        """设置ttk样式"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # 配置自定义样式
        style.configure('Title.TLabel', font=('楷体', 16, 'bold'))
        style.configure('Info.TLabel', font=('楷体', 9))
        style.configure('Success.TLabel', font=('楷体', 10, 'bold'), foreground='green')
        style.configure('Custom.TButton', font=('楷体', 10))
        style.configure('Section.TLabelframe', font=('楷体', 10, 'bold'))
        style.configure('Section.TLabelframe.Label', font=('楷体', 10, 'bold'))
        
        # 配置Checkbutton样式，使用✓而不是×
        style.configure('Switch.TCheckbutton', font=('楷体', 10))
        style.map('Switch.TCheckbutton',
                  indicatoron=[('selected', True)],
                  indicatormargin=[('selected', 4)])
        
    def create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="媒体文件统计", style='Title.TLabel')
        title_label.grid(row=0, column=0, pady=(0, 10))
        
        # 副标题
        """subtitle_label = ttk.Label(main_frame, text="分别统计图片和PDF文件夹，生成Excel报告", style='Info.TLabel')
        subtitle_label.grid(row=1, column=0, pady=(0, 20))"""
        
        # 使用说明
        """tip_label = ttk.Label(main_frame, text="💡 提示：勾选框显示 ✓ 表示启用统计，取消勾选则跳过该项", 
                             style='Info.TLabel', foreground='blue')
        tip_label.grid(row=2, column=0, pady=(0, 10))"""
        
        # 分隔线
        separator1 = ttk.Separator(main_frame, orient='horizontal')
        separator1.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=10)
        
        # ========== 图片文件夹选择框架 ==========
        image_frame = ttk.LabelFrame(main_frame, text="📁 图片文件夹设置", padding="10", style='Section.TLabelframe')
        image_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=10)
        image_frame.columnconfigure(0, weight=1)
        
        # 创建自定义Checkbutton（使用✓符号）
        self.include_images = tk.BooleanVar(value=True)
        image_check_frame = ttk.Frame(image_frame)
        image_check_frame.grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 5))
        
        # 使用Label模拟更清晰的复选框
        self.image_check_label = tk.Label(image_check_frame, text="☑", font=('Arial', 14), 
                                         fg='#366092', cursor='hand2')
        self.image_check_label.pack(side=tk.LEFT, padx=(0, 5))
        self.image_check_label.bind('<Button-1>', self.toggle_image_check)
        
        image_check_text = tk.Label(image_check_frame, text="启用图片统计（☑ 启用 / ☐ 禁用）", 
                                   font=('楷体', 10), cursor='hand2')
        image_check_text.pack(side=tk.LEFT)
        image_check_text.bind('<Button-1>', self.toggle_image_check)
        
        # 图片路径显示
        self.image_path_var = tk.StringVar()
        self.image_entry = ttk.Entry(image_frame, textvariable=self.image_path_var, font=('楷体', 10))
        self.image_entry.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        # 图片浏览按钮
        self.image_browse_btn = ttk.Button(image_frame, text="浏览", command=self.browse_image_folder, 
                                          style='Custom.TButton')
        self.image_browse_btn.grid(row=1, column=1)
        
        # 支持的图片格式标签
        image_formats = ttk.Label(image_frame, 
                                 text=f"支持格式: {', '.join(sorted(list(IMAGE_EXTENSIONS)[:8]))}等", 
                                 style='Info.TLabel', foreground='gray')
        image_formats.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=(5, 0))
        
        # ========== PDF文件夹选择框架 ==========
        pdf_frame = ttk.LabelFrame(main_frame, text="📄 PDF文件夹设置", padding="10", style='Section.TLabelframe')
        pdf_frame.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=10)
        pdf_frame.columnconfigure(0, weight=1)
        
        # 创建自定义Checkbutton（使用✓符号）
        self.include_pdfs = tk.BooleanVar(value=True)
        pdf_check_frame = ttk.Frame(pdf_frame)
        pdf_check_frame.grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 5))
        
        self.pdf_check_label = tk.Label(pdf_check_frame, text="☑", font=('Arial', 14), 
                                       fg='#366092', cursor='hand2')
        self.pdf_check_label.pack(side=tk.LEFT, padx=(0, 5))
        self.pdf_check_label.bind('<Button-1>', self.toggle_pdf_check)
        
        pdf_check_text = tk.Label(pdf_check_frame, text="启用PDF统计（☑ 启用 / ☐ 禁用）", 
                                 font=('楷体', 10), cursor='hand2')
        pdf_check_text.pack(side=tk.LEFT)
        pdf_check_text.bind('<Button-1>', self.toggle_pdf_check)
        
        # PDF路径显示
        self.pdf_path_var = tk.StringVar()
        self.pdf_entry = ttk.Entry(pdf_frame, textvariable=self.pdf_path_var, font=('楷体', 10))
        self.pdf_entry.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        # PDF浏览按钮
        self.pdf_browse_btn = ttk.Button(pdf_frame, text="浏览", command=self.browse_pdf_folder, 
                                        style='Custom.TButton')
        self.pdf_browse_btn.grid(row=1, column=1)
        
        # PDF信息标签
        pdf_info = ttk.Label(pdf_frame, text="统计所有子文件夹中的PDF文件页数和数量", 
                            style='Info.TLabel', foreground='gray')
        pdf_info.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=(5, 0))
        
        # ========== 输出设置框架 ==========
        output_frame = ttk.LabelFrame(main_frame, text="💾 输出设置", padding="10", style='Section.TLabelframe')
        output_frame.grid(row=6, column=0, sticky=(tk.W, tk.E), pady=10)
        output_frame.columnconfigure(1, weight=1)
        
        # 输出路径选择
        output_path_label = ttk.Label(output_frame, text="保存位置:")
        output_path_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        self.output_path_var = tk.StringVar(value=os.getcwd())
        output_entry = ttk.Entry(output_frame, textvariable=self.output_path_var, font=('楷体', 10))
        output_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        output_browse_btn = ttk.Button(output_frame, text="浏览", command=self.browse_output_folder, 
                                      style='Custom.TButton')
        output_browse_btn.grid(row=0, column=2)
        
        # ========== 进度条框架 ==========
        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=7, column=0, sticky=(tk.W, tk.E), pady=10)
        progress_frame.columnconfigure(0, weight=1)
        
        # 进度条
        self.progress = ttk.Progressbar(progress_frame, mode='indeterminate', length=400)
        self.progress.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # 状态标签
        self.status_var = tk.StringVar(value="就绪")
        status_label = ttk.Label(progress_frame, textvariable=self.status_var, style='Info.TLabel')
        status_label.grid(row=1, column=0, pady=(5, 0))
        
        # ========== 按钮框架 ==========
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=8, column=0, pady=15)
        
        # 开始统计按钮
        self.start_btn = ttk.Button(button_frame, text="开始统计", command=self.start_statistics, 
                                   style='Custom.TButton', width=15)
        self.start_btn.grid(row=0, column=0, padx=5)
        
        # 清空按钮
        clear_btn = ttk.Button(button_frame, text="清空路径", command=self.clear_paths, 
                              style='Custom.TButton', width=15)
        clear_btn.grid(row=0, column=1, padx=5)
        
        # 退出按钮
        exit_btn = ttk.Button(button_frame, text="退出", command=self.root.quit, 
                             style='Custom.TButton', width=15)
        exit_btn.grid(row=0, column=2, padx=5)
        
        # ========== 结果显示框架 ==========
        result_frame = ttk.LabelFrame(main_frame, text="统计结果", padding="10")
        result_frame.grid(row=9, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(9, weight=1)
        
        # 结果文本框
        self.result_text = tk.Text(result_frame, height=12, wrap=tk.WORD, font=('Consolas', 9))
        self.result_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 滚动条
        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.result_text['yscrollcommand'] = scrollbar.set
        
    def toggle_image_check(self, event=None):
        """切换图片统计复选框状态"""
        current_state = self.include_images.get()
        self.include_images.set(not current_state)
        
        # 更新显示符号
        if self.include_images.get():
            self.image_check_label.config(text="☑", fg='#366092')
            state = 'normal'
        else:
            self.image_check_label.config(text="☐", fg='gray')
            state = 'disabled'
            
        # 更新相关控件的状态
        self.image_entry.config(state=state)
        self.image_browse_btn.config(state=state)
        
    def toggle_pdf_check(self, event=None):
        """切换PDF统计复选框状态"""
        current_state = self.include_pdfs.get()
        self.include_pdfs.set(not current_state)
        
        # 更新显示符号
        if self.include_pdfs.get():
            self.pdf_check_label.config(text="☑", fg='#366092')
            state = 'normal'
        else:
            self.pdf_check_label.config(text="☐", fg='gray')
            state = 'disabled'
            
        # 更新相关控件的状态
        self.pdf_entry.config(state=state)
        self.pdf_browse_btn.config(state=state)
        
    def browse_image_folder(self):
        """浏览图片文件夹"""
        folder_path = filedialog.askdirectory(title="选择图片所在的文件夹")
        if folder_path:
            self.image_path_var.set(folder_path)
            
    def browse_pdf_folder(self):
        """浏览PDF文件夹"""
        folder_path = filedialog.askdirectory(title="选择PDF所在的文件夹")
        if folder_path:
            self.pdf_path_var.set(folder_path)
            
    def browse_output_folder(self):
        """浏览输出文件夹"""
        folder_path = filedialog.askdirectory(title="选择Excel文件保存位置")
        if folder_path:
            self.output_path_var.set(folder_path)
            
    def clear_paths(self):
        """清空所有路径"""
        self.image_path_var.set("")
        self.pdf_path_var.set("")
        self.output_path_var.set(os.getcwd())
        self.result_text.delete(1.0, tk.END)
        self.status_var.set("已清空路径")
        
    def start_statistics(self):
        """开始统计"""
        # 验证输入
        image_path = self.image_path_var.get().strip() if self.include_images.get() else None
        pdf_path = self.pdf_path_var.get().strip() if self.include_pdfs.get() else None
        
        # 检查是否至少选择了一项统计
        if not self.include_images.get() and not self.include_pdfs.get():
            messagebox.showwarning("警告", "请至少选择一项统计内容！")
            return
            
        # 验证已启用的路径
        if self.include_images.get() and not image_path:
            messagebox.showwarning("警告", "请选择图片文件夹！")
            return
            
        if self.include_pdfs.get() and not pdf_path:
            messagebox.showwarning("警告", "请选择PDF文件夹！")
            return
            
        # 验证路径是否存在
        if image_path and not os.path.exists(image_path):
            messagebox.showerror("错误", f"图片文件夹不存在：\n{image_path}")
            return
            
        if pdf_path and not os.path.exists(pdf_path):
            messagebox.showerror("错误", f"PDF文件夹不存在：\n{pdf_path}")
            return
            
        # 验证输出路径
        output_path = self.output_path_var.get().strip()
        if not os.path.exists(output_path):
            try:
                os.makedirs(output_path)
            except Exception as e:
                messagebox.showerror("错误", f"无法创建输出目录：\n{str(e)}")
                return
                
        # 在新线程中执行统计
        thread = threading.Thread(target=self.run_statistics, 
                                 args=(image_path, pdf_path, output_path))
        thread.daemon = True
        thread.start()
        
    def run_statistics(self, image_path, pdf_path, output_path):
        """执行统计任务"""
        try:
            # 更新UI状态
            self.root.after(0, self.update_ui_start)
            
            # 初始化结果
            image_data = []
            pdf_data = []
            image_stats = {'count': 0, 'folders': 0, 'total_size': 0}
            pdf_stats = {'count': 0, 'pages': 0, 'folders': 0, 'total_size': 0}
            
            # 统计图片
            if image_path:
                self.root.after(0, self.update_status, f"正在统计图片文件夹: {image_path}")
                image_data, image_stats = self.count_images(image_path)
                
            # 统计PDF
            if pdf_path:
                self.root.after(0, self.update_status, f"正在统计PDF文件夹: {pdf_path}")
                pdf_data, pdf_stats = self.count_pdfs(pdf_path)
                
            # 生成Excel文件
            self.root.after(0, self.update_status, "正在生成Excel报告...")
            excel_path = self.generate_excel(output_path, image_data, pdf_data, 
                                            image_stats, pdf_stats, image_path, pdf_path)
            
            # 显示结果
            self.root.after(0, self.display_results, excel_path, image_stats, pdf_stats, 
                          image_path, pdf_path)
            
        except Exception as e:
            self.root.after(0, self.show_error, str(e))
        finally:
            self.root.after(0, self.update_ui_finish)
            
    def count_images(self, root_path):
        """统计图片文件"""
        data = []
        folders_with_images = set()
        total_images = 0
        total_size = 0
        
        self.root.after(0, self.update_status, "正在扫描图片文件夹...")
        
        for foldername, subfolders, filenames in os.walk(root_path):
            image_files = [f for f in filenames 
                          if Path(f).suffix.lower() in IMAGE_EXTENSIONS]
            image_count = len(image_files)
            
            if image_count > 0:
                total_images += image_count
                rel_path = os.path.relpath(foldername, root_path)
                if rel_path == '.':
                    rel_path = '根目录'
                folders_with_images.add(rel_path)
                
                folder_size = sum(os.path.getsize(os.path.join(foldername, f)) 
                                for f in image_files) / (1024 * 1024)
                total_size += folder_size
                
                data.append({
                    '文件夹路径': rel_path,
                    '完整路径': foldername,
                    '图片数量': image_count,
                    '图片总大小(MB)': round(folder_size, 2)
                })
                
                self.root.after(0, self.update_status, f"发现图片: {rel_path} ({image_count} 张)")
                
        # 排序
        data.sort(key=lambda x: x['图片数量'], reverse=True)
        
        # 添加汇总行
        if data:
            data.append({
                '文件夹路径': '【总计】',
                '完整路径': f'共扫描 {len(folders_with_images)} 个包含图片的文件夹',
                '图片数量': total_images,
                '图片总大小(MB)': round(total_size, 2)
            })
            
        stats = {
            'count': total_images,
            'folders': len(folders_with_images),
            'total_size': round(total_size, 2)
        }
        
        return data, stats
        
    def count_pdfs(self, root_path):
        """统计PDF文件"""
        data = []
        folders_with_pdf = set()
        total_files = 0
        total_pages = 0
        total_size = 0
        
        self.root.after(0, self.update_status, "正在扫描PDF文件夹...")
        
        for foldername, subfolders, filenames in os.walk(root_path):
            pdf_files = [f for f in filenames if f.lower().endswith('.pdf')]
            
            if pdf_files:
                rel_path = os.path.relpath(foldername, root_path)
                if rel_path == '.':
                    rel_path = '根目录'
                folders_with_pdf.add(rel_path)
                
                for pdf_file in pdf_files:
                    try:
                        file_path = os.path.join(foldername, pdf_file)
                        file_size = os.path.getsize(file_path) / (1024 * 1024)
                        total_size += file_size
                        
                        with open(file_path, 'rb') as file:
                            pdf_reader = PyPDF2.PdfReader(file)
                            pages = len(pdf_reader.pages)
                            total_pages += pages
                            total_files += 1
                            
                            data.append({
                                '所在文件夹': rel_path,
                                '文件名': pdf_file,
                                '页数': pages,
                                '文件大小(MB)': round(file_size, 2),
                                '完整路径': file_path
                            })
                            
                            self.root.after(0, self.update_status, f"发现PDF: {rel_path}/{pdf_file} ({pages} 页)")
                    except Exception as e:
                        # 记录无法读取的PDF
                        data.append({
                            '所在文件夹': rel_path,
                            '文件名': pdf_file,
                            '页数': '读取失败',
                            '文件大小(MB)': round(file_size, 2),
                            '完整路径': file_path
                        })
                        
        # 排序
        data.sort(key=lambda x: (x['所在文件夹'], x['文件名']))
        
        # 添加汇总行
        if data:
            data.append({
                '所在文件夹': '【总计】',
                '文件名': f'共 {total_files} 个PDF文件',
                '页数': total_pages,
                '文件大小(MB)': round(total_size, 2),
                '完整路径': f'扫描范围: {root_path}'
            })
            
        stats = {
            'count': total_files,
            'pages': total_pages,
            'folders': len(folders_with_pdf),
            'total_size': round(total_size, 2)
        }
        
        return data, stats
        
    def generate_excel(self, output_path, image_data, pdf_data, image_stats, pdf_stats, 
                      image_path, pdf_path):
        """生成Excel文件"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f"媒体文件统计_{timestamp}.xlsx"
        excel_path = os.path.join(output_path, excel_filename)
        
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            # 写入图片统计工作表
            if image_data:
                df_images = pd.DataFrame(image_data)
                df_images.to_excel(writer, sheet_name='图片统计', index=False)
                self.format_excel_sheet(writer.sheets['图片统计'], 'image')
                
            # 写入PDF统计工作表
            if pdf_data:
                df_pdfs = pd.DataFrame(pdf_data)
                df_pdfs.to_excel(writer, sheet_name='PDF统计', index=False)
                self.format_excel_sheet(writer.sheets['PDF统计'], 'pdf')
                
            # 如果没有数据，创建一个空的工作表
            if not image_data and not pdf_data:
                empty_df = pd.DataFrame({'信息': ['未找到任何文件']})
                empty_df.to_excel(writer, sheet_name='统计结果', index=False)
                
        return excel_path
        
    def format_excel_sheet(self, worksheet, sheet_type):
        """格式化Excel工作表"""
        from openpyxl.styles import Font, PatternFill, Alignment
        
        # 设置列宽
        if sheet_type == 'image':
            worksheet.column_dimensions['A'].width = 40
            worksheet.column_dimensions['B'].width = 60
            worksheet.column_dimensions['C'].width = 15
            worksheet.column_dimensions['D'].width = 18
        else:  # pdf
            worksheet.column_dimensions['A'].width = 35
            worksheet.column_dimensions['B'].width = 45
            worksheet.column_dimensions['C'].width = 12
            worksheet.column_dimensions['D'].width = 18
            worksheet.column_dimensions['E'].width = 60
            
        # 设置表头样式
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_alignment = Alignment(horizontal="center", vertical="center")
        
        for cell in worksheet[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            
    def update_ui_start(self):
        """更新UI为开始状态"""
        self.start_btn.config(state='disabled')
        self.progress.start()
        self.result_text.delete(1.0, tk.END)
        
    def update_ui_finish(self):
        """更新UI为完成状态"""
        self.start_btn.config(state='normal')
        self.progress.stop()
        self.status_var.set("完成")
        
    def update_status(self, status):
        """更新状态标签"""
        self.status_var.set(status)
        self.result_text.insert(tk.END, f"{status}\n")
        self.result_text.see(tk.END)
        self.root.update_idletasks()
        
    def display_results(self, excel_path, image_stats, pdf_stats, image_path, pdf_path):
        """显示统计结果"""
        self.result_text.insert(tk.END, "\n" + "="*70 + "\n")
        self.result_text.insert(tk.END, "✅ 统计完成！\n\n")
        
        if image_path and image_stats['count'] > 0:
            self.result_text.insert(tk.END, f"📁 图片统计结果：\n")
            self.result_text.insert(tk.END, f"   源文件夹: {image_path}\n")
            self.result_text.insert(tk.END, f"   包含图片的文件夹: {image_stats['folders']} 个\n")
            self.result_text.insert(tk.END, f"   图片总数: {image_stats['count']:,} 张\n")
            self.result_text.insert(tk.END, f"   总大小: {image_stats['total_size']} MB\n\n")
        elif image_path:
            self.result_text.insert(tk.END, f"📁 图片统计: 未找到图片文件\n\n")
            
        if pdf_path and pdf_stats['count'] > 0:
            self.result_text.insert(tk.END, f"📄 PDF统计结果：\n")
            self.result_text.insert(tk.END, f"   源文件夹: {pdf_path}\n")
            self.result_text.insert(tk.END, f"   包含PDF的文件夹: {pdf_stats['folders']} 个\n")
            self.result_text.insert(tk.END, f"   PDF文件数: {pdf_stats['count']} 个\n")
            self.result_text.insert(tk.END, f"   总页数: {pdf_stats['pages']:,} 页\n")
            self.result_text.insert(tk.END, f"   总大小: {pdf_stats['total_size']} MB\n\n")
        elif pdf_path:
            self.result_text.insert(tk.END, f"📄 PDF统计: 未找到PDF文件\n\n")
            
        self.result_text.insert(tk.END, f"💾 报告已保存至:\n{excel_path}\n")
        self.result_text.insert(tk.END, "="*70 + "\n")
        self.result_text.see(tk.END)
        
        # 询问是否打开文件
        if messagebox.askyesno("完成", "统计完成！是否打开Excel文件？"):
            os.startfile(excel_path)
            
    def show_error(self, error_msg):
        """显示错误信息"""
        messagebox.showerror("错误", f"统计过程中出现错误：\n{error_msg}")
        self.result_text.insert(tk.END, f"\n❌ 错误：{error_msg}\n")
        
def main():
    """主函数"""
    root = tk.Tk()
    app = MediaStatisticsApp(root)
    root.mainloop()
    
if __name__ == "__main__":
    main()
