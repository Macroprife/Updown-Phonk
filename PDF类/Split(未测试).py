import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import PyPDF2
import os
from pathlib import Path

class PDFSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("传送门 - PDF批量分割工具")
        self.root.geometry("600x500")
        self.root.resizable(False, False)
        
        # 设置窗口背景
        self.root.configure(bg='#F0F8FF')
        
        # 创建主框架
        self.main_frame = tk.Frame(self.root, bg='#E8F4FD', relief='ridge', bd=2)
        self.main_frame.place(relx=0.5, rely=0.5, anchor="center", width=520, height=420)
        
        # 设置样式
        self.setup_styles()
        
        # 创建界面元素
        self.create_widgets()
        
    def setup_styles(self):
        """设置样式"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # 配置正楷字体
        style.configure('TButton', font=('正楷', 10), padding=5)
        style.configure('TLabel', font=('正楷', 10), background='#E8F4FD')
        style.configure('TEntry', font=('正楷', 10), padding=3)
        
    def create_widgets(self):
        """创建界面控件"""
        # 标题
        title_label = tk.Label(self.main_frame, text="传送门", 
                               font=("正楷", 28, "bold"), 
                               bg='#E8F4FD', fg='#2C5F8A')
        title_label.pack(pady=(20, 5))
        
        subtitle_label = tk.Label(self.main_frame, text="PDF批量分割工具", 
                                  font=("正楷", 12), 
                                  bg='#E8F4FD', fg='#4A90D9')
        subtitle_label.pack(pady=(0, 20))
        
        # Excel文件选择
        excel_frame = tk.Frame(self.main_frame, bg='#E8F4FD')
        excel_frame.pack(pady=8, padx=40, fill='x')
        
        tk.Label(excel_frame, text="Excel文件:", bg='#E8F4FD', 
                font=('正楷', 10), width=12, anchor='w').pack(side='left')
        self.excel_path_var = tk.StringVar()
        self.excel_entry = tk.Entry(excel_frame, textvariable=self.excel_path_var, 
                                    font=('正楷', 10), width=30)
        self.excel_entry.pack(side='left', padx=(0, 5))
        tk.Button(excel_frame, text="浏览", command=self.browse_excel, 
                 bg='#4A90D9', fg='white', font=('正楷', 9),
                 cursor='hand2', width=6).pack(side='left')
        
        # PDF文件夹选择
        pdf_frame = tk.Frame(self.main_frame, bg='#E8F4FD')
        pdf_frame.pack(pady=8, padx=40, fill='x')
        
        tk.Label(pdf_frame, text="PDF文件夹:", bg='#E8F4FD', 
                font=('正楷', 10), width=12, anchor='w').pack(side='left')
        self.pdf_dir_var = tk.StringVar()
        self.pdf_entry = tk.Entry(pdf_frame, textvariable=self.pdf_dir_var, 
                                  font=('正楷', 10), width=30)
        self.pdf_entry.pack(side='left', padx=(0, 5))
        tk.Button(pdf_frame, text="浏览", command=self.browse_pdf_dir, 
                 bg='#4A90D9', fg='white', font=('正楷', 9),
                 cursor='hand2', width=6).pack(side='left')
        
        # 输出文件夹选择
        output_frame = tk.Frame(self.main_frame, bg='#E8F4FD')
        output_frame.pack(pady=8, padx=40, fill='x')
        
        tk.Label(output_frame, text="输出文件夹:", bg='#E8F4FD', 
                font=('正楷', 10), width=12, anchor='w').pack(side='left')
        self.output_dir_var = tk.StringVar()
        self.output_entry = tk.Entry(output_frame, textvariable=self.output_dir_var, 
                                     font=('正楷', 10), width=30)
        self.output_entry.pack(side='left', padx=(0, 5))
        tk.Button(output_frame, text="浏览", command=self.browse_output_dir, 
                 bg='#4A90D9', fg='white', font=('正楷', 9),
                 cursor='hand2', width=6).pack(side='left')
        
        # 工作表名称
        sheet_frame = tk.Frame(self.main_frame, bg='#E8F4FD')
        sheet_frame.pack(pady=8, padx=40, fill='x')
        
        tk.Label(sheet_frame, text="工作表名:", bg='#E8F4FD', 
                font=('正楷', 10), width=12, anchor='w').pack(side='left')
        self.sheet_name_var = tk.StringVar(value="抽象")
        self.sheet_entry = tk.Entry(sheet_frame, textvariable=self.sheet_name_var, 
                                    font=('正楷', 10), width=30)
        self.sheet_entry.pack(side='left', padx=(0, 5))
        
        # 功能提示（简洁版）
        #tip_frame = tk.Frame(self.main_frame, bg='#FFF8DC', relief='groove', bd=1)
        #tip_frame.pack(pady=15, padx=40, fill='x')
        
        #tip_text = "📌 Excel列：总文件名、分件文件名、起始页、总页数、每份文件页数"
        #tk.Label(tip_frame, text=tip_text, bg='#FFF8DC', fg='#8B6914', 
        #       font=('正楷', 9), padx=10, pady=6).pack()
        
        # 进度条
        self.progress = ttk.Progressbar(self.main_frame, length=400, mode='determinate')
        self.progress.pack(pady=10)
        
        # 状态标签
        self.status_label = tk.Label(self.main_frame, text="就绪", bg='#E8F4FD', 
                                     fg='#666', font=('正楷', 9))
        self.status_label.pack(pady=5)
        
        # 执行按钮
        self.run_button = tk.Button(self.main_frame, text="开启传送", command=self.run_split,
                                    bg='#FF6B6B', fg='white', font=('正楷', 14, 'bold'),
                                    width=12, height=2, cursor='hand2')
        self.run_button.pack(pady=15)
        
    def browse_excel(self):
        """浏览Excel文件"""
        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if filename:
            self.excel_path_var.set(filename)
            
    def browse_pdf_dir(self):
        """浏览PDF文件夹"""
        dirname = filedialog.askdirectory(title="选择PDF文件夹")
        if dirname:
            self.pdf_dir_var.set(dirname)
            
    def browse_output_dir(self):
        """浏览输出文件夹"""
        dirname = filedialog.askdirectory(title="选择输出文件夹")
        if dirname:
            self.output_dir_var.set(dirname)
            
    def update_status(self, message, color='#666'):
        """更新状态信息"""
        self.status_label.config(text=message, fg=color)
        self.root.update()
        
    def run_split(self):
        """执行PDF分割"""
        # 验证输入
        if not self.excel_path_var.get():
            messagebox.showerror("错误", "请选择Excel文件！")
            return
        if not self.pdf_dir_var.get():
            messagebox.showerror("错误", "请选择PDF文件夹！")
            return
        if not self.output_dir_var.get():
            messagebox.showerror("错误", "请选择输出文件夹！")
            return
            
        self.run_button.config(state='disabled', bg='#ccc', text="传送中...")
        self.update_status("传送门开启中...", '#4A90D9')
        
        try:
            # 读取Excel数据
            df = pd.read_excel(self.excel_path_var.get(), sheet_name=self.sheet_name_var.get())
            
            # 验证必需列
            required_cols = ['总文件名', '分件文件名', '起始页', '总页数', '每份文件页数']
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                raise Exception(f"Excel缺少必需列：{missing_cols}")
            
            # 按总文件名分组
            groups = df.groupby('总文件名')
            total_groups = len(groups)
            
            self.progress['maximum'] = total_groups
            success_count = 0
            error_files = []
            
            for idx, (total_name, group_df) in enumerate(groups):
                self.progress['value'] = idx + 1
                self.update_status(f"传送中：{total_name} ({idx+1}/{total_groups})")
                
                try:
                    # 查找对应的PDF文件
                    pdf_path = None
                    for ext in ['.pdf', '.PDF']:
                        test_path = os.path.join(self.pdf_dir_var.get(), total_name + ext)
                        if os.path.exists(test_path):
                            pdf_path = test_path
                            break
                        # 尝试模糊匹配
                        for file in os.listdir(self.pdf_dir_var.get()):
                            if file.lower().endswith('.pdf') and total_name in file:
                                pdf_path = os.path.join(self.pdf_dir_var.get(), file)
                                break
                        if pdf_path:
                            break
                    
                    if not pdf_path:
                        raise Exception(f"找不到PDF文件：{total_name}")
                    
                    # 创建一级目录（总文件名）
                    output_subdir = os.path.join(self.output_dir_var.get(), total_name)
                    os.makedirs(output_subdir, exist_ok=True)
                    
                    # 读取PDF
                    with open(pdf_path, 'rb') as pdf_file:
                        pdf_reader = PyPDF2.PdfReader(pdf_file)
                        total_pages = len(pdf_reader.pages)
                        
                        # 验证总页数
                        expected_total = group_df.iloc[0]['总页数']
                        if total_pages != expected_total:
                            self.update_status(f"警告：{total_name} 实际页数({total_pages})与记录({expected_total})不符", '#FF8C00')
                        
                        # 按起始页排序
                        group_df = group_df.sort_values('起始页')
                        
                        # 分割PDF
                        for i, row in group_df.iterrows():
                            start_page = row['起始页'] - 1  # 转换为0索引
                            part_name = row['分件文件名']
                            expected_pages = row['每份文件页数']
                            
                            # 确定结束页
                            current_idx = group_df[group_df['分件文件名'] == part_name].index[0]
                            next_rows = group_df[group_df.index > current_idx]
                            
                            if len(next_rows) > 0:
                                end_page = next_rows.iloc[0]['起始页'] - 1
                            else:
                                end_page = total_pages
                            
                            # 创建新PDF
                            pdf_writer = PyPDF2.PdfWriter()
                            actual_pages = 0
                            for page_num in range(start_page, end_page):
                                pdf_writer.add_page(pdf_reader.pages[page_num])
                                actual_pages += 1
                            
                            # 验证页数
                            if actual_pages != expected_pages:
                                self.update_status(f"警告：{part_name} 实际页数({actual_pages})与预期({expected_pages})不符", '#FF8C00')
                            
                            # 保存PDF：直接以分件文件名命名，放在总文件名文件夹下
                            output_path = os.path.join(output_subdir, f"{part_name}.pdf")
                            with open(output_path, 'wb') as output_file:
                                pdf_writer.write(output_file)
                    
                    success_count += 1
                    
                except Exception as e:
                    error_files.append(f"{total_name}: {str(e)}")
                    
            # 显示结果
            self.progress['value'] = 0
            result_msg = f"传送完成！\n成功：{success_count} 个文件\n失败：{len(error_files)} 个文件"
            
            if error_files:
                result_msg += "\n\n错误详情：\n" + "\n".join(error_files[:5])
                if len(error_files) > 5:
                    result_msg += f"\n... 还有 {len(error_files)-5} 个错误"
            
            # 显示目录结构示例
            if success_count > 0:
                result_msg += "\n\n📁 输出目录结构示例：\n"
                result_msg += "输出文件夹/\n"
                result_msg += "└── 230-QQ0310-0001 李得林/\n"
                result_msg += "    ├── 230-QQ0310-0001-001.pdf\n"
                result_msg += "    ├── 230-QQ0310-0001-002.pdf\n"
                result_msg += "    └── ..."
                    
            messagebox.showinfo("传送结果", result_msg)
            self.update_status("传送完成", '#28A745')
            
        except Exception as e:
            messagebox.showerror("传送失败", f"处理失败：{str(e)}")
            self.update_status("传送失败", '#DC3545')
            
        finally:
            self.run_button.config(state='normal', bg='#FF6B6B', text="开启传送")
            self.progress['value'] = 0

def main():
    root = tk.Tk()
    app = PDFSplitterApp(root)
    root.mainloop()

if __name__ == "__main__":
    # 安装依赖提示
    try:
        import pandas
        import PyPDF2
    except ImportError:
        print("请先安装依赖：pip install pandas PyPDF2 openpyxl")
        exit(1)
    main()
