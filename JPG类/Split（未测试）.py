import os
import shutil
from pathlib import Path
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

class ImageSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("图片批量分割工具 v2.0")
        self.root.geometry("700x500")
        self.root.resizable(True, True)
        
        # 设置样式
        self.setup_styles()
        
        # 创建界面组件
        self.create_widgets()
        
        # 居中显示窗口
        self.center_window()
        
    def setup_styles(self):
        """设置界面样式"""
        self.font_default = ("楷体", 10)
        self.font_title = ("楷体", 12, "bold")
        self.bg_color = "#f0f0f0"
        
    def center_window(self):
        """使窗口居中显示"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
    def create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = tk.Frame(self.root, padx=20, pady=20, bg=self.bg_color)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = tk.Label(main_frame, text="抽卡回合", 
                               font=("楷体", 16, "bold"), 
                               fg="#2c3e50", bg=self.bg_color)
        title_label.pack(pady=(0, 20))
        
        # 说明文字
        #info_text = "功能说明：根据Excel配置，将源文件夹中的图片按页数分割到不同的目标文件夹\n"
        #info_text += "匹配规则：总文件名前15位（如：230-QQ0310-0001）与源文件夹名前15位完全匹配"
        #info_label = tk.Label(main_frame, text=info_text, 
        #                     font=("微软雅黑", 9), fg="#7f8c8d", 
        #                    bg=self.bg_color, justify=tk.LEFT)
        #info_label.pack(pady=(0, 20))
        
        # 文件选择框架
        file_frame = tk.LabelFrame(main_frame, text="文件路径设置", 
                                   font=self.font_title, 
                                   bg=self.bg_color, padx=10, pady=10)
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Excel文件选择
        excel_frame = tk.Frame(file_frame, bg=self.bg_color)
        excel_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(excel_frame, text="Excel文件：", font=self.font_default, 
                width=12, anchor=tk.W, bg=self.bg_color).pack(side=tk.LEFT)
        
        self.excel_path_var = tk.StringVar()
        self.excel_entry = tk.Entry(excel_frame, textvariable=self.excel_path_var, 
                                   font=self.font_default, width=50)
        self.excel_entry.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        
        tk.Button(excel_frame, text="浏览", command=self.browse_excel, 
                 font=self.font_default, width=8, bg="#3498db", fg="white").pack(side=tk.RIGHT)
        
        # 源文件夹选择
        source_frame = tk.Frame(file_frame, bg=self.bg_color)
        source_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(source_frame, text="源文件夹：", font=self.font_default, 
                width=12, anchor=tk.W, bg=self.bg_color).pack(side=tk.LEFT)
        
        self.source_path_var = tk.StringVar()
        self.source_entry = tk.Entry(source_frame, textvariable=self.source_path_var, 
                                    font=self.font_default, width=50)
        self.source_entry.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        
        tk.Button(source_frame, text="浏览", command=self.browse_source, 
                 font=self.font_default, width=8, bg="#3498db", fg="white").pack(side=tk.RIGHT)
        
        # 输出文件夹选择
        output_frame = tk.Frame(file_frame, bg=self.bg_color)
        output_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(output_frame, text="输出文件夹：", font=self.font_default, 
                width=12, anchor=tk.W, bg=self.bg_color).pack(side=tk.LEFT)
        
        self.output_path_var = tk.StringVar()
        self.output_entry = tk.Entry(output_frame, textvariable=self.output_path_var, 
                                    font=self.font_default, width=50)
        self.output_entry.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        
        tk.Button(output_frame, text="浏览", command=self.browse_output, 
                 font=self.font_default, width=8, bg="#3498db", fg="white").pack(side=tk.RIGHT)
        
        # 选项框架
        option_frame = tk.LabelFrame(main_frame, text="处理选项", 
                                    font=self.font_title, 
                                    bg=self.bg_color, padx=10, pady=10)
        option_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 工作表名称
        sheet_frame = tk.Frame(option_frame, bg=self.bg_color)
        sheet_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(sheet_frame, text="工作表名：", font=self.font_default, 
                width=12, anchor=tk.W, bg=self.bg_color).pack(side=tk.LEFT)
        
        self.sheet_name_var = tk.StringVar(value="抽象")
        self.sheet_entry = tk.Entry(sheet_frame, textvariable=self.sheet_name_var, 
                                   font=self.font_default, width=20)
        self.sheet_entry.pack(side=tk.LEFT)
        
        tk.Label(sheet_frame, text="（默认为'抽象'）", font=("楷体", 9), 
                fg="#7f8c8d", bg=self.bg_color).pack(side=tk.LEFT, padx=(10, 0))
        
        # 控制按钮框架
        button_frame = tk.Frame(main_frame, bg=self.bg_color)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.start_button = tk.Button(button_frame, text="开始抽卡", 
                                      command=self.start_processing,
                                      font=("楷体", 11, "bold"), 
                                      width=15, height=1,
                                      bg="#27ae60", fg="white")
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.clear_button = tk.Button(button_frame, text="清空日志", 
                                      command=self.clear_log,
                                      font=self.font_default, 
                                      width=10,
                                      bg="#95a5a6", fg="white")
        self.clear_button.pack(side=tk.LEFT, padx=5)
        
        # 日志显示区域
        log_frame = tk.LabelFrame(main_frame, text="处理日志", 
                                 font=self.font_title, 
                                 bg=self.bg_color, padx=10, pady=10)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        # 创建滚动条
        scrollbar = tk.Scrollbar(log_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.log_text = tk.Text(log_frame, font=("Consolas", 9), 
                               wrap=tk.WORD, 
                               yscrollcommand=scrollbar.set,
                               bg="white", fg="#2c3e50")
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        scrollbar.config(command=self.log_text.yview)
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, 
                                           maximum=100, length=400)
        self.progress_bar.pack(fill=tk.X, pady=(10, 0))
        
    def browse_excel(self):
        """浏览Excel文件"""
        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if filename:
            self.excel_path_var.set(filename)
            
    def browse_source(self):
        """浏览源文件夹"""
        foldername = filedialog.askdirectory(title="选择包含图片的源文件夹")
        if foldername:
            self.source_path_var.set(foldername)
            
    def browse_output(self):
        """浏览输出文件夹"""
        foldername = filedialog.askdirectory(title="选择输出文件夹")
        if foldername:
            self.output_path_var.set(foldername)
            
    def log_message(self, message, level="INFO"):
        """在日志区域添加消息"""
        import datetime
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        
        # 根据级别设置颜色
        color_map = {
            "INFO": "#2c3e50",
            "WARNING": "#f39c12",
            "ERROR": "#e74c3c",
            "SUCCESS": "#27ae60"
        }
        
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        
        # 滚动到底部
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def clear_log(self):
        """清空日志"""
        self.log_text.delete(1.0, tk.END)
        
    def update_progress(self, value, message=""):
        """更新进度条"""
        self.progress_var.set(value)
        if message:
            self.log_message(message)
        self.root.update_idletasks()
        
    def start_processing(self):
        """开始处理（在新线程中运行）"""
        # 验证输入
        excel_path = self.excel_path_var.get().strip()
        source_folder = self.source_path_var.get().strip()
        output_folder = self.output_path_var.get().strip()
        
        if not excel_path:
            messagebox.showerror("错误", "请选择Excel文件！")
            return
        
        if not source_folder:
            messagebox.showerror("错误", "请选择源文件夹！")
            return
            
        if not output_folder:
            messagebox.showerror("错误", "请选择输出文件夹！")
            return
            
        if not os.path.exists(excel_path):
            messagebox.showerror("错误", f"Excel文件不存在：{excel_path}")
            return
            
        if not os.path.exists(source_folder):
            messagebox.showerror("错误", f"源文件夹不存在：{source_folder}")
            return
            
        # 禁用开始按钮
        self.start_button.config(state=tk.DISABLED, text="处理中...")
        self.clear_log()
        
        # 在新线程中运行处理逻辑
        thread = threading.Thread(target=self.process_images, 
                                 args=(excel_path, source_folder, output_folder))
        thread.daemon = True
        thread.start()
        
    def process_images(self, excel_path, source_folder, output_folder):
        """处理图片的主逻辑"""
        try:
            # 创建输出文件夹
            os.makedirs(output_folder, exist_ok=True)
            self.log_message(f"输出文件夹已创建/确认: {output_folder}", "SUCCESS")
            
            # 读取Excel文件
            sheet_name = self.sheet_name_var.get().strip() or "抽象"
            self.log_message(f"正在读取Excel文件: {excel_path} (工作表: {sheet_name})")
            
            try:
                excel_data = pd.read_excel(excel_path, sheet_name=sheet_name)
            except Exception as e:
                self.log_message(f"读取Excel文件失败: {e}", "ERROR")
                self.log_message("请确保工作表名称正确", "WARNING")
                self.start_button.config(state=tk.NORMAL, text="开始处理")
                return
            
            # 检查必要的列
            required_columns = ['总文件名', '起始页', '分件文件名', '总页数', '每份文件页数']
            missing_columns = [col for col in required_columns if col not in excel_data.columns]
            
            if missing_columns:
                self.log_message(f"错误：Excel文件中缺少必要的列: {missing_columns}", "ERROR")
                self.log_message(f"当前列: {list(excel_data.columns)}", "WARNING")
                self.start_button.config(state=tk.NORMAL, text="开始处理")
                return
            
            self.update_progress(10, "Excel文件读取成功")
            
            # 清理总文件名数据，提取前15位用于精确匹配
            self.log_message("\n正在提取总文件名的前15位用于匹配...")
            excel_data['总文件名_匹配键'] = excel_data['总文件名'].apply(
                lambda x: self.extract_first_15_chars(str(x))
            )
            
            # 获取源文件夹中所有子文件夹的映射
            self.log_message("正在扫描源文件夹结构...")
            folder_mapping = self.scan_source_folders(source_folder)
            
            if not folder_mapping:
                self.log_message("错误：源文件夹中没有找到任何子文件夹", "ERROR")
                self.start_button.config(state=tk.NORMAL, text="开始处理")
                return
            
            self.log_message(f"找到 {len(folder_mapping)} 个源文件夹")
            for folder in list(folder_mapping.keys())[:10]:  # 显示前10个
                self.log_message(f"  - {folder}")
            if len(folder_mapping) > 10:
                self.log_message(f"  ... 还有 {len(folder_mapping) - 10} 个文件夹")
            
            self.update_progress(20, "源文件夹扫描完成")
            
            # 按匹配键分组
            self.log_message("\n正在按总文件名分组...")
            grouped = excel_data.groupby('总文件名_匹配键')
            
            # 存储所有分组信息用于勘误检查
            group_info = []
            processed_count = 0
            skipped_count = 0
            total_groups = len(grouped)
            
            # 处理每个分组
            for idx, (group_key, group_data) in enumerate(grouped):
                original_group_name = str(group_data.iloc[0]['总文件名']).strip()
                
                self.log_message(f"\n处理分组: {original_group_name}")
                self.log_message(f"匹配键(前15位): {group_key}")
                
                # 查找匹配的源文件夹（精确匹配前15位）
                matched_folder = self.find_matching_folder_by_prefix(folder_mapping, group_key)
                
                if not matched_folder:
                    self.log_message(f"  跳过：未找到前15位与 '{group_key}' 匹配的源文件夹", "WARNING")
                    skipped_count += 1
                    continue
                
                self.log_message(f"  找到匹配的源文件夹: {matched_folder}", "SUCCESS")
                
                # 获取分组中的理论总页数
                expected_total_pages = int(group_data.iloc[0]['总页数'])
                
                # 检查分组内总页数是否一致
                if not (group_data['总页数'] == expected_total_pages).all():
                    self.log_message(f"  警告：分组 '{original_group_name}' 中的总页数不一致！", "WARNING")
                    self.log_message(f"  总页数值: {group_data['总页数'].unique()}")
                
                # 获取源文件夹中的所有图片文件
                source_group_folder = folder_mapping[matched_folder]
                
                # 验证源文件夹是否存在
                if not os.path.exists(source_group_folder):
                    self.log_message(f"  错误：源文件夹不存在 - {source_group_folder}", "ERROR")
                    skipped_count += 1
                    continue
                    
                image_files = self.get_image_files(source_group_folder)
                
                if not image_files:
                    self.log_message(f"  警告：文件夹 '{matched_folder}' 中没有图片文件，跳过此分组", "WARNING")
                    skipped_count += 1
                    continue
                
                total_images = len(image_files)
                self.log_message(f"  找到 {total_images} 张图片")
                self.log_message(f"  理论总页数（总页数）: {expected_total_pages}")
                
                # 对分组数据按起始页排序
                group_data = group_data.sort_values('起始页')
                group_data = group_data.reset_index(drop=True)
                
                # 清理文件名，移除非法字符
                safe_group_name = self.sanitize_filename(original_group_name)
                
                # 创建分组的目标文件夹
                target_group_folder = os.path.join(output_folder, safe_group_name)
                try:
                    os.makedirs(target_group_folder, exist_ok=True)
                    self.log_message(f"  创建目标文件夹: {target_group_folder}")
                except Exception as e:
                    self.log_message(f"  创建目标文件夹失败: {e}", "ERROR")
                    skipped_count += 1
                    continue
                
                # 处理分组中的每一行
                processed_images = 0
                subfolder_image_counts = []
                
                for row_idx, row in group_data.iterrows():
                    start_page = int(row['起始页'])
                    folder_name = str(row['分件文件名']).strip()
                    expected_sub_pages = int(row['每份文件页数'])
                    
                    # 清理文件夹名
                    safe_folder_name = self.sanitize_filename(folder_name)
                    
                    # 计算开始索引
                    start_idx = start_page - 1
                    
                    # 如果是最后一行，处理所有剩余图片
                    if row_idx == len(group_data) - 1:
                        end_idx = total_images
                    else:
                        next_start_page = int(group_data.iloc[row_idx + 1]['起始页'])
                        end_idx = next_start_page - 1
                    
                    # 确保索引有效
                    if start_idx < 0:
                        start_idx = 0
                    if end_idx > total_images:
                        end_idx = total_images
                    
                    # 检查是否有图片需要处理
                    if start_idx >= total_images:
                        self.log_message(f"  警告：起始页码 {start_page} 超出图片总数 {total_images}，跳过此行", "WARNING")
                        subfolder_image_counts.append({
                            'folder_name': folder_name,
                            'expected': expected_sub_pages,
                            'actual': 0
                        })
                        continue
                    
                    # 创建目标子文件夹
                    target_subfolder = os.path.join(target_group_folder, safe_folder_name)
                    try:
                        os.makedirs(target_subfolder, exist_ok=True)
                    except Exception as e:
                        self.log_message(f"  创建子文件夹失败 '{safe_folder_name}': {e}", "ERROR")
                        subfolder_image_counts.append({
                            'folder_name': folder_name,
                            'expected': expected_sub_pages,
                            'actual': 0
                        })
                        continue
                    
                    # 计算需要复制的图片数量
                    images_to_copy = end_idx - start_idx
                    
                    self.log_message(f"  分割图片: {start_idx+1}-{end_idx} ({images_to_copy}张) 到 '{folder_name}'")
                    self.log_message(f"    预期页数: {expected_sub_pages} 页")
                    
                    copied_count = 0
                    for i in range(start_idx, end_idx):
                        try:
                            src_file = image_files[i]
                            dest_file = os.path.join(target_subfolder, os.path.basename(src_file))
                            shutil.copy2(src_file, dest_file)
                            copied_count += 1
                        except Exception as e:
                            self.log_message(f"    复制文件失败 {os.path.basename(src_file)}: {e}", "ERROR")
                    
                    self.log_message(f"    已复制 {copied_count} 张图片")
                    processed_images += copied_count
                    
                    # 记录子文件夹的图片数
                    subfolder_image_counts.append({
                        'folder_name': folder_name,
                        'expected': expected_sub_pages,
                        'actual': copied_count
                    })
                
                # 检查是否所有图片都被处理了
                if processed_images < total_images:
                    self.log_message(f"  警告：有 {total_images - processed_images} 张图片未被处理！", "WARNING")
                
                # 记录分组信息
                group_info.append({
                    'group_name': original_group_name,
                    'matched_folder': matched_folder,
                    'total_pages': total_images,
                    'expected_pages': expected_total_pages,
                    'file_count': len(group_data),
                    'processed_images': processed_images,
                    'subfolder_counts': subfolder_image_counts
                })
                
                processed_count += 1
                
                # 更新进度
                progress = 20 + (idx + 1) / total_groups * 60
                self.update_progress(progress, f"已完成 {idx + 1}/{total_groups} 个分组")
            
            # 执行勘误检查
            self.update_progress(85, "\n正在执行勘误检查...")
            self.run_error_check(group_info)
            
            # 输出统计信息
            self.update_progress(95, "\n处理统计:")
            self.log_message(f"  成功处理的分组: {processed_count}")
            self.log_message(f"  跳过的分组: {skipped_count}")
            self.log_message(f"  总分组数: {total_groups}")
            
            if group_info:
                self.log_message("\n分组处理详情:")
                for info in group_info:
                    status = "✓" if info['processed_images'] == info['total_pages'] else "✗"
                    page_match = "✓" if info['total_pages'] == info['expected_pages'] else "✗"
                    self.log_message(f"  {info['group_name']}: 处理{status}, 页数{page_match} ({info['processed_images']}/{info['total_pages']}张, 理论{info['expected_pages']}页)")
            
            self.update_progress(100, f"\n抽卡完成！已保存到: {output_folder}")
            messagebox.showinfo("完成", "抽卡完成！")
            
        except Exception as e:
            self.log_message(f"处理过程中发生错误: {str(e)}", "ERROR")
            messagebox.showerror("错误", f"处理失败：{str(e)}")
        finally:
            # 恢复开始按钮
            self.start_button.config(state=tk.NORMAL, text="开始处理")
    
    def extract_first_15_chars(self, text):
        """提取字符串的前15个字符"""
        if pd.isna(text):
            return ""
        text = str(text).strip()
        return text[:15] if len(text) >= 15 else text
    
    def scan_source_folders(self, base_folder):
        """扫描源文件夹，使用前15位作为匹配键"""
        folder_mapping = {}
        
        try:
            if os.path.exists(base_folder):
                for item in os.listdir(base_folder):
                    item_path = os.path.join(base_folder, item)
                    if os.path.isdir(item_path):
                        # 使用文件夹名的前15位作为匹配键
                        match_key = self.extract_first_15_chars(item)
                        if match_key:
                            folder_mapping[match_key] = item_path
                            self.log_message(f"    映射: '{match_key}' -> '{item}'")
        except Exception as e:
            self.log_message(f"扫描源文件夹时出错: {e}", "ERROR")
        
        return folder_mapping
    
    def find_matching_folder_by_prefix(self, folder_mapping, group_key):
        """根据前15位精确匹配文件夹"""
        # 直接匹配前15位
        if group_key in folder_mapping:
            return group_key
        
        # 如果没有完全匹配，尝试不区分大小写的匹配
        for key in folder_mapping.keys():
            if key.lower() == group_key.lower():
                return key
        
        return None
    
    def get_image_files(self, folder_path):
        """获取文件夹中的所有图片文件（按数字顺序排序）"""
        image_extensions = ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.tif', '.webp']
        image_files = []
        
        try:
            # 首先收集所有图片文件
            for file in os.listdir(folder_path):
                file_path = os.path.join(folder_path, file)
                if os.path.isfile(file_path):
                    ext = os.path.splitext(file)[1].lower()
                    if ext in image_extensions:
                        image_files.append(file_path)
            
            # 按文件名中的数字部分排序
            def extract_numbers(filename):
                """从文件名中提取数字用于排序"""
                basename = os.path.basename(filename)
                numbers = re.findall(r'\d+', basename)
                return [int(num) for num in numbers] if numbers else [0]
            
            # 按数字顺序排序
            image_files.sort(key=extract_numbers)
            
            # 如果没有数字，按字母顺序排序
            if not any(re.search(r'\d+', os.path.basename(f)) for f in image_files):
                image_files.sort(key=lambda x: os.path.basename(x))
        except Exception as e:
            self.log_message(f"获取图片文件时出错: {e}", "ERROR")
        
        return image_files
    
    def sanitize_filename(self, filename):
        """清理文件名，移除或替换非法字符"""
        illegal_chars = '<>:"/\\|?*'
        for char in illegal_chars:
            filename = filename.replace(char, '_')
        
        filename = filename.strip('. ')
        
        if not filename:
            filename = "unnamed"
        
        return filename
    
    def run_error_check(self, group_info):
        """执行勘误检查"""
        errors = []
        
        self.log_message("\n" + "="*50)
        self.log_message("勘误检查详情:")
        
        for info in group_info:
            group_name = info['group_name']
            total_pages = info['total_pages']
            expected_pages = info['expected_pages']
            file_count = info['file_count']
            subfolder_counts = info.get('subfolder_counts', [])
            
            self.log_message(f"\n分组 '{group_name}':")
            self.log_message(f"  理论总页数: {expected_pages}")
            self.log_message(f"  实际图片数: {total_pages}")
            self.log_message(f"  文件份数: {file_count}")
            
            # 检查总页数
            if total_pages != expected_pages:
                difference = abs(total_pages - expected_pages)
                if total_pages > expected_pages:
                    issue_desc = f"实际图片数({total_pages})多于理论总页数({expected_pages})，多出{difference}页"
                else:
                    issue_desc = f"实际图片数({total_pages})少于理论总页数({expected_pages})，缺少{difference}页"
                
                errors.append({
                    'group': group_name,
                    'issue': f"总页数不匹配: {issue_desc}",
                    'details': f"理论总页数: {expected_pages}, 实际图片数: {total_pages}"
                })
                self.log_message(f"  ✗ {issue_desc}", "ERROR")
            else:
                self.log_message(f"  ✓ 总页数匹配正确", "SUCCESS")
            
            # 检查子文件夹页数
            if subfolder_counts:
                self.log_message(f"  子文件夹页数检查:")
                for sub_info in subfolder_counts:
                    folder_name = sub_info['folder_name']
                    expected = sub_info['expected']
                    actual = sub_info['actual']
                    
                    if actual != expected:
                        if actual > expected:
                            detail = f"预期 {expected} 页，实际 {actual} 页，多出 {actual - expected} 页"
                        else:
                            detail = f"预期 {expected} 页，实际 {actual} 页，缺少 {expected - actual} 页"
                        
                        errors.append({
                            'group': group_name,
                            'issue': f"子文件夹 '{folder_name}' 页数不匹配",
                            'details': detail
                        })
                        self.log_message(f"    ✗ {folder_name}: 预期 {expected} 页，实际 {actual} 页", "ERROR")
                    else:
                        self.log_message(f"    ✓ {folder_name}: 预期 {expected} 页，实际 {actual} 页", "SUCCESS")
            
            # 检查是否完全处理
            if info['processed_images'] != total_pages:
                errors.append({
                    'group': group_name,
                    'issue': f"图片未完全处理",
                    'details': f"仅处理了 {info['processed_images']}/{total_pages} 张，有 {total_pages - info['processed_images']} 张未被处理"
                })
                self.log_message(f"  ✗ 图片未完全处理: {info['processed_images']}/{total_pages} 张", "ERROR")
            else:
                self.log_message(f"  ✓ 图片完全处理", "SUCCESS")
        
        # 输出汇总
        self.log_message("\n" + "="*50)
        if errors:
            self.log_message(f"发现 {len(errors)} 个勘误问题:", "WARNING")
            for error in errors:
                self.log_message(f"\n分组: {error['group']}", "WARNING")
                self.log_message(f"问题: {error['issue']}", "WARNING")
                self.log_message(f"详情: {error['details']}", "WARNING")
        else:
            self.log_message("✓ 勘误检查全部通过！所有分组总页数和子文件夹页数都匹配正确。", "SUCCESS")

def main():
    root = tk.Tk()
    app = ImageSplitterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
