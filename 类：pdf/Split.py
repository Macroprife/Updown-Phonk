import os
import pandas as pd
import PyPDF2
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import warnings
warnings.filterwarnings('ignore')

class PDFSplitter:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()  # 隐藏主窗口
        
    def select_file(self, file_type="Excel", extension="*.xlsx;*.xls"):
        """选择文件"""
        title = f"选择{file_type}文件"
        file_path = filedialog.askopenfilename(
            title=title,
            filetypes=[(f"{file_type}文件", extension), ("所有文件", "*.*")]
        )
        return file_path
    
    def select_folder(self, title="选择文件夹"):
        """选择文件夹"""
        folder_path = filedialog.askdirectory(title=title)
        return folder_path
    
    def split_pdf(self, excel_path, pdf_folder, output_folder):
        """主函数：分割PDF文件"""
        try:
            # 1. 读取Excel文件
            print(f"读取Excel文件: {excel_path}")
            excel_file = pd.ExcelFile(excel_path)
            
            # 查找名为"抽象"的工作表
            if '抽象' not in excel_file.sheet_names:
                # 尝试其他可能的名称
                sheet_names = excel_file.sheet_names
                print(f"未找到名为'抽象'的工作表，可用的工作表有: {sheet_names}")
                
                # 尝试自动查找包含关键字的表
                possible_sheets = [name for name in sheet_names if '抽象' in name] # type: ignore
                if not possible_sheets:
                    # 让用户选择工作表
                    choice = input(f"请选择工作表（输入序号）:\n" + 
                                 "\n".join([f"{i+1}. {name}" for i, name in enumerate(sheet_names)]))
                    sheet_name = sheet_names[int(choice)-1]
                else:
                    sheet_name = possible_sheets[0]
                    print(f"使用工作表: {sheet_name}")
            else:
                sheet_name = '抽象'
            
            # 读取工作表
            df = pd.read_excel(excel_path, sheet_name=sheet_name)
            print(f"成功读取工作表 '{sheet_name}'，共 {len(df)} 行数据")
            
            # 检查必要的列是否存在
            required_columns = ['图片文件名', '操作列', '操作档号']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                print(f"缺少必要的列: {missing_columns}")
                print(f"可用的列有: {list(df.columns)}")
                return False
            
            # 2. 按"操作档号"分组
            df_grouped = df.groupby('操作档号')
            print(f"找到 {len(df_grouped)} 个分组")
            
            # 3. 创建输出文件夹
            os.makedirs(output_folder, exist_ok=True)
            
            # 4. 处理每个分组
            total_pdfs_processed = 0
            skipped_groups = []
            
            for group_name, group_df in df_grouped:
                print(f"\n处理分组: {group_name}")
                
                # 创建分组文件夹（去除非法字符）
                safe_group_name = self.sanitize_filename(str(group_name))
                group_folder = os.path.join(output_folder, safe_group_name)
                os.makedirs(group_folder, exist_ok=True)
                
                # 获取该分组对应的PDF文件名（假设PDF文件名与操作档号的第一部分匹配）
                # 提取操作档号的基本部分（去掉中文等）
                base_name = str(group_name).split()[0] if ' ' in str(group_name) else str(group_name)
                
                # 在PDF文件夹中查找匹配的PDF文件
                matching_pdfs = []
                for pdf_file in os.listdir(pdf_folder):
                    if pdf_file.lower().endswith('.pdf'):
                        pdf_name_without_ext = os.path.splitext(pdf_file)[0]
                        # 多种匹配方式
                        if (base_name in pdf_name_without_ext or 
                            pdf_name_without_ext in base_name):
                            matching_pdfs.append(pdf_file)
                
                if not matching_pdfs:
                    print(f"  警告: 未找到与 '{base_name}' 匹配的PDF文件，跳过此分组")
                    skipped_groups.append(group_name)
                    continue
                
                # 使用第一个匹配的PDF文件
                pdf_file_name = matching_pdfs[0]
                pdf_path = os.path.join(pdf_folder, pdf_file_name)
                
                print(f"  使用PDF文件: {pdf_file_name}")
                
                # 打开PDF文件
                with open(pdf_path, 'rb') as pdf_file:
                    pdf_reader = PyPDF2.PdfReader(pdf_file)
                    total_pages = len(pdf_reader.pages)
                    print(f"  PDF总页数: {total_pages}")
                    
                    # 对分组内的每一行数据进行处理
                    group_df = group_df.sort_values('操作列')
                    
                    # 存储分割后的剩余页码范围
                    remaining_pages = list(range(total_pages))  # 0-based索引
                    
                    # 处理每个分割点
                    for i, row in group_df.iterrows():
                        try:
                            img_filename = str(row['图片文件名']).strip()
                            operation_col = int(row['操作列'])
                            
                            # 创建输出文件夹
                            img_folder = os.path.join(group_folder, img_filename)
                            os.makedirs(img_folder, exist_ok=True)
                            
                            # 确定分割范围
                            if i == group_df.index[0]:  # 第一行
                                start_page = 0  # PDF页数从0开始
                            else:
                                prev_operation = int(group_df.loc[prev_idx, '操作列']) # type: ignore
                                start_page = prev_operation - 1  # 转换为0-based索引
                            
                            if operation_col > total_pages:
                                print(f"  警告: 操作列值 {operation_col} 超出PDF总页数，使用最大页数")
                                operation_col = total_pages
                            
                            end_page = operation_col - 1  # 转换为0-based索引
                            
                            # 分割PDF
                            if start_page < end_page and start_page < total_pages:
                                pdf_writer = PyPDF2.PdfWriter()
                                
                                for page_num in range(start_page, end_page):
                                    if page_num < total_pages:
                                        pdf_writer.add_page(pdf_reader.pages[page_num])
                                        # 从剩余页码中移除
                                        if page_num in remaining_pages:
                                            remaining_pages.remove(page_num)
                                
                                # 保存分割后的PDF
                                output_pdf_path = os.path.join(img_folder, f"{img_filename}.pdf")
                                with open(output_pdf_path, 'wb') as output_pdf:
                                    pdf_writer.write(output_pdf)
                                
                                print(f"  已创建: {output_pdf_path} (页数: {end_page - start_page})")
                                total_pdfs_processed += 1
                            
                            prev_idx = i
                            
                        except Exception as e:
                            print(f"  处理行数据时出错: {e}")
                            continue
                    
                    # 处理剩余页码（最后一组）
                    if remaining_pages:
                        # 使用最后一个图片文件名加上"_剩余"
                        last_img_filename = str(group_df.iloc[-1]['图片文件名']).strip()
                        remaining_img_filename = f"{last_img_filename}_剩余"
                        
                        # 创建剩余页码的文件夹
                        remaining_folder = os.path.join(group_folder, remaining_img_filename)
                        os.makedirs(remaining_folder, exist_ok=True)
                        
                        # 创建剩余页码的PDF
                        pdf_writer = PyPDF2.PdfWriter()
                        for page_num in remaining_pages:
                            if page_num < total_pages:
                                pdf_writer.add_page(pdf_reader.pages[page_num])
                        
                        if len(pdf_writer.pages) > 0:
                            output_pdf_path = os.path.join(remaining_folder, f"{remaining_img_filename}.pdf")
                            with open(output_pdf_path, 'wb') as output_pdf:
                                pdf_writer.write(output_pdf)
                            
                            print(f"  已创建剩余页码文件: {output_pdf_path} (页数: {len(pdf_writer.pages)})")
                            total_pdfs_processed += 1
            
            # 5. 输出统计信息
            print(f"\n{'='*50}")
            print(f"处理完成!")
            print(f"总处理分组数: {len(df_grouped)}")
            print(f"成功处理PDF数: {total_pdfs_processed}")
            if skipped_groups:
                print(f"跳过的分组: {skipped_groups}")
            print(f"输出文件夹: {output_folder}")
            
            return True
            
        except Exception as e:
            print(f"处理过程中出错: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def sanitize_filename(self, filename):
        """去除文件名中的非法字符"""
        illegal_chars = r'<>:"/\|?*'
        for char in illegal_chars:
            filename = filename.replace(char, '')
        return filename.strip()
    
    def run(self):
        """运行主程序"""
        print("="*50)
        print("PDF分割工具")
        print("="*50)
        
        # 1. 选择Excel文件
        print("\n步骤1: 选择Excel文件")
        excel_path = self.select_file("Excel", "*.xlsx;*.xls")
        if not excel_path:
            print("未选择Excel文件，程序退出")
            return
        
        # 2. 选择PDF文件夹
        print("\n步骤2: 选择包含PDF文件的文件夹")
        pdf_folder = self.select_folder("选择包含PDF文件的文件夹")
        if not pdf_folder:
            print("未选择PDF文件夹，程序退出")
            return
        
        # 3. 选择输出文件夹
        print("\n步骤3: 选择输出文件夹")
        output_folder = self.select_folder("选择输出文件夹")
        if not output_folder:
            print("未选择输出文件夹，程序退出")
            return
        
        # 4. 处理文件
        print(f"\n开始处理...")
        print(f"Excel文件: {excel_path}")
        print(f"PDF文件夹: {pdf_folder}")
        print(f"输出文件夹: {output_folder}")
        
        success = self.split_pdf(excel_path, pdf_folder, output_folder)
        
        if success:
            messagebox.showinfo("完成", "PDF分割处理完成！")
        else:
            messagebox.showerror("错误", "处理过程中出现错误，请查看控制台输出")
        
        input("\n按Enter键退出...")

if __name__ == "__main__":
    # 检查必要的库
    try:
        import pandas
        import PyPDF2
    except ImportError as e:
        print(f"缺少必要的库: {e}")
        print("请安装以下库:")
        print("pip install pandas PyPDF2 openpyxl")
        input("按Enter键退出...")
        exit(1)
    
    # 运行程序
    app = PDFSplitter()
    app.run()