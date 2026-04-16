import pandas as pd
import os
from datetime import datetime
import re

def process_excel_file():
    """处理Excel/WPS表格，统计指定数据"""
    
    # 手动输入文件路径
    file_path = input("请输入Excel/WPS文件路径: ").strip()
    
    # 去除路径两端的引号（如果用户复制路径时带了引号）
    file_path = file_path.strip('"').strip("'")
    
    # 检查文件是否存在
    if not os.path.exists(file_path):
        print(f"错误：文件 '{file_path}' 不存在！")
        return
    
    try:
        # 读取Excel文件
        excel_file = pd.ExcelFile(file_path)
        
        # 检查是否存在名为"demo"的工作表
        if "demo" not in excel_file.sheet_names:
            print(f"错误：找不到名为'demo'的工作表！")
            print(f"可用的工作表有：{excel_file.sheet_names}")
            return
        
        # 读取demo工作表，保持日期列的原始格式
        df = pd.read_excel(file_path, sheet_name="demo")
        print(f"\n成功读取工作表 'demo'，共 {len(df)} 行数据")
        
        # 显示列名，方便用户确认
        print("\n表格列名：")
        for i, col in enumerate(df.columns, 1):
            print(f"{i}. {col}")
        
        # 检查必要的列是否存在
        required_columns = ["源文件名", "日期", "页号"]
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            print(f"\n错误：缺少必要的列：{missing_columns}")
            return
        
        # 保存原始日期格式
        original_dates = df["日期"].copy()
        
        # 将日期转换为可比较的格式（用于找最大最小值）
        df["日期_转换"] = df["日期"].apply(parse_date_for_comparison)
        
        # 创建结果DataFrame的列表
        result_data = []
        
        # 按源文件名分组处理
        for source_name, group_df in df.groupby("源文件名"):
            print(f"\n处理分组: {source_name}")
            
            # 统计该分组中的文件数量（行数）
            file_count = len(group_df)
            
            # 找到最大日期和最小日期（基于转换后的日期）
            max_date_idx = group_df["日期_转换"].idxmax()
            min_date_idx = group_df["日期_转换"].idxmin()
            
            # 获取原始格式的日期
            max_date_original = original_dates.loc[max_date_idx]
            min_date_original = original_dates.loc[min_date_idx]
            
            # 确保日期格式统一为8位数字字符串
            max_date_str = format_date_to_8digits(max_date_original)
            min_date_str = format_date_to_8digits(min_date_original)
            
            # 处理页数统计
            # 对分组按索引排序（保持原始顺序）
            group_df_sorted = group_df.sort_index()
            
            # 提取页号列
            page_numbers = group_df_sorted["页号"].astype(str).tolist()
            
            # 计算每个文件的页数
            page_counts = []
            
            for i in range(len(page_numbers)):
                if i < len(page_numbers) - 1:
                    # 不是最后一个文件：页数 = 下一个文件的起始页号 - 当前文件的起始页号
                    current_page = extract_start_page(page_numbers[i])
                    next_page = extract_start_page(page_numbers[i + 1])
                    
                    if current_page is not None and next_page is not None:
                        page_count = next_page - current_page
                    else:
                        page_count = 0
                    page_counts.append(page_count)
                else:
                    # 最后一个文件
                    if '-' in page_numbers[i]:
                        # 格式如"37-38"：页数 = 结束页 - 起始页 + 1
                        start_page, end_page = extract_page_range(page_numbers[i])
                        if start_page is not None and end_page is not None:
                            page_count = end_page - start_page + 1
                            
                            # 调整倒数第二个文件的页数
                            if i > 0:
                                prev_page = extract_start_page(page_numbers[i - 1])
                                if prev_page is not None and start_page is not None:
                                    page_counts[i - 1] = start_page - prev_page
                        else:
                            page_count = 1
                    else:
                        # 最后一个文件不是范围格式，无法确定页数，设为0或特殊标记
                        page_count = 0
                    
                    page_counts.append(page_count)
            
            # 为分组中的每一行添加统计信息
            for idx, (row_idx, row) in enumerate(group_df_sorted.iterrows()):
                result_row = row.drop("日期_转换").to_dict()  # 移除临时列
                result_row["文件总数"] = file_count
                result_row["最大日期"] = max_date_str
                result_row["最小日期"] = min_date_str
                result_row["页数"] = page_counts[idx] if idx < len(page_counts) else 0
                result_data.append(result_row)
            
            print(f"  - 文件数量: {file_count}")
            print(f"  - 日期范围: {min_date_str} 至 {max_date_str}")
            print(f"  - 页数统计: {page_counts}")
        
        # 创建结果DataFrame
        result_df = pd.DataFrame(result_data)
        
        # 确保日期列保持原始格式
        if "日期" in result_df.columns:
            result_df["日期"] = result_df["日期"].apply(format_date_to_8digits)
        
        # 生成输出文件名
        output_dir = os.path.dirname(file_path)
        output_filename = f"统计结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        output_path = os.path.join(output_dir, output_filename)
        
        # 保存结果
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            result_df.to_excel(writer, sheet_name='统计结果', index=False)
            
            # 创建汇总表
            summary_data = []
            for source_name, group_df in result_df.groupby("源文件名"):
                summary_row = {
                    "源文件名": source_name,
                    "文件总数": group_df["文件总数"].iloc[0],
                    "最大日期": group_df["最大日期"].iloc[0],
                    "最小日期": group_df["最小日期"].iloc[0],
                    "总页数": group_df["页数"].sum()
                }
                summary_data.append(summary_row)
            
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='汇总统计', index=False)
        
        print(f"\n✅ 处理完成！")
        print(f"结果已保存至: {output_path}")
        print(f"  - '统计结果' 工作表：包含原始数据和新增的统计列")
        print(f"  - '汇总统计' 工作表：按源文件名汇总的统计信息")
        
    except Exception as e:
        print(f"\n❌ 处理过程中出现错误：{str(e)}")
        import traceback
        traceback.print_exc()

def parse_date_for_comparison(date_value):
    """将日期转换为可比较的格式"""
    if pd.isna(date_value):
        return pd.NaT
    
    # 如果是字符串
    if isinstance(date_value, str):
        date_str = date_value.strip()
        
        # 处理8位数字格式：20170808
        if re.match(r'^\d{8}$', date_str):
            try:
                return pd.to_datetime(date_str, format='%Y%m%d')
            except:
                pass
        
        # 处理带分隔符的格式：2017-08-08, 2017/08/08
        for fmt in ['%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%d-%m-%Y', '%d/%m/%Y']:
            try:
                return pd.to_datetime(date_str, format=fmt)
            except:
                continue
    
    # 如果是数字类型（如20170808）
    if isinstance(date_value, (int, float)):
        date_str = str(int(date_value))
        if len(date_str) == 8:
            try:
                return pd.to_datetime(date_str, format='%Y%m%d')
            except:
                pass
    
    # 尝试自动解析
    try:
        return pd.to_datetime(date_value)
    except:
        return pd.NaT

def format_date_to_8digits(date_value):
    """将日期格式化为8位数字字符串"""
    if pd.isna(date_value):
        return ""
    
    # 如果已经是8位数字字符串
    if isinstance(date_value, str):
        date_str = date_value.strip()
        if re.match(r'^\d{8}$', date_str):
            return date_str
    
    # 如果是整数类型
    if isinstance(date_value, (int, float)):
        date_str = str(int(date_value))
        if len(date_str) == 8:
            return date_str
    
    # 如果是datetime对象
    if isinstance(date_value, (pd.Timestamp, datetime)):
        return date_value.strftime('%Y%m%d')
    
    # 尝试转换为字符串并清理
    date_str = str(date_value).strip()
    
    # 移除分隔符，只保留数字
    digits_only = re.sub(r'\D', '', date_str)
    
    # 如果是8位数字，返回
    if len(digits_only) == 8:
        return digits_only
    
    # 如果无法处理，返回原值
    return date_str

def extract_start_page(page_str):
    """提取起始页号"""
    if pd.isna(page_str) or page_str == '':
        return None
    
    page_str = str(page_str)
    
    # 如果是范围格式如"37-38"，提取前面的数字
    if '-' in page_str:
        match = re.match(r'(\d+)-', page_str)
        if match:
            return int(match.group(1))
    else:
        # 提取第一个数字
        match = re.search(r'\d+', page_str)
        if match:
            return int(match.group())
    
    return None

def extract_page_range(page_str):
    """提取页码范围，返回(起始页, 结束页)"""
    if pd.isna(page_str) or page_str == '':
        return None, None
    
    page_str = str(page_str)
    
    # 匹配格式如"37-38"
    match = re.match(r'(\d+)-(\d+)', page_str)
    if match:
        return int(match.group(1)), int(match.group(2))
    
    return None, None

if __name__ == "__main__":
    print("=" * 50)
    print("Excel/WPS表格数据统计工具")
    print("=" * 50)
    
    # 检查必要的库是否安装
    try:
        import openpyxl
    except ImportError:
        print("\n警告：未安装 openpyxl 库，正在尝试安装...")
        os.system("pip install openpyxl")
        print("请重新运行程序")
        exit()
    
    process_excel_file()
    
    input("\n按回车键退出...")
