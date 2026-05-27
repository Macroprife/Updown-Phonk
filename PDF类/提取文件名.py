import os
import pandas as pd
import re
from datetime import datetime
import tkinter as tk
from tkinter import filedialog

def clean_filename(filename):
    """
    规范化文件名：去掉扩展名、所有括号符号（中英文）和空格，保留括号内的内容
    
    Args:
        filename: 原始文件名（包含扩展名）
    
    Returns:
        规范化后的文件名（去掉括号符号、空格和扩展名，保留括号内内容）
    """
    # 先分离文件名和扩展名
    name_without_ext = os.path.splitext(filename)[0]
    
    # 去掉所有类型的括号符号（保留括号内的内容）
    cleaned = name_without_ext
    
    # 替换各种括号为空字符串（只去掉括号符号本身）
    bracket_chars = [
        '(', ')',           # 英文圆括号
        '（', '）',         # 中文圆括号
        '[', ']',           # 英文方括号
        '【', '】',         # 中文方括号
        '{', '}',           # 花括号
        '<', '>',           # 尖括号
        '《', '》',         # 书名号
        '「', '」',         # 日文引号
        '『', '』',         # 日文双引号
        '［', '］',         # 全角方括号
        '｛', '｝',         # 全角花括号
        '〈', '〉',         # 单书名号
    ]
    
    for char in bracket_chars:
        cleaned = cleaned.replace(char, '')
    
    # 去掉所有空格（包括半角空格、全角空格、制表符等）
    cleaned = re.sub(r'\s+', '', cleaned)  # 去掉所有空白字符
    cleaned = cleaned.replace('　', '')    # 去掉全角空格
    
    # 清理括号去除后可能留下的多余符号
    cleaned = re.sub(r'[-_]{2,}', '-', cleaned)  # 多个破折号或下划线合并为一个
    cleaned = cleaned.strip('-_')  # 去掉首尾的破折号和下划线
    
    return cleaned

def get_file_names(folder_path, extensions):
    """
    获取指定文件夹中所有指定扩展名的文件名
    
    Args:
        folder_path: 文件夹路径
        extensions: 文件扩展名列表
    
    Returns:
        包含文件名信息的列表
    """
    file_list = []
    
    # 遍历文件夹及其子文件夹
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            # 检查文件扩展名
            file_ext = os.path.splitext(file)[1].lower()
            if file_ext in extensions:
                full_path = os.path.join(root, file)
                relative_path = os.path.relpath(full_path, folder_path)
                
                # 生成规范文件名
                clean_name = clean_filename(file)
                
                file_info = {
                    '文件名': file,
                    '规范文件名': clean_name,
                    '完整路径': full_path,
                    '相对路径': relative_path,
                    '所在文件夹': os.path.dirname(relative_path) if os.path.dirname(relative_path) else '根目录',
                    '扩展名': file_ext,
                    '文件大小(KB)': round(os.path.getsize(full_path) / 1024, 2),
                    '提取时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
                file_list.append(file_info)
    
    return file_list

def sanitize_sheet_name(name):
    """
    处理工作表名称，确保符合Excel要求
    """
    invalid_chars = ['[', ']', ':', '*', '?', '/', '\\']
    for char in invalid_chars:
        name = name.replace(char, '_')
    
    if len(name) > 31:
        name = name[:31]
    
    return name

def select_folder_dialog():
    """打开文件夹选择对话框"""
    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        folder_path = filedialog.askdirectory(title="选择要提取的文件夹")
        root.destroy()
        return folder_path
    except:
        return None

def select_save_path():
    """选择保存路径"""
    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        file_path = filedialog.asksaveasfilename(
            title="保存Excel文件",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"批量文件清单_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        root.destroy()
        return file_path
    except:
        return None

def show_cleaning_demo():
    """
    演示文件名清理效果
    """
    print("\n" + "=" * 60)
    print("  文件名规范化示例")
    print("=" * 60)
    
    examples = [
        ("报告 (最终版).pdf", "去掉英文括号和空格，保留内容"),
        ("数据 【2024年】.xlsx", "去掉中文方括号和空格，保留内容"),
        ("文档 [草稿].docx", "去掉英文方括号和空格，保留内容"),
        ("计划 （待审核）.pdf", "去掉中文括号和空格，保留内容"),
        ("项目 {备份}.xlsx", "去掉花括号和空格，保留内容"),
        ("手册 <修订版>.pdf", "去掉尖括号和空格，保留内容"),
        ("材料 （第3版） [副本].xlsx", "去掉混合括号和空格，保留内容"),
        ("测试 (完整版) （含答案）.pdf", "去掉混合括号和空格，保留内容"),
        ("《项目总结》 最终版.pdf", "去掉书名号和空格，保留内容"),
        ("2024年 第一季度 报告.pdf", "去掉所有空格"),
        ("文件　名称　测试.pdf", "去掉全角空格"),
    ]
    
    for original, description in examples:
        cleaned = clean_filename(original)
        print(f"\n原始文件名：{original}")
        print(f"说明：{description}")
        print(f"规范后：{cleaned}")
    
    print("\n" + "=" * 60)
    input("按回车键继续...")

def process_multiple_folders():
    """
    主函数：连续处理多个文件夹
    """
    print("=" * 60)
    print("  文件批量提取工具 - 多文件夹版本（含规范文件名）")
    print("=" * 60)
    
    # 选择文件类型
    print("\n【第1步】选择要提取的文件类型：")
    print("1. 仅PDF文件")
    print("2. Excel表格文件（.xlsx, .xls）")
    print("3. CSV文件")
    print("4. 所有表格文件（Excel + CSV）")
    print("5. PDF + 所有表格文件")
    print("6. 自定义扩展名")
    print("7. 查看文件名规范化示例")
    
    while True:
        choice = input("\n请输入选择（1-7）：").strip()
        
        if choice == '7':
            show_cleaning_demo()
            continue
        
        extensions_map = {
            '1': ['.pdf'],
            '2': ['.xlsx', '.xls'],
            '3': ['.csv'],
            '4': ['.xlsx', '.xls', '.csv'],
            '5': ['.pdf', '.xlsx', '.xls', '.csv']
        }
        
        if choice in extensions_map:
            extensions = extensions_map[choice]
            break
        elif choice == '6':
            custom = input("请输入扩展名（用逗号分隔，如 .pdf,.xlsx）：").strip()
            extensions = [ext.strip().lower() for ext in custom.split(',') if ext.strip()]
            if extensions:
                extensions = [f".{ext}" if not ext.startswith('.') else ext for ext in extensions]
                break
            else:
                print("扩展名不能为空，请重新输入")
        else:
            print("无效选择，请重新输入")
    
    print(f"\n✅ 已选择文件类型：{', '.join(extensions)}")
    
    # 存储所有文件夹的数据
    all_folders_data = {}
    folder_count = 0
    
    print("\n【第2步】开始添加文件夹（可添加多个）")
    print("提示：您可以输入路径、输入 'browse' 浏览选择，或直接回车结束添加\n")
    
    while True:
        folder_input = input(f"请输入第 {folder_count + 1} 个文件夹路径（直接回车结束）：").strip()
        
        if not folder_input:
            if folder_count == 0:
                print("⚠️  尚未添加任何文件夹，请至少添加一个文件夹")
                continue
            else:
                break
        
        if folder_input.lower() == 'browse':
            folder_path = select_folder_dialog()
            if not folder_path:
                print("❌ 未选择文件夹，请重试")
                continue
            print(f"📁 已选择：{folder_path}")
        else:
            folder_path = folder_input
        
        if not os.path.exists(folder_path):
            print(f"❌ 错误：路径不存在 - {folder_path}")
            continue
        
        if not os.path.isdir(folder_path):
            print(f"❌ 错误：这不是一个文件夹 - {folder_path}")
            continue
        
        print(f"🔍 正在扫描文件夹 {folder_count + 1}：{folder_path}")
        file_list = get_file_names(folder_path, extensions)
        
        if not file_list:
            print(f"⚠️  该文件夹中没有找到符合条件的文件")
            continue_anyway = input("是否继续添加其他文件夹？(y/n，默认y)：").strip().lower()
            if continue_anyway == 'n':
                break
            continue
        
        folder_name = os.path.basename(folder_path)
        if not folder_name:
            folder_name = os.path.abspath(folder_path).replace(':\\', '_drive')
        
        original_name = folder_name
        counter = 1
        while folder_name in all_folders_data:
            folder_name = f"{original_name}_{counter}"
            counter += 1
        
        all_folders_data[folder_name] = {
            'path': folder_path,
            'files': file_list
        }
        
        folder_count += 1
        print(f"✅ 已添加文件夹：{folder_name}")
        print(f"   找到文件数：{len(file_list)} 个")
        
        # 显示清理示例
        print(f"\n   📝 文件名规范化示例（前5个有变化的）：")
        shown = 0
        for file in file_list[:10]:
            original_name = os.path.splitext(file['文件名'])[0]
            if original_name != file['规范文件名']:
                print(f"   {shown+1}. {file['文件名']}")
                print(f"      → {file['规范文件名']}{file['扩展名']}")
                shown += 1
                if shown >= 5:
                    break
        
        if shown == 0:
            print("   （所有文件名无需规范化）")
        
        print()
        
        continue_add = input("是否继续添加文件夹？(y/n，默认y)：").strip().lower()
        if continue_add == 'n':
            break
    
    if not all_folders_data:
        print("\n❌ 没有提取到任何文件，程序退出")
        return
    
    # 选择保存路径
    print("\n【第3步】选择输出文件保存位置")
    print("请输入保存路径，或输入 'browse' 浏览选择保存位置")
    
    save_input = input("保存路径：").strip()
    
    if save_input.lower() == 'browse' or not save_input:
        output_file = select_save_path()
        if not output_file:
            print("❌ 未选择保存路径，程序退出")
            return
    else:
        if not save_input.endswith('.xlsx'):
            save_input += '.xlsx'
        output_file = save_input
    
    output_dir = os.path.dirname(output_file)
    if output_dir and not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
        except Exception as e:
            print(f"❌ 无法创建输出目录：{e}")
            return
    
    # 生成Excel文件
    print(f"\n📝 正在生成Excel文件...")
    
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 创建汇总工作表
            summary_data = []
            for folder_name, data in all_folders_data.items():
                file_count = len(data['files'])
                
                ext_stats = {}
                for file in data['files']:
                    ext = file['扩展名']
                    ext_stats[ext] = ext_stats.get(ext, 0) + 1
                
                total_size = sum(file['文件大小(KB)'] for file in data['files'])
                
                summary_data.append({
                    '工作表名称': folder_name,
                    '文件夹路径': data['path'],
                    '文件总数': file_count,
                    '文件类型': ', '.join([f"{ext}({count})" for ext, count in ext_stats.items()]),
                    '总大小(KB)': round(total_size, 2),
                    '总大小(MB)': round(total_size / 1024, 2)
                })
            
            df_summary = pd.DataFrame(summary_data)
            df_summary.to_excel(writer, sheet_name='汇总信息', index=False)
            
            # 为每个文件夹创建工作表
            for folder_name, data in all_folders_data.items():
                df = pd.DataFrame(data['files'])
                # 调整列顺序
                column_order = ['文件名', '规范文件名', '扩展名', '完整路径', 
                              '相对路径', '所在文件夹', '文件大小(KB)', '提取时间']
                df = df[column_order]
                safe_name = sanitize_sheet_name(folder_name)
                df.to_excel(writer, sheet_name=safe_name, index=False)
                
                # 调整列宽
                worksheet = writer.sheets[safe_name]
                worksheet.column_dimensions['A'].width = 35  # 文件名
                worksheet.column_dimensions['B'].width = 35  # 规范文件名
                worksheet.column_dimensions['C'].width = 10  # 扩展名
                worksheet.column_dimensions['D'].width = 50  # 完整路径
                worksheet.column_dimensions['E'].width = 40  # 相对路径
                worksheet.column_dimensions['F'].width = 30  # 所在文件夹
                worksheet.column_dimensions['G'].width = 15  # 文件大小
                worksheet.column_dimensions['H'].width = 20  # 提取时间
        
        # 输出结果统计
        print("\n" + "=" * 60)
        print("  ✅ 文件提取完成！")
        print("=" * 60)
        print(f"\n📁 输出文件：{output_file}")
        
        total_files = sum(len(data['files']) for data in all_folders_data.values())
        print(f"📊 总共处理了 {len(all_folders_data)} 个文件夹")
        print(f"📄 总文件数：{total_files} 个")
        
        # 统计有多少文件名被修改了
        modified_count = 0
        for data in all_folders_data.values():
            for file in data['files']:
                original_without_ext = os.path.splitext(file['文件名'])[0]
                if original_without_ext != file['规范文件名']:
                    modified_count += 1
        
        if modified_count > 0:
            print(f"📝 其中 {modified_count} 个文件经过了规范化处理")
        
        print(f"\n📋 Excel工作表结构：")
        print("   • 汇总信息 - 所有文件夹的统计概览")
        
        for i, (folder_name, data) in enumerate(all_folders_data.items(), 1):
            count = len(data['files'])
            safe_name = sanitize_sheet_name(folder_name)
            print(f"   • {safe_name} - {count} 个文件")
        
        print(f"\n📋 规范化规则：")
        print("   • 去掉所有中英文括号符号：() （） [] 【】 {} 《》等")
        print("   • 去掉所有空格（包括半角和全角空格）")
        print("   • 保留括号内的所有内容")
        print("   • 去掉文件扩展名")
        print("   • 示例：报告 (最终版).pdf → 报告最终版")
        print("   • 示例：2024年 第一季度 报告.pdf → 2024年第一季度报告")
        
    except Exception as e:
        print(f"\n❌ 生成Excel文件时出错：{e}")
        import traceback
        traceback.print_exc()
        return

def main():
    """
    主循环：支持连续执行多次提取任务
    """
    while True:
        try:
            process_multiple_folders()
        except KeyboardInterrupt:
            print("\n\n⚠️  程序被用户中断")
            break
        except Exception as e:
            print(f"\n❌ 发生错误：{e}")
            import traceback
            traceback.print_exc()
        
        print("\n" + "=" * 60)
        again = input("\n是否开始新的提取任务？(y/n，默认n)：").strip().lower()
        if again != 'y':
            print("\n👋 感谢使用，再见！")
            break
        print("\n" * 2)

if __name__ == "__main__":
    # 检查依赖
    try:
        import pandas
        import openpyxl
    except ImportError as e:
        print("❌ 缺少必要的库，请运行以下命令安装：")
        print("pip install pandas openpyxl")
        print("\n安装完成后重新运行程序")
        input("\n按回车键退出...")
        exit()
    
    main()
