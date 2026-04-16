import os
import re
import shutil

def rename_files_in_directory(root_dir):
    """
    遍历指定目录及其子目录，按要求重命名文件
    """
    # 遍历所有子目录和文件
    for dirpath, dirnames, filenames in os.walk(root_dir):
        for filename in filenames:
            # 只处理PDF文件
            if not filename.lower().endswith('.pdf'):
                continue
            
            # 获取文件完整路径
            old_path = os.path.join(dirpath, filename)
            
            # 去除文件扩展名
            name_without_ext = os.path.splitext(filename)[0]
            
            # 定义匹配模式：
            # 1. 匹配类似 "230-QQ0311-0001-002" 的模式
            # 2. 匹配类似 "230-QQ0311-0001-017_剩余" 的模式
            pattern1 = r'^(\d{3}-QQ\d{4}-\d{4}-)(\d{3})$'  # 标准格式
            pattern2 = r'^(\d{3}-QQ\d{4}-\d{4}-)(\d{3})_剩余$'  # 带"_剩余"的格式
            
            match1 = re.match(pattern1, name_without_ext)
            match2 = re.match(pattern2, name_without_ext)
            
            new_name = None
            
            if match2:
                # 情况2：去除"_剩余"
                prefix = match2.group(1)
                number = match2.group(2)
                new_name = f"{prefix}{number}.pdf"
                print(f"去除'_剩余': {filename} -> {new_name}")
                
            elif match1:
                # 情况1：末三位数字减1
                prefix = match1.group(1)
                number_str = match1.group(2)
                
                try:
                    # 将数字转换为整数并减1
                    number_int = int(number_str)
                    new_number = number_int - 1
                    
                    # 格式化为3位数字，前面补零
                    if new_number >= 0:
                        new_number_str = f"{new_number:03d}"
                        new_name = f"{prefix}{new_number_str}.pdf"
                        print(f"数字减1: {filename} -> {new_name}")
                    else:
                        print(f"警告: 数字减1后为负数，跳过文件: {filename}")
                        continue
                except ValueError:
                    print(f"警告: 无法解析数字部分，跳过文件: {filename}")
                    continue
            
            if new_name:
                # 构建新的文件路径
                new_path = os.path.join(dirpath, new_name)
                
                try:
                    # 重命名文件
                    shutil.move(old_path, new_path)
                    print(f"成功重命名: {old_path} -> {new_path}")
                except Exception as e:
                    print(f"重命名失败: {old_path} -> {new_path}, 错误: {e}")

def main():
    # 设置要处理的根目录
    root_directory = "D:\老鸦山Split"
    
    if not os.path.exists(root_directory):
        print(f"错误: 目录不存在: {root_directory}")
        return
    
    print(f"开始处理目录: {root_directory}")
    print("-" * 50)
    
    rename_files_in_directory(root_directory)
    
    print("-" * 50)
    print("处理完成！")

if __name__ == "__main__":
    main()
