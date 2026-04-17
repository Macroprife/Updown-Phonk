import os

def find_empty_folders(directory):
    """查找目录中的所有空文件夹"""
    empty_folders = []
    
    for root, dirs, files in os.walk(directory):
        # 检查当前文件夹是否为空（没有文件和子文件夹）
        if not dirs and not files:
            empty_folders.append(root)
    
    return empty_folders

# 主程序
if __name__ == "__main__":
    # 手动输入文件夹路径
    folder_path = input("请输入要扫描的文件夹路径: ").strip()
    
    # 去除可能的引号
    folder_path = folder_path.strip('"').strip("'")
    
    # 检查路径是否存在
    if not os.path.exists(folder_path):
        print(f"错误：路径 '{folder_path}' 不存在！")
    elif not os.path.isdir(folder_path):
        print(f"错误：'{folder_path}' 不是一个文件夹！")
    else:
        print(f"\n正在扫描文件夹: {folder_path}")
        print("-" * 50)
        
        empty_dirs = find_empty_folders(folder_path)
        
        if empty_dirs:
            print(f"\n找到 {len(empty_dirs)} 个空文件夹：")
            for i, folder in enumerate(empty_dirs, 1):
                print(f"{i}. {folder}")
        else:
            print("\n没有找到空文件夹")
