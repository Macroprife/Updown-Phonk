import os

def remove_empty_folders(path):
    """
    简洁版：递归删除所有空文件夹
    """
    for root, dirs, files in os.walk(path, topdown=False):
        for dir in dirs:
            dir_path = os.path.join(root, dir)
            try:
                if not os.listdir(dir_path):
                    os.rmdir(dir_path)
                    print(f"删除: {dir_path}")
            except Exception as e:
                print(f"跳过 {dir_path}: {e}")

# 使用示例
if __name__ == "__main__":
    folder_to_clean = "D:\老鸦山Split" # 修改为你的目标文件夹
    
    # 先确认
    response = input(f"确定要删除 '{folder_to_clean}' 中的所有空子文件夹吗? (y/n): ")
    if response.lower() == 'y':
        remove_empty_folders(folder_to_clean)
        print("清理完成!")
    else:
        print("操作已取消")
