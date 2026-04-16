import os
import glob
import shutil

def rename_parent_folders_by_pdf(root_dir):
    """
    遍历根目录下的所有子文件夹，根据其中的PDF文件名重命名父文件夹
    
    Args:
        root_dir: 根目录路径
    """
    # 遍历根目录下的所有子文件夹
    for foldername, subfolders, filenames in os.walk(root_dir):
        # 在当前文件夹中查找PDF文件
        pdf_files = [f for f in filenames if f.lower().endswith('.pdf')]
        
        # 如果没有PDF文件，跳过
        if not pdf_files:
            continue
        
        # 如果有多个PDF文件，使用第一个PDF文件
        pdf_name = pdf_files[0]
        
        # 获取PDF文件名（不含扩展名）
        pdf_name_without_ext = os.path.splitext(pdf_name)[0]
        
        # 获取当前文件夹的父文件夹路径
        parent_dir = os.path.dirname(foldername)
        
        # 如果当前文件夹就是根目录，跳过（我们不希望重命名根目录）
        if parent_dir == root_dir:
            print(f"跳过根目录下的直接子文件夹: {foldername}")
            continue
        
        # 获取当前文件夹名
        current_folder_name = os.path.basename(foldername)
        
        # 如果当前文件夹名已经和PDF文件名相同（忽略扩展名），跳过
        if current_folder_name == pdf_name_without_ext:
            print(f"文件夹名已匹配，跳过: {foldername}")
            continue
        
        # 构建新的文件夹路径
        new_folder_path = os.path.join(parent_dir, pdf_name_without_ext)
        
        # 检查新文件夹名是否已存在（避免冲突）
        if os.path.exists(new_folder_path) and new_folder_path != foldername:
            print(f"目标文件夹已存在，跳过重命名 {foldername} -> {pdf_name_without_ext}")
            continue
        
        try:
            # 重命名文件夹
            shutil.move(foldername, new_folder_path)
            print(f"成功重命名: {current_folder_name} -> {pdf_name_without_ext}")
        except Exception as e:
            print(f"重命名失败 {foldername} -> {pdf_name_without_ext}: {str(e)}")

def main():
    # 设置要遍历的根目录
    root_directory = input("请输入要处理的根目录路径: ").strip()
    
    # 检查目录是否存在
    if not os.path.exists(root_directory):
        print(f"错误: 目录 '{root_directory}' 不存在")
        return
    
    # 询问用户是否确认执行
    print(f"\n将要处理目录: {root_directory}")
    print("此操作将根据PDF文件名重命名其所在的父文件夹")
    confirmation = input("是否继续? (y/n): ").strip().lower()
    
    if confirmation != 'y':
        print("操作已取消")
        return
    
    # 执行重命名操作
    rename_parent_folders_by_pdf(root_directory)
    print("\n处理完成!")

if __name__ == "__main__":
    main()
