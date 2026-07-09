import os
import shutil

def get_folder_name_from_path(path):
    """从路径中提取最后一级目录名"""
    # 去除路径末尾的分隔符
    path = path.rstrip(os.sep)
    # 获取最后一级目录名
    return os.path.basename(path)

def get_pdf_filenames(folder_path):
    """获取文件夹中所有PDF文件的文件名（不含扩展名）"""
    pdf_files = set()
    for file in os.listdir(folder_path):
        if file.lower().endswith('.pdf'):
            pdf_files.add(os.path.splitext(file)[0])  # 只保存文件名，不含.pdf
    return pdf_files

def copy_unmatched_pdfs(folder1, folder2, output_folder):
    """比较两个文件夹，复制没有匹配的PDF文件到新文件夹"""
    
    # 获取两个文件夹的目录名
    folder1_name = get_folder_name_from_path(folder1)
    folder2_name = get_folder_name_from_path(folder2)
    
    # 获取两个文件夹中的PDF文件名（不含扩展名）
    files1 = get_pdf_filenames(folder1)
    files2 = get_pdf_filenames(folder2)
    
    # 找出在folder1中但不在folder2中的文件
    unmatched_in_folder1 = files1 - files2
    # 找出在folder2中但不在folder1中的文件
    unmatched_in_folder2 = files2 - files1
    
    # 创建输出文件夹和两个子文件夹（使用原始文件夹名）
    os.makedirs(output_folder, exist_ok=True)
    subfolder1 = os.path.join(output_folder, folder1_name)
    subfolder2 = os.path.join(output_folder, folder2_name)
    os.makedirs(subfolder1, exist_ok=True)
    os.makedirs(subfolder2, exist_ok=True)
    
    # 复制folder1中不匹配的文件到子文件夹1
    for filename in unmatched_in_folder1:
        src = os.path.join(folder1, filename + '.pdf')
        dst = os.path.join(subfolder1, filename + '.pdf')
        shutil.copy2(src, dst)
        print(f"已复制: {filename}.pdf -> {folder1_name}")
    
    # 复制folder2中不匹配的文件到子文件夹2
    for filename in unmatched_in_folder2:
        src = os.path.join(folder2, filename + '.pdf')
        dst = os.path.join(subfolder2, filename + '.pdf')
        shutil.copy2(src, dst)
        print(f"已复制: {filename}.pdf -> {folder2_name}")
    
    print(f"\n完成！共复制 {len(unmatched_in_folder1) + len(unmatched_in_folder2)} 个不匹配的PDF文件")
    print(f"来自 '{folder1_name}' 的不匹配文件: {len(unmatched_in_folder1)} 个 -> 保存在 '{folder1_name}' 子文件夹")
    print(f"来自 '{folder2_name}' 的不匹配文件: {len(unmatched_in_folder2)} 个 -> 保存在 '{folder2_name}' 子文件夹")
    print(f"\n输出目录结构:")
    print(f"{output_folder}/")
    print(f"  ├── {folder1_name}/ ({len(unmatched_in_folder1)} 个文件)")
    print(f"  └── {folder2_name}/ ({len(unmatched_in_folder2)} 个文件)")

# 主程序
if __name__ == "__main__":
    print("=== PDF文件比较工具 ===")
    
    # 手动输入三个路径
    folder1 = input("请输入第一个文件夹路径: ").strip()
    folder2 = input("请输入第二个文件夹路径: ").strip()
    output_folder = input("请输入输出文件夹路径: ").strip()
    
    # 验证路径是否存在
    if not os.path.exists(folder1):
        print(f"错误：路径不存在 - {folder1}")
        exit(1)
    if not os.path.exists(folder2):
        print(f"错误：路径不存在 - {folder2}")
        exit(1)
    
    # 执行比较和复制操作
    try:
        copy_unmatched_pdfs(folder1, folder2, output_folder)
    except Exception as e:
        print(f"执行过程中出现错误: {e}")
    input("\n按回车键退出...")
