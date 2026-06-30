import os
import shutil

def get_pdf_filenames(folder_path):
    """获取文件夹中所有PDF文件的文件名（不含扩展名）"""
    pdf_files = set()
    for file in os.listdir(folder_path):
        if file.lower().endswith('.pdf'):
            pdf_files.add(os.path.splitext(file)[0])  # 只保存文件名，不含.pdf
    return pdf_files

def copy_unmatched_pdfs(folder1, folder2, output_folder):
    """比较两个文件夹，复制没有匹配的PDF文件到新文件夹"""
    
    # 获取两个文件夹中的PDF文件名（不含扩展名）
    files1 = get_pdf_filenames(folder1)
    files2 = get_pdf_filenames(folder2)
    
    # 找出在folder1中但不在folder2中的文件
    unmatched_in_folder1 = files1 - files2
    # 找出在folder2中但不在folder1中的文件
    unmatched_in_folder2 = files2 - files1
    
    # 确保输出文件夹存在
    os.makedirs(output_folder, exist_ok=True)
    
    # 复制folder1中不匹配的文件
    for filename in unmatched_in_folder1:
        src = os.path.join(folder1, filename + '.pdf')
        dst = os.path.join(output_folder, filename + '.pdf')
        shutil.copy2(src, dst)
        print(f"已复制: {filename}.pdf (来自文件夹1)")
    
    # 复制folder2中不匹配的文件
    for filename in unmatched_in_folder2:
        src = os.path.join(folder2, filename + '.pdf')
        dst = os.path.join(output_folder, filename + '.pdf')
        shutil.copy2(src, dst)
        print(f"已复制: {filename}.pdf (来自文件夹2)")
    
    print(f"\n完成！共复制 {len(unmatched_in_folder1) + len(unmatched_in_folder2)} 个不匹配的PDF文件")
    print(f"来自文件夹1的不匹配文件: {len(unmatched_in_folder1)} 个")
    print(f"来自文件夹2的不匹配文件: {len(unmatched_in_folder2)} 个")

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
