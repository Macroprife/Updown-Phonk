import os
import PyPDF2
from PyPDF2 import PdfReader, PdfWriter

def remove_first_page(input_pdf, output_pdf):
    """删除PDF的第一页，保存为新文件"""
    try:
        reader = PdfReader(input_pdf)
        writer = PdfWriter()

        total_pages = len(reader.pages)

        # 如果总页数 <= 1，删除第一页后没有剩余页面，跳过该文件
        if total_pages <= 1:
            print(f"⚠️ 跳过：{os.path.basename(input_pdf)} (只有{total_pages}页，删除后无内容)")
            return False

        # 从第2页开始到最后（索引从0开始，所以是 1 到 total_pages-1）
        for page_num in range(1, total_pages):
            writer.add_page(reader.pages[page_num])

        with open(output_pdf, "wb") as out_file:
            writer.write(out_file)
        
        print(f"✅ 成功：{os.path.basename(input_pdf)} (原{total_pages}页 → 新{total_pages-1}页)")
        return True
    
    except Exception as e:
        print(f"❌ 错误：{os.path.basename(input_pdf)} - {str(e)}")
        return False


def batch_process_folder():
    """批量处理文件夹内所有PDF文件"""
    
    # 手动输入源文件夹路径
    while True:
        source_folder = input("请输入源文件夹路径: ").strip()
        source_folder = source_folder.strip('"')  # 去除可能的引号
        
        if os.path.exists(source_folder):
            break
        else:
            print("❌ 路径不存在，请重新输入！")
    
    # 手动输入目标文件夹路径
    target_folder = input("请输入目标文件夹路径（输出位置）: ").strip()
    target_folder = target_folder.strip('"')  # 去除可能的引号
    
    # 创建目标文件夹（如果不存在）
    os.makedirs(target_folder, exist_ok=True)
    
    print(f"\n📁 源文件夹: {source_folder}")
    print(f"📁 目标文件夹: {target_folder}")
    print("-" * 50)
    
    # 统计信息
    total_files = 0
    success_count = 0
    skip_count = 0
    error_count = 0
    
    # 遍历文件夹内所有PDF文件
    for filename in os.listdir(source_folder):
        if not filename.lower().endswith(".pdf"):
            continue
        
        total_files += 1
        input_path = os.path.join(source_folder, filename)
        output_path = os.path.join(target_folder, filename)  # 保留原文件名
        
        print(f"\n📄 处理: {filename}")
        success = remove_first_page(input_path, output_path)
        
        if success:
            success_count += 1
        elif success is False:
            # 检查是否是因为页数不足而跳过
            try:
                reader = PdfReader(input_path)
                if len(reader.pages) <= 1:
                    skip_count += 1
                else:
                    error_count += 1
            except:
                error_count += 1
    
    # 输出统计信息
    print("\n" + "=" * 50)
    print("📊 处理完成统计：")
    print(f"   总PDF文件数: {total_files}")
    print(f"   ✅ 成功处理: {success_count}")
    print(f"   ⚠️  跳过（页数≤1）: {skip_count}")
    print(f"   ❌ 处理失败: {error_count}")
    print(f"📁 输出位置: {target_folder}")
    print("=" * 50)


if __name__ == "__main__":
    print("=" * 50)
    print("📚 PDF批量删除第一页工具")
    print("   功能：删除所有PDF的第一页")
    print("   输出：保留原文件名到新文件夹")
    print("=" * 50)
    print()
    
    batch_process_folder()
    
    input("\n按回车键退出...")