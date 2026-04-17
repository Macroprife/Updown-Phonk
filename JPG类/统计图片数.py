import os
import pandas as pd
from datetime import datetime
from pathlib import Path

# 支持的图片格式
IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.tif', 
                    '.webp', '.svg', '.ico', '.heic', '.heif', '.raw', '.cr2'}

def count_images_in_subfolders(root_path=None):
    # 手动输入路径
    if root_path is None:
        root_path = input("请输入文件夹路径: ").strip().strip('"').strip("'")
    
    # 检查文件夹是否存在
    if not os.path.exists(root_path):
        print(f"错误：路径 '{root_path}' 不存在！")
        return
    
    if not os.path.isdir(root_path):
        print(f"错误：'{root_path}' 不是一个文件夹！")
        return
    
    print(f"\n正在扫描文件夹: {root_path}")
    print("请稍候...\n")
    
    # 存储统计数据
    data = []
    total_images = 0
    total_folders = 0
    
    # 遍历所有子文件夹
    for foldername, subfolders, filenames in os.walk(root_path):
        # 计算当前文件夹中的图片数量
        image_files = [f for f in filenames 
                      if Path(f).suffix.lower() in IMAGE_EXTENSIONS]
        image_count = len(image_files)
        
        if image_count > 0:
            total_folders += 1
            total_images += image_count
            
            # 计算相对路径
            rel_path = os.path.relpath(foldername, root_path)
            if rel_path == '.':
                rel_path = '根目录'
            
            # 获取文件夹大小
            folder_size = sum(os.path.getsize(os.path.join(foldername, f)) 
                            for f in image_files) / (1024 * 1024)  # MB
            
            data.append({
                '文件夹路径': rel_path,
                '完整路径': foldername,
                '图片数量': image_count,
                '图片总大小(MB)': round(folder_size, 2)
            })
            
            print(f"📁 {rel_path:<50} {image_count:>6} 张图片")
    
    # 按图片数量降序排序
    data.sort(key=lambda x: x['图片数量'], reverse=True)
    
    # 添加汇总行
    data.append({
        '文件夹路径': '【总计】',
        '完整路径': f'共扫描 {total_folders} 个包含图片的文件夹',
        '图片数量': total_images,
        '图片总大小(MB)': ''
    })
    
    # 创建DataFrame
    df = pd.DataFrame(data)
    
    # 生成Excel文件名
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    folder_name = os.path.basename(root_path)
    excel_filename = f"图片统计_{folder_name}_{timestamp}.xlsx"
    excel_path = os.path.join(os.getcwd(), excel_filename)
    
    # 保存到Excel
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='图片统计', index=False)
        
        # 获取工作表对象
        worksheet = writer.sheets['图片统计']
        
        # 调整列宽
        worksheet.column_dimensions['A'].width = 40  # 文件夹路径
        worksheet.column_dimensions['B'].width = 60  # 完整路径
        worksheet.column_dimensions['C'].width = 15  # 图片数量
        worksheet.column_dimensions['D'].width = 18  # 图片大小
        
        # 设置表头样式
        from openpyxl.styles import Font, PatternFill, Alignment
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_alignment = Alignment(horizontal="center", vertical="center")
        
        for cell in worksheet[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
    
    # 打印统计结果
    print("\n" + "="*80)
    print(f"📊 统计完成！")
    print(f"📁 根目录: {root_path}")
    print(f"📂 包含图片的文件夹数: {total_folders} 个")
    print(f"🖼️  图片总数: {total_images:,} 张")
    print(f"💾 Excel文件已保存: {excel_filename}")
    print("="*80)
    
    # 显示TOP 5文件夹
    if len(data) > 1:
        print("\n🏆 图片数量 TOP 5 文件夹:")
        for i, item in enumerate(data[:5], 1):
            if item['文件夹路径'] != '【总计】':
                print(f"   {i}. {item['文件夹路径']}: {item['图片数量']} 张")
    
    # 询问是否打开Excel文件
    choice = input("\n是否打开Excel文件？(y/n): ").strip().lower()
    if choice == 'y':
        os.startfile(excel_path)
    
    # 询问是否继续统计其他文件夹
    print("\n" + "="*80)
    choice = input("是否继续统计其他文件夹？(y/n): ").strip().lower()
    if choice == 'y':
        count_images_in_subfolders()

# 直接运行
if __name__ == "__main__":
    print("="*80)
    print("                    图片统计工具（按子文件夹统计）")
    print("="*80)
    print(f"支持的图片格式: {', '.join(sorted(IMAGE_EXTENSIONS))}")
    print("-"*80)
    count_images_in_subfolders()
