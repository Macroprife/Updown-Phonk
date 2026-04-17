import os
import PyPDF2
import pandas as pd
from datetime import datetime

def count_pdf_pages_to_excel(folder_path=None):
    # 如果没有传入路径，则让用户手动输入
    if folder_path is None:
        folder_path = input("请输入文件夹路径: ").strip()
        folder_path = folder_path.strip('"').strip("'")
    
    # 检查文件夹是否存在
    if not os.path.exists(folder_path):
        print(f"错误：路径 '{folder_path}' 不存在！")
        return
    
    if not os.path.isdir(folder_path):
        print(f"错误：'{folder_path}' 不是一个文件夹！")
        return
    
    # 获取所有PDF文件
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        print("文件夹中没有PDF文件")
        return
    
    # 存储结果的数据列表
    data = []
    total_pages = 0
    
    print(f"\n正在统计PDF页数...")
    print("-" * 62)
    
    for pdf_file in sorted(pdf_files):
        try:
            file_path = os.path.join(folder_path, pdf_file)
            file_size = os.path.getsize(file_path) / (1024 * 1024)  # 转换为MB
            
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                pages = len(pdf_reader.pages)
                total_pages += pages
                
                data.append({
                    '文件名': pdf_file,
                    '页数': pages,
                    '文件大小(MB)': round(file_size, 2),
                    '完整路径': file_path
                })
                
                print(f"✓ {pdf_file:<40} {pages:>6} 页")
                
        except Exception as e:
            data.append({
                '文件名': pdf_file,
                '页数': '读取失败',
                '文件大小(MB)': round(file_size, 2),
                '完整路径': file_path
            })
            print(f"✗ {pdf_file:<40} 读取失败")
    
    # 添加汇总行
    data.append({
        '文件名': '【总计】',
        '页数': total_pages,
        '文件大小(MB)': '',
        '完整路径': f'共 {len(pdf_files)} 个PDF文件'
    })
    
    # 创建DataFrame
    df = pd.DataFrame(data)
    
    # 生成Excel文件名（包含时间戳）
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    folder_name = os.path.basename(folder_path)
    excel_filename = f"PDF页数统计_{folder_name}_{timestamp}.xlsx"
    
    # 保存到Excel
    excel_path = os.path.join(os.getcwd(), excel_filename)
    
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='PDF页数统计', index=False)
        
        # 获取工作表对象以调整格式
        worksheet = writer.sheets['PDF页数统计']
        
        # 调整列宽
        worksheet.column_dimensions['A'].width = 40  # 文件名
        worksheet.column_dimensions['B'].width = 12  # 页数
        worksheet.column_dimensions['C'].width = 15  # 文件大小
        worksheet.column_dimensions['D'].width = 50  # 完整路径
    
    print("-" * 62)
    print(f"\n📊 统计完成！")
    print(f"📁 文件夹: {folder_path}")
    print(f"📄 PDF文件数: {len(pdf_files)} 个")
    print(f"📑 总页数: {total_pages} 页")
    print(f"💾 Excel文件已保存: {excel_filename}")
    
    # 询问是否打开Excel文件
    choice = input("\n是否打开Excel文件？(y/n): ").strip().lower()
    if choice == 'y':
        os.startfile(excel_path)
    
    # 询问是否继续统计其他文件夹
    print("\n" + "="*62)
    choice = input("是否继续统计其他文件夹？(y/n): ").strip().lower()
    if choice == 'y':
        count_pdf_pages_to_excel()

# 直接运行
if __name__ == "__main__":
    print("="*62)
    print("           PDF 页数统计工具（带Excel导出）")
    print("="*62)
    count_pdf_pages_to_excel()
