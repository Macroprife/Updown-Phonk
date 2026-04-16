import os
from PyPDF2 import PdfReader

def count_pdf_pages(folder_path):
    total_pages = 0
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_path = os.path.join(root, file)
                try:
                    reader = PdfReader(pdf_path)
                    total_pages += len(reader.pages)
                except Exception as e:
                    # 跳过损坏或无法读取的 PDF
                    print(f"警告：无法读取 {pdf_path} - {e}")
    return total_pages

if __name__ == "__main__":
    folder = input("请输入文件夹路径: ").strip()
    if os.path.isdir(folder):
        print(count_pdf_pages(folder))
    else:
        print("路径无效")
