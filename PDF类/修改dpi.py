import fitz
from PIL import Image
import os
from pathlib import Path

def process_brute_force(in_path, out_path):
    """暴力重采样模式"""
    doc = fitz.open(in_path)
    images = []
    zoom = 300 / 72
    matrix = fitz.Matrix(zoom, zoom)
    for page in doc:
        pix = page.get_pixmap(matrix=matrix)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
    
    images[0].save(out_path, "PDF", resolution=300.0, save_all=True, 
                   append_images=images[1:], optimize=True, quality=80)
    doc.close()

def process_metadata(in_path, out_path):
    """属性微调模式"""
    doc = fitz.open(in_path)
    factor = 299 / 300
    for page in doc:
        r = page.rect
        page.set_mediabox(fitz.Rect(r.x0, r.y0, r.x1 * factor, r.y1 * factor))
    doc.save(out_path, garbage=3, deflate=True)
    doc.close()

def batch_processor(mode):
    in_root = input("请输入【输入文件夹】根路径: ").strip('"').strip("'")
    out_root = input("请输入【输出文件夹】根路径: ").strip('"').strip("'")
    
    if not os.path.exists(in_root):
        print("错误：输入路径不存在！")
        return

    # 递归遍历目录
    for root, dirs, files in os.walk(in_root):
        for file in files:
            if file.lower().endswith(".pdf"):
                # 获取文件的完整路径
                in_file_path = os.path.join(root, file)
                
                # 计算相对路径，以便在输出文件夹中重建相同的目录结构
                rel_path = os.path.relpath(root, in_root)
                out_dir = os.path.join(out_root, rel_path)
                os.makedirs(out_dir, exist_ok=True)
                
                out_file_path = os.path.join(out_dir, file)
                
                print(f"正在处理: {file} ...")
                try:
                    if mode == '1':
                        process_brute_force(in_file_path, out_file_path)
                    else:
                        process_metadata(in_file_path, out_file_path)
                    print(f"  -> 成功: {out_file_path}")
                except Exception as e:
                    print(f"  -> 失败: {file}, 错误: {e}")

if __name__ == "__main__":
    while True:
        print("\n" + "="*50)
        print("      PDF 批量 DPI 修改工具")
        print("="*50)
        print("1. 暴力重采样 (300DPI 稳过审，生成新文件)")
        print("2. 属性微调 (轻量修改，通过率视情况)")
        print("0. 退出")
        
        choice = input("\n请选择模式 (1/2/0): ")
        if choice in ['1', '2']:
            batch_processor(choice)
            print("\n所有任务处理完毕！")
        elif choice == '0':
            break
        else:
            print("无效输入，请重试。")
