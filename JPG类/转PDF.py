import os
from PIL import Image

def images_to_pdf(input_path, output_path=None, dpi=150):
    """
    将文件夹中的图片合并为PDF
    """
    image_extensions = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp'}
    
    folder_name = os.path.basename(input_path.rstrip(os.sep))
    
    if output_path is None:
        output_path = os.path.join(os.path.dirname(input_path), f"{folder_name}.pdf")
    
    # 收集所有图片文件并排序
    images = []
    print(f"正在扫描文件夹: {input_path}")
    for file in sorted(os.listdir(input_path)):
        if os.path.splitext(file)[1].lower() in image_extensions:
            file_path = os.path.join(input_path, file)
            try:
                img = Image.open(file_path)
                if img.mode in ('RGBA', 'LA', 'P'):
                    img = img.convert('RGB')
                images.append(img)
                print(f"  ✓ 已加载: {file}")
            except Exception as e:
                print(f"  ✗ 跳过 {file}: {e}")
    
    if not images:
        print(f"⚠ 警告: {folder_name} 中没有找到图片文件")
        return False
    
    # 保存为PDF
    images[0].save(
        output_path,
        "PDF",
        save_all=True,
        append_images=images[1:],
        resolution=dpi,
        quality=95
    )
    print(f"✓ PDF已创建: {output_path} (包含{len(images)}张图片, DPI={dpi})\n")
    return True

def batch_convert(root_folder, dpi=150):
    """
    批量转换root_folder下所有子文件夹的图片为PDF
    """
    if not os.path.exists(root_folder):
        print(f"错误: 路径不存在 - {root_folder}")
        return
    
    success_count = 0
    fail_count = 0
    folder_list = []
    
    # 先列出所有子文件夹
    for item in os.listdir(root_folder):
        item_path = os.path.join(root_folder, item)
        if os.path.isdir(item_path):
            folder_list.append(item)
    
    if not folder_list:
        print(f"在 {root_folder} 中没有找到子文件夹")
        return
    
    print(f"\n找到 {len(folder_list)} 个文件夹:")
    for i, folder in enumerate(folder_list, 1):
        print(f"  {i}. {folder}")
    
    print(f"\n开始转换 (DPI={dpi})...")
    print("="*50)
    
    for folder_name in folder_list:
        folder_path = os.path.join(root_folder, folder_name)
        print(f"\n处理文件夹 [{folder_name}]")
        if images_to_pdf(folder_path, dpi=dpi):
            success_count += 1
        else:
            fail_count += 1
    
    print("="*50)
    print(f"\n转换完成！")
    print(f"  成功: {success_count} 个")
    print(f"  失败: {fail_count} 个")

def main():
    print("="*50)
    print("图片文件夹批量转PDF工具")
    print("="*50)
    
    while True:
        # 手动输入路径
        root_folder = input("\n请输入包含图片文件夹的根目录路径: ").strip()
        
        # 去除可能的引号
        root_folder = root_folder.strip('"').strip("'")
        
        if not root_folder:
            print("路径不能为空，请重新输入")
            continue
        
        if not os.path.exists(root_folder):
            print(f"路径不存在: {root_folder}")
            retry = input("是否重新输入？(y/n): ").lower()
            if retry != 'y':
                return
            continue
        
        if not os.path.isdir(root_folder):
            print(f"这不是一个文件夹: {root_folder}")
            retry = input("是否重新输入？(y/n): ").lower()
            if retry != 'y':
                return
            continue
        
        break
    
    # 输入DPI
    while True:
        dpi_input = input("\n请输入DPI值 (默认150，直接回车使用默认值): ").strip()
        
        if not dpi_input:
            dpi = 150
            break
        
        try:
            dpi = int(dpi_input)
            if dpi <= 0:
                print("DPI必须大于0")
                continue
            if dpi > 1200:
                confirm = input("DPI值较大(>1200)，生成的文件可能会很大，确认继续？(y/n): ")
                if confirm.lower() != 'y':
                    continue
            break
        except ValueError:
            print("请输入有效的数字")
    
    # 确认信息
    print(f"\n配置信息:")
    print(f"  根目录: {root_folder}")
    print(f"  输出DPI: {dpi}")
    confirm = input("\n确认开始转换？(y/n): ").lower()
    
    if confirm == 'y':
        batch_convert(root_folder, dpi)
    else:
        print("已取消")
    
    input("\n按回车键退出...")

if __name__ == "__main__":
    main()
