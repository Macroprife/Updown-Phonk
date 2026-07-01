import os
from PIL import Image
import gc
import time

def images_to_pdf(input_dir, output_dir, dpi=300):
    """
    将每个子文件夹中的图片转换为对应的PDF文件
    保持每张图片的原始横竖方向
    
    Args:
        input_dir: 输入根目录路径
        output_dir: 输出根目录路径
        dpi: PDF的分辨率，默认300
    """
    # 创建输出目录
    os.makedirs(output_dir, exist_ok=True)
    
    # 支持的图片格式
    image_extensions = {'.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.gif', '.webp'}
    
    # 先统计所有需要处理的文件夹
    all_folders = []
    for root, dirs, files in os.walk(input_dir):
        if root == input_dir:
            continue
        all_folders.append(root)
    
    total_folders = len(all_folders)
    if total_folders == 0:
        print("未找到子文件夹！")
        return
    
    print(f"找到 {total_folders} 个子文件夹需要处理\n")
    
    # 遍历输入目录中的所有子文件夹
    for idx, root in enumerate(all_folders, 1):
        try:
            # 获取相对于输入目录的子文件夹路径
            relative_path = os.path.relpath(root, input_dir)
            folder_name = os.path.basename(root)
            
            # 收集当前子文件夹中的所有图片
            image_files = []
            for file in sorted(os.listdir(root)):  # 排序保证顺序一致
                file_path = os.path.join(root, file)
                if os.path.isfile(file_path):
                    ext = os.path.splitext(file)[1].lower()
                    if ext in image_extensions:
                        image_files.append(file_path)
            
            # 如果没有图片，跳过
            if not image_files:
                print(f"[{idx}/{total_folders}] 跳过空文件夹或无图片文件夹: {relative_path}")
                continue
            
            # 处理多层级子文件夹时的路径问题
            if os.path.dirname(relative_path):
                # 如果有父文件夹层级，使用完整相对路径作为文件名
                pdf_filename = relative_path.replace(os.sep, '_') + '.pdf'
            else:
                pdf_filename = folder_name + '.pdf'
            
            pdf_path = os.path.join(output_dir, pdf_filename)
            
            # 检查PDF是否已存在，避免重复处理
            if os.path.exists(pdf_path):
                print(f"[{idx}/{total_folders}] 跳过已存在: {pdf_filename}")
                continue
            
            print(f"[{idx}/{total_folders}] 处理中: {relative_path} ({len(image_files)} 张图片)")
            
            # 打开所有图片并转换为RGB模式，保持原始方向和尺寸
            images = []
            
            for img_idx, img_path in enumerate(image_files, 1):
                try:
                    img = Image.open(img_path)
                    
                    # 保留EXIF方向信息并自动旋转
                    try:
                        # 获取EXIF数据
                        exif = img._getexif()
                        if exif is not None:
                            orientation = exif.get(0x0112)  # 274是方向标签
                            if orientation is not None:
                                # 根据EXIF方向旋转图片
                                orientation_map = {
                                    2: Image.FLIP_LEFT_RIGHT,
                                    3: 180,
                                    4: Image.FLIP_TOP_BOTTOM,
                                    5: Image.TRANSPOSE,
                                    6: Image.ROTATE_270,
                                    7: Image.TRANSVERSE,
                                    8: Image.ROTATE_90
                                }
                                if orientation in orientation_map:
                                    if isinstance(orientation_map[orientation], int):
                                        img = img.rotate(orientation_map[orientation], expand=True)
                                    else:
                                        img = img.transpose(orientation_map[orientation])
                    except (AttributeError, KeyError, IndexError, TypeError):
                        # 如果没有EXIF数据或读取失败，保持原样
                        pass
                    
                    # 获取原始尺寸
                    orig_width, orig_height = img.size
                    
                    # 优化图片模式转换，减少内存使用
                    if img.mode in ('RGBA', 'LA', 'P'):
                        if img.mode == 'P':
                            img = img.convert('RGBA')
                        # 创建白色背景
                        rgb_img = Image.new('RGB', (orig_width, orig_height), (255, 255, 255))
                        rgb_img.paste(img, mask=img.split()[-1] if img.mode in ('RGBA', 'LA') else None)
                        img.close()  # 关闭原始图片
                        img = rgb_img
                    elif img.mode != 'RGB':
                        old_img = img
                        img = img.convert('RGB')
                        old_img.close()  # 关闭原始图片
                    
                    images.append(img)
                    
                    # 显示进度
                    if img_idx % 5 == 0 or img_idx == len(image_files):
                        print(f"  已加载: {img_idx}/{len(image_files)} 张图片")
                        
                except Exception as e:
                    print(f"  警告: 处理图片 {os.path.basename(img_path)} 时出错: {str(e)}")
                    continue
            
            # 保存为PDF，每张图片保持各自的横竖方向
            if images:
                try:
                    # 第一张图片作为基础，其余图片追加
                    images[0].save(
                        pdf_path,
                        save_all=True,
                        append_images=images[1:] if len(images) > 1 else [],
                        resolution=dpi,
                        quality=85,  # 降低默认质量，减少文件大小和内存使用
                        optimize=True  # 优化PDF文件大小
                    )
                    print(f"  ✓ 已创建: {pdf_filename} (包含 {len(images)} 张图片)")
                except Exception as e:
                    print(f"  ✗ 保存PDF时出错: {str(e)}")
                finally:
                    # 立即释放图片内存
                    for img in images:
                        img.close()
                    images.clear()
                
                # 释放内存
                gc.collect()
            
        except Exception as e:
            print(f"✗ 处理 {relative_path} 时出错: {str(e)}")
            gc.collect()
            continue

def main():
    print("=" * 60)
    print("图片批量转PDF工具（优化版-保持原始横竖方向）")
    print("=" * 60)
    
    # 手动输入输入路径
    while True:
        input_dir = input("\n请输入图片所在的根目录路径: ").strip().strip('"').strip("'")
        if os.path.exists(input_dir):
            break
        print(f"错误: 路径 '{input_dir}' 不存在，请重新输入!")
    
    # 手动输入输出路径
    while True:
        output_dir = input("请输入PDF输出目录路径: ").strip().strip('"').strip("'")
        if output_dir:
            break
        print("错误: 输出路径不能为空，请重新输入!")
    
    # 手动输入DPI
    while True:
        dpi_input = input("请输入DPI（直接回车默认为300）: ").strip()
        if not dpi_input:
            dpi = 300
            break
        try:
            dpi = int(dpi_input)
            if dpi > 0 and dpi <= 1200:  # 限制最大DPI
                break
            elif dpi > 1200:
                print("警告: DPI过大可能导致内存不足，建议不超过1200")
                confirm = input("是否继续使用此DPI? (y/n): ").strip().lower()
                if confirm == 'y':
                    break
            else:
                print("错误: DPI必须大于0，请重新输入!")
        except ValueError:
            print("错误: 请输入有效的数字!")
    
    print("\n" + "-" * 60)
    print(f"输入目录: {input_dir}")
    print(f"输出目录: {output_dir}")
    print(f"DPI: {dpi}")
    print("-" * 60 + "\n")
    
    start_time = time.time()
    images_to_pdf(input_dir, output_dir, dpi)
    
    elapsed_time = time.time() - start_time
    print("\n" + "-" * 60)
    print(f"转换完成! 耗时: {elapsed_time:.1f} 秒")
    print("-" * 60)
    
    input("\n按回车键退出...")

if __name__ == "__main__":
    main()
