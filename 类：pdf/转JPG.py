import os
import fitz  # PyMuPDF
from PIL import Image
import io
from datetime import datetime
import logging
import sys

def setup_logging():
    """设置简单的控制台输出（不生成日志文件）"""
    # 只使用控制台输出，不生成日志文件
    # 创建一个简单的logger，只输出到控制台
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)
    
    # 清除已有的处理器
    logger.handlers.clear()
    
    # 添加控制台处理器
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)
    
    return logger

# 重新配置logging，不生成文件
logging.basicConfig(handlers=[logging.StreamHandler()], level=logging.INFO)

def verify_conversion(pdf_path, output_folder, expected_pages):
    
    try:
        # 获取输出文件夹中所有图片文件（支持.jpg和.jpeg）
        image_extensions = ['.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.gif']
        image_files = [f for f in os.listdir(output_folder) 
                      if os.path.splitext(f)[1].lower() in image_extensions]
        
        actual_pages = len(image_files)
        
        if actual_pages == expected_pages:
            return True, actual_pages, "转换成功，页数一致"
        else:
            error_msg = f"页数不一致！PDF有{expected_pages}页，但只生成{actual_pages}张图片"
            return False, actual_pages, error_msg
            
    except Exception as e:
        return False, 0, f"验证过程中出错: {str(e)}"

def sanitize_filename(filename):
  
    # Windows文件名中不允许的字符
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    # 移除前后空格
    filename = filename.strip()
    # 如果文件名为空，返回默认名称
    if not filename:
        filename = "image"
    return filename

def get_image_prefix(pdf_path, max_length=15):
   
    # 获取PDF文件名（不含扩展名）
    pdf_filename = os.path.splitext(os.path.basename(pdf_path))[0]
    
    # 清理文件名中的不合法字符
    clean_name = sanitize_filename(pdf_filename)
    
    # 限制长度（前15个字符）
    if len(clean_name) > max_length:
        clean_name = clean_name[:max_length]
        print(f"信息: 文件名 '{pdf_filename}' 超过{max_length}字符，截取前{max_length}字符: '{clean_name}'")
    
    return clean_name

def save_image_with_dpi(pix, image_path, format='PNG', dpi=300):

    if format.upper() == 'PNG':
        # 对于PNG，使用PIL保存并设置DPI
        img_data = pix.tobytes("png")
        img = Image.open(io.BytesIO(img_data))
        
        # 设置DPI信息（PNG使用dpi参数）
        img.save(image_path, 'PNG', dpi=(dpi, dpi))
        
    elif format.upper() in ['JPEG', 'JPG']:
        # 对于JPEG，使用PIL保存并设置DPI，统一保存为.jpg扩展名
        img_data = pix.tobytes("jpeg")
        img = Image.open(io.BytesIO(img_data))
        
        # 确保文件扩展名为.jpg
        if image_path.lower().endswith('.jpeg'):
            image_path = image_path[:-5] + '.jpg'
        
        # 设置DPI信息
        img.save(image_path, 'JPEG', quality=95, dpi=(dpi, dpi))
        
    elif format.upper() == 'TIFF':
        # 对于TIFF
        img_data = pix.tobytes("png")
        img = Image.open(io.BytesIO(img_data))
        img.save(image_path, 'TIFF', dpi=(dpi, dpi))
        
    else:
        # 其他格式（BMP等）直接保存，但可能没有DPI信息
        pix.save(image_path)

def convert_pdf_to_images(pdf_path, output_folder, dpi=150, format='PNG'):
  
    result = {
        'pdf_path': pdf_path,
        'output_folder': output_folder,
        'expected_pages': 0,
        'actual_pages': 0,
        'success': False,
        'error': None,
        'verification_passed': False,
        'verification_message': ''
    }
    
    try:
        # 确保输出文件夹存在
        os.makedirs(output_folder, exist_ok=True)
        
        # 打开PDF文件
        pdf_document = fitz.open(pdf_path)
        result['expected_pages'] = len(pdf_document)
        
        print(f"开始转换: {os.path.basename(pdf_path)} (共{result['expected_pages']}页)")
        
        # 获取图片前缀（PDF文件名前15个字符）
        image_prefix = get_image_prefix(pdf_path, max_length=15)
        print(f"使用图片前缀: {image_prefix} (基于PDF文件名前15个字符)")
        print(f"目标分辨率: {dpi} DPI")
        
        # 统一处理格式名称
        display_format = format.upper()
        if display_format in ['JPEG', 'JPG']:
            file_extension = 'jpg'  # 统一使用.jpg扩展名
            display_format = 'JPEG'  # 内部仍使用JPEG格式
        else:
            file_extension = format.lower()
        
        # 转换每一页
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            
            # 设置缩放因子以获得高分辨率
            zoom = dpi / 72
            mat = fitz.Matrix(zoom, zoom)
            
            # 渲染页面为图片
            pix = page.get_pixmap(matrix=mat)
            
            # 生成图片文件名：前缀_001，前缀_002，...
            image_number = str(page_num + 1).zfill(3)  # 使用3位数字，如001, 002, 010, 100
            image_name = f"{image_prefix}_{image_number}.{file_extension}"
            image_path = os.path.join(output_folder, image_name)
            
            # 保存图片并设置DPI信息
            save_image_with_dpi(pix, image_path, format, dpi)
            
            print(f"  已保存: {image_name} (DPI: {dpi})")
        
        # 关闭PDF文档
        pdf_document.close()
        
        # 验证转换结果
        verification_result = verify_conversion(pdf_path, output_folder, result['expected_pages'])
        result['actual_pages'] = verification_result[1]
        result['verification_passed'] = verification_result[0]
        result['verification_message'] = verification_result[2]
        
        if result['verification_passed']:
            result['success'] = True
            print(f"✓ 转换完成: {os.path.basename(pdf_path)} -> {output_folder}")
            print(f"  验证结果: {result['verification_message']}")
        else:
            result['error'] = result['verification_message']
            print(f"✗ 转换失败: {os.path.basename(pdf_path)} - {result['verification_message']}")
        
        return result
        
    except Exception as e:
        result['error'] = str(e)
        print(f"转换过程中出错 ({os.path.basename(pdf_path)}): {str(e)}")
        return result

def batch_convert_pdfs(input_folder, output_base_folder, dpi=150, format='PNG'):
    """
    批量转换文件夹中的所有PDF文件
    
    Args:
        input_folder: 包含PDF文件的输入文件夹
        output_base_folder: 输出基础文件夹
        dpi: 图片分辨率
        format: 图片格式
        
    Returns:
        list: 所有转换结果的列表
    """
    all_results = []
    
    # 确保输出基础文件夹存在
    os.makedirs(output_base_folder, exist_ok=True)
    
    # 查找所有PDF文件
    pdf_files = []
    for root, dirs, files in os.walk(input_folder):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(root, file))
    
    total_pdfs = len(pdf_files)
    print(f"找到 {total_pdfs} 个PDF文件需要转换")
    
    if total_pdfs == 0:
        print("未找到任何PDF文件！")
        return all_results
    
    # 转换每个PDF文件
    for i, pdf_path in enumerate(pdf_files, 1):
        print(f"\n处理文件 {i}/{total_pdfs}: {os.path.basename(pdf_path)}")
        
        # 为每个PDF创建输出文件夹（使用PDF文件名，不含扩展名）
        pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_folder = os.path.join(output_base_folder, pdf_name)
        
        # 执行转换
        result = convert_pdf_to_images(pdf_path, output_folder, dpi, format)
        all_results.append(result)
    
    # 输出汇总信息
    successful = sum(1 for r in all_results if r['success'])
    failed = len(all_results) - successful
    
    print("\n" + "=" * 60)
    print("批量转换完成汇总")
    print("=" * 60)
    print(f"总PDF文件数: {len(all_results)}")
    print(f"成功转换: {successful}")
    print(f"转换失败: {failed}")
    
    if failed > 0:
        print("\n失败的文件列表:")
        for r in all_results:
            if not r['success']:
                print(f"  - {os.path.basename(r['pdf_path'])}: {r['error']}")
    
    print("=" * 60)
    
    return all_results

def convert_single_pdf(pdf_path, output_folder, dpi=150, format='PNG'):
    """
    转换单个PDF文件
    
    Args:
        pdf_path: PDF文件路径
        output_folder: 输出文件夹路径
        dpi: 图片分辨率
        format: 图片格式
        
    Returns:
        dict: 转换结果
    """
    print(f"开始转换单个PDF文件: {pdf_path}")
    
    # 确保输出文件夹存在
    os.makedirs(output_folder, exist_ok=True)
    
    # 执行转换
    result = convert_pdf_to_images(pdf_path, output_folder, dpi, format)
    
    return result

def get_user_input():
    """获取用户输入的路径和参数"""
    print("\n" + "=" * 60)
    print("PDF转图片工具")
    print("=" * 60)
    
    # 选择转换模式
    print("\n请选择转换模式：")
    print("1. 批量转换（转换文件夹中的所有PDF）")
    print("2. 单个PDF转换")
    
    while True:
        mode = input("\n请输入选项 (1 或 2): ").strip()
        if mode in ['1', '2']:
            break
        print("输入无效，请重新输入！")
    
    # 获取输入路径
    while True:
        input_path = input("\n请输入PDF文件路径或包含PDF的文件夹路径: ").strip()
        # 去除路径两端的引号（如果用户复制了带引号的路径）
        input_path = input_path.strip('"').strip("'")
        
        if os.path.exists(input_path):
            break
        else:
            print(f"路径不存在: {input_path}")
            print("请重新输入正确的路径！")
    
    # 获取输出路径
    default_output = input_path + "_output" if mode == '2' else input_path + "_output"
    output_path = input(f"\n请输入输出文件夹路径 (直接回车使用默认: {default_output}): ").strip()
    output_path = output_path.strip('"').strip("'")
    
    if not output_path:
        output_path = default_output
    
    # 获取DPI
    while True:
        dpi_input = input("\n请输入图片分辨率DPI (直接回车使用默认300): ").strip()
        if not dpi_input:
            dpi = 300
            break
        try:
            dpi = int(dpi_input)
            if 72 <= dpi <= 600:
                break
            else:
                print("DPI建议在72-600之间，请重新输入！")
        except ValueError:
            print("请输入有效的数字！")
    
    # 获取图片格式
    print("\n请选择图片格式：")
    print("1. PNG (推荐，无损压缩，支持DPI元数据)")
    print("2. JPG (有损压缩，文件较小，支持DPI元数据，扩展名为.jpg)")
    print("3. BMP (无压缩，文件较大，不支持DPI元数据)")
    
    format_map = {'1': 'PNG', '2': 'JPG', '3': 'BMP'}  # 统一使用JPG
    while True:
        format_choice = input("请输入选项 (1, 2 或 3，直接回车使用PNG): ").strip()
        if not format_choice:
            image_format = 'PNG'
            break
        if format_choice in format_map:
            image_format = format_map[format_choice]
            if image_format == 'BMP':
                print("注意：BMP格式不支持DPI元数据，检测工具可能仍显示-1。建议使用PNG或JPG格式。")
            elif image_format == 'JPG':
                print("注意：将使用.jpg扩展名保存文件，适合nhdeep检测工具。")
            break
        print("输入无效，请重新输入！")
    
    return {
        'mode': mode,
        'input_path': input_path,
        'output_path': output_path,
        'dpi': dpi,
        'format': image_format
    }

def main():
    """主函数"""
    print("\n欢迎使用PDF转图片工具！")
    print("=" * 60)
    print("图片命名规则：PDF文件名（前15字符）_001, PDF文件名（前15字符）_002, ...")
    print("重要提示：图片将包含正确的DPI元数据，可被检测工具识别")
    print("注意：本程序不生成任何日志文件，所有信息仅显示在控制台")
    print("=" * 60)
    
    # 获取用户输入
    user_config = get_user_input()
    
    print(f"\n开始PDF转图片转换")
    print(f"转换模式: {'批量转换' if user_config['mode'] == '1' else '单个转换'}")
    print(f"输入路径: {user_config['input_path']}")
    print(f"输出路径: {user_config['output_path']}")
    print(f"分辨率: {user_config['dpi']} DPI")
    print(f"图片格式: {user_config['format']}")
    print("-" * 60)
    
    try:
        if user_config['mode'] == '1':
            # 批量转换模式
            if not os.path.isdir(user_config['input_path']):
                print(f"\n错误：批量转换模式需要输入文件夹路径，但提供的路径不是文件夹: {user_config['input_path']}")
                return
            
            # 执行批量转换
            all_results = batch_convert_pdfs(
                input_folder=user_config['input_path'],
                output_base_folder=user_config['output_path'],
                dpi=user_config['dpi'],
                format=user_config['format']
            )
            
            if all_results:
                print(f"\n✓ 批量转换完成！共处理 {len(all_results)} 个PDF文件")
                print(f"  成功: {sum(1 for r in all_results if r['success'])} 个")
                print(f"  失败: {sum(1 for r in all_results if not r['success'])} 个")
                print(f"  输出位置: {user_config['output_path']}")
            else:
                print("\n未找到任何PDF文件，转换结束。")
                
        else:
            # 单个PDF转换模式
            if not os.path.isfile(user_config['input_path']):
                print(f"\n错误：单个转换模式需要输入PDF文件路径，但提供的路径不是文件: {user_config['input_path']}")
                return
            
            # 执行单个转换
            result = convert_single_pdf(
                pdf_path=user_config['input_path'],
                output_folder=user_config['output_path'],
                dpi=user_config['dpi'],
                format=user_config['format']
            )
            
            # 显示结果
            print("\n" + "=" * 60)
            print("转换结果")
            print("=" * 60)
            if result['success']:
                print(f"✓ 转换成功！")
                print(f"  PDF文件: {os.path.basename(result['pdf_path'])}")
                print(f"  页数: {result['expected_pages']}")
                print(f"  输出位置: {result['output_folder']}")
                print(f"  {result['verification_message']}")
                print(f"  图片DPI: {user_config['dpi']} (已写入元数据)")
                
                # 显示生成的图片文件列表
                print(f"\n生成的图片文件:")
                image_files = [f for f in os.listdir(result['output_folder']) 
                              if f.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp'))]
                for img_file in sorted(image_files):
                    print(f"  - {img_file}")
            else:
                print(f"✗ 转换失败！")
                print(f"  错误信息: {result['error']}")
            print("=" * 60)
            
    except Exception as e:
        print(f"\n发生错误: {str(e)}")
    
    print("\n按回车键退出...")
    input()

if __name__ == "__main__":
    
    main()
