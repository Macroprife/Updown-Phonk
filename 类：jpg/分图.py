import os
import shutil
from pathlib import Path
import pandas as pd
import re

def split_images_based_on_excel():
    """
    主函数：根据Excel表格配置分割图片
    """
    # 获取用户输入路径
    excel_path = input("请输入Excel表格文件路径：").strip('"').strip()
    source_folder = input("请输入包含图片的源文件夹路径：").strip('"').strip()
    output_base_folder = input("请输入输出文件夹路径：").strip('"').strip()
    
    # 验证路径
    if not os.path.exists(excel_path):
        print(f"错误：Excel文件路径不存在 - {excel_path}")
        return
    
    if not os.path.exists(source_folder):
        print(f"错误：源文件夹路径不存在 - {source_folder}")
        return
    
    # 创建输出文件夹
    try:
        os.makedirs(output_base_folder, exist_ok=True)
        print(f"输出文件夹已创建/确认: {output_base_folder}")
    except Exception as e:
        print(f"创建输出文件夹失败: {e}")
        return
    
    try:
        # 读取Excel文件
        print(f"正在读取Excel文件: {excel_path}")
        excel_data = pd.read_excel(excel_path, sheet_name='抽象')
    except Exception as e:
        print(f"读取Excel文件失败: {e}")
        print("请确保工作表名称为'抽象'")
        return
    
    # 检查必要的列
    required_columns = ['操作档号', '操作列', '图片文件名', '每份页数']
    for col in required_columns:
        if col not in excel_data.columns:
            print(f"错误：Excel文件中缺少必要的列 '{col}'")
            print(f"当前列: {list(excel_data.columns)}")
            return
    
    # 清理操作档号数据，提取关键部分用于匹配
    print("\n正在清理操作档号数据...")
    excel_data['操作档号_清理'] = excel_data['操作档号'].apply(clean_group_name)
    
    # 获取源文件夹中所有子文件夹的映射
    print("正在扫描源文件夹结构...")
    folder_mapping = scan_source_folders(source_folder)
    
    if not folder_mapping:
        print("错误：源文件夹中没有找到任何子文件夹")
        return
    
    print(f"找到 {len(folder_mapping)} 个源文件夹")
    for folder in list(folder_mapping.keys())[:5]:  # 显示前5个
        print(f"  - {folder}")
    
    # 按清理后的操作档号分组
    print("\n正在按操作档号分组...")
    grouped = excel_data.groupby('操作档号_清理')
    
    # 存储所有分组信息用于勘误检查
    group_info = []
    processed_count = 0
    skipped_count = 0
    
    # 处理每个分组
    for group_key, group_data in grouped:
        original_group_name = str(group_data.iloc[0]['操作档号']).strip()
        
        print(f"\n处理分组: {original_group_name}")
        print(f"清理后的关键名: {group_key}")
        
        # 查找匹配的源文件夹
        matched_folder = find_matching_folder(folder_mapping, group_key)
        
        if not matched_folder:
            print(f"  跳过：未找到与 '{group_key}' 匹配的源文件夹")
            skipped_count += 1
            continue
        
        print(f"  找到匹配的源文件夹: {matched_folder}")
        
        # 获取分组中的理论总页数（每份页数列的值）
        expected_total_pages = int(group_data.iloc[0]['每份页数'])
        
        # 检查分组内每份页数是否一致
        if not (group_data['每份页数'] == expected_total_pages).all():
            print(f"  警告：分组 '{original_group_name}' 中的每份页数不一致！")
            print(f"  每份页数值: {group_data['每份页数'].unique()}")
        
        # 获取源文件夹中的所有图片文件
        source_group_folder = folder_mapping[matched_folder]
        
        # 验证源文件夹是否存在
        if not os.path.exists(source_group_folder):
            print(f"  错误：源文件夹不存在 - {source_group_folder}")
            skipped_count += 1
            continue
            
        image_files = get_image_files(source_group_folder)
        
        if not image_files:
            print(f"  警告：文件夹 '{matched_folder}' 中没有图片文件，跳过此分组")
            skipped_count += 1
            continue
        
        total_images = len(image_files)
        print(f"  找到 {total_images} 张图片")
        print(f"  理论总页数（每份页数）: {expected_total_pages}")
        
        # 对分组数据按操作列排序
        group_data = group_data.sort_values('操作列')
        group_data = group_data.reset_index(drop=True)
        
        # 清理文件名，移除非法字符
        safe_group_name = sanitize_filename(original_group_name)
        
        # 创建分组的目标文件夹（使用清理后的操作档号名称）
        target_group_folder = os.path.join(output_base_folder, safe_group_name)
        try:
            os.makedirs(target_group_folder, exist_ok=True)
            print(f"  创建目标文件夹: {target_group_folder}")
        except Exception as e:
            print(f"  创建目标文件夹失败: {e}")
            skipped_count += 1
            continue
        
        # 处理分组中的每一行
        processed_images = 0
        
        for row_idx, row in group_data.iterrows():
            start_page = int(row['操作列'])
            folder_name = str(row['图片文件名']).strip()
            
            # 清理文件夹名
            safe_folder_name = sanitize_filename(folder_name)
            
            # 计算开始索引（Excel中的页码从1开始，Python列表索引从0开始）
            start_idx = start_page - 1
            
            # 如果是最后一行，处理所有剩余图片
            if row_idx == len(group_data) - 1:
                end_idx = total_images  # 处理到最后一页
            else:
                # 获取下一行的操作列作为当前行的结束位置
                next_start_page = int(group_data.iloc[row_idx + 1]['操作列'])
                end_idx = next_start_page - 1
            
            # 确保索引有效
            if start_idx < 0:
                start_idx = 0
            if end_idx > total_images:
                end_idx = total_images
            
            # 检查是否有图片需要处理
            if start_idx >= total_images:
                print(f"  警告：起始页码 {start_page} 超出图片总数 {total_images}，跳过此行")
                continue
            
            # 创建目标子文件夹
            target_subfolder = os.path.join(target_group_folder, safe_folder_name)
            try:
                os.makedirs(target_subfolder, exist_ok=True)
            except Exception as e:
                print(f"  创建子文件夹失败 '{safe_folder_name}': {e}")
                continue
            
            # 计算需要复制的图片数量
            images_to_copy = end_idx - start_idx
            
            print(f"  分割图片: {start_idx+1}-{end_idx} ({images_to_copy}张) 到 '{folder_name}'")
            
            copied_count = 0
            for i in range(start_idx, end_idx):
                try:
                    src_file = image_files[i]
                    dest_file = os.path.join(target_subfolder, os.path.basename(src_file))
                    
                    # 复制文件（保持原文件名不变）
                    shutil.copy2(src_file, dest_file)
                    copied_count += 1
                except Exception as e:
                    print(f"    复制文件失败 {os.path.basename(src_file)}: {e}")
            
            print(f"  已复制 {copied_count} 张图片")
            processed_images += copied_count
        
        # 检查是否所有图片都被处理了
        if processed_images < total_images:
            print(f"  警告：有 {total_images - processed_images} 张图片未被处理！")
        
        # 记录分组信息用于勘误检查
        group_info.append({
            'group_name': original_group_name,
            'matched_folder': matched_folder,
            'total_pages': total_images,  # 实际图片数
            'expected_pages': expected_total_pages,  # 理论总页数（来自每份页数列）
            'file_count': len(group_data),  # 文件份数
            'processed_images': processed_images
        })
        
        processed_count += 1
    
    # 执行勘误检查
    print("\n" + "="*50)
    print("正在执行勘误检查...")
    run_error_check(group_info)
    
    # 输出统计信息
    print("\n" + "="*50)
    print("处理统计:")
    print(f"  成功处理的分组: {processed_count}")
    print(f"  跳过的分组: {skipped_count}")
    print(f"  总分组数: {len(grouped)}")
    
    # 检查每个分组的所有图片是否都被处理
    if group_info:
        print("\n分组处理详情:")
        for info in group_info:
            status = "✓ 完全处理" if info['processed_images'] == info['total_pages'] else "✗ 部分处理"
            page_match = "✓ 页数匹配" if info['total_pages'] == info['expected_pages'] else "✗ 页数不符"
            print(f"  {info['group_name']}: {status}, {page_match} ({info['processed_images']}/{info['total_pages']}张, 理论{info['expected_pages']}页)")
    
    print(f"\n处理完成！所有分割后的图片已保存到: {output_base_folder}")

def sanitize_filename(filename):
    """
    清理文件名，移除或替换非法字符
    """
    # Windows文件名中的非法字符
    illegal_chars = '<>:"/\\|?*'
    for char in illegal_chars:
        filename = filename.replace(char, '_')
    
    # 移除首尾空格和点号
    filename = filename.strip('. ')
    
    # 如果文件名为空，使用默认名称
    if not filename:
        filename = "unnamed"
    
    return filename

def clean_group_name(group_name):
    """
    清理操作档号，提取关键部分用于匹配
    """
    if pd.isna(group_name):
        return ""
    
    # 转换为字符串
    group_str = str(group_name).strip()
    
    # 提取数字部分（通常是最稳定的标识符）
    digits = re.findall(r'\d+', group_str)
    cleaned = ''.join(digits) if digits else group_str
    
    # 如果包含中文，尝试提取前面的编号部分
    if any('\u4e00' <= char <= '\u9fff' for char in group_str):
        # 分割字符串，取第一部分（通常是编号）
        parts = group_str.split()
        if parts:
            cleaned = parts[0]
    
    # 清理特殊字符
    cleaned = re.sub(r'[^\w\-]', '', cleaned)
    cleaned = cleaned.lower()
    
    return cleaned

def scan_source_folders(base_folder):
    """
    扫描源文件夹中的所有子文件夹，并创建清理后的名称映射
    """
    folder_mapping = {}
    
    try:
        if os.path.exists(base_folder):
            for item in os.listdir(base_folder):
                item_path = os.path.join(base_folder, item)
                if os.path.isdir(item_path):
                    # 清理文件夹名
                    cleaned_name = clean_group_name(item)
                    if cleaned_name:  # 只添加非空名称
                        folder_mapping[cleaned_name] = item_path
                    
                    # 同时添加原始名称的清理版本
                    original_cleaned = clean_group_name(item)
                    if original_cleaned:
                        folder_mapping[original_cleaned] = item_path
    except Exception as e:
        print(f"扫描源文件夹时出错: {e}")
    
    return folder_mapping

def find_matching_folder(folder_mapping, group_key):
    """
    在文件夹映射中查找匹配的文件夹
    """
    # 完全匹配
    if group_key in folder_mapping:
        return group_key
    
    # 尝试更宽松的匹配：比较数字部分
    group_digits = ''.join(re.findall(r'\d+', group_key))
    if group_digits:
        for folder_name in folder_mapping.keys():
            folder_digits = ''.join(re.findall(r'\d+', folder_name))
            if group_digits and folder_digits and group_digits == folder_digits:
                return folder_name
    
    # 部分匹配
    for folder_name in folder_mapping.keys():
        if group_key in folder_name or folder_name in group_key:
            return folder_name
    
    return None

def get_image_files(folder_path):
    """
    获取文件夹中的所有图片文件（按数字顺序排序）
    """
    image_extensions = ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.tif', '.webp']
    image_files = []
    
    try:
        # 首先收集所有图片文件
        for file in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file)
            if os.path.isfile(file_path):
                ext = os.path.splitext(file)[1].lower()
                if ext in image_extensions:
                    image_files.append(file_path)
        
        # 按文件名中的数字部分排序
        def extract_numbers(filename):
            """从文件名中提取数字用于排序"""
            basename = os.path.basename(filename)
            numbers = re.findall(r'\d+', basename)
            return [int(num) for num in numbers] if numbers else [0]
        
        # 按数字顺序排序
        image_files.sort(key=extract_numbers)
        
        # 如果没有数字，按字母顺序排序
        if not any(re.search(r'\d+', os.path.basename(f)) for f in image_files):
            image_files.sort(key=lambda x: os.path.basename(x))
    except Exception as e:
        print(f"获取图片文件时出错: {e}")
    
    return image_files

def run_error_check(group_info):
    """
    执行勘误检查
    根据修正后的逻辑：
    - 每份页数是指该分组应该有的理论总页数
    - 直接比较每份页数与实际图片数
    """
    errors = []
    
    print("\n勘误检查详情:")
    
    for info in group_info:
        group_name = info['group_name']
        total_pages = info['total_pages']  # 实际图片数
        expected_pages = info['expected_pages']  # 理论总页数
        file_count = info['file_count']  # 文件份数
        
        print(f"\n分组 '{group_name}':")
        print(f"  理论总页数（每份页数列的值）: {expected_pages}")
        print(f"  实际图片数: {total_pages}")
        print(f"  文件份数: {file_count}")
        
        # 检查实际图片数是否等于理论总页数
        if total_pages != expected_pages:
            difference = abs(total_pages - expected_pages)
            if total_pages > expected_pages:
                issue_desc = f"实际图片数({total_pages})多于理论总页数({expected_pages})，多出{difference}页"
            else:
                issue_desc = f"实际图片数({total_pages})少于理论总页数({expected_pages})，缺少{difference}页"
            
            errors.append({
                'group': group_name,
                'issue': f"页数不匹配: {issue_desc}",
                'details': f"理论总页数: {expected_pages}, 实际图片数: {total_pages}, 文件份数: {file_count}"
            })
            print(f"  ✗ {issue_desc}")
        else:
            print(f"  ✓ 页数匹配正确（理论{expected_pages}页 = 实际{total_pages}页）")
        
        # 检查是否所有图片都被处理
        if info['processed_images'] != total_pages:
            errors.append({
                'group': group_name,
                'issue': f"图片未完全处理: 仅处理了 {info['processed_images']}/{total_pages} 张图片",
                'details': f"有 {total_pages - info['processed_images']} 张图片未被处理"
            })
            print(f"  ✗ 图片未完全处理: 仅处理了 {info['processed_images']}/{total_pages} 张")
        else:
            print(f"  ✓ 图片完全处理")
    
    # 输出勘误结果汇总
    if errors:
        print("\n" + "="*50)
        print("发现以下勘误问题:")
        for error in errors:
            print(f"\n分组: {error['group']}")
            print(f"问题: {error['issue']}")
            print(f"详情: {error['details']}")
    else:
        print("\n" + "="*50)
        print("✓ 勘误检查全部通过！所有分组页数匹配正确且图片完全处理。")

if __name__ == "__main__":
    split_images_based_on_excel()
