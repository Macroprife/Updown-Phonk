from pathlib import Path
import shutil

def flatten_pdfs(root_path):
    """
    将指定目录下倒数第二级目录中的PDF文件移到上一级
    支持任意层级结构：零级/一级/二级/.../xx.pdf -> 零级/一级/.../xx.pdf
    """
    root = Path(root_path)
    
    if not root.exists():
        print(f"错误：路径不存在 - {root}")
        return False
    
    if not root.is_dir():
        print(f"错误：不是目录 - {root}")
        return False
    
    moved_count = 0
    skipped_count = 0
    
    # 收集所有PDF文件，记录它们的路径层级
    pdf_files_info = []
    
    # 递归遍历所有子目录
    for pdf_file in root.rglob('*.pdf'):
        # 获取相对于根目录的路径
        relative_path = pdf_file.relative_to(root)
        # 计算路径深度（有多少级目录）
        depth = len(relative_path.parts)
        
        # 只处理至少在两级子目录下的文件（即至少有一级可以去掉）
        if depth >= 2:
            # 目标路径：去掉倒数第二级目录
            # 例如：一级/二级/xx.pdf -> 一级/xx.pdf
            # 例如：一级/二级/三级/xx.pdf -> 一级/三级/xx.pdf
            
            # 获取父目录（文件当前所在目录）
            current_parent = pdf_file.parent
            # 获取上级目录的上级（倒数第二级的父目录）
            target_parent = current_parent.parent
            
            # 如果目标父目录就是根目录
            if target_parent == root:
                target = root / pdf_file.name
            else:
                # 保留除倒数第二级外的路径结构
                target = target_parent / pdf_file.name
            
            pdf_files_info.append((pdf_file, target, current_parent))
    
    # 按文件移动后的路径长度排序，先处理路径短的文件
    pdf_files_info.sort(key=lambda x: len(str(x[1])))
    
    for source, target, current_parent in pdf_files_info:
        # 确保目标目录存在
        target.parent.mkdir(parents=True, exist_ok=True)
        
        # 处理重名文件
        if target.exists():
            stem = target.stem
            suffix = target.suffix
            counter = 1
            while target.exists():
                target = target.parent / f"{stem}_{counter}{suffix}"
                counter += 1
            print(f"文件已存在，重命名为: {target.name}")
        
        try:
            shutil.move(str(source), str(target))
            print(f"✓ 已移动: {source.relative_to(root)} -> {target.relative_to(root)}")
            moved_count += 1
            
            # 尝试删除空目录
            try:
                if not any(current_parent.iterdir()):
                    current_parent.rmdir()
                    print(f"✓ 已删除空目录: {current_parent.relative_to(root)}")
            except:
                pass
                
        except Exception as e:
            print(f"✗ 移动失败 {source.name}: {e}")
            skipped_count += 1
    
    print(f"\n完成！移动了 {moved_count} 个文件，跳过 {skipped_count} 个文件")
    return True

def preview_flatten_pdfs(root_path):
    """
    预览将要进行的操作
    """
    root = Path(root_path)
    
    if not root.exists():
        print(f"错误：路径不存在 - {root}")
        return False
    
    if not root.is_dir():
        print(f"错误：不是目录 - {root}")
        return False
    
    print(f"\n根目录: {root}")
    print("=" * 60)
    
    operations = []
    
    for pdf_file in root.rglob('*.pdf'):
        relative_path = pdf_file.relative_to(root)
        depth = len(relative_path.parts)
        
        if depth >= 2:
            current_parent = pdf_file.parent
            target_parent = current_parent.parent
            
            if target_parent == root:
                target = root / pdf_file.name
            else:
                target = target_parent / pdf_file.name
            
            operations.append((pdf_file, target, relative_path, target.relative_to(root)))
    
    if not operations:
        print("\n未找到需要移动的PDF文件")
        print("注意：文件需要在至少二级子目录下才会被移动")
        return False
    
    # 按目标路径分组显示
    current_dir = None
    for source, target, source_rel, target_rel in sorted(operations, key=lambda x: str(x[1])):
        parent_rel = source.parent.relative_to(root)
        if current_dir != parent_rel:
            current_dir = parent_rel
            print(f"\n当前目录: {parent_rel}/")
        
        if target.exists():
            print(f"  {source.name} -> {target_rel} (需重命名)")
        else:
            print(f"  {source.name} -> {target_rel}")
    
    print("\n" + "=" * 60)
    print(f"共找到 {len(operations)} 个文件需要移动")
    return True

def main():
    """主函数：手动输入路径并选择操作"""
    print("=" * 60)
    print("PDF文件层级调整工具（通用版）")
    print("功能：自动去掉倒数第二级目录，保持其他层级不变")
    print("=" * 60)
    print("\n示例：")
    print("  项目/子项目/文档/文件.pdf -> 项目/文档/文件.pdf")
    print("  根目录/一级/二级/文件.pdf -> 根目录/一级/文件.pdf")
    print("  根目录/一级/文件.pdf       -> 根目录/文件.pdf")
    
    while True:
        print("\n请输入根目录路径：")
        print("（输入 'q' 退出程序）")
        
        root_path = input("\n路径: ").strip().strip('"').strip("'")
        
        if root_path.lower() == 'q':
            print("程序已退出")
            break
        
        if not root_path:
            print("路径不能为空，请重新输入")
            continue
        
        root = Path(root_path)
        
        if not root.exists():
            print(f"\n错误：路径不存在 - {root_path}")
            continue
        
        if not root.is_dir():
            print(f"\n错误：不是有效的目录 - {root_path}")
            continue
        
        while True:
            print(f"\n当前根目录: {root}")
            print("\n请选择操作：")
            print("1. 预览将要进行的操作")
            print("2. 执行移动操作")
            print("3. 重新输入路径")
            print("4. 退出程序")
            
            choice = input("\n请输入选项 (1-4): ").strip()
            
            if choice == '1':
                print("\n【预览模式】")
                preview_flatten_pdfs(root)
                input("\n按回车键继续...")
                
            elif choice == '2':
                print("\n【执行模式】")
                confirm = input(f"确认要去掉 {root} 下所有PDF文件的倒数第二级目录吗？(y/n): ").lower()
                if confirm == 'y':
                    print()
                    flatten_pdfs(root)
                    input("\n按回车键继续...")
                else:
                    print("操作已取消")
                    
            elif choice == '3':
                break
                
            elif choice == '4':
                print("程序已退出")
                return
                
            else:
                print("无效选项，请重新输入")

if __name__ == "__main__":
    main()
