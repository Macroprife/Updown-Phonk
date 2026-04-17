import pandas as pd
from pathlib import Path

# 手动输入文件路径
file_path = input("请输入Excel文件路径：").strip().strip('"')

# 读取Excel文件中的“抽象”工作表
df = pd.read_excel(file_path, sheet_name="抽象")

# 按“总图片文件夹名”列分组，为每组内的每一行生成序号后缀
df['新列名'] = df.groupby('总图片文件夹名').cumcount() + 1
df['新列名'] = df['总图片文件夹名'] + '-' + df['新列名'].astype(str).str.zfill(3)

# 保存结果（可选：覆盖原文件或另存为新文件）
output_path = Path(file_path).parent / f"{Path(file_path).stem}_已处理.xlsx"
df.to_excel(output_path, index=False)

print(f"处理完成！结果已保存至：{output_path}")
