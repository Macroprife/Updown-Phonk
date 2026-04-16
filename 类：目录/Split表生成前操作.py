import pandas as pd
from openpyxl import load_workbook
import os

# 手动输入文件路径
file_path = input("请输入Excel文件路径：").strip().strip('"')

# 读取名为"抽象"的工作表
try:
    df = pd.read_excel(file_path, sheet_name='抽象')
except ValueError:
    print("错误：文件中没有名为'抽象'的工作表")
    exit()

# 检查是否存在"源文件名"列
if '源文件名' not in df.columns:
    print("错误：工作表中没有名为'源文件名'的列")
    exit()

# 新建一列，提取"源文件名"列前15个字符
df['总图片文件夹名'] = df['源文件名'].astype(str).str[:15]

# 生成新表格"BeforeSplit"
output_path = os.path.join(os.path.dirname(file_path), 'BeforeSplit.xlsx')
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='抽象', index=False)

print(f"处理完成！已生成新表格：{output_path}")
print(f"新增列名：'总图片文件夹名'")
print(f"工作表名：'抽象'")
