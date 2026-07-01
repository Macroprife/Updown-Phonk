# 档案处理工具集

扫描、统计、转换、分割档案相关文件（Excel / JPG / PDF）的桌面工具包。

## 快速开始

```bash
pip install -r requirements.txt
python 档案处理工具.py
```

## 工具清单

### 📊 目录与数据

| 工具 | 说明 |
|------|------|
| 合并表格 | 扫描文件夹中所有 Excel，提取卷内目录并生成汇总统计表 |
| 生成成品表 | 将合并表格输出转为案卷级 + 文件级成品表 |
| 统计 PDF 与图片 | 分别统计图片和 PDF 文件的数量、大小，生成 Excel 报告 |

### 🖼️ JPG 处理

| 工具 | 说明 |
|------|------|
| 图片转 PDF | 将每个子文件夹中的图片批量合并为 PDF |
| 图片分割 | 根据 Excel 配置按页数将图片批量拆分到不同目录 |

### 📄 PDF 处理

| 工具 | 说明 |
|------|------|
| PDF 转图片 | 批量将 PDF 转换为 JPG/PNG 图片 |
| PDF 删页 | 批量删除 PDF 的第一页或首尾页 |
| PDF 分割 | 根据 Excel 配置将 PDF 按页数拆分为多个文件 |
| 复制不同名文件 | 比较两个文件夹，复制互不匹配的 PDF 文件 |
| PDF 层级迁移 | 去掉 PDF 文件路径中的倒数第二级目录 |
| 扫描空文件夹 | 递归扫描并列出所有空文件夹 |
| 提取文件名 | 批量提取文件名到 Excel |
| 统计 PDF 总页数 | 递归统计文件夹中所有 PDF 的总页数 |

## 打包为 exe

```bash
pip install pyinstaller
pyinstaller --onefile --name "档案处理工具" 档案处理工具.py
```

打包后把 `dist/档案处理工具.exe` 放到本项目根目录，和子文件夹一起分发即可。

## 依赖

- Python 3.8+
- pandas, openpyxl, xlrd
- Pillow
- PyMuPDF (fitz)
- PyPDF2
- tkinter（Python 自带）

## 目录结构

```
图片目录扫描/
├── 档案处理工具.py          # 统一启动菜单
├── requirements.txt
├── 目录类/                  # Excel 数据合并与统计
├── JPG类/                   # 图片转换与分割
└── PDF类/                   # PDF 转换、分割、统计
```
