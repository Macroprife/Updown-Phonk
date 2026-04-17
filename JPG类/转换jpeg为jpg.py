import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

class JPEGtoJPGConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("JPEG 转 JPG 批量重命名工具")
        self.root.geometry("600x500")
        
        # 创建界面元素
        self.create_widgets()
    
    def create_widgets(self):
        # 标题
        title = tk.Label(self.root, text="JPEG 转 JPG 批量重命名工具", 
                        font=("Arial", 16, "bold"))
        title.pack(pady=10)
        
        # 说明文字
        info = tk.Label(self.root, text="将文件夹及子文件夹中的所有 .jpeg 文件重命名为 .jpg", 
                       font=("Arial", 10))
        info.pack(pady=5)
        
        # 路径选择框架
        path_frame = tk.Frame(self.root)
        path_frame.pack(pady=20, padx=20, fill=tk.X)
        
        tk.Label(path_frame, text="目标文件夹:", font=("Arial", 10)).pack(side=tk.LEFT)
        
        self.path_var = tk.StringVar()
        self.path_entry = tk.Entry(path_frame, textvariable=self.path_var, width=40)
        self.path_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        browse_btn = tk.Button(path_frame, text="浏览", command=self.browse_folder)
        browse_btn.pack(side=tk.RIGHT)
        
        # 执行按钮
        self.execute_btn = tk.Button(self.root, text="开始转换", command=self.start_conversion,
                                     bg="#4CAF50", fg="white", font=("Arial", 12),
                                     width=15, height=2)
        self.execute_btn.pack(pady=10)
        
        # 输出文本框
        tk.Label(self.root, text="处理日志:", font=("Arial", 10, "bold")).pack(anchor=tk.W, padx=20)
        
        self.output_text = scrolledtext.ScrolledText(self.root, width=70, height=15)
        self.output_text.pack(pady=5, padx=20, fill=tk.BOTH, expand=True)
        
        # 退出按钮
        quit_btn = tk.Button(self.root, text="退出", command=self.root.quit,
                            width=10)
        quit_btn.pack(pady=10)
    
    def browse_folder(self):
        folder = filedialog.askdirectory(title="选择要处理的文件夹")
        if folder:
            self.path_var.set(folder)
    
    def log_message(self, message, tag=None):
        self.output_text.insert(tk.END, message + "\n", tag)
        self.output_text.see(tk.END)
        self.root.update()
    
    def start_conversion(self):
        target_folder = self.path_var.get().strip()
        
        if not target_folder:
            messagebox.showerror("错误", "请选择目标文件夹")
            return
        
        if not os.path.exists(target_folder):
            messagebox.showerror("错误", "文件夹路径不存在")
            return
        
        self.output_text.delete(1.0, tk.END)
        self.execute_btn.config(state=tk.DISABLED)
        
        self.log_message(f"开始处理文件夹: {target_folder}")
        self.log_message("="*50)
        
        count = 0
        errors = 0
        
        for root, dirs, files in os.walk(target_folder):
            for filename in files:
                if filename.lower().endswith('.jpeg'):
                    old_path = os.path.join(root, filename)
                    new_filename = filename[:-5] + '.jpg'
                    new_path = os.path.join(root, new_filename)
                    
                    try:
                        if os.path.exists(new_path):
                            self.log_message(f"⚠ 跳过: {filename} (目标文件已存在)")
                            errors += 1
                        else:
                            os.rename(old_path, new_path)
                            self.log_message(f"✓ 已重命名: {filename} -> {new_filename}")
                            count += 1
                    except Exception as e:
                        self.log_message(f"✗ 重命名失败: {filename} - {e}")
                        errors += 1
        
        self.log_message("="*50)
        self.log_message(f"处理完成！成功: {count} 个，失败/跳过: {errors} 个")
        
        self.execute_btn.config(state=tk.NORMAL)
        messagebox.showinfo("完成", f"处理完成！\n成功: {count} 个\n失败/跳过: {errors} 个")

if __name__ == "__main__":
    root = tk.Tk()
    app = JPEGtoJPGConverter(root)
    root.mainloop()
