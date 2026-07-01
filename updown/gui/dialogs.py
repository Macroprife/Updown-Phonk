"""
独立对话框工具
==============
给没有 GUI 窗口的控制台程序提供文件/文件夹选择对话框。
用法：
    from updown.gui.dialogs import pick_folder, pick_save_file
    folder = pick_folder("选择文件夹")
    path   = pick_save_file("保存文件", ".xlsx", ...)
"""

import tkinter as tk
from tkinter import filedialog


def pick_folder(*, title="选择文件夹", topmost=True) -> str | None:
    """弹出文件夹选择对话框，返回路径或 None"""
    root = tk.Tk()
    root.withdraw()
    if topmost:
        root.attributes("-topmost", True)
    try:
        path = filedialog.askdirectory(title=title)
        return path if path else None
    finally:
        root.destroy()


def pick_open_file(*, title="选择文件",
                   filetypes=None,
                   topmost=True) -> str | None:
    """弹出文件选择对话框，返回路径或 None"""
    if filetypes is None:
        filetypes = [("所有文件", "*.*")]
    root = tk.Tk()
    root.withdraw()
    if topmost:
        root.attributes("-topmost", True)
    try:
        path = filedialog.askopenfilename(title=title, filetypes=filetypes)
        return path if path else None
    finally:
        root.destroy()


def pick_save_file(*, title="保存文件",
                   defaultextension=".xlsx",
                   filetypes=None,
                   initialfile=None,
                   topmost=True) -> str | None:
    """弹出保存文件对话框，返回路径或 None"""
    if filetypes is None:
        filetypes = [("所有文件", "*.*")]
    root = tk.Tk()
    root.withdraw()
    if topmost:
        root.attributes("-topmost", True)
    try:
        kwargs = dict(title=title, defaultextension=defaultextension,
                      filetypes=filetypes)
        if initialfile:
            kwargs["initialfile"] = initialfile
        path = filedialog.asksaveasfilename(**kwargs)
        return path if path else None
    finally:
        root.destroy()
