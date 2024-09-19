import os.path
from tkinter import filedialog
import tkinter as tk
from doc import open_folder, convert_data
from excel_data import *


if __name__ == "__main__":
    # 创建主窗口
    root = tk.Tk()
    root.title("Folder Opener")

    # 设置窗口大小
    root.geometry("600x300")

    # 创建一个不可编辑的text，用于显示选择的文件夹路径
    file_path = tk.StringVar()
    file_path.set("请选择数据文件")
    folder_label = tk.Label(root, textvariable=file_path)
    folder_label.pack(pady=20)

    # 创建一个按钮，点击时会调用open_folder函数
    open_button = tk.Button(root, text="Open Folder", command=lambda: open_folder(file_path))
    open_button.pack(pady=10)

    log_text = tk.Text(root, height=10)

    # 开始转换的按钮
    convert_button = tk.Button(root, text="开始转换", command=lambda: convert_data(file_path, log_text))
    convert_button.pack(pady=10)

    # 显示日志的文本框
    log_text.pack(pady=20)

    # 写入日志
    log_text.insert(tk.END, "日志：\n")

    # 运行主事件循环
    root.mainloop()
