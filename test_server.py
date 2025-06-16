import os
import tkinter as tk
from tkinter import filedialog, messagebox

def test_file_browser():
    """测试文件浏览对话框"""
    print("正在测试文件浏览对话框...")
    
    root = tk.Tk()
    root.withdraw()
    
    # 测试Excel文件选择
    print("请选择一个Excel文件...")
    excel_file = filedialog.askopenfilename(
        title="选择Excel文件",
        filetypes=[("Excel文件", "*.xlsx;*.xls")]
    )
    if excel_file:
        print(f"已选择Excel文件: {excel_file}")
    else:
        print("未选择Excel文件")
    
    # 测试Word文件选择
    print("请选择一个Word文件...")
    word_file = filedialog.askopenfilename(
        title="选择Word模板",
        filetypes=[("Word文件", "*.docx")]
    )
    if word_file:
        print(f"已选择Word文件: {word_file}")
    else:
        print("未选择Word文件")
    
    # 测试目录选择
    print("请选择一个输出目录...")
    output_dir = filedialog.askdirectory(title="选择输出目录")
    if output_dir:
        print(f"已选择输出目录: {output_dir}")
    else:
        print("未选择输出目录")
    
    root.destroy()
    
    return excel_file, word_file, output_dir

if __name__ == "__main__":
    excel_file, word_file, output_dir = test_file_browser()
    
    # 显示测试结果
    root = tk.Tk()
    root.title("测试结果")
    root.geometry("600x300")
    
    tk.Label(root, text="文件浏览测试结果", font=("Arial", 16)).pack(pady=10)
    
    tk.Label(root, text=f"Excel文件: {excel_file if excel_file else '未选择'}", anchor="w").pack(fill="x", padx=20, pady=5)
    tk.Label(root, text=f"Word文件: {word_file if word_file else '未选择'}", anchor="w").pack(fill="x", padx=20, pady=5)
    tk.Label(root, text=f"输出目录: {output_dir if output_dir else '未选择'}", anchor="w").pack(fill="x", padx=20, pady=5)
    
    tk.Button(root, text="确定", command=root.destroy).pack(pady=20)
    
    root.mainloop() 