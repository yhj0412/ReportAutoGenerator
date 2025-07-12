#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试GUI模板选择功能
"""

import tkinter as tk
from tkinter import ttk
import sys
import os

# 添加当前目录到路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def test_template_selection():
    """测试模板选择功能"""
    
    # 创建测试窗口
    root = tk.Tk()
    root.title("模板选择测试")
    root.geometry("600x400")
    
    # 创建模板选择框
    frame = ttk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
    
    # 模板选择
    template_label = ttk.Label(frame, text="模板选择:")
    template_label.pack(pady=5)
    
    template_var = tk.StringVar()
    template_combobox = ttk.Combobox(frame, textvariable=template_var, 
                                   values=["模板1", "模板2"], state="readonly")
    template_combobox.set("模板2")  # 默认选择模板2
    template_combobox.pack(pady=5)
    
    # 参数显示区域
    params_frame = ttk.LabelFrame(frame, text="参数设置")
    params_frame.pack(fill=tk.BOTH, expand=True, pady=10)
    
    # 基础参数
    basic_frame = ttk.Frame(params_frame)
    basic_frame.pack(fill=tk.X, padx=10, pady=5)
    
    ttk.Label(basic_frame, text="工程名称:").pack(side=tk.LEFT)
    project_entry = ttk.Entry(basic_frame, width=30)
    project_entry.pack(side=tk.LEFT, padx=5)
    
    ttk.Label(basic_frame, text="委托单位:").pack(side=tk.LEFT, padx=(20, 0))
    client_entry = ttk.Entry(basic_frame, width=20)
    client_entry.pack(side=tk.LEFT, padx=5)
    
    # 模板1专用参数（初始隐藏）
    mode1_frame = ttk.Frame(params_frame)
    
    ttk.Label(mode1_frame, text="检测单位:").pack(side=tk.LEFT)
    inspection_unit_entry = ttk.Entry(mode1_frame, width=20)
    inspection_unit_entry.pack(side=tk.LEFT, padx=5)
    
    ttk.Label(mode1_frame, text="检测标准:").pack(side=tk.LEFT, padx=(20, 0))
    inspection_standard_entry = ttk.Entry(mode1_frame, width=20)
    inspection_standard_entry.pack(side=tk.LEFT, padx=5)
    
    # 状态显示
    status_label = ttk.Label(frame, text="当前选择: 模板2")
    status_label.pack(pady=10)
    
    def on_template_change(event=None):
        """模板选择变化时的回调函数"""
        selected_template = template_var.get()
        
        if selected_template == "模板1":
            # 显示模板1专用参数
            mode1_frame.pack(fill=tk.X, padx=10, pady=5)
            status_label.config(text="当前选择: 模板1 (显示检测单位和检测标准参数)")
        else:
            # 隐藏模板1专用参数
            mode1_frame.pack_forget()
            status_label.config(text="当前选择: 模板2 (隐藏检测单位和检测标准参数)")
    
    # 绑定事件
    template_combobox.bind("<<ComboboxSelected>>", on_template_change)
    
    # 测试按钮
    def test_template1():
        template_var.set("模板1")
        on_template_change()
    
    def test_template2():
        template_var.set("模板2")
        on_template_change()
    
    button_frame = ttk.Frame(frame)
    button_frame.pack(pady=10)
    
    ttk.Button(button_frame, text="测试模板1", command=test_template1).pack(side=tk.LEFT, padx=5)
    ttk.Button(button_frame, text="测试模板2", command=test_template2).pack(side=tk.LEFT, padx=5)
    
    # 说明文本
    info_text = """
模板选择功能说明：
- 模板1：对应 NDT_result_mode1.py，需要5个参数（工程名称、委托单位、检测单位、检测标准、检测方法）
- 模板2：对应 NDT_result.py，需要3个参数（工程名称、委托单位、检测方法）

当选择模板1时，会显示额外的"检测单位"和"检测标准"参数输入框。
当选择模板2时，这些额外参数会被隐藏。
"""
    
    info_label = ttk.Label(frame, text=info_text, justify=tk.LEFT)
    info_label.pack(pady=10)
    
    root.mainloop()

if __name__ == "__main__":
    test_template_selection()
