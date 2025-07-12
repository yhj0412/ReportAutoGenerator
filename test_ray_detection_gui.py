#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试射线检测委托台账GUI模板功能
"""

import tkinter as tk
from tkinter import ttk
import sys
import os

# 添加当前目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def test_ray_detection_template():
    """测试射线检测委托台账模板功能"""
    
    # 创建测试窗口
    root = tk.Tk()
    root.title("射线检测委托台账模板测试")
    root.geometry("800x600")
    
    # 创建主框架
    main_frame = ttk.Frame(root)
    main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    # 模板选择
    template_frame = ttk.LabelFrame(main_frame, text="模板选择")
    template_frame.pack(fill=tk.X, pady=(0, 10))
    
    template_var = tk.StringVar()
    template_combobox = ttk.Combobox(template_frame, textvariable=template_var,
                                   values=["模板1", "模板2"], state="readonly", width=15)
    template_combobox.set("模板2")  # 默认选择模板2
    template_combobox.pack(padx=10, pady=10)
    
    # 参数显示区域
    params_frame = ttk.LabelFrame(main_frame, text="参数设置")
    params_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
    
    # 创建参数显示文本框
    params_text = tk.Text(params_frame, wrap=tk.WORD, height=15)
    params_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    def update_params_display():
        """更新参数显示"""
        selected_template = template_var.get()
        params_text.delete(1.0, tk.END)
        
        if selected_template == "模板1":
            params_info = """模板1 - 8个参数：

1. 工程名称 (project_name)
2. 委托单位 (client_name) 
3. 检测标准 (inspection_standard)
4. 验收规范 (acceptance_specification)
5. 检测方法 (inspection_method)
6. 检测技术等级 (inspection_tech_level)
7. 外观检查 (appearance_check)
8. 坡口形式 (groove_type)

对应文件：Ray_Detection_mode1.py
Word模板：生成器/word/1_射线检测委托台账_Mode1.docx
输出目录：生成器/输出报告/1_射线检测委托台账/1_射线检测委托台账_Mode1/

功能特点：
- 支持8个传参功能
- 支持委托日期填入功能（Excel A列最晚日期填入到Word文档的4个指定位置）
"""
        else:
            params_info = """模板2 - 5个参数：

1. 工程名称 (project_name)
2. 检测类别号 (inspection_category)
3. 检测标准 (inspection_standard)
4. 检测方法 (inspection_method)
5. 坡口形式 (groove_type)

对应文件：Ray_Detection.py
Word模板：生成器/word/1_射线检测委托台账_Mode2.docx
输出目录：生成器/输出报告/1_射线检测委托台账_Mode2/

功能特点：
- 支持5个传参功能
- 标准的Excel到Word文档填充
"""
        
        params_text.insert(1.0, params_info)
    
    # 绑定模板选择变化事件
    template_combobox.bind("<<ComboboxSelected>>", lambda e: update_params_display())
    
    # 初始显示
    update_params_display()
    
    # 测试按钮
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill=tk.X, pady=(0, 10))
    
    def test_template_switch():
        """测试模板切换"""
        current = template_var.get()
        new_template = "模板1" if current == "模板2" else "模板2"
        template_combobox.set(new_template)
        update_params_display()
        print(f"切换到{new_template}")
    
    test_button = ttk.Button(button_frame, text="测试模板切换", command=test_template_switch)
    test_button.pack(side=tk.LEFT, padx=5)
    
    close_button = ttk.Button(button_frame, text="关闭", command=root.destroy)
    close_button.pack(side=tk.RIGHT, padx=5)
    
    # 状态栏
    status_frame = ttk.Frame(main_frame)
    status_frame.pack(fill=tk.X)
    
    status_label = ttk.Label(status_frame, text="射线检测委托台账模板功能测试 - 准备就绪")
    status_label.pack(side=tk.LEFT)
    
    print("射线检测委托台账模板测试窗口已启动")
    print("- 模板1：8个参数，调用Ray_Detection_mode1.py")
    print("- 模板2：5个参数，调用Ray_Detection.py")
    
    root.mainloop()

if __name__ == "__main__":
    test_ray_detection_template()
