import sys
import os
import tkinter as tk
from tkinter import filedialog, ttk, scrolledtext, font
import threading
import io
from contextlib import redirect_stdout
from datetime import datetime

# 导入NDT_result模块
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
import NDT_result

class RedirectText:
    """用于重定向stdout到Text控件"""
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.buffer = ""

    def write(self, string):
        self.buffer += string
        # 在主线程中更新UI
        self.text_widget.after(10, self.update_text_widget)
    
    def update_text_widget(self):
        if self.buffer:
            self.text_widget.configure(state='normal')
            self.text_widget.insert(tk.END, self.buffer)
            self.text_widget.see(tk.END)  # 自动滚动到最新内容
            self.text_widget.configure(state='disabled')
            self.buffer = ""
    
    def flush(self):
        pass

class NDTResultGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("NDT结果生成器")
        self.root.geometry("1200x700")
        
        # 获取系统中文字体
        self.default_font = self.get_chinese_font()
        
        # 设置整体样式
        self.configure_styles()
        
        # 创建主框架
        self.main_frame = ttk.Frame(root)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建左侧功能模块区和右侧内容区
        self.create_sidebar()
        self.create_content_area()
        
        # 底部状态栏
        self.create_status_bar()
        
        # 默认选中第二个功能模块
        self.select_module(1)  # RT结果通知单台账
    
    def get_chinese_font(self):
        """获取系统中可用的中文字体"""
        # 常见的中文字体列表，按优先级排序
        chinese_fonts = [
            "Microsoft YaHei", "微软雅黑",  # 微软雅黑
            "SimHei", "黑体",              # 黑体
            "SimSun", "宋体",              # 宋体
            "KaiTi", "楷体",               # 楷体
            "NSimSun", "新宋体",           # 新宋体
            "FangSong", "仿宋",            # 仿宋
            "Arial Unicode MS",           # Arial Unicode
            "Heiti SC", "Heiti TC",       # 苹果系统黑体
            "PingFang SC", "PingFang TC", # 苹果系统平方
            "Noto Sans CJK SC",           # Google Noto字体
            "WenQuanYi Micro Hei"         # 文泉驿微米黑
        ]
        
        # 获取系统可用字体
        available_fonts = font.families()
        
        # 查找第一个可用的中文字体
        for font_name in chinese_fonts:
            if font_name in available_fonts:
                return font_name
        
        # 如果没有找到合适的中文字体，返回默认字体
        return "TkDefaultFont"
        
    def configure_styles(self):
        """配置样式"""
        style = ttk.Style()
        
        # 定义字体
        title_font = (self.default_font, 14, "bold")
        header_font = (self.default_font, 12, "bold")
        normal_font = (self.default_font, 10)
        small_font = (self.default_font, 9)
        
        # 基本样式
        style.configure("TFrame", background="#f5f5f5")
        style.configure("TLabel", background="#f5f5f5", font=normal_font)
        style.configure("TButton", font=normal_font)
        style.configure("TEntry", font=normal_font)
        style.configure("Header.TLabel", font=header_font)
        
        # 左侧菜单样式
        style.configure("Sidebar.TFrame", background="#e8e8e8")
        style.configure("Module.TButton", font=normal_font, padding=10)
        style.configure("ModuleActive.TButton", font=(self.default_font, 10, "bold"), 
                        background="#d0d8ff", padding=10)
        
        # 内容区样式
        style.configure("Content.TFrame", background="#ffffff")
        style.configure("ContentHeader.TLabel", font=title_font, foreground="#333333")
        
        # 标签框样式
        style.configure("TLabelframe", font=normal_font)
        style.configure("TLabelframe.Label", font=header_font, background="#f5f5f5")
        
        # 按钮样式
        style.configure("Submit.TButton", font=(self.default_font, 10, "bold"))
        style.configure("Action.TButton", font=normal_font)
        
        # 设置全局字体
        self.root.option_add("*Font", normal_font)
        
    def create_sidebar(self):
        """创建左侧功能模块区"""
        self.sidebar = ttk.Frame(self.main_frame, style="Sidebar.TFrame", width=220)
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y, padx=0, pady=0)
        self.sidebar.pack_propagate(False)  # 防止宽度被内部组件改变
        
        # 功能模块标题
        module_title_frame = ttk.Frame(self.sidebar, style="Sidebar.TFrame")
        module_title_frame.pack(fill=tk.X, pady=(15, 10), padx=10)
        
        module_title = ttk.Label(module_title_frame, text="功能模块", 
                               style="Header.TLabel", background="#e8e8e8")
        module_title.pack(side=tk.LEFT, padx=5)
        
        # 分隔线
        separator = ttk.Separator(self.sidebar, orient='horizontal')
        separator.pack(fill=tk.X, padx=10, pady=5)
        
        # 功能模块按钮
        self.module_buttons = []
        modules = [
            "1. 射线检测委托台账",
            "2. RT结果通知单台账",
            "3. 表面结果通知单台账",
            "4. 射线检测记录",
            "5. 射线检测记录续"
        ]
        
        modules_frame = ttk.Frame(self.sidebar, style="Sidebar.TFrame")
        modules_frame.pack(fill=tk.X, padx=5, pady=5)
        
        for i, module in enumerate(modules):
            btn = ttk.Button(modules_frame, text=module, style="Module.TButton",
                           command=lambda idx=i: self.select_module(idx))
            btn.pack(fill=tk.X, pady=3)
            self.module_buttons.append(btn)
    
    def create_content_area(self):
        """创建右侧内容区"""
        self.content_frame = ttk.Frame(self.main_frame, style="Content.TFrame")
        self.content_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建各个模块的内容框架
        self.module_frames = []
        for i in range(5):
            frame = ttk.Frame(self.content_frame)
            if i != 1:  # 默认只显示第二个模块（RT结果通知单台账）
                frame.pack_forget()
            self.module_frames.append(frame)
        
        # 创建射线检测委托台账模块的内容
        self.create_ray_detection_frame(self.module_frames[0])
        
        # 创建RT结果通知单台账模块的内容
        self.create_rt_result_frame(self.module_frames[1])
        
        # 创建表面结果通知单台账模块的内容
        self.create_surface_defect_frame(self.module_frames[2])
    
    def create_ray_detection_frame(self, parent_frame):
        """创建射线检测委托台账模块的内容"""
        parent_frame.pack(fill=tk.BOTH, expand=True)
        
        # 模块标题
        header_frame = ttk.Frame(parent_frame)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        header_label = ttk.Label(header_frame, text="射线检测委托台账", 
                               style="ContentHeader.TLabel")
        header_label.pack(side=tk.LEFT, padx=5)
        
        # 参数设置区域
        params_frame = ttk.LabelFrame(parent_frame, text="参数设置")
        params_frame.pack(fill=tk.X, pady=(0, 10), padx=5)
        
        # 创建参数行
        params_grid = ttk.Frame(params_frame)
        params_grid.pack(fill=tk.X, padx=15, pady=15)
        
        # 第一行参数
        row1_frame = ttk.Frame(params_grid)
        row1_frame.pack(fill=tk.X, pady=5)
        
        # 工程名称
        project_label = ttk.Label(row1_frame, text="工程名称")
        project_label.pack(side=tk.LEFT, padx=(0, 5))
        self.ray_project_entry = ttk.Entry(row1_frame, width=20)
        self.ray_project_entry.pack(side=tk.LEFT, padx=(0, 20))
        
        # 检测类别号
        category_label = ttk.Label(row1_frame, text="检测类别号")
        category_label.pack(side=tk.LEFT, padx=(0, 5))
        self.ray_category_entry = ttk.Entry(row1_frame, width=20)
        self.ray_category_entry.pack(side=tk.LEFT)
        
        # 第二行参数
        row2_frame = ttk.Frame(params_grid)
        row2_frame.pack(fill=tk.X, pady=5)
        
        # 检测标准
        standard_label = ttk.Label(row2_frame, text="检测标准")
        standard_label.pack(side=tk.LEFT, padx=(0, 5))
        self.ray_standard_entry = ttk.Entry(row2_frame, width=20)
        self.ray_standard_entry.pack(side=tk.LEFT, padx=(0, 20))
        
        # 检测方法
        method_label = ttk.Label(row2_frame, text="检测方法")
        method_label.pack(side=tk.LEFT, padx=(0, 5))
        self.ray_method_entry = ttk.Entry(row2_frame, width=20)
        self.ray_method_entry.insert(0, "RT")  # 默认值
        self.ray_method_entry.pack(side=tk.LEFT)
        
        # 第三行参数
        row3_frame = ttk.Frame(params_grid)
        row3_frame.pack(fill=tk.X, pady=5)
        
        # 坡口形式
        groove_label = ttk.Label(row3_frame, text="坡口形式")
        groove_label.pack(side=tk.LEFT, padx=(0, 5))
        self.ray_groove_entry = ttk.Entry(row3_frame, width=20)
        self.ray_groove_entry.pack(side=tk.LEFT)
        
        # 文件选择区域
        files_frame = ttk.LabelFrame(parent_frame, text="文件选择")
        files_frame.pack(fill=tk.X, pady=(0, 10), padx=5)
        
        # Excel文件选择
        excel_frame = ttk.Frame(files_frame)
        excel_frame.pack(fill=tk.X, padx=15, pady=(15, 8))
        excel_label = ttk.Label(excel_frame, text="选择输入文件(xlsx)*")
        excel_label.pack(side=tk.LEFT, padx=(0, 5))
        self.ray_excel_path = tk.StringVar()
        self.ray_excel_path.set("生成器/Excel/1_生成器委托.xlsx")  # 默认值
        excel_entry = ttk.Entry(excel_frame, textvariable=self.ray_excel_path)
        excel_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        excel_button = ttk.Button(excel_frame, text="浏览...", command=self.browse_ray_excel)
        excel_button.pack(side=tk.LEFT)
        
        # Word模板选择
        word_frame = ttk.Frame(files_frame)
        word_frame.pack(fill=tk.X, padx=15, pady=8)
        word_label = ttk.Label(word_frame, text="选择Word模板文件(docx)*")
        word_label.pack(side=tk.LEFT, padx=(0, 5))
        self.ray_word_path = tk.StringVar()
        self.ray_word_path.set("生成器/wod/1_射线检测委托台账_Mode2.docx")  # 默认值
        word_entry = ttk.Entry(word_frame, textvariable=self.ray_word_path)
        word_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        word_button = ttk.Button(word_frame, text="浏览...", command=self.browse_ray_word)
        word_button.pack(side=tk.LEFT)
        
        # 输出文件夹选择
        output_frame = ttk.Frame(files_frame)
        output_frame.pack(fill=tk.X, padx=15, pady=(8, 15))
        output_label = ttk.Label(output_frame, text="选择输出文件夹*")
        output_label.pack(side=tk.LEFT, padx=(0, 5))
        self.ray_output_path = tk.StringVar()
        output_entry = ttk.Entry(output_frame, textvariable=self.ray_output_path)
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        output_button = ttk.Button(output_frame, text="浏览...", command=self.browse_ray_output)
        output_button.pack(side=tk.LEFT)
        
        # 提交按钮
        submit_frame = ttk.Frame(parent_frame)
        submit_frame.pack(fill=tk.X, pady=10)
        self.ray_submit_button = ttk.Button(submit_frame, text="提交", 
                                        style="Submit.TButton", command=self.process_ray_data)
        self.ray_submit_button.pack(side=tk.RIGHT, padx=10)
        
        # 日志区域
        log_frame = ttk.LabelFrame(parent_frame, text="执行日志")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10), padx=5)
        
        # 创建滚动文本框
        self.ray_log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10)
        self.ray_log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.ray_log_text.configure(state='disabled')
        
        # 设置日志文本字体
        self.ray_log_text.configure(font=(self.default_font, 9))
        
        # 日志操作按钮
        log_buttons_frame = ttk.Frame(log_frame)
        log_buttons_frame.pack(fill=tk.X, padx=5, pady=(0, 5))
        
        export_button = ttk.Button(log_buttons_frame, text="导出", 
                                 style="Action.TButton", command=self.export_ray_log)
        export_button.pack(side=tk.RIGHT, padx=5)
        
        clear_button = ttk.Button(log_buttons_frame, text="清空", 
                                style="Action.TButton", command=self.clear_ray_log)
        clear_button.pack(side=tk.RIGHT, padx=5)
        
        # 设置日志重定向
        self.ray_redirect = RedirectText(self.ray_log_text)
    
    def create_rt_result_frame(self, parent_frame):
        """创建RT结果通知单台账模块的内容"""
        parent_frame.pack(fill=tk.BOTH, expand=True)
        
        # 模块标题
        header_frame = ttk.Frame(parent_frame)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        header_label = ttk.Label(header_frame, text="RT结果通知单台账", 
                                style="ContentHeader.TLabel")
        header_label.pack(side=tk.LEFT, padx=5)
        
        # 参数设置区域
        params_frame = ttk.LabelFrame(parent_frame, text="参数设置")
        params_frame.pack(fill=tk.X, pady=(0, 10), padx=5)
        
        # 创建参数行
        params_grid = ttk.Frame(params_frame)
        params_grid.pack(fill=tk.X, padx=15, pady=15)
        
        # 第一行参数
        row1_frame = ttk.Frame(params_grid)
        row1_frame.pack(fill=tk.X, pady=5)
        
        # 工程名称
        project_label = ttk.Label(row1_frame, text="工程名称")
        project_label.pack(side=tk.LEFT, padx=(0, 5))
        self.project_entry = ttk.Entry(row1_frame, width=40)
        self.project_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        # 第二行参数
        row2_frame = ttk.Frame(params_grid)
        row2_frame.pack(fill=tk.X, pady=5)
        
        # 委托单位
        client_label = ttk.Label(row2_frame, text="委托单位")
        client_label.pack(side=tk.LEFT, padx=(0, 5))
        self.client_entry = ttk.Entry(row2_frame, width=20)
        self.client_entry.pack(side=tk.LEFT, padx=(0, 30))
        
        # 检测方法
        method_label = ttk.Label(row2_frame, text="检测方法")
        method_label.pack(side=tk.LEFT, padx=(0, 5))
        self.method_entry = ttk.Entry(row2_frame, width=20)
        self.method_entry.insert(0, "RT")  # 默认值
        self.method_entry.pack(side=tk.LEFT)
        
        # 文件选择区域
        files_frame = ttk.LabelFrame(parent_frame, text="文件选择")
        files_frame.pack(fill=tk.X, pady=(0, 10), padx=5)
        
        # Excel文件选择
        excel_frame = ttk.Frame(files_frame)
        excel_frame.pack(fill=tk.X, padx=15, pady=(15, 8))
        excel_label = ttk.Label(excel_frame, text="选择输入文件(xlsx)*")
        excel_label.pack(side=tk.LEFT, padx=(0, 5))
        self.excel_path = tk.StringVar()
        excel_entry = ttk.Entry(excel_frame, textvariable=self.excel_path)
        excel_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        excel_button = ttk.Button(excel_frame, text="浏览...", command=self.browse_excel)
        excel_button.pack(side=tk.LEFT)
        
        # Word模板选择
        word_frame = ttk.Frame(files_frame)
        word_frame.pack(fill=tk.X, padx=15, pady=8)
        word_label = ttk.Label(word_frame, text="选择Word模板文件(docx)*")
        word_label.pack(side=tk.LEFT, padx=(0, 5))
        self.word_path = tk.StringVar()
        word_entry = ttk.Entry(word_frame, textvariable=self.word_path)
        word_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        word_button = ttk.Button(word_frame, text="浏览...", command=self.browse_word)
        word_button.pack(side=tk.LEFT)
        
        # 输出文件夹选择
        output_frame = ttk.Frame(files_frame)
        output_frame.pack(fill=tk.X, padx=15, pady=(8, 15))
        output_label = ttk.Label(output_frame, text="选择输出文件夹*")
        output_label.pack(side=tk.LEFT, padx=(0, 5))
        self.output_path = tk.StringVar()
        output_entry = ttk.Entry(output_frame, textvariable=self.output_path)
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        output_button = ttk.Button(output_frame, text="浏览...", command=self.browse_output)
        output_button.pack(side=tk.LEFT)
        
        # 提交按钮
        submit_frame = ttk.Frame(parent_frame)
        submit_frame.pack(fill=tk.X, pady=10)
        self.submit_button = ttk.Button(submit_frame, text="提交", 
                                      style="Submit.TButton", command=self.process_data)
        self.submit_button.pack(side=tk.RIGHT, padx=10)
        
        # 日志区域
        log_frame = ttk.LabelFrame(parent_frame, text="执行日志")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10), padx=5)
        
        # 创建滚动文本框
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.log_text.configure(state='disabled')
        
        # 设置日志文本字体
        self.log_text.configure(font=(self.default_font, 9))
        
        # 日志操作按钮
        log_buttons_frame = ttk.Frame(log_frame)
        log_buttons_frame.pack(fill=tk.X, padx=5, pady=(0, 5))
        
        export_button = ttk.Button(log_buttons_frame, text="导出", 
                                 style="Action.TButton", command=self.export_log)
        export_button.pack(side=tk.RIGHT, padx=5)
        
        clear_button = ttk.Button(log_buttons_frame, text="清空", 
                                style="Action.TButton", command=self.clear_log)
        clear_button.pack(side=tk.RIGHT, padx=5)
        
        # 设置日志重定向
        self.redirect = RedirectText(self.log_text)
    
    def create_surface_defect_frame(self, parent_frame):
        """创建表面结果通知单台账模块的内容"""
        parent_frame.pack(fill=tk.BOTH, expand=True)
        
        # 模块标题
        header_frame = ttk.Frame(parent_frame)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        header_label = ttk.Label(header_frame, text="表面结果通知单台账", 
                                style="ContentHeader.TLabel")
        header_label.pack(side=tk.LEFT, padx=5)
        
        # 参数设置区域
        params_frame = ttk.LabelFrame(parent_frame, text="参数设置")
        params_frame.pack(fill=tk.X, pady=(0, 10), padx=5)
        
        # 创建参数行
        params_grid = ttk.Frame(params_frame)
        params_grid.pack(fill=tk.X, padx=15, pady=15)
        
        # 第一行参数
        row1_frame = ttk.Frame(params_grid)
        row1_frame.pack(fill=tk.X, pady=5)
        
        # 工程名称
        project_label = ttk.Label(row1_frame, text="工程名称")
        project_label.pack(side=tk.LEFT, padx=(0, 5))
        self.surface_project_entry = ttk.Entry(row1_frame, width=40)
        self.surface_project_entry.pack(side=tk.LEFT)
        
        # 第二行参数
        row2_frame = ttk.Frame(params_grid)
        row2_frame.pack(fill=tk.X, pady=5)
        
        # 委托单位
        client_label = ttk.Label(row2_frame, text="委托单位")
        client_label.pack(side=tk.LEFT, padx=(0, 5))
        self.surface_client_entry = ttk.Entry(row2_frame, width=20)
        self.surface_client_entry.pack(side=tk.LEFT, padx=(0, 30))
        
        # 检测方法
        method_label = ttk.Label(row2_frame, text="检测方法")
        method_label.pack(side=tk.LEFT, padx=(0, 5))
        self.surface_method_entry = ttk.Entry(row2_frame, width=20)
        self.surface_method_entry.insert(0, "表面检测")  # 默认值
        self.surface_method_entry.pack(side=tk.LEFT)
        
        # 文件选择区域
        files_frame = ttk.LabelFrame(parent_frame, text="文件选择")
        files_frame.pack(fill=tk.X, pady=(0, 10), padx=5)
        
        # Excel文件选择
        excel_frame = ttk.Frame(files_frame)
        excel_frame.pack(fill=tk.X, padx=15, pady=(15, 8))
        excel_label = ttk.Label(excel_frame, text="选择输入文件(xlsx)*")
        excel_label.pack(side=tk.LEFT, padx=(0, 5))
        self.surface_excel_path = tk.StringVar()
        self.surface_excel_path.set("生成器/Excel/3_生成器表面结果.xlsx")  # 默认值
        excel_entry = ttk.Entry(excel_frame, textvariable=self.surface_excel_path)
        excel_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        excel_button = ttk.Button(excel_frame, text="浏览...", command=self.browse_surface_excel)
        excel_button.pack(side=tk.LEFT)
        
        # Word模板选择
        word_frame = ttk.Frame(files_frame)
        word_frame.pack(fill=tk.X, padx=15, pady=8)
        word_label = ttk.Label(word_frame, text="选择Word模板文件(docx)*")
        word_label.pack(side=tk.LEFT, padx=(0, 5))
        self.surface_word_path = tk.StringVar()
        self.surface_word_path.set("生成器/wod/3_表面结果通知单台账_Mode2.docx")  # 默认值
        word_entry = ttk.Entry(word_frame, textvariable=self.surface_word_path)
        word_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        word_button = ttk.Button(word_frame, text="浏览...", command=self.browse_surface_word)
        word_button.pack(side=tk.LEFT)
        
        # 输出文件夹选择
        output_frame = ttk.Frame(files_frame)
        output_frame.pack(fill=tk.X, padx=15, pady=(8, 15))
        output_label = ttk.Label(output_frame, text="选择输出文件夹*")
        output_label.pack(side=tk.LEFT, padx=(0, 5))
        self.surface_output_path = tk.StringVar()
        output_entry = ttk.Entry(output_frame, textvariable=self.surface_output_path)
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        output_button = ttk.Button(output_frame, text="浏览...", command=self.browse_surface_output)
        output_button.pack(side=tk.LEFT)
        
        # 提交按钮
        submit_frame = ttk.Frame(parent_frame)
        submit_frame.pack(fill=tk.X, pady=10)
        self.surface_submit_button = ttk.Button(submit_frame, text="提交", 
                                        style="Submit.TButton", command=self.process_surface_data)
        self.surface_submit_button.pack(side=tk.RIGHT, padx=10)
        
        # 日志区域
        log_frame = ttk.LabelFrame(parent_frame, text="执行日志")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10), padx=5)
        
        # 创建滚动文本框
        self.surface_log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10)
        self.surface_log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.surface_log_text.configure(state='disabled')
        
        # 设置日志文本字体
        self.surface_log_text.configure(font=(self.default_font, 9))
        
        # 日志操作按钮
        log_buttons_frame = ttk.Frame(log_frame)
        log_buttons_frame.pack(fill=tk.X, padx=5, pady=(0, 5))
        
        export_button = ttk.Button(log_buttons_frame, text="导出", 
                                 style="Action.TButton", command=self.export_surface_log)
        export_button.pack(side=tk.RIGHT, padx=5)
        
        clear_button = ttk.Button(log_buttons_frame, text="清空", 
                                style="Action.TButton", command=self.clear_surface_log)
        clear_button.pack(side=tk.RIGHT, padx=5)
        
        # 设置日志重定向
        self.surface_redirect = RedirectText(self.surface_log_text)
    
    def create_status_bar(self):
        """创建状态栏"""
        status_frame = ttk.Frame(self.root)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 左侧状态信息
        self.status_var = tk.StringVar()
        self.status_var.set("状态: 失败")
        status_label = ttk.Label(status_frame, textvariable=self.status_var, 
                               relief=tk.SUNKEN, anchor=tk.W)
        status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 右侧处理信息
        self.process_var = tk.StringVar()
        self.process_var.set("0/0 文件已处理")
        process_label = ttk.Label(status_frame, textvariable=self.process_var, 
                                relief=tk.SUNKEN, anchor=tk.E)
        process_label.pack(side=tk.RIGHT, padx=(5, 0))
    
    def select_module(self, index):
        """选择功能模块"""
        # 更新按钮样式
        for i, btn in enumerate(self.module_buttons):
            if i == index:
                btn.configure(style="ModuleActive.TButton")
            else:
                btn.configure(style="Module.TButton")
        
        # 显示选中的模块内容
        for i, frame in enumerate(self.module_frames):
            if i == index:
                frame.pack(fill=tk.BOTH, expand=True)
            else:
                frame.pack_forget()
    
    def browse_excel(self):
        """浏览选择Excel文件"""
        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        if filename:
            self.excel_path.set(filename)
            
    def browse_word(self):
        """浏览选择Word模板文件"""
        filename = filedialog.askopenfilename(
            title="选择Word模板文件",
            filetypes=[("Word文件", "*.docx *.doc")]
        )
        if filename:
            self.word_path.set(filename)
            
    def browse_output(self):
        """浏览选择输出文件夹"""
        directory = filedialog.askdirectory(title="选择输出文件夹")
        if directory:
            self.output_path.set(directory)
    
    def clear_log(self):
        """清空日志"""
        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')
    
    def export_log(self):
        """导出日志"""
        filename = filedialog.asksaveasfilename(
            title="导出日志",
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt")]
        )
        if filename:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(self.log_text.get(1.0, tk.END))
            self.show_log(f"日志已导出到: {filename}")
            
    def process_data(self):
        """处理数据并生成报告"""
        # 获取输入值
        excel_path = self.excel_path.get()
        word_path = self.word_path.get()
        output_path = self.output_path.get()
        project_name = self.project_entry.get()
        client_name = self.client_entry.get()
        inspection_method = self.method_entry.get()
        
        # 验证输入
        if not excel_path or not os.path.exists(excel_path):
            self.show_log("错误: 请选择有效的Excel文件")
            return
            
        if not word_path or not os.path.exists(word_path):
            self.show_log("错误: 请选择有效的Word模板文件")
            return
        
        if not output_path:
            # 使用默认输出路径
            output_path = os.path.join("生成器", "输出报告")
            self.output_path.set(output_path)
            self.show_log(f"未指定输出文件夹，使用默认路径: {output_path}")
            
            # 确保目录存在
            if not os.path.exists(output_path):
                try:
                    os.makedirs(output_path)
                    self.show_log(f"创建输出目录: {output_path}")
                except Exception as e:
                    self.show_log(f"创建目录失败: {e}")
                    return
        
        # 禁用提交按钮，避免重复提交
        self.submit_button.configure(state='disabled')
        self.status_var.set("状态: 处理中...")
        
        # 显示开始信息
        self.show_log(f"开始处理数据: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.show_log(f"Excel文件: {excel_path}")
        self.show_log(f"Word模板: {word_path}")
        self.show_log(f"输出路径: {output_path}")
        self.show_log(f"工程名称: {project_name}")
        self.show_log(f"委托单位: {client_name}")
        self.show_log(f"检测方法: {inspection_method}")
        self.show_log("="*50)
        
        # 在后台线程中处理数据
        threading.Thread(target=self.run_process, args=(
            excel_path, word_path, output_path, project_name, client_name, inspection_method
        )).start()
        
    def run_process(self, excel_path, word_path, output_path, project_name, client_name, inspection_method):
        """在后台线程中运行数据处理"""
        try:
            # 重定向标准输出到日志区
            with redirect_stdout(self.redirect):
                # 调用NDT_result模块的处理函数
                success = NDT_result.process_excel_to_word(
                    excel_path, word_path, output_path, project_name, client_name, inspection_method
                )
            
            # 在主线程中更新UI
            self.root.after(0, self.process_completed, success)
            
        except Exception as e:
            # 在主线程中显示错误
            self.root.after(0, self.show_error, str(e))
            
    def process_completed(self, success):
        """处理完成后的回调"""
        if success:
            self.status_var.set("状态: 完成")
            self.show_log("\n处理成功完成!")
        else:
            self.status_var.set("状态: 失败")
            self.show_log("\n处理失败!")
            
        # 重新启用提交按钮
        self.submit_button.configure(state='normal')
        
    def show_error(self, error_msg):
        """显示错误信息"""
        self.show_log(f"\n错误: {error_msg}")
        self.status_var.set("状态: 处理出错")
        self.submit_button.configure(state='normal')
        
    def show_log(self, message):
        """在日志区显示消息"""
        self.log_text.configure(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)  # 自动滚动到最新内容
        self.log_text.configure(state='disabled')

    def browse_ray_excel(self):
        """浏览选择射线检测委托台账Excel文件"""
        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        if filename:
            self.ray_excel_path.set(filename)
            
    def browse_ray_word(self):
        """浏览选择射线检测委托台账Word模板文件"""
        filename = filedialog.askopenfilename(
            title="选择Word模板文件",
            filetypes=[("Word文件", "*.docx *.doc")]
        )
        if filename:
            self.ray_word_path.set(filename)
            
    def browse_ray_output(self):
        """浏览选择射线检测委托台账输出文件夹"""
        directory = filedialog.askdirectory(title="选择输出文件夹")
        if directory:
            self.ray_output_path.set(directory)

    def clear_ray_log(self):
        """清空射线检测委托台账日志"""
        self.ray_log_text.configure(state='normal')
        self.ray_log_text.delete(1.0, tk.END)
        self.ray_log_text.configure(state='disabled')

    def export_ray_log(self):
        """导出射线检测委托台账日志"""
        filename = filedialog.asksaveasfilename(
            title="导出日志",
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt")]
        )
        if filename:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(self.ray_log_text.get(1.0, tk.END))
            self.show_ray_log(f"日志已导出到: {filename}")

    def show_ray_log(self, message):
        """在射线检测委托台账日志区显示消息"""
        self.ray_log_text.configure(state='normal')
        self.ray_log_text.insert(tk.END, message + "\n")
        self.ray_log_text.see(tk.END)  # 自动滚动到最新内容
        self.ray_log_text.configure(state='disabled')

    def process_ray_data(self):
        """处理射线检测委托台账数据"""
        # 获取输入值
        excel_path = self.ray_excel_path.get()
        word_path = self.ray_word_path.get()
        output_path = self.ray_output_path.get()
        project_name = self.ray_project_entry.get()
        category = self.ray_category_entry.get()
        standard = self.ray_standard_entry.get()
        method = self.ray_method_entry.get()
        groove = self.ray_groove_entry.get()
        
        # 验证输入
        if not excel_path or not os.path.exists(excel_path):
            self.show_ray_log("错误: 请选择有效的Excel文件")
            return
        
        if not word_path or not os.path.exists(word_path):
            self.show_ray_log("错误: 请选择有效的Word模板文件")
            return
        
        if not output_path:
            # 使用默认输出路径
            template_name = os.path.splitext(os.path.basename(word_path))[0]
            output_path = os.path.join("生成器", "输出报告", template_name)
            self.ray_output_path.set(output_path)
            self.show_ray_log(f"未指定输出文件夹，使用默认路径: {output_path}")
            
            # 确保目录存在
            if not os.path.exists(output_path):
                try:
                    os.makedirs(output_path)
                    self.show_ray_log(f"创建输出目录: {output_path}")
                except Exception as e:
                    self.show_ray_log(f"创建目录失败: {e}")
                    return
        
        # 禁用提交按钮，避免重复提交
        self.ray_submit_button.configure(state='disabled')
        self.status_var.set("状态: 处理中...")
        
        # 显示开始信息
        self.show_ray_log(f"开始处理数据: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.show_ray_log(f"Excel文件: {excel_path}")
        self.show_ray_log(f"Word模板: {word_path}")
        self.show_ray_log(f"输出路径: {output_path}")
        self.show_ray_log(f"工程名称: {project_name}")
        self.show_ray_log(f"检测类别号: {category}")
        self.show_ray_log(f"检测标准: {standard}")
        self.show_ray_log(f"检测方法: {method}")
        self.show_ray_log(f"坡口形式: {groove}")
        self.show_ray_log("="*50)
        
        # 在后台线程中处理数据
        threading.Thread(target=self.run_ray_process, args=(
            excel_path, word_path, output_path, project_name, category, 
            standard, method, groove
        )).start()

    def run_ray_process(self, excel_path, word_path, output_path, project_name, category, 
                      standard, method, groove):
        """在后台线程中运行射线检测委托台账处理"""
        try:
            # 导入Ray_Detection模块
            sys.path.append(os.path.dirname(os.path.abspath(__file__)))
            import Ray_Detection
            
            # 重定向标准输出到日志区
            with redirect_stdout(self.ray_redirect):
                # 调用Ray_Detection模块的处理函数
                success = Ray_Detection.process_excel_to_word(
                    excel_path, word_path, output_path, project_name, category, 
                    standard, method, groove
                )
            
            # 在主线程中更新UI
            self.root.after(0, self.process_ray_completed, success)
            
        except Exception as e:
            # 在主线程中显示错误
            self.root.after(0, self.show_ray_error, str(e))

    def process_ray_completed(self, success):
        """射线检测委托台账处理完成后的回调"""
        if success:
            self.status_var.set("状态: 完成")
            self.show_ray_log("\n处理成功完成!")
        else:
            self.status_var.set("状态: 失败")
            self.show_ray_log("\n处理失败!")
            
        # 重新启用提交按钮
        self.ray_submit_button.configure(state='normal')
        
    def show_ray_error(self, error_msg):
        """显示射线检测委托台账错误信息"""
        self.show_ray_log(f"\n错误: {error_msg}")
        self.status_var.set("状态: 处理出错")
        self.ray_submit_button.configure(state='normal')

    def browse_surface_excel(self):
        """浏览选择表面结果通知单台账Excel文件"""
        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        if filename:
            self.surface_excel_path.set(filename)
            
    def browse_surface_word(self):
        """浏览选择表面结果通知单台账Word模板文件"""
        filename = filedialog.askopenfilename(
            title="选择Word模板文件",
            filetypes=[("Word文件", "*.docx *.doc")]
        )
        if filename:
            self.surface_word_path.set(filename)
            
    def browse_surface_output(self):
        """浏览选择表面结果通知单台账输出文件夹"""
        directory = filedialog.askdirectory(title="选择输出文件夹")
        if directory:
            self.surface_output_path.set(directory)

    def process_surface_data(self):
        """处理表面结果通知单台账数据"""
        # 获取输入值
        excel_path = self.surface_excel_path.get()
        word_path = self.surface_word_path.get()
        output_path = self.surface_output_path.get()
        project_name = self.surface_project_entry.get()
        client_name = self.surface_client_entry.get()
        inspection_method = self.surface_method_entry.get()
        
        # 验证输入
        if not excel_path or not os.path.exists(excel_path):
            self.show_surface_log("错误: 请选择有效的Excel文件")
            return
        
        if not word_path or not os.path.exists(word_path):
            self.show_surface_log("错误: 请选择有效的Word模板文件")
            return
        
        if not output_path:
            # 使用默认输出路径
            output_path = os.path.join("生成器", "输出报告", "3_表面结果通知单台账_Mode2")
            self.surface_output_path.set(output_path)
            self.show_surface_log(f"未指定输出文件夹，使用默认路径: {output_path}")
            
            # 确保目录存在
            if not os.path.exists(output_path):
                try:
                    os.makedirs(output_path)
                    self.show_surface_log(f"创建输出目录: {output_path}")
                except Exception as e:
                    self.show_surface_log(f"创建目录失败: {e}")
                    return
        
        # 禁用提交按钮，避免重复提交
        self.surface_submit_button.configure(state='disabled')
        self.status_var.set("状态: 处理中...")
        
        # 显示开始信息
        self.show_surface_log(f"开始处理数据: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.show_surface_log(f"Excel文件: {excel_path}")
        self.show_surface_log(f"Word模板: {word_path}")
        self.show_surface_log(f"输出路径: {output_path}")
        self.show_surface_log(f"工程名称: {project_name}")
        self.show_surface_log(f"委托单位: {client_name}")
        self.show_surface_log(f"检测方法: {inspection_method}")
        self.show_surface_log("="*50)
        
        # 在后台线程中处理数据
        threading.Thread(target=self.run_surface_process, args=(
            excel_path, word_path, output_path, project_name, client_name, inspection_method
        )).start()

    def run_surface_process(self, excel_path, word_path, output_path, project_name, client_name, inspection_method):
        """在后台线程中运行表面结果通知单台账处理"""
        try:
            # 导入Surface_Defect模块
            sys.path.append(os.path.dirname(os.path.abspath(__file__)))
            import Surface_Defect
            
            # 重定向标准输出到日志区
            with redirect_stdout(self.surface_redirect):
                # 调用Surface_Defect模块的处理函数
                success = Surface_Defect.process_excel_to_word(
                    excel_path, word_path, output_path, project_name, client_name, inspection_method
                )
            
            # 在主线程中更新UI
            self.root.after(0, self.process_surface_completed, success)
            
        except Exception as e:
            # 在主线程中显示错误
            self.root.after(0, self.show_surface_error, str(e))

    def process_surface_completed(self, success):
        """表面结果通知单台账处理完成后的回调"""
        if success:
            self.status_var.set("状态: 完成")
            self.show_surface_log("\n处理成功完成!")
        else:
            self.status_var.set("状态: 失败")
            self.show_surface_log("\n处理失败!")
            
        # 重新启用提交按钮
        self.surface_submit_button.configure(state='normal')
        
    def show_surface_error(self, error_msg):
        """显示表面结果通知单台账错误信息"""
        self.show_surface_log(f"\n错误: {error_msg}")
        self.status_var.set("状态: 处理出错")
        self.surface_submit_button.configure(state='normal')
        
    def clear_surface_log(self):
        """清空表面结果通知单台账日志"""
        self.surface_log_text.configure(state='normal')
        self.surface_log_text.delete(1.0, tk.END)
        self.surface_log_text.configure(state='disabled')
    
    def export_surface_log(self):
        """导出表面结果通知单台账日志"""
        filename = filedialog.asksaveasfilename(
            title="导出日志",
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt")]
        )
        if filename:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(self.surface_log_text.get(1.0, tk.END))
            self.show_surface_log(f"日志已导出到: {filename}")
    
    def show_surface_log(self, message):
        """在表面结果通知单台账日志区显示消息"""
        self.surface_log_text.configure(state='normal')
        self.surface_log_text.insert(tk.END, message + "\n")
        self.surface_log_text.see(tk.END)  # 自动滚动到最新内容
        self.surface_log_text.configure(state='disabled')

if __name__ == "__main__":
    root = tk.Tk()
    app = NDTResultGUI(root)
    root.mainloop()