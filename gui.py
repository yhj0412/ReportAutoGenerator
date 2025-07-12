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
import NDT_result_mode1

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

        # 添加技术支持信息区域
        self.create_sidebar_support_info()

    def create_sidebar_support_info(self):
        """在左侧边栏创建技术支持信息"""
        # 添加一些垂直间距
        spacer_frame = ttk.Frame(self.sidebar, style="Sidebar.TFrame", height=20)
        spacer_frame.pack(fill=tk.X, pady=10)

        # 技术支持信息区域
        support_frame = ttk.Frame(self.sidebar, style="Sidebar.TFrame")
        support_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=10, pady=(10, 15))

        # 分隔线
        separator = ttk.Separator(support_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=(0, 10))

        # 技术支持标题
        support_title = ttk.Label(support_frame, text="📞 技术支持",
                                font=(self.default_font, 10, "bold"),
                                foreground="#2c5aa0",
                                background="#e8e8e8")
        support_title.pack(anchor=tk.W, pady=(0, 5))

        # 技术支持联系信息
        support_contact = ttk.Label(support_frame,
                                  text="如遇到问题，请联系\n1594445261@qq.com\n技术支持",
                                  font=(self.default_font, 8),
                                  foreground="#666666",
                                  background="#e8e8e8",
                                  justify=tk.LEFT)
        support_contact.pack(anchor=tk.W, pady=(0, 5))

        # 版权信息
        copyright_info = ttk.Label(support_frame,
                                 text="© 2025 NDT结果生成器\n保留所有权利",
                                 font=(self.default_font, 7),
                                 foreground="#888888",
                                 background="#e8e8e8",
                                 justify=tk.LEFT)
        copyright_info.pack(anchor=tk.W)

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
        
        # 创建射线检测记录模块的内容
        self.create_radio_test_frame(self.module_frames[3])
        
        # 创建射线检测记录续模块的内容
        self.create_radio_renewal_frame(self.module_frames[4])
    
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

        # 第一行参数 - 模板选择
        ray_row0_frame = ttk.Frame(params_grid)
        ray_row0_frame.pack(fill=tk.X, pady=5)

        # 模板选择
        ray_template_label = ttk.Label(ray_row0_frame, text="模板")
        ray_template_label.pack(side=tk.LEFT, padx=(0, 5))
        self.ray_template_var = tk.StringVar()
        self.ray_template_combobox = ttk.Combobox(ray_row0_frame, textvariable=self.ray_template_var,
                                                values=["模板1", "模板2"], state="readonly", width=15)
        self.ray_template_combobox.set("模板2")  # 默认选择模板2
        self.ray_template_combobox.pack(side=tk.LEFT, padx=(0, 20))
        self.ray_template_combobox.bind("<<ComboboxSelected>>", self.on_ray_template_change)

        # 第二行参数
        ray_row1_frame = ttk.Frame(params_grid)
        ray_row1_frame.pack(fill=tk.X, pady=5)

        # 工程名称
        project_label = ttk.Label(ray_row1_frame, text="工程名称")
        project_label.pack(side=tk.LEFT, padx=(0, 5))
        self.ray_project_entry = ttk.Entry(ray_row1_frame, width=40)
        self.ray_project_entry.pack(side=tk.LEFT, padx=(0, 5))

        # 第三行参数
        ray_row2_frame = ttk.Frame(params_grid)
        ray_row2_frame.pack(fill=tk.X, pady=5)

        # 检测类别号 (模板2专用)
        self.ray_category_label = ttk.Label(ray_row2_frame, text="检测类别号")
        self.ray_category_label.pack(side=tk.LEFT, padx=(0, 5))
        self.ray_category_entry = ttk.Entry(ray_row2_frame, width=20)
        self.ray_category_entry.pack(side=tk.LEFT, padx=(0, 30))

        # 委托单位 (模板1专用，初始隐藏)
        self.ray_client_label = ttk.Label(ray_row2_frame, text="委托单位")
        self.ray_client_entry = ttk.Entry(ray_row2_frame, width=20)

        # 检测方法
        method_label = ttk.Label(ray_row2_frame, text="检测方法")
        method_label.pack(side=tk.LEFT, padx=(0, 5))
        self.ray_method_entry = ttk.Entry(ray_row2_frame, width=20)
        self.ray_method_entry.insert(0, "RT")  # 默认值
        self.ray_method_entry.pack(side=tk.LEFT)

        # 第四行参数
        ray_row3_frame = ttk.Frame(params_grid)
        ray_row3_frame.pack(fill=tk.X, pady=5)

        # 检测标准
        standard_label = ttk.Label(ray_row3_frame, text="检测标准")
        standard_label.pack(side=tk.LEFT, padx=(0, 5))
        self.ray_standard_entry = ttk.Entry(ray_row3_frame, width=20)
        self.ray_standard_entry.pack(side=tk.LEFT, padx=(0, 30))

        # 坡口形式
        groove_label = ttk.Label(ray_row3_frame, text="坡口形式")
        groove_label.pack(side=tk.LEFT, padx=(0, 5))
        self.ray_groove_entry = ttk.Entry(ray_row3_frame, width=20)
        self.ray_groove_entry.pack(side=tk.LEFT)

        # 第五行参数 - 模板1专用参数（初始隐藏）
        self.ray_row4_frame = ttk.Frame(params_grid)

        # 验收规范
        self.ray_acceptance_label = ttk.Label(self.ray_row4_frame, text="验收规范")
        self.ray_acceptance_label.pack(side=tk.LEFT, padx=(0, 5))
        self.ray_acceptance_entry = ttk.Entry(self.ray_row4_frame, width=20)
        self.ray_acceptance_entry.pack(side=tk.LEFT, padx=(0, 30))

        # 检测技术等级
        self.ray_tech_level_label = ttk.Label(self.ray_row4_frame, text="检测技术等级")
        self.ray_tech_level_label.pack(side=tk.LEFT, padx=(0, 5))
        self.ray_tech_level_entry = ttk.Entry(self.ray_row4_frame, width=20)
        self.ray_tech_level_entry.pack(side=tk.LEFT)

        # 第六行参数 - 模板1专用参数（初始隐藏）
        self.ray_row5_frame = ttk.Frame(params_grid)

        # 外观检查
        self.ray_appearance_label = ttk.Label(self.ray_row5_frame, text="外观检查")
        self.ray_appearance_label.pack(side=tk.LEFT, padx=(0, 5))
        self.ray_appearance_entry = ttk.Entry(self.ray_row5_frame, width=20)
        self.ray_appearance_entry.pack(side=tk.LEFT)
        
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
        self.ray_word_path.set("生成器/word/1_射线检测委托台账_Mode2.docx")  # 默认值
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

    def on_ray_template_change(self, event=None):
        """射线检测委托台账模板选择变化时的回调函数"""
        selected_template = self.ray_template_var.get()

        if selected_template == "模板1":
            # 显示模板1专用参数
            self.ray_row4_frame.pack(fill=tk.X, pady=5)
            self.ray_row5_frame.pack(fill=tk.X, pady=5)

            # 隐藏检测类别号，显示委托单位
            self.ray_category_label.pack_forget()
            self.ray_category_entry.pack_forget()
            self.ray_client_label.pack(side=tk.LEFT, padx=(0, 5))
            self.ray_client_entry.pack(side=tk.LEFT, padx=(0, 30))

            # 更新Word模板默认路径
            self.ray_word_path.set("生成器/word/1_射线检测委托台账_Mode1.docx")

            print("切换到模板1，显示8个参数：工程名称、委托单位、检测标准、验收规范、检测方法、检测技术等级、外观检查、坡口形式")
        else:
            # 隐藏模板1专用参数
            self.ray_row4_frame.pack_forget()
            self.ray_row5_frame.pack_forget()

            # 显示检测类别号，隐藏委托单位
            self.ray_client_label.pack_forget()
            self.ray_client_entry.pack_forget()
            self.ray_category_label.pack(side=tk.LEFT, padx=(0, 5))
            self.ray_category_entry.pack(side=tk.LEFT, padx=(0, 30))

            # 更新Word模板默认路径
            self.ray_word_path.set("生成器/word/1_射线检测委托台账_Mode2.docx")

            print("切换到模板2，显示5个参数：工程名称、检测类别号、检测标准、检测方法、坡口形式")

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

        # 第一行参数 - 模板选择
        row0_frame = ttk.Frame(params_grid)
        row0_frame.pack(fill=tk.X, pady=5)

        # 模板选择
        template_label = ttk.Label(row0_frame, text="模板")
        template_label.pack(side=tk.LEFT, padx=(0, 5))
        self.template_var = tk.StringVar()
        self.template_combobox = ttk.Combobox(row0_frame, textvariable=self.template_var,
                                            values=["模板1", "模板2"], state="readonly", width=15)
        self.template_combobox.set("模板2")  # 默认选择模板2
        self.template_combobox.pack(side=tk.LEFT, padx=(0, 20))
        self.template_combobox.bind("<<ComboboxSelected>>", self.on_template_change)

        # 第二行参数
        row1_frame = ttk.Frame(params_grid)
        row1_frame.pack(fill=tk.X, pady=5)

        # 工程名称
        project_label = ttk.Label(row1_frame, text="工程名称")
        project_label.pack(side=tk.LEFT, padx=(0, 5))
        self.project_entry = ttk.Entry(row1_frame, width=40)
        self.project_entry.pack(side=tk.LEFT, padx=(0, 5))

        # 第三行参数
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

        # 第四行参数 - 模板1专用参数（初始隐藏）
        self.row3_frame = ttk.Frame(params_grid)

        # 检测单位
        inspection_unit_label = ttk.Label(self.row3_frame, text="检测单位")
        inspection_unit_label.pack(side=tk.LEFT, padx=(0, 5))
        self.inspection_unit_entry = ttk.Entry(self.row3_frame, width=20)
        self.inspection_unit_entry.pack(side=tk.LEFT, padx=(0, 30))

        # 检测标准
        inspection_standard_label = ttk.Label(self.row3_frame, text="检测标准")
        inspection_standard_label.pack(side=tk.LEFT, padx=(0, 5))
        self.inspection_standard_entry = ttk.Entry(self.row3_frame, width=20)
        self.inspection_standard_entry.pack(side=tk.LEFT)
        
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

    def on_template_change(self, event=None):
        """RT结果通知单台账模板选择变化时的回调函数"""
        selected_template = self.template_var.get()

        if selected_template == "模板1":
            # 显示模板1专用参数
            self.row3_frame.pack(fill=tk.X, pady=5)
            print("切换到模板1，显示检测单位和检测标准参数")
        else:
            # 隐藏模板1专用参数
            self.row3_frame.pack_forget()
            print("切换到模板2，隐藏检测单位和检测标准参数")

    def on_surface_template_change(self, event=None):
        """表面结果通知单台账模板选择变化时的回调函数"""
        selected_template = self.surface_template_var.get()

        if selected_template == "模板1":
            # 显示模板1专用参数
            self.surface_row3_frame.pack(fill=tk.X, pady=5)
            # 更新Word模板默认路径
            self.surface_word_path.set("生成器/word/3_表面结果通知单台账_Mode1.docx")
            print("切换到模板1，显示检测单位和检测标准参数")
        else:
            # 隐藏模板1专用参数
            self.surface_row3_frame.pack_forget()
            # 更新Word模板默认路径
            self.surface_word_path.set("生成器/word/3_表面结果通知单台账_Mode2.docx")
            print("切换到模板2，隐藏检测单位和检测标准参数")
    
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

        # 第一行参数 - 模板选择
        surface_row0_frame = ttk.Frame(params_grid)
        surface_row0_frame.pack(fill=tk.X, pady=5)

        # 模板选择
        surface_template_label = ttk.Label(surface_row0_frame, text="模板")
        surface_template_label.pack(side=tk.LEFT, padx=(0, 5))
        self.surface_template_var = tk.StringVar()
        self.surface_template_combobox = ttk.Combobox(surface_row0_frame, textvariable=self.surface_template_var,
                                                    values=["模板1", "模板2"], state="readonly", width=15)
        self.surface_template_combobox.set("模板2")  # 默认选择模板2
        self.surface_template_combobox.pack(side=tk.LEFT, padx=(0, 20))
        self.surface_template_combobox.bind("<<ComboboxSelected>>", self.on_surface_template_change)

        # 第二行参数
        surface_row1_frame = ttk.Frame(params_grid)
        surface_row1_frame.pack(fill=tk.X, pady=5)

        # 工程名称
        project_label = ttk.Label(surface_row1_frame, text="工程名称")
        project_label.pack(side=tk.LEFT, padx=(0, 5))
        self.surface_project_entry = ttk.Entry(surface_row1_frame, width=40)
        self.surface_project_entry.pack(side=tk.LEFT)

        # 第三行参数
        surface_row2_frame = ttk.Frame(params_grid)
        surface_row2_frame.pack(fill=tk.X, pady=5)

        # 委托单位
        client_label = ttk.Label(surface_row2_frame, text="委托单位")
        client_label.pack(side=tk.LEFT, padx=(0, 5))
        self.surface_client_entry = ttk.Entry(surface_row2_frame, width=40)
        self.surface_client_entry.pack(side=tk.LEFT)

        # 第四行参数 - 模板1专用参数（初始隐藏）
        self.surface_row3_frame = ttk.Frame(params_grid)

        # 检测单位
        surface_inspection_unit_label = ttk.Label(self.surface_row3_frame, text="检测单位")
        surface_inspection_unit_label.pack(side=tk.LEFT, padx=(0, 5))
        self.surface_inspection_unit_entry = ttk.Entry(self.surface_row3_frame, width=20)
        self.surface_inspection_unit_entry.pack(side=tk.LEFT, padx=(0, 30))

        # 检测标准
        surface_inspection_standard_label = ttk.Label(self.surface_row3_frame, text="检测标准")
        surface_inspection_standard_label.pack(side=tk.LEFT, padx=(0, 5))
        self.surface_inspection_standard_entry = ttk.Entry(self.surface_row3_frame, width=20)
        self.surface_inspection_standard_entry.pack(side=tk.LEFT)
        
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
        self.surface_word_path.set("生成器/word/3_表面结果通知单台账_Mode2.docx")  # 默认值
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
    
    def create_radio_test_frame(self, parent_frame):
        """创建射线检测记录模块的内容"""
        parent_frame.pack(fill=tk.BOTH, expand=True)
        
        # 模块标题
        header_frame = ttk.Frame(parent_frame)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        header_label = ttk.Label(header_frame, text="射线检测记录", 
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
        self.radio_project_entry = ttk.Entry(row1_frame, width=40)
        self.radio_project_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        # 第二行参数
        row2_frame = ttk.Frame(params_grid)
        row2_frame.pack(fill=tk.X, pady=5)
        
        # 委托单位
        client_label = ttk.Label(row2_frame, text="委托单位")
        client_label.pack(side=tk.LEFT, padx=(0, 5))
        self.radio_client_entry = ttk.Entry(row2_frame, width=20)
        self.radio_client_entry.pack(side=tk.LEFT, padx=(0, 20))
        
        # 操作指导书编号
        guide_label = ttk.Label(row2_frame, text="操作指导书编号")
        guide_label.pack(side=tk.LEFT, padx=(0, 5))
        self.radio_guide_entry = ttk.Entry(row2_frame, width=20)
        self.radio_guide_entry.pack(side=tk.LEFT)
        
        # 第三行参数
        row3_frame = ttk.Frame(params_grid)
        row3_frame.pack(fill=tk.X, pady=5)
        
        # 承包单位
        contract_label = ttk.Label(row3_frame, text="承包单位")
        contract_label.pack(side=tk.LEFT, padx=(0, 5))
        self.radio_contract_entry = ttk.Entry(row3_frame, width=20)
        self.radio_contract_entry.pack(side=tk.LEFT, padx=(0, 20))
        
        # 设备型号
        equipment_label = ttk.Label(row3_frame, text="设备型号")
        equipment_label.pack(side=tk.LEFT, padx=(0, 5))
        self.radio_equipment_entry = ttk.Entry(row3_frame, width=20)
        self.radio_equipment_entry.pack(side=tk.LEFT)
        
        # 文件选择区域
        files_frame = ttk.LabelFrame(parent_frame, text="文件选择")
        files_frame.pack(fill=tk.X, pady=(0, 10), padx=5)
        
        # Excel文件选择
        excel_frame = ttk.Frame(files_frame)
        excel_frame.pack(fill=tk.X, padx=15, pady=(15, 8))
        excel_label = ttk.Label(excel_frame, text="选择输入文件(xlsx)*")
        excel_label.pack(side=tk.LEFT, padx=(0, 5))
        self.radio_excel_path = tk.StringVar()
        self.radio_excel_path.set("生成器/Excel/4_生成器台账-射线检测记录.xlsx")  # 默认值
        excel_entry = ttk.Entry(excel_frame, textvariable=self.radio_excel_path)
        excel_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        excel_button = ttk.Button(excel_frame, text="浏览...", command=self.browse_radio_excel)
        excel_button.pack(side=tk.LEFT)
        
        # Word模板选择
        word_frame = ttk.Frame(files_frame)
        word_frame.pack(fill=tk.X, padx=15, pady=8)
        word_label = ttk.Label(word_frame, text="选择Word模板文件(docx)*")
        word_label.pack(side=tk.LEFT, padx=(0, 5))
        self.radio_word_path = tk.StringVar()
        self.radio_word_path.set("生成器/word/4_射线检测记录.docx")  # 默认值
        word_entry = ttk.Entry(word_frame, textvariable=self.radio_word_path)
        word_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        word_button = ttk.Button(word_frame, text="浏览...", command=self.browse_radio_word)
        word_button.pack(side=tk.LEFT)
        
        # 输出文件夹选择
        output_frame = ttk.Frame(files_frame)
        output_frame.pack(fill=tk.X, padx=15, pady=(8, 15))
        output_label = ttk.Label(output_frame, text="选择输出文件夹*")
        output_label.pack(side=tk.LEFT, padx=(0, 5))
        self.radio_output_path = tk.StringVar()
        output_entry = ttk.Entry(output_frame, textvariable=self.radio_output_path)
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        output_button = ttk.Button(output_frame, text="浏览...", command=self.browse_radio_output)
        output_button.pack(side=tk.LEFT)
        
        # 提交按钮
        submit_frame = ttk.Frame(parent_frame)
        submit_frame.pack(fill=tk.X, pady=10)
        self.radio_submit_button = ttk.Button(submit_frame, text="提交", 
                                        style="Submit.TButton", command=self.process_radio_data)
        self.radio_submit_button.pack(side=tk.RIGHT, padx=10)
        
        # 日志区域
        log_frame = ttk.LabelFrame(parent_frame, text="执行日志")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10), padx=5)
        
        # 创建滚动文本框
        self.radio_log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10)
        self.radio_log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.radio_log_text.configure(state='disabled')
        
        # 设置日志文本字体
        self.radio_log_text.configure(font=(self.default_font, 9))
        
        # 日志操作按钮
        log_buttons_frame = ttk.Frame(log_frame)
        log_buttons_frame.pack(fill=tk.X, padx=5, pady=(0, 5))
        
        export_button = ttk.Button(log_buttons_frame, text="导出", 
                                 style="Action.TButton", command=self.export_radio_log)
        export_button.pack(side=tk.RIGHT, padx=5)
        
        clear_button = ttk.Button(log_buttons_frame, text="清空", 
                                style="Action.TButton", command=self.clear_radio_log)
        clear_button.pack(side=tk.RIGHT, padx=5)
        
        # 设置日志重定向
        self.radio_redirect = RedirectText(self.radio_log_text)
    
    def create_radio_renewal_frame(self, parent_frame):
        """创建射线检测记录续模块的内容"""
        parent_frame.pack(fill=tk.BOTH, expand=True)
        
        # 模块标题
        header_frame = ttk.Frame(parent_frame)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        header_label = ttk.Label(header_frame, text="射线检测记录续", 
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
        self.radio_renewal_project_entry = ttk.Entry(row1_frame, width=40)
        self.radio_renewal_project_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        # 第二行参数
        row2_frame = ttk.Frame(params_grid)
        row2_frame.pack(fill=tk.X, pady=5)
        
        # 委托单位
        client_label = ttk.Label(row2_frame, text="委托单位")
        client_label.pack(side=tk.LEFT, padx=(0, 5))
        self.radio_renewal_client_entry = ttk.Entry(row2_frame, width=20)
        self.radio_renewal_client_entry.pack(side=tk.LEFT, padx=(0, 20))
        
        # 操作指导书编号
        guide_label = ttk.Label(row2_frame, text="操作指导书编号")
        guide_label.pack(side=tk.LEFT, padx=(0, 5))
        self.radio_renewal_guide_entry = ttk.Entry(row2_frame, width=20)
        self.radio_renewal_guide_entry.pack(side=tk.LEFT)
        
        # 文件选择区域
        files_frame = ttk.LabelFrame(parent_frame, text="文件选择")
        files_frame.pack(fill=tk.X, pady=(0, 10), padx=5)
        
        # Excel文件选择
        excel_frame = ttk.Frame(files_frame)
        excel_frame.pack(fill=tk.X, padx=15, pady=(15, 8))
        excel_label = ttk.Label(excel_frame, text="选择输入文件(xlsx)*")
        excel_label.pack(side=tk.LEFT, padx=(0, 5))
        self.radio_renewal_excel_path = tk.StringVar()
        self.radio_renewal_excel_path.set("生成器/Excel/5_生成器台账-射线检测记录续.xlsx")  # 默认值
        excel_entry = ttk.Entry(excel_frame, textvariable=self.radio_renewal_excel_path)
        excel_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        excel_button = ttk.Button(excel_frame, text="浏览...", command=self.browse_radio_renewal_excel)
        excel_button.pack(side=tk.LEFT)
        
        # Word模板选择
        word_frame = ttk.Frame(files_frame)
        word_frame.pack(fill=tk.X, padx=15, pady=8)
        word_label = ttk.Label(word_frame, text="选择Word模板文件(docx)*")
        word_label.pack(side=tk.LEFT, padx=(0, 5))
        self.radio_renewal_word_path = tk.StringVar()
        self.radio_renewal_word_path.set("生成器/word/5_射线检测记录续.docx")  # 默认值
        word_entry = ttk.Entry(word_frame, textvariable=self.radio_renewal_word_path)
        word_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        word_button = ttk.Button(word_frame, text="浏览...", command=self.browse_radio_renewal_word)
        word_button.pack(side=tk.LEFT)
        
        # 输出文件夹选择
        output_frame = ttk.Frame(files_frame)
        output_frame.pack(fill=tk.X, padx=15, pady=(8, 15))
        output_label = ttk.Label(output_frame, text="选择输出文件夹*")
        output_label.pack(side=tk.LEFT, padx=(0, 5))
        self.radio_renewal_output_path = tk.StringVar()
        output_entry = ttk.Entry(output_frame, textvariable=self.radio_renewal_output_path)
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        output_button = ttk.Button(output_frame, text="浏览...", command=self.browse_radio_renewal_output)
        output_button.pack(side=tk.LEFT)
        
        # 提交按钮
        submit_frame = ttk.Frame(parent_frame)
        submit_frame.pack(fill=tk.X, pady=10)
        self.radio_renewal_submit_button = ttk.Button(submit_frame, text="提交", 
                                        style="Submit.TButton", command=self.process_radio_renewal_data)
        self.radio_renewal_submit_button.pack(side=tk.RIGHT, padx=10)
        
        # 日志区域
        log_frame = ttk.LabelFrame(parent_frame, text="执行日志")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10), padx=5)
        
        # 创建滚动文本框
        self.radio_renewal_log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10)
        self.radio_renewal_log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.radio_renewal_log_text.configure(state='disabled')
        
        # 设置日志文本字体
        self.radio_renewal_log_text.configure(font=(self.default_font, 9))
        
        # 日志操作按钮
        log_buttons_frame = ttk.Frame(log_frame)
        log_buttons_frame.pack(fill=tk.X, padx=5, pady=(0, 5))
        
        export_button = ttk.Button(log_buttons_frame, text="导出", 
                                 style="Action.TButton", command=self.export_radio_renewal_log)
        export_button.pack(side=tk.RIGHT, padx=5)
        
        clear_button = ttk.Button(log_buttons_frame, text="清空", 
                                style="Action.TButton", command=self.clear_radio_renewal_log)
        clear_button.pack(side=tk.RIGHT, padx=5)
        
        # 设置日志重定向
        self.radio_renewal_redirect = RedirectText(self.radio_renewal_log_text)
    
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
        selected_template = self.template_var.get()

        # 获取模板1专用参数
        inspection_unit = self.inspection_unit_entry.get() if selected_template == "模板1" else None
        inspection_standard = self.inspection_standard_entry.get() if selected_template == "模板1" else None

        # 验证输入
        if not excel_path or not os.path.exists(excel_path):
            self.show_log("错误: 请选择有效的Excel文件")
            return

        if not word_path or not os.path.exists(word_path):
            self.show_log("错误: 请选择有效的Word模板文件")
            return

        if not output_path:
            # 根据模板类型使用不同的默认输出路径
            if selected_template == "模板1":
                output_path = os.path.join("生成器", "输出报告", "2_RT结果通知单台账_Mode1")
            else:
                output_path = os.path.join("生成器", "输出报告", "2_RT结果通知单台账_Mode2")

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
        self.show_log(f"选择模板: {selected_template}")
        self.show_log(f"Excel文件: {excel_path}")
        self.show_log(f"Word模板: {word_path}")
        self.show_log(f"输出路径: {output_path}")
        self.show_log(f"工程名称: {project_name}")
        self.show_log(f"委托单位: {client_name}")
        self.show_log(f"检测方法: {inspection_method}")
        if selected_template == "模板1":
            self.show_log(f"检测单位: {inspection_unit}")
            self.show_log(f"检测标准: {inspection_standard}")
        self.show_log("="*50)

        # 在后台线程中处理数据
        if selected_template == "模板1":
            threading.Thread(target=self.run_process_mode1, args=(
                excel_path, word_path, output_path, project_name, client_name,
                inspection_unit, inspection_standard, inspection_method
            )).start()
        else:
            threading.Thread(target=self.run_process, args=(
                excel_path, word_path, output_path, project_name, client_name, inspection_method
            )).start()
        
    def run_process(self, excel_path, word_path, output_path, project_name, client_name, inspection_method):
        """在后台线程中运行数据处理 - 模板2"""
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

    def run_process_mode1(self, excel_path, word_path, output_path, project_name, client_name,
                         inspection_unit, inspection_standard, inspection_method):
        """在后台线程中运行数据处理 - 模板1"""
        try:
            # 重定向标准输出到日志区
            with redirect_stdout(self.redirect):
                # 调用NDT_result_mode1模块的处理函数
                success = NDT_result_mode1.process_excel_to_word(
                    excel_path, word_path, output_path, project_name, client_name,
                    inspection_unit, inspection_standard, inspection_method
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
        method = self.ray_method_entry.get()
        standard = self.ray_standard_entry.get()
        groove = self.ray_groove_entry.get()

        # 根据模板选择获取不同的参数
        selected_template = self.ray_template_var.get()

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
        self.show_ray_log(f"选择模板: {selected_template}")
        self.show_ray_log(f"工程名称: {project_name}")
        self.show_ray_log(f"检测标准: {standard}")
        self.show_ray_log(f"检测方法: {method}")
        self.show_ray_log(f"坡口形式: {groove}")

        if selected_template == "模板1":
            # 模板1的8个参数
            client_name = self.ray_client_entry.get()
            acceptance_spec = self.ray_acceptance_entry.get()
            tech_level = self.ray_tech_level_entry.get()
            appearance_check = self.ray_appearance_entry.get()

            self.show_ray_log(f"委托单位: {client_name}")
            self.show_ray_log(f"验收规范: {acceptance_spec}")
            self.show_ray_log(f"检测技术等级: {tech_level}")
            self.show_ray_log(f"外观检查: {appearance_check}")
            self.show_ray_log("="*50)

            # 在后台线程中处理数据
            threading.Thread(target=self.run_ray_mode1_process, args=(
                excel_path, word_path, output_path, project_name, client_name,
                standard, acceptance_spec, method, tech_level, appearance_check, groove
            )).start()
        else:
            # 模板2的5个参数
            category = self.ray_category_entry.get()

            self.show_ray_log(f"检测类别号: {category}")
            self.show_ray_log("="*50)

            # 在后台线程中处理数据
            threading.Thread(target=self.run_ray_mode2_process, args=(
                excel_path, word_path, output_path, project_name, category,
                standard, method, groove
            )).start()

    def run_ray_mode1_process(self, excel_path, word_path, output_path, project_name, client_name,
                            standard, acceptance_spec, method, tech_level, appearance_check, groove):
        """在后台线程中运行射线检测委托台账模板1处理"""
        try:
            # 导入Ray_Detection_mode1模块
            sys.path.append(os.path.dirname(os.path.abspath(__file__)))
            import Ray_Detection_mode1

            # 重定向标准输出到日志区
            with redirect_stdout(self.ray_redirect):
                # 调用Ray_Detection_mode1模块的处理函数
                success = Ray_Detection_mode1.process_excel_to_word(
                    excel_path, word_path, output_path, project_name, client_name,
                    standard, acceptance_spec, method, tech_level, appearance_check, groove
                )

            # 在主线程中更新UI
            self.root.after(0, self.process_ray_completed, success)

        except Exception as e:
            # 在主线程中显示错误
            self.root.after(0, self.show_ray_error, str(e))

    def run_ray_mode2_process(self, excel_path, word_path, output_path, project_name, category,
                            standard, method, groove):
        """在后台线程中运行射线检测委托台账模板2处理"""
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
        selected_template = self.surface_template_var.get()

        # 获取模板1专用参数
        inspection_unit = self.surface_inspection_unit_entry.get() if selected_template == "模板1" else None
        inspection_standard = self.surface_inspection_standard_entry.get() if selected_template == "模板1" else None

        # 验证输入
        if not excel_path or not os.path.exists(excel_path):
            self.show_surface_log("错误: 请选择有效的Excel文件")
            return

        if not word_path or not os.path.exists(word_path):
            self.show_surface_log("错误: 请选择有效的Word模板文件")
            return

        if not output_path:
            # 根据模板类型使用不同的默认输出路径
            if selected_template == "模板1":
                output_path = os.path.join("生成器", "输出报告", "3_表面结果通知单台账", "3_表面结果通知单台账_Mode1")
            else:
                output_path = os.path.join("生成器", "输出报告", "3_表面结果通知单台账", "3_表面结果通知单台账_Mode2")

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
        self.show_surface_log(f"选择模板: {selected_template}")
        self.show_surface_log(f"Excel文件: {excel_path}")
        self.show_surface_log(f"Word模板: {word_path}")
        self.show_surface_log(f"输出路径: {output_path}")
        self.show_surface_log(f"工程名称: {project_name}")
        self.show_surface_log(f"委托单位: {client_name}")
        if selected_template == "模板1":
            self.show_surface_log(f"检测单位: {inspection_unit}")
            self.show_surface_log(f"检测标准: {inspection_standard}")
        self.show_surface_log("="*50)

        # 在后台线程中处理数据
        threading.Thread(target=self.run_surface_process, args=(
            excel_path, word_path, output_path, project_name, client_name,
            selected_template, inspection_unit, inspection_standard
        )).start()

    def run_surface_process(self, excel_path, word_path, output_path, project_name, client_name,
                           selected_template, inspection_unit, inspection_standard):
        """在后台线程中运行表面结果通知单台账处理"""
        try:
            # 根据选择的模板导入不同的模块
            sys.path.append(os.path.dirname(os.path.abspath(__file__)))

            # 重定向标准输出到日志区
            with redirect_stdout(self.surface_redirect):
                if selected_template == "模板1":
                    # 导入Surface_Defect_mode1模块
                    import Surface_Defect_mode1
                    # 调用Surface_Defect_mode1模块的处理函数
                    success = Surface_Defect_mode1.process_excel_to_word(
                        excel_path, word_path, output_path, project_name, client_name,
                        inspection_unit, inspection_standard
                    )
                else:
                    # 导入Surface_Defect模块
                    import Surface_Defect
                    # 调用Surface_Defect模块的处理函数
                    success = Surface_Defect.process_excel_to_word(
                        excel_path, word_path, output_path, project_name, client_name
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

    def browse_radio_excel(self):
        """浏览选择射线检测记录Excel文件"""
        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        if filename:
            self.radio_excel_path.set(filename)
            
    def browse_radio_word(self):
        """浏览选择射线检测记录Word模板文件"""
        filename = filedialog.askopenfilename(
            title="选择Word模板文件",
            filetypes=[("Word文件", "*.docx *.doc")]
        )
        if filename:
            self.radio_word_path.set(filename)
            
    def browse_radio_output(self):
        """浏览选择射线检测记录输出文件夹"""
        directory = filedialog.askdirectory(title="选择输出文件夹")
        if directory:
            self.radio_output_path.set(directory)

    def process_radio_data(self):
        """处理射线检测记录数据"""
        # 获取输入值
        excel_path = self.radio_excel_path.get()
        word_path = self.radio_word_path.get()
        output_path = self.radio_output_path.get()
        project_name = self.radio_project_entry.get()
        client_name = self.radio_client_entry.get()
        guide_number = self.radio_guide_entry.get()
        contract_name = self.radio_contract_entry.get()
        equipment_model = self.radio_equipment_entry.get()
        
        # 验证输入
        if not excel_path or not os.path.exists(excel_path):
            self.show_radio_log("错误: 请选择有效的Excel文件")
            return
        
        if not word_path or not os.path.exists(word_path):
            self.show_radio_log("错误: 请选择有效的Word模板文件")
            return
        
        if not output_path:
            # 使用默认输出路径
            output_path = os.path.join("生成器", "输出报告", "4_射线检测记录")
            self.radio_output_path.set(output_path)
            self.show_radio_log(f"未指定输出文件夹，使用默认路径: {output_path}")
            
            # 确保目录存在
            if not os.path.exists(output_path):
                try:
                    os.makedirs(output_path)
                    self.show_radio_log(f"创建输出目录: {output_path}")
                except Exception as e:
                    self.show_radio_log(f"创建目录失败: {e}")
                    return
        
        # 禁用提交按钮，避免重复提交
        self.radio_submit_button.configure(state='disabled')
        self.status_var.set("状态: 处理中...")
        
        # 显示开始信息
        self.show_radio_log(f"开始处理数据: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.show_radio_log(f"Excel文件: {excel_path}")
        self.show_radio_log(f"Word模板: {word_path}")
        self.show_radio_log(f"输出路径: {output_path}")
        self.show_radio_log(f"工程名称: {project_name}")
        self.show_radio_log(f"委托单位: {client_name}")
        self.show_radio_log(f"操作指导书编号: {guide_number}")
        self.show_radio_log(f"承包单位: {contract_name}")
        self.show_radio_log(f"设备型号: {equipment_model}")
        self.show_radio_log("="*50)
        
        # 在后台线程中处理数据
        threading.Thread(target=self.run_radio_process, args=(
            excel_path, word_path, output_path, project_name, client_name, guide_number, 
            contract_name, equipment_model
        )).start()

    def run_radio_process(self, excel_path, word_path, output_path, project_name, client_name, guide_number, 
                          contract_name, equipment_model):
        """在后台线程中运行射线检测记录处理"""
        try:
            # 导入Radio_test模块
            sys.path.append(os.path.dirname(os.path.abspath(__file__)))
            import Radio_test
            
            # 重定向标准输出到日志区
            with redirect_stdout(self.radio_redirect):
                # 调用Radio_test模块的处理函数
                success = Radio_test.process_excel_to_word(
                    excel_path, word_path, output_path, project_name, client_name, guide_number, 
                    contract_name, equipment_model
                )
            
            # 在主线程中更新UI
            self.root.after(0, self.process_radio_completed, success)
            
        except Exception as e:
            # 在主线程中显示错误
            self.root.after(0, self.show_radio_error, str(e))

    def process_radio_completed(self, success):
        """射线检测记录处理完成后的回调"""
        if success:
            self.status_var.set("状态: 完成")
            self.show_radio_log("\n处理成功完成!")
        else:
            self.status_var.set("状态: 失败")
            self.show_radio_log("\n处理失败!")
            
        # 重新启用提交按钮
        self.radio_submit_button.configure(state='normal')
        
    def show_radio_error(self, error_msg):
        """显示射线检测记录错误信息"""
        self.show_radio_log(f"\n错误: {error_msg}")
        self.status_var.set("状态: 处理出错")
        self.radio_submit_button.configure(state='normal')

    def clear_radio_log(self):
        """清空射线检测记录日志"""
        self.radio_log_text.configure(state='normal')
        self.radio_log_text.delete(1.0, tk.END)
        self.radio_log_text.configure(state='disabled')
    
    def export_radio_log(self):
        """导出射线检测记录日志"""
        filename = filedialog.asksaveasfilename(
            title="导出日志",
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt")]
        )
        if filename:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(self.radio_log_text.get(1.0, tk.END))
            self.show_radio_log(f"日志已导出到: {filename}")
    
    def show_radio_log(self, message):
        """在射线检测记录日志区显示消息"""
        self.radio_log_text.configure(state='normal')
        self.radio_log_text.insert(tk.END, message + "\n")
        self.radio_log_text.see(tk.END)  # 自动滚动到最新内容
        self.radio_log_text.configure(state='disabled')

    def browse_radio_renewal_excel(self):
        """浏览选择射线检测记录续Excel文件"""
        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        if filename:
            self.radio_renewal_excel_path.set(filename)
            
    def browse_radio_renewal_word(self):
        """浏览选择射线检测记录续Word模板文件"""
        filename = filedialog.askopenfilename(
            title="选择Word模板文件",
            filetypes=[("Word文件", "*.docx *.doc")]
        )
        if filename:
            self.radio_renewal_word_path.set(filename)
            
    def browse_radio_renewal_output(self):
        """浏览选择射线检测记录续输出文件夹"""
        directory = filedialog.askdirectory(title="选择输出文件夹")
        if directory:
            self.radio_renewal_output_path.set(directory)

    def process_radio_renewal_data(self):
        """处理射线检测记录续数据"""
        # 获取输入值
        excel_path = self.radio_renewal_excel_path.get()
        word_path = self.radio_renewal_word_path.get()
        output_path = self.radio_renewal_output_path.get()
        project_name = self.radio_renewal_project_entry.get()
        client_name = self.radio_renewal_client_entry.get()
        guide_number = self.radio_renewal_guide_entry.get()
        
        # 验证输入
        if not excel_path or not os.path.exists(excel_path):
            self.show_radio_renewal_log("错误: 请选择有效的Excel文件")
            return
        
        if not word_path or not os.path.exists(word_path):
            self.show_radio_renewal_log("错误: 请选择有效的Word模板文件")
            return
        
        if not output_path:
            # 使用默认输出路径
            output_path = os.path.join("生成器", "输出报告", "5_射线检测记录续")
            self.radio_renewal_output_path.set(output_path)
            self.show_radio_renewal_log(f"未指定输出文件夹，使用默认路径: {output_path}")
            
            # 确保目录存在
            if not os.path.exists(output_path):
                try:
                    os.makedirs(output_path)
                    self.show_radio_renewal_log(f"创建输出目录: {output_path}")
                except Exception as e:
                    self.show_radio_renewal_log(f"创建目录失败: {e}")
                    return
        
        # 禁用提交按钮，避免重复提交
        self.radio_renewal_submit_button.configure(state='disabled')
        self.status_var.set("状态: 处理中...")
        
        # 显示开始信息
        self.show_radio_renewal_log(f"开始处理数据: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.show_radio_renewal_log(f"Excel文件: {excel_path}")
        self.show_radio_renewal_log(f"Word模板: {word_path}")
        self.show_radio_renewal_log(f"输出路径: {output_path}")
        self.show_radio_renewal_log(f"工程名称: {project_name}")
        self.show_radio_renewal_log(f"委托单位: {client_name}")
        self.show_radio_renewal_log(f"操作指导书编号: {guide_number}")
        self.show_radio_renewal_log("="*50)
        
        # 在后台线程中处理数据
        threading.Thread(target=self.run_radio_renewal_process, args=(
            excel_path, word_path, output_path, project_name, client_name, guide_number
        )).start()

    def run_radio_renewal_process(self, excel_path, word_path, output_path, project_name, client_name, guide_number):
        """在后台线程中运行射线检测记录续处理"""
        try:
            # 导入Radio_test_renewal模块
            sys.path.append(os.path.dirname(os.path.abspath(__file__)))
            import Radio_test_renewal
            
            # 重定向标准输出到日志区
            with redirect_stdout(self.radio_renewal_redirect):
                # 调用Radio_test_renewal模块的处理函数
                success = Radio_test_renewal.process_excel_to_word(
                    excel_path, word_path, output_path, project_name, client_name, guide_number
                )
            
            # 在主线程中更新UI
            self.root.after(0, self.process_radio_renewal_completed, success)
            
        except Exception as e:
            # 在主线程中显示错误
            self.root.after(0, self.show_radio_renewal_error, str(e))

    def process_radio_renewal_completed(self, success):
        """射线检测记录续处理完成后的回调"""
        if success:
            self.status_var.set("状态: 完成")
            self.show_radio_renewal_log("\n处理成功完成!")
        else:
            self.status_var.set("状态: 失败")
            self.show_radio_renewal_log("\n处理失败!")
            
        # 重新启用提交按钮
        self.radio_renewal_submit_button.configure(state='normal')
        
    def show_radio_renewal_error(self, error_msg):
        """显示射线检测记录续错误信息"""
        self.show_radio_renewal_log(f"\n错误: {error_msg}")
        self.status_var.set("状态: 处理出错")
        self.radio_renewal_submit_button.configure(state='normal')

    def clear_radio_renewal_log(self):
        """清空射线检测记录续日志"""
        self.radio_renewal_log_text.configure(state='normal')
        self.radio_renewal_log_text.delete(1.0, tk.END)
        self.radio_renewal_log_text.configure(state='disabled')
    
    def export_radio_renewal_log(self):
        """导出射线检测记录续日志"""
        filename = filedialog.asksaveasfilename(
            title="导出日志",
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt")]
        )
        if filename:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(self.radio_renewal_log_text.get(1.0, tk.END))
            self.show_radio_renewal_log(f"日志已导出到: {filename}")
    
    def show_radio_renewal_log(self, message):
        """在射线检测记录续日志区显示消息"""
        self.radio_renewal_log_text.configure(state='normal')
        self.radio_renewal_log_text.insert(tk.END, message + "\n")
        self.radio_renewal_log_text.see(tk.END)  # 自动滚动到最新内容
        self.radio_renewal_log_text.configure(state='disabled')

if __name__ == "__main__":
    root = tk.Tk()
    app = NDTResultGUI(root)
    root.mainloop()