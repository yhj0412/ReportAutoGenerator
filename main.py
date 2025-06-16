import os
import sys
import logging
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import webbrowser
from http.server import HTTPServer, SimpleHTTPRequestHandler
import json
from urllib.parse import parse_qs, urlparse
import socket
import time
import queue

# 导入报告生成器
from report_generator import generate_delegation_reports

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%H:%M:%S'
)
logger = logging.getLogger(__name__)

# 获取一个可用的端口
def get_free_port():
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.bind(('', 0))
    port = s.getsockname()[1]
    s.close()
    return port

# 文件浏览请求队列
file_browse_queue = queue.Queue()
file_response_queue = queue.Queue()

# 自定义HTTP请求处理器
class ReportGeneratorHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, log_queue=None, **kwargs):
        self.log_queue = log_queue
        super().__init__(*args, **kwargs)
    
    def log_message(self, format, *args):
        # 覆盖默认的日志方法，避免在控制台输出访问日志
        pass
    
    def do_GET(self):
        """处理GET请求"""
        # 提供静态文件服务
        if self.path == '/':
            self.path = '/ui_design.html'
        
        # 调用父类方法处理静态文件
        return SimpleHTTPRequestHandler.do_GET(self)
    
    def do_POST(self):
        """处理POST请求"""
        if self.path == '/api/generate':
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length).decode('utf-8')
            params = json.loads(post_data)
            
            # 解析参数
            module = params.get('module', '')
            excel_file = params.get('excel_file', '')
            word_template = params.get('word_template', '')
            output_dir = params.get('output_dir', '')
            
            # 验证参数
            if not excel_file or not word_template or not output_dir:
                self.send_error(400, "Missing required parameters")
                return
            
            # 定义日志回调函数
            def log_callback(message):
                if self.log_queue:
                    self.log_queue.append(message)
            
            # 处理不同模块的生成请求
            result = False
            if module == 'module1':  # 射线检测委托台账
                result = generate_delegation_reports(excel_file, word_template, output_dir, log_callback)
            # TODO: 添加其他模块的处理
            
            # 返回结果
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            response = {'success': result}
            self.wfile.write(json.dumps(response).encode())
        
        elif self.path == '/api/browse-file':
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length).decode('utf-8')
            params = json.loads(post_data)
            
            file_type = params.get('type', 'excel')
            
            # 将请求放入队列
            request_id = int(time.time() * 1000)  # 使用时间戳作为请求ID
            file_browse_queue.put((request_id, file_type))
            logger.info(f"已将文件浏览请求 {request_id} 放入队列")
            
            # 等待响应
            timeout = 60  # 60秒超时
            start_time = time.time()
            response_path = None
            
            while time.time() - start_time < timeout:
                try:
                    # 尝试从响应队列获取结果，不阻塞
                    resp_id, path = file_response_queue.get(block=False)
                    if resp_id == request_id:
                        response_path = path
                        logger.info(f"已收到文件浏览响应 {resp_id}")
                        break
                    else:
                        # 如果不是当前请求的响应，放回队列
                        file_response_queue.put((resp_id, path))
                except queue.Empty:
                    # 队列为空，等待一段时间再尝试
                    time.sleep(0.1)
            
            if response_path is None:
                logger.error(f"文件浏览请求 {request_id} 超时")
            
            # 返回结果
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            response = {'path': response_path if response_path else ""}
            self.wfile.write(json.dumps(response).encode())
        
        elif self.path == '/api/get-logs':
            # 返回日志
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            logs = self.log_queue.copy() if self.log_queue else []
            # 清空日志队列
            if self.log_queue:
                self.log_queue.clear()
            response = {'logs': logs}
            self.wfile.write(json.dumps(response).encode())
        
        else:
            self.send_error(404, "API endpoint not found")

def file_browse_worker():
    """文件浏览工作线程"""
    logger.info("文件浏览工作线程已启动")
    
    while True:
        try:
            # 从请求队列获取文件浏览请求
            request_id, file_type = file_browse_queue.get()
            logger.info(f"处理文件浏览请求 {request_id}, 类型: {file_type}")
            
            # 创建一个新的Tkinter根窗口用于文件对话框
            root = tk.Tk()
            root.withdraw()  # 隐藏主窗口
            
            # 将窗口提升到前面
            root.attributes('-topmost', True)
            root.update()
            
            # 在主线程中打开文件对话框
            if file_type == 'excel':
                file_path = filedialog.askopenfilename(
                    title="选择Excel文件",
                    filetypes=[("Excel文件", "*.xlsx;*.xls")],
                    parent=root  # 指定父窗口
                )
            elif file_type == 'word':
                file_path = filedialog.askopenfilename(
                    title="选择Word模板",
                    filetypes=[("Word文件", "*.docx")],
                    parent=root  # 指定父窗口
                )
            elif file_type == 'dir':
                file_path = filedialog.askdirectory(
                    title="选择输出目录",
                    parent=root  # 指定父窗口
                )
            else:
                file_path = ""
            
            # 销毁根窗口
            root.destroy()
            
            # 将结果放入响应队列
            file_response_queue.put((request_id, file_path))
            logger.info(f"文件浏览请求 {request_id} 处理完成: {file_path}")
            
        except Exception as e:
            logger.error(f"文件浏览工作线程出错: {str(e)}")
        
        time.sleep(0.1)

def run_server(port, log_queue):
    """启动HTTP服务器"""
    # 创建自定义的请求处理器类
    handler_class = lambda *args, **kwargs: ReportGeneratorHandler(*args, log_queue=log_queue, **kwargs)
    
    # 设置服务器地址和端口
    server_address = ('', port)
    
    # 创建HTTP服务器
    httpd = HTTPServer(server_address, handler_class)
    
    logger.info(f"服务器启动在 http://localhost:{port}")
    
    # 运行服务器
    httpd.serve_forever()

def main():
    """主函数"""
    # 获取可用端口
    port = get_free_port()
    
    # 创建日志队列
    log_queue = []
    
    # 创建并启动服务器线程
    server_thread = threading.Thread(target=run_server, args=(port, log_queue))
    server_thread.daemon = True
    server_thread.start()
    
    # 打开浏览器
    webbrowser.open(f"http://localhost:{port}/ui_design.html")
    
    # 在主线程中运行文件浏览工作线程
    # 这确保文件对话框可以正常显示
    file_browse_worker()

if __name__ == "__main__":
    # 初始化tkinter
    # 注意: 不要在主程序创建和销毁根窗口，这可能导致问题
    # 每次文件浏览请求时创建新的根窗口
    
    main() 