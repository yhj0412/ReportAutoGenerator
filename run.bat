@echo off
echo 正在启动Excel到Word数据填充工具...
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未检测到Python安装，请先安装Python 3.6或更高版本。
    echo 您可以从 https://www.python.org/downloads/ 下载安装。
    pause
    exit /b 1
)

REM 检查是否已安装依赖
echo 正在检查依赖...
python -c "import pandas; import docx" >nul 2>&1
if %errorlevel% neq 0 (
    echo 正在安装依赖包...
    pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo 安装依赖包失败，请检查网络连接或手动安装。
        pause
        exit /b 1
    )
)

echo 所有依赖已就绪，正在启动程序...
echo.
echo 注意: 当点击"浏览..."按钮后，如果文件选择窗口未显示在前台，
echo 请查看任务栏或切换程序窗口(按Alt+Tab)以找到文件选择对话框。
echo.

python main.py

if %errorlevel% neq 0 (
    echo 程序运行出错，请查看上面的错误信息。
    pause
)

exit /b 0 