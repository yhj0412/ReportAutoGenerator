@echo off
chcp 65001 > nul
echo NDT结果生成器 - 打包工具
echo ================================

echo 检查Python环境...
python --version
if errorlevel 1 (
    echo 错误: 未找到Python环境
    echo 请确保已安装Python并添加到PATH环境变量
    pause
    exit /b 1
)

echo.
echo 安装依赖...
pip install -r requirements.txt
if errorlevel 1 (
    echo 警告: 依赖安装可能有问题，继续尝试打包...
)

echo.
echo 开始打包...
python build_simple.py

echo.
echo 打包完成！
echo 可执行文件位置: dist\NDT结果生成器.exe
echo.
pause
