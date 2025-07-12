#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简化的打包脚本
"""

import os
import sys
import subprocess

def main():
    print("NDT结果生成器 - 简化打包")
    print("=" * 50)
    
    # 检查PyInstaller
    try:
        import PyInstaller
        print("✓ PyInstaller已安装")
    except ImportError:
        print("安装PyInstaller...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    # 使用PyInstaller打包
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--onefile',  # 打包成单个文件
        '--windowed',  # 不显示控制台
        '--name=NDT结果生成器',
        '--add-data=生成器;生成器',  # 添加生成器文件夹
        '--hidden-import=pandas',
        '--hidden-import=openpyxl',
        '--hidden-import=docx',
        '--hidden-import=NDT_result',
        '--hidden-import=NDT_result_mode1',
        '--hidden-import=Surface_Defect',
        '--hidden-import=Surface_Defect_mode1',
        '--hidden-import=Ray_Detection',
        '--hidden-import=Ray_Detection_mode1',
        '--hidden-import=Radio_test',
        '--hidden-import=Radio_test_renewal',
        'gui.py'
    ]
    
    print("开始打包...")
    print(f"命令: {' '.join(cmd)}")
    
    result = subprocess.run(cmd)
    
    if result.returncode == 0:
        print("\n✓ 打包成功!")
        print("可执行文件: dist/NDT结果生成器.exe")
    else:
        print("\n✗ 打包失败!")
    
    return result.returncode == 0

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
