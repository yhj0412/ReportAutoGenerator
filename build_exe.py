#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
NDT结果生成器 - 打包脚本
使用PyInstaller将项目打包成exe文件
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

def check_dependencies():
    """检查必要的依赖是否已安装"""
    print("检查依赖...")
    
    try:
        import pandas
        import docx
        import openpyxl
        print("✓ 核心依赖已安装")
    except ImportError as e:
        print(f"✗ 缺少依赖: {e}")
        print("请运行: pip install -r requirements.txt")
        return False
    
    try:
        import PyInstaller
        print("✓ PyInstaller已安装")
    except ImportError:
        print("✗ PyInstaller未安装")
        print("正在安装PyInstaller...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller==6.3.0"])
    
    return True

def clean_build_dirs():
    """清理之前的构建目录"""
    print("清理构建目录...")
    
    dirs_to_clean = ['build', 'dist', '__pycache__']
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"✓ 已清理 {dir_name}")
    
    # 清理.spec文件
    spec_files = [f for f in os.listdir('.') if f.endswith('.spec')]
    for spec_file in spec_files:
        os.remove(spec_file)
        print(f"✓ 已清理 {spec_file}")

def create_pyinstaller_spec():
    """创建PyInstaller配置文件"""
    print("创建PyInstaller配置文件...")
    
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# 分析主程序
a = Analysis(
    ['gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('生成器', '생성기'),  # 包含模板文件夹
        ('requirements.txt', '.'),
    ],
    hiddenimports=[
        'pandas',
        'openpyxl',
        'docx',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.scrolledtext',
        'tkinter.font',
        'threading',
        'datetime',
        'io',
        'contextlib',
        'NDT_result',
        'NDT_result_mode1',
        'Surface_Defect',
        'Surface_Defect_mode1',
        'Ray_Detection',
        'Ray_Detection_mode1',
        'Radio_test',
        'Radio_test_renewal',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# 收集所有文件
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# 创建可执行文件
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='NDT结果生成器',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 不显示控制台窗口
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 可以添加图标文件路径
)
'''
    
    with open('NDT结果生成器.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)
    
    print("✓ 已创建 NDT结果生成器.spec")

def build_executable():
    """构建可执行文件"""
    print("开始构建可执行文件...")
    
    # 使用PyInstaller构建
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--clean',
        '--noconfirm',
        'NDT结果生成器.spec'
    ]
    
    print(f"执行命令: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
    
    if result.returncode == 0:
        print("✓ 构建成功!")
        return True
    else:
        print("✗ 构建失败!")
        print("错误输出:")
        print(result.stderr)
        return False

def copy_additional_files():
    """复制额外的文件到dist目录"""
    print("复制额外文件...")
    
    dist_dir = Path('dist')
    if not dist_dir.exists():
        print("✗ dist目录不存在")
        return False
    
    # 复制生成器文件夹
    src_generator = Path('生成器')
    dst_generator = dist_dir / '생성기'
    
    if src_generator.exists():
        if dst_generator.exists():
            shutil.rmtree(dst_generator)
        shutil.copytree(src_generator, dst_generator)
        print("✓ 已复制生成器文件夹")
    
    # 复制README文件
    readme_files = ['README.md', '需求.md', '需求文档_详细版.md']
    for readme in readme_files:
        if os.path.exists(readme):
            shutil.copy2(readme, dist_dir)
            print(f"✓ 已复制 {readme}")
    
    return True

def create_batch_file():
    """创建启动批处理文件"""
    print("创建启动脚本...")
    
    batch_content = '''@echo off
chcp 65001 > nul
echo 启动NDT结果生成器...
echo.
"NDT结果生成器.exe"
if errorlevel 1 (
    echo.
    echo 程序异常退出，按任意键关闭窗口...
    pause > nul
)
'''
    
    with open('dist/启动NDT结果生成器.bat', 'w', encoding='utf-8') as f:
        f.write(batch_content)
    
    print("✓ 已创建启动脚本")

def main():
    """主函数"""
    print("=" * 60)
    print("NDT结果生成器 - 打包工具")
    print("=" * 60)
    
    # 检查依赖
    if not check_dependencies():
        return False
    
    # 清理构建目录
    clean_build_dirs()
    
    # 创建配置文件
    create_pyinstaller_spec()
    
    # 构建可执行文件
    if not build_executable():
        return False
    
    # 复制额外文件
    if not copy_additional_files():
        return False
    
    # 创建启动脚本
    create_batch_file()
    
    print("\n" + "=" * 60)
    print("✓ 打包完成!")
    print("可执行文件位置: dist/NDT结果生成器.exe")
    print("启动脚本位置: dist/启动NDT结果生成器.bat")
    print("=" * 60)
    
    return True

if __name__ == "__main__":
    success = main()
    if not success:
        print("\n打包失败，请检查错误信息")
        sys.exit(1)
    else:
        print("\n打包成功，可以分发dist文件夹中的内容")
