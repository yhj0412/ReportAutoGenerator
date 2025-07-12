#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
最小化打包脚本 - 只打包程序，不包含模板文件
"""

import os
import sys
import subprocess
import shutil

def main():
    print("NDT结果生成器 - 最小化打包")
    print("=" * 50)
    
    # 清理之前的构建
    if os.path.exists('build'):
        shutil.rmtree('build')
        print("✓ 已清理build目录")
    
    if os.path.exists('dist'):
        shutil.rmtree('dist')
        print("✓ 已清理dist目录")
    
    # 检查PyInstaller
    try:
        import PyInstaller
        print("✓ PyInstaller已安装")
    except ImportError:
        print("安装PyInstaller...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    # 使用PyInstaller打包 - 最小化配置
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--onefile',  # 打包成单个文件
        '--windowed',  # 不显示控制台
        '--name=NDT结果生成器',
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
    print("注意：此版本不包含模板文件，用户需要手动提供模板")
    
    result = subprocess.run(cmd)
    
    if result.returncode == 0:
        print("\n✓ 打包成功!")
        print("可执行文件: dist/NDT结果生成器.exe")
        
        # 创建使用说明
        readme_content = """# NDT结果生成器 使用说明

## 重要提示
此版本为最小化打包版本，不包含模板文件。

## 使用前准备
1. 将exe文件放在项目根目录
2. 确保"生成器"文件夹与exe文件在同一目录
3. 生成器文件夹应包含：
   - Excel/ (Excel模板文件夹)
   - word/ (Word模板文件夹)
   - wod/ (Word模板文件夹)

## 目录结构示例
```
NDT结果生成器.exe
生成器/
├── Excel/
│   ├── 1_生成器委托.xlsx
│   ├── 2_生成器结果.xlsx
│   └── 3_生成器表面结果.xlsx
├── word/
│   └── (Word模板文件)
└── wod/
    └── (Word模板文件)
```

## 故障排除
- 如果提示找不到模板文件，请检查目录结构
- 确保所有模板文件路径正确
- 运行时请给予程序读写权限
"""
        
        with open('dist/使用说明.txt', 'w', encoding='utf-8') as f:
            f.write(readme_content)
        
        print("✓ 已创建使用说明.txt")
        
    else:
        print("\n✗ 打包失败!")
    
    return result.returncode == 0

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
