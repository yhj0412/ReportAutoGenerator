#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
创建完整的分发包
"""

import os
import shutil
import zipfile
from datetime import datetime

def show_existing_distributions():
    """显示现有的分发包信息"""
    if not os.path.exists("分发包"):
        return

    # 获取所有分发包目录和压缩包
    distributions = []
    zip_files = []

    for item in os.listdir("分发包"):
        item_path = os.path.join("分发包", item)
        if os.path.isdir(item_path) and item.startswith("NDT结果生成器"):
            # 获取目录信息
            size = get_directory_size(item_path)
            mtime = os.path.getmtime(item_path)
            distributions.append({
                'name': item,
                'type': '目录',
                'size': size,
                'mtime': mtime,
                'path': item_path
            })
        elif item.endswith('.zip') and item.startswith("NDT结果生成器"):
            # 获取压缩包信息
            size = os.path.getsize(item_path)
            mtime = os.path.getmtime(item_path)
            zip_files.append({
                'name': item,
                'type': '压缩包',
                'size': size,
                'mtime': mtime,
                'path': item_path
            })

    # 合并并按时间排序
    all_items = distributions + zip_files
    all_items.sort(key=lambda x: x['mtime'], reverse=True)

    if all_items:
        print("\n📁 现有分发包:")
        print("-" * 80)
        print(f"{'名称':<40} {'类型':<8} {'大小':<12} {'修改时间':<20}")
        print("-" * 80)

        for item in all_items[:10]:  # 只显示最新的10个
            size_str = format_size(item['size'])
            mtime_str = datetime.fromtimestamp(item['mtime']).strftime('%Y-%m-%d %H:%M:%S')
            print(f"{item['name']:<40} {item['type']:<8} {size_str:<12} {mtime_str:<20}")

        if len(all_items) > 10:
            print(f"... 还有 {len(all_items) - 10} 个历史版本")

        print("-" * 80)
        print(f"总计: {len(distributions)} 个目录, {len(zip_files)} 个压缩包")
        print()

def get_directory_size(directory):
    """计算目录大小"""
    total_size = 0
    try:
        for dirpath, dirnames, filenames in os.walk(directory):
            for filename in filenames:
                filepath = os.path.join(dirpath, filename)
                if os.path.exists(filepath):
                    total_size += os.path.getsize(filepath)
    except:
        pass
    return total_size

def format_size(size_bytes):
    """格式化文件大小"""
    if size_bytes == 0:
        return "0 B"

    size_names = ["B", "KB", "MB", "GB"]
    i = 0
    while size_bytes >= 1024 and i < len(size_names) - 1:
        size_bytes /= 1024.0
        i += 1

    return f"{size_bytes:.1f} {size_names[i]}"

def create_distribution():
    """创建完整的分发包"""
    print("创建NDT结果生成器分发包")
    print("=" * 50)

    # 创建带时间戳的唯一分发目录名
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    dist_name = f"NDT结果生成器_v1.0_{timestamp}"
    dist_dir = f"分发包/{dist_name}"

    # 检查是否存在同名目录，如果存在则添加序号
    counter = 1
    original_dist_name = dist_name

    while os.path.exists(dist_dir):
        dist_name = f"{original_dist_name}_{counter:02d}"
        dist_dir = f"分发包/{dist_name}"
        counter += 1
        if counter > 99:  # 防止无限循环
            print("⚠ 警告: 同名目录过多，请手动清理")
            break

    # 创建分发包根目录（如果不存在）
    os.makedirs("分发包", exist_ok=True)

    # 创建具体的分发目录
    os.makedirs(dist_dir, exist_ok=True)
    print(f"✓ 创建分发目录: {dist_dir}")

    # 显示历史分发包信息
    show_existing_distributions()
    
    # 复制exe文件
    if os.path.exists("dist/NDT结果生成器.exe"):
        shutil.copy2("dist/NDT结果生成器.exe", dist_dir)
        print("✓ 复制主程序")
    else:
        print("✗ 未找到exe文件，请先运行打包")
        return False
    
    # 复制生成器文件夹
    if os.path.exists("生成器"):
        # 先清理临时文件
        print("清理临时文件...")
        for root, dirs, files in os.walk("生成器"):
            for file in files:
                if file.startswith("~$"):
                    file_path = os.path.join(root, file)
                    try:
                        # 尝试修改文件属性后删除
                        os.chmod(file_path, 0o777)
                        os.remove(file_path)
                        print(f"✓ 清理临时文件: {file}")
                    except Exception as e:
                        print(f"⚠ 无法删除临时文件 {file}: {e}")
                        pass
        
        dest_generator_dir = os.path.join(dist_dir, "生成器")

        # 自定义复制函数，跳过临时文件
        def ignore_temp_files(dir, files):
            return [f for f in files if f.startswith("~$")]

        try:
            shutil.copytree("生成器", dest_generator_dir, ignore=ignore_temp_files, dirs_exist_ok=True)
            print("✓ 复制模板文件夹")
        except Exception as e:
            print(f"❌ 复制模板文件夹失败: {e}")
            return False
    else:
        print("✗ 未找到生成器文件夹")
        return False

    # 复制PDF文档
    pdf_files = [
        "NDT结果生成器用户操作手册.pdf",
        "NDT结果生成器系统部署指南.pdf",
        "NDT结果生成器系统需求文档.pdf",
        "NDT结果生成器系统详细设计文档.pdf"
    ]

    # 创建文档目录
    docs_dir = os.path.join(dist_dir, "文档")
    os.makedirs(docs_dir, exist_ok=True)

    copied_pdfs = []
    for pdf_file in pdf_files:
        if os.path.exists(pdf_file):
            shutil.copy2(pdf_file, docs_dir)
            copied_pdfs.append(pdf_file)
            print(f"✓ 复制文档: {pdf_file}")
        else:
            print(f"⚠ 未找到文档: {pdf_file}")

    if copied_pdfs:
        print(f"✓ 共复制 {len(copied_pdfs)} 个PDF文档到文档目录")
    else:
        print("⚠ 未找到任何PDF文档")

    # 创建详细的使用说明
    readme_content = f"""# NDT结果生成器 v1.0

## 📋 软件介绍
NDT结果生成器是一个专业的无损检测报告生成工具，支持多种检测类型的报告自动化生成。

## 🚀 快速开始
1. 双击 `NDT结果生成器.exe` 启动程序
2. 选择对应的功能模块
3. 填写必要参数
4. 选择输入文件和输出路径
5. 点击提交生成报告

## 📁 目录结构
```
NDT结果生成器/
├── NDT结果生成器.exe          # 主程序
├── 生成器/                    # 模板文件夹
│   ├── Excel/                # Excel模板
│   ├── word/                 # Word模板
│   ├── wod/                  # Word模板
│   └── 输出报告/             # 默认输出目录
├── 文档/                      # 系统文档
│   ├── NDT结果生成器用户操作手册.pdf
│   ├── NDT结果生成器系统部署指南.pdf
│   ├── NDT结果生成器系统需求文档.pdf
│   └── NDT结果生成器系统详细设计文档.pdf
├── 启动NDT结果生成器.bat      # 启动脚本
├── 使用说明.txt              # 简要说明
└── README.md                 # 详细说明（本文件）
```

## 🔧 功能模块

### 1. 射线检测委托台账
- **模板1**: 8个参数（工程名称、委托单位、检测标准、验收规范、检测方法、检测技术等级、外观检查、坡口形式）
- **模板2**: 5个参数（工程名称、检测类别号、检测标准、检测方法、坡口形式）

### 2. RT结果通知单台账
- **模板1**: 5个参数（工程名称、委托单位、检测单位、检测方法、检测标准）
- **模板2**: 3个参数（工程名称、委托单位、检测方法）

### 3. 表面结果通知单台账
- **模板1**: 4个参数（工程名称、委托单位、检测单位、检测标准）
- **模板2**: 2个参数（工程名称、委托单位）

### 4. 射线检测记录
- 射线检测记录生成功能

### 5. 射线检测记录续
- 射线检测记录续表生成功能

## 📚 系统文档
分发包中包含完整的系统文档，位于 `文档/` 目录：

### 用户文档
- **用户操作手册**: 详细的软件使用指南，包含每个功能模块的操作步骤
- **系统部署指南**: 软件安装、配置和部署的详细说明

### 技术文档
- **系统需求文档**: 软件功能需求和技术规格说明
- **系统详细设计文档**: 软件架构设计和技术实现细节

建议用户在使用前先阅读用户操作手册，技术人员可参考技术文档了解系统实现。

## ⚙️ 系统要求
- Windows 7 或更高版本
- 至少 2GB 内存
- 100MB 可用磁盘空间
- 不需要安装Python环境

## 📝 使用注意事项

### 文件路径
- 确保Excel输入文件格式正确
- Word模板文件路径正确
- 输出目录有写入权限

### 数据格式
- Excel文件应包含必要的列（委托单编号、完成日期等）
- 数据格式应符合模板要求

### 权限设置
- 首次运行可能被杀毒软件拦截，请添加到白名单
- 确保对输出目录有写入权限

## 🔍 故障排除

### 程序无法启动
1. 检查是否被杀毒软件拦截
2. 确保系统满足最低要求
3. 尝试以管理员身份运行

### 找不到模板文件
1. 确保"生成器"文件夹与exe文件在同一目录
2. 检查模板文件是否完整
3. 重新下载完整安装包

### 生成失败
1. 检查Excel文件格式是否正确
2. 确保所有必填参数已填写
3. 检查输出目录权限
4. 查看日志区域的错误信息

## 📞 技术支持
如遇到问题，请联系1594445261@qq.com技术支持并提供：
- 错误截图
- 输入文件样例
- 详细操作步骤

## 📄 版本信息
- 版本: v1.0
- 构建日期: {datetime.now().strftime('%Y-%m-%d')}
- 支持的文件格式: .xlsx, .docx

## 🔄 更新日志
### v1.0 ({datetime.now().strftime('%Y-%m-%d')})
- 初始版本发布
- 支持5个主要功能模块
- 支持模板选择功能
- 支持检测级别值自动映射
- 完整的GUI界面
- 详细的日志输出

---
© 2024 NDT结果生成器. 保留所有权利。
"""
    
    with open(os.path.join(dist_dir, "README.md"), 'w', encoding='utf-8') as f:
        f.write(readme_content)
    print("✓ 创建详细说明文档")
    
    # 复制使用说明
    if os.path.exists("dist/使用说明.txt"):
        shutil.copy2("dist/使用说明.txt", dist_dir)
        print("✓ 复制使用说明")
    
    # 创建启动脚本
    batch_content = f"""@echo off
chcp 65001 > nul
title NDT结果生成器 v1.0
echo.
echo ========================================
echo   NDT结果生成器 v1.0
echo   构建日期: {datetime.now().strftime('%Y-%m-%d')}
echo ========================================
echo.
echo 正在启动程序...
echo.

"NDT结果生成器.exe"

if errorlevel 1 (
    echo.
    echo 程序异常退出！
    echo 可能的原因：
    echo 1. 被杀毒软件拦截
    echo 2. 缺少必要文件
    echo 3. 权限不足
    echo.
    echo 请检查上述问题后重试
    echo.
    pause
)
"""
    
    with open(os.path.join(dist_dir, "启动NDT结果生成器.bat"), 'w', encoding='utf-8') as f:
        f.write(batch_content)
    print("✓ 创建启动脚本")
    
    # 创建压缩包
    zip_path = f"分发包/{dist_name}.zip"
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(dist_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arc_path = os.path.relpath(file_path, "分发包")
                zipf.write(file_path, arc_path)
    
    print(f"✓ 创建压缩包: {zip_path}")
    
    # 显示文件大小
    exe_size = os.path.getsize(os.path.join(dist_dir, "NDT结果生成器.exe")) / (1024*1024)
    zip_size = os.path.getsize(zip_path) / (1024*1024)
    
    print("\n" + "=" * 50)
    print("✓ 分发包创建完成!")
    print(f"程序大小: {exe_size:.1f} MB")
    print(f"压缩包大小: {zip_size:.1f} MB")
    print(f"分发目录: {dist_dir}")
    print(f"压缩包: {zip_path}")
    print("=" * 50)
    
    return True

if __name__ == "__main__":
    create_distribution()
