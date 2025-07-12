#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
检查Git文件状态
验证.gitignore文件的效果，显示哪些文件会被Git跟踪
"""

import os
import subprocess
import glob

def check_git_status():
    """检查Git状态和.gitignore效果"""
    print("=== Git文件状态检查 ===\n")
    
    # 检查是否在Git仓库中
    if not os.path.exists('.git'):
        print("⚠️ 当前目录不是Git仓库")
        print("请先运行: git init")
        return False
    
    # 检查.gitignore文件
    if not os.path.exists('.gitignore'):
        print("❌ 未找到.gitignore文件")
        return False
    
    print("✅ 找到.gitignore文件")
    
    # 显示.gitignore的主要规则
    print("\n📋 .gitignore主要规则:")
    print("-" * 40)
    
    with open('.gitignore', 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    # 提取主要的排除规则
    important_rules = []
    for line in lines:
        line = line.strip()
        if line and not line.startswith('#') and not line.startswith('='):
            important_rules.append(line)
    
    # 显示前20个重要规则
    for rule in important_rules[:20]:
        print(f"  - {rule}")
    
    if len(important_rules) > 20:
        print(f"  ... 还有 {len(important_rules) - 20} 个规则")
    
    return True

def analyze_project_files():
    """分析项目文件分类"""
    print("\n📁 项目文件分析:")
    print("-" * 50)
    
    # 文件分类
    categories = {
        'Python源码': [],
        'Markdown文档': [],
        'PDF文档': [],
        'Excel模板': [],
        'Word模板': [],
        '配置文件': [],
        '临时文件': [],
        '构建产物': [],
        '其他文件': []
    }
    
    # 遍历所有文件
    for root, dirs, files in os.walk('.'):
        # 跳过隐藏目录和Git目录
        dirs[:] = [d for d in dirs if not d.startswith('.')]
        
        for file in files:
            if file.startswith('.'):
                continue
                
            file_path = os.path.join(root, file)
            relative_path = os.path.relpath(file_path, '.')
            
            # 分类文件
            if file.endswith('.py'):
                categories['Python源码'].append(relative_path)
            elif file.endswith('.md'):
                categories['Markdown文档'].append(relative_path)
            elif file.endswith('.pdf'):
                categories['PDF文档'].append(relative_path)
            elif file.endswith('.xlsx') or file.endswith('.xls'):
                categories['Excel模板'].append(relative_path)
            elif file.endswith('.docx') or file.endswith('.doc'):
                categories['Word模板'].append(relative_path)
            elif file in ['requirements.txt', '.gitignore'] or file.endswith('.spec'):
                categories['配置文件'].append(relative_path)
            elif (file.startswith('~$') or file.endswith('.tmp') or 
                  file.endswith('.pyc') or '__pycache__' in relative_path):
                categories['临时文件'].append(relative_path)
            elif ('build' in relative_path or 'dist' in relative_path or 
                  '分发包' in relative_path or file.endswith('.exe')):
                categories['构建产物'].append(relative_path)
            else:
                categories['其他文件'].append(relative_path)
    
    # 显示分类结果
    for category, files in categories.items():
        if files:
            print(f"\n{category} ({len(files)} 个文件):")
            for file in sorted(files)[:10]:  # 只显示前10个
                print(f"  - {file}")
            if len(files) > 10:
                print(f"  ... 还有 {len(files) - 10} 个文件")

def simulate_git_add():
    """模拟git add操作，显示哪些文件会被添加"""
    print("\n🔍 模拟Git添加操作:")
    print("-" * 50)
    
    try:
        # 运行git status --porcelain来获取文件状态
        result = subprocess.run(['git', 'status', '--porcelain'], 
                              capture_output=True, text=True, encoding='utf-8')
        
        if result.returncode == 0:
            lines = result.stdout.strip().split('\n')
            if lines and lines[0]:
                print("未跟踪的文件:")
                for line in lines:
                    if line.startswith('??'):
                        file_path = line[3:]
                        print(f"  + {file_path}")
            else:
                print("✅ 所有文件都已被跟踪或被忽略")
        else:
            print("❌ 无法获取Git状态")
            
    except FileNotFoundError:
        print("❌ Git命令不可用")
    except Exception as e:
        print(f"❌ 执行Git命令时出错: {e}")

def show_recommended_files():
    """显示推荐包含在Git中的文件"""
    print("\n✅ 推荐包含在Git中的文件:")
    print("-" * 50)
    
    recommended_patterns = [
        '*.py',
        '*.md', 
        '*.pdf',
        'requirements.txt',
        '*.spec',
        '生成器/Excel/*.xlsx',
        '生成器/word/*.docx',
        '生成器/示意图/*'
    ]
    
    recommended_files = []
    
    for pattern in recommended_patterns:
        files = glob.glob(pattern, recursive=True)
        recommended_files.extend(files)
    
    # 添加生成器目录下的文件
    for root, dirs, files in os.walk('生成器'):
        for file in files:
            if not file.startswith('~$') and not file.endswith('.tmp'):
                file_path = os.path.join(root, file)
                if file_path not in recommended_files:
                    recommended_files.append(file_path)
    
    # 按类型分组显示
    file_types = {}
    for file in recommended_files:
        if os.path.exists(file):
            ext = os.path.splitext(file)[1].lower()
            if ext not in file_types:
                file_types[ext] = []
            file_types[ext].append(file)
    
    for ext, files in sorted(file_types.items()):
        print(f"\n{ext or '无扩展名'} 文件 ({len(files)} 个):")
        for file in sorted(files)[:8]:  # 只显示前8个
            print(f"  - {file}")
        if len(files) > 8:
            print(f"  ... 还有 {len(files) - 8} 个文件")

def show_git_commands():
    """显示推荐的Git命令"""
    print("\n🚀 推荐的Git操作命令:")
    print("-" * 50)
    
    commands = [
        ("初始化Git仓库", "git init"),
        ("添加.gitignore", "git add .gitignore"),
        ("添加README", "git add README.md"),
        ("添加所有Python文件", "git add *.py"),
        ("添加文档文件", "git add *.md *.pdf"),
        ("添加模板文件", "git add 生成器/"),
        ("添加配置文件", "git add requirements.txt *.spec"),
        ("查看状态", "git status"),
        ("提交更改", "git commit -m \"Initial commit: NDT结果生成器项目\""),
        ("查看忽略的文件", "git status --ignored"),
    ]
    
    for desc, cmd in commands:
        print(f"{desc}:")
        print(f"  {cmd}")
        print()

if __name__ == "__main__":
    success = check_git_status()
    
    if success:
        analyze_project_files()
        simulate_git_add()
        show_recommended_files()
    
    show_git_commands()
    
    print("\n📝 总结:")
    print("1. .gitignore文件已创建，包含完整的排除规则")
    print("2. 推荐先添加核心文件（源码、文档、模板）")
    print("3. 避免添加构建产物和临时文件")
    print("4. 定期检查Git状态确保文件正确管理")
    print("\n🎯 下一步: 运行 'git init' 开始版本控制")
