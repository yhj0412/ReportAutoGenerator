# Git版本控制说明

## 📋 .gitignore 文件设计说明

为NDT结果生成器项目创建的 `.gitignore` 文件，旨在排除不必要的文件，保持Git仓库的整洁和高效。

## 🚫 **排除的文件类型**

### **1. Python 运行时文件**
```
__pycache__/          # Python字节码缓存
*.pyc                 # 编译的Python文件
build/                # 构建目录
dist/                 # 分发目录
*.egg-info/           # 包信息
```

**原因**: 这些文件是Python运行时自动生成的，不需要版本控制。

### **2. 分发包和构建产物**
```
分发包/               # 分发包目录
*.exe                 # 可执行文件
build/                # PyInstaller构建目录
dist/                 # 输出目录
```

**原因**: 这些是编译和打包的产物，可以通过源代码重新生成。

### **3. 临时和测试文件**
```
test_*.py             # 测试脚本
测试*.py              # 中文测试文件
验证*.py              # 验证脚本
*完成总结.md          # 临时文档
```

**原因**: 这些是开发过程中的临时文件，不是项目的核心代码。

### **4. Office临时文件**
```
~$*.xlsx              # Excel临时文件
~$*.docx              # Word临时文件
*.tmp                 # 临时文件
```

**原因**: Office软件打开文件时自动创建的临时文件。

### **5. 操作系统文件**
```
Thumbs.db             # Windows缩略图
.DS_Store             # macOS系统文件
*~                    # Linux备份文件
```

**原因**: 操作系统自动生成的文件，与项目无关。

### **6. IDE配置文件**
```
.vscode/              # VS Code配置
.idea/                # PyCharm配置
*.sublime-*           # Sublime Text配置
```

**原因**: 个人开发环境配置，不应影响其他开发者。

## ✅ **保留的重要文件**

### **1. 核心源代码**
- `gui.py` - 主界面程序
- `NDT_result.py` - NDT结果处理
- `Surface_Defect.py` - 表面缺陷处理
- `Ray_Detection.py` - 射线检测处理
- `create_distribution.py` - 分发包创建

### **2. 项目文档**
- `README.md` - 项目说明
- `需求文档_详细版.md` - 需求规格
- `用户操作手册.md` - 使用指南
- `系统部署指南.md` - 部署说明
- `详细设计文档.md` - 技术设计

### **3. 模板文件**
- `生成器/Excel/*.xlsx` - Excel模板
- `生成器/word/*.docx` - Word模板
- `生成器/示意图/` - 示意图文件

### **4. 配置文件**
- `requirements.txt` - Python依赖
- `NDT结果生成器.spec` - PyInstaller配置

### **5. 重要文档**
- `NDT结果生成器用户操作手册.pdf`
- `NDT结果生成器系统部署指南.pdf`
- `NDT结果生成器系统需求文档.pdf`
- `NDT结果生成器系统详细设计文档.pdf`

## 📁 **目录结构管理**

### **保留的目录结构**
```
NDT结果生成器/
├── gui.py                    # ✅ 保留
├── NDT_result.py            # ✅ 保留
├── create_distribution.py   # ✅ 保留
├── requirements.txt         # ✅ 保留
├── README.md               # ✅ 保留
├── 生成器/                  # ✅ 保留
│   ├── Excel/              # ✅ 保留模板
│   ├── word/               # ✅ 保留模板
│   ├── 示意图/             # ✅ 保留
│   └── 输出报告/           # ✅ 保留目录结构
└── *.pdf                   # ✅ 保留文档
```

### **排除的目录**
```
__pycache__/                # ❌ 排除
build/                      # ❌ 排除
dist/                       # ❌ 排除
分发包/                     # ❌ 排除
.vscode/                    # ❌ 排除
```

## 🔧 **特殊处理规则**

### **1. 输出报告目录**
```gitignore
# 排除生成的报告文件，但保留目录结构
生成器/输出报告/*/*.docx
生成器/输出报告/*/*.pdf
生成器/输出报告/*/*.xlsx

# 保留目录结构
!生成器/输出报告/*/
!生成器/输出报告/*/.gitkeep
```

### **2. 临时文档处理**
```gitignore
# 排除临时生成的文档
*完成总结.md
*验证报告.md
*测试*.md
```

### **3. 压缩文件处理**
```gitignore
# 排除压缩文件（除非是项目必需的）
*.zip
*.rar
*.7z
```

## 🚀 **Git使用建议**

### **1. 初始化仓库**
```bash
git init
git add .gitignore
git add README.md
git commit -m "Initial commit: Add .gitignore and README"
```

### **2. 添加核心文件**
```bash
# 添加源代码
git add *.py

# 添加文档
git add *.md
git add *.pdf

# 添加模板文件
git add 生成器/

# 添加配置文件
git add requirements.txt
git add *.spec
```

### **3. 提交策略**
```bash
# 功能开发
git add <modified_files>
git commit -m "feat: 添加新功能描述"

# Bug修复
git commit -m "fix: 修复具体问题描述"

# 文档更新
git commit -m "docs: 更新文档内容"

# 重构代码
git commit -m "refactor: 重构模块名称"
```

## 📊 **版本控制统计**

### **预期包含的文件**
- **源代码**: ~15个Python文件
- **文档**: ~10个Markdown文件 + 4个PDF文件
- **模板**: ~20个Excel/Word模板文件
- **配置**: requirements.txt, .spec文件

### **预期排除的文件**
- **构建产物**: build/, dist/, 分发包/
- **临时文件**: __pycache__/, *.pyc, ~$*
- **测试文件**: test_*.py, 测试*.py
- **IDE配置**: .vscode/, .idea/

## ⚠️ **注意事项**

### **1. 敏感信息**
如果项目中包含敏感信息（如API密钥、数据库密码），请确保：
- 使用环境变量或配置文件
- 将配置文件添加到 `.gitignore`
- 提供配置文件模板

### **2. 大文件处理**
如果有大型文件（>100MB）：
- 考虑使用Git LFS
- 或将文件存储在外部位置
- 在README中说明获取方式

### **3. 团队协作**
- 定期更新 `.gitignore`
- 与团队成员同步排除规则
- 避免提交个人配置文件

## 🎯 **总结**

这个 `.gitignore` 文件设计原则：

1. **保留核心**: 所有源代码、文档、模板文件
2. **排除临时**: 构建产物、缓存、临时文件
3. **保护隐私**: IDE配置、个人设置
4. **保持整洁**: 避免不必要的文件污染仓库

通过合理的 `.gitignore` 配置，确保Git仓库只包含项目的核心文件，提高版本控制的效率和团队协作的便利性。
