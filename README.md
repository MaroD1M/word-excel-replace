# 📋 Word+Excel 批量替换工具

🚀 **基于 Streamlit 的智能文档批量处理工具** - 让 Word 和 Excel 的批量替换变得简单高效！

![Streamlit](https://img.shields.io/badge/Streamlit-1.51.0-FF4B4B?style=for-the-badge&logo=streamlit)
![Python](https://img.shields.io/badge/Python-3.10-3776AB?style=for-the-badge&logo=python)
![Docker](https://img.shields.io/badge/Docker-Enabled-2496ED?style=for-the-badge&logo=docker)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)

## ✨ 核心功能

🔄 **智能批量替换** - 基于 Excel 数据批量替换 Word 文档内容  
📊 **完美表格支持** - 支持 Word 表格内的文字替换，保持表格结构  
🎨 **原格式保留** - 替换后保持原有字体、颜色、样式不变  
🖥️ **可视化界面** - 友好的 Web 操作界面，无需编程经验  
📦 **灵活下载** - 支持单个文件下载和 ZIP 压缩包批量下载  
💾 **结果持久化** - 替换结果在会话期间持久保存  

## 🚀 快速部署【访问地址: http://localhost:12344 (部署后)】🌐 

### 🐳 Docker 一键部署（推荐）

```bash
# 使用最新版本
docker run -p 12344:8501 ghcr.io/MaroD1M/word-excel-replace-tool:latest

# 或者使用 Docker Compose
git clone https://github.com/MaroD1M/word-excel-replace-tool.git
cd word-excel-replace-tool
docker-compose up -d
⚡ 本地开发运行
bash
# 1. 克隆项目
git clone https://github.com/MaroD1M/word-excel-replace-tool.git
cd word-excel-replace-tool

# 2. 安装依赖
pip install -r requirements.txt

# 3. 启动应用
streamlit run app/main.py
🎯 Portainer 部署
🖱️ 在 Portainer 中进入 "Stacks"

📝 粘贴以下配置：

yaml
version: '3.8'
services:
  word-excel-replace:
    image: ghcr.io/MaroD1M/word-excel-replace-tool:latest
    container_name: word-excel-replace-tool
    ports:
      - "12344:8501"
    restart: unless-stopped
    environment:
      - STREAMLIT_SERVER_HEADLESS=true
      - STREAMLIT_BROWSER_GATHER_USAGE_STATS=false
🎉 点击 "Deploy" 完成部署

📖 详细使用指南
🎯 第一步：上传文件
文件类型	支持格式	说明
📄 Word 模板	.docx	包含需要替换的关键字的模板文档
📊 Excel 数据	.xlsx / .xls	包含替换数据的表格文件
👀 第二步：文档预览与复制
🔍 Word 预览：查看文档内容，支持表格显示

📋 复制关键字：选中 Word 中的关键字，按 Ctrl+C 复制

👁️ Excel 预览：查看数据结构，确认列名和内容

⚙️ 第三步：设置替换规则
📝 粘贴关键字：将复制的关键字粘贴到输入框

🎯 选择数据列：选择对应的 Excel 数据列

➕ 添加规则：点击"添加规则"按钮

📋 规则管理：可添加多个规则，支持删除和清空

🚀 第四步：执行替换
📛 文件名设置：

选择文件名核心字段（取自 Excel）

设置文件名前缀（如"合同_"）

🎯 替换范围：

🔢 全部行：处理所有数据

📍 指定行：自定义处理范围

⚡ 开始替换：点击"开始批量替换"按钮

📥 第五步：下载结果
下载方式	说明	适用场景
📄 单文件下载	分页显示，逐一下载	需要选择性下载
📦 批量下载	ZIP 压缩包一键下载	批量处理所有文件
🛠️ 技术架构
text
word-excel-replace-tool/
├── 🐍 app/
│   └── main.py                 # 主应用程序
├── ⚙️ .github/
│   └── workflows/
│       └── docker-publish.yml  # 自动构建配置
├── 🐳 Dockerfile               # 容器化配置
├── 🎯 docker-compose.yml       # 容器编排
├── 📋 requirements.txt         # 依赖管理
├── 🚀 deploy.sh               # 部署脚本
└── 📖 README.md               # 项目文档
🎯 技术栈
技术	版本	用途
🐍 Python	3.10+	后端逻辑处理
🎈 Streamlit	1.51.0	Web 界面框架
📊 Pandas	2.3.3+	数据处理
📄 python-docx	1.2.0+	Word 文档处理
📊 openpyxl	3.1.5+	Excel 文件处理
🐳 Docker	20.10+	容器化部署
🔄 GitHub Actions	-	自动化构建
⚡ 性能建议
💡 最佳实践：

📊 单次处理建议不超过 1000 行数据

💾 文件大小建议控制在 50MB 以内

🖥️ 确保服务器有 2GB+ 可用内存

🔄 大文件建议分批次处理

🛡️ 安全提示：

🔒 所有处理在内存中进行

🗑️ 会话结束自动清理数据

🌐 建议在可信网络环境使用

🐛 常见问题
❓ 替换后格式乱了？
✅ 解决方案：工具会保留原格式，请检查关键字是否包含特殊字符

❓ 表格内容没有替换？
✅ 解决方案：确保复制的关键字来自表格单元格内

❓ 下载按钮点击无效？
✅ 解决方案：刷新页面重新执行替换，或检查浏览器下载设置

❓ 构建失败？
✅ 解决方案：检查 Dockerfile 和 requirements.txt 配置

🔧 开发与构建
🏗️ 本地开发
bash
# 1. 创建虚拟环境
python -m venv venv
source venv/bin/activate  # Linux/Mac
# venv\Scripts\activate  # Windows

# 2. 安装依赖
pip install -r requirements.txt

# 3. 运行开发服务器
streamlit run app/main.py
🐳 构建镜像
bash
# 本地构建测试
docker build -t word-excel-tool .

# 使用 GitHub Actions 自动构建
git tag v1.0.0
git push origin v1.0.0
⚙️ 环境配置
创建 .env 文件：

bash
GITHUB_USERNAME=你的GitHub用户名
APP_VERSION=latest
EXTERNAL_PORT=12344

⭐ 如果这个项目对你有帮助，请给我们一个 Star！ ⭐
