# Word-Excel 批量替换工具

# 📋 Word+Excel 批量替换工具

🚀 **智能文档批量处理工具** - 让Word和Excel的批量替换变得简单高效！

## ✨ 核心功能

- 🔄 **批量替换** - 基于Excel数据批量替换Word文档内容

- 📊 **表格支持** - 完美支持Word表格内的文字替换

- 🎨 **格式保留** - 替换后保持原有字体、颜色、样式不变

- 🖥️ **可视化界面** - 友好的Web操作界面，无需编程经验

- 📦 **灵活下载** - 支持单个文件下载和ZIP压缩包批量下载

- 💾 **结果持久化** - 替换结果在会话期间持久保存

## 🚀 快速开始

### 🐳 Docker 一键部署（推荐）

```Bash
# 使用最新版本
docker run -p 12344:8501 ghcr.io/你的用户名/word-excel-replace-tool:latest
# 或使用 docker-compose
git clone https://github.com/你的用户名/word-excel-replace-tool.git
cd word-excel-replace-tool
docker-compose up -d

```

### ⚡ 本地运行

```Bash
git clone https://github.com/你的用户名/word-excel-replace-tool.git
cd word-excel-replace-tool
pip install -r requirements.txt
streamlit run app/main.py
```

## 📖 使用流程

1. 📄 **上传文件**

    - Word模板：上传.docx格式的模板文档

    - Excel数据：上传.xlsx/.xls格式的数据文件

2. 👀 **预览文档**

    - 查看Word文档内容（含表格）

    - 选中关键字并按Ctrl+C复制

    - 预览Excel数据结构

3. ⚙️ **设置规则**

    - 粘贴关键字到输入框

    - 选择对应的Excel数据列

    - 添加替换规则（可添加多个）

4. 🚀 **执行替换**

    - 设置文件名格式和前缀

    - 选择替换范围（全部行或指定行）

    - 点击开始批量替换

5. 📥 **下载结果**

    - 单文件下载：分页显示，逐一下载

    - 批量下载：ZIP压缩包一键下载

## 🛠️ 技术栈

|技术|用途|
|---|---|
|🐍 Python 3.10+|后端逻辑处理|
|🎈 Streamlit 1.51.0|Web界面框架|
|📊 Pandas|数据处理|
|📄 python-docx|Word文档处理|
|📊 openpyxl|Excel文件处理|
|🐳 Docker|容器化部署|
## ⚡ 使用建议

### 💡 最佳实践：

- 单次处理建议不超过1000行数据

- 文件大小建议控制在50MB以内

- 确保服务器有2GB+可用内存

- 大文件建议分批次处理

### 🛡️ 安全提示：

- 所有处理在内存中进行

- 会话结束自动清理数据

- 建议在可信网络环境使用

## 🔧 开发构建

### 🐳 构建镜像

```Bash
docker build -t word-excel-tool .
```

### 🏗️ 本地开发

```Bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
# venv\Scripts\activate  # Windows
pip install -r requirements.txt
streamlit run app/main.py
```

---

⭐ 如果这个项目对你有帮助，请给我们一个Star！

🌐 访问地址：[http://localhost:12344](http://localhost:12344)（部署后）
> （文档及代码内容由 AI 生成）
