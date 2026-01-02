# 使用多阶段构建来减小镜像体积
# 第一阶段：构建阶段，安装所有依赖
FROM python:3.10-alpine AS builder

WORKDIR /app

# 安装编译依赖和系统包
RUN apk add --no-cache --virtual .build-deps \
    gcc \
    g++ \
    musl-dev \
    libxml2-dev \
    libxslt-dev \
    && apk add --no-cache \
    libxml2 \
    libxslt

# 复制 requirements.txt
COPY requirements.txt .

# 升级 pip 并安装依赖到一个临时目录
RUN pip install --upgrade pip setuptools wheel && \
    pip install --no-cache-dir --prefix=/install -r requirements.txt

# 清理构建依赖和临时文件
RUN apk del .build-deps && \
    rm -rf /tmp/* /var/tmp/*

# 第二阶段：运行阶段，只保留必要的依赖
FROM python:3.10-alpine

WORKDIR /app

# 安装运行时依赖
RUN apk add --no-cache \
    libxml2 \
    libxslt \
    libgcc \
    libstdc++

# 从构建阶段复制安装好的依赖
COPY --from=builder /install /usr/local

# 复制应用代码
COPY app/ ./app/

# 清理Python环境中不需要的文件
RUN find /usr/local/lib/python3.10 -name "__pycache__" -type d -exec rm -rf {} + && \
    find /usr/local/lib/python3.10 -name "*.pyc" -delete && \
    rm -rf /usr/local/lib/python3.10/site-packages/pip /usr/local/lib/python3.10/site-packages/setuptools

# 暴露端口
EXPOSE 8501

# 使用内置的streamlit健康检查，避免安装curl
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD python -c "import requests; response = requests.get('http://localhost:8501/_stcore/health'); response.raise_for_status()"

# 启动命令
ENTRYPOINT ["streamlit", "run", "app/main.py", "--server.port=8501", "--server.address=0.0.0.0"]