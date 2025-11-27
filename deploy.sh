#!/bin/bash

set -e

# 配置变量
GITHUB_USERNAME="MaroD1M"  # 请替换为你的GitHub用户名
IMAGE_NAME="ghcr.io/$GITHUB_USERNAME/word-excel-replace-tool:latest"
CONTAINER_NAME="word-excel-replace-tool"
PORT="8501"

echo "🚀 开始部署 Word+Excel 批量替换工具..."

# 检查 Docker 是否安装
if ! command -v docker &> /dev/null; then
    echo "❌ 未检测到 Docker，请先安装 Docker"
    exit 1
fi

# 拉取最新镜像
echo "📥 拉取最新镜像..."
docker pull $IMAGE_NAME

# 停止并删除现有容器
echo "🔄 清理现有容器..."
docker stop $CONTAINER_NAME 2>/dev/null || true
docker rm $CONTAINER_NAME 2>/dev/null || true

# 启动新容器
echo "🎯 启动容器..."
docker run -d \
  --name $CONTAINER_NAME \
  -p $PORT:8501 \
  --restart unless-stopped \
  $IMAGE_NAME

echo "✅ 部署完成！"
echo "🌐 访问地址: http://localhost:$PORT"
echo "📝 容器名称: $CONTAINER_NAME"
echo "🛑 停止命令: docker stop $CONTAINER_NAME"
echo "📊 查看日志: docker logs -f $CONTAINER_NAME"
