#!/bin/bash

IMAGE_NAME="ghcr.io/你的github用户名/onedrive-deduper:latest"
DATA_DIR="/opt/sharepoint-deduper/data"

echo "[$(date)] 开始拉取最新镜像..."
docker pull $IMAGE_NAME

echo "[$(date)] 启动去重容器..."
docker run --rm   -e CLIENT_ID="你的_CLIENT_ID"   -e SHAREPOINT_HOST="你的租户.sharepoint.com"   -e SHAREPOINT_SITE="/sites/CommunicationSite"   -e TARGET_FOLDER="/照片"   -e DRY_RUN="false"   -e TG_BOT_TOKEN="你的TG机器人Token"   -e TG_CHAT_ID="你的TG账号ID"   -v $DATA_DIR:/app/data   $IMAGE_NAME

echo "[$(date)] 清理旧版本悬空镜像..."
docker image prune -f --filter "label=org.opencontainers.image.source=https://github.com/你的github用户名/onedrive-deduper"

echo "[$(date)] 任务执行完毕。"
