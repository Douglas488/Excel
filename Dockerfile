# syntax=docker/dockerfile:1
FROM python:3.10-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

# 安装运行依赖
COPY requirements.api.txt /app/
RUN pip install -r requirements.api.txt

# 复制源代码（仅后端所需文件）
COPY api_server.py /app/

# 运行 uvicorn（fly.io 默认监听 8080）
ENV PORT=8080
CMD ["uvicorn", "api_server:app", "--host", "0.0.0.0", "--port", "8080"]
