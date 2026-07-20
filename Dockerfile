FROM python:3.11-slim

WORKDIR /app

# 安装依赖
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 复制项目文件
COPY . .

# 创建必要目录
RUN mkdir -p .uploads .trae

# 暴露端口
EXPOSE 8080

# 健康检查
HEALTHCHECK --interval=5m --timeout=3s \
  CMD curl -f http://localhost:8080/ || exit 1

# 启动安全版服务器
CMD ["python", "协作服务器_安全版.py", "8080"]
