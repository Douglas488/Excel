# Docker 部署指南

## 🐳 使用 Docker 部署 Excel 处理工具

### 本地测试 Docker 构建

1. **构建镜像**：
   ```bash
   docker build -t excel-processor .
   ```

2. **运行容器**：
   ```bash
   docker run -p 5000:5000 excel-processor
   ```

3. **测试 API**：
   - 访问：`http://localhost:5000/api/health`
   - 应该返回：`{"status": "healthy", "timestamp": "..."}`

### 部署到云平台

#### 1. Render (推荐)
- 在 Render 中创建新的 Web Service
- 连接 GitHub 仓库
- 选择 "Docker" 作为环境
- Render 会自动检测 Dockerfile 并构建

#### 2. Railway
```bash
# 安装 Railway CLI
npm install -g @railway/cli

# 登录并部署
railway login
railway init
railway up
```

#### 3. Fly.io
```bash
# 安装 Fly CLI
curl -L https://fly.io/install.sh | sh

# 登录并部署
fly auth login
fly launch
fly deploy
```

#### 4. Google Cloud Run
```bash
# 构建并推送到 Google Container Registry
gcloud builds submit --tag gcr.io/PROJECT-ID/excel-processor

# 部署到 Cloud Run
gcloud run deploy --image gcr.io/PROJECT-ID/excel-processor --platform managed
```

#### 5. AWS ECS/Fargate
```bash
# 构建并推送到 ECR
aws ecr get-login-password --region us-east-1 | docker login --username AWS --password-stdin ACCOUNT.dkr.ecr.us-east-1.amazonaws.com
docker tag excel-processor:latest ACCOUNT.dkr.ecr.us-east-1.amazonaws.com/excel-processor:latest
docker push ACCOUNT.dkr.ecr.us-east-1.amazonaws.com/excel-processor:latest
```

### Docker 配置说明

#### Dockerfile 特性
- **基础镜像**: Python 3.9-slim（轻量级）
- **工作目录**: /app
- **端口**: 5000
- **进程管理**: Gunicorn（2个worker，120秒超时）

#### 优化特性
- 多阶段构建减少镜像大小
- 缓存 Python 依赖安装
- 非 root 用户运行（安全性）
- 健康检查支持

### 环境变量配置

在云平台中设置以下环境变量：

```bash
# 必需
PORT=5000

# 可选
PYTHONUNBUFFERED=1
FLASK_ENV=production
```

### 监控和日志

#### 查看容器日志
```bash
docker logs CONTAINER_ID
```

#### 进入容器调试
```bash
docker exec -it CONTAINER_ID /bin/bash
```

#### 健康检查
```bash
curl http://localhost:5000/api/health
```

### 故障排除

#### 常见问题

1. **端口冲突**：
   - 确保端口 5000 未被占用
   - 或修改 Dockerfile 中的 EXPOSE 端口

2. **内存不足**：
   - 增加 Docker 内存限制
   - 或减少 Gunicorn worker 数量

3. **依赖安装失败**：
   - 检查 requirements.txt 格式
   - 确保网络连接正常

4. **CORS 问题**：
   - 确认 app.py 中的 CORS 配置正确
   - 检查云平台的网络设置

### 性能优化

#### 生产环境建议
- 使用多进程部署（Gunicorn workers）
- 配置反向代理（Nginx）
- 启用 HTTPS
- 设置适当的超时时间
- 监控资源使用情况

#### 扩展性考虑
- 使用负载均衡器
- 配置数据库连接池
- 实现缓存机制
- 设置自动扩缩容

### 安全建议

1. **容器安全**：
   - 使用非 root 用户运行
   - 定期更新基础镜像
   - 扫描安全漏洞

2. **网络安全**：
   - 配置防火墙规则
   - 使用 HTTPS
   - 限制访问来源

3. **数据安全**：
   - 加密敏感数据
   - 定期备份
   - 访问控制

---

🎉 **现在你可以使用 Docker 在任何支持容器的平台上部署你的 Excel 处理工具！**
