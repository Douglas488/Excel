# Render 部署修复指南

## 🚨 当前问题
Render 部署时出现 `ModuleNotFoundError: No module named 'flask_cors'` 错误。

## 🔧 解决方案

### 1. 强制重新部署
在 Render 控制台中：
1. 进入你的服务设置
2. 点击 "Manual Deploy" → "Deploy latest commit"
3. 或者点击 "Settings" → "Build & Deploy" → "Clear build cache" → 重新部署

### 2. 检查 requirements.txt
确保文件内容为：
```
pandas==2.0.3
openpyxl==3.1.2
flask==2.3.3
flask-cors==4.0.0
werkzeug==2.3.7
gunicorn==21.2.0
```

### 3. 备用方案：移除 flask-cors 依赖
如果 Render 仍然无法安装 flask-cors，可以：

1. **修改 requirements.txt**：
```
pandas==2.0.3
openpyxl==3.1.2
flask==2.3.3
werkzeug==2.3.7
gunicorn==21.2.0
```

2. **app.py 已经支持无 flask-cors 运行**：
   - 代码会自动检测 flask-cors 是否可用
   - 如果不可用，会使用手动 CORS 处理
   - 功能完全相同，只是实现方式不同

### 4. 验证部署
部署成功后，访问：
- `https://excel-processor-api.onrender.com/api/health`
- 应该返回：`{"status": "healthy", "timestamp": "..."}`

### 5. 如果仍然失败
尝试以下步骤：

1. **清理 Render 缓存**：
   - 在 Render 控制台删除服务
   - 重新创建服务

2. **使用 Docker 部署**（可选）：
   ```dockerfile
   FROM python:3.9-slim
   WORKDIR /app
   COPY requirements.txt .
   RUN pip install -r requirements.txt
   COPY . .
   CMD ["gunicorn", "app:app"]
   ```

3. **检查 Render 日志**：
   - 查看构建日志中的 pip install 输出
   - 确认所有依赖都正确安装

## 🎯 预期结果
部署成功后，你的 Vercel 前端应该能够正常调用 Render API，不再出现 CORS 错误。

## 📞 如果问题持续
1. 检查 Render 服务状态
2. 查看详细构建日志
3. 尝试使用不同的 Python 版本（3.9 而不是 3.13）
4. 考虑使用其他部署平台（如 Railway、Heroku）
