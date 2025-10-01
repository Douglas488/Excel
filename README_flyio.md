# 部署到 fly.io（作为后端 API）

## 1. 代码结构
- `api_server.py`: FastAPI 后端服务，提供 `/healthz` 与 `/process` 接口
- `requirements.api.txt`: 后端依赖清单
- `Dockerfile`: 构建容器镜像
- `.dockerignore`: 构建忽略
- `fly.toml`: fly.io 配置示例（`app` 字段可用 `fly launch` 生成后替换）

## 2. 推送到 GitHub
```bash
# 初始化仓库（如尚未初始化）
git init
git add .
git commit -m "feat: add fastapi backend for fly.io"
# 关联远程并推送（示例）
git remote add origin https://github.com/<your_name>/<your_repo>.git
git branch -M main
git push -u origin main
```

## 3. 安装 flyctl 并初始化应用
- 安装参考：https://fly.io/docs/hands-on/install-flyctl/
```bash
fly auth signup   # 或 fly auth login
fly launch        # 在项目根目录执行，选择使用现有 Dockerfile
# 记下自动生成/选择的 app 名称，更新 fly.toml 的 app 字段
```

## 4. 部署
```bash
fly deploy
fly status
fly logs
```
- 部署成功后，会得到一个公共域名，例如：`https://<app>.fly.dev`

## 5. 健康检查
```bash
curl https://<app>.fly.dev/healthz
# {"status":"ok"}
```

## 6. 处理 Excel（接口：/process）
`/process` 接收 multipart/form-data：
- `file`: 待处理的 Excel 文件（二进制）
- 其他表单字段（可选，默认与桌面版一致）：
  - `sku_sheet`(默认 Sheet1) `sku_title_col`(B) `sku_col`(C)
  - `cost_sheet`(Sheet2) `cost_sku_col`(A) `cost_col`(B)
  - `output_sheet`(Order details) `output_title_col`(A) `output_sku_col`(B) `output_cost_col`(D)
  - `start_row`(2) `end_row`(5000)

### 示例：curl
```bash
curl -X POST \
  -F "file=@./your.xlsx" \
  -F "sku_sheet=Sheet1" -F "sku_title_col=B" -F "sku_col=C" \
  -F "cost_sheet=Sheet2" -F "cost_sku_col=A" -F "cost_col=B" \
  -F "output_sheet=Order details" -F "output_title_col=A" -F "output_sku_col=B" -F "output_cost_col=D" \
  -F "start_row=2" -F "end_row=5000" \
  -o processed.xlsx \
  https://<app>.fly.dev/process
```

### 示例：Python 客户端
```python
import requests

url = "https://<app>.fly.dev/process"
files = {"file": ("your.xlsx", open("your.xlsx", "rb"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
data = {
    "sku_sheet": "Sheet1",
    "sku_title_col": "B",
    "sku_col": "C",
    "cost_sheet": "Sheet2",
    "cost_sku_col": "A",
    "cost_col": "B",
    "output_sheet": "Order details",
    "output_title_col": "A",
    "output_sku_col": "B",
    "output_cost_col": "D",
    "start_row": 2,
    "end_row": 5000,
}
resp = requests.post(url, files=files, data=data)
open("processed.xlsx", "wb").write(resp.content)
print("Found-SKU:", resp.headers.get("X-Found-SKU"))
print("Found-Cost:", resp.headers.get("X-Found-Cost"))
```

## 7. 本地运行（可选）
```bash
pip install -r requirements.api.txt
uvicorn api_server:app --host 0.0.0.0 --port 8080 --reload
```
- 访问：`http://127.0.0.1:8080/healthz`

## 8. 作为后端被调用
- 前端或其他客户端只需按第 6 节以 HTTP 调用 `/process` 上传 Excel，并保存返回的二进制流为 `processed.xlsx` 即可。
