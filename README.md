# Excel数据处理工具

一个用于处理Excel数据的工具，支持SKU查找与成本获取功能。

## 功能特性

- 📊 Excel文件上传与解析
- 🔍 SKU数据查找
- 💰 成本数据匹配
- 📤 结果导出
- 🌐 Web API接口

## 部署方式

### 1. 本地桌面版

运行 `excel_processor.py` 或双击 `启动程序.bat`

### 2. Web API版

#### 本地运行
```bash
pip install -r requirements.txt
python app.py
```

#### Render部署

1. **创建GitHub仓库**
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/你的用户名/excel-processor.git
   git push -u origin main
   ```

2. **在Render部署**
   - 访问 [Render.com](https://render.com)
   - 注册/登录账户
   - 点击 "New +" → "Web Service"
   - 连接GitHub仓库
   - 配置如下：
     - **Name**: excel-processor-api
     - **Environment**: Python 3
     - **Build Command**: `pip install -r requirements.txt`
     - **Start Command**: `gunicorn app:app`
     - **Plan**: Free

3. **环境变量**（可选）
   - `PYTHON_VERSION`: 3.9.0

## API接口

### 基础信息
- **Base URL**: `https://你的应用名.onrender.com`
- **Content-Type**: `application/json`

### 接口列表

#### 1. 健康检查
```http
GET /api/health
```

#### 2. 上传Excel文件
```http
POST /api/upload
Content-Type: multipart/form-data

file: [Excel文件]
```

#### 3. 处理Excel数据
```http
POST /api/process
Content-Type: application/json

{
  "file": {
    "content": "base64编码的文件内容"
  },
  "sku_config": {
    "sheet": "Sheet1",
    "title_col": "B",
    "sku_col": "C"
  },
  "cost_config": {
    "sheet": "Sheet2", 
    "sku_col": "A",
    "cost_col": "B"
  },
  "output_config": {
    "sheet": "Order details",
    "title_col": "A",
    "sku_col": "B", 
    "cost_col": "D",
    "start_row": 2,
    "end_row": 5000
  }
}
```

## 使用示例

### Python客户端示例

```python
import requests
import base64

# 上传文件
with open('data.xlsx', 'rb') as f:
    file_content = base64.b64encode(f.read()).decode('utf-8')

# 处理数据
response = requests.post('https://你的应用名.onrender.com/api/process', json={
    "file": {"content": file_content},
    "sku_config": {
        "sheet": "Sheet1",
        "title_col": "B", 
        "sku_col": "C"
    },
    "cost_config": {
        "sheet": "Sheet2",
        "sku_col": "A",
        "cost_col": "B"
    },
    "output_config": {
        "sheet": "Order details",
        "title_col": "A",
        "sku_col": "B",
        "cost_col": "D"
    }
})

result = response.json()
print(f"处理完成: {result['processed_rows']}行")
print(f"找到SKU: {result['found_sku']}个")
print(f"找到成本: {result['found_cost']}个")

# 保存结果
output_content = base64.b64decode(result['output_file']['content'])
with open('result.xlsx', 'wb') as f:
    f.write(output_content)
```

### JavaScript客户端示例

```javascript
// 上传和处理文件
async function processExcel(file) {
    const formData = new FormData();
    formData.append('file', file);
    
    // 上传文件获取工作表信息
    const uploadResponse = await fetch('/api/upload', {
        method: 'POST',
        body: formData
    });
    const uploadResult = await uploadResponse.json();
    
    // 处理数据
    const processResponse = await fetch('/api/process', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            file: {
                content: await fileToBase64(file)
            },
            sku_config: {
                sheet: "Sheet1",
                title_col: "B",
                sku_col: "C"
            },
            cost_config: {
                sheet: "Sheet2",
                sku_col: "A", 
                cost_col: "B"
            },
            output_config: {
                sheet: "Order details",
                title_col: "A",
                sku_col: "B",
                cost_col: "D"
            }
        })
    });
    
    return await processResponse.json();
}

function fileToBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => resolve(reader.result.split(',')[1]);
        reader.onerror = error => reject(error);
    });
}
```

## 项目结构

```
├── app.py                 # Flask Web API
├── excel_processor.py     # 桌面版GUI应用
├── requirements.txt      # Python依赖
├── render.yaml          # Render部署配置
├── 启动程序.bat          # 桌面版启动脚本
├── 打包程序.bat          # 桌面版打包脚本
├── 01.png               # 操作说明图片
├── 02.png
├── 03.png
└── README.md            # 项目说明
```

## 注意事项

1. **文件大小限制**: 最大50MB
2. **支持格式**: .xlsx, .xls
3. **免费版限制**: Render免费版有资源限制
4. **数据安全**: 处理后的文件会临时存储，请及时下载

## 故障排除

### 常见问题

1. **部署失败**
   - 检查requirements.txt依赖
   - 确认Python版本兼容性

2. **API调用失败**
   - 检查Content-Type设置
   - 验证JSON格式正确性

3. **文件处理错误**
   - 确认Excel文件格式正确
   - 检查工作表名称和列设置

## 许可证

MIT License
