# Front - Sheet Fill 智能文档填充系统前端

## 概述

Front 是 Sheet Fill 智能文档填充系统的 Web 前端模块，基于 FastAPI 框架构建，提供了直观易用的 Web 界面，允许用户通过浏览器上传文档、处理表格并下载结果。该模块采用前后端分离的架构设计，支持文件上传、实时处理状态监控和结果下载等功能。

## 项目结构

```
front/
├── main.py              # FastAPI 后端服务主程序
├── requirements.txt     # Python 依赖包列表
├── README.md           # 项目说明文档
└── public/             # 静态前端资源目录
    └── index.html      # 主页面 HTML 文件
```

## 技术架构

### 后端框架
- **FastAPI**: 现代、快速的 Web 框架
- **Uvicorn**: ASGI 服务器
- **python-docx**: Word 文档处理
- **pandas**: 数据处理
- **pathlib**: 文件路径操作

### 前端技术
- **HTML5**: 现代 Web 标准
- **CSS3**: 样式设计和动画
- **JavaScript (ES6+)**: 交互逻辑
- **Fetch API**: 异步 HTTP 请求

## 核心功能模块

### 1. Web 服务 (main.py)

#### 应用配置
```python
app = FastAPI(title="表格填写系统", description="上传资料文件和表格文件，自动填写表格")

# CORS 配置 - 支持跨域请求
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
```

#### 主要 API 端点

##### 1. 静态资源服务
- **`GET /`**: 返回主页面 HTML
- **`/static/*`**: 静态文件服务

##### 2. 文件上传 API
- **`POST /upload/material`**: 上传资料文件（支持多文件）
- **`POST /upload/table`**: 上传表格文件（单文件）

**资料文件上传示例**:
```python
@app.post("/upload/material")
async def upload_material_files(files: List[UploadFile] = File(...)):
    # 支持多个资料文件上传
    # 文件重名自动处理
    # 返回文件信息和上传状态
```

**表格文件上传示例**:
```python
@app.post("/upload/table")
async def upload_table_file(file: UploadFile = File(...)):
    # 单个表格文件上传
    # 支持格式: .xlsx, .xls, .csv, .doc, .docx
    # 自动替换旧文件
```

##### 3. 文件管理 API
- **`GET /files/status`**: 获取上传文件状态
- **`DELETE /clear`**: 清空所有上传文件

##### 4. 处理和下载 API
- **`POST /process`**: 执行文档填充处理
- **`GET /download/{file_id}`**: 下载处理结果

**文档处理流程**:
```python
@app.post("/process")
async def process_files():
    # 1. 验证文件完整性
    # 2. 调用 backend 处理模块
    # 3. 生成处理结果
    # 4. 返回下载链接
```

### 2. 前端界面 (public/index.html)

#### 界面设计特点
- **响应式设计**: 支持多种屏幕尺寸
- **现代化 UI**: 渐变色彩和流畅动画
- **用户友好**: 直观的操作流程
- **状态反馈**: 实时显示处理状态

#### 主要功能区域

##### 1. 文件状态显示
- 已上传文件列表
- 文件大小和状态
- 删除和重新上传功能

##### 2. 处理控制区域
- 开始处理按钮
- 处理进度显示
- 结果下载链接

#### JavaScript 交互逻辑

##### 文件上传处理
```javascript
async function uploadMaterialFiles() {
    const formData = new FormData();
    const files = document.getElementById('materialFiles').files;
    
    for (let file of files) {
        formData.append('files', file);
    }
    
    const response = await fetch('/upload/material', {
        method: 'POST',
        body: formData
    });
    
    // 处理响应和UI更新
}
```

##### 文档处理请求
```javascript
async function processFiles() {
    showLoading(true);
    
    try {
        const response = await fetch('/process', {
            method: 'POST'
        });
        
        const result = await response.json();
        
        if (response.ok) {
            showDownloadLink(result.download_url);
        } else {
            showError(result.detail);
        }
    } finally {
        showLoading(false);
    }
}
```

## 目录和文件管理

### 自动创建目录
系统启动时自动创建必要目录：
```python
UPLOAD_DIR = Path("uploads")      # 上传文件存储
OUTPUT_DIR = Path("output")       # 处理结果存储
PUBLIC_DIR = Path("front/public") # 静态资源目录
```

### 文件存储策略
- **上传文件**: 保存在 `uploads/` 目录
- **处理结果**: 存储在 `output/` 目录
- **文件重名**: 自动添加 UUID 前缀避免冲突
- **临时文件**: 处理完成后自动清理

### 支持的文件格式

#### 资料文件
- **文本文件**: .txt, .md
- **Word 文档**: .doc, .docx
- **Excel 表格**: .xlsx, .xls
- **其他格式**: 根据 backend 支持情况

#### 表格文件
- **Excel**: .xlsx, .xls
- **CSV**: .csv
- **Word**: .doc, .docx（包含表格的文档）

## 部署和运行

### 1. 环境准备

#### 安装依赖
```bash
cd front
pip install -r requirements.txt
```

### 2. 启动服务

#### 开发模式
```bash
python main.py
```

### 3. 访问应用
- **主页**: http://localhost:8000

## API 接口文档

### 文件上传接口

#### 上传资料文件
```http
POST /upload/material
Content-Type: multipart/form-data

Body: files (multiple files)

Response:
{
    "message": "成功上传 2 个资料文件",
    "files": [
        {
            "id": "uuid-string",
            "original_name": "sample_data.txt",
            "saved_path": "uploads/sample_data.txt",
            "size": 1024
        }
    ]
}
```

#### 上传表格文件
```http
POST /upload/table
Content-Type: multipart/form-data

Body: file (single file)

Response:
{
    "message": "表格文件上传成功",
    "file": {
        "id": "uuid-string",
        "original_name": "form.xlsx",
        "saved_path": "uploads/form.xlsx",
        "size": 2048
    }
}
```

### 文档处理接口
```http
POST /process

Response:
{
    "message": "表格填写完成",
    "download_url": "/download/uuid-string",
    "result": {
        "output_filename": "filled_form.xlsx"
    }
}
```

### 文件下载接口
```http
GET /download/{file_id}

Response: File download stream
Content-Disposition: attachment; filename="filled_form.xlsx"
```

## 前端界面功能

### 1. 文件上传区域
- **拖拽上传**: 支持拖拽文件到指定区域
- **多文件选择**: 资料文件支持批量上传
- **格式验证**: 自动检查文件格式
- **进度显示**: 上传进度实时反馈

### 2. 文件管理面板
- **文件列表**: 显示已上传文件
- **文件信息**: 文件名、大小、状态
- **删除功能**: 单独删除或批量清空
- **状态指示**: 上传成功/失败状态

### 3. 处理控制面板
- **处理按钮**: 一键启动文档处理
- **进度指示**: 处理状态实时显示
- **结果展示**: 处理完成后显示结果
- **下载链接**: 直接下载处理结果
