# Backend - Sheet Fill 智能文档填充系统后端

## 概述

Backend 是 Sheet Fill 智能文档填充系统的核心后端模块，负责文档处理、AI 分析、RAG 检索、PDF 转换和系统监控等核心功能。该模块采用模块化设计，各组件职责清晰，易于维护和扩展。

## 项目结构

```
backend/
├── main.py                    # 主程序入口和控制器
├── ai_client.py              # AI客户端，处理与大语言模型的交互
├── config.py                 # 配置管理模块
├── document_processor.py     # 文档处理核心模块
├── pdf_processor.py          # PDF转换和图像处理模块
├── rag_engine.py             # 检索增强生成引擎
├── monitor.py                # 系统资源监控模块
├── requirements copy.txt     # Python依赖包列表
├── ___init__.py              # Python包初始化文件
├── __pycache__/              # Python缓存目录
└── monitor_output/           # 监控数据输出目录
```

## 核心模块详解

### 1. main.py - 主控制器
**功能**: 系统的主要入口点和流程控制器
- **类**: `DocumentFiller`, `Logger`
- **主要功能**:
  - 命令行参数解析
  - 系统初始化和组件协调
  - 文档处理流程控制
  - 日志管理和彩色输出
  - 监控系统集成

**关键方法**:
```python
def fill_document(self, knowledge_file: str, form_file: str) -> str
def process_single_document(self, knowledge_file: str, document_path: str) -> str
```

### 2. ai_client.py - AI客户端
**功能**: 与大型语言模型进行交互的核心模块
- **类**: `AIClient`
- **依赖**: OpenAI API, RAGEngine
- **主要功能**:
  - 多模态文档分析（文本+图像）
  - 空白字段智能识别
  - RAG增强的内容生成
  - 最终填充决策

**关键方法**:
```python
def analyze_empty_fields_with_images(self, doc_text: str, page_images: List[str]) -> Dict
def analyze_empty_fields_by_index(self, doc_text: str, fields_with_index: List[Dict]) -> Dict
def final_fill_decision(self, all_fields: List[Dict], described_fields: List[Dict]) -> Dict
def build_rag_from_files(self, knowledge_files: List[str]) -> None
```

### 3. config.py - 配置管理
**功能**: 系统配置和环境变量管理
- **类**: `Config`
- **配置项**:
  - OpenAI API 设置
  - 文件路径配置
  - 高亮颜色设置
  - 目录自动创建

**配置示例**:
```python
OPENAI_API_KEY = 'your-api-key'
OPENAI_BASE_URL = 'http://your-endpoint/v1/'
OPENAI_MODEL = 'gemini-2.5-pro-preview-06-05'
INPUT_DIR = './examples'
OUTPUT_DIR = './output'
```

### 4. document_processor.py - 文档处理器
**功能**: Word和Excel文档的核心处理模块
- **类**: `DocumentProcessor`
- **支持格式**: .docx, .xlsx, .xls
- **主要功能**:
  - 文档内容提取
  - 空白字段检测和编号
  - 智能字段填充
  - 格式保持

**关键方法**:
```python
def find_and_number_all_fields(self, file_path: str) -> Tuple[List[Dict], str]
def extract_document_content(self, file_path: str) -> str
def fill_document(self, file_path: str, filled_cells: List[Dict]) -> str
def restore_cells_content_from_indexed(self, file_path: str) -> str
```

### 5. rag_engine.py - RAG检索引擎
**功能**: 基于语义检索的知识库管理
- **类**: `RAGEngine`
- **技术栈**: Sentence Transformers + FAISS
- **主要功能**:
  - 文本向量化
  - 语义相似度搜索
  - 知识库索引构建
  - 文档分块处理

**关键方法**:
```python
def load_txt_knowledge(self, txt_path: str, chunk_size: int = 300) -> List[Dict]
def semantic_search(self, queries: List[str], top_k: int = 5) -> List[List[Dict]]
def add_documents(self, documents: List[Dict]) -> None
def save_index(self) -> None
```

### 6. pdf_processor.py - PDF处理器
**功能**: PDF转换和图像处理
- **类**: `PDFProcessor`
- **主要功能**:
  - Word/Excel到PDF转换
  - PDF页面截图生成
  - 图像编码处理
  - 多平台兼容性

**关键方法**:
```python
def process_document_with_images(self, document_path: str) -> Dict[str, Any]
def convert_docx_to_pdf(self, docx_path: str) -> str
def generate_page_screenshots(self, pdf_path: str) -> List[str]
def encode_image_to_base64(self, image_path: str) -> str
```

### 7. monitor.py - 系统监控器
**功能**: 实时系统资源监控和可视化
- **类**: `SystemMonitor`
- **监控指标**:
  - CPU使用率
  - 内存使用情况
  - 磁盘空间
  - 网络I/O

**关键方法**:
```python
def start_monitoring(self) -> None
def stop_monitoring(self) -> None
def get_system_info(self) -> Dict[str, Any]
def save_monitoring_data(self, output_dir: str) -> str
def generate_charts(self, output_dir: str) -> str
```

## 依赖管理

### 核心依赖
```
# 文档处理
python-docx          # Word文档处理
openpyxl            # Excel文档处理
pandas              # 数据分析

# AI和机器学习
openai              # OpenAI API客户端
sentence-transformers # 文本向量化
faiss-cpu           # 向量检索引擎
numpy               # 数值计算
scikit-learn        # 机器学习工具

# 图像和PDF处理
pdf2image           # PDF转图像
Pillow              # 图像处理
PyMuPDF             # PDF操作
pytesseract         # OCR文字识别

# 系统监控
psutil              # 系统信息获取
matplotlib          # 数据可视化

# 其他工具
colorama            # 终端彩色输出
requests            # HTTP请求
python-dotenv       # 环境变量管理
pywin32             # Windows系统调用
```

### 安装依赖
```bash
pip install -r "requirements.txt"
```

## 使用方式

### 1. 作为独立模块运行
```bash
cd backend
python main.py --knowledge ../examples/sample_data.txt --forms ../examples/sample.docx
```

### 2. 作为Python包导入
```python
from backend.main import DocumentFiller
from backend.ai_client import AIClient
from backend.document_processor import DocumentProcessor

# 创建文档填充器
filler = DocumentFiller()
result = filler.fill_document(knowledge_file, form_file)
```

### 3. 组件级使用
```python
# 使用文档处理器
processor = DocumentProcessor()
fields, numbered_doc = processor.find_and_number_all_fields(doc_path)

# 使用AI客户端
ai_client = AIClient()
analysis = ai_client.analyze_empty_fields_with_images(doc_text, images)

# 使用RAG引擎
rag = RAGEngine()
rag.load_txt_knowledge(knowledge_file)
results = rag.semantic_search(["查询内容"])
```

## 数据流和处理流程

### 1. 文档预处理阶段
```
输入文档 → find_and_number_all_fields → 字段编号文档
         ↓
转换为PDF → generate_page_screenshots → 页面截图
```

### 2. AI分析阶段
```
文档内容 + 页面图像 → analyze_empty_fields_with_images → 字段描述
                  ↓
字段描述 → RAG语义检索 → 相关知识片段
         ↓
最终决策 → final_fill_decision → 填充方案
```

### 3. 文档填充阶段
```
原始文档 + 填充方案 → fill_document → 填充后文档
                   ↓
格式恢复 → restore_cells_content → 最终输出
```

## 输出文件

### 1. 处理过程文件
- `mid_docs/`: 中间处理文档
- `temp/`: 临时文件（PDF、截图等）

### 2. 监控数据
- `monitor_output/monitor_data_*.json`: 原始监控数据
- `monitor_output/monitor_charts_*.png`: 可视化图表

### 3. 日志文件
- `temp/document_fill_log_*.txt`: 详细处理日志
