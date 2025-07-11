# Intelligent Document Filling System (RAG Enhanced, TXT Knowledge)

This project is an AI-powered system for automatically filling in tables in Word/Excel documents using a Retrieval-Augmented Generation (RAG) approach. It leverages a large language model (LLM) and a text file as the knowledge base—no database is required.

## Features
- **Automatic detection of blank fields in tables (Word/Excel)**
- **PDF conversion and page-by-page screenshot generation**
- **Multi-modal analysis combining document images and text content**
- **Semantic analysis of fields using LLM (e.g., GPT)**
- **RAG-based retrieval from a plain text knowledge base (txt file)**
- **Automatic content generation and filling of tables**
- **Support for both horizontal and vertical (label-value) table formats**
- **Real-time system resource monitoring (CPU, Memory, Disk)**
- **Automatic generation of monitoring charts and reports**
- **No database required**

## Directory Structure
```
project_root/
├── main.py
├── rag_engine.py
├── ai_client.py
├── document_processor.py
├── pdf_processor.py          # PDF processing and screenshot functionality
├── monitor.py               # System resource monitoring
├── test_monitor.py          # Monitor testing script
├── config.py
├── requirements.txt
├── examples/
│   ├── sample_data.txt 
│   └── sample.docx 
├── temp/                    # Temporary files directory (PDFs, images, etc.)
├── monitor_output/          # Monitoring data and charts
└── ...
```

## How to Use

### 1. Prepare Your Knowledge Base and Document
- Place your knowledge base as a plain text file (e.g., `sample_data.txt`) in the `examples/` directory.
- Place the document to be filled (e.g., `sample.docx`) in the same directory.
- Each line in the txt file can be a fact, description, or any information you want the AI to use. For best results, use clear, concise lines (e.g., `Name: John Doe`).

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```

### 3. Run the System
```bash
# 基本用法
python main.py --knowledge examples/sample_data.txt --forms examples/sample.docx

# 禁用监控
python main.py --knowledge examples/sample_data.txt --forms examples/sample.docx --no-monitor

# 自定义监控间隔（秒）
python main.py --knowledge examples/sample_data.txt --forms examples/sample.docx --monitor-interval 10

# 测试监控功能
python test_monitor.py
```

### 4. Output
- The system will automatically:
  - Build a RAG index from your txt file
  - **Start monitoring system resources when document processing begins**
  - Detect blank fields in your document
  - Convert document to PDF and generate page screenshots
  - Use the LLM with both images and text content for field analysis
  - Use the LLM and RAG to generate the best fill content
  - Save the filled document as `sample_filled.docx` in the same directory
  - **Stop monitoring and generate charts when processing completes**
  - **Save monitoring data and charts to `monitor_output/` directory**

## Configuration
- By default, the system looks for `examples/sample_data.txt` as the knowledge base and processes all `.docx`/`.xlsx` files in the `examples/` directory.
- You can change the knowledge base path or input directory in `main.py` or `config.py` as needed.
- PDF processing settings can be adjusted in `pdf_processor.py`.

## Knowledge Base (TXT) Format Tips
- Each line should be a fact, description, or key-value pair.
- Example:
  ```
  Name: John Doe
  Gender: Male
  Department: IT
  Email: john.doe@example.com
  John Doe is a software engineer with 5 years of experience in AI.
  ```
- Long paragraphs are supported, but concise lines improve retrieval accuracy.

## No Database Required
- The system does **not** require any SQL or NoSQL database.
- All knowledge is loaded from the txt file and indexed for semantic search.

## Supported Table Formats
- Horizontal: `| Name | _____ | Gender | _____ | ... |`
- Vertical (label-value):
  | Name   | _____ |
  |--------|-------|
  | Gender | _____ |
- The LLM will infer the field type and description from the table context and your knowledge base.

## Advanced
- You can use any txt file as your knowledge base—just update the path in `main.py`.
- The system uses sentence-transformers and FAISS for semantic retrieval.
- The LLM (e.g., OpenAI GPT) is used for both field analysis and content generation.
- Multi-modal analysis combines document images with text for better understanding.

## System Monitoring
The system includes a comprehensive monitoring module that tracks:
- **CPU Usage**: Real-time CPU utilization percentage
- **Memory Usage**: Memory consumption in MB and percentage
- **Disk Usage**: Disk space utilization percentage
- **Network I/O**: Network bytes sent/received and packet counts

### Monitoring Features:
- **Real-time monitoring**: Collects data at configurable intervals (default: 5 seconds)
- **Automatic start/stop**: Monitoring starts when document processing begins and stops when processing completes
- **Data persistence**: Saves monitoring data to JSON files
- **Chart generation**: Creates comprehensive charts showing resource usage over time
- **Thread-safe**: Uses threading for non-blocking data collection
- **Configurable**: Adjustable monitoring interval and maximum record count

### Monitoring Output:
- **Data files**: `monitor_output/monitor_data_YYYYMMDD_HHMMSS.json`
- **Chart files**: `monitor_output/monitor_charts_YYYYMMDD_HHMMSS.png`
- **Summary reports**: Printed to console with statistics and averages

## Troubleshooting
- If the system does not fill fields as expected, try making your txt knowledge base more explicit (e.g., use more key-value lines).
- For best results, use clear and unambiguous field names in your tables and knowledge base.
- **PDF Processing Issues**: If PDF conversion fails, the system will automatically fall back to text-only analysis.
- **Memory Usage**: Large documents with many pages may require more memory for image processing.

## Example
Suppose your `sample_data.txt` contains:
```
Name: John Doe
Gender: Male
Department: IT
Email: john.doe@example.com
John Doe is a software engineer with 5 years of experience in AI.
```
And your `sample.docx` table has:
| Name   | _____ |
|--------|-------|
| Gender | _____ |
| Email  | _____ |

The system will automatically:
1. Convert the document to PDF
2. Generate screenshots of each page
3. Analyze both the visual layout and text content
4. Fill in the blanks with the most relevant information from the txt file

---

For further customization or integration, please refer to the code comments or contact the maintainer. 