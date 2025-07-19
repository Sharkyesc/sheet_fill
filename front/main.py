from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from typing import List
import os
import shutil
import subprocess
from pathlib import Path
import uuid
import pandas as pd
from docx import Document

app = FastAPI(title="表格填写系统", description="上传资料文件和表格文件，自动填写表格")

# 配置 CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 在生产环境中应该设置具体的域名
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 创建必要的目录
UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("output")
PUBLIC_DIR = Path("front/public")
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# 挂载静态文件目录
app.mount("/static", StaticFiles(directory=PUBLIC_DIR), name="static")

# 存储上传的文件信息
uploaded_files = {
    "material_files": [],
    "table_file": None
}

@app.get("/", response_class=HTMLResponse)
async def root():
    """返回前端页面"""
    try:
        index_file = PUBLIC_DIR / "index.html"
        if index_file.exists():
            with open(index_file, "r", encoding="utf-8") as f:
                return HTMLResponse(content=f.read())
        else:
            return HTMLResponse(content="<h1>前端文件未找到</h1><p>请确保 public/index.html 文件存在</p>")
    except Exception as e:
        return HTMLResponse(content=f"<h1>错误</h1><p>{str(e)}</p>")

@app.get("/api")
async def api_root():
    """API 信息接口"""
    return {"message": "表格填写系统 API"}

@app.post("/upload/material")
async def upload_material_files(files: List[UploadFile] = File(...)):
    """上传资料文件（可多个）"""
    try:
        saved_files = []
        for file in files:
            if not file.filename:
                continue
            
            # 使用原始文件名，如果重名则添加时间戳
            file_id = str(uuid.uuid4())
            original_name = file.filename
            file_path = UPLOAD_DIR / original_name
            
            # 如果文件已存在，添加唯一标识避免冲突
            if file_path.exists():
                name_parts = Path(original_name).stem
                extension = Path(original_name).suffix
                unique_name = f"{name_parts}_{file_id[:8]}{extension}"
                file_path = UPLOAD_DIR / unique_name
            
            # 保存文件
            with open(file_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            
            file_info = {
                "id": file_id,
                "original_name": file.filename,
                "saved_path": str(file_path),
                "size": file_path.stat().st_size
            }
            
            uploaded_files["material_files"].append(file_info)
            saved_files.append(file_info)
        
        return {
            "message": f"成功上传 {len(saved_files)} 个资料文件",
            "files": saved_files
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"上传失败: {str(e)}")

@app.post("/upload/table")
async def upload_table_file(file: UploadFile = File(...)):
    """上传表格文件（仅一个）"""
    try:
        if not file.filename:
            raise HTTPException(status_code=400, detail="未选择文件")
        
        # 检查文件扩展名
        allowed_extensions = {'.xlsx', '.xls', '.csv', '.doc', '.docx'}
        file_extension = Path(file.filename).suffix.lower()
        if file_extension not in allowed_extensions:
            raise HTTPException(
                status_code=400, 
                detail=f"不支持的文件格式。支持的格式: {', '.join(allowed_extensions)}"
            )
        
        # 如果已经有表格文件，删除旧的
        if uploaded_files["table_file"]:
            old_path = Path(uploaded_files["table_file"]["saved_path"])
            if old_path.exists():
                old_path.unlink()
        
        # 使用原始文件名，如果重名则添加时间戳
        file_id = str(uuid.uuid4())
        original_name = file.filename
        file_path = UPLOAD_DIR / original_name
        
        # 如果文件已存在，添加唯一标识避免冲突
        if file_path.exists():
            name_parts = Path(original_name).stem
            extension = Path(original_name).suffix
            unique_name = f"{name_parts}_{file_id[:8]}{extension}"
            file_path = UPLOAD_DIR / unique_name
        
        # 保存文件
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        file_info = {
            "id": file_id,
            "original_name": file.filename,
            "saved_path": str(file_path),
            "size": file_path.stat().st_size
        }
        
        uploaded_files["table_file"] = file_info
        
        return {
            "message": "表格文件上传成功",
            "file": file_info
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"上传失败: {str(e)}")

@app.get("/files/status")
async def get_files_status():
    """获取已上传文件的状态"""
    return {
        "material_files": uploaded_files["material_files"],
        "table_file": uploaded_files["table_file"],
        "material_count": len(uploaded_files["material_files"]),
        "has_table": uploaded_files["table_file"] is not None
    }

@app.post("/process")
async def process_files():
    """处理文件并填写表格"""
    try:
        if not uploaded_files["table_file"]:
            raise HTTPException(status_code=400, detail="请先上传表格文件")
        
        if not uploaded_files["material_files"]:
            raise HTTPException(status_code=400, detail="请先上传资料文件")
        
        # 获取表格文件路径和资料文件路径
        table_file_path = uploaded_files["table_file"]["saved_path"]
        table_file_name = uploaded_files["table_file"]["original_name"]
        table_file_id = uploaded_files["table_file"]["id"]
        
        # 假设第一个资料文件是知识库文件
        knowledge_file_path = uploaded_files["material_files"][0]["saved_path"]
 
        main_args = ['--knowledge', knowledge_file_path, '--forms', table_file_path]
        print("开始运行 main.py ... 参数:", main_args)
        subprocess.run(['python', 'backend/main.py'] + main_args)
    
        # 查找生成的输出文件（文件名保持不变，位于output目录下）
        output_file_path = OUTPUT_DIR / table_file_name
        if not output_file_path.exists():
            raise Exception("输出文件未生成")
        
        return {
            "message": "表格填写完成",
            "download_url": f"/download/{table_file_id}",
            "result": {
                "output_filename": table_file_name
            }
        }

    except subprocess.TimeoutExpired:
        raise HTTPException(status_code=500, detail="处理超时，请稍后重试")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"处理失败: {str(e)}")

@app.get("/download/{file_id}")
async def download_file(file_id: str):
    """下载处理后的文件"""
    try:
        # 通过file_id找到对应的原始文件名
        if uploaded_files["table_file"] and uploaded_files["table_file"]["id"] == file_id:
            original_filename = uploaded_files["table_file"]["original_name"]
            output_file_path = OUTPUT_DIR / original_filename
            
            if not output_file_path.exists():
                raise HTTPException(status_code=404, detail="处理后的文件不存在")
            
            return FileResponse(
                path=output_file_path,
                filename=f"filled_{original_filename}",
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            raise HTTPException(status_code=404, detail="文件ID不存在")
            
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"下载失败: {str(e)}")

@app.delete("/clear")
async def clear_files():
    """清空所有上传的文件"""
    try:
        # 删除上传的文件
        for file_info in uploaded_files["material_files"]:
            file_path = Path(file_info["saved_path"])
            if file_path.exists():
                file_path.unlink()
        
        if uploaded_files["table_file"]:
            file_path = Path(uploaded_files["table_file"]["saved_path"])
            if file_path.exists():
                file_path.unlink()
        
        # 清空记录
        uploaded_files["material_files"] = []
        uploaded_files["table_file"] = None
        
        return {"message": "已清空所有文件"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"清空失败: {str(e)}")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
