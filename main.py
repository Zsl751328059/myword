from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pdf2docx import Converter
from docx import Document
from docx.shared import Pt
from docx.enum.table import WD_ALIGN_VERTICAL
import os
import uuid
import shutil
import time
import zipfile
import io
from pathlib import Path

app = FastAPI(title="PDF转Word接口", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("outputs")

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

def cleanup_old_files(max_age_minutes: int = 30):
    now = time.time()
    for dir_path in [UPLOAD_DIR, OUTPUT_DIR]:
        for file in dir_path.glob("*"):
            if file.is_file() and now - file.stat().st_mtime > max_age_minutes * 60:
                file.unlink()

def optimize_docx(docx_path):
    doc = Document(docx_path)
    
    for table in doc.tables:
        table.autofit = True
        for row in table.rows:
            for cell in row.cells:
                tcPr = cell._tc.get_or_add_tcPr()
                for border in tcPr.xpath('w:top|w:bottom|w:left|w:right'):
                    tcPr.remove(border)
    
    doc.save(docx_path)

@app.post("/convert", summary="PDF转Word")
async def convert_pdf_to_word(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="只支持PDF文件")
    
    cleanup_old_files()
    
    file_id = str(uuid.uuid4())
    input_path = UPLOAD_DIR / f"{file_id}.pdf"
    output_path = OUTPUT_DIR / f"{file_id}.docx"
    
    try:
        # 检查上传的文件
        if not file:
            raise HTTPException(status_code=400, detail="未接收到文件")
        
        # 保存上传的文件
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        # 验证PDF文件是否有效
        if input_path.stat().st_size == 0:
            raise HTTPException(status_code=400, detail="上传的PDF文件为空")
        
        # 执行PDF到Word的转换
        cv = None
        try:
            cv = Converter(str(input_path))
            cv.convert(
                str(output_path),
                start=0,
                end=None,
                layout=True,
                recover=True
            )
        finally:
            if cv:
                cv.close()
        
        # 检查输出文件是否存在
        if not output_path.exists():
            raise HTTPException(status_code=500, detail="转换完成后输出文件不存在")
        
        # 优化Word文档
        optimize_docx(str(output_path))
        
        return FileResponse(
            path=output_path,
            filename=f"{Path(file.filename).stem}.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    
    except HTTPException:
        # 重新抛出HTTP异常而不删除文件，以便保留错误信息
        raise
    except ImportError as e:
        # 特别处理导入错误
        if input_path.exists():
            input_path.unlink()
        if output_path.exists():
            output_path.unlink()
        raise HTTPException(status_code=500, detail=f"依赖库缺失: {str(e)}")
    except Exception as e:
        # 记录详细错误信息
        print(f"PDF转换错误详情: {str(e)}")
        print(f"输入文件路径: {input_path}")
        print(f"输出文件路径: {output_path}")
        print(f"输入文件是否存在: {input_path.exists()}")
        print(f"输出文件是否存在: {output_path.exists()}")
        
        if input_path.exists():
            input_path.unlink()
        if output_path.exists():
            output_path.unlink()
        raise HTTPException(status_code=500, detail=f"转换失败: {str(e)}")

@app.get("/", summary="健康检查")
async def health_check():
    return {"status": "ok", "message": "PDF转Word服务运行正常"}

@app.post("/compress", summary="文件压缩")
async def compress_files(files: list[UploadFile] = File(...), level: str = "medium"):
    if not files:
        raise HTTPException(status_code=400, detail="请至少选择一个文件")
    
    level_map = {
        "low": 1,
        "medium": 5,
        "high": 9
    }
    
    if level not in level_map:
        raise HTTPException(status_code=400, detail="压缩等级只能是 low、medium 或 high")
    
    compress_level = level_map[level]
    
    cleanup_old_files()
    
    file_id = str(uuid.uuid4())
    zip_buffer = io.BytesIO()
    
    try:
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED, compresslevel=compress_level) as zip_file:
            for file in files:
                file_content = await file.read()
                zip_file.writestr(file.filename, file_content)
        
        zip_buffer.seek(0)
        output_path = OUTPUT_DIR / f"{file_id}.zip"
        
        with open(output_path, "wb") as f:
            f.write(zip_buffer.getvalue())
        
        return FileResponse(
            path=output_path,
            filename=f"compressed_{file_id[:8]}.zip",
            media_type="application/zip"
        )
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"压缩失败: {str(e)}")
    finally:
        zip_buffer.close()

@app.post("/compress-pdf", summary="PDF专用压缩")
async def compress_pdf(file: UploadFile = File(...), level: str = "medium"):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="只支持PDF文件")
    
    cleanup_old_files()
    
    file_id = str(uuid.uuid4())
    input_path = UPLOAD_DIR / f"{file_id}.pdf"
    output_path = OUTPUT_DIR / f"{file_id}_compressed.pdf"
    
    try:
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        original_size = input_path.stat().st_size
        
        import fitz
        doc = None
        try:
            doc = fitz.open(str(input_path))
            
            if level == "high":
                doc.save(str(output_path), garbage=4, deflate=True, clean=True)
            elif level == "low":
                doc.save(str(output_path), garbage=1, deflate=True)
            else:
                doc.save(str(output_path), garbage=2, deflate=True, clean=True)
        finally:
            if doc:
                doc.close()
        
        compressed_size = output_path.stat().st_size
        reduction = ((original_size - compressed_size) / original_size) * 100 if original_size > 0 else 0
        
        return {
            "success": True,
            "original_size": original_size,
            "compressed_size": compressed_size,
            "reduction": round(reduction, 2),
            "download_url": f"/download/{file_id}_compressed.pdf"
        }
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PDF压缩失败: {str(e)}")
    finally:
        if input_path.exists():
            input_path.unlink()

@app.get("/download/{filename}", summary="下载压缩后的PDF")
async def download_file(filename: str):
    output_path = OUTPUT_DIR / filename
    if not output_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    
    return FileResponse(
        path=output_path,
        filename=filename,
        media_type="application/pdf"
    )