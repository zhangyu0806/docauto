#!/usr/bin/env python3
"""
企业文档自动化 - FastAPI服务
生产级API，支持PDF处理、Excel清洗、Word模板生成
"""

import os
import sys
import uuid
import shutil
from pathlib import Path
from typing import Optional
from datetime import datetime

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

sys.path.insert(0, str(Path(__file__).parent))

from pdf_to_excel_v2 import PDFProcessorV2
from excel_cleaner import ExcelCleaner
from template_generator import WordTemplateGenerator

# 目录配置
BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

app = FastAPI(
    title="企业文档自动化API",
    description="PDF→Excel、Excel清洗、Word模板生成",
    version="2.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


def save_upload(file: UploadFile, prefix: str = "") -> Path:
    """保存上传文件，返回路径"""
    file_id = str(uuid.uuid4())[:8]
    safe_name = file.filename.replace("/", "_").replace("\\", "_")
    path = UPLOAD_DIR / f"{prefix}{file_id}_{safe_name}"
    with open(path, "wb") as f:
        shutil.copyfileobj(file.file, f)
    return path


@app.get("/", response_class=HTMLResponse)
async def index():
    """首页"""
    html_path = BASE_DIR / "templates" / "index.html"
    if html_path.exists():
        return HTMLResponse(html_path.read_text(encoding="utf-8"))
    return HTMLResponse("<h1>企业文档自动化API</h1><p>访问 /docs 查看API文档</p>")


@app.get("/health")
async def health():
    return {"status": "ok", "time": datetime.now().isoformat()}


@app.post("/api/pdf-to-excel")
async def api_pdf_to_excel(file: UploadFile = File(...)):
    """PDF转Excel"""
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "只支持PDF文件")

    upload_path = save_upload(file, "pdf_")
    output_name = f"{Path(file.filename).stem}.xlsx"
    output_path = OUTPUT_DIR / f"{uuid.uuid4().hex[:8]}_{output_name}"

    try:
        processor = PDFProcessorV2()
        result = processor.extract_tables(str(upload_path))

        if not result.tables:
            raise HTTPException(400, "未从PDF中提取到表格数据")

        import pandas as pd
        with pd.ExcelWriter(str(output_path), engine="openpyxl") as writer:
            all_data = pd.concat(result.tables, ignore_index=True)
            all_data.to_excel(writer, sheet_name="全部数据", index=False)
            for i, df in enumerate(result.tables, 1):
                df.to_excel(writer, sheet_name=f"表格{i}", index=False)

        return FileResponse(
            str(output_path),
            filename=output_name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    finally:
        upload_path.unlink(missing_ok=True)


@app.post("/api/excel-clean")
async def api_excel_clean(
    file: UploadFile = File(...),
    remove_dup: bool = Form(True),
    remove_empty: bool = Form(True),
    standardize_date: bool = Form(True),
    standardize_phone: bool = Form(True),
):
    """Excel数据清洗"""
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(400, "只支持Excel文件")

    upload_path = save_upload(file, "excel_")
    output_name = f"{Path(file.filename).stem}_cleaned.xlsx"
    output_path = OUTPUT_DIR / f"{uuid.uuid4().hex[:8]}_{output_name}"

    try:
        cleaner = ExcelCleaner()
        df = cleaner.load_excel(str(upload_path))
        if df is None:
            raise HTTPException(400, "Excel文件读取失败")

        df_cleaned = cleaner.apply_all_cleaning(
            df,
            remove_dup=remove_dup,
            remove_empty=remove_empty,
            standardize_date=standardize_date,
            standardize_phone=standardize_phone,
        )
        cleaner.save_excel(df_cleaned, str(output_path))

        return FileResponse(
            str(output_path),
            filename=output_name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    finally:
        upload_path.unlink(missing_ok=True)


@app.post("/api/template-generate")
async def api_template_generate(
    template: UploadFile = File(...),
    data: UploadFile = File(...),
):
    """Word模板批量生成"""
    if not template.filename.lower().endswith(".docx"):
        raise HTTPException(400, "模板必须是.docx文件")

    template_path = save_upload(template, "tpl_")
    data_path = save_upload(data, "data_")
    batch_dir = OUTPUT_DIR / f"batch_{uuid.uuid4().hex[:8]}"
    batch_dir.mkdir(exist_ok=True)

    try:
        generator = WordTemplateGenerator(str(template_path))

        if data.filename.lower().endswith(".json"):
            import json
            with open(data_path, "r", encoding="utf-8") as f:
                d = json.load(f)
            if isinstance(d, list):
                generator.batch_generate(d, str(batch_dir))
            else:
                generator.generate_document(d, str(batch_dir / "output.docx"))
        elif data.filename.lower().endswith((".xlsx", ".xls")):
            generator.generate_from_excel(str(data_path), output_dir=str(batch_dir))
        else:
            raise HTTPException(400, "数据文件需要是JSON或Excel格式")

        # 打包为zip
        import zipfile
        zip_path = OUTPUT_DIR / f"documents_{uuid.uuid4().hex[:8]}.zip"
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for f in batch_dir.iterdir():
                if f.is_file():
                    zf.write(f, f.name)

        return FileResponse(str(zip_path), filename="generated_documents.zip", media_type="application/zip")
    finally:
        template_path.unlink(missing_ok=True)
        data_path.unlink(missing_ok=True)
        shutil.rmtree(batch_dir, ignore_errors=True)


if __name__ == "__main__":
    import uvicorn
    print("=" * 50)
    print("企业文档自动化API v2.0")
    print(f"本地: http://127.0.0.1:5002")
    print(f"公网: http://38.55.133.19:5002")
    print(f"API文档: http://38.55.133.19:5002/docs")
    print("=" * 50)
    uvicorn.run(app, host="0.0.0.0", port=5002)
