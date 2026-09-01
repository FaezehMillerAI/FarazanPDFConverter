"""
FastAPI Web Application backend for FarazanPDFConverter.
Provides endpoints for upload, pre-flight inspection, conversion, and file download.
"""

import os
import shutil
import uuid
from typing import Optional
from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.requests import Request

from omnipdf.core.converter import OmniConverter, ConversionMode
from omnipdf.core.inspector import PDFInspector

app = FastAPI(
    title="FarazanPDFConverter - PDF to Word Converter",
    description="Advanced layout-preserving PDF to editable Word (DOCX) converter",
    version="1.0.0",
)

import tempfile

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
STATIC_DIR = os.path.join(BASE_DIR, "static")
TEMP_DIR = os.path.join(tempfile.gettempdir(), "omnipdf_temp")

os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)
os.makedirs(TEMPLATES_DIR, exist_ok=True)

templates = Jinja2Templates(directory=TEMPLATES_DIR)
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    """Render the main conversion interface."""
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/api/inspect")
async def inspect_endpoint(file: UploadFile = File(...)):
    """Upload and inspect a PDF file before conversion."""
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Only PDF files are supported.")

    file_id = str(uuid.uuid4())
    pdf_temp_path = os.path.join(TEMP_DIR, f"{file_id}_{file.filename}")

    try:
        with open(pdf_temp_path, "wb") as f:
            shutil.copyfileobj(file.file, f)

        inspector = PDFInspector(pdf_temp_path)
        report = inspector.inspect()

        return JSONResponse(content={
            "file_id": file_id,
            "filename": file.filename,
            "report": report.to_dict(),
        })
    except Exception as e:
        if os.path.exists(pdf_temp_path):
            os.remove(pdf_temp_path)
        raise HTTPException(status_code=500, detail=f"Inspection failed: {str(e)}")


@app.post("/api/convert")
async def convert_endpoint(
    file: Optional[UploadFile] = File(None),
    file_id: Optional[str] = Form(None),
    filename: Optional[str] = Form(None),
    mode: str = Form("auto"),
    convert_math: bool = Form(True),
    highlight_code: bool = Form(True),
    extract_images: bool = Form(True),
    image_dpi: int = Form(300),
):
    """Convert a PDF file to DOCX."""
    if file:
        file_id = str(uuid.uuid4())
        filename = file.filename
        pdf_path = os.path.join(TEMP_DIR, f"{file_id}_{filename}")
        with open(pdf_path, "wb") as f:
            shutil.copyfileobj(file.file, f)
    elif file_id and filename:
        pdf_path = os.path.join(TEMP_DIR, f"{file_id}_{filename}")
        if not os.path.exists(pdf_path):
            raise HTTPException(status_code=404, detail="Uploaded file session expired. Please re-upload.")
    else:
        raise HTTPException(status_code=400, detail="No file provided for conversion.")

    base_name, _ = os.path.splitext(filename)
    docx_filename = f"{base_name}.docx"
    docx_path = os.path.join(TEMP_DIR, f"{file_id}_{docx_filename}")

    try:
        converter = OmniConverter(
            convert_math=convert_math,
            highlight_code=highlight_code,
            extract_images=extract_images,
            image_dpi=image_dpi,
        )

        mode_enum = ConversionMode(mode.lower())
        result = converter.convert(
            pdf_path=pdf_path,
            docx_path=docx_path,
            mode=mode_enum,
        )

        return JSONResponse(content={
            "status": "success",
            "file_id": file_id,
            "download_filename": docx_filename,
            "download_url": f"/api/download/{file_id}_{docx_filename}",
            "result": result,
        })
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Conversion error: {str(e)}")


@app.get("/api/download/{stored_filename}")
async def download_file(stored_filename: str):
    """Download the converted DOCX file."""
    file_path = os.path.join(TEMP_DIR, stored_filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="Requested file not found or expired.")

    # Extract user-facing download name
    parts = stored_filename.split("_", 1)
    user_name = parts[1] if len(parts) > 1 else stored_filename

    return FileResponse(
        path=file_path,
        filename=user_name,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
