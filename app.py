import os
import shutil
import datetime
import tempfile
from typing import List, Optional
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from pydantic import BaseModel, Field

from debt_processor import DebtReportProcessor
from wcf_integrator import WCFIntegrator, LOCAL_WCF_EXCEL, LOCAL_DATA_DIR
from letter_service import (
    build_letter_docx,
    docx_to_pdf,
    load_letter_config,
    save_letter_config,
    DEFAULT_LETTER_CONFIG,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "Input")
OUTPUT_DIR = os.path.join(BASE_DIR, "Output")
STATIC_DIR = os.path.join(BASE_DIR, "static")

os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(LOCAL_DATA_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)

app = FastAPI(title="SSO Debt Report System", version="1.0.0")

# Serve static assets
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

# Shared Processor & WCF Integrator Instance
wcf_integrator = WCFIntegrator()
processor = DebtReportProcessor(wcf_integrator=wcf_integrator)

@app.get("/", response_class=HTMLResponse)
async def serve_index():
    index_path = os.path.join(STATIC_DIR, "index.html")
    headers = {"Cache-Control": "no-cache, no-store, must-revalidate"}
    if os.path.exists(index_path):
        with open(index_path, "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read(), headers=headers)
    return HTMLResponse("<h1>SSO Debt Report System</h1><p>Static index.html not found.</p>", headers=headers)

@app.get("/api/system-info")
async def get_system_info():
    """ส่งข้อมูลสถานะการเชื่อมโยงระบบ WCF Master Data และไฟล์ Input ที่มีอยู่"""
    # ตรวจสอบการอัปเดตไฟล์ Master Data อัตโนมัติ
    wcf_integrator.check_and_reload_if_modified()

    input_files = []
    if os.path.isdir(INPUT_DIR):
        for fname in os.listdir(INPUT_DIR):
            if fname.startswith(("~$", ".")):
                continue
            if not fname.lower().endswith((".xlsx", ".xls", ".csv")):
                continue
            fpath = os.path.join(INPUT_DIR, fname)
            if not os.path.isfile(fpath):
                continue
            size_kb = round(os.path.getsize(fpath) / 1024, 1)
            input_files.append({
                "filename": fname,
                "size_kb": size_kb
            })

    return {
        "wcf_connected": wcf_integrator.is_loaded,
        "wcf_records_count": wcf_integrator.total_records,
        "wcf_source_path": wcf_integrator.excel_path or "ไม่พบไฟล์ DATA-EMP-WCF.xlsx",
        "wcf_last_modified": wcf_integrator.last_modified_str,
        "wcf_size_kb": wcf_integrator.file_size_kb,
        "input_files": input_files
    }

@app.post("/api/upload-master-data")
async def upload_master_data(file: UploadFile = File(...)):
    """
    รับไฟล์ DATA-EMP-WCF.xlsx ใหม่ และอัปเดตฐานข้อมูล Master Data ทันที
    """
    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="กรุณาอัปโหลดไฟล์ Master Data นามสกุล .xlsx เท่านั้น")

    target_path = LOCAL_WCF_EXCEL
    suffix = os.path.splitext(file.filename or "")[1].lower() or ".xlsx"
    fd, tmp_path = tempfile.mkstemp(suffix=suffix)
    os.close(fd)
    try:
        with open(tmp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        probe = WCFIntegrator(tmp_path)
        if not probe.is_loaded or probe.total_records < 1:
            raise HTTPException(status_code=400, detail="ไฟล์ที่อัปโหลดไม่มี Sheet 'WCF' หรือโครงสร้างข้อมูลไม่ถูกต้อง")
        os.makedirs(os.path.dirname(target_path), exist_ok=True)
        os.replace(tmp_path, target_path)
        tmp_path = ""
        success = wcf_integrator.load_master_data(target_path)
        if not success:
            raise HTTPException(status_code=400, detail="ไฟล์ที่อัปโหลดไม่มี Sheet 'WCF' หรือโครงสร้างข้อมูลไม่ถูกต้อง")

        return JSONResponse(content={
            "success": True,
            "message": f"อัปเดต Master Data สำเร็จ! โหลดข้อมูลแล้ว {wcf_integrator.total_records:,} สถานประกอบการ",
            "wcf_records_count": wcf_integrator.total_records,
            "wcf_last_modified": wcf_integrator.last_modified_str,
            "wcf_size_kb": wcf_integrator.file_size_kb
        })
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"เกิดข้อผิดพลาดในการบันทึก Master Data: {str(e)}")
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.remove(tmp_path)

@app.post("/api/reload-master-data")
async def reload_master_data():
    """บังคับโหลดข้อมูล Master Data ใหม่อีกครั้งจากไฟล์เดิม"""
    success = wcf_integrator.load_master_data()
    return JSONResponse(content={
        "success": success,
        "wcf_records_count": wcf_integrator.total_records,
        "wcf_last_modified": wcf_integrator.last_modified_str,
        "wcf_size_kb": wcf_integrator.file_size_kb,
        "message": f"รีโหลดข้อมูลสำเร็จ: {wcf_integrator.total_records:,} สถานประกอบการ" if success else "ไม่สามารถรีโหลดข้อมูลได้"
    })

@app.post("/api/upload-and-process")
async def upload_and_process(file: UploadFile = File(...)):
    """รับไฟล์รายงานหนี้อัปโหลด ประมวลผล และส่งผลลัพธ์กลับแบบ JSON"""
    if not file.filename.lower().endswith(('.xlsx', '.xls', '.csv')):
        raise HTTPException(status_code=400, detail="กรุณาอัปโหลดไฟล์นามสกุล .xlsx, .xls หรือ .csv เท่านั้น")

    safe_filename = os.path.basename(file.filename)
    temp_path = os.path.join(INPUT_DIR, safe_filename)
    try:
        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        result = processor.process(temp_path, output_dir=OUTPUT_DIR)
        return JSONResponse(content=result)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"เกิดข้อผิดพลาดในการประมวลผล: {str(e)}")

class ExistingProcessRequest(BaseModel):
    filename: str

@app.post("/api/process-existing")
async def process_existing(req: ExistingProcessRequest):
    """ประมวลผลไฟล์ที่มีอยู่แล้วในโฟลเดอร์ Input"""
    safe_filename = os.path.basename(req.filename)
    target_path = os.path.abspath(os.path.join(INPUT_DIR, safe_filename))
    if not target_path.startswith(os.path.abspath(INPUT_DIR)) or not os.path.isfile(target_path):
        raise HTTPException(status_code=404, detail="ไม่พบไฟล์ที่ระบุในโฟลเดอร์ Input")

    try:
        result = processor.process(target_path, output_dir=OUTPUT_DIR)
        return JSONResponse(content=result)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"เกิดข้อผิดพลาดในการประมวลผล: {str(e)}")

@app.get("/api/download/{filename}")
async def download_file(filename: str):
    """ดาวน์โหลดไฟล์ผลลัพธ์ Excel"""
    safe_filename = os.path.basename(filename)
    file_path = os.path.abspath(os.path.join(OUTPUT_DIR, safe_filename))
    if not file_path.startswith(os.path.abspath(OUTPUT_DIR)) or not os.path.isfile(file_path):
        raise HTTPException(status_code=404, detail="ไม่พบไฟล์ที่ต้องการดาวน์โหลด")

    return FileResponse(
        path=file_path,
        filename=safe_filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


class LetterCaseIn(BaseModel):
    name: str = ""
    account: str
    branch: str = ""
    periods: List[str] = Field(default_factory=list)
    street: str = ""
    district: str = ""
    province: str = ""
    zipcode: str = ""


class LetterConfigIn(BaseModel):
    doc_code: Optional[str] = None
    office_name: Optional[str] = None
    office_location: Optional[str] = None
    office_address: Optional[str] = None
    signatory_name: Optional[str] = None
    signatory_position: Optional[str] = None
    owner_division: Optional[str] = None
    phone: Optional[str] = None
    fax: Optional[str] = None


@app.get("/api/letter-config")
def get_letter_config():
    """ดึงข้อมูลตั้งค่าหนังสือราชการ 8 รายการ พร้อมค่าเริ่มต้น."""
    return {
        "status": "success",
        "config": load_letter_config(),
        "defaults": DEFAULT_LETTER_CONFIG,
    }


@app.post("/api/letter-config")
def update_letter_config(cfg: LetterConfigIn):
    """บันทึกข้อมูลตั้งค่าหนังสือราชการ 8 รายการ."""
    saved = save_letter_config(cfg.model_dump(exclude_unset=True))
    return {
        "status": "success",
        "message": "บันทึกข้อมูลหนังสือราชการสำเร็จ",
        "config": saved,
    }


class LetterGenerateIn(BaseModel):
    format: str
    issue_date: str
    appoint_date: str
    warn_start: int
    invite_start: Optional[int] = None
    config: Optional[LetterConfigIn] = None
    cases: List[LetterCaseIn]


@app.post("/api/letters/generate")
def generate_letters(req: LetterGenerateIn):
    """พิมพ์หนังสือเตือน+เชิญตามรายที่เลือก เป็น Word หรือ PDF."""
    fmt = (req.format or "").lower().strip()
    if fmt not in ("docx", "pdf"):
        raise HTTPException(status_code=400, detail="รูปแบบต้องเป็น docx หรือ pdf")
    if not req.cases:
        raise HTTPException(status_code=400, detail="ไม่มีรายการที่จะพิมพ์")
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    docx_name = f"Letter_Batch_{stamp}.docx"
    docx_path = os.path.join(OUTPUT_DIR, docx_name)
    try:
        cfg_dict = req.config.model_dump(exclude_unset=True) if req.config else None
        build_letter_docx(
            [case.model_dump() for case in req.cases],
            issue_date=req.issue_date,
            appoint_date=req.appoint_date,
            warn_start=req.warn_start,
            invite_start=req.invite_start,
            output_path=docx_path,
            config=cfg_dict,
        )
        if fmt == "pdf":
            pdf_name = f"Letter_Batch_{stamp}.pdf"
            pdf_path = os.path.join(OUTPUT_DIR, pdf_name)
            docx_to_pdf(docx_path, pdf_path)
            return FileResponse(path=pdf_path, filename=pdf_name, media_type="application/pdf")
        return FileResponse(
            path=docx_path,
            filename=docx_name,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except Exception:
        raise HTTPException(status_code=500, detail="สร้างหนังสือไม่สำเร็จ")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8001)
