import os
import shutil
import datetime
from typing import Optional
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from pydantic import BaseModel

from debt_processor import DebtReportProcessor
from wcf_integrator import WCFIntegrator, LOCAL_WCF_EXCEL, LOCAL_DATA_DIR

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
    if os.path.exists(index_path):
        with open(index_path, "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    return HTMLResponse("<h1>SSO Debt Report System</h1><p>Static index.html not found.</p>")

@app.get("/api/system-info")
async def get_system_info():
    """ส่งข้อมูลสถานะการเชื่อมโยงระบบ WCF Master Data และไฟล์ Input ที่มีอยู่"""
    # ตรวจสอบการอัปเดตไฟล์ Master Data อัตโนมัติ
    wcf_integrator.check_and_reload_if_modified()

    input_files = []
    if os.path.exists(INPUT_DIR):
        for fname in os.listdir(INPUT_DIR):
            if fname.lower().endswith(('.xlsx', '.xls', '.csv')):
                fpath = os.path.join(INPUT_DIR, fname)
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
    try:
        # บันทึกไฟล์ใหม่ทับที่ data/DATA-EMP-WCF.xlsx
        with open(target_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # โหลดเข้าสู่ Memory ทันที
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
    if not target_path.startswith(os.path.abspath(INPUT_DIR)) or not os.path.exists(target_path):
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
    if not file_path.startswith(os.path.abspath(OUTPUT_DIR)) or not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="ไม่พบไฟล์ที่ต้องการดาวน์โหลด")

    return FileResponse(
        path=file_path,
        filename=safe_filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8001)
