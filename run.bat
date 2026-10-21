@echo off
chcp 65001 > nul
title SSO Debt Report System - ระบบประมวลผลรายงานหนี้ค้างชำระ

cd /d "%~dp0"

echo ===============================================================
echo     SSO Debt Report & Data Transformation System
echo     ระบบประมวลผลรายงานหนี้ค้างชำระสถานประกอบการ
echo ===============================================================
echo.
echo [*] ตรวจสอบสภาพแวดล้อม Python...
python --version > nul 2>&1
if errorlevel 1 (
    echo [ERROR] ไม่พบคำสั่ง Python ในระบบ (กรุณาติดตั้ง Python หรือตั้งค่า PATH)
    pause
    exit /b 1
)

echo [*] กำลังเปิดเว็บเบราว์เซอร์ไปยัง http://localhost:8001/ ...
start "" /b cmd /c "timeout /t 2 /nobreak > nul && start http://localhost:8001/"

echo.
echo ===============================================================
echo  ระบบพร้อมใช้งาน! เปิดหน้าเว็บที่: http://localhost:8001/
echo  (หากต้องการปิดระบบ ให้ปิดหน้าต่าง Command Prompt นี้ หรือกด Ctrl+C)
echo ===============================================================
echo.

python app.py

pause
