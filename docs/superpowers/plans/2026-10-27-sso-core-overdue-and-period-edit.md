# SSO CORE Overdue Check & Overdue Period Editor Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** เพิ่มฟีเจอร์ตรวจสอบงวดค้างชำระจริงจาก SSO CORE แบบเลือกตรวจสอบได้ในหน้ารายงานหนี้ และเพิ่มฟีเจอร์แก้ไขงวดค้างชำระรายบัญชีที่ซิงค์ข้อมูลทั้งหน้ารายงานหนี้ หน้าพิมพ์หนังสือ และไฟล์ Excel ส่งออก

**Architecture:** นำ `overdue_calculator.py` และ `sso_client.py` จาก `SSO_M33_Overdue_System` มาติดตั้งใน `SSO_Debt_Report_System` เพิ่ม API `/api/sso-core/check` และ `/api/records/re-export` ใน `app.py` และเพิ่ม UI ปุ่มตรวจสอบ SSO CORE, การสลับแหล่งข้อมูล Excel/SSO CORE, และ Modal แก้ไขงวดค้างชำระใน `index.html`, `app.js`, และ `letters.js`

**Tech Stack:** Python (FastAPI, Pydantic, httpx, Playwright, openpyxl, pandas), Vanilla JS, HTML5/CSS3, pytest

---

### Task 1: SSO CORE Client & Overdue Calculator Integration + Backend Endpoints

**Files:**
- Create: `projects/SSO_Debt_Report_System/overdue_calculator.py`
- Create: `projects/SSO_Debt_Report_System/sso_client.py`
- Modify: `projects/SSO_Debt_Report_System/debt_processor.py`
- Modify: `projects/SSO_Debt_Report_System/app.py`
- Test: `projects/SSO_Debt_Report_System/tests/test_sso_core_and_edit.py`

- [ ] **Step 1: Write failing tests for SSO CORE check endpoint, period formatting/normalization, and re-export endpoint**
- [ ] **Step 2: Run pytest to verify tests fail**
- [ ] **Step 3: Implement `overdue_calculator.py`, `sso_client.py`, `debt_processor.py` enhancements, and `/api/sso-core/check` + `/api/records/re-export` in `app.py`**
- [ ] **Step 4: Run pytest to verify all backend tests pass**

---

### Task 2: Frontend UI for SSO CORE Verification & Period Editor (Debt Desk + Letter Desk)

**Files:**
- Modify: `projects/SSO_Debt_Report_System/static/index.html`
- Modify: `projects/SSO_Debt_Report_System/static/css/style.css`
- Modify: `projects/SSO_Debt_Report_System/static/js/app.js`
- Modify: `projects/SSO_Debt_Report_System/static/js/letters.js`
- Test: `projects/SSO_Debt_Report_System/tests/test_sso_core_and_edit.py`

- [ ] **Step 1: Add UI tests verifying HTML controls and JS hooks exist**
- [ ] **Step 2: Update `index.html` and `style.css` with SSO CORE toolbar, source badges, and Edit Periods modal**
- [ ] **Step 3: Update `app.js` and `letters.js` to support SSO CORE batch verification, source switching (Excel vs SSO CORE), period editing modal, default >=1 period filtering in Letter Desk, and automatic Excel re-export**
- [ ] **Step 4: Run full pytest suite and verify all tests pass**

