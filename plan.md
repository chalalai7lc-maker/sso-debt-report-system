# แผนการแก้ไขบั๊ก (Bug Fix Implementation Plan) — SSO_Debt_Report_System

## 1. ข้อมูลสภาพแวดล้อมและขอบเขตงาน (Environment & Scope)
- **ระบบปฏิบัติการเป้าหมาย**: Windows 11 (รันผ่าน `run.bat` หรือ PowerShell / Command Prompt)
- **โฟลเดอร์โปรเจกต์**: `SSO_Debt_Report_System`
- **ผู้ดำเนินการแก้ไขโค้ด**: Hermes Agent (ลลิสา)
- **ขอบเขตงาน**: แก้ไขบั๊กและจุดบกพร่องที่ตรวจพบทั้ง 10 รายการให้ครบถ้วน โดยยังไม่เพิ่มฟีเจอร์ใหม่นอกเหนือจากรายการแก้ไขบั๊กนี้

## 2. สรุปไฟล์ที่ต้องดำเนินการแก้ไข

| ลำดับ | ไฟล์ที่ต้องแก้ไข | รายการบั๊กที่เกี่ยวข้อง |
| :--- | :--- | :--- |
| 1 | `requirements.txt` | บั๊กข้อที่ 3 (เพิ่มไลบรารี `xlrd` รองรับไฟล์ `.xls` และ `pytest` สำหรับชุดทดสอบ) |
| 2 | `wcf_integrator.py` | บั๊กข้อที่ 6 (ทำความสะอาดคีย์ `เลขสปก.`, เติม `.zfill(6)` ให้รหัสสาขา, ตัด `.0` ท้ายรหัสไปรษณีย์ และสแกนหาแถวหัวตารางอัตโนมัติ) |
| 3 | `debt_processor.py` | บั๊กข้อที่ 1, 2, 3, 4, 5, 9 (แปลงตัวเลขที่มีคอมม่าคั่น, อ่าน CSV/XLS ที่มีหัวกระดาษหลายบรรทัด, ดึงคอลัมน์ขั้นตอนและวันที่บันทึกขั้นตอน, ส่งข้อมูลครบทุกแถวให้ตารางค้นหา, เพิ่มแถวสรุปรวมท้ายตาราง Excel) |
| 4 | `app.py` | บั๊กข้อที่ 7 (ป้องกัน Path Traversal ในทุก Endpoint ที่รับชื่อไฟล์จากภายนอก) |
| 5 | `static/js/app.js` | บั๊กข้อที่ 4, 8 (แก้ไอคอนเพี้ยน `ðŸ“„`, ป้องกัน XSS ด้วย `escapeHtml`, รองรับการค้นหาครบทุกรายการ) |
| 6 | `tests/test_debt_processor.py` | บั๊กข้อที่ 10 (ใช้ `tempfile.TemporaryDirectory()` ป้องกันไฟล์ขยะสะสมใน `Output/` และเพิ่มเคสทดสอบบั๊กที่แก้ไข) |

## 3. รายละเอียดการแก้ไขรายไฟล์ (Step-by-Step Instructions)

### ขั้นตอนที่ 1: แก้ไขไฟล์ `requirements.txt` (บั๊กข้อที่ 3)
- **ปัญหา**: ระบบเปิดรับไฟล์ `.xls` ในหน้าเว็บและ API แต่ไม่มีไลบรารี `xlrd` สำหรับอ่านไฟล์ Excel 97-2003
- **การแก้ไข**: อัปเดตไฟล์ `requirements.txt` ให้มีรายการดังนี้

```text
fastapi>=0.110.0
uvicorn>=0.28.0
pandas>=2.2.0
openpyxl>=3.1.2
xlrd>=2.0.1
python-multipart>=0.0.9
pydantic>=2.0.0
pytest>=8.0.0
```

### ขั้นตอนที่ 2: แก้ไขไฟล์ `wcf_integrator.py` (บั๊กข้อที่ 6)
- **ปัญหา**:
  1. การสร้างคีย์สำรองจาก `เลขที่บัญชี` + `เลขที่สาขา` (บรรทัด 69) ไม่ได้ตัด `.0` และไม่ได้เติม `.zfill(6)` ให้รหัสสาขา ทำให้คีย์ไม่ตรงกับฝั่งรายงานหนี้ (เช่น `4700020784;0` ไม่ตรงกับ `4700020784;000000`)
  2. การใช้ `.replace(".0", "")` กับรหัสไปรษณีย์ (บรรทัด 81) อาจตัดตัวเลขกลางข้อความผิดพลาด ควรตัดเฉพาะท้ายข้อความ
  3. ไฟล์ `DATA-EMP-WCF.xlsx` บางชุดมีบรรทัดแสดง Path/FileName อยู่ด้านบนก่อนบรรทัด Header จริง ทำให้ `pd.read_excel(..., sheet_name="WCF")` อ่านชื่อคอลัมน์ไม่พบ
- **การแก้ไข**:
  1. เพิ่มฟังก์ชัน `_clean_acc_and_branch(acc_val: Any, branch_val: Any) -> Tuple[str, str]` ด้านบนคลาส `WCFIntegrator` (อย่าลืม `from typing import Optional, Dict, Any, Tuple`)
  2. ปรับปรุง `load_master_data()` และ `lookup()` ตามโค้ดด้านล่าง

```python
def _clean_acc_and_branch(acc_val: Any, branch_val: Any) -> Tuple[str, str]:
    acc_str = str(acc_val or "").strip()
    if acc_str.endswith(".0"):
        acc_str = acc_str[:-2]
    branch_str = str(branch_val or "").strip()
    if branch_str.endswith(".0"):
        branch_str = branch_str[:-2]
    if branch_str.isdigit():
        branch_str = branch_str.zfill(6)
    return acc_str, branch_str
```

- ในเมธอด `load_master_data()` ปรับส่วนการอ่าน DataFrame และการสร้าง `master[key]`:

```python
        try:
            df = pd.read_excel(target_path, sheet_name="WCF", dtype=str)
            df = df.fillna("")

            # กรณีแถวแรกไม่ใช่ Header (เช่น มีแถว Path / FileName อยู่ด้านบน)
            if "เลขสปก." not in df.columns and "เลขที่บัญชี" not in df.columns:
                raw_df = pd.read_excel(target_path, sheet_name="WCF", header=None, dtype=str).fillna("")
                header_idx = None
                for idx in range(min(10, len(raw_df))):
                    row_values = [str(v).strip() for v in raw_df.iloc[idx].values]
                    if "เลขสปก." in row_values or "เลขที่บัญชี" in row_values:
                        header_idx = idx
                        break
                if header_idx is not None:
                    df = raw_df.iloc[header_idx + 1:].copy()
                    df.columns = [str(v).strip() for v in raw_df.iloc[header_idx].values]
                else:
                    return False

            master = {}
            for _, row in df.iterrows():
                raw_key = str(row.get("เลขสปก.", "")).strip()
                if ";" in raw_key:
                    parts = raw_key.split(";", 1)
                    acc_c, br_c = _clean_acc_and_branch(parts[0], parts[1])
                    key = f"{acc_c};{br_c}" if acc_c else ""
                else:
                    acc_c, br_c = _clean_acc_and_branch(row.get("เลขที่บัญชี", ""), row.get("เลขที่สาขา", ""))
                    key = f"{acc_c};{br_c}" if acc_c else ""

                if not key or not key.split(";")[0].isdigit():
                    continue

                zipcode = str(row.get("รหัสไปรษณีย์", "")).strip()
                if zipcode.endswith(".0"):
                    zipcode = zipcode[:-2]

                master[key] = {
                    "ชื่อสปก_จริง": str(row.get("ชื่อสปก.", "")).strip(),
                    "ที่อยู่": str(row.get("ที่อยู่", "")).strip(),
                    "ตำบล": str(row.get("ตำบล", "")).strip(),
                    "อำเภอ": str(row.get("อำเภอ", "")).strip(),
                    "จังหวัด": str(row.get("จังหวัด", "")).strip(),
                    "รหัสไปรษณีย์": zipcode,
                    "โทรศัพท์": str(row.get("โทรศัพท์", "")).strip(),
                    "เลขพาณิชย์": str(row.get("เลขพาณิชย์", "")).strip()
                }
```

- ในเมธอด `lookup(self, acc_no: str, branch_no: str)` ปรับการสร้าง `ref_key`:

```python
        acc_c, br_c = _clean_acc_and_branch(acc_no, branch_no)
        ref_key = f"{acc_c};{br_c}"
```

### ขั้นตอนที่ 3: แก้ไขไฟล์ `debt_processor.py` (บั๊กข้อที่ 1, 2, 3, 4, 5, 9)

#### 3.1 เพิ่มฟังก์ชัน `parse_float_safe` (แก้ไขบั๊กข้อที่ 1)
- **ปัญหา**: `float("220,668.00")` เกิด `ValueError` เมื่อข้อมูลมีเครื่องหมายจุลภาคคั่นหลักพัน ทำให้ในโหมด Excel ถูกปัดเป็น `0.00` เงียบๆ และในโหมด CSV โปรแกรมแครชทันที
- **การแก้ไข**: เพิ่มฟังก์ชัน `parse_float_safe(val: Any) -> float` ด้านบนไฟล์ `debt_processor.py`

```python
def parse_float_safe(val: Any) -> float:
    """แปลงค่าตัวเลขหรือข้อความจำนวนเงินที่มีเครื่องหมายจุลภาค (,) ให้เป็น float อย่างปลอดภัย"""
    if val is None or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        return 0.0 if pd.isna(val) else float(val)
    cleaned = str(val).replace(",", "").strip()
    if not cleaned or cleaned in ("-", "None", "nan"):
        return 0.0
    try:
        return float(cleaned)
    except ValueError:
        return 0.0
```

#### 3.2 ปรับปรุง `parse_debt_report()` ให้รองรับ `.xlsx`, `.xls`, และ `.csv` แบบมีหัวรายงานหลายบรรทัด พร้อมดึงคอลัมน์ `ขั้นตอน` (แก้ไขบั๊กข้อที่ 2, 3, 5)
- **ปัญหา**:
  1. ไฟล์ `.csv` และ `.xls` จากรายงาน `CN200090-1` มีหัวกระดาษ 4–5 บรรทัดแรก ทำให้ `pd.read_csv()` เดิมหาคอลัมน์ `"เลขที่บัญชี"` ไม่พบ
  2. โค้ดเดิมใช้ `openpyxl` กับไฟล์ `.xls` ทำให้เกิด `InvalidFileException`
  3. โค้ดเดิมข้ามคอลัมน์ที่ 18 (`ขั้นตอน`) และ 19 (`วันที่บันทึกขั้นตอน`)
- **การแก้ไข**: แทนที่เมธอด `parse_debt_report()` ด้วยโค้ดต่อไปนี้

```python
    def parse_debt_report(self, file_path: str) -> List[Dict[str, Any]]:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"ไม่พบไฟล์: {file_path}")

        ext = os.path.splitext(file_path)[1].lower()
        rows_matrix: List[List[Any]] = []

        if ext == ".xlsx":
            wb = openpyxl.load_workbook(file_path, data_only=True)
            ws = wb.active
            for r in range(1, ws.max_row + 1):
                rows_matrix.append([ws.cell(r, c).value for c in range(1, ws.max_column + 1)])
        elif ext == ".xls":
            df_raw = pd.read_excel(file_path, header=None, dtype=str).fillna("")
            rows_matrix = df_raw.values.tolist()
        elif ext == ".csv":
            df_raw = None
            for enc in ("utf-8-sig", "utf-8", "cp874", "tis-620"):
                try:
                    df_raw = pd.read_csv(file_path, header=None, dtype=str, encoding=enc, on_bad_lines="skip").fillna("")
                    break
                except Exception:
                    continue
            if df_raw is None:
                raise ValueError("ไม่สามารถอ่านไฟล์ CSV ได้ (Encoding ไม่รองรับ)")
            rows_matrix = df_raw.values.tolist()
        else:
            raise ValueError(f"ไม่รองรับนามสกุลไฟล์: {ext}")

        # ค้นหาบรรทัด Header ที่มีคำว่า 'เลขที่บัญชี'
        header_row_idx = 0
        col_map: Dict[str, int] = {}
        for r_idx, row in enumerate(rows_matrix[:20]):
            row_strs = [str(v or "").strip() for v in row]
            if any("เลขที่บัญชี" in s for s in row_strs):
                header_row_idx = r_idx
                for c_idx, val in enumerate(row_strs):
                    if "เลขที่บัญชี" in val and "acc" not in col_map:
                        col_map["acc"] = c_idx
                    elif val == "สาขา" or "เลขที่สาขา" in val:
                        col_map["branch"] = c_idx
                    elif "ชื่อสถานประกอบการ" in val:
                        col_map["name"] = c_idx
                    elif "สถานะ" in val and "สปก" in val:
                        col_map["status"] = c_idx
                    elif "งวดที่ค้าง" in val:
                        col_map["period"] = c_idx
                    elif "หนี้ประเมิน" in val:
                        col_map["est_debt"] = c_idx
                    elif "หนี้จริง" in val:
                        col_map["real_debt"] = c_idx
                    elif "เงินเพิ่ม" in val:
                        col_map["surcharge"] = c_idx
                    elif "รวมค้างชำระ" in val:
                        col_map["total_debt"] = c_idx
                    elif val == "ขั้นตอน" or "ขั้นตอนการติดตาม" in val:
                        col_map["step"] = c_idx
                    elif "วันที่บันทึก" in val:
                        col_map["step_date"] = c_idx
                    elif "ประเภทกิจการ" in val:
                        col_map["biz_type"] = c_idx
                break

        # ค่า Default ตำแหน่งคอลัมน์ตามโครงสร้างมาตรฐาน CN200090-1 (0-based index)
        c_acc = col_map.get("acc", 3)
        c_br = col_map.get("branch", 5)
        c_name = col_map.get("name", 6)
        c_status = col_map.get("status", 9)
        c_period = col_map.get("period", 12)
        c_est = col_map.get("est_debt", 13)
        c_real = col_map.get("real_debt", 14)
        c_sur = col_map.get("surcharge", 15)
        c_tot = col_map.get("total_debt", 16)
        c_step = col_map.get("step", 17)
        c_step_date = col_map.get("step_date", 18)
        c_biz = col_map.get("biz_type", 19)

        def _get_cell(row_list: List[Any], idx: int) -> Any:
            return row_list[idx] if 0 <= idx < len(row_list) else None

        records = []
        for r_idx in range(header_row_idx + 1, len(rows_matrix)):
            row = rows_matrix[r_idx]
            acc_raw = _get_cell(row, c_acc)
            if acc_raw is None:
                continue
            acc_str = str(acc_raw).strip()
            if acc_str.endswith(".0"):
                acc_str = acc_str[:-2]
            if not acc_str.isdigit() or acc_str in ("", "เลขที่บัญชี"):
                continue

            br_raw = _get_cell(row, c_br)
            branch_str = str(br_raw).strip() if br_raw is not None else ""
            if branch_str.endswith(".0"):
                branch_str = branch_str[:-2]
            branch_str = branch_str.zfill(6) if branch_str.isdigit() else branch_str

            name = str(_get_cell(row, c_name) or "").strip()
            status = str(_get_cell(row, c_status) or "").strip()
            period = str(_get_cell(row, c_period) or "").strip()

            est_debt = parse_float_safe(_get_cell(row, c_est))
            real_debt = parse_float_safe(_get_cell(row, c_real))
            surcharge = parse_float_safe(_get_cell(row, c_sur))
            raw_total = _get_cell(row, c_tot)
            total_debt = parse_float_safe(raw_total) if raw_total not in (None, "") else (est_debt + real_debt + surcharge)

            step = str(_get_cell(row, c_step) or "").strip()
            step_date = str(_get_cell(row, c_step_date) or "").strip()
            biz_type = str(_get_cell(row, c_biz) or "").strip()

            records.append({
                "row_index": r_idx + 1,
                "acc": acc_str,
                "branch": branch_str,
                "name_report": name,
                "status": status,
                "period": period,
                "est_debt": est_debt,
                "real_debt": real_debt,
                "surcharge": surcharge,
                "total_debt": total_debt,
                "step": step,
                "step_date": step_date,
                "biz_type": biz_type
            })

        return records
```

#### 3.3 ส่งข้อมูลครบทุกแถวใน `preview_data` และเพิ่มคอลัมน์ `ขั้นตอน` ใน `process()` (แก้ไขบั๊กข้อที่ 4 และ 5)
- ในเมธอด `process()` ของ `debt_processor.py`:
  1. ใน `groups = defaultdict(...)` เพิ่ม `"steps": []` และ `"step_date": ""`
  2. ในลูป `for rec in raw_records:` เพิ่ม:
     ```python
     if rec.get("step") and rec["step"] not in g["steps"]:
         g["steps"].append(rec["step"])
     if rec.get("step_date"):
         g["step_date"] = rec["step_date"]
     ```
  3. ใน `item = { ... }` เพิ่มคอลัมน์ `"ขั้นตอน": ", ".join(g["steps"])` และ `"วันที่บันทึกขั้นตอน": g["step_date"]` ต่อจาก `"สถานะ สปก."`
  4. ใน `return { ... }` บรรทัดที่ 274 เปลี่ยนจาก `"preview_data": df_result.head(100).to_dict(orient="records")` เป็น:
     ```python
     "preview_data": df_result.to_dict(orient="records")
     ```

#### 3.4 เพิ่มแถวสรุปยอดรวม (`รวมทั้งสิ้น`) ท้ายตารางใน `export_to_excel()` (แก้ไขบั๊กข้อที่ 9)
- ในเมธอด `export_to_excel()` ของ `debt_processor.py`:
  1. เพิ่ม `"ขั้นตอน"` ลงในเซต `center_cols`
  2. หลังจบลูปเขียนแถวข้อมูล (`for r_idx, row_data in enumerate(df.values, start=2):`) ให้เพิ่มโค้ดสร้างแถวสรุปยอดรวมก่อนส่วนปรับความกว้างคอลัมน์:

```python
        # เขียนแถวสรุปผลรวมท้ายตาราง (Summary Row)
        if len(df) > 0:
            sum_row_idx = len(df) + 2
            ws.row_dimensions[sum_row_idx].height = 24
            summary_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
            summary_font = Font(name="Tahoma", size=10, bold=True, color="0F172A")
            double_bottom_border = Border(
                left=Side(style="thin", color="CBD5E1"),
                right=Side(style="thin", color="CBD5E1"),
                top=Side(style="thin", color="64748B"),
                bottom=Side(style="double", color="0F172A")
            )

            for c_idx, col_name in enumerate(columns, start=1):
                cell = ws.cell(sum_row_idx, c_idx)
                cell.fill = summary_fill
                cell.font = summary_font
                cell.border = double_bottom_border

                if c_idx == 1:
                    cell.value = "รวมทั้งสิ้น"
                    cell.alignment = align_center
                elif col_name in currency_cols:
                    cell.value = round(float(df[col_name].sum()), 2)
                    cell.number_format = "#,##0.00"
                    cell.alignment = align_right
                elif col_name in integer_cols:
                    cell.value = int(df[col_name].sum())
                    cell.number_format = "#,##0"
                    cell.alignment = align_center
                else:
                    cell.value = ""
```

### ขั้นตอนที่ 4: แก้ไขไฟล์ `app.py` (บั๊กข้อที่ 7)
- **ปัญหา**: นำ `file.filename`, `req.filename`, และ `filename` มาต่อ Path ด้วย `os.path.join()` โดยไม่คัดกรอง `../` (Path Traversal)
- **การแก้ไข**:
  1. เพิ่มฟังก์ชัน `_safe_join(base_dir: str, filename: str) -> str` ใน `app.py`:
     ```python
     def _safe_join(base_dir: str, filename: str) -> str:
         safe_name = os.path.basename(filename or "")
         if not safe_name or safe_name in (".", ".."):
             raise HTTPException(status_code=400, detail="ชื่อไฟล์ไม่ถูกต้อง")
         full_path = os.path.abspath(os.path.join(base_dir, safe_name))
         if not full_path.startswith(os.path.abspath(base_dir)):
             raise HTTPException(status_code=400, detail="ไม่อนุญาตให้เข้าถึงพาธนอกโฟลเดอร์ที่กำหนด")
         return full_path
     ```
  2. เปลี่ยนจุดเรียกใช้ใน 3 Endpoints:
     - ใน `upload_and_process`: `temp_path = _safe_join(INPUT_DIR, file.filename)`
     - ใน `process_existing`: `target_path = _safe_join(INPUT_DIR, req.filename)`
     - ใน `download_file`: `file_path = _safe_join(OUTPUT_DIR, filename)`

### ขั้นตอนที่ 5: แก้ไขไฟล์ `static/js/app.js` (บั๊กข้อที่ 4 และ 8)
- **ปัญหา**:
  1. บรรทัดที่ 78 แสดงไอคอนเพี้ยนเป็น `ðŸ“„ ${f.filename}`
  2. การแสดงผลตารางและชื่อไฟล์ด้วย `innerHTML` ไม่ได้ทำ HTML Escaping
- **การแก้ไข**:
  1. เพิ่มฟังก์ชัน `escapeHtml(val)` ใน `static/js/app.js`:
     ```javascript
     function escapeHtml(val) {
         if (val === null || val === undefined) return "";
         return String(val)
             .replace(/&/g, "&amp;")
             .replace(/</g, "&lt;")
             .replace(/>/g, "&gt;")
             .replace(/"/g, "&quot;")
             .replace(/'/g, "&#039;");
     }
     ```
  2. แก้ไขบรรทัดที่ 76–80 ใน `renderInputFiles(files)`:
     ```javascript
     item.innerHTML = `
         <span class="file-item-name" title="${escapeHtml(f.filename)}">
             📄 ${escapeHtml(f.filename)}
         </span>
         <span class="file-item-size">${escapeHtml(f.size_kb)} KB</span>
     `;
     ```
  3. ใน `renderTable(records)` (บรรทัดที่ 202–220) ครอบค่าข้อความด้วย `escapeHtml(...)` ทุกคอลัมน์

### ขั้นตอนที่ 6: แก้ไขไฟล์ `tests/test_debt_processor.py` (บั๊กข้อที่ 10)
- **ปัญหา**: `test_process_actual_file` สร้างไฟล์จริงทิ้งไว้ในโฟลเดอร์ `Output/` ทุกครั้งที่รันทดสอบ และจำนวนแถวใน Excel เพิ่มขึ้น 1 แถวจากแถวสรุปยอดรวม (`รวมทั้งสิ้น`)
- **การแก้ไข**:
  1. `import tempfile` ที่หัวไฟล์ `tests/test_debt_processor.py`
  2. ใน `test_process_actual_file` ให้รันภายใต้ `with tempfile.TemporaryDirectory() as tmp_dir:` แล้วเรียก `result = processor.process(input_path, output_dir=tmp_dir)`
  3. ปรับ `self.assertEqual(ws.max_row, 153)` เป็น `self.assertEqual(ws.max_row, 154)` (1 Header + 152 Data Rows + 1 Summary Row) และตรวจค่า `ws.cell(154, 1).value == "รวมทั้งสิ้น"`
  4. เพิ่มเคสทดสอบ `test_parse_float_safe_with_commas` เพื่อยืนยันการแปลงค่า `"220,668.00"`

## 4. ขั้นตอนการตรวจสอบหลังแก้ไขบน Windows 11

1. ติดตั้งไลบรารีเพิ่มเติมในสภาพแวดล้อมของโปรเจกต์:
   ```powershell
   pip install -r requirements.txt
   ```
2. รันชุดทดสอบทั้งหมดเพื่อยืนยันความถูกต้อง:
   ```powershell
   pytest -v
   ```
3. ทดสอบเปิดใช้งานระบบผ่านสคริปต์เริ่มต้น:
   ```powershell
   .\run.bat
   ```

คำแนะนำและคำเตือนด้านความปลอดภัยสำหรับรันบน Windows 11

ตรวจสอบ Virtual Environment ก่อนติดตั้งแพ็กเกจ: ก่อนรัน pip install -r requirements.txt บน Windows 11 ควรเปิดใช้งาน Virtual Environment ของโปรเจกต์ก่อนเสมอ เพื่อป้องกันการเขียนทับเวอร์ชันไลบรารี (pandas, openpyxl, fastapi) ของระบบหลัก

จำกัดสิทธิ์การแก้ไขไฟล์ของ Hermes Agent: กำหนดขอบเขตโฟลเดอร์ทำงานของ Hermes Agent (ลลิสา) ให้อยู่เฉพาะในโฟลเดอร์ SSO_Debt_Report_System เท่านั้น เพื่อป้องกันการแก้ไขไฟล์นอกโปรเจกต์โดยไม่ตั้งใจ