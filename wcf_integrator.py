import os
import datetime
import pandas as pd
from typing import Optional, Dict, Any

# ตำแหน่ง Master Data ภายในโครงการนี้ (ลำดับความสำคัญสูงสุด)
LOCAL_DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
LOCAL_WCF_EXCEL = os.path.join(LOCAL_DATA_DIR, "DATA-EMP-WCF.xlsx")

# ตำแหน่ง Master Data อ้างอิงจากโครงการ SSO_WCF_Report_System (Fallback สำรอง)
DEFAULT_WCF_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "SSO_WCF_Report_System", "data"))
DEFAULT_WCF_EXCEL = os.path.join(DEFAULT_WCF_DIR, "DATA-EMP-WCF.xlsx")

def _strip_dot0(value) -> str:
    text = str(value or "").strip()
    return text[:-2] if text.endswith(".0") else text


class WCFIntegrator:
    """
    คลาสสำหรับเชื่อมโยงข้อมูลหลักสถานประกอบการ (Master Data) จาก DATA-EMP-WCF.xlsx
    โดยอ่านข้อมูล Sheet 'WCF' เพื่อดึงชื่อสถานประกอบการจริง, ที่อยู่, ตำบล, อำเภอ, จังหวัด, รหัสไปรษณีย์
    รองรับการตรวจจับการเปลี่ยนแปลงของไฟล์อัตโนมัติ (Auto-reload on file change)
    """

    def __init__(self, excel_path: Optional[str] = None):
        if excel_path and os.path.exists(excel_path):
            self.excel_path = excel_path
        elif os.path.exists(LOCAL_WCF_EXCEL):
            self.excel_path = LOCAL_WCF_EXCEL
        elif os.path.exists(DEFAULT_WCF_EXCEL):
            self.excel_path = DEFAULT_WCF_EXCEL
        else:
            self.excel_path = LOCAL_WCF_EXCEL

        self._master_dict: Dict[str, Dict[str, str]] = {}
        self._last_mtime: float = 0.0
        self.is_loaded = False

        if os.path.exists(self.excel_path):
            self.load_master_data()

    def check_and_reload_if_modified(self) -> bool:
        """ตรวจสอบว่าไฟล์ถูกแก้ไขหรือไม่ หากมีการแก้ไขจะโหลดข้อมูลใหม่ทันที"""
        if not self.excel_path or not os.path.exists(self.excel_path):
            return False

        try:
            current_mtime = os.path.getmtime(self.excel_path)
            if current_mtime != self._last_mtime:
                print(f"[*] พบการเปลี่ยนแปลงในไฟล์ Master Data ({self.excel_path}) กำลังโหลดข้อมูลใหม่...")
                return self.load_master_data()
            return True
        except Exception as e:
            print(f"[ERROR] ตรวจสอบไฟล์ Master Data ล้มเหลว: {e}")
            return False

    def load_master_data(self, excel_path: Optional[str] = None) -> bool:
        """โหลดข้อมูล WCF Master Data เข้าสู่ Memory Dictionary เพื่อการค้นหาความเร็วสูง"""
        target_path = excel_path or self.excel_path
        if not target_path or not os.path.exists(target_path):
            return False

        try:
            # ค้นหาแถว Header ของ Sheet 'WCF' อัตโนมัติ (เผื่อกรณีไฟล์มีคำอธิบายด้านบน)
            header_row = 0
            for r in range(10):
                try:
                    df_probe = pd.read_excel(target_path, sheet_name="WCF", nrows=1, skiprows=r)
                    cols = [str(c).strip() for c in df_probe.columns]
                    if any("เลขสปก." in c or "เลขที่บัญชี" in c for c in cols):
                        header_row = r
                        break
                except Exception:
                    break

            df = pd.read_excel(target_path, sheet_name="WCF", skiprows=header_row, dtype=str)
            df = df.fillna("")

            # สร้างคีย์ 'เลขสปก.' รูปแบบ 'เลขที่บัญชี;เลขที่สาขา'
            if "เลขสปก." not in df.columns:
                if "เลขที่บัญชี" in df.columns and "เลขที่สาขา" in df.columns:
                    df["เลขสปก."] = df["เลขที่บัญชี"].map(_strip_dot0) + ";" + df["เลขที่สาขา"].map(_strip_dot0)
                else:
                    return False

            master = {}
            for _, row in df.iterrows():
                raw_key = str(row.get("เลขสปก.", "")).strip()
                if not raw_key or raw_key.lower() == "nan":
                    continue

                if ";" in raw_key:
                    parts = raw_key.split(";", 1)
                    k_acc = _strip_dot0(parts[0])
                    k_br = _strip_dot0(parts[1])
                    k_br = k_br.zfill(6) if k_br.isdigit() else k_br
                    clean_key = f"{k_acc};{k_br}"
                else:
                    clean_key = raw_key

                zip_code = str(row.get("รหัสไปรษณีย์", "")).strip()
                if zip_code.endswith(".0"):
                    zip_code = zip_code[:-2]

                master[clean_key] = {
                    "ชื่อสปก_จริง": str(row.get("ชื่อสปก.", "")).strip(),
                    "ที่อยู่": str(row.get("ที่อยู่", "")).strip(),
                    "ตำบล": str(row.get("ตำบล", "")).strip(),
                    "อำเภอ": str(row.get("อำเภอ", "")).strip(),
                    "จังหวัด": str(row.get("จังหวัด", "")).strip(),
                    "รหัสไปรษณีย์": zip_code,
                    "โทรศัพท์": str(row.get("โทรศัพท์", "")).strip(),
                    "เลขพาณิชย์": str(row.get("เลขพาณิชย์", "")).strip()
                }

            self._master_dict = master
            self.excel_path = target_path
            self._last_mtime = os.path.getmtime(target_path)
            self.is_loaded = True
            print(f"[OK] โหลด WCF Master Data สำเร็จ: {len(master)} สถานประกอบการ")
            return True
        except Exception as e:
            print(f"[ERROR] ไม่สามารถโหลด WCF Master Data จาก {target_path}: {e}")
            return False

    def lookup(self, acc_no: str, branch_no: str) -> Dict[str, str]:
        """
        ค้นหาข้อมูลสถานประกอบการด้วย เลขที่บัญชี และ สาขา (เช่น 4700005424;000000)
        หากตรวจพบว่าไฟล์ Master Data มีการแก้ไข จะ reload อัตโนมัติก่อนค้นหา
        """
        self.check_and_reload_if_modified()

        acc_clean = _strip_dot0(acc_no)
        br_clean = _strip_dot0(branch_no)
        br_clean = br_clean.zfill(6) if br_clean.isdigit() else br_clean

        ref_key = f"{acc_clean};{br_clean}"
        if ref_key in self._master_dict:
            return self._master_dict[ref_key]

        # ไม่พบคู่ตรง คืนค่าว่างเปล่า
        return {
            "ชื่อสปก_จริง": "",
            "ที่อยู่": "",
            "ตำบล": "",
            "อำเภอ": "",
            "จังหวัด": "",
            "รหัสไปรษณีย์": "",
            "โทรศัพท์": "",
            "เลขพาณิชย์": ""
        }

    @property
    def total_records(self) -> int:
        return len(self._master_dict)

    @property
    def last_modified_str(self) -> str:
        if self._last_mtime and self.excel_path and os.path.exists(self.excel_path):
            dt = datetime.datetime.fromtimestamp(self._last_mtime)
            return dt.strftime("%d/%m/%Y %H:%M:%S")
        return "-"

    @property
    def file_size_kb(self) -> float:
        if self.excel_path and os.path.exists(self.excel_path):
            return round(os.path.getsize(self.excel_path) / 1024, 1)
        return 0.0
