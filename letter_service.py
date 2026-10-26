"""สร้างหนังสือเตือน + เชิญจากฟอร์มราชการ โดยไม่ทิ้งข้อมูลตัวอย่างของต้นฉบับ."""

import datetime
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from copy import deepcopy
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from lxml import etree

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "letter_warn_invite.docx")
CONFIG_PATH = os.path.join(BASE_DIR, "data", "letter_config.json")

DEFAULT_LETTER_CONFIG = {
    "doc_code": "สน ๐๐๓1.1/สท",
    "office_name": "สำนักงานประกันสังคมจังหวัดสกลนคร สาขาสว่างแดนดิน",
    "office_location": "ถนนนิตโย อำเภอสว่างแดนดิน สน 47110",
    "office_address": "เลขที่ 273/2 หมู่ที่ 20 ตำบลสว่างแดนดิน อำเภอสว่างแดนดิน จังหวัดสกลนคร",
    "signatory_name": "นายชลาลัย ประโมสี",
    "signatory_position": "นักวิชาการแรงงานปฏิบัติการ",
    "owner_division": "ส่วนงานเงินสมทบและการตรวจสอบ",
    "phone": "0 4277 6833 ต่อ 105",
    "fax": "0 4277 6834",
}


def load_letter_config() -> Dict[str, str]:
    """โหลดการตั้งค่าข้อมูลหนังสือราชการ หากไม่มีหรืออ่านไม่ได้ให้ใช้ default."""
    cfg = dict(DEFAULT_LETTER_CONFIG)
    if os.path.exists(CONFIG_PATH):
        try:
            import json
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                saved = json.load(f)
                if isinstance(saved, dict):
                    for k in DEFAULT_LETTER_CONFIG:
                        if k in saved and saved[k] is not None:
                            cfg[k] = str(saved[k]).strip()
        except Exception:
            pass
    return cfg


def save_letter_config(cfg: Dict[str, Any]) -> Dict[str, str]:
    """บันทึกการตั้งค่าข้อมูลหนังสือราชการลงไฟล์ data/letter_config.json."""
    import json
    merged = dict(DEFAULT_LETTER_CONFIG)
    for k in DEFAULT_LETTER_CONFIG:
        if k in cfg and cfg[k] is not None:
            merged[k] = str(cfg[k]).strip()
    os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(merged, f, ensure_ascii=False, indent=2)
    return merged


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
NS = {"w": W_NS}

THAI_MONTHS = [
    "", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม",
]
# ข้อความตัวอย่างในไฟล์ต้นฉบับ — ใช้ตอนทำ template เท่านั้น ห้ามเหลือในไฟล์ที่พิมพ์
_SAMPLE_FULL_NAME = "โรงเรียนวัดพระธาตุศรีมงคล และ ศูนย์อบรมเด็กก่อนเกณฑ์วัดพระธาตุศรีมงคล"
_SAMPLE_ENV_NAME = "โรงเรียนวัดพระธาตุศรีมงคล และ ศูนย์อบรมเ…"
_SAMPLE_STREET = "125 ม. 3 ถนนวาริชภูมิ-พังโคน"
_SAMPLE_ACCOUNT = "4790001695"


def assign_numbers(count: int, warn_start: int, invite_start: Optional[int]) -> List[Tuple[int, int]]:
    """คืน (เลขเตือน, เลขเชิญ) ตามลำดับราย.

    กรอกเฉพาะเลขเตือน: สลับคู่ในชุดเดียว
        รายที่ i → เตือน = start + 2i, เชิญ = start + 2i + 1
    กรอกทั้งสองช่อง: คนละชุด
        เตือน = warn_start + i, เชิญ = invite_start + i
    """
    if count < 0:
        raise ValueError("จำนวนรายติดลบไม่ได้")
    if warn_start < 1:
        raise ValueError("เลขเริ่มต้นหนังสือเตือนต้องเป็นจำนวนเต็มตั้งแต่ 1")
    if invite_start is not None and invite_start < 1:
        raise ValueError("เลขเริ่มต้นหนังสือเชิญต้องเป็นจำนวนเต็มตั้งแต่ 1")
    rows = []
    for i in range(count):
        if invite_start is None:
            rows.append((warn_start + 2 * i, warn_start + 2 * i + 1))
        else:
            rows.append((warn_start + i, invite_start + i))
    return rows


def format_doc_no(number: int) -> str:
    """เลขในช่อง สท ตามที่กรอก ไม่เติมศูนย์หน้า."""
    if number < 1:
        raise ValueError("เลขหนังสือต้องเป็นจำนวนเต็มตั้งแต่ 1")
    return str(number)


def format_periods_letter(period_list: List[str]) -> str:
    """งวดในหนังสือแบบช่วงเต็ม: มิถุนายน-สิงหาคม 2569. ปีซ้ำไม่พิมพ์สองครั้ง."""
    parsed = []
    for raw in period_list or []:
        parts = str(raw).strip().split("/")
        if len(parts) != 2:
            continue
        try:
            month = int(parts[0])
            year = int(parts[1])
        except ValueError:
            continue
        if 1 <= month <= 12:
            parsed.append((year, month))
    parsed = sorted(set(parsed))
    if not parsed:
        return "—"
    groups: List[List[Tuple[int, int]]] = [[parsed[0]]]
    for item in parsed[1:]:
        prev_year, prev_month = groups[-1][-1]
        year, month = item
        if (year - prev_year) * 12 + (month - prev_month) == 1:
            groups[-1].append(item)
        else:
            groups.append([item])
    parts = []
    for group in groups:
        start_year, start_month = group[0]
        end_year, end_month = group[-1]
        if len(group) == 1:
            parts.append(f"{THAI_MONTHS[start_month]} {start_year}")
        elif start_year == end_year:
            parts.append(f"{THAI_MONTHS[start_month]}-{THAI_MONTHS[end_month]} {end_year}")
        else:
            parts.append(
                f"{THAI_MONTHS[start_month]} {start_year}-{THAI_MONTHS[end_month]} {end_year}"
            )
    return ", ".join(parts)


def parse_ce_iso(value: str) -> Tuple[int, int, int]:
    """YYYY-MM-DD (ค.ศ.) → (วัน, เดือน, ปี พ.ศ.)."""
    text = (value or "").strip()
    match = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", text)
    if not match:
        raise ValueError("วันที่ต้องเป็นรูปแบบ YYYY-MM-DD")
    year_ce, month, day = (int(match.group(1)), int(match.group(2)), int(match.group(3)))
    if not (1900 <= year_ce <= 2400):
        raise ValueError("วันที่ไม่ถูกต้อง")
    try:
        datetime.date(year_ce, month, day)
    except ValueError as exc:
        raise ValueError("วันที่ไม่ถูกต้อง") from exc
    return day, month, year_ce + 543


def format_issue_date(iso_date: str) -> str:
    day, month, year_be = parse_ce_iso(iso_date)
    return f"{day} {THAI_MONTHS[month]} {year_be}"


def format_appt_date(iso_date: str) -> str:
    day, month, year_be = parse_ce_iso(iso_date)
    return f"{day:02d}/{month:02d}/{year_be}"


def _strip_float_suffix(value: Any) -> str:
    text = str(value or "").strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def _digits(value: Any) -> str:
    return re.sub(r"\D", "", _strip_float_suffix(value))


def norm_account(value: Any) -> str:
    return _digits(value)


def norm_branch(value: Any) -> str:
    digits = _digits(value) or "0"
    return digits.zfill(6)


def norm_zip(value: Any) -> str:
    digits = _digits(value)
    return "" if set(digits) <= {"0"} else digits


def _is_bangkok(province: str) -> bool:
    cleaned = (province or "").strip()
    return "กรุงเทพ" in cleaned or cleaned in ("กทม", "กทม.")


def _strip_prefix(text: str, prefixes: Tuple[str, ...]) -> str:
    cleaned = (text or "").strip()
    for prefix in prefixes:
        if cleaned.startswith(prefix):
            cleaned = cleaned[len(prefix):].strip()
    return cleaned


def envelope_name(name: str, limit: int = 36) -> str:
    text = re.sub(r"\s+", " ", (name or "").strip()) or "สถานประกอบการ"
    if len(text) <= limit:
        return text
    return text[: limit - 1].rstrip() + "…"


def _set_exact(root, old: str, new: str, expected_min: int = 1) -> int:
    count = 0
    for node in root.iter(f"{W}t"):
        if node.text == old:
            node.text = new
            count += 1
    if count < expected_min:
        raise ValueError(f"ไม่พบข้อความ '{old}' ในฟอร์ม (พบ {count})")
    return count


def _paragraphs(root):
    body = root.find("w:body", NS)
    return [child for child in list(body) if child.tag == f"{W}p"]


def _para_text(paragraph) -> str:
    return "".join(node.text or "" for node in paragraph.iter(f"{W}t"))


def tokenize_template(source_path: str, dest_path: str) -> None:
    """แปลงไฟล์ตัวอย่างเป็น template ที่มี token และไม่มีข้อมูลนายจ้างตัวอย่าง."""
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    with zipfile.ZipFile(source_path) as zin:
        xml = zin.read("word/document.xml")
        other = {name: zin.read(name) for name in zin.namelist() if name != "word/document.xml"}
    root = etree.fromstring(xml)

    for paragraph in _paragraphs(root):
        text = _para_text(paragraph)
        if "การนำเงินสมทบเดือน" not in text:
            continue
        replaced = False
        for node in paragraph.iter(f"{W}t"):
            if node.text == "ก.ค.":
                node.text = "{{PERIODS}}"
                replaced = True
            elif node.text in {",", "ส.ค. ", "2569"}:
                node.text = ""
        if not replaced:
            raise ValueError("ไม่พบช่องงวดเดือนในย่อหน้าหนังสือ")

    for paragraph in _paragraphs(root):
        if "เลขที่บัญชี" not in _para_text(paragraph):
            continue
        for node in paragraph.iter(f"{W}t"):
            if node.text == "000000":
                node.text = "{{BRANCH}}"

    for paragraph in _paragraphs(root):
        if _SAMPLE_STREET not in _para_text(paragraph):
            continue
        for node in paragraph.iter(f"{W}t"):
            if node.text == "สกลนคร":
                node.text = "{{PROVINCE}}"

    _set_exact(root, "01747", "{{WARN_NO}}")
    _set_exact(root, "01748", "{{INVITE_NO}}")
    _set_exact(root, "18 กันยายน 2569", "{{ISSUE_DATE}}", 2)
    _set_exact(root, "18/10/2569", "{{APPT}}")
    _set_exact(root, _SAMPLE_ACCOUNT, "{{ACC}}", 2)
    _set_exact(root, _SAMPLE_FULL_NAME, "{{NAME}}", 1)
    _set_exact(root, _SAMPLE_ENV_NAME, "{{ENV_NAME}}", 2)
    _set_exact(root, _SAMPLE_STREET, "{{STREET}}", 2)
    _set_exact(root, "วาริชภูมิ", "{{DISTRICT}}", 2)
    _set_exact(root, "47150", "{{ZIP}}", 2)
    _set_exact(root, "ผู้อำนวยการ", "", 1)

    required = ("{{WARN_NO}}", "{{INVITE_NO}}", "{{ISSUE_DATE}}", "{{APPT}}", "{{ACC}}", "{{BRANCH}}", "{{PERIODS}}", "{{NAME}}", "{{ENV_NAME}}", "{{STREET}}", "{{DISTRICT}}", "{{PROVINCE}}", "{{ZIP}}")
    flat = "".join(node.text or "" for node in root.iter(f"{W}t"))
    missing = [token for token in required if token not in flat]
    if missing:
        raise ValueError("template ขาด token: " + ", ".join(missing))
    leaked = [token for token in (_SAMPLE_ACCOUNT, "01747", "01748", "โรงเรียนวัดพระธาตุ", "47150", "18/10/2569") if token in flat]
    if leaked:
        raise ValueError("template ยังเหลือข้อมูลตัวอย่าง")

    new_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(dest_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, blob in other.items():
            zout.writestr(name, blob)
        zout.writestr("word/document.xml", new_xml)


def _page_break():
    paragraph = etree.Element(f"{W}p")
    run = etree.SubElement(paragraph, f"{W}r")
    br = etree.SubElement(run, f"{W}br")
    br.set(f"{W}type", "page")
    return paragraph


def _unwrap_sdts(root) -> None:
    sdt_tag = f"{W}sdt"
    content_tag = f"{W}sdtContent"
    for sdt in list(root.iter(sdt_tag)):
        parent = sdt.getparent()
        if parent is None:
            continue
        content = sdt.find(content_tag)
        children = list(content) if content is not None else []
        index = list(parent).index(sdt)
        parent.remove(sdt)
        for offset, child in enumerate(children):
            parent.insert(index + offset, child)


def _drop_hidden_shapes(root) -> None:
    for node in list(root.iter()):
        if not str(node.tag).endswith("}docPr") or node.get("hidden") != "1":
            continue
        run = node
        while run is not None and not (str(run.tag).endswith("}r") and run.getparent() is not None and str(run.getparent().tag).endswith("}p")):
            run = run.getparent()
        if run is not None and run.getparent() is not None:
            run.getparent().remove(run)


def _strip_merge_fields(root) -> None:
    for r in list(root.iter(f"{W}r")):
        parent = r.getparent()
        if parent is None:
            continue
        if r.find(f"{W}fldChar") is not None or r.find(f"{W}instrText") is not None:
            parent.remove(r)
            continue
        children = [c for c in r if c.tag != f"{W}rPr"]
        if len(children) == 1 and children[0].tag == f"{W}t" and not (children[0].text or ""):
            parent.remove(r)

    for p in root.iter(f"{W}p"):
        txt = "".join(n.text or "" for n in p.iter(f"{W}t"))
        if "{{ENV_NAME}}" in txt and not list(p.iter(f"{W}drawing")):
            runs = p.findall(f"{W}r")
            for idx, r in enumerate(runs[:-1]):
                t1 = r.find(f"{W}t")
                t2 = runs[idx + 1].find(f"{W}t")
                if t1 is not None and t2 is not None and (t1.text or "").endswith(" ") and (t2.text or "") == " ":
                    p.remove(runs[idx + 1])
                    break
        elif "{{NAME}}" in txt:
            runs = p.findall(f"{W}r")
            for idx, r in enumerate(runs[:-1]):
                has_tab = r.find(f"{W}tab") is not None
                t2 = runs[idx + 1].find(f"{W}t")
                if has_tab and t2 is not None and (t2.text or "") == " ":
                    p.remove(runs[idx + 1])
                    break


def _sanitize_document_xml(xml: bytes) -> bytes:
    root = etree.fromstring(xml)
    _unwrap_sdts(root)
    _drop_hidden_shapes(root)
    _strip_merge_fields(root)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _clean_package_part(name: str, blob: bytes) -> bytes:
    if name == "word/settings.xml":
        root = etree.fromstring(blob)
        for node in list(root.iter(f"{W}mailMerge")):
            parent = node.getparent()
            if parent is not None:
                parent.remove(node)
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    if name == "word/_rels/settings.xml.rels":
        root = etree.fromstring(blob)
        for node in list(root):
            if "mailMerge" in (node.get("Type") or ""):
                root.remove(node)
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    if name == "docProps/core.xml":
        root = etree.fromstring(blob)
        for node in root.iter("{http://purl.org/dc/elements/1.1/}description"):
            node.text = ""
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return blob


def _find_soffice() -> Optional[str]:
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found
    if os.name == "nt":
        for candidate in (
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ):
            if os.path.exists(candidate):
                return candidate
    return None


def _apply_tokens(nodes, mapping: Dict[str, str], is_bkk: bool = False) -> None:
    for node in nodes:
        for text_node in node.iter(f"{W}t"):
            current = text_node.text or ""
            if not current:
                continue
            if is_bkk and current == "อ.":
                current = "เขต"
            elif is_bkk and current == " จ.":
                current = " "
            elif current in mapping:
                current = mapping[current]
            else:
                for k, v in mapping.items():
                    if k in current:
                        current = current.replace(k, v)
            text_node.text = current


def _split_office_name(name: str) -> Tuple[str, str]:
    """แยกชื่อหน่วยงาน 2 บรรทัดสำหรับหัวหนังสือ เช่น ('สำนักงานประกันสังคมจังหวัดสกลนคร', 'สาขาสว่างแดนดิน')."""
    text = (name or "").strip()
    if " สาขา" in text:
        parts = text.split(" สาขา", 1)
        return parts[0].strip(), f"สาขา{parts[1].strip()}"
    return text, ""


def _prepare_template_tokens(root) -> None:
    """แปลงจุดข้อความราชการ 8 รายการใน document.xml เป็น token เพื่อรองรับการตั้งค่า."""
    paras = list(root.iter(f"{W}p"))
    if len(paras) < 52:
        return

    # 1. รหัสเอกสาร (DOC_CODE) ใน P[1] และ P[29]
    for p_idx in (1, 29):
        runs = list(paras[p_idx].iter(f"{W}r"))
        if len(runs) > 6 and runs[1].find(f"{W}t") is not None:
            runs[1].find(f"{W}t").text = "{{DOC_CODE}}"
            for r in runs[2:7]:
                t = r.find(f"{W}t")
                if t is not None:
                    t.text = ""

    # 2. ชื่อหน่วยงาน 2 บรรทัดหัวกระดาษ (OFFICE_NAME_L1, OFFICE_NAME_L2)
    # P[1] runs: run 9
    r_p1 = list(paras[1].iter(f"{W}r"))
    if len(r_p1) > 11 and r_p1[9].find(f"{W}t") is not None:
        r_p1[9].find(f"{W}t").text = "{{OFFICE_NAME_L1}}"
        for r in r_p1[10:12]:
            t = r.find(f"{W}t")
            if t is not None:
                t.text = ""

    # P[29] runs: run 11
    r_p29 = list(paras[29].iter(f"{W}r"))
    if len(r_p29) > 13 and r_p29[11].find(f"{W}t") is not None:
        r_p29[11].find(f"{W}t").text = "{{OFFICE_NAME_L1}}"
        for r in r_p29[12:14]:
            t = r.find(f"{W}t")
            if t is not None:
                t.text = ""

    # P[2] and P[30]
    r_p2 = list(paras[2].iter(f"{W}r"))
    if len(r_p2) > 1 and r_p2[1].find(f"{W}t") is not None:
        r_p2[1].find(f"{W}t").text = "{{OFFICE_NAME_L2}}"
    r_p30 = list(paras[30].iter(f"{W}r"))
    if len(r_p30) > 1 and r_p30[1].find(f"{W}t") is not None:
        r_p30[1].find(f"{W}t").text = "{{OFFICE_NAME_L2}}"

    # 3. ที่ตั้งหน่วยงาน (OFFICE_LOCATION) ใน P[3] และ P[31]
    for p_idx in (3, 31):
        r_p = list(paras[p_idx].iter(f"{W}r"))
        if len(r_p) > 1 and r_p[1].find(f"{W}t") is not None:
            r_p[1].find(f"{W}t").text = "{{OFFICE_LOCATION}}"
            for r in r_p[2:]:
                t = r.find(f"{W}t")
                if t is not None:
                    t.text = ""

    # 4. ชื่อหน่วยงานในเนื้อหา (OFFICE_NAME) ใน P[7] และ P[35]
    r_p7 = list(paras[7].iter(f"{W}r"))
    if len(r_p7) > 2 and r_p7[1].find(f"{W}t") is not None:
        r_p7[1].find(f"{W}t").text = "ด้วย{{OFFICE_NAME}}"
        t = r_p7[2].find(f"{W}t")
        if t is not None:
            t.text = ""

    r_p35 = list(paras[35].iter(f"{W}r"))
    if len(r_p35) > 3 and r_p35[2].find(f"{W}t") is not None:
        r_p35[2].find(f"{W}t").text = "{{OFFICE_NAME}}"
        t = r_p35[3].find(f"{W}t")
        if t is not None:
            t.text = ""

    # 5. ที่อยู่หน่วยงาน (OFFICE_ADDRESS) ในหนังสือเตือน P[8]
    r_p8 = list(paras[8].iter(f"{W}r"))
    if len(r_p8) > 26 and r_p8[8].find(f"{W}t") is not None:
        r_p8[8].find(f"{W}t").text = "ณ {{OFFICE_NAME}} {{OFFICE_ADDRESS}} "
        for r in r_p8[9:27]:
            t = r.find(f"{W}t")
            if t is not None:
                t.text = ""

    # 6. สถานที่นัดพบในหนังสือเชิญ P[36]
    r_p36 = list(paras[36].iter(f"{W}r"))
    if len(r_p36) > 9 and r_p36[6].find(f"{W}t") is not None:
        r_p36[6].find(f"{W}t").text = "{{OFFICE_NAME}}"
        for idx in (7, 9):
            t = r_p36[idx].find(f"{W}t")
            if t is not None:
                t.text = ""

    # 7. ผู้ลงนามและตำแหน่ง (SIGNATORY_NAME, SIGNATORY_POSITION)
    if len(paras) > 14:
        r_p13 = list(paras[13].iter(f"{W}r"))
        if len(r_p13) > 1 and r_p13[1].find(f"{W}t") is not None:
            r_p13[1].find(f"{W}t").text = "{{SIGNATORY_NAME}}"
        r_p14 = list(paras[14].iter(f"{W}r"))
        if len(r_p14) > 1 and r_p14[1].find(f"{W}t") is not None:
            r_p14[1].find(f"{W}t").text = "{{SIGNATORY_POSITION}}"

    if len(paras) > 45:
        r_p44 = list(paras[44].iter(f"{W}r"))
        if len(r_p44) > 1 and r_p44[1].find(f"{W}t") is not None:
            r_p44[1].find(f"{W}t").text = "{{SIGNATORY_NAME}}"
        r_p45 = list(paras[45].iter(f"{W}r"))
        if len(r_p45) > 1 and r_p45[1].find(f"{W}t") is not None:
            r_p45[1].find(f"{W}t").text = "{{SIGNATORY_POSITION}}"

    # 8. ส่วนราชการเจ้าของเรื่อง (OWNER_DIVISION)
    for p_idx in (18, 49):
        if len(paras) > p_idx:
            r_p = list(paras[p_idx].iter(f"{W}r"))
            if len(r_p) > 0 and r_p[0].find(f"{W}t") is not None:
                r_p[0].find(f"{W}t").text = "{{OWNER_DIVISION}}"

    # 9. โทรศัพท์ (PHONE)
    for p_idx in (19, 50):
        if len(paras) > p_idx:
            r_p = list(paras[p_idx].iter(f"{W}r"))
            if len(r_p) > 0 and r_p[0].find(f"{W}t") is not None:
                r_p[0].find(f"{W}t").text = "โทรศัพท์  {{PHONE}}"
                for r in r_p[1:]:
                    t = r.find(f"{W}t")
                    if t is not None:
                        t.text = ""

    # 10. โทรสาร (FAX)
    for p_idx in (20, 51):
        if len(paras) > p_idx:
            r_p = list(paras[p_idx].iter(f"{W}r"))
            if len(r_p) > 0 and r_p[0].find(f"{W}t") is not None:
                r_p[0].find(f"{W}t").text = "โทรสาร  {{FAX}}"
                for r in r_p[1:]:
                    t = r.find(f"{W}t")
                    if t is not None:
                        t.text = ""



def build_letter_docx(
    cases: List[Dict[str, Any]],
    issue_date: str,
    appoint_date: str,
    warn_start: int,
    invite_start: Optional[int],
    output_path: str,
    template_path: str = TEMPLATE_PATH,
    config: Optional[Dict[str, Any]] = None,
) -> str:
    if not cases:
        raise ValueError("ไม่มีรายการที่จะพิมพ์")
    if len(cases) > 400:
        raise ValueError("พิมพ์ได้ไม่เกิน 400 รายต่อครั้ง")
    if not os.path.exists(template_path):
        raise ValueError("ไม่พบแม่แบบหนังสือ")

    cfg = load_letter_config()
    if config:
        for k, v in config.items():
            if v is not None and k in DEFAULT_LETTER_CONFIG:
                cfg[k] = str(v).strip()

    office_l1, office_l2 = _split_office_name(cfg.get("office_name", ""))

    numbers = assign_numbers(len(cases), warn_start, invite_start)
    issue_text = format_issue_date(issue_date)
    appt_text = format_appt_date(appoint_date)

    with zipfile.ZipFile(template_path) as zin:
        xml = _sanitize_document_xml(zin.read("word/document.xml"))
        extras = [(info, _clean_package_part(info.filename, zin.read(info.filename))) for info in zin.infolist() if info.filename != "word/document.xml"]

    root = etree.fromstring(xml)
    _prepare_template_tokens(root)
    body = root.find("w:body", NS)
    children = list(body)
    sect = children[-1]
    block = children[:-1]
    for child in children:
        body.remove(child)

    # หากชื่อหน่วยงานมีบรรทัดเดียว (ไม่มี office_l2) ให้นำย่อหน้าบรรทัดที่ 2 ออก
    # เพื่อให้บรรทัดที่ตั้งขยับขึ้นมาเป็นบรรทัดที่ 2 แทน
    if not office_l2:
        l2_warn = None
        l2_invite = None
        for el in block:
            if el.tag == f"{W}p":
                txt = "".join(el.itertext()).strip()
                if txt == "{{OFFICE_NAME_L2}}":
                    if l2_warn is None:
                        l2_warn = el
                    else:
                        l2_invite = el
        if l2_warn is not None and l2_warn in block:
            block.remove(l2_warn)
        if l2_invite is not None and l2_invite in block:
            block.remove(l2_invite)

    for index, case in enumerate(cases):
        if index:
            body.append(_page_break())
        clones = [deepcopy(el) for el in block]
        name = re.sub(r"\s+", " ", str(case.get("name") or "").strip()) or "สถานประกอบการ"
        periods = format_periods_letter(case.get("periods") or [])
        street = _strip_prefix(str(case.get("street") or ""), ("ที่อยู่",)) or "-"
        raw_province = str(case.get("province") or "")
        is_bkk = _is_bangkok(raw_province)
        district = _strip_prefix(str(case.get("district") or ""), ("อำเภอ", "อ.", "เขต")) or "-"
        province = "กรุงเทพมหานคร" if is_bkk else (_strip_prefix(raw_province, ("จังหวัด", "จ.")) or "-")
        warn_no, invite_no = numbers[index]
        _apply_tokens(clones, {
            "{{DOC_CODE}}": cfg.get("doc_code", "สน ๐๐๓1.1/สท"),
            "{{OFFICE_NAME_L1}}": office_l1,
            "{{OFFICE_NAME_L2}}": office_l2,
            "{{OFFICE_LOCATION}}": cfg.get("office_location", ""),
            "{{OFFICE_NAME}}": cfg.get("office_name", ""),
            "{{OFFICE_ADDRESS}}": cfg.get("office_address", ""),
            "{{SIGNATORY_NAME}}": cfg.get("signatory_name", ""),
            "{{SIGNATORY_POSITION}}": cfg.get("signatory_position", ""),
            "{{OWNER_DIVISION}}": cfg.get("owner_division", ""),
            "{{PHONE}}": cfg.get("phone", ""),
            "{{FAX}}": cfg.get("fax", ""),
            "{{WARN_NO}}": format_doc_no(warn_no),
            "{{INVITE_NO}}": format_doc_no(invite_no),
            "{{ISSUE_DATE}}": issue_text,
            "{{APPT}}": appt_text,
            "{{NAME}}": name,
            "{{ENV_NAME}}": envelope_name(name),
            "{{ACC}}": norm_account(case.get("account")),
            "{{BRANCH}}": norm_branch(case.get("branch")),
            "{{PERIODS}}": periods,
            "{{STREET}}": street[:80],
            "{{DISTRICT}}": district[:40],
            "{{PROVINCE}}": province[:40],
            "{{ZIP}}": norm_zip(case.get("zipcode")),
        }, is_bkk=is_bkk)
        for el in clones:
            body.append(el)
    body.append(sect)


    new_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    flat = new_xml.decode("utf-8")
    if "{{" in flat:
        raise ValueError("สร้างหนังสือไม่สำเร็จ: ยังมีช่องว่าง")

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for info, blob in extras:
            zout.writestr(info, blob)
        zout.writestr("word/document.xml", new_xml)
    return output_path


def docx_to_pdf(docx_path: str, pdf_path: str, timeout: int = 180) -> str:
    soffice = _find_soffice()
    if not soffice:
        raise ValueError("เครื่องนี้ไม่มี LibreOffice สำหรับแปลง PDF")
    outdir = os.path.dirname(os.path.abspath(pdf_path))
    os.makedirs(outdir, exist_ok=True)
    profile = tempfile.mkdtemp(prefix="lo-letter-")
    cmd = [
        soffice, "--headless", "--norestore", "--nofirststartwizard",
        f"-env:UserInstallation={Path(profile).as_uri()}",
        "--convert-to", "pdf",
        "--outdir", outdir,
        os.path.abspath(docx_path),
    ]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, check=False)
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    produced = os.path.join(outdir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
    if proc.returncode != 0 or not os.path.exists(produced):
        raise ValueError("แปลง PDF ไม่สำเร็จ")
    if os.path.abspath(produced) != os.path.abspath(pdf_path):
        shutil.move(produced, pdf_path)
    return pdf_path
