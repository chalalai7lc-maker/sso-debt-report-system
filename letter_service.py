"""สร้างหนังสือเตือน + เชิญจากฟอร์มราชการ โดยไม่ทิ้งข้อมูลตัวอย่างของต้นฉบับ."""

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
    if not (1 <= month <= 12 and 1 <= day <= 31 and 1900 <= year_ce <= 2400):
        raise ValueError("วันที่ไม่ถูกต้อง")
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
    return _digits(value)


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


def _sanitize_document_xml(xml: bytes) -> bytes:
    root = etree.fromstring(xml)
    _unwrap_sdts(root)
    _drop_hidden_shapes(root)
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


def _apply_tokens(nodes, mapping: Dict[str, str]) -> None:
    for node in nodes:
        for text_node in node.iter(f"{W}t"):
            current = text_node.text or ""
            if current in mapping:
                text_node.text = mapping[current]


def build_letter_docx(
    cases: List[Dict[str, Any]],
    issue_date: str,
    appoint_date: str,
    warn_start: int,
    invite_start: Optional[int],
    output_path: str,
    template_path: str = TEMPLATE_PATH,
) -> str:
    if not cases:
        raise ValueError("ไม่มีรายการที่จะพิมพ์")
    if len(cases) > 400:
        raise ValueError("พิมพ์ได้ไม่เกิน 400 รายต่อครั้ง")
    if not os.path.exists(template_path):
        raise ValueError("ไม่พบแม่แบบหนังสือ")

    numbers = assign_numbers(len(cases), warn_start, invite_start)
    issue_text = format_issue_date(issue_date)
    appt_text = format_appt_date(appoint_date)

    with zipfile.ZipFile(template_path) as zin:
        xml = _sanitize_document_xml(zin.read("word/document.xml"))
        extras = [(info, _clean_package_part(info.filename, zin.read(info.filename))) for info in zin.infolist() if info.filename != "word/document.xml"]

    root = etree.fromstring(xml)
    body = root.find("w:body", NS)
    children = list(body)
    sect = children[-1]
    block = children[:-1]
    for child in children:
        body.remove(child)

    for index, case in enumerate(cases):
        if index:
            body.append(_page_break())
        clones = [deepcopy(el) for el in block]
        name = re.sub(r"\s+", " ", str(case.get("name") or "").strip()) or "สถานประกอบการ"
        periods = format_periods_letter(case.get("periods") or [])
        street = _strip_prefix(str(case.get("street") or ""), ("ที่อยู่",)) or "-"
        district = _strip_prefix(str(case.get("district") or ""), ("อำเภอ", "อ.", "เขต")) or "-"
        province = _strip_prefix(str(case.get("province") or ""), ("จังหวัด", "จ.")) or "-"
        warn_no, invite_no = numbers[index]
        _apply_tokens(clones, {
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
        })
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
