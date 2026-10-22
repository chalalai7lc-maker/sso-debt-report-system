import os
import sys
import tempfile
import unittest
import zipfile

from lxml import etree

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from letter_service import (
    assign_numbers,
    build_letter_docx,
    format_doc_no,
    format_issue_date,
    format_periods_letter,
)

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _docx_text(path):
    xml = zipfile.ZipFile(path).read("word/document.xml")
    root = etree.fromstring(xml)
    return "".join(node.text or "" for node in root.iter(f"{W}t"))


class TestLetterNumbers(unittest.TestCase):
    def test_interleaved_when_only_warning_start(self):
        self.assertEqual(assign_numbers(2, 1, None), [(1, 2), (3, 4)])

    def test_parallel_when_both_starts(self):
        self.assertEqual(assign_numbers(2, 1, 100), [(1, 100), (2, 101)])

    def test_doc_no_no_padding(self):
        self.assertEqual(format_doc_no(1), "1")
        self.assertEqual(format_doc_no(1747), "1747")

    def test_period_letter_format(self):
        self.assertEqual(format_periods_letter(["07/2569", "08/2569"]), "ก.ค.,ส.ค. 2569")
        self.assertEqual(format_periods_letter(["12/2568", "01/2569"]), "ธ.ค. 2568,ม.ค. 2569")

    def test_issue_date_be(self):
        self.assertEqual(format_issue_date("2026-09-18"), "18 กันยายน 2569")


class TestLetterDocx(unittest.TestCase):
    def test_two_cases_replace_sample_and_numbers(self):
        cases = [
            {
                "name": "บริษัททดสอบ จำกัด",
                "account": "1000000001",
                "branch": "1",
                "periods": ["07/2569", "08/2569"],
                "street": "99 ถนนทดสอบ",
                "district": "เมือง",
                "province": "สกลนคร",
                "zipcode": "47000",
            },
            {
                "name": "ห้างทดสอบ",
                "account": "4790000002",
                "branch": "000010",
                "periods": ["01/2569"],
                "street": "1 หมู่ 2",
                "district": "สว่างแดนดิน",
                "province": "สกลนคร",
                "zipcode": "47110",
            },
        ]
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "letters.docx")
            build_letter_docx(cases, "2026-09-18", "2026-10-18", 1, None, path)
            text = _docx_text(path)
            self.assertIn("สท 1", text)
            self.assertIn("สท 2", text)
            self.assertIn("สท 3", text)
            self.assertIn("สท 4", text)
            self.assertIn("บริษัททดสอบ จำกัด", text)
            self.assertIn("ห้างทดสอบ", text)
            self.assertIn("1000000001", text)
            self.assertIn("ก.ค.,ส.ค. 2569", text)
            self.assertIn("18 กันยายน 2569", text)
            self.assertIn("18/10/2569", text)
            self.assertNotIn("{{", text)
            self.assertNotIn("4790001695", text)
            self.assertNotIn("โรงเรียนวัดพระธาตุ", text)
