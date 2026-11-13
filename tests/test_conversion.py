from __future__ import annotations

import copy
import io
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from zipfile import ZipFile

import pdfplumber
from openpyxl import load_workbook
from pypdf import PdfReader, PdfWriter

from approval_generator.demo import create_demo, sample_xlsx
from approval_generator.extract import extract_quote, items_from_rows, read_metadata
from approval_generator.models import Item, Mapping, Profile, Quote, UserError, number
from approval_generator.render import render, save_result, xml, NS


class ConversionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory(prefix="approval-test-")
        cls.root = Path(cls.temp.name)
        cls.files = create_demo(cls.root / "demo")

    @classmethod
    def tearDownClass(cls):
        cls.temp.cleanup()

    def quote(self, kind="source_xlsx"):
        quote = extract_quote(str(self.files[kind]))
        quote.reviewed = True
        return quote

    def profile(self, kind="excel_profile"):
        return Profile.load(str(self.files[kind]))

    def test_amount_parsing_keeps_blank_distinct_from_zero(self):
        self.assertIsNone(number(""))
        self.assertIsNone(number("-"))
        self.assertEqual(number("0"), 0)
        self.assertEqual(number("₩1,250.50"), Decimal("1250.50"))
        self.assertEqual(number("(2,000)"), -2000)
        self.assertEqual(number("\\19,000"), 19000)
        self.assertEqual(number("(\\ 38,000)"), 38000)
        for invalid in ["NaN", "Infinity", "1개", True, "1e20"]:
            with self.subTest(value=invalid), self.assertRaises(UserError):
                number(invalid)

    def test_xlsx_extraction_and_metadata(self):
        quote = self.quote()
        self.assertEqual(len(quote.items), 2)
        self.assertEqual(quote.supplier, "샘플상사")
        self.assertEqual(quote.quotation_date, "2026-11-13")
        self.assertEqual(quote.items[0].amount, 100000)
        self.assertEqual(quote.subtotal, 250000)
        self.assertEqual(quote.vat, 25000)
        self.assertEqual(quote.total, 275000)
        self.assertEqual(quote.issues(), [])

    def test_pdf_extraction_and_metadata(self):
        quote = self.quote("source_pdf")
        self.assertEqual(len(quote.items), 2)
        self.assertEqual(quote.supplier, "샘플상사")
        self.assertEqual(quote.subtotal, 250000)
        self.assertEqual(quote.vat, 25000)
        self.assertEqual(quote.total, 275000)
        self.assertEqual(quote.issues(), [])

    def test_duplicate_line_items_are_preserved(self):
        rows = [["품명", "수량", "단가", "금액"], ["부품", "1", "1000", "1000"], ["부품", "1", "1000", "1000"]]
        self.assertEqual(len(items_from_rows(rows, [])), 2)

    def test_unknown_vat_not_assumed(self):
        path = self.root / "no-vat.xlsx"
        sample_xlsx(path, {"A1": "품명", "B1": "수량", "C1": "단가", "D1": "금액", "A2": "품목", "B2": 1, "C2": 1000, "D2": 1000})
        quote = extract_quote(str(path))
        self.assertIsNone(quote.vat)
        self.assertIsNone(quote.total)

    def test_vat_included_line_totals_are_not_taxed_twice(self):
        quote = self.quote()
        quote.amount_basis = "inclusive"
        quote.subtotal = Decimal("227273")
        quote.vat = Decimal("22727")
        quote.total = Decimal("250000")
        self.assertEqual(quote.issues(), [])
        quote.amount_basis = "exclusive"
        self.assertTrue(quote.issues())

    def test_korean_spaced_headers_and_currency_summaries(self):
        rows = [["품 목", "수 량", "단 가", "공급대가\n(VAT 포함)"],
                ["부품", "2", "\\19,000", "(\\ 38,000)"],
                ["합계 (VAT 별도)", "(\\ 34,545)"], ["부가세", "(\\ 3,455)"],
                ["최종 견적가 (VAT 포함)", "(\\ 38,000)"]]
        quote = Quote()
        quote.items = items_from_rows(rows, [])
        read_metadata(quote, rows, "2026년 11월 12일\n합 계 금 액  \\38,000")
        self.assertEqual(len(quote.items), 1)
        self.assertEqual(quote.items[0].amount, 38000)
        self.assertEqual((quote.subtotal, quote.vat, quote.total), (34545, 3455, 38000))
        self.assertEqual(quote.quotation_date, "2026-11-12")

    def test_unknown_pdf_glyph_requires_review(self):
        quote = self.quote()
        quote.items[0].description = "2(cid:1070)1 Ratio"
        self.assertTrue(any("문자" in issue for issue in quote.issues()))

    def test_user_default_unit_only_fills_empty_units(self):
        quote = self.quote()
        quote.items[0].unit = ""
        quote.items[1].unit = "세트"
        profile = self.profile()
        profile.default_unit = "개"
        book = load_workbook(io.BytesIO(render(profile, quote)), data_only=True)
        self.assertEqual(book["양식"]["C7"].value, "개")
        self.assertEqual(book["양식"]["C8"].value, "세트")
        self.assertEqual(quote.items[0].unit, "")
        book.close()

    def test_all_input_output_combinations(self):
        for source in ["source_pdf", "source_xlsx"]:
            for target in ["pdf_profile", "excel_profile"]:
                with self.subTest(source=source, target=target):
                    data = render(self.profile(target), self.quote(source))
                    if target == "excel_profile":
                        workbook = load_workbook(io.BytesIO(data), data_only=True)
                        try:
                            self.assertEqual(workbook["양식"]["B3"].value, "샘플상사")
                            self.assertEqual(workbook["양식"]["F15"].value, 275000)
                        finally:
                            workbook.close()
                    else:
                        with pdfplumber.open(io.BytesIO(data)) as pdf:
                            text = pdf.pages[0].extract_text()
                            self.assertIn("샘플상사", text)
                            self.assertIn("275,000", text)
                            self.assertIn("무선 키보드", text)

    def test_xlsx_preserves_nonmodified_archive_parts(self):
        profile = self.profile()
        with ZipFile(profile.template) as original, ZipFile(io.BytesIO(render(profile, self.quote()))) as result:
            self.assertEqual(original.namelist(), result.namelist())
            for name in original.namelist():
                if name not in {"xl/worksheets/sheet1.xml", "xl/workbook.xml"}:
                    self.assertEqual(original.read(name), result.read(name), name)
            before = xml(original.read("xl/worksheets/sheet1.xml"))
            after = xml(result.read("xl/worksheets/sheet1.xml"))
            for node in ["cols", "mergeCells", "pageMargins", "pageSetup"]:
                from lxml import etree
                self.assertEqual(etree.tostring(before.find("s:" + node, NS)), etree.tostring(after.find("s:" + node, NS)))

    def test_formula_cells_are_protected(self):
        path = self.root / "formula.xlsx"
        sample_xlsx(path, {}, formulas={"F13": "SUM(F7:F11)"})
        profile = Profile.create(str(path))
        profile.mappings = [Mapping(field="subtotal", sheet="양식", cell="F13")]
        with self.assertRaisesRegex(UserError, "수식"):
            render(profile, self.quote())

    def test_text_is_never_executed_as_formula(self):
        quote = self.quote()
        quote.supplier = '=HYPERLINK("example")'
        data = render(self.profile(), quote)
        workbook = load_workbook(io.BytesIO(data), data_only=False)
        self.assertEqual(workbook["양식"]["B3"].data_type, "s")
        workbook.close()

    def test_merged_cell_anchor_required(self):
        profile = self.profile()
        profile.mappings[0].cell = "C3"
        with self.assertRaisesRegex(UserError, "병합"):
            render(profile, self.quote())

    def test_item_overflow_fails_without_truncation(self):
        profile = self.profile()
        for mapping in profile.mappings:
            if mapping.field.startswith("items."):
                mapping.count = 1
        with self.assertRaisesRegex(UserError, "품목은 2개"):
            render(profile, self.quote())

    def test_unreviewed_and_amount_mismatch_block_generation(self):
        quote = self.quote()
        quote.reviewed = False
        with self.assertRaisesRegex(UserError, "검토"):
            render(self.profile(), quote)
        quote.reviewed = True
        quote.total = Decimal("1")
        with self.assertRaisesRegex(UserError, "합계금액"):
            render(self.profile(), quote)

    def test_missing_mapped_value_is_not_silently_empty(self):
        quote = self.quote()
        quote.supplier = ""
        with self.assertRaisesRegex(UserError, "공급업체"):
            render(self.profile(), quote)

    def test_duplicate_excel_targets_block_generation(self):
        profile = self.profile()
        profile.mappings.append(copy.deepcopy(profile.mappings[0]))
        with self.assertRaisesRegex(UserError, "같은 셀"):
            render(profile, self.quote())

    def test_pdf_text_overflow_and_area_overlap(self):
        profile = self.profile("pdf_profile")
        profile.mappings[0].height = 10
        with self.assertRaisesRegex(UserError, "다 들어가지"):
            render(profile, self.quote())
        profile = self.profile("pdf_profile")
        profile.mappings.append(copy.deepcopy(profile.mappings[0]))
        with self.assertRaisesRegex(UserError, "겹칩니다"):
            render(profile, self.quote())

    def test_pdf_outside_page_is_rejected(self):
        profile = self.profile("pdf_profile")
        profile.mappings[0].x = 1000
        with self.assertRaisesRegex(UserError, "범위를 벗어"):
            render(profile, self.quote())

    def test_rotated_cropped_pdf_placement(self):
        path = self.root / "rotated.pdf"
        writer = PdfWriter()
        page = writer.add_blank_page(width=600, height=800)
        page.cropbox.lower_left = (30, 50)
        page.cropbox.upper_right = (570, 750)
        page.rotate(90)
        writer.write(path)
        profile = Profile.create(str(path))
        profile.mappings = [Mapping(field="supplier", x=40, y=60, width=200, height=25)]
        data = render(profile, self.quote())
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            page = pdf.pages[0]
            # 화면의 cropbox 기준 왼쪽 40pt, 위쪽 60pt에 위치해야 한다.
            words = page.extract_words()
            self.assertEqual(words[0]["text"], "샘플상사")
            self.assertAlmostEqual(words[0]["x0"] - page.cropbox[0], 42, delta=2)
            self.assertAlmostEqual(words[0]["top"] - page.cropbox[1], 62, delta=4)

    def test_profile_is_bound_to_template_content(self):
        profile = self.profile()
        profile.template_hash = "changed"
        with self.assertRaisesRegex(UserError, "다릅니다"):
            render(profile, self.quote())

    def test_output_never_overwrites(self):
        folder = self.root / "results"
        first = save_result(self.profile(), self.quote(), str(folder))
        original = first.read_bytes()
        second = save_result(self.profile(), self.quote(), str(folder))
        self.assertNotEqual(first, second)
        self.assertEqual(first.read_bytes(), original)

    def test_output_folder_must_be_explicitly_provided(self):
        for folder in ("", "   ", None):
            with self.subTest(folder=folder), self.assertRaisesRegex(UserError, "저장 폴더.*직접"):
                save_result(self.profile(), self.quote(), folder)

    def test_blank_pdf_requires_manual_entry(self):
        path = self.root / "blank.pdf"
        writer = PdfWriter()
        writer.add_blank_page(width=595, height=842)
        writer.write(path)
        quote = extract_quote(str(path))
        self.assertEqual(quote.items, [])
        self.assertFalse(quote.reviewed)
        self.assertTrue(any("OCR" in warning for warning in quote.warnings))


if __name__ == "__main__":
    unittest.main()
