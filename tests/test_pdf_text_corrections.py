import os
import unittest
from decimal import Decimal
from pathlib import Path

from approval_generator.extract import extract_quote
from approval_generator.models import Item, Quote
from approval_generator.pdf_text_corrections import apply_verified_corrections


SOURCE_HASH = "9e82960ed5afffe871468a0ab4eb17a16353a88d33613d7bd6a90025d8dca16b"
BROKEN_NAME = "2(cid:1070)1 Ratio Bevel Gear Set (6mm D-Bore Pinion Gear)"
FIXED_NAME = "2:1 Ratio Bevel Gear Set (6mm D-Bore Pinion Gear)"
REAL_PDF = Path(os.environ.get("APPROVAL_TEST_PDF", "tests/private-fixtures/verified-quote.pdf"))


class PdfTextCorrectionTests(unittest.TestCase):
    def test_verified_source_and_description_are_corrected_once(self):
        quote = Quote(items=[Item(description=BROKEN_NAME)])
        apply_verified_corrections(quote, SOURCE_HASH)
        apply_verified_corrections(quote, SOURCE_HASH)
        self.assertEqual(quote.items[0].description, FIXED_NAME)
        self.assertEqual(len(quote.warnings), 1)
        self.assertFalse(quote.reviewed)

    def test_same_cid_in_another_pdf_is_not_guessed(self):
        quote = Quote(items=[Item(description=BROKEN_NAME)])
        apply_verified_corrections(quote, "different-pdf")
        self.assertEqual(quote.items[0].description, BROKEN_NAME)
        self.assertEqual(quote.warnings, [])

    def test_unverified_product_text_is_not_replaced(self):
        quote = Quote(items=[Item(description="다른 부품 (cid:1070)")])
        apply_verified_corrections(quote, SOURCE_HASH)
        self.assertEqual(quote.items[0].description, "다른 부품 (cid:1070)")

    def test_unknown_glyph_reports_item_number_and_field(self):
        quote = Quote(items=[Item(description="정상 품목", amount=Decimal(100)),
                             Item(description="다른 품목", specification="(cid:99)", amount=Decimal(100))])
        self.assertEqual(len(quote.issues()), 1)
        self.assertIn("2번 품목의 규격", quote.issues()[0])

    @unittest.skipUnless(REAL_PDF.is_file(), "사용자가 제공한 PDF가 있을 때 수행하는 회귀 검사")
    def test_real_pdf_can_be_reviewed_without_a_manual_launch_patch(self):
        quote = extract_quote(str(REAL_PDF))
        self.assertEqual(len(quote.items), 20)
        self.assertEqual(quote.items[9].description, FIXED_NAME)
        self.assertEqual(quote.total, Decimal(1714000))
        self.assertEqual(quote.issues(), [])
        self.assertIn("(cid:1070)", quote.raw_text)  # 원문은 그대로 보존
        self.assertFalse(quote.reviewed)


if __name__ == "__main__":
    unittest.main()
