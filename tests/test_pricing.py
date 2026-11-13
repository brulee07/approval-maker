import copy
import io
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

import pdfplumber
from openpyxl import load_workbook

from approval_generator.demo import create_demo
from approval_generator.models import Item, Mapping, Profile, Quote, UserError
from approval_generator.pricing import allocate_vat, approval_total, prepare_approval_quote
from approval_generator.render import render


D = Decimal


def quote_with_tax(quantity="3", price="300", vat="101"):
    amount = D(quantity) * D(price)
    return Quote(source="별도부가세.pdf", supplier="시험 업체", title="구매 품의", quotation_date="2026-11-13",
                 subtotal=amount, vat=D(vat), total=amount + D(vat), reviewed=True,
                 items=[Item(description="같은 품목", specification="표준형", unit="개", quantity=D(quantity),
                             unit_price=D(price), amount=amount, note="원본 비고")])


class PricingTests(unittest.TestCase):
    def test_tax_is_included_in_unit_price_and_amount(self):
        output = prepare_approval_quote(quote_with_tax("2", "10000", "2000"))
        self.assertEqual(len(output.items), 1)
        self.assertEqual(output.items[0].unit_price, D(11000))
        self.assertEqual(output.items[0].amount, D(22000))
        self.assertEqual(output.total, D(22000))
        self.assertEqual(output.subtotal, output.total)
        self.assertIsNone(output.vat)

    def test_integer_price_split_preserves_quantity_total_and_product(self):
        output = prepare_approval_quote(quote_with_tax())
        self.assertEqual([(i.quantity, i.unit_price, i.amount) for i in output.items],
                         [(D(1), D(333), D(333)), (D(2), D(334), D(668))])
        self.assertEqual(sum(i.quantity for i in output.items), D(3))
        self.assertEqual(sum(i.quantity * i.unit_price for i in output.items), D(1001))
        for item in output.items:
            self.assertEqual((item.description, item.specification, item.unit, item.note),
                             ("같은 품목", "표준형", "개", "원본 비고"))

    def test_tax_remainders_are_allocated_without_losing_a_won(self):
        self.assertEqual(allocate_vat([D(3333)] * 3, D(1000)), [D(334), D(333), D(333)])
        self.assertEqual(allocate_vat([D(1)] * 7, D(3)), [D(1)] * 3 + [D(0)] * 4)

    def test_included_tax_is_not_added_twice(self):
        source = quote_with_tax("2", "11000", "2000")
        source.amount_basis = "inclusive"
        source.subtotal, source.total = D(20000), D(22000)
        output = prepare_approval_quote(source)
        self.assertEqual(output.items, source.items)
        self.assertEqual(output.total, D(22000))
        self.assertIsNone(output.vat)

    def test_source_is_unchanged_and_conversion_is_repeatable(self):
        source = quote_with_tax()
        snapshot = copy.deepcopy(source)
        first = prepare_approval_quote(source)
        second = prepare_approval_quote(source)
        again = prepare_approval_quote(first)
        self.assertEqual(source, snapshot)
        self.assertEqual(first.items, second.items)
        self.assertEqual(first.items, again.items)
        self.assertEqual(first.total, again.total)

    def test_uses_stated_tax_instead_of_assuming_ten_percent(self):
        output = prepare_approval_quote(quote_with_tax("1", "1000", "80"))
        self.assertEqual(output.items[0].unit_price, D(1080))

    def test_unknown_tax_is_rejected(self):
        source = quote_with_tax()
        source.vat = source.total = None
        with self.assertRaisesRegex(UserError, "합계금액.*비어"):
            prepare_approval_quote(source)

    def test_tax_can_be_derived_from_explicit_gross_total(self):
        source = quote_with_tax()
        source.vat = None
        self.assertEqual(prepare_approval_quote(source).total, D(1001))

    def test_zero_tax_does_not_create_an_extra_charge(self):
        output = prepare_approval_quote(quote_with_tax(vat="0"))
        self.assertEqual(len(output.items), 1)
        self.assertEqual(output.items[0].unit_price, D(300))
        self.assertEqual(output.total, D(900))

    def test_discount_rows_receive_signed_tax(self):
        source = quote_with_tax("1", "70", "6")
        source.items.append(Item(description="할인", quantity=D(1), unit_price=D(-10), amount=D(-10)))
        source.subtotal, source.total = D(60), D(66)
        output = prepare_approval_quote(source)
        self.assertEqual([i.amount for i in output.items], [D(77), D(-11)])
        self.assertEqual(output.total, D(66))

    def test_return_quantity_preserved_when_split(self):
        output = prepare_approval_quote(quote_with_tax("-3", "300", "-101"))
        self.assertEqual(sum(i.quantity for i in output.items), D(-3))
        self.assertEqual(sum(i.amount for i in output.items), D(-1001))
        self.assertTrue(all(i.quantity < 0 for i in output.items))

    def test_fractional_quantity_preserved_when_split(self):
        output = prepare_approval_quote(quote_with_tax("0.3", "300", "10"))
        self.assertEqual(sum(i.quantity for i in output.items), D("0.3"))
        self.assertEqual(sum(i.quantity * i.unit_price for i in output.items), D(100))
        self.assertTrue(all(i.unit_price == i.unit_price.to_integral_value() for i in output.items))

    def test_zero_quantity_and_noninteger_won_tax_need_correction(self):
        for source in [quote_with_tax("0", "100", "0"), quote_with_tax("1", "100", "10.5")]:
            with self.subTest(source=source), self.assertRaises(UserError):
                prepare_approval_quote(source)

    def test_one_won_source_total_mismatch_blocks_conversion(self):
        for difference in (D(-1), D(1)):
            source = quote_with_tax()
            source.total += difference
            with self.subTest(difference=difference), self.assertRaises(UserError):
                prepare_approval_quote(source)

    def test_missing_original_total_is_not_replaced_by_computed_total(self):
        for basis in ("exclusive", "inclusive"):
            source = quote_with_tax()
            source.amount_basis = basis
            source.total = None
            with self.subTest(basis=basis), self.assertRaisesRegex(UserError, "비교"):
                prepare_approval_quote(source)

    def test_inclusive_rounding_is_split_to_match_source_total_exactly(self):
        source = quote_with_tax("3", "333", "0")
        source.amount_basis = "inclusive"
        source.items[0].amount = source.subtotal = source.total = D(1000)
        snapshot = copy.deepcopy(source)
        output = prepare_approval_quote(source)
        self.assertEqual(sum(i.quantity for i in output.items), D(3))
        self.assertEqual([(i.quantity, i.unit_price) for i in output.items], [(D(2), D(333)), (D(1), D(334))])
        self.assertEqual(approval_total(output), source.total)
        self.assertEqual(source, snapshot)

    def test_output_row_calculation_rejects_even_one_won_difference(self):
        output = prepare_approval_quote(quote_with_tax())
        output.items[0].unit_price += D(1)
        with self.assertRaisesRegex(UserError, "출력 행.*다릅니다"):
            approval_total(output)


class PricingOutputTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory(prefix="approval-pricing-test-")
        cls.files = create_demo(Path(cls.temp.name))

    @classmethod
    def tearDownClass(cls):
        cls.temp.cleanup()

    def test_xlsx_has_integer_gross_prices_and_no_separate_tax_value(self):
        profile = Profile.load(str(self.files["excel_profile"]))
        data = render(profile, quote_with_tax())
        book = load_workbook(io.BytesIO(data), data_only=True)
        try:
            sheet = book["양식"]
            self.assertEqual((sheet["D7"].value, sheet["E7"].value), (1, 333))
            self.assertEqual((sheet["D8"].value, sheet["E8"].value), (2, 334))
            self.assertEqual(sheet["F7"].value + sheet["F8"].value, 1001)
            self.assertEqual(sheet["F13"].value, 1001)
            self.assertEqual(sheet["F15"].value, 1001)
            self.assertIsNone(sheet["F14"].value)
        finally:
            book.close()

    def test_pdf_has_split_rows_and_no_separate_tax_value(self):
        profile = Profile.load(str(self.files["pdf_profile"]))
        data = render(profile, quote_with_tax())
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            text = pdf.pages[0].extract_text()
            self.assertEqual(text.count("같은 품목"), 2)
            self.assertIn("333", text)
            self.assertIn("334", text)
            self.assertIn("1,001", text)
            self.assertNotIn("101", text)

    def test_template_capacity_is_checked_after_price_splitting(self):
        profile = Profile.load(str(self.files["excel_profile"]))
        for mapping in profile.mappings:
            if mapping.field.startswith("items."):
                mapping.count = 1
        with self.assertRaisesRegex(UserError, "품목은 2개"):
            render(profile, quote_with_tax())

    def test_price_mapping_requires_quantity_for_vat_conversion(self):
        source = quote_with_tax()
        source.items[0].quantity = None
        profile = Profile.load(str(self.files["excel_profile"]))
        with self.assertRaisesRegex(UserError, "수량과 금액"):
            render(profile, source)

    def test_saved_xlsx_quantity_times_price_matches_quote_total_after_rounding(self):
        source = quote_with_tax("3", "333", "0")
        source.amount_basis = "inclusive"
        source.items[0].amount = source.subtotal = source.total = D(1000)
        data = render(Profile.load(str(self.files["excel_profile"])), source)
        book = load_workbook(io.BytesIO(data), data_only=True)
        try:
            sheet = book["양식"]
            calculated = sum(D(str(sheet[f"D{row}"].value)) * D(str(sheet[f"E{row}"].value)) for row in (7, 8))
            self.assertEqual(calculated, source.total)
            self.assertEqual(D(str(sheet["F15"].value)), calculated)
        finally:
            book.close()

    def test_inclusive_missing_quantity_cannot_export_quantity_and_price(self):
        source = quote_with_tax("1", "1000", "0")
        source.amount_basis = "inclusive"
        source.items[0].quantity = None
        with self.assertRaisesRegex(UserError, "수량이 필요"):
            render(Profile.load(str(self.files["excel_profile"])), source)


if __name__ == "__main__":
    unittest.main()
