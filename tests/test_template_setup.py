import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from approval_generator.gui import App
from approval_generator.models import Item, Mapping, Profile, Quote, UserError
from approval_generator.templates import profile_for_template


class SavedTemplateTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.template = self.folder / "양식.xlsx"
        book = Workbook()
        book.active.append(["내용", "규격", "단위", "수량", "예상단가"])
        book.save(self.template)
        book.close()
        self.saved = Profile.create(str(self.template))
        self.saved.default_unit = "개"
        self.saved.mappings = [Mapping(field="items.description", sheet="Sheet", cell="A2", count=1000),
                               Mapping(field="items.unit_price", sheet="Sheet", cell="E2", count=1000)]
        self.saved.save(str(self.folder / "설정.json"))

    def test_saved_locations_and_default_unit_restore_for_identical_copy(self):
        copied = self.folder / "이름변경.xlsx"
        copied.write_bytes(self.template.read_bytes())
        (self.folder / "손상.json").write_text("{bad", encoding="utf-8")
        selected, message = profile_for_template(copied, self.folder)
        self.assertEqual(selected.template, str(copied.resolve()))
        self.assertEqual(selected.mappings, self.saved.mappings)
        self.assertEqual(selected.default_unit, "개")
        self.assertIn("자동", message)
        selected.check_template()

    def test_changed_template_does_not_reuse_old_locations(self):
        book = Workbook()
        book.active.append(["변경된 양식"])
        book.save(self.template)
        book.close()
        selected, _ = profile_for_template(self.template, self.folder)
        self.assertEqual(selected.mappings, [])

    def test_conflicting_saved_settings_require_explicit_selection(self):
        self.saved.default_unit = "대"
        self.saved.save(str(self.folder / "다른설정.json"))
        selected, message = profile_for_template(self.template, self.folder)
        self.assertEqual(selected.mappings, [])
        self.assertIn("여러 개", message)

    def test_duplicate_identical_settings_are_not_ambiguous(self):
        self.saved.save(str(self.folder / "복사본.json"))
        selected, _ = profile_for_template(self.template, self.folder)
        self.assertEqual(selected.mappings, self.saved.mappings)


class TemplateWorkflowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = App()
        cls.app.withdraw()

    @classmethod
    def tearDownClass(cls):
        cls.app.update_idletasks()
        for callback in cls.app.tk.splitlist(cls.app.tk.call("after", "info")):
            cls.app.after_cancel(callback)
        cls.app.destroy()

    def setUp(self):
        self.app.profile = None
        self.app.quotes = [Quote(source="검토한견적서.pdf", reviewed=True, subtotal=Decimal(1000),
                                 vat=Decimal(100), total=Decimal(1100),
                                 items=[Item(description="부품", quantity=Decimal(1),
                                             unit_price=Decimal(1000), amount=Decimal(1000))])]
        self.app._load_quote(0)
        self.app.tabs.select(self.app.export_tab)

    def test_missing_mapping_redirects_before_batch_and_preserves_review(self):
        self.app.profile = Profile("template.xlsx", "hash")
        with patch.object(self.app, "_run") as run:
            with self.assertRaisesRegex(UserError, "열 제목으로 자동 연결"):
                self.app.generate()
            run.assert_not_called()
        self.assertEqual(self.app.tabs.select(), str(self.app.mapping_tab))
        self.assertTrue(self.app.quotes[0].reviewed)

    def test_template_picker_restores_mapping_and_shows_ready_state(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "양식.xlsx"
            book = Workbook()
            book.active.append(["내용", "단위", "수량", "단가"])
            book.save(path)
            book.close()
            profile = Profile.create(str(path))
            profile.default_unit = "개"
            profile.mappings = [Mapping(field="items.description", sheet="Sheet", cell="A2")]
            profile.save(str(Path(folder) / "연결.json"))
            with patch("approval_generator.gui.filedialog.askopenfilename", return_value=str(path)), \
                 patch("approval_generator.gui.profile_for_template", side_effect=lambda p, _: profile_for_template(p, folder)):
                self.app.choose_template()
            self.assertEqual(self.app.profile.mappings, profile.mappings)
            self.assertEqual(self.app.default_unit.get(), "개")
            self.assertIn("1개 연결됨", self.app.export_template.get())
            self.assertIn("자동", self.app.workflow_notice.get())
            self.assertTrue(self.app.quotes[0].reviewed)


if __name__ == "__main__":
    unittest.main()
