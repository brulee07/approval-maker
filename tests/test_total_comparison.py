import json
import queue
import tempfile
import unittest
from decimal import Decimal as D
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

from approval_generator.demo import create_demo
from approval_generator.gui import App
from approval_generator.models import Item, Profile, Quote


class TotalComparisonTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = App()
        cls.app.withdraw()
        cls.initial_output_folder = cls.app.output_dir.get()

    @classmethod
    def tearDownClass(cls):
        cls.app.update_idletasks()
        for callback in cls.app.tk.splitlist(cls.app.tk.call("after", "info")):
            cls.app.after_cancel(callback)
        cls.app.destroy()

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        files = create_demo(self.folder / "demo")
        self.app.profile = Profile.load(str(files["excel_profile"]))
        self.app.default_unit.set("개")
        self.app.generation_records = {}
        self.app.output_dir.set(str(self.folder / "output"))
        self.app.quotes = [Quote(source="부가세별도.pdf", title="시험 품의", supplier="시험 업체",
                                 quotation_date="2026-11-13", subtotal=D(900), vat=D(101), total=D(1001), reviewed=True,
                                 items=[Item(description="부품", quantity=D(3), unit_price=D(300), amount=D(900))])]
        self.app._load_quote(0)
        self.app.tabs.select(self.app.review_tab)
        self.app.update()

    def generate(self):
        def run(_label, task, done):
            results = task()
            while True:
                try:
                    self.app.events.get_nowait()
                except queue.Empty:
                    break
            done(results)
        with patch.object(self.app, "_run", side_effect=run):
            self.app.generate()

    def test_totals_appear_on_export_tab_without_opening_preview(self):
        with patch.object(self.app, "preview_pricing") as preview:
            self.app.tabs.select(self.app.export_tab)
            self.app.update()
            preview.assert_not_called()
        self.assertEqual([self.app.results.set("0", key) for key in ("source_total", "output_total", "difference", "comparison", "status")],
                         ["1,001", "1,001", "0", "일치(예정)", "생성 전"])
        self.assertEqual(self.app.result_paths, {})

    def test_no_default_output_folder_and_generation_requires_selection(self):
        self.assertEqual(self.initial_output_folder, "")
        self.app.output_dir.set("")
        with patch.object(self.app, "_run") as run:
            with self.assertRaisesRegex(ValueError, "저장 폴더"):
                self.app.generate()
            run.assert_not_called()
        with patch("approval_generator.gui.open_file") as opened:
            with self.assertRaisesRegex(ValueError, "폴더 선택"):
                self.app.open_output_folder()
            opened.assert_not_called()

    def test_user_selected_output_folder_is_used(self):
        chosen = self.folder / "직접 선택한 폴더"
        chosen.mkdir()
        with patch("approval_generator.gui.filedialog.askdirectory", return_value=str(chosen)):
            self.app.choose_output()
        self.generate()
        self.assertEqual(Path(self.app.result_paths["0"]).parent, chosen)

    def test_successful_file_totals_report_and_table_agree(self):
        self.generate()
        row = self.app.results
        self.assertEqual([row.set("0", key) for key in ("source_total", "output_total", "difference", "comparison", "status")],
                         ["1,001", "1,001", "0", "일치", "성공"])
        path = self.app.result_paths["0"]
        book = load_workbook(path, data_only=True)
        try:
            sheet = book["양식"]
            self.assertEqual(sum(sheet[f"D{r}"].value * sheet[f"E{r}"].value for r in (7, 8)), 1001)
            self.assertEqual(sheet["F15"].value, 1001)
        finally:
            book.close()
        report = json.loads(next((self.folder / "output").glob("생성결과*.json")).read_text(encoding="utf-8"))
        self.assertEqual(report["results"][0]["total_difference"], "0")
        self.app.tabs.select(self.app.export_tab)
        self.app.update()
        self.assertEqual(self.app.results.set("0", "status"), "성공")
        self.app.results.selection_set("0")
        with patch("approval_generator.gui.open_file") as opened:
            self.app.open_result()
            opened.assert_called_once_with(path)

    def test_one_won_mismatch_is_visible_and_produces_no_approval(self):
        self.app.quotes[0].total = D(1002)
        self.app._load_quote(0)
        self.app.tabs.select(self.app.export_tab)
        self.app.update()
        self.assertEqual(self.app.results.set("0", "comparison"), "불일치")
        self.assertEqual(self.app.results.set("0", "difference"), "-1")
        self.generate()
        self.assertEqual(self.app.results.set("0", "status"), "실패")
        self.assertEqual(self.app.result_paths, {})
        self.assertEqual(list((self.folder / "output").glob("품의서*.xlsx")), [])

    def test_editing_input_removes_previous_success_from_current_comparison(self):
        self.generate()
        self.app.fields["title"].set("수정된 제목")
        self.app._commit()
        self.app._refresh_export_comparison()
        self.assertEqual(self.app.results.set("0", "status"), "검토 필요")
        self.assertEqual(self.app.results.set("0", "comparison"), "일치(예정)")
        self.assertEqual(self.app.result_paths, {})

    def test_error_details_open_from_message_column(self):
        self.app.quotes[0].total = None
        self.app._load_quote(0)
        self.app._refresh_export_comparison()
        self.app.results.selection_set("0")
        with patch("approval_generator.gui.messagebox.showinfo") as info:
            self.app.open_result()
            self.assertIn("합계금액이 비어", info.call_args.args[1])
        self.assertEqual(self.app.results.set("0", "source_total"), "미입력")
        self.assertEqual(self.app.results.set("0", "comparison"), "확인 필요")


if __name__ == "__main__":
    unittest.main()
