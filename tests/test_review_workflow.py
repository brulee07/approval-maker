import unittest
from decimal import Decimal
from unittest.mock import patch

from approval_generator.gui import App
from approval_generator.models import Item, Mapping, Profile, Quote


def sample_quote(name):
    return Quote(source=name, subtotal=Decimal(1000), vat=Decimal(100), total=Decimal(1100),
                 items=[Item(description="부품", quantity=Decimal(1), unit_price=Decimal(1000), amount=Decimal(1000))])


class ReviewWorkflowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = App()
        cls.app.withdraw()

    @classmethod
    def tearDownClass(cls):
        for callback in cls.app.tk.splitlist(cls.app.tk.call("after", "info")):
            cls.app.after_cancel(callback)
        cls.app.destroy()

    def setUp(self):
        self.app.profile = None
        self.app.quotes = [sample_quote("견적서1.pdf")]
        self.app._load_quote(0)
        self.app.tabs.select(self.app.review_tab)
        self.app.update()

    def test_review_button_goes_to_template_step_without_mapping(self):
        self.app.review_button.invoke()
        self.app.update()
        self.assertTrue(self.app.quotes[0].reviewed)
        self.assertEqual(self.app.tabs.select(), str(self.app.mapping_tab))
        self.assertIn("검토 완료", self.app.quote_list.get(0))
        self.assertIn("양식 선택", self.app.workflow_notice.get())

    def test_review_button_goes_to_export_when_mapping_is_ready(self):
        self.app.profile = Profile("template.xlsx", "hash", [Mapping()])
        self.app.review_button.invoke()
        self.app.update()
        self.assertEqual(self.app.tabs.select(), str(self.app.export_tab))
        self.assertIn("전체 품의서 생성", self.app.workflow_notice.get())
        self.assertEqual(self.app.results.set("0", "status"), "생성 전")
        self.assertEqual(self.app.result_paths, {})  # 검토만으로 파일 생성하지 않음

    def test_missing_mappings_still_go_to_template_step(self):
        self.app.profile = Profile("template.xlsx", "hash")
        self.app.review_button.invoke()
        self.assertEqual(self.app.tabs.select(), str(self.app.mapping_tab))

    def test_batch_review_moves_to_next_unreviewed_quote(self):
        self.app.quotes.append(sample_quote("견적서2.pdf"))
        self.app.review_button.invoke()
        self.app.update()
        self.assertTrue(self.app.quotes[0].reviewed)
        self.assertFalse(self.app.quotes[1].reviewed)
        self.assertEqual(self.app.current, 1)
        self.assertEqual(self.app.tabs.select(), str(self.app.review_tab))
        self.assertIn("1/2개", self.app.workflow_notice.get())
        self.app.review_button.invoke()
        self.app.update()
        self.assertTrue(all(q.reviewed for q in self.app.quotes))
        self.assertEqual(self.app.tabs.select(), str(self.app.mapping_tab))

    def test_invalid_amount_stays_on_review_and_shows_error(self):
        self.app.fields["total"].set("999")
        with patch("approval_generator.gui.messagebox.showerror") as showerror:
            self.app.review_button.invoke()
        self.assertFalse(self.app.quotes[0].reviewed)
        self.assertEqual(self.app.tabs.select(), str(self.app.review_tab))
        self.assertIn("완료하지 못했습니다", self.app.workflow_notice.get())
        showerror.assert_called_once()

    def test_editing_reviewed_quote_restores_pending_state(self):
        self.app.review_button.invoke()
        self.app.tabs.select(self.app.review_tab)
        self.app.fields["title"].set("수정된 품의 제목")
        self.app.update()
        self.assertFalse(self.app.quotes[0].reviewed)
        self.assertIn("검토 필요", self.app.quote_list.get(0))
        self.assertIn("검토 필요", self.app.review_state.get())
        self.assertIn("다시", self.app.workflow_notice.get())


if __name__ == "__main__":
    unittest.main()
