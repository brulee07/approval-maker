from __future__ import annotations

import copy
import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pypdfium2 as pdfium
from openpyxl.utils import get_column_letter
from PIL import ImageTk

from . import __version__
from .extract import extract_quote
from .models import ALL_FIELDS, FIELDS, ITEM_FIELDS, NUMERIC_FIELDS, Item, Mapping, Profile, Quote, UserError, display, number
from .pricing import approval_total, prepare_approval_quote
from .render import pdf_size, save_result
from .templates import preview_grid, profile_for_template, sheet_names, table_mappings


FILE_TYPES = [("견적서", "*.pdf *.xlsx *.xlsm *.xls"), ("모든 파일", "*.*")]
TEMPLATE_TYPES = [("품의서 양식", "*.pdf *.xlsx *.xlsm *.xls")]


def open_file(path: str):
    if sys.platform == "win32":
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


class App(tk.Tk):
    def __init__(self):
        from .bootstrap import prepare_tk
        prepare_tk()
        super().__init__()
        self.title(f"품의서 메이커 v{__version__} · 견적서 일괄 변환")
        self.geometry("1380x900")
        self.minsize(1100, 760)
        self.configure(bg="#f4f6fa")
        self.quotes: list[Quote] = []
        self.current: int | None = None
        self.profile: Profile | None = None
        self.loading = False
        self.events = queue.Queue()
        self.busy = False
        self.pdf_image = None
        self.preview_scale = 1.0
        self.drag_start = None
        self.pdf_rect = None
        self.result_paths = {}
        self.generation_records = {}
        self._style()
        self._build()
        self.after(100, self._poll)

    def _style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", font=("맑은 고딕", 10), background="#f4f6fa", foreground="#1e293b")
        style.configure("TFrame", background="#f4f6fa")
        style.configure("TLabel", background="#f4f6fa")
        style.configure("Title.TLabel", font=("맑은 고딕", 22, "bold"))
        style.configure("Muted.TLabel", foreground="#64748b")
        style.configure("Notice.TLabel", background="#e8effc", foreground="#17439c", padding=(12, 9))
        style.configure("Reviewed.TLabel", foreground="#166534", font=("맑은 고딕", 10, "bold"))
        style.configure("Pending.TLabel", foreground="#92400e", font=("맑은 고딕", 10, "bold"))
        style.configure("TButton", padding=(12, 7))
        style.configure("Accent.TButton", background="#2458d3", foreground="white")
        style.map("Accent.TButton", background=[("active", "#1944ad")])
        style.configure("Treeview", rowheight=31, fieldbackground="white", background="white")
        style.configure("Treeview.Heading", font=("맑은 고딕", 10, "bold"), padding=6)
        style.configure("TNotebook.Tab", padding=(20, 10))

    def _build(self):
        header = ttk.Frame(self, padding=(24, 18, 24, 12))
        header.pack(fill="x")
        ttk.Label(header, text="품의서 메이커", style="Title.TLabel").pack(side="left")
        ttk.Label(header, text="견적서 검토 → 양식 연결 → 파일 생성", style="Muted.TLabel").pack(side="left", padx=24)
        ttk.Button(header, text="사용 안내", command=self.help).pack(side="right")
        ttk.Button(header, text="샘플로 체험", command=lambda: self.guard(self.load_demo)).pack(side="right", padx=8)
        self.workflow_notice = tk.StringVar(value="1단계: 견적서를 추가하고 내용을 확인한 뒤 '검토 완료'를 누르세요.")
        ttk.Label(self, textvariable=self.workflow_notice, style="Notice.TLabel", wraplength=1040).pack(fill="x", padx=24, pady=(0, 10))
        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill="both", expand=True, padx=24, pady=(0, 12))
        self.review_tab = ttk.Frame(self.tabs, padding=16)
        self.mapping_tab = ttk.Frame(self.tabs, padding=16)
        self.export_tab = ttk.Frame(self.tabs, padding=16)
        for frame, text in [(self.review_tab, "1  견적서 검토"), (self.mapping_tab, "2  양식 연결"), (self.export_tab, "3  일괄 생성")]:
            self.tabs.add(frame, text=text)
        self.status = tk.StringVar(value="견적서를 추가해 주세요. 파일은 이 컴퓨터에서 처리됩니다.")
        ttk.Label(self, textvariable=self.status, style="Muted.TLabel", padding=(24, 0, 24, 12)).pack(side="bottom", fill="x", before=self.tabs)
        self._review_ui()
        self._mapping_ui()
        self._export_ui()
        self.tabs.bind("<<NotebookTabChanged>>", lambda _event: self.guard(self._on_tab_changed))

    def _on_tab_changed(self):
        if self.tabs.select() == str(self.export_tab) and not self.busy:
            self._commit()
            self._refresh_export_comparison()

    def guard(self, callback):
        try:
            return callback()
        except Exception as exc:
            messagebox.showerror("확인이 필요합니다", str(exc), parent=self)
            return None

    def _run(self, label, task, done):
        if self.busy:
            return
        self.busy = True
        progress = tk.Toplevel(self)
        progress.title("처리 중")
        progress.transient(self)
        progress.resizable(False, False)
        progress.protocol("WM_DELETE_WINDOW", lambda: None)
        progress.geometry(f"480x140+{self.winfo_rootx() + 200}+{self.winfo_rooty() + 200}")
        self.progress_label = ttk.Label(progress, text=label, padding=20, wraplength=440)
        self.progress_label.pack(fill="x")
        bar = ttk.Progressbar(progress, mode="indeterminate")
        bar.pack(fill="x", padx=20, pady=8)
        bar.start()
        progress.grab_set()

        def worker():
            try:
                self.events.put(("done", task(), done, progress))
            except Exception as exc:
                self.events.put(("error", exc, None, progress))
        threading.Thread(target=worker, daemon=True).start()

    def _poll(self):
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "status":
                    self.progress_label.configure(text=event[1])
                    continue
                kind, value, callback, window = event
                window.grab_release()
                window.destroy()
                self.busy = False
                if kind == "error":
                    messagebox.showerror("처리 실패", str(value), parent=self)
                else:
                    self.guard(lambda: callback(value))
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _review_ui(self):
        toolbar = ttk.Frame(self.review_tab)
        toolbar.pack(fill="x", pady=(0, 12))
        ttk.Button(toolbar, text="+ 견적서 추가", style="Accent.TButton", command=self.add_quotes).pack(side="left")
        self.ocr = tk.BooleanVar(value=False)
        ttk.Checkbutton(toolbar, text="스캔 PDF에 OCR 사용 (Tesseract 필요)", variable=self.ocr).pack(side="left", padx=16)
        ttk.Label(toolbar, text="PDF · XLSX · XLSM · XLS", style="Muted.TLabel").pack(side="right")
        body = ttk.Panedwindow(self.review_tab, orient="horizontal")
        body.pack(fill="both", expand=True)
        sidebar = ttk.Frame(body, width=240)
        body.add(sidebar, weight=1)
        ttk.Label(sidebar, text="견적서 목록").pack(anchor="w", pady=(0, 6))
        self.quote_list = tk.Listbox(sidebar, font=("맑은 고딕", 10), width=30, bd=0, highlightthickness=1,
                                    highlightbackground="#d7dfeb", selectbackground="#dbe7ff", selectforeground="#1647aa", exportselection=False)
        self.quote_list.pack(fill="both", expand=True)
        self.quote_list.bind("<<ListboxSelect>>", self.select_quote)
        actions = ttk.Frame(sidebar)
        actions.pack(fill="x", pady=8)
        ttk.Button(actions, text="원본 열기", command=lambda: self.guard(self.open_source)).pack(side="left")
        ttk.Button(actions, text="제거", command=self.remove_quote).pack(side="right")
        editor = ttk.Frame(body, padding=(16, 0, 0, 0))
        body.add(editor, weight=4)
        self.review_title = tk.StringVar(value="견적서를 선택해 주세요")
        ttk.Label(editor, textvariable=self.review_title, font=("맑은 고딕", 12, "bold")).pack(anchor="w", pady=(0, 8))
        self.review_state = tk.StringVar(value="검토할 견적서가 없습니다.")
        self.review_state_label = ttk.Label(editor, textvariable=self.review_state, style="Pending.TLabel")
        self.review_state_label.pack(anchor="w", pady=(0, 6))
        form = ttk.Frame(editor)
        form.pack(fill="x")
        self.fields = {}
        for idx, (key, label) in enumerate(FIELDS.items()):
            row, pair = divmod(idx, 3)
            ttk.Label(form, text=label).grid(row=row * 2, column=pair, sticky="w", pady=(5, 2))
            value = tk.StringVar()
            value.trace_add("write", self._mark_dirty)
            self.fields[key] = value
            ttk.Entry(form, textvariable=value).grid(row=row * 2 + 1, column=pair, sticky="ew", padx=(0, 12))
            form.columnconfigure(pair, weight=1)
        item_toolbar = ttk.Frame(editor)
        item_toolbar.pack(fill="x", pady=(14, 6))
        ttk.Label(item_toolbar, text="품목", font=("맑은 고딕", 11, "bold")).pack(side="left")
        ttk.Button(item_toolbar, text="+ 추가", command=lambda: self.edit_item()).pack(side="left", padx=(12, 4))
        ttk.Button(item_toolbar, text="수정", command=lambda: self.edit_item(True)).pack(side="left")
        ttk.Button(item_toolbar, text="삭제", command=self.remove_item).pack(side="left", padx=4)
        ttk.Button(item_toolbar, text="금액 계산", command=lambda: self.guard(self.calculate)).pack(side="right")
        self.vat_included = tk.BooleanVar(value=False)
        self.vat_included.trace_add("write", self._mark_dirty)
        ttk.Checkbutton(item_toolbar, text="원본 품목 금액에 부가세 포함", variable=self.vat_included).pack(side="right", padx=8)
        items_frame = ttk.Frame(editor)
        items_frame.pack(fill="both", expand=True)
        self.items_tree = ttk.Treeview(items_frame, columns=list(ITEM_FIELDS), show="headings", height=6, selectmode="browse")
        for key, label in ITEM_FIELDS.items():
            self.items_tree.heading(key, text=label)
            self.items_tree.column(key, width=140 if key == "description" else 90, minwidth=50, anchor="e" if key in NUMERIC_FIELDS else "w")
        items_scroll = ttk.Scrollbar(items_frame, orient="vertical", command=self.items_tree.yview)
        items_scroll.pack(side="right", fill="y")
        self.items_tree.configure(yscrollcommand=items_scroll.set)
        self.items_tree.pack(fill="both", expand=True)
        self.items_tree.bind("<Double-1>", lambda _event: self.edit_item(True))
        self.review_notes = tk.Text(editor, height=4, wrap="word", font=("맑은 고딕", 9), bg="#eef2f8", relief="flat", padx=8, pady=8)
        self.review_notes.pack(side="bottom", fill="x", pady=8, before=items_frame)
        footer = ttk.Frame(editor)
        footer.pack(side="bottom", fill="x", before=self.review_notes)
        ttk.Button(footer, text="추출 원문 보기", command=self.show_raw).pack(side="left")
        self.review_button = ttk.Button(footer, text="검토 완료 → 다음 단계", style="Accent.TButton", command=lambda: self.guard(self.mark_reviewed), state="disabled")
        self.review_button.pack(side="right")

    def _mark_dirty(self, *_args):
        if not self.loading and self.current is not None:
            self.quotes[self.current].reviewed = False
            self._refresh_list()
            self.workflow_notice.set("내용이 변경되었습니다. 확인 후 '검토 완료'를 다시 눌러 주세요.")

    def _refresh_list(self):
        self.quote_list.delete(0, "end")
        for quote in self.quotes:
            self.quote_list.insert("end", f"[{'검토 완료' if quote.reviewed else '검토 필요'}] {Path(quote.source).name}")
        if self.current is not None and self.current < len(self.quotes):
            self.quote_list.selection_set(self.current)
        self.queue_summary.set(f"견적서 {len(self.quotes)}개  /  검토 완료 {sum(q.reviewed for q in self.quotes)}개") if hasattr(self, "queue_summary") else None
        if self.current is None:
            self.review_state.set("검토할 견적서가 없습니다.")
            self.review_state_label.configure(style="Pending.TLabel")
            self.review_button.configure(state="disabled", text="검토 완료 → 다음 단계")
        elif self.quotes[self.current].reviewed:
            self.review_state.set("검토 완료 · 다음 단계로 진행할 수 있습니다.")
            self.review_state_label.configure(style="Reviewed.TLabel")
            self.review_button.configure(state="normal", text="다음 단계로 이동")
        else:
            self.review_state.set("검토 필요 · 원본과 내용을 확인해 주세요.")
            self.review_state_label.configure(style="Pending.TLabel")
            self.review_button.configure(state="normal", text="검토 완료 → 다음 단계")

    def _commit(self):
        if self.current is None:
            return
        values = {key: number(var.get()) if key in NUMERIC_FIELDS else var.get().strip() for key, var in self.fields.items()}
        quote = self.quotes[self.current]
        quote.amount_basis = "inclusive" if self.vat_included.get() else "exclusive"
        for key, value in values.items():
            setattr(quote, key, value)

    def _load_quote(self, index: int | None):
        self.current = index
        self.loading = True
        quote = self.quotes[index] if index is not None else Quote()
        for key, value in self.fields.items():
            value.set(display(getattr(quote, key)))
        self.vat_included.set(quote.amount_basis == "inclusive")
        self.loading = False
        self.review_title.set(Path(quote.source).name if quote.source else "견적서를 선택해 주세요")
        self._refresh_items()
        self._notes(quote)
        self._refresh_list()

    def _notes(self, quote):
        notes = quote.warnings + quote.issues()
        if quote.source:
            notes.append("검토 화면은 원본 금액입니다. 품의서 출력 시 부가세를 품목에 포함하고, 필요한 경우 정수 단가로 행을 나눕니다.")
        self.review_notes.configure(state="normal")
        self.review_notes.delete("1.0", "end")
        self.review_notes.insert("1.0", "\n".join(notes) if quote.source else "파일을 추가하면 자동 추출 결과가 표시됩니다.")
        self.review_notes.configure(state="disabled")

    def select_quote(self, _event=None):
        selection = self.quote_list.curselection()
        if selection and selection[0] != self.current:
            index = selection[0]
            try:
                self._commit()
            except UserError as exc:
                messagebox.showerror("숫자 확인", str(exc), parent=self)
                self._refresh_list()
                return
            self._load_quote(index)

    def add_quotes(self):
        paths = filedialog.askopenfilenames(title="견적서 선택 (여러 파일 가능)", filetypes=FILE_TYPES)
        if not paths:
            return
        use_ocr = self.ocr.get()
        existing = {q.source for q in self.quotes}
        paths = [p for p in paths if str(Path(p).resolve()) not in existing]
        if not paths:
            return

        def work():
            results, errors = [], []
            for idx, path in enumerate(paths, 1):
                self.events.put(("status", f"견적서 읽는 중 ({idx}/{len(paths)})\n{Path(path).name}"))
                try:
                    results.append(extract_quote(path, use_ocr))
                except Exception as exc:
                    errors.append(f"{Path(path).name}: {exc}")
            return results, errors

        def done(result):
            quotes, errors = result
            self.quotes.extend(quotes)
            if self.current is None and self.quotes:
                self._load_quote(0)
            else:
                self._refresh_list()
            self.status.set(f"{len(quotes)}개 견적서를 읽었습니다. 원본과 대조한 뒤 검토 완료를 눌러 주세요.")
            self.workflow_notice.set("견적서를 불러왔습니다. 내용을 확인하고 '검토 완료'를 누르면 다음 단계로 이동합니다.")
            if errors:
                messagebox.showerror("일부 파일 읽기 실패", "\n\n".join(errors), parent=self)
        self._run("견적서를 읽고 있습니다…", work, done)

    def remove_quote(self):
        if self.current is None:
            return
        del self.quotes[self.current]
        self._load_quote(min(self.current, len(self.quotes) - 1) if self.quotes else None)

    def open_source(self):
        if self.current is not None:
            open_file(self.quotes[self.current].source)

    def _refresh_items(self):
        self.items_tree.delete(*self.items_tree.get_children())
        if self.current is not None:
            for idx, item in enumerate(self.quotes[self.current].items):
                self.items_tree.insert("", "end", iid=str(idx), values=[display(getattr(item, key)) for key in ITEM_FIELDS])

    def edit_item(self, existing=False):
        if self.current is None:
            return
        selection = self.items_tree.selection()
        if existing and not selection:
            return
        index = int(selection[0]) if existing else None
        quote = self.quotes[self.current]
        item = quote.items[index] if index is not None else Item()
        dialog = tk.Toplevel(self)
        dialog.title("품목 수정" if existing else "품목 추가")
        dialog.transient(self)
        dialog.grab_set()
        form = ttk.Frame(dialog, padding=20)
        form.pack(fill="both", expand=True)
        values = {}
        for idx, (key, label) in enumerate(ITEM_FIELDS.items()):
            ttk.Label(form, text=label).grid(row=idx, column=0, sticky="w", pady=6)
            values[key] = tk.StringVar(value=display(getattr(item, key)))
            ttk.Entry(form, textvariable=values[key], width=48).grid(row=idx, column=1, padx=16)

        def save():
            try:
                result = Item(**{key: number(value.get()) if key in NUMERIC_FIELDS else value.get().strip() for key, value in values.items()})
                if not result.description:
                    raise UserError("품명을 입력해 주세요.")
                if result.amount is None and result.quantity is not None and result.unit_price is not None:
                    result.amount = result.quantity * result.unit_price
                if index is None:
                    quote.items.append(result)
                else:
                    quote.items[index] = result
                quote.reviewed = False
                self._refresh_items()
                self._refresh_list()
                self._notes(quote)
                dialog.destroy()
            except UserError as exc:
                messagebox.showerror("품목 확인", str(exc), parent=dialog)
        ttk.Button(form, text="저장", style="Accent.TButton", command=save).grid(row=len(ITEM_FIELDS), column=1, sticky="e", pady=(16, 0))

    def remove_item(self):
        selection = self.items_tree.selection()
        if self.current is not None and selection:
            quote = self.quotes[self.current]
            del quote.items[int(selection[0])]
            quote.reviewed = False
            self._refresh_items()
            self._refresh_list()
            self._notes(quote)

    def calculate(self):
        if self.current is None:
            return
        self._commit()
        quote = self.quotes[self.current]
        for item in quote.items:
            if item.amount is None and item.quantity is not None and item.unit_price is not None:
                item.amount = item.quantity * item.unit_price
        if not quote.items or any(item.amount is None for item in quote.items):
            raise UserError("각 품목의 금액을 먼저 입력해 주세요.")
        if quote.amount_basis == "inclusive":
            quote.total = sum(i.amount for i in quote.items)
            if quote.vat is not None:
                quote.subtotal = quote.total - quote.vat
        else:
            quote.subtotal = sum(i.amount for i in quote.items)
            if quote.vat is not None:
                quote.total = quote.subtotal + quote.vat
        quote.reviewed = False
        self._load_quote(self.current)
        self.status.set("품목 합계를 선택한 부가세 포함/별도 기준으로 계산했습니다. 부가세 자체는 변경하지 않았습니다.")

    def mark_reviewed(self):
        if self.current is None:
            raise UserError("견적서를 추가해 주세요.")
        self._commit()
        quote = self.quotes[self.current]
        self._notes(quote)
        issues = quote.issues()
        if issues:
            self.workflow_notice.set("검토를 완료하지 못했습니다. 안내된 항목을 수정한 뒤 다시 눌러 주세요.")
            for index, item in enumerate(quote.items):
                if any("(cid:" in getattr(item, key) for key in ("description", "specification", "unit", "note")):
                    self.items_tree.selection_set(str(index))
                    self.items_tree.focus(str(index))
                    self.items_tree.see(str(index))
                    break
            raise UserError("\n".join(issues))
        quote.reviewed = True
        self._refresh_list()
        completed_name = Path(quote.source).name
        order = list(range(self.current + 1, len(self.quotes))) + list(range(self.current))
        pending = next((index for index in order if not self.quotes[index].reviewed), None)
        if pending is not None:
            self._load_quote(pending)
            self.quote_list.see(pending)
            self.tabs.select(self.review_tab)
            done = sum(item.reviewed for item in self.quotes)
            message = f"{completed_name} 검토 완료 ({done}/{len(self.quotes)}개). 다음 견적서의 내용을 확인해 주세요."
        elif self.profile is not None and self.profile.mappings:
            self.tabs.select(self.export_tab)
            message = f"견적서 {len(self.quotes)}개 검토 완료. 저장 폴더를 확인하고 '전체 품의서 생성'을 눌러 주세요."
        else:
            self.tabs.select(self.mapping_tab)
            message = f"견적서 {len(self.quotes)}개 검토 완료. '양식 선택' 또는 '연결 설정 불러오기'로 품의서 양식을 연결해 주세요."
        self.workflow_notice.set(message)
        self.status.set(message)

    def show_raw(self):
        if self.current is None:
            return
        dialog = tk.Toplevel(self)
        dialog.title("추출 원문 · 원본 파일과 대조해 주세요")
        dialog.geometry("980x700")
        content = tk.Text(dialog, wrap="none", font=("맑은 고딕", 10), padx=16, pady=16)
        scrollbar = ttk.Scrollbar(dialog, command=content.yview)
        scrollbar.pack(side="right", fill="y")
        content.configure(yscrollcommand=scrollbar.set)
        content.pack(fill="both", expand=True)
        content.insert("1.0", self.quotes[self.current].raw_text or "읽을 수 있는 텍스트가 없습니다. 원본 열기로 확인하고 직접 입력해 주세요.")
        content.configure(state="disabled")

    def _mapping_ui(self):
        toolbar = ttk.Frame(self.mapping_tab)
        toolbar.pack(fill="x", pady=(0, 10))
        ttk.Button(toolbar, text="양식 선택", style="Accent.TButton", command=lambda: self.guard(self.choose_template)).pack(side="left")
        ttk.Button(toolbar, text="연결 설정 불러오기", command=lambda: self.guard(self.load_profile)).pack(side="left", padx=8)
        ttk.Button(toolbar, text="연결 설정 저장", command=lambda: self.guard(self.save_profile)).pack(side="left")
        ttk.Button(toolbar, text="열 제목으로 자동 연결", command=lambda: self.guard(self.auto_map)).pack(side="left", padx=8)
        ttk.Button(toolbar, text="양식 원본 열기", command=lambda: self.guard(lambda: open_file(self.profile.template) if self.profile else None)).pack(side="right")
        self.template_label = tk.StringVar(value="사용자 품의서 양식(PDF, XLSX 또는 XLS)을 선택해 주세요.")
        ttk.Label(self.mapping_tab, textvariable=self.template_label, style="Muted.TLabel").pack(anchor="w", pady=(0, 10))
        body = ttk.Panedwindow(self.mapping_tab, orient="horizontal")
        body.pack(fill="both", expand=True)
        body.bind("<Configure>", lambda event: body.sashpos(0, min(440, max(360, event.width // 3))) if len(body.panes()) == 2 else None)
        config = ttk.Frame(body, width=420, padding=(0, 0, 16, 0))
        body.add(config, weight=2)
        self.mapping_hint = tk.StringVar(value="엑셀은 셀을 클릭하고, PDF는 입력할 빈 영역을 드래그합니다.")
        ttk.Label(config, textvariable=self.mapping_hint, wraplength=410).pack(anchor="w", pady=(0, 8))
        default_row = ttk.Frame(config)
        default_row.pack(fill="x", pady=(0, 6))
        ttk.Label(default_row, text="빈 단위 기본값").pack(side="left")
        self.default_unit = tk.StringVar()
        ttk.Entry(default_row, textvariable=self.default_unit, width=12).pack(side="left", padx=10)
        form = ttk.Frame(config)
        form.pack(fill="x")
        self.map_values = {}
        self.map_field = tk.StringVar(value=ALL_FIELDS["supplier"])
        ttk.Label(form, text="입력할 항목").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Combobox(form, textvariable=self.map_field, values=list(ALL_FIELDS.values()), state="readonly", width=24).grid(row=0, column=1, sticky="ew", pady=4)
        definitions = [("sheet", "엑셀 시트", ""), ("cell", "시작 셀", "C5"), ("page", "PDF 페이지", "1"),
                       ("x", "왼쪽 x (pt)", "40"), ("y", "위쪽 y (pt)", "100"), ("width", "영역 폭 (pt)", "180"),
                       ("height", "영역 높이 (pt)", "24"), ("count", "품목 칸 수", "1"),
                       ("step", "반복 간격 (행 / pt)", "1"), ("font_size", "PDF 글자 크기", "10"), ("align", "PDF 정렬", "left")]
        self.map_widgets = {}
        for idx, (key, label, default) in enumerate(definitions, 1):
            ttk.Label(form, text=label).grid(row=idx, column=0, sticky="w", pady=3)
            var = tk.StringVar(value=default)
            self.map_values[key] = var
            if key in {"sheet", "align"}:
                widget = ttk.Combobox(form, textvariable=var, state="readonly", values=["left", "center", "right"] if key == "align" else [])
                if key == "sheet":
                    widget.bind("<<ComboboxSelected>>", lambda _event: self.guard(self.show_template))
            else:
                widget = ttk.Entry(form, textvariable=var)
            widget.grid(row=idx, column=1, sticky="ew", pady=3)
            self.map_widgets[key] = widget
        form.columnconfigure(1, weight=1)
        actions = ttk.Frame(config)
        actions.pack(fill="x", pady=10)
        ttk.Button(actions, text="연결 추가", command=lambda: self.guard(self.add_mapping)).pack(side="left")
        ttk.Button(actions, text="선택 수정", command=lambda: self.guard(lambda: self.add_mapping(True))).pack(side="left", padx=4)
        ttk.Button(actions, text="삭제", command=self.remove_mapping).pack(side="right")
        self.mapping_tree = ttk.Treeview(config, columns=("field", "target"), show="headings", height=5, selectmode="browse")
        self.mapping_tree.heading("field", text="항목")
        self.mapping_tree.heading("target", text="입력 위치")
        self.mapping_tree.column("field", width=110)
        self.mapping_tree.column("target", width=230)
        mapping_scroll = ttk.Scrollbar(config, orient="vertical", command=self.mapping_tree.yview)
        mapping_scroll.pack(side="right", fill="y")
        self.mapping_tree.configure(yscrollcommand=mapping_scroll.set)
        self.mapping_tree.pack(fill="both", expand=True)
        self.mapping_tree.bind("<<TreeviewSelect>>", self.select_mapping)
        preview = ttk.Frame(body)
        body.add(preview, weight=4)
        nav = ttk.Frame(preview)
        nav.pack(fill="x", pady=(0, 6))
        ttk.Label(nav, text="양식 미리보기", font=("맑은 고딕", 11, "bold")).pack(side="left")
        ttk.Button(nav, text="현재 시트 / 페이지 보기", command=lambda: self.guard(self.show_template)).pack(side="right")
        self.preview = ttk.Frame(preview)
        self.preview.pack(fill="both", expand=True)

    def choose_template(self):
        path = filedialog.askopenfilename(title="빈 품의서 양식 선택", filetypes=TEMPLATE_TYPES)
        if not path:
            return
        candidate, message = profile_for_template(path, Path(__file__).resolve().parent.parent / "profiles")
        self.profile = candidate
        self.setup_template()
        self._refresh_mappings()
        self.workflow_notice.set(message)
        self.status.set(message)

    def setup_template(self):
        if self.profile is None:
            return
        self.profile.check_template()
        is_pdf = Path(self.profile.template).suffix.lower() == ".pdf"
        for key, widget in self.map_widgets.items():
            enabled = key in {"count", "step"} or (is_pdf and key not in {"sheet", "cell"}) or (not is_pdf and key in {"sheet", "cell"})
            widget.configure(state=("readonly" if key in {"sheet", "align"} else "normal") if enabled else "disabled")
        if not is_pdf:
            sheets = sheet_names(self.profile.template)
            self.map_widgets["sheet"].configure(values=sheets)
            mapped_sheet = next((m.sheet for m in self.profile.mappings if m.sheet in sheets), sheets[0])
            self.map_values["sheet"].set(mapped_sheet)
            self.map_values["step"].set("1")
            self.mapping_hint.set("셀을 클릭해 입력 위치를 선택하세요. 병합 셀은 왼쪽 위 셀을 지정합니다. 품목은 시작 셀·칸 수·행 간격을 설정하세요.")
        else:
            self.map_values["page"].set("1")
            self.map_values["step"].set("28")
            self.mapping_hint.set("빈 입력 영역을 드래그하세요. 품목은 첫 번째 행의 영역과 칸 수·행 높이(pt)를 설정합니다. 기존 글자는 지워지지 않습니다.")
        self.template_label.set(self.profile.template)
        self.default_unit.set(self.profile.default_unit)
        self.show_template()

    def show_template(self):
        if self.profile is None:
            return
        for widget in self.preview.winfo_children():
            widget.destroy()
        if Path(self.profile.template).suffix.lower() == ".pdf":
            page_number = int(self.map_values["page"].get())
            pdf_size(self.profile.template, page_number)
            document = pdfium.PdfDocument(self.profile.template)
            try:
                page = document[page_number - 1]
                try:
                    self.preview_scale = min(1.5, max(600, self.preview.winfo_width() - 24) / page.get_width())
                    bitmap = page.render(scale=self.preview_scale)
                    try:
                        image = bitmap.to_pil().copy()
                    finally:
                        bitmap.close()
                finally:
                    page.close()
            finally:
                document.close()
            canvas = tk.Canvas(self.preview, background="#e2e8f0", highlightthickness=0)
            scrollbar = ttk.Scrollbar(self.preview, orient="vertical", command=canvas.yview)
            scrollbar.pack(side="right", fill="y")
            horizontal = ttk.Scrollbar(self.preview, orient="horizontal", command=canvas.xview)
            horizontal.pack(side="bottom", fill="x")
            canvas.configure(yscrollcommand=scrollbar.set, xscrollcommand=horizontal.set)
            canvas.pack(fill="both", expand=True)
            self.pdf_image = ImageTk.PhotoImage(image)
            canvas.create_image(0, 0, image=self.pdf_image, anchor="nw")
            canvas.configure(scrollregion=(0, 0, image.width, image.height))
            self.pdf_canvas = canvas
            self.pdf_rect = None
            canvas.bind("<ButtonPress-1>", self.pdf_press)
            canvas.bind("<B1-Motion>", self.pdf_drag)
            canvas.bind("<ButtonRelease-1>", self.pdf_release)
            for mapping in self.profile.mappings:
                if mapping.page == page_number:
                    for idx in range(mapping.count if mapping.field.startswith("items.") else 1):
                        s = self.preview_scale
                        y = mapping.y + (idx * mapping.step if mapping.field.startswith("items.") else 0)
                        canvas.create_rectangle(mapping.x * s, y * s, (mapping.x + mapping.width) * s, (y + mapping.height) * s, outline="#2563eb", width=1, tags="mapped")
        else:
            sheet_name = self.map_values["sheet"].get()
            rows = preview_grid(self.profile.template, sheet_name)
            columns = [get_column_letter(c) for c in range(1, len(rows[0]) + 1)]
            if rows:
                tree = ttk.Treeview(self.preview, columns=columns, show="tree headings")
                tree.heading("#0", text="행")
                tree.column("#0", width=50, stretch=False)
                for col in columns:
                    tree.heading(col, text=col)
                    tree.column(col, width=100, minwidth=80, stretch=False)
                for idx, row in enumerate(rows, 1):
                    tree.insert("", "end", iid=str(idx), text=str(idx), values=[display(v) for v in row])
                vertical = ttk.Scrollbar(self.preview, orient="vertical", command=tree.yview)
                vertical.pack(side="right", fill="y")
                horizontal = ttk.Scrollbar(self.preview, orient="horizontal", command=tree.xview)
                horizontal.pack(side="bottom", fill="x")
                tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
                tree.pack(fill="both", expand=True)
                tree.bind("<ButtonRelease-1>", lambda event: self.pick_cell(tree, event))

    def pick_cell(self, tree, event):
        row, col = tree.identify_row(event.y), tree.identify_column(event.x)
        if row and col != "#0":
            self.map_values["cell"].set(f"{get_column_letter(int(col[1:]))}{row}")

    def pdf_press(self, event):
        canvas = self.pdf_canvas
        self.drag_start = (canvas.canvasx(event.x), canvas.canvasy(event.y))
        if self.pdf_rect:
            canvas.delete(self.pdf_rect)
        self.pdf_rect = canvas.create_rectangle(*self.drag_start, *self.drag_start, outline="#e26918", width=2)

    def pdf_drag(self, event):
        if self.drag_start:
            self.pdf_canvas.coords(self.pdf_rect, *self.drag_start, self.pdf_canvas.canvasx(event.x), self.pdf_canvas.canvasy(event.y))

    def pdf_release(self, event):
        if not self.drag_start:
            return
        start_x, start_y = self.drag_start
        end_x, end_y = self.pdf_canvas.canvasx(event.x), self.pdf_canvas.canvasy(event.y)
        s = self.preview_scale
        for key, value in {"x": min(start_x, end_x) / s, "y": min(start_y, end_y) / s,
                           "width": abs(end_x - start_x) / s, "height": abs(end_y - start_y) / s}.items():
            self.map_values[key].set(f"{value:.1f}")
        self.drag_start = None

    def add_mapping(self, replace=False):
        if self.profile is None:
            raise UserError("양식을 먼저 선택해 주세요.")
        field = next(key for key, label in ALL_FIELDS.items() if label == self.map_field.get())
        values = {key: var.get().strip() for key, var in self.map_values.items()}
        for key in ("x", "y", "width", "height", "step", "font_size"):
            values[key] = float(values[key])
        for key in ("count", "page"):
            values[key] = int(values[key])
        if not field.startswith("items."):
            values["count"] = 1
        mapping = Mapping(field=field, **values)
        from .render import cell_address, validate_mappings
        validate_mappings(Profile(self.profile.template, self.profile.template_hash, [mapping]), Quote())
        if Path(self.profile.template).suffix.lower() != ".pdf":
            cell_address(mapping.cell)
        if replace:
            selected = self.mapping_tree.selection()
            if not selected:
                raise UserError("수정할 연결을 목록에서 선택해 주세요.")
            self.profile.mappings[int(selected[0])] = mapping
        else:
            self.profile.mappings.append(mapping)
        self._refresh_mappings()
        if Path(self.profile.template).suffix.lower() == ".pdf":
            self.show_template()

    def _refresh_mappings(self):
        self.mapping_tree.delete(*self.mapping_tree.get_children())
        if not self.profile:
            self.export_template.set("양식이 선택되지 않았습니다.")
            return
        count = len(self.profile.mappings)
        connection = f"입력 위치 {count}개 연결됨" if count else "입력 위치 연결 필요"
        self.export_template.set(f"출력 형식: {Path(self.profile.template).suffix.upper()[1:]}  ·  {Path(self.profile.template).name}  ·  {connection}")
        pdf = Path(self.profile.template).suffix.lower() == ".pdf"
        for idx, mapping in enumerate(self.profile.mappings):
            target = f"{mapping.page}쪽 ({mapping.x:g}, {mapping.y:g})" if pdf else f"{mapping.sheet}!{mapping.cell}"
            if mapping.field.startswith("items."):
                target += f" / {mapping.count}칸"
            self.mapping_tree.insert("", "end", iid=str(idx), values=(ALL_FIELDS.get(mapping.field, mapping.field), target))

    def select_mapping(self, _event=None):
        selected = self.mapping_tree.selection()
        if self.profile and selected:
            mapping = self.profile.mappings[int(selected[0])]
            self.map_field.set(ALL_FIELDS.get(mapping.field, mapping.field))
            for key, var in self.map_values.items():
                var.set(str(getattr(mapping, key)))

    def remove_mapping(self):
        selected = self.mapping_tree.selection()
        if self.profile and selected:
            del self.profile.mappings[int(selected[0])]
            self._refresh_mappings()
            self.guard(self.show_template)

    def save_profile(self):
        if self.profile is None:
            raise UserError("양식을 선택해 주세요.")
        path = filedialog.asksaveasfilename(title="연결 설정 저장", defaultextension=".json", initialfile="품의서_연결설정.json", filetypes=[("연결 설정", "*.json")])
        if path:
            self.profile.default_unit = self.default_unit.get().strip()
            self.profile.save(path)
            self.status.set(f"연결 설정 저장: {path}")

    def load_profile(self):
        path = filedialog.askopenfilename(title="연결 설정 불러오기", filetypes=[("연결 설정", "*.json")])
        if not path:
            return
        candidate = Profile.load(path)
        if not Path(candidate.template).is_file():
            replacement = filedialog.askopenfilename(title="이 설정에 사용한 원본 양식의 새 위치를 선택하세요", filetypes=TEMPLATE_TYPES)
            if not replacement:
                return
            candidate.template = replacement
        candidate.check_template()
        self.profile = candidate
        self.setup_template()
        self._refresh_mappings()

        message = f"저장된 입력 위치 {len(candidate.mappings)}개를 불러왔습니다. '3 일괄 생성'에서 품의서를 생성하세요." if candidate.mappings else "입력 위치가 없는 설정입니다. '열 제목으로 자동 연결' 또는 직접 연결을 완료해 주세요."
        self.workflow_notice.set(message)
        self.status.set(message)

    def _export_ui(self):
        ttk.Label(self.export_tab, text="견적서별 품의서 일괄 생성", font=("맑은 고딕", 18, "bold")).pack(anchor="w", pady=(4, 12))
        ttk.Label(self.export_tab, text="품의서에는 부가세 포함 단가·금액을 입력합니다. 별도 부가세는 품목에 배분하며, 정수 단가로 나누어떨어지지 않으면 같은 품목을 여러 행으로 표시합니다.", wraplength=1050).pack(anchor="w")
        self.queue_summary = tk.StringVar(value="견적서 0개  /  검토 완료 0개")
        ttk.Label(self.export_tab, textvariable=self.queue_summary, padding=(0, 18)).pack(anchor="w")
        self.export_template = tk.StringVar(value="양식이 선택되지 않았습니다.")
        ttk.Label(self.export_tab, textvariable=self.export_template, style="Muted.TLabel").pack(anchor="w")
        destination = ttk.Frame(self.export_tab)
        destination.pack(fill="x", pady=18)
        ttk.Label(destination, text="저장 폴더 (직접 지정)").pack(side="left", padx=(0, 12))
        self.output_dir = tk.StringVar(value="")
        ttk.Entry(destination, textvariable=self.output_dir).pack(side="left", fill="x", expand=True)
        ttk.Button(destination, text="폴더 선택", command=self.choose_output).pack(side="left", padx=8)
        actions = ttk.Frame(self.export_tab)
        actions.pack(fill="x", pady=(0, 16))
        ttk.Button(actions, text="품의서 금액 미리보기", command=lambda: self.guard(self.preview_pricing)).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="선택 견적서로 시험 생성", command=lambda: self.guard(lambda: self.generate(True))).pack(side="left")
        ttk.Button(actions, text="전체 품의서 생성", style="Accent.TButton", command=lambda: self.guard(self.generate)).pack(side="left", padx=10)
        ttk.Button(actions, text="결과 폴더 열기", command=lambda: self.guard(self.open_output_folder)).pack(side="right")
        ttk.Label(self.export_tab, text="총액 비교 · 단위: 원 (부가세 포함)", font=("맑은 고딕", 11, "bold")).pack(anchor="w", pady=(0, 8))
        results_frame = ttk.Frame(self.export_tab)
        results_frame.pack(fill="both", expand=True)
        self.results = ttk.Treeview(results_frame, columns=("source", "source_total", "output_total", "difference", "comparison", "status", "message"), show="headings")
        for key, label, width in [("source", "견적서", 220), ("source_total", "견적서 총액", 135),
                                  ("output_total", "품의서 총액", 135), ("difference", "차액", 85),
                                  ("comparison", "총액 비교", 105), ("status", "생성 결과", 90),
                                  ("message", "저장 위치 / 확인 사항", 270)]:
            self.results.heading(key, text=label)
            self.results.column(key, width=width, minwidth=75, anchor="e" if key in {"source_total", "output_total", "difference"} else "w")
        result_scroll = ttk.Scrollbar(results_frame, orient="vertical", command=self.results.yview)
        result_scroll.pack(side="right", fill="y")
        self.results.configure(yscrollcommand=result_scroll.set)
        self.results.pack(fill="both", expand=True)
        self.results.tag_configure("matched", foreground="#166534")
        self.results.tag_configure("error", foreground="#b91c1c")
        self.results.bind("<Double-1>", lambda _event: self.guard(self.open_result))
        ttk.Label(self.export_tab, text="'일치(예정)'는 생성 전 계산값입니다. 생성이 성공하면 '일치'로 표시됩니다. 차액은 품의서 총액 − 견적서 총액입니다.\n성공한 결과를 두 번 클릭하면 파일을 엽니다. 기존 파일은 덮어쓰지 않습니다.", style="Muted.TLabel", wraplength=1050).pack(side="bottom", anchor="w", pady=12, before=results_frame)

    def _export_signature(self, quote):
        return copy.deepcopy((quote, self.profile, self.default_unit.get().strip()))

    def _refresh_export_comparison(self):
        self.results.delete(*self.results.get_children())
        self.result_paths = {}
        for index, quote in enumerate(self.quotes):
            record = self.generation_records.get(id(quote))
            result = record[1] if record and record[0] == self._export_signature(quote) else None
            source_total = quote.total
            output_total = difference = None
            comparison, status, message, tag = "확인 필요", "생성 전", "생성 전 계산값입니다.", ""
            try:
                prepared = prepare_approval_quote(quote, self.default_unit.get().strip())
                output_total = approval_total(prepared)
                difference = output_total - source_total
                comparison = "일치(예정)"
                if not quote.reviewed:
                    status, message = "검토 필요", "원본과 대조한 뒤 검토 완료를 눌러 주세요."
            except UserError as exc:
                # 원본 합계가 맞지 않는 경우에도 확인 가능한 예정 합계와 차액을 보여 준다.
                if quote.items and all(item.amount is not None for item in quote.items):
                    base = sum((item.amount for item in quote.items), Decimal(0))
                    if quote.amount_basis == "inclusive":
                        output_total = base
                    elif quote.amount_basis == "exclusive" and quote.vat is not None:
                        output_total = base + quote.vat
                    if output_total is not None and source_total is not None:
                        difference = output_total - source_total
                comparison = "불일치" if difference is not None and difference != 0 else "확인 필요"
                status, message, tag = "생성 불가", str(exc), "error"
            if result:
                status, message = result["status"], result["message"]
                if status == "성공":
                    source_total = number(result["source_total"])
                    output_total = number(result["output_total"])
                    difference = output_total - source_total
                    comparison, tag = ("일치", "matched") if difference == 0 else ("불일치", "error")
                    self.result_paths[str(index)] = result["output"]
                else:
                    tag = "error"
            self.results.insert("", "end", iid=str(index), values=(Path(quote.source).name,
                                display(source_total) if source_total is not None else "미입력",
                                display(output_total) if output_total is not None else "—",
                                display(difference) if difference is not None else "—",
                                comparison, status, message), tags=(tag,) if tag else ())

    def choose_output(self):
        path = filedialog.askdirectory(title="결과 저장 폴더 선택")
        if path:
            self.output_dir.set(path)

    def open_output_folder(self):
        folder = self.output_dir.get().strip()
        if not folder:
            raise UserError("'폴더 선택'으로 결과 저장 폴더를 먼저 지정해 주세요.")
        destination = Path(folder).resolve()
        if not destination.is_dir():
            raise UserError("지정한 폴더가 아직 없습니다. 품의서를 생성하거나 기존 폴더를 선택해 주세요.")
        open_file(str(destination))

    def preview_pricing(self):
        self._commit()
        if self.current is None:
            raise UserError("미리 볼 견적서를 먼저 선택해 주세요.")
        source = self.quotes[self.current]
        prepared = prepare_approval_quote(source, self.default_unit.get().strip())
        output_total = approval_total(prepared)
        dialog = tk.Toplevel(self)
        dialog.title("품의서 금액 미리보기 · 부가세 포함")
        dialog.geometry("1000x620")
        dialog.minsize(760, 420)
        dialog.transient(self)
        body = ttk.Frame(dialog, padding=16)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=Path(source.source).name, font=("맑은 고딕", 12, "bold")).pack(anchor="w")
        ttk.Label(body, text=f"원본 {len(source.items)}개 품목 → 출력 {len(prepared.items)}행  /  총액 {display(prepared.total)}원 (부가세 포함)", padding=(0, 10)).pack(anchor="w")
        ttk.Label(body, text="같은 품목의 행이 나뉘어도 수량 합계와 금액 합계는 유지됩니다. 별도 부가세 금액은 출력하지 않습니다.", wraplength=950).pack(anchor="w", pady=(0, 12))
        # 목록보다 먼저 하단 공간을 확보해 스크롤하거나 창을 줄여도 합계를 표시한다.
        total_bar = ttk.Frame(body, padding=(12, 12))
        total_bar.pack(side="bottom", fill="x", pady=(12, 0))
        total_bar.columnconfigure(1, weight=1)
        ttk.Label(total_bar, text="견적서 총액 (부가세 포함)", font=("맑은 고딕", 11)).grid(row=0, column=0, sticky="w")
        ttk.Label(total_bar, text=f"{display(source.total)}원", font=("맑은 고딕", 12)).grid(row=0, column=1, sticky="e")
        ttk.Label(total_bar, text="품의서 총합계 금액 (부가세 포함)", font=("맑은 고딕", 12, "bold")).grid(row=1, column=0, sticky="w", pady=6)
        ttk.Label(total_bar, text=f"{display(output_total)}원", font=("맑은 고딕", 16, "bold"),
                  foreground="#17439c").grid(row=1, column=1, sticky="e", pady=6)
        ttk.Label(total_bar, text="총액 검증", font=("맑은 고딕", 11)).grid(row=2, column=0, sticky="w")
        ttk.Label(total_bar, text=f"일치 · 차액 {display(output_total - source.total)}원", foreground="#166534",
                  font=("맑은 고딕", 11, "bold")).grid(row=2, column=1, sticky="e")
        container = ttk.Frame(body)
        container.pack(fill="both", expand=True)
        tree = ttk.Treeview(container, columns=("index", "description", "quantity", "unit", "price", "amount"), show="headings")
        for key, label, width in [("index", "행", 40), ("description", "품명", 400), ("quantity", "수량", 80),
                                  ("unit", "단위", 50), ("price", "단가 (부가세 포함)", 150), ("amount", "금액 (부가세 포함)", 160)]:
            tree.heading(key, text=label)
            tree.column(key, width=width, minwidth=40, anchor="w" if key in {"description", "unit"} else "e")
        scroll = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        scroll.pack(side="right", fill="y")
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(fill="both", expand=True)
        for index, item in enumerate(prepared.items, 1):
            tree.insert("", "end", values=(index, " ".join(item.description.split()), display(item.quantity), item.unit,
                                            display(item.unit_price), display(item.amount)))

    def generate(self, single=False):
        self._commit()
        if not self.profile:
            self.tabs.select(self.mapping_tab)
            raise UserError("양식을 선택하고 입력 위치를 연결해 주세요.")
        if not self.profile.mappings:
            self.tabs.select(self.mapping_tab)
            message = "양식의 입력 위치가 아직 연결되지 않았습니다.\n'열 제목으로 자동 연결' 또는 '연결 설정 불러오기'를 누른 뒤 다시 생성해 주세요."
            self.workflow_notice.set(message.replace("\n", " "))
            self.status.set(message.replace("\n", " "))
            raise UserError(message)
        if not self.quotes:
            raise UserError("견적서를 추가해 주세요.")
        selected = [self.quotes[self.current]] if single and self.current is not None else self.quotes
        unreviewed = [Path(q.source).name for q in selected if not q.reviewed]
        if unreviewed:
            raise UserError("검토 완료가 필요합니다:\n" + "\n".join(unreviewed))
        folder = self.output_dir.get().strip()
        if not folder:
            raise UserError("저장 폴더를 선택해 주세요.")
        self.profile.default_unit = self.default_unit.get().strip()
        profile, quotes = copy.deepcopy(self.profile), copy.deepcopy(selected)
        source_ids = [id(quote) for quote in selected]
        signatures = [self._export_signature(quote) for quote in selected]
        self.results.delete(*self.results.get_children())
        self.result_paths = {}

        def work():
            results = []
            for idx, quote in enumerate(quotes, 1):
                self.events.put(("status", f"품의서 생성 중 ({idx}/{len(quotes)})\n{Path(quote.source).name}"))
                try:
                    prepared = prepare_approval_quote(quote, profile.default_unit)
                    path = save_result(profile, quote, folder)
                    results.append({"source": quote.source, "status": "성공", "output": str(path), "message": str(path),
                                    "warnings": prepared.warnings, "source_amount_basis": quote.amount_basis,
                                    "output_item_count": len(prepared.items), "output_total": str(prepared.total),
                                    "source_total": str(quote.total), "total_difference": str(approval_total(prepared) - quote.total),
                                    "output_items": [{"description": item.description, "quantity": display(item.quantity),
                                                      "unit_price": display(item.unit_price), "amount": display(item.amount)} for item in prepared.items]})
                except Exception as exc:
                    results.append({"source": quote.source, "status": "실패", "output": "", "message": str(exc), "warnings": quote.warnings})
            report = {"created_at": datetime.now().astimezone().isoformat(), "template": profile.template, "template_sha256": profile.template_hash, "results": results}
            try:
                Path(folder).mkdir(parents=True, exist_ok=True)
                report_path = Path(folder) / ("생성결과_" + datetime.now().strftime("%Y%m%d_%H%M%S_%f") + ".json")
                report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
            except OSError as exc:
                results.append({"source": "처리 기록", "status": "실패", "output": "", "message": str(exc)})
            return results

        def done(results):
            for source_id, signature, result in zip(source_ids, signatures, results):
                self.generation_records[source_id] = (signature, result)
            self._refresh_export_comparison()
            for idx, result in enumerate(results[len(source_ids):]):
                self.results.insert("", "end", iid=f"report_{idx}", values=(result["source"], "—", "—", "—", "확인 필요", result["status"], result["message"]), tags=("error",))
            success = sum(r["status"] == "성공" for r in results)
            self.status.set(f"생성 완료: 성공 {success}개 / 실패 {len(results) - success}개. 결과 파일을 열어 입력 위치와 인쇄 모양을 확인해 주세요.")
        self._run("품의서를 생성하고 있습니다…", work, done)

    def open_result(self):
        selected = self.results.selection()
        if selected and selected[0] in self.result_paths:
            open_file(self.result_paths[selected[0]])
        elif selected:
            messagebox.showinfo("확인 사항", self.results.set(selected[0], "message"), parent=self)

    def help(self):
        messagebox.showinfo("사용 안내", "1. 견적서를 여러 개 추가하고 추출 결과를 원본과 대조합니다.\n"
                            "2. 품목·금액을 수정하고 각 견적서의 '검토 완료'를 누릅니다.\n"
                            "3. 빈 품의서 양식을 선택합니다. 엑셀 셀 또는 PDF 영역에 항목을 연결합니다.\n"
                            "4. 품목은 첫 칸, 칸 수, 반복 간격을 지정합니다. 연결 설정을 저장하면 재사용할 수 있습니다.\n"
                            "5. 시험 생성 후 원본과 비교하고 전체 품의서를 생성합니다.\n\n"
                            "PDF, XLSX, XLSM, XLS는 양식과 같은 형식으로 생성됩니다. XLS 출력에는 Microsoft Excel이 필요합니다.\n"
                            "별도 부가세는 품목에 포함합니다. 정수 단가를 위해 행을 나눌 수 있으며, 품의서 금액 미리보기에서 확인할 수 있습니다.\n"
                            "스캔 PDF의 OCR은 Tesseract 한국어 설치가 필요합니다.\n"
                            "엑셀 미리보기는 셀 위치 확인용입니다. 최종 서식은 Excel에서 확인하세요.\n"
                            "검토 중인 내용은 프로그램 종료 시 사라집니다. 연결 설정만 파일로 저장됩니다.", parent=self)

    def load_demo(self):
        from .demo import create_demo
        self._commit()
        folder = Path.cwd() / "examples" / ("체험_" + datetime.now().strftime("%Y%m%d_%H%M%S_%f"))
        files = create_demo(folder)
        start = len(self.quotes)
        for name in ("source_xlsx", "source_pdf"):
            self.quotes.append(extract_quote(str(files[name])))
        self.profile = Profile.load(str(files["excel_profile"]))
        self.setup_template()
        self._refresh_mappings()
        self._load_quote(start)
        self.tabs.select(self.review_tab)
        self.status.set("가상 견적서 2개와 연결된 엑셀 양식을 불러왔습니다. 각각 검토 완료 후 생성해 보세요. 체험 폴더에 PDF 연결 설정도 있습니다.")

    def auto_map(self):
        if not self.profile:
            raise UserError("양식을 먼저 선택해 주세요.")
        self.profile.mappings = table_mappings(self.profile.template, self.map_values["sheet"].get())
        self._refresh_mappings()
        self.workflow_notice.set(f"입력 위치 {len(self.profile.mappings)}개 연결 완료. 빈 단위 기본값을 확인한 뒤 '3 일괄 생성'에서 품의서를 생성하세요.")
        self.status.set("첫 행의 열 제목으로 연결했습니다. 2행부터 최대 1,000행을 입력합니다. 부가세 배분으로 나뉜 행도 포함합니다. 기본 단위와 연결 내용을 확인해 주세요.")


def main():
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except (OSError, AttributeError):
            pass
    App().mainloop()


if __name__ == "__main__":
    main()
