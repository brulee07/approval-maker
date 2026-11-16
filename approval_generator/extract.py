from __future__ import annotations

import re
import shutil
import subprocess
import tempfile
from datetime import date, datetime
from pathlib import Path

import pdfplumber
from openpyxl import load_workbook

from .models import Item, Quote, UserError, fingerprint, number
from .pdf_text_corrections import apply_verified_corrections


def normalize(value) -> str:
    return re.sub(r"[\s:：()\[\]·/._-]+", "", str(value or "")).lower()


HEADERS = {
    "description": {"품명", "품목", "품목명", "제품명", "물품명", "내역", "내용", "description", "item", "품명및규격"},
    "specification": {"규격", "사양", "모델", "모델명", "spec", "specification"},
    "unit": {"단위", "unit"},
    "quantity": {"수량", "qty", "quantity"},
    "unit_price": {"단가", "예상단가", "unitprice", "price"},
    "amount": {"금액", "공급가액", "공급가", "amount", "가격", "공급대가", "공급대가vat포함", "금액vat포함"},
    "note": {"비고", "remark", "remarks"},
}
LABELS = {
    "supplier": ["공급업체", "업체명", "상호", "회사명", "supplier", "vendor"],
    "customer": ["수신처", "수신", "고객명", "customer"],
    "quotation_date": ["견적일자", "견적일", "작성일자", "작성일", "발행일", "date"],
    "quotation_number": ["견적번호", "문서번호", "quotation no"],
    "subtotal": ["공급가액", "공급가액합계", "공급가합계", "합계 (VAT 별도)", "소계", "subtotal"],
    "vat": ["부가가치세", "부가세", "세액", "vat"],
    "total": ["최종 견적가 (VAT 포함)", "총견적금액", "합계금액", "총합계", "총금액", "견적금액", "합계", "total"],
}
SUMMARY = {normalize(x) for k in ("subtotal", "vat", "total") for x in LABELS[k]} | {"계", "소계", "총계"}


def text_value(value) -> str:
    if value is None:
        return ""
    if isinstance(value, (date, datetime)):
        return value.strftime("%Y-%m-%d")
    return str(value).strip()


def read_metadata(quote: Quote, rows: list[list], raw_text: str):
    # 셀 경계가 있는 엑셀/PDF 표에서는 인접 셀을 우선 사용한다.
    for row in rows:
        for index, cell in enumerate(row):
            label = normalize(cell)
            for key, aliases in LABELS.items():
                if getattr(quote, key) not in (None, "") or label not in {normalize(a) for a in aliases}:
                    continue
                candidates = [text_value(v) for v in row[index + 1:] if text_value(v)]
                if not candidates:
                    continue
                candidate = candidates[0]
                if key in {"subtotal", "vat", "total"}:
                    for candidate in candidates:
                        try:
                            parsed = number(candidate)
                            if parsed is not None:
                                setattr(quote, key, parsed)
                                break
                        except UserError:
                            pass
                elif not any(normalize(candidate) in values for values in HEADERS.values()):
                    setattr(quote, key, candidate)

    for key, aliases in LABELS.items():
        if getattr(quote, key) not in (None, ""):
            continue
        for alias in sorted(aliases, key=len, reverse=True):
            spaced_alias = r"\s*".join(re.escape(char) for char in alias if not char.isspace())
            pattern = rf"(?:^|\s){spaced_alias}\s*[:：]?\s+([^\n]+)"
            match = re.search(pattern, raw_text, flags=re.I | re.M)
            if not match:
                # '상호:회사명'처럼 공백이 없는 항목
                match = re.search(rf"(?:^|\s){spaced_alias}\s*[:：]\s*([^\n]+)", raw_text, flags=re.I | re.M)
            if not match:
                continue
            candidate = match.group(1).strip() if key in {"subtotal", "vat", "total"} else re.split(r"\s{2,}", match.group(1).strip())[0]
            if key in {"subtotal", "vat", "total"}:
                found = re.match(r"\(?\s*[\\₩￦]?\s*-?\d[\d,]*(?:\.\d+)?\s*\)?\s*원?", candidate)
                if not found:
                    continue
                try:
                    setattr(quote, key, number(found.group()))
                except UserError:
                    continue
            else:
                setattr(quote, key, candidate[:300])
            break
    if not quote.quotation_date:
        match = re.search(r"(20\d{2})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일", raw_text)
        if match:
            try:
                quote.quotation_date = date(*map(int, match.groups())).isoformat()
            except ValueError:
                pass


def items_from_rows(rows: list[list], warnings: list[str]) -> list[Item]:
    items = []
    columns = {}
    for row in rows:
        row = list(row)
        header = {}
        for index, cell in enumerate(row):
            for field, aliases in HEADERS.items():
                if normalize(cell) in aliases:
                    header[field] = index
        if "description" in header and ({"quantity", "unit_price", "amount"} & header.keys()):
            columns = header
            continue
        if not columns:
            continue
        nonempty = [normalize(x) for x in row if text_value(x)]
        if not nonempty:
            continue
        if nonempty[0] in SUMMARY:
            columns = {}
            continue
        values = {key: text_value(row[idx]) if idx < len(row) else "" for key, idx in columns.items()}
        if normalize(values.get("description", "")) in SUMMARY:
            columns = {}
            continue
        if not values.get("description"):
            if any(values.get(k) for k in ("quantity", "unit_price", "amount")):
                warnings.append("품명이 없는 숫자 행이 있습니다. 원문에서 누락 품목을 확인해 주세요.")
            continue
        item = Item()
        has_number = False
        for key, value in values.items():
            if key in {"quantity", "unit_price", "amount"}:
                try:
                    parsed = number(value)
                    setattr(item, key, parsed)
                    has_number = has_number or parsed is not None
                except UserError:
                    warnings.append(f"'{values['description']}': {key} 값을 읽지 못했습니다 ({value}).")
            else:
                setattr(item, key, value)
        if not has_number:
            continue
        if item.amount is None and item.quantity is not None and item.unit_price is not None:
            item.amount = item.quantity * item.unit_price
            warnings.append(f"'{item.description}': 금액을 수량 × 단가로 계산했습니다.")
        items.append(item)
    return items


def tesseract_path() -> str | None:
    detected = shutil.which("tesseract")
    if detected:
        return detected
    path = Path("C:/Program Files/Tesseract-OCR/tesseract.exe")
    return str(path) if path.is_file() else None


def ocr_page(page) -> tuple[str, list[list[str]]]:
    executable = tesseract_path()
    if not executable:
        raise UserError("스캔 PDF의 OCR에는 Tesseract와 한국어(kor) 언어 파일이 필요합니다. OCR을 설치하거나 원문을 보며 직접 입력해 주세요.")
    with tempfile.TemporaryDirectory(prefix="quote-ocr-") as folder:
        image = Path(folder) / "page.png"
        page.to_image(resolution=200).save(str(image))
        try:
            result = subprocess.run(
                [executable, str(image), "stdout", "-l", "kor+eng", "--psm", "6", "-c", "preserve_interword_spaces=1"],
                capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=120,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except subprocess.TimeoutExpired as exc:
            raise UserError("OCR 처리 시간이 초과되었습니다. 해당 페이지를 직접 확인해 주세요.") from exc
        if result.returncode:
            raise UserError("OCR에 실패했습니다. Tesseract의 kor/eng 언어 파일을 확인해 주세요. " + result.stderr[-300:])
    text = result.stdout
    return text, [re.split(r"\s{2,}|\t", line.strip()) for line in text.splitlines() if line.strip()]


def extract_quote(path: str, use_ocr: bool = False) -> Quote:
    path_obj = Path(path)
    quote = Quote(source=str(path_obj.resolve()), title=f"{path_obj.stem} 구매 품의")
    tables = []
    texts = []
    suffix = path_obj.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        workbook = load_workbook(path_obj, read_only=True, data_only=True)
        try:
            for sheet in workbook:
                if sheet.max_row > 10000 or sheet.max_column > 200:
                    raise UserError("견적서의 사용 범위가 너무 큽니다. 10,000행 / 200열 이내의 견적 영역만 별도 파일로 저장해 주세요.")
                rows = [list(row) for row in sheet.iter_rows(values_only=True)]
                tables.append(rows)
                texts.append(f"[{sheet.title}]\n" + "\n".join("\t".join(text_value(c) for c in row) for row in rows))
        finally:
            workbook.close()
        quote.warnings.append("엑셀 수식은 파일에 저장된 계산 결과를 읽습니다. 원본을 Excel에서 계산·저장한 뒤 가져와 주세요.")
    elif suffix == ".xls":
        try:
            import xlrd
        except ImportError as exc:
            raise UserError(".xls 엑셀 파일을 읽는 구성 요소(xlrd)가 설치되지 않았습니다. 프로그램을 종료한 뒤 같은 폴더의 install.bat을 실행하고 start.bat으로 다시 시작해 주세요.") from exc
        workbook = xlrd.open_workbook(path)
        try:
            for sheet in workbook.sheets():
                if sheet.nrows > 10000 or sheet.ncols > 200:
                    raise UserError("견적 영역만 별도 .xlsx 파일로 저장해 주세요.")
                rows = []
                for row_idx in range(sheet.nrows):
                    row = []
                    for cell in sheet.row(row_idx):
                        row.append(xlrd.xldate_as_datetime(cell.value, workbook.datemode) if cell.ctype == xlrd.XL_CELL_DATE else cell.value)
                    rows.append(row)
                tables.append(rows)
                texts.append("\n".join("\t".join(text_value(c) for c in row) for row in rows))
        finally:
            workbook.release_resources()
    elif suffix == ".pdf":
        with pdfplumber.open(path) as pdf:
            for index, page in enumerate(pdf.pages, 1):
                text = page.extract_text(layout=True) or ""
                found_tables = page.extract_tables()
                if not text.strip():
                    if use_ocr:
                        try:
                            text, ocr_rows = ocr_page(page)
                            found_tables = [ocr_rows]
                            quote.warnings.append(f"{index}쪽은 OCR로 읽었습니다. 문자와 숫자를 원본과 대조해 주세요.")
                        except UserError as exc:
                            quote.warnings.append(f"{index}쪽: {exc}")
                    else:
                        quote.warnings.append(f"{index}쪽에 읽을 수 있는 텍스트가 없습니다. OCR을 사용하거나 직접 입력해 주세요.")
                texts.append(f"[{index}쪽]\n{text}")
                if found_tables and any(items_from_rows(table, []) for table in found_tables):
                    tables.extend(found_tables)
                else:
                    tables.append([re.split(r"\s{2,}|\t", line.strip()) for line in text.splitlines() if line.strip()])
    else:
        raise UserError("견적서는 PDF, XLSX, XLSM, XLS 형식을 지원합니다.")

    quote.raw_text = "\n\n".join(texts)
    for table in tables:
        quote.items.extend(items_from_rows(table, quote.warnings))
        if any(normalize(cell) in {"공급대가vat포함", "금액vat포함"} for row in table for cell in row):
            quote.amount_basis = "inclusive"
    if quote.amount_basis == "inclusive":
        quote.warnings.append("품목 금액은 VAT 포함입니다. 품목 합계를 부가세 포함 합계금액과 대조합니다.")
    read_metadata(quote, [row for table in tables for row in table], quote.raw_text)
    if not quote.items:
        quote.warnings.append("품목 표를 자동 인식하지 못했습니다. 원문을 확인하여 품목을 추가해 주세요.")
    if quote.items and all(i.amount is not None for i in quote.items):
        if quote.amount_basis == "inclusive" and quote.total is None:
            quote.total = sum(i.amount for i in quote.items)
            quote.warnings.append("부가세 포함 합계금액을 품목 금액의 합으로 계산했습니다.")
        elif quote.amount_basis == "exclusive" and quote.subtotal is None:
            quote.subtotal = sum(i.amount for i in quote.items)
            quote.warnings.append("공급가액을 품목 금액의 합으로 계산했습니다. 부가세 포함 여부를 확인해 주세요.")
    if quote.total is None and quote.subtotal is not None and quote.vat is not None:
        quote.total = quote.subtotal + quote.vat
        quote.warnings.append("합계금액을 공급가액 + 부가세로 계산했습니다.")
    if suffix == ".pdf" and any("(cid:" in item.description for item in quote.items):
        apply_verified_corrections(quote, fingerprint(path_obj))
    quote.warnings = list(dict.fromkeys(quote.warnings))
    return quote
