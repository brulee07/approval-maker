from __future__ import annotations

import io
import math
import os
import posixpath
import re
from collections import defaultdict
from decimal import Decimal
from pathlib import Path
from zipfile import ZipFile

from lxml import etree
from openpyxl.utils.cell import column_index_from_string, get_column_letter, range_boundaries
from pypdf import PdfReader, PdfWriter, Transformation
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

from .models import ALL_FIELDS, Mapping, Profile, Quote, UserError, display
from .pricing import prepare_approval_quote


S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML = "http://www.w3.org/XML/1998/namespace"
NS = {"s": S}


def xml(data):
    return etree.fromstring(data, parser=etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False))


def serialize(element):
    return etree.tostring(element, encoding="UTF-8", xml_declaration=True, standalone=True)


def worksheet_paths(archive: ZipFile) -> dict[str, str]:
    workbook = xml(archive.read("xl/workbook.xml"))
    if workbook.tag != f"{{{S}}}workbook":
        raise UserError("이 엑셀 XML 형식은 지원하지 않습니다. Excel에서 일반 .xlsx로 다시 저장해 주세요.")
    rels = xml(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {rel.get("Id"): rel.get("Target") for rel in rels if rel.get("TargetMode") != "External"}
    result = {}
    for sheet in workbook.findall("s:sheets/s:sheet", NS):
        target = targets.get(sheet.get(f"{{{R}}}id"), "")
        location = posixpath.normpath(target.lstrip("/") if target.startswith("/") else "xl/" + target)
        if location.startswith("xl/worksheets/"):
            result[sheet.get("name")] = location
    return result


def cell_address(address: str) -> tuple[int, int]:
    match = re.fullmatch(r"([A-Za-z]{1,3})([1-9]\d{0,6})", address.strip())
    if not match:
        raise UserError(f"셀 주소가 올바르지 않습니다: {address} (예: C12)")
    col, row = column_index_from_string(match[1].upper()), int(match[2])
    if col > 16384 or row > 1048576:
        raise UserError(f"엑셀 범위를 벗어난 셀입니다: {address}")
    return col, row


def validate_mappings(profile: Profile, quote: Quote):
    if not profile.mappings:
        raise UserError("양식의 입력 위치를 하나 이상 연결해 주세요.")
    for mapping in profile.mappings:
        if mapping.field not in ALL_FIELDS:
            raise UserError(f"알 수 없는 연결 항목입니다: {mapping.field}")
        if not isinstance(mapping.count, int) or not 1 <= mapping.count <= 2000:
            raise UserError("품목 칸 수는 1~2,000 사이 정수여야 합니다.")
        if not math.isfinite(mapping.step) or mapping.step <= 0:
            raise UserError("반복 간격은 0보다 커야 합니다.")
        if mapping.field.startswith("items.") and len(quote.items) > mapping.count:
            raise UserError(f"{ALL_FIELDS[mapping.field]}: 양식은 {mapping.count}개까지 입력할 수 있지만 품목은 {len(quote.items)}개입니다. 양식/연결 설정을 늘려 주세요.")
        if mapping.align not in {"left", "center", "right"}:
            raise UserError("정렬은 left, center, right 중 하나여야 합니다.")


def excel_updates(profile: Profile, quote: Quote) -> dict[str, dict[str, object]]:
    updates = defaultdict(dict)
    for mapping in profile.mappings:
        col, row = cell_address(mapping.cell)
        repeated = mapping.field.startswith("items.")
        if repeated and int(mapping.step) != mapping.step:
            raise UserError("엑셀 행 간격은 정수여야 합니다.")
        for index in range(mapping.count if repeated else 1):
            address = f"{get_column_letter(col)}{row + index * int(mapping.step)}"
            cell_address(address)
            if address in updates[mapping.sheet]:
                raise UserError(f"두 연결이 같은 셀을 사용합니다: {mapping.sheet}!{address}")
            updates[mapping.sheet][address] = quote.value(mapping.field, index)
    return dict(updates)


def update_sheet(root, updates: dict[str, object], sheet: str):
    data = root.find("s:sheetData", NS)
    if data is None:
        raise UserError(f"입력할 수 없는 시트입니다: {sheet}")
    merged = [range_boundaries(m.get("ref")) for m in root.findall("s:mergeCells/s:mergeCell", NS)]
    rows = {int(row.get("r")): row for row in data}
    for address, value in updates.items():
        col, row_number = cell_address(address)
        for left, top, right, bottom in merged:
            if left <= col <= right and top <= row_number <= bottom and (col, row_number) != (left, top):
                raise UserError(f"{sheet}!{address}는 병합 셀의 내부입니다. 왼쪽 위 셀 {get_column_letter(left)}{top}에 연결해 주세요.")
        row = rows.get(row_number)
        if row is None:
            row = etree.Element(f"{{{S}}}row", r=str(row_number))
            following = next((r for n, r in sorted(rows.items()) if n > row_number), None)
            if following is not None:
                data.insert(data.index(following), row)
            else:
                data.append(row)
            rows[row_number] = row
        cell = next((c for c in row if c.get("r") == address), None)
        if cell is None:
            cell = etree.Element(f"{{{S}}}c", r=address)
            following = next((c for c in row if c.get("r") and cell_address(c.get("r"))[0] > col), None)
            if following is not None:
                row.insert(row.index(following), cell)
            else:
                row.append(cell)
        if cell.find("s:f", NS) is not None:
            raise UserError(f"{sheet}!{address}에는 수식이 있습니다. 수식을 유지하려면 해당 셀의 연결을 삭제하고 수식의 입력 셀에 연결해 주세요.")
        for child in list(cell):
            if child.tag in {f"{{{S}}}v", f"{{{S}}}is"}:
                cell.remove(child)
        cell.attrib.pop("t", None)
        if value is None or value == "":
            continue
        if isinstance(value, (Decimal, int, float)):
            cell.set("t", "n")
            etree.SubElement(cell, f"{{{S}}}v").text = str(value)
        else:
            if len(str(value)) > 32767:
                raise UserError(f"{sheet}!{address}: 엑셀 셀의 글자 수 제한을 초과했습니다.")
            cell.set("t", "inlineStr")
            text = etree.SubElement(etree.SubElement(cell, f"{{{S}}}is"), f"{{{S}}}t")
            text.set(f"{{{XML}}}space", "preserve")
            text.text = str(value)  # '='로 시작하는 입력도 수식이 아닌 문자열로 저장
    dimension = root.find("s:dimension", NS)
    if dimension is not None:
        old = range_boundaries(dimension.get("ref", "A1"))
        coordinates = [cell_address(c) for c in updates]
        left = min([old[0]] + [c for c, _ in coordinates])
        top = min([old[1]] + [r for _, r in coordinates])
        right = max([old[2]] + [c for c, _ in coordinates])
        bottom = max([old[3]] + [r for _, r in coordinates])
        dimension.set("ref", f"{get_column_letter(left)}{top}:{get_column_letter(right)}{bottom}")


def render_xlsx(profile: Profile, quote: Quote) -> bytes:
    changes = excel_updates(profile, quote)
    output = io.BytesIO()
    with ZipFile(profile.template) as archive:
        if any(name.startswith("_xmlsignatures/") for name in archive.namelist()):
            raise UserError("전자서명된 엑셀 양식은 지원하지 않습니다. 서명되지 않은 빈 양식을 사용해 주세요.")
        paths = worksheet_paths(archive)
        replacements = {}
        for sheet, updates in changes.items():
            if sheet not in paths:
                raise UserError(f"양식에서 시트를 찾을 수 없습니다: {sheet}")
            root = xml(archive.read(paths[sheet]))
            update_sheet(root, updates, sheet)
            replacements[paths[sheet]] = serialize(root)
        workbook = xml(archive.read("xl/workbook.xml"))
        calc = workbook.find("s:calcPr", NS)
        if calc is None:
            calc = etree.Element(f"{{{S}}}calcPr")
            # calcPr은 extLst 등 마지막 확장 요소보다 앞에 있어야 한다.
            trailing = {"oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"}
            following = next((child for child in workbook if etree.QName(child).localname in trailing), None)
            if following is not None:
                workbook.insert(workbook.index(following), calc)
            else:
                workbook.append(calc)
        calc.set("calcMode", "auto")
        calc.set("fullCalcOnLoad", "1")
        calc.set("forceFullCalc", "1")
        replacements["xl/workbook.xml"] = serialize(workbook)
        with ZipFile(output, "w") as destination:
            for info in archive.infolist():
                destination.writestr(info, replacements.get(info.filename, archive.read(info.filename)))
    return output.getvalue()


def korean_font() -> str:
    name = "ApprovalKorean"
    if name in pdfmetrics.getRegisteredFontNames():
        return name
    candidates = [os.environ.get("APPROVAL_FONT", ""), "C:/Windows/Fonts/malgun.ttf", "/usr/share/fonts/truetype/nanum/NanumGothic.ttf"]
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            pdfmetrics.registerFont(TTFont(name, candidate))
            return name
    raise UserError("한글 PDF 출력용 글꼴이 없습니다. 맑은 고딕 또는 나눔고딕 TTF 경로를 APPROVAL_FONT 환경변수로 지정해 주세요.")


def wrap_text(text: str, width: float, font: str, size: float) -> list[str]:
    lines = []
    for paragraph in text.split("\n"):
        current = ""
        for char in paragraph:
            if pdfmetrics.stringWidth(char, font, size) > width:
                raise UserError("PDF 입력 영역의 폭이 글자보다 좁습니다. 영역을 넓혀 주세요.")
            if current and pdfmetrics.stringWidth(current + char, font, size) > width:
                lines.append(current)
                current = char
            else:
                current += char
        lines.append(current)
    return lines


def pdf_size(path: str, page_number: int) -> tuple[float, float]:
    reader = PdfReader(path)
    if reader.is_encrypted:
        raise UserError("암호화된 PDF는 지원하지 않습니다. 암호를 해제한 사본을 사용해 주세요.")
    if not 1 <= page_number <= len(reader.pages):
        raise UserError("PDF 페이지 번호를 확인해 주세요.")
    page = reader.pages[page_number - 1]
    if page.rotation:
        page.transfer_rotation_to_content()
    return float(page.cropbox.width), float(page.cropbox.height)


def render_pdf(profile: Profile, quote: Quote) -> bytes:
    reader = PdfReader(profile.template)
    if reader.is_encrypted:
        raise UserError("암호화된 PDF 양식은 지원하지 않습니다.")
    if any(field.get("/FT") == "/Sig" for field in (reader.get_fields() or {}).values()):
        raise UserError("전자서명 필드가 있는 PDF는 지원하지 않습니다. 서명되지 않은 빈 양식을 사용해 주세요.")
    writer = PdfWriter()
    writer.clone_document_from_reader(reader)
    font = korean_font()
    by_page = defaultdict(list)
    for mapping in profile.mappings:
        if not isinstance(mapping.page, int) or not 1 <= mapping.page <= len(writer.pages):
            raise UserError(f"PDF 페이지 번호를 확인해 주세요: {mapping.page}")
        if not all(math.isfinite(v) for v in (mapping.x, mapping.y, mapping.width, mapping.height, mapping.font_size)):
            raise UserError("PDF 입력 위치에는 유한한 숫자를 입력해 주세요.")
        if mapping.width <= 4 or mapping.height <= 4 or not 6 <= mapping.font_size <= 72:
            raise UserError("PDF 영역은 4pt보다 커야 하고 글자 크기는 6~72pt여야 합니다.")
        by_page[mapping.page - 1].append(mapping)
    for page_index, mappings in by_page.items():
        page = writer.pages[page_index]
        if page.rotation:
            page.transfer_rotation_to_content()
        width, height = float(page.cropbox.width), float(page.cropbox.height)
        buffer = io.BytesIO()
        overlay = canvas.Canvas(buffer, pagesize=(width, height))
        used = []
        for mapping in mappings:
            repeated = mapping.field.startswith("items.")
            for index in range(mapping.count if repeated else 1):
                x, y = mapping.x, mapping.y + (index * mapping.step if repeated else 0)
                rectangle = (x, y, x + mapping.width, y + mapping.height)
                if x < 0 or y < 0 or rectangle[2] > width + .01 or rectangle[3] > height + .01:
                    raise UserError(f"{ALL_FIELDS[mapping.field]}: PDF {page_index + 1}쪽의 범위를 벗어납니다.")
                if any(max(x, a) < min(rectangle[2], c) and max(y, b) < min(rectangle[3], d) for a, b, c, d in used):
                    raise UserError(f"{ALL_FIELDS[mapping.field]}: PDF 입력 영역이 다른 연결 영역과 겹칩니다.")
                used.append(rectangle)
                value = display(quote.value(mapping.field, index))
                if not value:
                    continue
                lines = wrap_text(value, mapping.width - 4, font, mapping.font_size)
                leading = mapping.font_size * 1.35
                if len(lines) * leading > mapping.height - 4:
                    raise UserError(f"{ALL_FIELDS[mapping.field]}: '{value[:35]}' 내용이 PDF 영역에 다 들어가지 않습니다. 높이/폭을 늘리거나 글자 크기를 조정해 주세요.")
                overlay.setFont(font, mapping.font_size)
                for line_index, line in enumerate(lines):
                    baseline = height - y - 2 - mapping.font_size - line_index * leading
                    if mapping.align == "right":
                        overlay.drawRightString(x + mapping.width - 2, baseline, line)
                    elif mapping.align == "center":
                        overlay.drawCentredString(x + mapping.width / 2, baseline, line)
                    else:
                        overlay.drawString(x + 2, baseline, line)
        overlay.save()
        stamp = PdfReader(buffer).pages[0]
        page.merge_transformed_page(stamp, Transformation().translate(float(page.cropbox.left), float(page.cropbox.bottom)))
    result = io.BytesIO()
    writer.write(result)
    return result.getvalue()


def render(profile: Profile, quote: Quote) -> bytes:
    profile.check_template()
    if not quote.reviewed:
        raise UserError("견적서 내용을 검토하고 '검토 완료'를 눌러 주세요.")
    quote = prepare_approval_quote(quote, profile.default_unit)
    validate_mappings(profile, quote)
    missing = sorted({ALL_FIELDS[m.field] for m in profile.mappings if m.field != "vat" and not m.field.startswith("items.") and quote.value(m.field) in (None, "")})
    if missing:
        raise UserError("연결한 항목이 비어 있습니다: " + ", ".join(missing))
    if any(m.field == "items.unit_price" for m in profile.mappings):
        missing_price = [str(i) for i, item in enumerate(quote.items, 1) if item.unit_price is None]
        if missing_price:
            raise UserError("부가세 포함 단가를 출력하려면 수량과 금액을 확인해 주세요. 단가가 없는 출력 행: " + ", ".join(missing_price))
        if any(m.field == "items.quantity" for m in profile.mappings):
            missing_quantity = [str(i) for i, item in enumerate(quote.items, 1) if item.quantity is None]
            if missing_quantity:
                raise UserError("수량 × 단가로 품의서 총액을 확인하려면 수량이 필요합니다. 수량이 없는 출력 행: " + ", ".join(missing_quantity))
    suffix = Path(profile.template).suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        return render_xlsx(profile, quote)
    if suffix == ".pdf":
        return render_pdf(profile, quote)
    if suffix == ".xls":
        from .legacy_excel import render_xls
        return render_xls(profile, quote)
    raise UserError("품의서 양식은 XLS, XLSX, XLSM, PDF를 지원합니다.")


def save_result(profile: Profile, quote: Quote, folder: str) -> Path:
    if not folder or not str(folder).strip():
        raise UserError("결과 저장 폴더를 직접 지정해 주세요.")
    data = render(profile, quote)
    destination = Path(folder).resolve()
    destination.mkdir(parents=True, exist_ok=True)
    stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", Path(quote.source).stem).rstrip(". ")[:90] or "견적서"
    suffix = Path(profile.template).suffix.lower()
    for index in range(1, 10000):
        name = f"품의서_{stem}{'_' + str(index) if index > 1 else ''}{suffix}"
        path = destination / name
        try:
            with path.open("xb") as stream:
                try:
                    stream.write(data)
                except BaseException:
                    stream.close()
                    path.unlink(missing_ok=True)
                    raise
            return path
        except FileExistsError:
            continue
    raise UserError("같은 이름의 결과물이 너무 많습니다. 다른 출력 폴더를 선택해 주세요.")
