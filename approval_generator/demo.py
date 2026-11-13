"""개인정보가 없는 체험용 파일을 생성한다. python -m approval_generator.demo"""
from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree
from reportlab.pdfgen import canvas

from .models import Mapping, Profile
from .render import S, korean_font, serialize


def sample_xlsx(path: Path, cells: dict, merges=(), formulas=None):
    formulas = formulas or {}
    ns = {None: S}
    root = etree.Element(f"{{{S}}}worksheet", nsmap=ns)
    etree.SubElement(root, f"{{{S}}}dimension", ref="A1:G20")
    views = etree.SubElement(root, f"{{{S}}}sheetViews")
    etree.SubElement(views, f"{{{S}}}sheetView", workbookViewId="0", showGridLines="0")
    columns = etree.SubElement(root, f"{{{S}}}cols")
    for idx, width in enumerate([25, 20, 10, 12, 17, 19, 20], 1):
        etree.SubElement(columns, f"{{{S}}}col", min=str(idx), max=str(idx), width=str(width), customWidth="1")
    data = etree.SubElement(root, f"{{{S}}}sheetData")
    for row_idx in range(1, 21):
        row = etree.SubElement(data, f"{{{S}}}row", r=str(row_idx), ht="28", customHeight="1")
        for col in "ABCDEFG":
            address = f"{col}{row_idx}"
            cell = etree.SubElement(row, f"{{{S}}}c", r=address, s="1" if col in "DEF" else "0")
            if address in formulas:
                etree.SubElement(cell, f"{{{S}}}f").text = formulas[address]
                etree.SubElement(cell, f"{{{S}}}v").text = "0"
            elif address in cells:
                value = cells[address]
                if isinstance(value, (int, float)):
                    cell.set("t", "n")
                    etree.SubElement(cell, f"{{{S}}}v").text = str(value)
                else:
                    cell.set("t", "inlineStr")
                    etree.SubElement(etree.SubElement(cell, f"{{{S}}}is"), f"{{{S}}}t").text = str(value)
    if merges:
        merge_cells = etree.SubElement(root, f"{{{S}}}mergeCells", count=str(len(merges)))
        for reference in merges:
            etree.SubElement(merge_cells, f"{{{S}}}mergeCell", ref=reference)
    etree.SubElement(root, f"{{{S}}}pageMargins", left="0.3", right="0.3", top="0.4", bottom="0.4", header="0.2", footer="0.2")
    etree.SubElement(root, f"{{{S}}}pageSetup", paperSize="9", orientation="portrait", fitToWidth="1", fitToHeight="1")
    files = {
        "[Content_Types].xml": '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
        "_rels/.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": f'<workbook xmlns="{S}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="양식" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="124519" fullCalcOnLoad="1"/></workbook>',
        "xl/_rels/workbook.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
        "xl/styles.xml": f'<styleSheet xmlns="{S}"><fonts count="1"><font><sz val="11"/><name val="맑은 고딕"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>',
        "xl/worksheets/sheet1.xml": serialize(root),
    }
    with ZipFile(path, "w", compression=ZIP_DEFLATED) as archive:
        for name, value in files.items():
            archive.writestr(name, value.encode("utf-8") if isinstance(value, str) else value)


def create_demo(folder: str | Path) -> dict[str, Path]:
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    source_cells = {"A1": "견적서 (체험용 가상 데이터)", "A2": "공급업체", "B2": "샘플상사", "D2": "견적일", "E2": "2026-11-13",
                    "A5": "품명", "B5": "규격", "C5": "단위", "D5": "수량", "E5": "단가", "F5": "금액", "G5": "비고",
                    "A6": "무선 키보드", "B6": "표준형", "C6": "개", "D6": 2, "E6": 50000, "F6": 100000,
                    "A7": "모니터", "B7": "24인치", "C7": "대", "D7": 1, "E7": 150000, "F7": 150000,
                    "A9": "공급가액", "F9": 250000, "A10": "부가세", "F10": 25000, "A11": "합계금액", "F11": 275000}
    template_cells = {"A1": "구 매 품 의 서", "A3": "공급업체", "D3": "견적일", "A4": "품의 제목",
                      "A6": "품명", "B6": "규격", "C6": "단위", "D6": "수량", "E6": "단가", "F6": "금액", "G6": "비고",
                      "A13": "품목 합계", "A15": "합계금액", "A17": "품의 사유", "A19": "단가와 금액은 부가세 포함입니다."}
    source_xlsx = folder / "체험_견적서.xlsx"
    template_xlsx = folder / "체험_품의서양식.xlsx"
    sample_xlsx(source_xlsx, source_cells, ["A1:G1"])
    sample_xlsx(template_xlsx, template_cells, ["A1:G1", "B3:C3", "E3:G3", "B4:G4", "B17:G18", "A19:G19"])
    profile = Profile.create(str(template_xlsx))
    for field, cell in {"supplier": "B3", "quotation_date": "E3", "title": "B4", "subtotal": "F13", "vat": "F14", "total": "F15"}.items():
        profile.mappings.append(Mapping(field=field, sheet="양식", cell=cell))
    for field, col in zip(["description", "specification", "unit", "quantity", "unit_price", "amount", "note"], "ABCDEFG"):
        profile.mappings.append(Mapping(field="items." + field, sheet="양식", cell=col + "7", count=5))
    excel_profile = folder / "체험_엑셀연결.json"
    profile.save(str(excel_profile))

    font = korean_font()
    source_pdf = folder / "체험_견적서.pdf"
    drawing = canvas.Canvas(str(source_pdf), pagesize=(595, 842))
    drawing.setFont(font, 22)
    drawing.drawString(40, 790, "견적서 (체험용 가상 데이터)")
    drawing.setFont(font, 11)
    drawing.drawString(40, 750, "공급업체: 샘플상사")
    drawing.drawString(330, 750, "견적일: 2026-11-13")
    xs = [40, 200, 280, 325, 390, 475, 555]
    rows = [["품명", "규격", "단위", "수량", "단가", "금액"],
            ["무선 키보드", "표준형", "개", "2", "50,000", "100,000"],
            ["모니터", "24인치", "대", "1", "150,000", "150,000"]]
    for y in [710, 680, 650, 620]:
        drawing.line(xs[0], y, xs[-1], y)
    for x in xs:
        drawing.line(x, 710, x, 620)
    for row_index, row in enumerate(rows):
        for x, value in zip(xs, row):
            drawing.drawString(x + 4, 690 - row_index * 30, value)
    for offset, text in enumerate(["공급가액: 250,000", "부가세: 25,000", "합계금액: 275,000"]):
        drawing.drawString(340, 580 - 25 * offset, text)
    drawing.save()

    template_pdf = folder / "체험_품의서양식.pdf"
    drawing = canvas.Canvas(str(template_pdf), pagesize=(595, 842))
    drawing.setFont(font, 24)
    drawing.drawCentredString(297.5, 785, "구 매 품 의 서")
    drawing.setFont(font, 11)
    drawing.drawString(40, 740, "공급업체")
    drawing.drawString(330, 740, "견적일")
    drawing.drawString(40, 705, "품의 제목")
    for y in range(660, 479, -30):
        drawing.line(40, y, 555, y)
    for x in xs:
        drawing.line(x, 660, x, 480)
    for x, label in zip(xs, rows[0]):
        drawing.drawString(x + 4, 640, label)
    for y, label in [(440, "품목 합계"), (380, "합계금액")]:
        drawing.drawString(345, y, label)
        drawing.line(340, y - 8, 555, y - 8)
    drawing.setFont(font, 9)
    drawing.drawString(40, 465, "단가와 금액은 부가세 포함입니다.")
    drawing.drawString(40, 60, "체험용 빈 양식입니다. 실제 사용 시 회사의 품의서 양식을 등록해 주세요.")
    drawing.save()
    profile = Profile.create(str(template_pdf))
    for field, x, y, width in [("supplier", 108, 88, 210), ("quotation_date", 385, 88, 160), ("title", 108, 123, 437),
                               ("subtotal", 430, 390, 115), ("vat", 430, 420, 115), ("total", 430, 450, 115)]:
        profile.mappings.append(Mapping(field=field, x=x, y=y, width=width, height=24, align="right" if field in {"subtotal", "vat", "total"} else "left"))
    for field, x, width in [("description", 42, 156), ("specification", 202, 76), ("unit", 282, 41), ("quantity", 327, 61), ("unit_price", 392, 81), ("amount", 477, 76)]:
        profile.mappings.append(Mapping(field="items." + field, x=x, y=214, width=width, height=24, count=5, step=30,
                                        align="right" if field in {"quantity", "unit_price", "amount"} else "left"))
    pdf_profile = folder / "체험_PDF연결.json"
    profile.save(str(pdf_profile))
    return {"source_xlsx": source_xlsx, "source_pdf": source_pdf, "template_xlsx": template_xlsx,
            "template_pdf": template_pdf, "excel_profile": excel_profile, "pdf_profile": pdf_profile}


if __name__ == "__main__":
    paths = create_demo(Path.cwd() / "examples")
    print("체험 파일 생성 완료:", paths["source_xlsx"].parent)
