from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from .extract import HEADERS, normalize
from .models import Mapping, Profile, UserError


def profile_for_template(path, profiles_folder):
    """선택한 양식과 내용이 같은 저장 설정만 복원한다. 충돌하면 사용자가 선택한다."""
    candidate = Profile.create(path)
    matches = []
    for saved_path in sorted(Path(profiles_folder).glob("*.json")):
        try:
            saved = Profile.load(str(saved_path))
        except (OSError, ValueError, TypeError, AttributeError):
            continue
        if saved.template_hash != candidate.template_hash or not saved.mappings:
            continue
        if not any(saved.mappings == other.mappings and saved.default_unit == other.default_unit for other in matches):
            matches.append(saved)
    if len(matches) == 1:
        restored = matches[0]
        restored.template = candidate.template
        return restored, f"저장된 연결 설정 {len(restored.mappings)}개를 자동으로 불러왔습니다. 연결 내용을 확인한 뒤 '3 일괄 생성'으로 이동하세요."
    if matches:
        return candidate, "이 양식의 연결 설정이 여러 개입니다. '연결 설정 불러오기'에서 사용할 설정을 선택해 주세요."
    return candidate, "양식 파일을 선택했습니다. '열 제목으로 자동 연결' 또는 '연결 설정 불러오기'로 입력 위치를 연결해 주세요."


def sheet_names(path):
    if Path(path).suffix.lower() == ".xls":
        from .legacy_excel import read_xls
        book = read_xls(path)
        try:
            return book.sheet_names()
        finally:
            book.release_resources()
    book = load_workbook(path, read_only=True)
    try:
        return book.sheetnames
    finally:
        book.close()


def preview_grid(path, sheet_name):
    if Path(path).suffix.lower() == ".xls":
        from .legacy_excel import read_xls
        book = read_xls(path)
        try:
            sheet = book.sheet_by_name(sheet_name)
            columns = min(max(sheet.ncols, 12), 60)
            return [list(sheet.row_values(r, end_colx=min(sheet.ncols, columns))) + [None] * max(0, columns - sheet.ncols)
                    if r < sheet.nrows else [None] * columns for r in range(min(max(sheet.nrows, 50), 1000))]
        finally:
            book.release_resources()
    book = load_workbook(path, read_only=True, data_only=False)
    try:
        sheet = book[sheet_name]
        return [list(row) for row in sheet.iter_rows(max_row=min(max(sheet.max_row, 50), 1000), max_col=min(max(sheet.max_column, 12), 60), values_only=True)]
    finally:
        book.close()


def table_mappings(path, sheet_name):
    if Path(path).suffix.lower() == ".pdf":
        raise UserError("열 제목 자동 연결은 첫 행에 열 제목이 있는 엑셀 표 양식에 사용합니다.")
    grid = preview_grid(path, sheet_name)
    mapping = []
    for index, cell in enumerate(grid[0], 1):
        for key, aliases in HEADERS.items():
            if normalize(cell) in aliases:
                mapping.append(Mapping(field="items." + key, sheet=sheet_name, cell=get_column_letter(index) + "2", count=1000))
    fields = {m.field for m in mapping}
    if "items.description" not in fields or not fields & {"items.quantity", "items.unit_price", "items.amount"}:
        raise UserError("첫 행에서 품목 열 제목을 찾지 못했습니다. 입력 위치를 직접 연결해 주세요.")
    if any(value not in (None, "") for row in grid[1:] for value in row):
        raise UserError("자동 연결은 첫 행에만 열 제목이 있는 빈 표 양식에 사용합니다. 기존 내용이 있는 양식은 직접 연결해 주세요.")
    return mapping
