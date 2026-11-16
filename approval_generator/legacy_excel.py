from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from decimal import Decimal
from pathlib import Path

from .models import UserError


def read_xls(path):
    try:
        import xlrd
    except ImportError as exc:
        raise UserError(".xls 엑셀 파일을 읽는 구성 요소(xlrd)가 설치되지 않았습니다. 프로그램을 종료한 뒤 같은 폴더의 install.bat을 실행하고 start.bat으로 다시 시작해 주세요.") from exc
    return xlrd.open_workbook(str(path), formatting_info=True)


def render_xls(profile, quote):
    if sys.platform != "win32":
        raise UserError("XLS 양식 출력에는 Windows와 Microsoft Excel이 필요합니다.")
    from .render import cell_address, excel_updates
    updates = excel_updates(profile, quote)
    book = read_xls(profile.template)
    try:
        for name, cells in updates.items():
            if name not in book.sheet_names():
                raise UserError(f"시트를 찾을 수 없습니다: {name}")
            sheet = book.sheet_by_name(name)
            for address, value in list(cells.items()):
                col, row = cell_address(address)
                if row > 65536 or col > 256:
                    raise UserError(f"XLS의 행/열 제한을 벗어났습니다: {address}")
                # 내용이 원래 없는 양식의 빈 행 수천 개를 불필요하게 만들지 않는다.
                if value in (None, "") and (row > sheet.nrows or col > sheet.ncols):
                    del cells[address]
                elif isinstance(value, Decimal):
                    cells[address] = int(value) if value == value.to_integral() else float(value)
    finally:
        book.release_resources()
    with tempfile.TemporaryDirectory(prefix="approval-xls-") as folder:
        data = Path(folder) / "cells.json"
        result = Path(folder) / "result.xls"
        data.write_text(json.dumps(updates, ensure_ascii=False), encoding="utf-8")
        command = ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(Path(__file__).with_name("native_excel.ps1")),
                   "-InputPath", str(Path(profile.template).resolve()), "-OutputPath", str(result), "-DataPath", str(data)]
        try:
            completed = subprocess.run(command, capture_output=True, encoding="utf-8", errors="replace",
                                       timeout=180, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        except subprocess.TimeoutExpired as exc:
            raise UserError("Excel의 XLS 저장 시간이 초과되었습니다. Excel 대화상자 또는 설치 상태를 확인해 주세요.") from exc
        if completed.returncode or not result.is_file():
            raise UserError("XLS 저장에 실패했습니다. 이 PC에 Microsoft Excel이 설치되어 있어야 합니다.\n" + (completed.stdout + completed.stderr)[-1500:])
        return result.read_bytes()
