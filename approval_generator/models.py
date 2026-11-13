from __future__ import annotations

import hashlib
import json
import re
from dataclasses import asdict, dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path


FIELDS = {
    "title": "품의 제목", "supplier": "공급업체", "customer": "견적 수신자",
    "quotation_date": "견적일", "quotation_number": "견적번호",
    "requester": "기안자", "department": "부서", "purpose": "품의 사유",
    "notes": "비고", "subtotal": "공급가액", "vat": "부가세", "total": "합계금액",
}
ITEM_FIELDS = {
    "description": "품명", "specification": "규격", "unit": "단위", "quantity": "수량",
    "unit_price": "단가", "amount": "금액", "note": "비고",
}
NUMERIC_FIELDS = {"subtotal", "vat", "total", "quantity", "unit_price", "amount"}
ALL_FIELDS = {**FIELDS, "items.index": "품목 · 순번", **{f"items.{k}": f"품목 · {v}" for k, v in ITEM_FIELDS.items()}}
ALL_FIELDS.update({"subtotal": "품목 합계 (부가세 포함)", "vat": "부가세 금액 칸 비우기",
                   "items.unit_price": "품목 · 단가 (부가세 포함)", "items.amount": "품목 · 금액 (부가세 포함)"})


class UserError(ValueError):
    """사용자에게 바로 보여줄 수 있는 입력/변환 오류."""


def number(value) -> Decimal | None:
    if value is None or str(value).strip() == "":
        return None
    if isinstance(value, bool):
        raise UserError("숫자에 참/거짓 값을 사용할 수 없습니다.")
    text = str(value).strip().replace(",", "")
    # 일부 한국어 PDF의 원화 기호는 역슬래시로 추출되고, 회계 표시의
    # 양수도 '(\\ 38,000)' 형태로 추출된다. 기호 없는 '(1,000)'은 음수다.
    won_brackets = bool(re.fullmatch(r"\(\s*[\\₩￦]\s*\d+(?:\.\d+)?\s*\)", text))
    text = text.replace("₩", "").replace("￦", "").replace("\\", "").strip()
    text = re.sub(r"^(?:KRW|금)\s*|\s*(?:원|원정)$", "", text, flags=re.I).strip()
    if text.startswith("(") and text.endswith(")"):
        text = ("" if won_brackets else "-") + text[1:-1].strip()
    if text == "-":
        return None  # 대시를 임의로 0원으로 간주하지 않음
    try:
        result = Decimal(text)
        if not result.is_finite() or abs(result) >= Decimal("1e16"):
            raise InvalidOperation
        return result
    except InvalidOperation as exc:
        raise UserError(f"숫자를 확인해 주세요: {value}") from exc


def display(value) -> str:
    if value is None:
        return ""
    if isinstance(value, Decimal):
        return f"{value:,.0f}" if value == value.to_integral() else format(value, ",f").rstrip("0").rstrip(".")
    return str(value)


@dataclass
class Item:
    description: str = ""
    specification: str = ""
    unit: str = ""
    quantity: Decimal | None = None
    unit_price: Decimal | None = None
    amount: Decimal | None = None
    note: str = ""


@dataclass
class Quote:
    source: str = ""
    title: str = ""
    supplier: str = ""
    customer: str = ""
    quotation_date: str = ""
    quotation_number: str = ""
    requester: str = ""
    department: str = ""
    purpose: str = ""
    notes: str = ""
    subtotal: Decimal | None = None
    vat: Decimal | None = None
    total: Decimal | None = None
    amount_basis: str = "exclusive"
    items: list[Item] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    raw_text: str = ""
    reviewed: bool = False

    def issues(self) -> list[str]:
        errors = []
        if not self.items:
            errors.append("품목이 없습니다. 품목을 추가해 주세요.")
        for index, item in enumerate(self.items, 1):
            if not item.description.strip():
                errors.append(f"{index}번 품목의 품명이 비어 있습니다.")
            if item.amount is None:
                errors.append(f"{index}번 품목의 금액이 비어 있습니다.")
            if item.amount is not None and item.quantity is not None and item.unit_price is not None:
                if abs(item.quantity * item.unit_price - item.amount) > Decimal("1"):
                    errors.append(f"{index}번 품목: 수량 × 단가와 금액이 다릅니다. 할인 등이 있으면 금액을 유지하고 단가를 비워 주세요.")
            for key in ("description", "specification", "unit", "note"):
                if "(cid:" in getattr(item, key):
                    errors.append(f"{index}번 품목의 {ITEM_FIELDS[key]}에 읽지 못한 PDF 문자가 있습니다. 해당 품목을 두 번 클릭하고 원본과 대조해 '(cid:...)' 부분을 수정해 주세요.")
        control = self.total if self.amount_basis == "inclusive" else self.subtotal
        control_name = "합계금액(부가세 포함)" if self.amount_basis == "inclusive" else "공급가액"
        if self.items and all(i.amount is not None for i in self.items) and control is not None:
            if sum((i.amount for i in self.items), Decimal(0)) != control:
                errors.append(f"품목 금액의 합과 {control_name}이 다릅니다. 금액 기준을 확인하고 할인/배송비도 별도 품목으로 입력해 주세요.")
        if self.subtotal is not None and self.vat is not None and self.total is not None:
            if self.subtotal + self.vat != self.total:
                errors.append("공급가액 + 부가세와 합계금액이 다릅니다.")
        return errors

    def value(self, key: str, index: int = 0):
        if key.startswith("items."):
            if index >= len(self.items):
                return None
            attribute = key.split(".", 1)[1]
            return Decimal(index + 1) if attribute == "index" else getattr(self.items[index], attribute)
        return getattr(self, key)


def fingerprint(path: str | Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


@dataclass
class Mapping:
    field: str = "supplier"
    sheet: str = ""
    cell: str = "A1"
    page: int = 1
    x: float = 0
    y: float = 0
    width: float = 100
    height: float = 24
    count: int = 1
    step: float = 1
    font_size: float = 10
    align: str = "left"


@dataclass
class Profile:
    template: str
    template_hash: str
    mappings: list[Mapping] = field(default_factory=list)
    version: int = 1
    default_unit: str = ""

    @classmethod
    def create(cls, template: str):
        return cls(str(Path(template).resolve()), fingerprint(template))

    def check_template(self):
        if not Path(self.template).is_file():
            raise UserError("저장한 양식 파일을 찾을 수 없습니다. 양식을 다시 선택해 주세요.")
        if fingerprint(self.template) != self.template_hash:
            raise UserError("저장한 연결 설정과 양식 파일이 다릅니다. 변경된 양식의 입력 위치를 다시 확인해 주세요.")

    def save(self, path: str):
        data = asdict(self)
        data["template"] = str(Path(self.template).resolve())
        Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    @classmethod
    def load(cls, path: str):
        try:
            data = json.loads(Path(path).read_text(encoding="utf-8"))
            if data.get("version") != 1:
                raise UserError("지원하지 않는 연결 설정 버전입니다.")
            data["mappings"] = [Mapping(**m) for m in data["mappings"]]
            if not Path(data["template"]).is_absolute():
                data["template"] = str(Path(path).parent / data["template"])
            return cls(**data)
        except (KeyError, TypeError, json.JSONDecodeError) as exc:
            raise UserError("올바른 연결 설정 JSON 파일이 아닙니다.") from exc
