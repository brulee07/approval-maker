"""견적 원본을 보존하면서 품의서용 부가세 포함 품목을 계산한다."""
from __future__ import annotations

import copy
from decimal import Decimal
from fractions import Fraction

from .models import Item, Quote, UserError, display


def allocate_vat(amounts: list[Decimal], vat: Decimal) -> list[Decimal]:
    """공급가액 비례 배분. 원 미만 잔여가 큰 순서로 1원씩 더해 합계를 보존한다."""
    if vat != vat.to_integral_value():
        raise UserError("정수 단가로 부가세를 배분하려면 부가세 금액을 원 단위로 입력해 주세요.")
    if vat == 0:
        return [Decimal(0)] * len(amounts)
    base = sum(amounts, Decimal(0))
    if base == 0:
        raise UserError("품목 공급가액 합계가 0원이어서 부가세를 비례 배분할 수 없습니다.")
    shares = [Fraction(vat) * Fraction(amount) / Fraction(base) for amount in amounts]
    whole = [share.numerator // share.denominator for share in shares]
    remaining = int(vat) - sum(whole)
    order = sorted(range(len(shares)), key=lambda i: (-(shares[i] - whole[i]), i))
    for index in order[:remaining]:
        whole[index] += 1
    return [Decimal(value) for value in whole]


def integer_priced_rows(item: Item, gross_amount: Decimal, index: int) -> list[Item]:
    """필요하면 두 정수 단가로 나눠 원래 수량과 부가세 포함 금액을 보존한다."""
    result = copy.deepcopy(item)
    result.amount = gross_amount
    quantity = item.quantity
    if quantity is None:
        # 금액만 있는 양식은 허용하되, 공급가 단가가 출력되는 일은 막는다.
        result.unit_price = None
        return [result]
    if quantity == 0:
        raise UserError(f"{index}번 품목의 수량이 0입니다. 부가세 포함 단가 계산 전에 수량을 확인해 주세요.")
    if quantity == quantity.to_integral_value() and gross_amount != gross_amount.to_integral_value():
        raise UserError(f"{index}번 품목의 금액에 소수 원이 있습니다. 정수 수량·단가로 총액을 유지할 수 없어 금액 확인이 필요합니다.")
    exact_price = Fraction(gross_amount) / Fraction(quantity)
    low_price = Decimal(exact_price.numerator // exact_price.denominator)
    high_quantity = gross_amount - quantity * low_price
    low_quantity = quantity - high_quantity
    rows = []
    for row_quantity, price in ((low_quantity, low_price), (high_quantity, low_price + 1)):
        if row_quantity == 0:
            continue
        row = copy.deepcopy(result)
        row.quantity = row_quantity
        row.unit_price = price
        row.amount = row_quantity * price
        rows.append(row)
    return rows


def approval_total(quote: Quote) -> Decimal:
    """출력 행에서 합계를 계산하고, 단가가 있는 행은 수량 × 단가와 정확히 대조한다."""
    total = Decimal(0)
    for index, item in enumerate(quote.items, 1):
        if item.amount is None:
            raise UserError(f"{index}번 출력 행의 금액이 비어 있습니다.")
        amount = item.amount
        if item.quantity is not None and item.unit_price is not None:
            amount = item.quantity * item.unit_price
            if amount != item.amount:
                raise UserError(f"{index}번 출력 행: 수량 × 단가 {display(amount)}원과 품목 금액 {display(item.amount)}원이 다릅니다. 파일을 생성하지 않습니다.")
        total += amount
    return total


def prepare_approval_quote(source: Quote, default_unit: str = "") -> Quote:
    """출력 전용 사본. 모든 출력 형식이 같은 계산 결과를 사용한다."""
    if source.total is None:
        raise UserError("견적서 합계금액이 비어 있어 품의서 총액과 비교할 수 없습니다. 원본의 부가세 포함 합계금액을 확인하여 입력해 주세요.")
    issues = source.issues()
    if issues:
        raise UserError("\n".join(issues))
    if source.amount_basis not in {"inclusive", "exclusive"}:
        raise UserError("원본 품목 금액의 부가세 포함/별도 기준을 확인해 주세요.")
    result = copy.deepcopy(source)
    base = sum((item.amount for item in source.items), Decimal(0))
    if source.amount_basis == "exclusive":
        vat = source.vat
        if vat is None:
            vat = source.total - base
            result.warnings.append("부가세를 원본 합계금액에서 품목 공급가액 합계를 빼서 확인했습니다.")
        total = base + vat
        if source.total != total:
            raise UserError("품목 공급가액 + 부가세와 견적서 합계금액이 다릅니다.")
        allocations = allocate_vat([item.amount for item in source.items], vat)
        result.items = []
        for index, (item, tax) in enumerate(zip(source.items, allocations), 1):
            rows = integer_priced_rows(item, item.amount + tax, index)
            result.items.extend(rows)
            if len(rows) > 1:
                details = " + ".join(f"단가 {display(row.unit_price)}원 × 수량 {display(row.quantity)}" for row in rows)
                result.warnings.append(f"{index}번 품목을 정수 단가로 {len(rows)}행에 나누었습니다: {details}")
        result.warnings.append(f"별도 부가세 {display(vat)}원을 품목별로 배분했습니다. 품의서에는 부가세 포함 단가·금액만 입력합니다.")
    else:
        total = source.total
        if total != base:
            raise UserError("부가세 포함 품목 금액 합계와 견적서 총액이 다릅니다.")
        result.items = []
        for index, item in enumerate(source.items, 1):
            if item.quantity is not None and (item.unit_price is None or item.quantity * item.unit_price != item.amount):
                rows = integer_priced_rows(item, item.amount, index)
                result.items.extend(rows)
                result.warnings.append(f"{index}번 품목은 원본 품목 금액을 유지하고 정수 단가 {len(rows)}행으로 맞췄습니다. 수량 합계와 총액은 동일합니다.")
            else:
                result.items.append(copy.deepcopy(item))
        result.warnings.append("원본 품목 금액에 이미 부가세가 포함되어 있어 추가로 더하지 않았습니다.")
    for item in result.items:
        if not item.unit and default_unit:
            item.unit = default_unit
    result.amount_basis = "inclusive"
    result.total = total
    result.subtotal = total
    result.vat = None  # 기존 연결 설정에 부가세 셀이 있어도 별도 금액을 쓰지 않는다.
    output_total = approval_total(result)
    if output_total != source.total:
        raise UserError(f"견적서 총액 {display(source.total)}원과 품의서 총액 {display(output_total)}원이 다릅니다. 차액 {display(output_total - source.total)}원. 파일을 생성하지 않습니다.")
    return result
