"""원본 페이지를 시각적으로 대조한 PDF 문자 보정만 적용한다.

CID는 글꼴마다 의미가 다르므로 같은 번호를 다른 PDF에 일반화하지 않는다.
파일 내용의 SHA-256과 추출된 전체 품명이 모두 일치해야 보정한다.
"""

from .models import Quote


VERIFIED_DESCRIPTIONS = {
    "9e82960ed5afffe871468a0ab4eb17a16353a88d33613d7bd6a90025d8dca16b": {
        "2(cid:1070)1 Ratio Bevel Gear Set (6mm D-Bore Pinion Gear)":
            "2:1 Ratio Bevel Gear Set (6mm D-Bore Pinion Gear)",
    },
}


def apply_verified_corrections(quote: Quote, source_sha256: str) -> None:
    corrections = VERIFIED_DESCRIPTIONS.get(source_sha256, {})
    for index, item in enumerate(quote.items, 1):
        corrected = corrections.get(item.description)
        if corrected is not None:
            item.description = corrected
            quote.warnings.append(
                f"{index}번 품명: 원본 화면과 대조해 확인한 PDF 특수문자를 자동 보정했습니다 ({corrected})."
            )
