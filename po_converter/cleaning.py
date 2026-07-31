# -*- coding: utf-8 -*-
"""업체별 원본 데이터를 표준 양식에 맞게 정리하는 유틸리티."""
from __future__ import annotations

import re
from typing import Optional, Tuple


def clean_text(value) -> str:
    """공백 정리 후 문자열로 반환. None/NaN → ''."""
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() == "nan":
        return ""
    # 여러 개의 공백을 하나로
    return re.sub(r"\s+", " ", text)


def normalize_phone(value) -> str:
    """전화번호 구분자를 통일( ~ → - )하고 공백을 제거한다.

    원본이 이미 숫자만 있거나 하이픈이 있는 경우 그대로 두어 원본 충실도를 유지한다.
    """
    text = clean_text(value)
    if not text:
        return ""
    text = text.replace("~", "-").replace(" ", "")
    return text


def clean_postcode(value) -> str:
    """우편번호를 5자리 XXX-XX 형태로 정리한다.

    - 숫자만 추출해 5자리이면 ``XXX-XX`` 로 포맷
    - 그 외에는 이중 하이픈(``--``)만 단일화해 원본을 최대한 보존
    """
    text = clean_text(value)
    if not text:
        return ""
    digits = re.sub(r"\D", "", text)
    # 엑셀이 숫자로 저장하며 앞자리 0을 떨어뜨린 경우(4자리) → 5자리로 복원
    if len(digits) == 4:
        digits = "0" + digits
    # 기준 통합본 형식: 5자리 XXX-XX (예: 036-32)
    if len(digits) == 5:
        return f"{digits[:3]}-{digits[3:]}"
    if len(digits) == 6:  # 구 우편번호 XXX-XXX
        return f"{digits[:3]}-{digits[3:]}"
    return re.sub(r"-{2,}", "-", text)


# 상품명 끝에 붙은 수량 표기: "/3개", "▶ 기본/1개", "- 2개" 등
_QTY_TAIL = re.compile(r"\s*[/·\-]?\s*(\d+)\s*개\s*$")


def split_product_qty(value, default_qty: int = 1) -> Tuple[str, int]:
    """상품명 문자열에서 끝에 붙은 "/N개" 수량을 분리한다.

    반환: (수량이 제거된 상품명, 수량).  수량 표기가 없으면 default_qty 사용.
    """
    text = clean_text(value)
    if not text:
        return "", default_qty
    m = _QTY_TAIL.search(text)
    if m:
        qty = int(m.group(1))
        product = text[: m.start()].strip()
        # 끝에 남은 구분자 정리
        product = re.sub(r"[\s/·\-▶]+$", "", product).strip()
        return product, qty
    return text, default_qty


# 주소 앞부분의 "[509-51]" 형태 우편번호 추출
_ADDR_POST = re.compile(r"^\s*\[\s*(\d{3})\s*-?\s*(\d{2})\s*\]\s*")


def extract_postcode_from_address(value) -> Tuple[str, str]:
    """"[509-51] 경남 ..." 형태 주소에서 우편번호와 나머지 주소를 분리한다.

    반환: (우편번호 'XXX-XX' 또는 '', 앞 우편번호가 제거된 주소).
    """
    text = clean_text(value)
    if not text:
        return "", ""
    m = _ADDR_POST.match(text)
    if m:
        postcode = f"{m.group(1)}-{m.group(2)}"
        address = text[m.end():].strip()
        return postcode, address
    return "", text


# 주소 끝에 덧붙은 "  이름 010-1234-5678" 패턴(수령인/연락처 중복 표기)
_ADDR_TAIL_NAMEPHONE = re.compile(
    r"\s{2,}[가-힣A-Za-z]{2,5}\s+0\d[\d~\-\s]{7,}$"
)
# 주소 끝에 전화번호만 덧붙은 경우
_ADDR_TAIL_PHONE = re.compile(r"\s{2,}0\d[\d~\-\s]{7,}$")


def clean_address(value, *, strip_trailing_contact: bool = False) -> str:
    """주소 정리.

    - ``~`` → ``-`` (일부 업체가 하이픈 대신 물결표 사용)
    - strip_trailing_contact=True 이면 끝에 중복으로 붙은 이름/전화번호 제거
    """
    if value is None:
        return ""
    # 끝에 덧붙은 "이름 전화번호"는 공백 정리(이중공백 축약) 전에 제거해야 한다.
    text = str(value).strip()
    if text.lower() == "nan":
        return ""
    if strip_trailing_contact:
        text = _ADDR_TAIL_NAMEPHONE.sub("", text)
        text = _ADDR_TAIL_PHONE.sub("", text)
    text = text.replace("~", "-")
    return re.sub(r"\s+", " ", text).strip()


def parse_int(value, default: Optional[int] = 1) -> Optional[int]:
    """수량 등 정수 파싱. 실패 시 default."""
    text = clean_text(value)
    if not text:
        return default
    m = re.search(r"\d+", text)
    return int(m.group()) if m else default


def append_qty(product: str, qty) -> str:
    """수량이 별도 열에 있는 양식(유니어·패션지오)에서 2개 이상이면 상품명 끝에 '/N개' 표기.

    (기준 260730 양식은 수량을 상품명 안에 두므로 정보 손실을 막기 위해 붙인다.)
    """
    product = (product or "").strip()
    try:
        n = int(qty)
    except (TypeError, ValueError):
        return product
    if n > 1:
        return f"{product} /{n}개".strip()
    return product
