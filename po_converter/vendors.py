# -*- coding: utf-8 -*-
"""업체별 발주 양식 정의: 자동 감지 시그니처 + 표준 양식 매핑.

새 업체를 추가하려면 아래 ``VENDORS`` 리스트에 ``Vendor`` 하나를 등록하면 된다.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional

from .cleaning import (
    clean_address,
    clean_postcode,
    clean_text,
    extract_postcode_from_address,
    normalize_phone,
    parse_int,
    split_product_qty,
)


def normalize_header(value) -> str:
    """헤더 문자열 정규화: 앞뒤 공백 제거 + 내부 공백 모두 제거."""
    if value is None:
        return ""
    return re.sub(r"\s+", "", str(value).strip())


class Row:
    """한 데이터 행에 대한 접근자.

    ``h(header)`` 로 헤더 이름(공백 무시)으로, ``pos(idx)`` 로 열 위치(1-based)로 값을 얻는다.
    """

    def __init__(self, values: List, header_map: Dict[str, int]):
        self._values = values           # 0-based 리스트
        self._header_map = header_map    # 정규화 헤더 → 1-based 열 인덱스

    def h(self, header: str):
        idx = self._header_map.get(normalize_header(header))
        if idx is None:
            return None
        return self.pos(idx)

    def pos(self, idx: int):
        if 1 <= idx <= len(self._values):
            return self._values[idx - 1]
        return None


@dataclass
class Vendor:
    key: str                                  # 내부 식별자
    default_name: str                         # 파일명에 쓸 기본 업체명
    signature: List[str]                      # 이 헤더들이 모두 있으면 이 업체
    map_row: Callable[[Row], dict]            # 한 행 → 표준 필드 dict
    name_from_headers: Optional[Callable[[Dict[str, int]], Optional[str]]] = None
    name_from_sheet: Optional[Callable[[str], Optional[str]]] = None
    date_header: Optional[str] = None         # 주문일자 등 날짜가 담긴 헤더
    name_confident: bool = True               # 업체명을 원본에서 확정할 수 있는지
    signature_norm: List[str] = field(default_factory=list)

    def __post_init__(self):
        self.signature_norm = [normalize_header(s) for s in self.signature]

    def matches(self, present: set) -> bool:
        return all(s in present for s in self.signature_norm)


# ---------------------------------------------------------------------------
# 업체별 행 매핑
# ---------------------------------------------------------------------------
def _map_unier(row: Row) -> dict:
    """유니어: 수신인/연락처1/우편번호/주소/상품명/실수량/배송메모/택배사명/송장번호."""
    name = clean_text(row.h("수신인"))
    phone = normalize_phone(row.h("연락처1"))
    product = clean_text(row.h("상품명"))
    attr = clean_text(row.h("속성1 속성2"))
    if attr:
        product = f"{product} {attr}".strip()
    return {
        "orderer": name,
        "recipient": name,
        "phone": phone,
        "phone2": phone,
        "postcode": clean_postcode(row.h("우편번호")),
        "address": clean_address(row.h("주소"), strip_trailing_contact=True),
        "message": clean_text(row.h("배송메모")),
        "product": product,
        "qty": parse_int(row.h("실수량")),
        "courier": clean_text(row.h("택배사명")),
        "invoice": clean_text(row.h("송장번호")),
    }


def _map_blueberry(row: Row) -> dict:
    """블루베리 퓨레 업체: 수령인명/수령인연락처/주소/(무헤더 상품)/배송시 요청사항."""
    name = clean_text(row.h("수령인명"))
    phone = normalize_phone(row.h("수령인연락처"))
    # 상품 열은 헤더가 비어 있어 위치(D=4열)로 접근
    product_raw = row.h("주문상품명") or row.pos(4)
    product, qty = split_product_qty(product_raw)
    return {
        "orderer": name,
        "recipient": name,
        "phone": phone,
        "phone2": phone,
        "postcode": "",
        "address": clean_address(row.h("주소")),
        "message": clean_text(row.h("배송시 요청사항")),
        "product": product,
        "qty": qty,
        "courier": "",
        "invoice": "",
    }


def _map_chikjeup(row: Row) -> dict:
    """칡즙 업체: 수령인/수령인연락처/주소/주문상품명/비고."""
    name = clean_text(row.h("수령인"))
    phone = normalize_phone(row.h("수령인연락처"))
    product, qty = split_product_qty(row.h("주문상품명"))
    return {
        "orderer": name,
        "recipient": name,
        "phone": phone,
        "phone2": phone,
        "postcode": "",
        "address": clean_address(row.h("주소")),
        "message": clean_text(row.h("비고")),
        "product": product,
        "qty": qty,
        "courier": "",
        "invoice": "",
    }


def _map_fashiongeo(row: Row) -> dict:
    """패션지오: 주문자명/수취인 명/수취인 휴대폰번호/배송지([우편]주소)/품목명/수량."""
    orderer = clean_text(row.h("주문자명"))
    recipient = clean_text(row.h("수취인 명"))
    phone = normalize_phone(row.h("수취인 휴대폰번호") or row.h("수취인 전화번호"))
    postcode, address = extract_postcode_from_address(row.h("배송지"))
    return {
        "orderer": orderer or recipient,
        "recipient": recipient,
        "phone": phone,
        "phone2": phone,
        "postcode": postcode,
        "address": address,
        "message": clean_text(row.h("배송메세지")),
        "product": clean_text(row.h("품목명")),
        "qty": parse_int(row.h("수량")),
        "courier": clean_text(row.h("택배사")),
        "invoice": clean_text(row.h("송장번호")),
    }


# ---------------------------------------------------------------------------
# 업체명 자동 추출
# ---------------------------------------------------------------------------
def _unier_name(header_map: Dict[str, int]) -> Optional[str]:
    for norm in header_map:
        if "식별번호" in norm:
            m = re.search(r"식별번호[-\s]*([가-힣A-Za-z0-9]+)", norm)
            if m:
                return m.group(1)
    return None


def _fashiongeo_name(sheet_name: str) -> Optional[str]:
    # 예) "20260729_패션지오(주) 발주 건" → "패션지오"
    m = re.search(r"\d{6,8}_\s*([^_]+?)\s*발주", sheet_name or "")
    if m:
        name = m.group(1).strip()
        return re.sub(r"\(주\)|주식회사", "", name).strip()
    return None


# ---------------------------------------------------------------------------
# 등록된 업체 목록 (구체적인 시그니처가 먼저 오도록 정렬)
# ---------------------------------------------------------------------------
VENDORS: List[Vendor] = [
    Vendor(
        key="unier",
        default_name="유니어",
        signature=["수신인", "상품명", "실수량"],
        map_row=_map_unier,
        name_from_headers=_unier_name,
    ),
    Vendor(
        key="fashiongeo",
        default_name="패션지오",
        signature=["품목명", "수취인 명", "배송지"],
        map_row=_map_fashiongeo,
        name_from_sheet=_fashiongeo_name,
        date_header="주문일자",
    ),
    Vendor(
        key="chikjeup",
        default_name="칡즙",
        signature=["수령인", "주문상품명", "비고"],
        map_row=_map_chikjeup,
        name_confident=False,
    ),
    Vendor(
        key="blueberry",
        default_name="블루베리퓨레",
        signature=["수령인명", "수령인연락처", "배송시요청사항"],
        map_row=_map_blueberry,
        name_confident=False,
    ),
]


def detect_vendor(present_headers: set) -> Optional[Vendor]:
    """헤더 집합에 가장 잘 맞는 업체를 반환(가장 많은 시그니처가 일치하는 업체)."""
    best = None
    best_score = 0
    for vendor in VENDORS:
        if vendor.matches(present_headers):
            score = len(vendor.signature_norm)
            if score > best_score:
                best, best_score = vendor, score
    return best
