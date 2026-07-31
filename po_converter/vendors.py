# -*- coding: utf-8 -*-
"""업체별 발주 양식 정의: 자동 감지 시그니처 + 표준 양식(260730) 매핑.

표준 5열: 수령인명 · 수령인연락처 · 주소 · 주문상품명 · 배송시 요청사항
새 업체를 추가하려면 아래 ``VENDORS`` 리스트에 ``Vendor`` 하나를 등록하면 된다.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional

from .cleaning import (
    append_qty,
    clean_address,
    clean_text,
    extract_postcode_from_address,
    normalize_phone,
    parse_int,
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
        self._values = values
        self._header_map = header_map

    def h(self, header: str):
        idx = self._header_map.get(normalize_header(header))
        return self.pos(idx) if idx else None

    def pos(self, idx):
        if idx and 1 <= idx <= len(self._values):
            return self._values[idx - 1]
        return None


@dataclass
class Vendor:
    key: str
    default_name: str
    signature: List[str]
    map_row: Callable[[Row], dict]
    name_from_headers: Optional[Callable[[Dict[str, int]], Optional[str]]] = None
    name_from_sheet: Optional[Callable[[str], Optional[str]]] = None
    date_header: Optional[str] = None
    name_confident: bool = True
    signature_norm: List[str] = field(default_factory=list)

    def __post_init__(self):
        self.signature_norm = [normalize_header(s) for s in self.signature]

    def matches(self, present: set) -> bool:
        return all(s in present for s in self.signature_norm)


# ---------------------------------------------------------------------------
# 업체별 행 매핑 → 표준 5열
# ---------------------------------------------------------------------------
def _map_unier(row: Row) -> dict:
    """유니어: 수신인/연락처1/주소/상품명(+실수량)/배송메모."""
    product = clean_text(row.h("상품명"))
    attr = clean_text(row.h("속성1 속성2"))
    if attr:
        product = f"{product} {attr}".strip()
    product = append_qty(product, parse_int(row.h("실수량")))
    return {
        "recipient": clean_text(row.h("수신인")),
        "phone": normalize_phone(row.h("연락처1")),
        "address": clean_address(row.h("주소"), strip_trailing_contact=True),
        "product": product,
        "message": clean_text(row.h("배송메모")),
    }


def _map_blueberry(row: Row) -> dict:
    """블루베리 퓨레: 수령인명/수령인연락처/주소/(무헤더 상품)/배송시 요청사항."""
    return {
        "recipient": clean_text(row.h("수령인명")),
        "phone": normalize_phone(row.h("수령인연락처")),
        "address": clean_address(row.h("주소")),
        "product": clean_text(row.h("주문상품명") or row.pos(4)),  # 수량은 상품명에 포함된 채로 유지
        "message": clean_text(row.h("배송시 요청사항")),
    }


def _map_chikjeup(row: Row) -> dict:
    """칡즙: 수령인/수령인연락처/주소/주문상품명/비고."""
    return {
        "recipient": clean_text(row.h("수령인")),
        "phone": normalize_phone(row.h("수령인연락처")),
        "address": clean_address(row.h("주소")),
        "product": clean_text(row.h("주문상품명")),  # 수량은 상품명에 포함된 채로 유지
        "message": clean_text(row.h("비고")),
    }


def _map_fashiongeo(row: Row) -> dict:
    """패션지오: 수취인 명/수취인 휴대폰번호/배송지([우편]주소)/품목명(+수량)/배송메세지."""
    pa = extract_postcode_from_address(row.h("배송지"))  # 앞의 [509-51] 우편번호 제거
    product = append_qty(clean_text(row.h("품목명")), parse_int(row.h("수량")))
    return {
        "recipient": clean_text(row.h("수취인 명")),
        "phone": normalize_phone(row.h("수취인 휴대폰번호") or row.h("수취인 전화번호")),
        "address": pa[1],
        "product": product,
        "message": clean_text(row.h("배송메세지")),
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
    m = re.search(r"\d{6,8}_\s*([^_]+?)\s*발주", sheet_name or "")
    if m:
        return re.sub(r"\(주\)|주식회사", "", m.group(1)).strip()
    return None


# ---------------------------------------------------------------------------
# 등록된 업체 목록
# ---------------------------------------------------------------------------
VENDORS: List[Vendor] = [
    Vendor(key="unier", default_name="유니어",
           signature=["수신인", "상품명", "실수량"], map_row=_map_unier,
           name_from_headers=_unier_name),
    Vendor(key="fashiongeo", default_name="패션지오",
           signature=["품목명", "수취인 명", "배송지"], map_row=_map_fashiongeo,
           name_from_sheet=_fashiongeo_name, date_header="주문일자"),
    Vendor(key="chikjeup", default_name="칡즙",
           signature=["수령인", "주문상품명", "비고"], map_row=_map_chikjeup,
           name_confident=False),
    Vendor(key="blueberry", default_name="블루베리퓨레",
           signature=["수령인명", "수령인연락처", "배송시 요청사항"], map_row=_map_blueberry,
           name_confident=False),
]


def detect_vendor(present_headers: set) -> Optional[Vendor]:
    """헤더 집합에 가장 잘 맞는 업체를 반환(가장 많은 시그니처가 일치)."""
    best, best_score = None, 0
    for vendor in VENDORS:
        if vendor.matches(present_headers) and len(vendor.signature_norm) > best_score:
            best, best_score = vendor, len(vendor.signature_norm)
    return best
