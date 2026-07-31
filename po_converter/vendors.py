# -*- coding: utf-8 -*-
"""업체별 발주 양식 정의: 자동 감지 시그니처 + 최종 통합본(12열) 매핑.

표준 필드: orderer(주문인명) · recipient(수령인명) · phone(수령인핸드폰) · postcode · address ·
message(배송메세지) · product(상품정보) · qty(주문수량) · courier(택배사) · invoice(송장번호)
- 주문인 정보가 없는 양식은 orderer 를 recipient 로 채운다(앞 열이 비지 않도록).
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
    if value is None:
        return ""
    return re.sub(r"\s+", "", str(value).strip())


class Row:
    """한 데이터 행 접근자. ``h(*headers)`` 후보 헤더 중 먼저 값이 있는 것, ``pos(idx)`` 열 위치."""

    def __init__(self, values: List, header_map: Dict[str, int]):
        self._values = values
        self._header_map = header_map

    def h(self, *headers: str):
        for header in headers:
            idx = self._header_map.get(normalize_header(header))
            if idx:
                v = self.pos(idx)
                if v not in (None, ""):
                    return v
        return None

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
    name_from_filename: Optional[Callable[[str], Optional[str]]] = None
    date_header: Optional[str] = None
    name_confident: bool = True
    signature_norm: List[str] = field(default_factory=list)

    def __post_init__(self):
        self.signature_norm = [normalize_header(s) for s in self.signature]

    def matches(self, present: set) -> bool:
        return all(s in present for s in self.signature_norm)


def _std(orderer, recipient, phone, postcode, address, message, product, qty, courier, invoice):
    """표준 행 생성. 주문인명이 비면 수령인명으로 채운다."""
    recipient = clean_text(recipient)
    orderer = clean_text(orderer) or recipient
    return {
        "orderer": orderer, "recipient": recipient, "phone": normalize_phone(phone),
        "postcode": postcode, "address": address, "message": clean_text(message),
        "product": product, "qty": qty, "courier": clean_text(courier), "invoice": clean_text(invoice),
    }


# ---------------------------------------------------------------------------
# 업체별 행 매핑
# ---------------------------------------------------------------------------
def _map_ys(row: Row) -> dict:
    """이미 최종양식(12열)인 파일: 그대로 통과 + 가벼운 정리."""
    return _std(
        orderer=row.h("주문인명", "주문자명"),
        recipient=row.h("수령인명", "수취인명"),
        phone=row.h("수령인핸드폰번호", "수령인핸드폰", "수령인연락처"),
        postcode=clean_postcode(row.h("우편번호", "우편")),
        address=clean_address(row.h("주소", "배송지")),
        message=row.h("배송메세지", "배송메모", "전언"),
        product=clean_text(row.h("상품정보", "상품명")),
        qty=parse_int(row.h("주문수량", "수량")),
        courier=row.h("택배사", "택배사명"), invoice=row.h("송장번호", "운송장번호"),
    )


def _map_unier(row: Row) -> dict:
    product = clean_text(row.h("상품명"))
    attr = clean_text(row.h("속성1 속성2"))
    if attr:
        product = f"{product} {attr}".strip()
    name = row.h("수신인")
    return _std(
        orderer=name, recipient=name, phone=row.h("연락처1"),
        postcode=clean_postcode(row.h("우편번호")),
        address=clean_address(row.h("주소"), strip_trailing_contact=True),
        message=row.h("배송메모"), product=product, qty=parse_int(row.h("실수량")),
        courier=row.h("택배사명"), invoice=row.h("송장번호"),
    )


def _map_blueberry(row: Row) -> dict:
    product, qty = split_product_qty(row.h("주문상품명") or row.pos(4))
    name = row.h("수령인명")
    return _std(
        orderer=name, recipient=name, phone=row.h("수령인연락처"),
        postcode="", address=clean_address(row.h("주소")),
        message=row.h("배송시 요청사항"), product=product, qty=qty, courier="", invoice="",
    )


def _map_chikjeup(row: Row) -> dict:
    product, qty = split_product_qty(row.h("주문상품명"))
    name = row.h("수령인")
    return _std(
        orderer=name, recipient=name, phone=row.h("수령인연락처"),
        postcode="", address=clean_address(row.h("주소")),
        message=row.h("비고"), product=product, qty=qty, courier="", invoice="",
    )


def _map_fashiongeo(row: Row) -> dict:
    pa = extract_postcode_from_address(row.h("배송지"))
    return _std(
        orderer=row.h("주문자명"), recipient=row.h("수취인 명"),
        phone=row.h("수취인 휴대폰번호", "수취인 전화번호"),
        postcode=clean_postcode(pa[0]), address=pa[1], message=row.h("배송메세지"),
        product=clean_text(row.h("품목명")), qty=parse_int(row.h("수량")),
        courier=row.h("택배사"), invoice=row.h("송장번호"),
    )


def _map_daon(row: Row) -> dict:
    return _std(
        orderer=row.h("주문인"), recipient=row.h("받는인"),
        phone=row.h("받는인핸드폰", "받는인연락처"),
        postcode=clean_postcode(row.h("우편", "우편번호")),
        address=clean_address(row.h("배송지", "주소")),
        message=row.h("전언", "배송메세지", "배송메모"),
        product=clean_text(row.h("상품명", "상품정보")), qty=parse_int(row.h("수량", "주문수량")),
        courier=row.h("택배사", "택배사명"), invoice=row.h("송장번호", "운송장번호"),
    )


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


def _daon_name(stem: str) -> Optional[str]:
    first = re.split(r"[_\s]", stem or "", 1)[0]
    first = re.sub(r"주식회사|\(주\)", "", first).strip()
    return first or None


def _ys_name(stem: str) -> Optional[str]:
    s = re.sub(r"(?i)^ys[_\s]*", "", stem or "")
    s = re.sub(r"\d{4,8}", "", s)
    s = re.sub(r"발주\s*수정본|발주서|발주건|발주|수정본", "", s)
    s = re.sub(r"[_\s]+", " ", s).strip()
    return s or None


# ---------------------------------------------------------------------------
# 등록된 업체 목록
# ---------------------------------------------------------------------------
VENDORS: List[Vendor] = [
    Vendor(key="ys", default_name="YS",
           signature=["주문인명", "수령인명", "상품정보"], map_row=_map_ys,
           name_from_filename=_ys_name),
    Vendor(key="unier", default_name="유니어",
           signature=["수신인", "상품명", "실수량"], map_row=_map_unier,
           name_from_headers=_unier_name),
    Vendor(key="fashiongeo", default_name="패션지오",
           signature=["품목명", "수취인 명", "배송지"], map_row=_map_fashiongeo,
           name_from_sheet=_fashiongeo_name, date_header="주문일자"),
    Vendor(key="daon", default_name="다온에프앤씨",
           signature=["주문인", "받는인", "상품명"], map_row=_map_daon,
           name_from_filename=_daon_name),
    Vendor(key="chikjeup", default_name="칡즙",
           signature=["수령인", "주문상품명", "비고"], map_row=_map_chikjeup,
           name_confident=False),
    Vendor(key="blueberry", default_name="블루베리퓨레",
           signature=["수령인명", "수령인연락처", "배송시 요청사항"], map_row=_map_blueberry,
           name_confident=False),
]


def detect_vendor(present_headers: set) -> Optional[Vendor]:
    best, best_score = None, 0
    for vendor in VENDORS:
        if vendor.matches(present_headers) and len(vendor.signature_norm) > best_score:
            best, best_score = vendor, len(vendor.signature_norm)
    return best
