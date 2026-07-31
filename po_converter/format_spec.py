# -*- coding: utf-8 -*-
"""표준 발주 양식(최종 통합본) 정의 및 스타일링.

기준 파일(``260730``)의 열 구성/서식을 그대로 재현한다.
- 열: 수령인명 · 수령인연락처 · 주소 · 주문상품명 · 배송시 요청사항 (5열)
- 글꼴: 맑은 고딕 11pt, 테두리/채움 없음(원본과 동일한 단순 서식)
"""
from __future__ import annotations

from copy import copy

from openpyxl.styles import Font

# ---------------------------------------------------------------------------
# 표준 양식 열 정의 (기준 260730 파일과 동일)
# ---------------------------------------------------------------------------
# 표준 필드 키 → 화면 헤더 문자열.
# (기준 파일 C에서는 상품 열 머리글이 비어 있으나, 의미가 분명하도록 '주문상품명'으로 표기)
STANDARD_COLUMNS = [
    ("recipient", "수령인명"),
    ("phone",     "수령인연락처"),
    ("address",   "주소"),
    ("product",   "주문상품명"),
    ("message",   "배송시 요청사항"),
]

FIELD_KEYS = [key for key, _ in STANDARD_COLUMNS]

# 열 너비 (기준 260730 파일에서 추출)
COLUMN_WIDTHS = {"A": 13.0, "B": 13.0, "C": 13.0, "D": 72.38, "E": 13.0}

FONT_NAME = "맑은 고딕"
FONT_SIZE = 11


def write_standard_sheet(ws, rows):
    """``rows`` (표준 필드 dict 리스트)를 표준 서식으로 ``ws`` 에 기록한다."""
    base_font = Font(name=FONT_NAME, size=FONT_SIZE)

    # 헤더 -----------------------------------------------------------------
    for c, (_key, header) in enumerate(STANDARD_COLUMNS, start=1):
        cell = ws.cell(row=1, column=c, value=header)
        cell.font = copy(base_font)

    # 데이터 ----------------------------------------------------------------
    for r, row in enumerate(rows, start=2):
        for c, key in enumerate(FIELD_KEYS, start=1):
            cell = ws.cell(row=r, column=c, value=row.get(key))
            cell.font = copy(base_font)
            # 전화번호는 텍스트로 고정(앞자리 0 보존)
            if key == "phone":
                cell.number_format = "@"

    # 열 너비 ---------------------------------------------------------------
    for col_letter, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    return ws
