# -*- coding: utf-8 -*-
"""표준 발주 양식(최종 통합본) 정의 및 스타일링.

기준 파일(``0730`` / 260730 통합본)의 열 구성/서식을 그대로 재현한다.
- 열(12): NO · 주문인명 · 수령인명 · 수령인핸드폰번호 · 수령인핸드폰번호 · 우편번호 ·
  주소 · 배송메세지 · 상품정보 · 주문수량 · 택배사 · 송장번호
- 글꼴 맑은 고딕 10pt, 전 셀 얇은 테두리, 택배사/송장번호 헤더 연한 주황
- 주문인명은 채워서 출력(주문인 정보가 없으면 수령인명으로 채움)
- 수령인핸드폰번호 두 열 모두 수령인 전화번호(기준 파일과 동일)
"""
from __future__ import annotations

from copy import copy

from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Side

# ---------------------------------------------------------------------------
# 표준 양식 열 정의 (기준 통합본과 동일). phone 키가 두 번 → 두 열 모두 수령인 전화.
# ---------------------------------------------------------------------------
STANDARD_COLUMNS = [
    ("no",        "NO"),
    ("orderer",   "주문인명"),
    ("recipient", "수령인명"),
    ("phone",     "수령인핸드폰번호"),
    ("phone",     "수령인핸드폰번호"),
    ("postcode",  "우편번호"),
    ("address",   "주소"),
    ("message",   "배송메세지"),
    ("product",   "상품정보"),
    ("qty",       "주문수량"),
    ("courier",   "택배사"),
    ("invoice",   "송장번호"),
]

FIELD_KEYS = [key for key, _ in STANDARD_COLUMNS]

COLUMN_WIDTHS = {
    "A": 4.25, "B": 12.625, "C": 13.0, "D": 15.0, "E": 13.0, "F": 8.375,
    "G": 23.25, "H": 6.5, "I": 42.25, "J": 8.0, "K": 13.0, "L": 14.375,
}

_CENTER_COLS = {1, 2, 3, 10}      # NO, 주문인명, 수령인명, 주문수량
_TEXT_COLS = {4, 5, 6, 12}        # 수령인핸드폰x2, 우편번호, 송장번호
_HEADER_FILL_COLS = {11, 12}      # 택배사, 송장번호

FONT_NAME = "맑은 고딕"
FONT_SIZE = 10
HEADER_ROW_HEIGHT = 30.75

_THIN = Side(style="thin", color="FF000000")
THIN_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
HEADER_FILL = PatternFill(patternType="solid", fgColor=Color(theme=5, tint=0.7999816888943144))


def write_standard_sheet(ws, rows):
    """``rows`` (표준 필드 dict 리스트)를 표준 서식으로 ``ws`` 에 기록한다."""
    base_font = Font(name=FONT_NAME, size=FONT_SIZE)

    for c, (_key, header) in enumerate(STANDARD_COLUMNS, start=1):
        cell = ws.cell(row=1, column=c, value=header)
        cell.font = copy(base_font)
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if c in _HEADER_FILL_COLS:
            cell.fill = copy(HEADER_FILL)
    ws.row_dimensions[1].height = HEADER_ROW_HEIGHT

    for r, row in enumerate(rows, start=2):
        for c, key in enumerate(FIELD_KEYS, start=1):
            cell = ws.cell(row=r, column=c, value=row.get(key))
            cell.font = copy(base_font)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(
                horizontal="center" if c in _CENTER_COLS else "left", vertical="center")
            if c in _TEXT_COLS:
                cell.number_format = "@"

    for col_letter, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    return ws
