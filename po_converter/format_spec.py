# -*- coding: utf-8 -*-
"""표준 발주 양식(최종 통합본) 정의 및 스타일링.

기준 파일(``0730`` 통합본)의 열 구성/서식을 그대로 재현한다.
- 글꼴: 맑은 고딕 10pt
- 헤더 행 높이 30.75, 모든 셀 얇은 테두리
- 택배사/송장번호 헤더에 연한 주황 채움(theme 5, tint 0.8)
"""
from __future__ import annotations

from copy import copy

from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 표준 양식 열 정의 (기준 통합본과 동일한 순서)
# ---------------------------------------------------------------------------
# 표준 필드 키 → 화면에 표시되는 헤더 문자열.
# ``phone`` 이 두 번(수령인핸드폰번호) 반복되는 것도 기준 파일과 동일하게 재현한다.
STANDARD_COLUMNS = [
    ("no",           "NO"),
    ("orderer",      "주문인명"),
    ("recipient",    "수령인명"),
    ("phone",        "수령인핸드폰번호"),
    ("phone2",       "수령인핸드폰번호"),
    ("postcode",     "우편번호"),
    ("address",      "주소"),
    ("message",      "배송메세지"),
    ("product",      "상품정보"),
    ("qty",          "주문수량"),
    ("courier",      "택배사"),
    ("invoice",      "송장번호"),
]

FIELD_KEYS = [key for key, _ in STANDARD_COLUMNS]

# 열 너비 (기준 파일에서 추출)
COLUMN_WIDTHS = {
    "A": 4.25, "B": 12.625, "C": 13.0, "D": 15.0, "E": 13.0, "F": 8.375,
    "G": 23.25, "H": 6.5, "I": 42.25, "J": 8.0, "K": 13.0, "L": 14.375,
}

# 가운데 정렬할 열(1-based). 나머지는 왼쪽 정렬.
_CENTER_COLS = {1, 2, 3, 10}       # NO, 주문인명, 수령인명, 주문수량

# 헤더에 연한 주황 채움을 넣을 열(1-based): 택배사, 송장번호
_HEADER_FILL_COLS = {11, 12}

FONT_NAME = "맑은 고딕"
FONT_SIZE = 10
HEADER_ROW_HEIGHT = 30.75

_THIN = Side(style="thin", color="FF000000")
THIN_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
HEADER_FILL = PatternFill(
    patternType="solid",
    fgColor=Color(theme=5, tint=0.7999816888943144),
)


def _alignment(col_idx: int, *, header: bool) -> Alignment:
    horizontal = "center" if col_idx in _CENTER_COLS else "left"
    # 헤더는 기본 가운데 정렬(전화번호 열 제외)
    if header:
        horizontal = "left" if col_idx in (4, 5) else "center"
    return Alignment(horizontal=horizontal, vertical="center", wrap_text=header)


def write_standard_sheet(ws, rows):
    """``rows`` (표준 필드 dict 리스트)를 표준 서식으로 ``ws`` 에 기록한다."""
    base_font = Font(name=FONT_NAME, size=FONT_SIZE)

    # 헤더 -----------------------------------------------------------------
    for c, (_key, header) in enumerate(STANDARD_COLUMNS, start=1):
        cell = ws.cell(row=1, column=c, value=header)
        cell.font = copy(base_font)
        cell.border = THIN_BORDER
        cell.alignment = _alignment(c, header=True)
        if c in _HEADER_FILL_COLS:
            cell.fill = copy(HEADER_FILL)
    ws.row_dimensions[1].height = HEADER_ROW_HEIGHT

    # 데이터 ----------------------------------------------------------------
    for r, row in enumerate(rows, start=2):
        for c, key in enumerate(FIELD_KEYS, start=1):
            value = row.get(key)
            cell = ws.cell(row=r, column=c, value=value)
            cell.font = copy(base_font)
            cell.border = THIN_BORDER
            cell.alignment = _alignment(c, header=False)
            # 전화번호/우편번호/송장번호는 텍스트로 고정(앞자리 0·서식 손실 방지)
            if key in ("phone", "phone2", "postcode", "invoice"):
                cell.number_format = "@"

    # 열 너비 ---------------------------------------------------------------
    for col_letter, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    return ws
