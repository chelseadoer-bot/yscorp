# -*- coding: utf-8 -*-
"""합성(가짜) 데이터로 변환 파이프라인을 검증한다. (실제 고객정보 미포함)

표준 5열: 수령인명 · 수령인연락처 · 주소 · 주문상품명 · 배송시 요청사항
실행:  python -m pytest tests/  또는  python tests/test_convert.py
"""
import os
import sys
import tempfile

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from po_converter import convert_workbook, write_output  # noqa: E402
from po_converter.cleaning import (  # noqa: E402
    append_qty,
    clean_address,
    extract_postcode_from_address,
)


def _make(path, headers, rows, sheet_title="Sheet1"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title
    for c, h in enumerate(headers, start=1):
        ws.cell(1, c, h)
    for r, row in enumerate(rows, start=2):
        for c, v in enumerate(row, start=1):
            ws.cell(r, c, v)
    wb.save(path)
    return path


def test_cleaning_helpers():
    # 주소 끝 "이름 전화번호" 제거 (이중공백 구분)
    assert "홍길동" not in clean_address("서울 어딘구 1  홍길동 010~1111~2222", strip_trailing_contact=True)
    # ~ → -
    assert clean_address("서울 101~1~2호") == "서울 101-1-2호"
    # 주소 앞 [우편번호] 분리
    post, addr = extract_postcode_from_address("[509-51] 경남 김해시 ...")
    assert post == "509-51" and addr.startswith("경남")
    # 수량 2 이상이면 상품명 끝에 /N개
    assert append_qty("카무트 3박스", 2) == "카무트 3박스 /2개"
    assert append_qty("카무트 3박스", 1) == "카무트 3박스"   # 1개는 붙이지 않음
    print("  ✓ cleaning helpers")


def test_unier():
    p = _make(
        os.path.join(tempfile.gettempdir(), "_t_unier_260730.xlsx"),
        ["", "수신인", "연락처1", "연락처2", "우편번호", "주  소", "상품코드",
         "상품명", "속성1 속성2", "발주업체 업체연락처", "실수량", "", "",
         "배송메모", "송장번호", "택배사명", "", "업체연락처",
         "★식별번호-유니어(삭제마세요)"],
        [["유니어", "홍길동", "010-1111-2222", "010-1111-2222", "54136",
          "서울시 어딘구 어딘로 1  홍길동 010~1111~2222", "H1", "테스트상품 3박스",
          " ", "비즈제이 02-000-0000", 2, "", "", "문앞", "1111-2222-3333",
          "CJ대한통운", "", "02-000-0000", "a032026-07-300001"]],
    )
    r = convert_workbook(p)
    assert r.vendor.key == "unier" and r.vendor_name == "유니어"
    assert r.date == "260730"   # 파일명에서 날짜 추출
    row = r.rows[0]
    assert row["recipient"] == "홍길동"
    assert row["phone"] == "010-1111-2222"
    assert "홍길동 010" not in row["address"]        # 끝 이름·전화 제거
    assert row["product"] == "테스트상품 3박스 /2개"   # 실수량 2 → /2개
    assert row["message"] == "문앞"
    assert set(row.keys()) == {"recipient", "phone", "address", "product", "message"}
    print("  ✓ 유니어")


def test_blueberry():
    p = _make(
        os.path.join(tempfile.gettempdir(), "_t_blue_260730.xlsx"),
        ["수령인명", "수령인연락처", "주소", "", "배송시 요청사항"],
        [["김철수", "010-3333-4444", "부산시 어딘구 어딘로 2", "테스트 퓨레 ▶ 기본/6개", "문 앞"]],
    )
    r = convert_workbook(p)
    assert r.vendor.key == "blueberry" and r.name_confident is False
    row = r.rows[0]
    assert row["recipient"] == "김철수"
    assert row["product"] == "테스트 퓨레 ▶ 기본/6개"   # 수량 표기 그대로 유지
    assert row["message"] == "문 앞"
    print("  ✓ 블루베리")


def test_chikjeup():
    p = _make(
        os.path.join(tempfile.gettempdir(), "_t_chik_260730.xlsx"),
        ["수령인", "수령인연락처", "주소", "주문상품명", "비고"],
        [["이영희", "0502-5555-6666", "대구시 어딘구 어딘로 3", "테스트 칡즙 /1개", "대문 앞"]],
    )
    r = convert_workbook(p)
    assert r.vendor.key == "chikjeup"
    row = r.rows[0]
    assert row["recipient"] == "이영희"
    assert row["product"] == "테스트 칡즙 /1개"   # 그대로 유지
    assert row["message"] == "대문 앞"
    print("  ✓ 칡즙")


def test_fashiongeo():
    p = _make(
        os.path.join(tempfile.gettempdir(), "_t_fg.xlsx"),
        ["주문일자", "수집일자", "주문번호(쇼핑몰)", "주문번호(사방넷)", "품번코드",
         "품목명", "수량", "주문자명", "주문자 전화번호", "수취인 명",
         "수취인 전화번호", "수취인 휴대폰번호", "배송지", "배송메세지", "택배사",
         "송장번호", "주문상태", "물류메세지"],
        [["20260729", "20260729", "1", 2, "C1", "테스트 파우더", 2, "박주문",
          "010-7777-8888", "박수취", "010-7777-8888", "010-7777-8888",
          "[509-51] 경남 김해시 어딘로 4", "문 앞", "CJ대한통운", "9999-0000-1111",
          "신규주문", ""]],
        sheet_title="20260729_테스트업체(주) 발주 건",
    )
    r = convert_workbook(p)
    assert r.vendor.key == "fashiongeo" and r.vendor_name == "테스트업체"
    assert r.date == "260729"   # 데이터 주문일자 우선
    row = r.rows[0]
    assert row["recipient"] == "박수취"
    assert row["address"] == "경남 김해시 어딘로 4"          # [509-51] 제거
    assert row["product"] == "테스트 파우더 /2개"            # 수량 2 → /2개
    print("  ✓ 패션지오")


def test_output_format():
    p = _make(
        os.path.join(tempfile.gettempdir(), "_t_chik2_260730.xlsx"),
        ["수령인", "수령인연락처", "주소", "주문상품명", "비고"],
        [["이영희", "0502-5555-6666", "대구시 어딘구 3", "테스트 칡즙 /1개", "문 앞"]],
    )
    r = convert_workbook(p)
    out = os.path.join(tempfile.gettempdir(), "_t_out.xlsx")
    write_output(r.rows, out)
    ws = openpyxl.load_workbook(out).active
    headers = [ws.cell(1, c).value for c in range(1, 6)]
    assert headers == ["수령인명", "수령인연락처", "주소", "주문상품명", "배송시 요청사항"]
    assert ws["A1"].font.name == "맑은 고딕" and ws["A1"].font.size == 11
    assert ws["A2"].value == "이영희"
    print("  ✓ 출력 양식(헤더/글꼴 11pt)")


if __name__ == "__main__":
    fns = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    print(f"테스트 {len(fns)}개 실행")
    for fn in fns:
        fn()
    print("모든 테스트 통과 ✅")
