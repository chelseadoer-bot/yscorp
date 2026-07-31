# -*- coding: utf-8 -*-
"""합성(가짜) 데이터로 변환 파이프라인을 검증한다. (실제 고객정보 미포함)

최종양식 12열: NO·주문인명·주문인핸드폰번호·수령인명·수령인핸드폰번호·우편번호·
주소·배송메세지·상품정보·주문수량·택배사·송장번호
실행:  python -m pytest tests/  또는  python tests/test_convert.py
"""
import os
import sys
import tempfile

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from po_converter import convert_workbook, write_output  # noqa: E402
from po_converter.cleaning import (  # noqa: E402
    clean_address,
    clean_postcode,
    extract_postcode_from_address,
    split_product_qty,
)

FINAL_HEADERS = ["NO", "주문인명", "주문인핸드폰번호", "수령인명", "수령인핸드폰번호",
                 "우편번호", "주소", "배송메세지", "상품정보", "주문수량", "택배사", "송장번호"]


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
    # 우편번호: 하이픈 없는 5자리, 4자리는 앞 0 복원
    assert clean_postcode("492--08") == "49208"
    assert clean_postcode(6582) == "06582"
    assert clean_postcode("38409") == "38409"
    # 상품명 끝 수량 분리
    assert split_product_qty("테스트 퓨레 ▶ 기본/6개") == ("테스트 퓨레 ▶ 기본", 6)
    assert split_product_qty("y 칡즙 /1개") == ("y 칡즙", 1)
    # 주소 앞 [우편번호] 분리 / 끝 이름·전화 제거
    post, addr = extract_postcode_from_address("[509-51] 경남 김해시 ...")
    assert post == "509-51" and addr.startswith("경남")
    assert "홍길동" not in clean_address("서울 어딘구 1  홍길동 010~1111~2222", strip_trailing_contact=True)
    print("  ✓ cleaning helpers")


def test_unier():
    p = _make(
        os.path.join(tempfile.gettempdir(), "_t_unier_260730.xlsx"),
        ["", "수신인", "연락처1", "연락처2", "우편번호", "주  소", "상품코드",
         "상품명", "속성1 속성2", "발주업체 업체연락처", "실수량", "", "",
         "배송메모", "송장번호", "택배사명", "", "업체연락처",
         "★식별번호-유니어(삭제마세요)"],
        [["유니어", "홍길동", "010-1111-2222", "010-1111-2222", "082--60",
          "서울시 어딘구 어딘로 1  홍길동 010~1111~2222", "H1", "테스트상품 3박스",
          " ", "비즈제이 02-000-0000", 2, "", "", "문앞", "1111-2222-3333",
          "CJ대한통운", "", "02-000-0000", "a032026-07-300001"]],
    )
    r = convert_workbook(p)
    assert r.vendor.key == "unier" and r.vendor_name == "유니어" and r.date == "260730"
    row = r.rows[0]
    assert row["orderer"] == "" and row["orderer_phone"] == ""   # 유니어는 주문인 없음
    assert row["recipient"] == "홍길동" and row["recipient_phone"] == "010-1111-2222"
    assert row["postcode"] == "08260"                            # 하이픈 없는 5자리
    assert "홍길동 010" not in row["address"]
    assert row["product"] == "테스트상품 3박스" and row["qty"] == 2
    assert row["courier"] == "CJ대한통운" and row["invoice"] == "1111-2222-3333"
    print("  ✓ 유니어")


def test_blueberry():
    p = _make(
        os.path.join(tempfile.gettempdir(), "_t_blue_260730.xlsx"),
        ["수령인명", "수령인연락처", "주소", "", "배송시 요청사항"],
        [["김철수", "010-3333-4444", "부산시 어딘구 2", "테스트 퓨레 ▶ 기본/6개", "문 앞"]],
    )
    r = convert_workbook(p)
    assert r.vendor.key == "blueberry" and r.name_confident is False
    row = r.rows[0]
    assert row["recipient"] == "김철수" and row["orderer"] == ""
    assert row["product"] == "테스트 퓨레 ▶ 기본" and row["qty"] == 6   # 수량 분리
    print("  ✓ 블루베리")


def test_chikjeup():
    p = _make(
        os.path.join(tempfile.gettempdir(), "_t_chik_260730.xlsx"),
        ["수령인", "수령인연락처", "주소", "주문상품명", "비고"],
        [["이영희", "0502-5555-6666", "대구시 어딘구 3", "테스트 칡즙 /1개", "대문 앞"]],
    )
    r = convert_workbook(p)
    row = r.rows[0]
    assert r.vendor.key == "chikjeup"
    assert row["recipient"] == "이영희" and row["product"] == "테스트 칡즙" and row["qty"] == 1
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
    assert r.vendor.key == "fashiongeo" and r.vendor_name == "테스트업체" and r.date == "260729"
    row = r.rows[0]
    assert row["orderer"] == "박주문" and row["recipient"] == "박수취"
    assert row["postcode"] == "50951"                 # [509-51] → 하이픈 없는 5자리
    assert row["address"].startswith("경남")
    assert row["qty"] == 2 and row["courier"] == "CJ대한통운"
    print("  ✓ 패션지오")


def test_daon():
    p = _make(
        os.path.join(tempfile.gettempdir(), "_t_daon.xlsx"),
        ["상품명", "수량", "주문인", "받는인", "주문인연락처", "주문인핸드폰",
         "받는인연락처", "받는인핸드폰", "우편", "배송지", "전언", "쇼핑몰",
         "택배사", "송장번호"],
        [["김소형원방의 호박팥차 (1gX100티백)", 1, "박주문", "박영미",
          "010-0000-0000", "010-0000-0000", "010-5536-2123", "010-5536-2123",
          "38409", "경상북도 경산시 하양읍 서사리 278", "문 앞", "스마트스토어",
          "CJ대한통운", "6993-8654-5884"]],
        sheet_title="주문내역",
    )
    p2 = os.path.join(tempfile.gettempdir(), "주식회사다온에프앤씨_주와이에스_20260730.xlsx")
    os.replace(p, p2)
    r = convert_workbook(p2)
    assert r.vendor.key == "daon"
    assert r.vendor_name == "다온에프앤씨"       # 파일명에서 추출
    assert r.date == "260730"
    row = r.rows[0]
    assert row["orderer"] == "박주문" and row["orderer_phone"] == "010-0000-0000"
    assert row["recipient"] == "박영미" and row["recipient_phone"] == "010-5536-2123"
    assert row["postcode"] == "38409" and row["address"].startswith("경상북도")
    assert row["message"] == "문 앞"              # 전언 → 배송메세지
    assert row["product"].startswith("김소형원방") and row["qty"] == 1
    assert row["courier"] == "CJ대한통운" and row["invoice"] == "6993-8654-5884"
    print("  ✓ 다온에프앤씨")


def test_ys_already_final():
    # 이미 최종양식(12열)인 파일. 헤더명이 조금 달라도(주문인핸드폰) 인식.
    p = _make(
        os.path.join(tempfile.gettempdir(), "_t_ys.xlsx"),
        ["NO", "주문인명", "주문인핸드폰", "수령인명", "수령인핸드폰번호", "우편번호",
         "주소", "배송메세지", "상품정보", "주문수량", "택배사", "송장번호"],
        [[1, "", "", "박영미", "010-5536-2123", "38409", "경상북도 경산시 하양읍 서사리 278",
          "문 앞", "김소형원방의 호박팥차 (1gX100티백)", 1, "CJ대한통운", "6993-8654-5884"]],
    )
    p2 = os.path.join(tempfile.gettempdir(), "ys_미페마발주_0730.xlsx")
    os.replace(p, p2)
    r = convert_workbook(p2)
    assert r.vendor.key == "ys"
    assert r.vendor_name == "미페마"      # 파일명에서 추출
    assert r.date == "260730"             # 0730(MMDD) → 올해 YY 붙임
    row = r.rows[0]
    assert row["orderer"] == "" and row["recipient"] == "박영미"
    assert row["recipient_phone"] == "010-5536-2123"       # 주문인핸드폰(짧은 헤더) 인식
    assert row["postcode"] == "38409" and row["invoice"] == "6993-8654-5884"
    print("  ✓ ys(최종양식 통과)")


def test_output_format():
    p = _make(
        os.path.join(tempfile.gettempdir(), "_t_chik2_260730.xlsx"),
        ["수령인", "수령인연락처", "주소", "주문상품명", "비고"],
        [["이영희", "0502-5555-6666", "대구시 3", "테스트 칡즙 /1개", "문 앞"]],
    )
    r = convert_workbook(p)
    out = os.path.join(tempfile.gettempdir(), "_t_out.xlsx")
    write_output(r.rows, out)
    ws = openpyxl.load_workbook(out).active
    assert [ws.cell(1, c).value for c in range(1, 13)] == FINAL_HEADERS
    assert ws["A1"].font.name == "맑은 고딕" and ws["A1"].font.size == 10
    assert ws["A2"].value == 1     # NO 자동 채번
    print("  ✓ 출력 양식(12열/글꼴/번호)")


if __name__ == "__main__":
    fns = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    print(f"테스트 {len(fns)}개 실행")
    for fn in fns:
        fn()
    print("모든 테스트 통과 ✅")
