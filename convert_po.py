#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""발주양식 통합 변환기 (CLI).

사용 예:
    python convert_po.py 원본1.xlsx 원본2.xlsx
    python convert_po.py ./inputs/*.xlsx -o ./outputs
    python convert_po.py 원본.xlsx --vendor 업체명 --date 260730

각 원본 파일을 표준 발주 양식으로 변환하여
'날짜_업체명_발주 수정본.xlsx' 로 저장한다.
"""
from __future__ import annotations

import argparse
import datetime
import sys
from pathlib import Path

from po_converter import build_output_filename, convert_workbook, write_output


def _today_yymmdd() -> str:
    return datetime.date.today().strftime("%y%m%d")


def process_file(path: Path, out_dir: Path, *, date=None, vendor=None, quiet=False) -> Path | None:
    try:
        result = convert_workbook(str(path))
    except Exception as exc:  # noqa: BLE001
        print(f"  ✗ {path.name}: {exc}", file=sys.stderr)
        return None

    vendor_name = vendor or result.vendor_name
    file_date = date or result.date or _today_yymmdd()

    out_name = build_output_filename(file_date, vendor_name)
    out_path = out_dir / out_name
    write_output(result.rows, str(out_path), sheet_title=f"{file_date}_{vendor_name}")

    if not quiet:
        flag = "" if (vendor or result.name_confident) else "  ⚠ 업체명 자동추정(확인 필요)"
        print(f"  ✓ {path.name}")
        print(f"      → {out_name}  ({len(result.rows)}건, 양식='{result.vendor.key}'){flag}")
    return out_path


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="여러 업체의 발주 엑셀을 표준 양식으로 통합 변환합니다.",
    )
    parser.add_argument("inputs", nargs="+", help="변환할 원본 xlsx 파일들")
    parser.add_argument("-o", "--out-dir", default=".", help="출력 폴더 (기본: 현재 폴더)")
    parser.add_argument("--date", help="파일명 날짜 YYMMDD (기본: 원본에서 추출 또는 오늘)")
    parser.add_argument("--vendor", help="업체명 강제 지정 (기본: 자동 감지)")
    args = parser.parse_args(argv)

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # glob 미확장 셸 대응
    paths: list[Path] = []
    for pattern in args.inputs:
        p = Path(pattern)
        if p.exists():
            paths.append(p)
        else:
            paths.extend(sorted(Path().glob(pattern)))

    if not paths:
        print("변환할 파일을 찾을 수 없습니다.", file=sys.stderr)
        return 1

    print(f"발주양식 통합 변환 시작 — {len(paths)}개 파일")
    ok = 0
    for path in paths:
        if process_file(path, out_dir, date=args.date, vendor=args.vendor):
            ok += 1
    print(f"완료: {ok}/{len(paths)}개 변환  →  {out_dir.resolve()}")
    return 0 if ok == len(paths) else 2


if __name__ == "__main__":
    raise SystemExit(main())
