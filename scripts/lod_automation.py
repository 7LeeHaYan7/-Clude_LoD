"""
LoD 실험 데이터 자동 집계 스크립트

폴더 구조:
  <ROOT>/5000cp/1_1 .. 16_1/*FinalResult*.xls(x)
  <ROOT>/1670cp/1_1 .. 16_1/*FinalResult*.xls(x)
  <ROOT>/556cp/1_1 .. 16_1/*FinalResult*.xls(x)
  <ROOT>/185cp/1_1 .. 16_1/*FinalResult*.xls(x)
  <ROOT>/62cp/1_1 .. 16_1/*FinalResult*.xls(x)
  <ROOT>/21cp/1_1 .. 16_1/*FinalResult*.xls(x)

각 FinalResult 파일의 B18~B21 값을 확인하여 Valid인 행에서 지정된 셀을
읽어 타겟 LoD 템플릿(xlsx)의 해당 병원체 시트, 해당 행/열에 기록한다.

원본 raw 파일과 원본 템플릿 파일은 절대 수정하지 않는다.
결과는 항상 템플릿을 복사한 새 파일에 저장하고, 처리 내역은 별도 로그
엑셀로 남긴다.
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import openpyxl

try:
    import xlrd
except ImportError:  # pragma: no cover
    xlrd = None


TOP_FOLDERS_ORDER = ["5000cp", "1670cp", "556cp", "185cp", "62cp", "21cp"]

# 상위 폴더 -> 타겟 시트에서 Ct 값을 쓸 열
CT_COLUMN = {
    "5000cp": "D",
    "1670cp": "F",
    "556cp": "H",
    "185cp": "J",
    "62cp": "L",
    "21cp": "N",
}

SUBFOLDER_COUNT = 16  # 1_1 ~ 16_1
FIRST_DATA_ROW = 6  # No.1 이 위치한 타겟 시트 행

# (FinalResult 파일의 B{row} 검사행, [(값을 읽을 열, 타겟 시트 병원체명), ...])
EXTRACTION_RULES: list[tuple[int, list[tuple[str, str]]]] = [
    (18, [("F", "CPA"), ("E", "CPE")]),
    (19, [("C", "Giardia"), ("D", "CPV2")]),
    (20, [("C", "Campylobacter"), ("E", "Salmonella")]),
    (21, [("C", "CECoV")]),
]

TARGET_SHEET_PREFIX = "260909_"

SUBFOLDER_RE = re.compile(r"^(\d+)_1$")


def col_letter_to_index(letter: str) -> int:
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
    return idx


class SourceWorkbook:
    """FinalResult 원본 파일(.xls/.xlsx)을 읽기 전용으로 여는 래퍼. 원본은 절대 쓰지 않는다."""

    def __init__(self, path: Path):
        self.path = path
        ext = path.suffix.lower()
        if ext == ".xls":
            if xlrd is None:
                raise RuntimeError("xlrd 패키지가 필요합니다 (pip install xlrd)")
            self._book = xlrd.open_workbook(str(path), formatting_info=False)
            self._sheet = self._book.sheet_by_index(0)
            self._mode = "xls"
        elif ext in (".xlsx", ".xlsm"):
            self._book = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
            self._sheet = self._book.worksheets[0]
            self._mode = "xlsx"
        else:
            raise ValueError(f"지원하지 않는 원본 파일 형식: {path.suffix}")

    def cell(self, col_letter: str, row: int):
        if self._mode == "xls":
            r, c = row - 1, col_letter_to_index(col_letter) - 1
            if r >= self._sheet.nrows or c >= self._sheet.ncols:
                return None
            return self._sheet.cell_value(r, c)
        return self._sheet[f"{col_letter}{row}"].value

    def close(self):
        if self._mode == "xlsx":
            self._book.close()


def clean_value(raw):
    """앞뒤 공백 제거 후 숫자로 변환 가능하면 float, 아니면(UD 등) 문자열 그대로."""
    if raw is None:
        return None
    s = str(raw).strip()
    if s == "":
        return None
    try:
        return float(s)
    except ValueError:
        return s


@dataclass
class LogRow:
    top_folder: str
    sub_folder: str
    check_row: str = ""
    judgement: str = ""
    source_cell: str = ""
    raw_value: object = ""
    written_value: object = ""
    target_sheet: str = ""
    target_cell: str = ""
    status: str = ""
    message: str = ""

    def as_dict(self):
        return {
            "상위폴더": self.top_folder,
            "하위폴더": self.sub_folder,
            "검사행": self.check_row,
            "판정": self.judgement,
            "소스셀": self.source_cell,
            "원본값": self.raw_value,
            "기록값": self.written_value,
            "타겟시트": self.target_sheet,
            "타겟셀": self.target_cell,
            "상태": self.status,
            "메시지": self.message,
        }


def find_final_result_files(sub_dir: Path) -> list[Path]:
    return sorted(
        p for p in sub_dir.iterdir()
        if p.is_file() and "finalresult" in p.name.lower()
    )


def find_subfolders(top_dir: Path, n: int) -> list[Path]:
    """폴더명이 정확히 'N_1'이거나 끝이 'N_1'로 끝나는 폴더를 찾는다.
    (예: 'BBB014_..._5000cp-15_1' 도 n=15 대상으로 매칭됨)
    11_1이 n=1의 '1_1'과 잘못 매칭되지 않도록, 숫자 앞 글자가 숫자면 제외한다."""
    suffix = f"{n}_1"
    matches = []
    for d in sorted(top_dir.iterdir()):
        if not d.is_dir():
            continue
        if d.name == suffix:
            matches.append(d)
            continue
        if d.name.endswith(suffix):
            boundary_idx = len(d.name) - len(suffix)
            if boundary_idx == 0 or not d.name[boundary_idx - 1].isdigit():
                matches.append(d)
    return matches


def process(root: Path, template_path: Path, out_path: Path) -> list[LogRow]:
    logs: list[LogRow] = []

    wb = openpyxl.load_workbook(template_path)  # 서식/수식 보존을 위해 복사본에 직접 로드

    for top in TOP_FOLDERS_ORDER:
        top_dir = root / top
        ct_col = CT_COLUMN[top]

        if not top_dir.is_dir():
            logs.append(LogRow(top, "", status="오류", message="상위 폴더를 찾을 수 없음"))
            continue

        for n in range(1, SUBFOLDER_COUNT + 1):
            sub_name = f"{n}_1"
            target_row = FIRST_DATA_ROW + (n - 1)

            sub_dirs = find_subfolders(top_dir, n)
            if len(sub_dirs) == 0:
                logs.append(LogRow(top, sub_name, status="오류", message="하위 폴더를 찾을 수 없음"))
                continue
            if len(sub_dirs) > 1:
                names = ", ".join(d.name for d in sub_dirs)
                logs.append(LogRow(top, sub_name, status="오류",
                                    message=f"'{sub_name}'로 끝나는 하위 폴더가 {len(sub_dirs)}개 발견됨: {names}"))
                continue
            sub_dir = sub_dirs[0]

            matches = find_final_result_files(sub_dir)
            if len(matches) == 0:
                logs.append(LogRow(top, sub_name, status="오류", message="FinalResult 파일을 찾을 수 없음"))
                continue
            if len(matches) > 1:
                names = ", ".join(m.name for m in matches)
                logs.append(LogRow(top, sub_name, status="오류",
                                    message=f"FinalResult 파일이 {len(matches)}개 발견됨: {names}"))
                continue

            src_path = matches[0]
            try:
                src = SourceWorkbook(src_path)
            except Exception as e:
                logs.append(LogRow(top, sub_name, status="오류", message=f"파일 열기 실패: {e}"))
                continue

            for row, extractions in EXTRACTION_RULES:
                b_val = src.cell("B", row)
                status = str(b_val).strip() if b_val is not None else ""

                if status.lower() == "valid":
                    for col, pathogen in extractions:
                        raw_val = src.cell(col, row)
                        value = clean_value(raw_val)
                        target_sheet_name = f"{TARGET_SHEET_PREFIX}{pathogen}"

                        if target_sheet_name not in wb.sheetnames:
                            logs.append(LogRow(
                                top, sub_name, check_row=f"B{row}", judgement="Valid",
                                source_cell=f"{col}{row}", raw_value=raw_val,
                                status="오류", message=f"타겟 시트 '{target_sheet_name}' 없음",
                            ))
                            continue

                        ws = wb[target_sheet_name]
                        target_cell = f"{ct_col}{target_row}"
                        ws[target_cell] = value

                        logs.append(LogRow(
                            top, sub_name, check_row=f"B{row}", judgement="Valid",
                            source_cell=f"{col}{row}", raw_value=raw_val, written_value=value,
                            target_sheet=target_sheet_name, target_cell=target_cell,
                            status="기록완료",
                        ))

                elif status.lower() == "invalid":
                    logs.append(LogRow(
                        top, sub_name, check_row=f"B{row}", judgement="Invalid",
                        status="건너뜀",
                    ))
                else:
                    logs.append(LogRow(
                        top, sub_name, check_row=f"B{row}", judgement=status or "(빈 값)",
                        status="경고", message="Valid/Invalid가 아닌 값 - 건너뜀",
                    ))

            src.close()

    wb.save(out_path)
    return logs


def write_log(logs: list[LogRow], log_path: Path):
    log_wb = openpyxl.Workbook()
    ws = log_wb.active
    ws.title = "처리내역"

    headers = ["상위폴더", "하위폴더", "검사행", "판정", "소스셀", "원본값",
               "기록값", "타겟시트", "타겟셀", "상태", "메시지"]
    ws.append(headers)
    for log in logs:
        d = log.as_dict()
        ws.append([d[h] for h in headers])

    for col_cells in ws.columns:
        length = max((len(str(c.value)) for c in col_cells if c.value is not None), default=8)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max(length + 2, 10), 40)

    # 요약 시트
    summary = log_wb.create_sheet("요약")
    total = len(logs)
    written = sum(1 for l in logs if l.status == "기록완료")
    skipped = sum(1 for l in logs if l.status == "건너뜀")
    errors = [l for l in logs if l.status == "오류"]
    warnings = [l for l in logs if l.status == "경고"]

    summary.append(["항목", "값"])
    summary.append(["총 로그 행 수", total])
    summary.append(["기록 완료", written])
    summary.append(["Invalid로 건너뜀", skipped])
    summary.append(["오류 건수", len(errors)])
    summary.append(["경고 건수", len(warnings)])
    summary.append([])
    summary.append(["오류 목록"])
    for l in errors:
        summary.append([f"{l.top_folder}/{l.sub_folder}", l.message])
    summary.append([])
    summary.append(["경고 목록"])
    for l in warnings:
        summary.append([f"{l.top_folder}/{l.sub_folder} {l.check_row}", l.message])

    log_wb.save(log_path)


def main():
    parser = argparse.ArgumentParser(description="LoD FinalResult 데이터를 타겟 엑셀로 자동 집계")
    parser.add_argument("--root", required=True, help="1st 폴더 경로 (상위 폴더 5000cp 등이 들어있는 경로)")
    parser.add_argument("--template", required=True, help="타겟 LoD 템플릿 xlsx 경로 (읽기 전용, 수정되지 않음)")
    parser.add_argument("--out", help="결과를 저장할 새 파일 경로 (기본: 템플릿명_결과_타임스탬프.xlsx)")
    parser.add_argument("--log", help="처리 로그를 저장할 경로 (기본: 결과파일명_로그.xlsx)")
    args = parser.parse_args()

    root = Path(args.root)
    template_path = Path(args.template)

    if not root.is_dir():
        sys.exit(f"오류: ROOT 폴더를 찾을 수 없습니다: {root}")
    if not template_path.is_file():
        sys.exit(f"오류: 템플릿 파일을 찾을 수 없습니다: {template_path}")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    if args.out:
        out_path = Path(args.out)
    else:
        out_path = template_path.with_name(f"{template_path.stem}_결과_{timestamp}{template_path.suffix}")

    if args.log:
        log_path = Path(args.log)
    else:
        log_path = out_path.with_name(f"{out_path.stem}_로그.xlsx")

    # 템플릿은 절대 직접 열어서 쓰지 않고, 항상 먼저 새 경로로 복사한 뒤 그 사본을 채운다.
    shutil.copy2(template_path, out_path)

    logs = process(root, out_path, out_path)

    write_log(logs, log_path)

    written = sum(1 for l in logs if l.status == "기록완료")
    skipped = sum(1 for l in logs if l.status == "건너뜀")
    errors = sum(1 for l in logs if l.status == "오류")
    warnings = sum(1 for l in logs if l.status == "경고")

    print(f"완료: 결과 파일 -> {out_path}")
    print(f"      로그 파일 -> {log_path}")
    print(f"      기록 {written}건 / Invalid 건너뜀 {skipped}건 / 오류 {errors}건 / 경고 {warnings}건")
    if errors:
        print("오류가 발생했습니다. 로그 파일의 '요약' 시트를 확인하세요.")


if __name__ == "__main__":
    main()
