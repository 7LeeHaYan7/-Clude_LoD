"""
LoD 실험 데이터 자동 집계 스크립트 (범용 버전)

lod_automation_Canine-Ent.py / lod_automation_Canine-Res.py 와 같은 파이프라인(폴더 탐색 ->
Valid 판정 -> Ct 기록 -> 기울기/상수 -> Probit LoD -> 그래프)을 그대로 쓰되, 아래 두 가지를
하드코딩 대신 실제 입력을 보고 동적으로 결정한다.

1. 농도(폴더명) 자동 인식
   <ROOT> 바로 아래의 폴더들(예: 5000cp, 556cp, 185cp, 7cp, 1670copies ...)에서 숫자를
   추출해 큰 값부터 내림차순으로 정렬한다. 농도 단계 개수나 구체적인 숫자를 코드에 미리
   적어두지 않는다. 각 타겟 시트의 5행에서 'Conc.'/'Ct' 가 번갈아 나오는 열 쌍을 왼쪽부터
   순서대로 찾아서, 내림차순 농도 목록과 1:1로 매칭한다 (가장 큰 농도 -> 가장 왼쪽 쌍).
   매칭된 Conc. 열에는 실제 숫자를, Ct 열에는 FinalResult에서 읽은 값을 쓴다.
   표준곡선 회귀에 쓰는 R/S/T(농도/Log/평균Ct) 작은 표도 농도 단계 수에 맞춰 직접 써서,
   기존처럼 시트 어딘가의 AVG 행을 참조하는 간접 수식에 의존하지 않는다.

2. 타겟(병원체) 매핑: <ROOT>/Target_input.xlsx
   1번째 시트에 'Target'/'Data' 두 열이 있고, 각 행이 "병원체명, FinalResult에서 읽을 셀
   주소(예: F18)"를 나타낸다. 셀 주소의 행 번호가 곧 FinalResult의 B{row} 검사행이 되므로
   (예: F18, E18 -> 18행이 Valid인지로 같이 판정), 같은 행 번호를 쓰는 항목들은 자동으로
   하나의 '검사행' 그룹으로 묶인다. 타겟 시트 이름은 병원체명과 동일(접두사 없음)하다고
   가정한다.

원본 raw 파일과 원본 템플릿 파일은 절대 수정하지 않는다. 결과는 항상 템플릿을 복사한 새
파일에 저장하고, 처리 내역은 별도 로그 엑셀로 남긴다.
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
import tempfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter, column_index_from_string

try:
    import xlrd
except ImportError:  # pragma: no cover
    xlrd = None

try:
    import numpy as np
    import statsmodels.api as sm
    from scipy.stats import norm
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    PROBIT_LIBS_AVAILABLE = True
except ImportError:  # pragma: no cover
    PROBIT_LIBS_AVAILABLE = False


# ---- 타겟 시트 구조: 고정된 절대 셀/열 (병원체 수·농도 수와 무관) ----
FIRST_DATA_ROW = 6  # No.1 이 위치한 타겟 시트 행
TARGET_SHEET_PREFIX = ""

SLOPE_CELL = "AD18"
INTERCEPT_CELL = "AD19"
REGRESSION_CONC_COL = "R"   # 작은 요약표: 농도
REGRESSION_LOG_COL = "S"    # 작은 요약표: Log(농도)
REGRESSION_CTAVG_COL = "T"  # 작은 요약표: 평균 Ct
REGRESSION_FIRST_ROW = FIRST_DATA_ROW  # 6행부터 농도 단계 수만큼 사용

LOD_LOGX_CELL = "AC22"
LOD_CI_LOW_CELL = "AC23"
LOD_CI_HIGH_CELL = "AC24"
LOD_TARGET_P = 0.95
LOD_PLOT_ANCHOR = "R40"

TARGET_INPUT_FILENAME = "Target_input.xlsx"

RETEST_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

TRAILING_NUMBER_RE = re.compile(r"(\d+)_1$")
LEADING_NUMBER_RE = re.compile(r"(\d+(?:\.\d+)?)")
CELL_ADDRESS_RE = re.compile(r"^\s*([A-Za-z]+)\s*(\d+)\s*$")


def col_letter_to_index(letter: str) -> int:
    return column_index_from_string(letter)


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
    previous_value: object = ""
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
            "기존기록값": self.previous_value,
            "기록값": self.written_value,
            "타겟시트": self.target_sheet,
            "타겟셀": self.target_cell,
            "상태": self.status,
            "메시지": self.message,
        }


# ======================================================================
# 1. 농도 폴더 자동 인식
# ======================================================================

def parse_concentration_value(folder_name: str) -> float | None:
    """폴더명에서 숫자를 추출한다. '5000cp'->5000.0, '556.5cp'->556.5, '7copies'->7.0.
    숫자를 찾을 수 없으면 None."""
    m = LEADING_NUMBER_RE.search(folder_name)
    if m is None:
        return None
    return float(m.group(1))


@dataclass
class ConcentrationFolder:
    name: str
    value: float
    path: Path


def discover_concentration_folders(root: Path) -> tuple[list[ConcentrationFolder], list[str]]:
    """root 바로 아래의 폴더들에서 농도를 추출해 큰 값부터 내림차순으로 정렬해 반환한다.
    두 번째 반환값은 숫자를 추출하지 못해 농도 폴더로 보지 않고 건너뛴 폴더명 목록."""
    folders: list[ConcentrationFolder] = []
    skipped: list[str] = []
    seen_values: dict[float, str] = {}

    for d in sorted(p for p in root.iterdir() if p.is_dir()):
        value = parse_concentration_value(d.name)
        if value is None:
            skipped.append(d.name)
            continue
        if value in seen_values:
            raise ValueError(
                f"농도 폴더 '{d.name}'과 '{seen_values[value]}'이(가) 같은 농도값({value})으로 "
                "해석됩니다 - 폴더명을 확인해주세요."
            )
        seen_values[value] = d.name
        folders.append(ConcentrationFolder(name=d.name, value=value, path=d))

    folders.sort(key=lambda f: f.value, reverse=True)
    return folders, skipped


# ======================================================================
# 2. Target_input.xlsx 기반 타겟 매핑
# ======================================================================

@dataclass
class TargetEntry:
    pathogen: str
    col: str
    row: int


def load_target_config(path: Path) -> tuple[list[str], list[tuple[int, list[tuple[str, str]]]]]:
    """Target_input.xlsx를 읽어 (PATHOGENS 순서 목록, EXTRACTION_RULES)를 만든다.

    기대하는 구조: 1번째 시트, 1행에 'Target'(병원체명)과 'Data'(FinalResult에서 읽을
    셀 주소, 예 F18) 두 열이 있고, 2행부터 각 병원체 한 줄씩. 셀 주소의 행 번호가 그대로
    FinalResult의 검사행(B{row})이 된다 - 같은 행 번호를 쓰는 항목들은 자동으로 같은
    검사행 그룹으로 묶인다."""
    if not path.is_file():
        raise FileNotFoundError(f"타겟 매핑 파일을 찾을 수 없습니다: {path}")

    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        raise ValueError(f"타겟 매핑 파일을 열 수 없습니다 ({path}): {e}")

    if not wb.worksheets:
        raise ValueError(f"타겟 매핑 파일에 시트가 없습니다: {path}")
    ws = wb.worksheets[0]

    header_row = 1
    target_col = None
    data_col = None
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is None:
            continue
        key = str(v).strip().lower()
        # 정확히 'target'/'data'가 아니어도 'Ct Data cell'처럼 단어가 포함돼 있으면 인정한다
        # (실제 파일마다 헤더 문구가 조금씩 다른 경우가 있음).
        if target_col is None and "target" in key:
            target_col = c
        elif data_col is None and "data" in key:
            data_col = c
    if target_col is None or data_col is None:
        raise ValueError(
            f"타겟 매핑 파일({path}) 1행에서 'Target'/'Data' 열을 찾을 수 없습니다. "
            f"실제 1행: {[ws.cell(row=header_row, column=c).value for c in range(1, ws.max_column + 1)]}"
        )

    pathogens: list[str] = []
    entries: list[TargetEntry] = []
    for r in range(header_row + 1, ws.max_row + 1):
        pathogen_raw = ws.cell(row=r, column=target_col).value
        data_raw = ws.cell(row=r, column=data_col).value
        if pathogen_raw is None and data_raw is None:
            continue  # 빈 행은 건너뜀
        if pathogen_raw is None or data_raw is None:
            raise ValueError(f"타겟 매핑 파일 {r}행: Target/Data 중 한쪽이 비어 있습니다.")

        pathogen = str(pathogen_raw).strip()
        m = CELL_ADDRESS_RE.match(str(data_raw))
        if not m:
            raise ValueError(
                f"타겟 매핑 파일 {r}행: Data 값 '{data_raw}'을(를) 셀 주소(예: F18)로 "
                "해석할 수 없습니다."
            )
        col = m.group(1).upper()
        row = int(m.group(2))

        if pathogen not in pathogens:
            pathogens.append(pathogen)
        entries.append(TargetEntry(pathogen=pathogen, col=col, row=row))

    if not entries:
        raise ValueError(f"타겟 매핑 파일({path})에 유효한 데이터 행이 없습니다.")

    rules_by_row: dict[int, list[tuple[str, str]]] = {}
    for e in entries:
        rules_by_row.setdefault(e.row, []).append((e.col, e.pathogen))
    extraction_rules = [(row, cols) for row, cols in sorted(rules_by_row.items())]

    return pathogens, extraction_rules


def write_slope_intercept_formulas(wb: openpyxl.Workbook, pathogens: list[str],
                                    regression_x_range: str, regression_y_range: str,
                                    logs: list[LogRow]):
    """각 병원체 시트에 표준곡선 기울기/상수 수식을 입력한다."""
    for pathogen in pathogens:
        sheet_name = f"{TARGET_SHEET_PREFIX}{pathogen}"
        if sheet_name not in wb.sheetnames:
            logs.append(LogRow("", "", target_sheet=sheet_name, status="오류",
                                message="기울기/상수 입력 대상 시트 없음"))
            continue
        ws = wb[sheet_name]
        ws[SLOPE_CELL] = f"=SLOPE({regression_y_range},{regression_x_range})"
        ws[INTERCEPT_CELL] = f"=INTERCEPT({regression_y_range},{regression_x_range})"
        logs.append(LogRow("", "", target_sheet=sheet_name, target_cell=SLOPE_CELL,
                            status="수식입력",
                            message=f"기울기=SLOPE({regression_y_range},{regression_x_range})"))
        logs.append(LogRow("", "", target_sheet=sheet_name, target_cell=INTERCEPT_CELL,
                            status="수식입력",
                            message=f"상수=INTERCEPT({regression_y_range},{regression_x_range})"))


# ======================================================================
# 농도 Conc./Ct 열 쌍 자동 탐지 (타겟 시트 5행)
# ======================================================================

def detect_conc_ct_pairs(ws, header_row: int = 5, max_col: int = 60) -> list[tuple[str, str]]:
    """header_row 행에서 'No.' 바로 다음 칸부터 'Conc.'/'Ct' 가 번갈아 나오는 열 쌍을
    왼쪽부터 순서대로 찾는다. 'Conc.'/'Ct' 가 아닌 값이 나오면 그 앞까지만 반환한다."""
    no_col = None
    for c in range(1, max_col + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is not None and str(v).strip().lower() in ("no.", "no"):
            no_col = c
            break
    start_col = (no_col + 1) if no_col is not None else 3  # 'No.'를 못 찾으면 C열부터 시도

    pairs: list[tuple[str, str]] = []
    c = start_col
    while c + 1 <= max_col:
        v1 = ws.cell(row=header_row, column=c).value
        v2 = ws.cell(row=header_row, column=c + 1).value
        if (v1 is not None and str(v1).strip().lower() == "conc."
                and v2 is not None and str(v2).strip().lower() == "ct"):
            pairs.append((get_column_letter(c), get_column_letter(c + 1)))
            c += 2
        else:
            break
    return pairs


def assign_conc_ct_columns(pairs: list[tuple[str, str]], n_needed: int,
                            sheet_name: str) -> list[tuple[str, str]]:
    if n_needed == 0:
        raise ValueError("사용할 농도 단계가 0개입니다 - 농도 폴더를 찾지 못했는지 확인해주세요.")
    if n_needed > len(pairs):
        raise ValueError(
            f"타겟 시트 '{sheet_name}'의 {5}행에 'Conc./Ct' 열 쌍이 {len(pairs)}개뿐인데, "
            f"농도 폴더는 {n_needed}개 발견됐습니다. 템플릿에 열 쌍을 더 추가해주세요."
        )
    return pairs[:n_needed]


def write_concentration_summary(wb: openpyxl.Workbook, pathogens: list[str],
                                 conc_values_desc: list[float], last_data_row: int,
                                 logs: list[LogRow]) -> tuple[str, str]:
    """각 병원체 시트에 (1) 메인 데이터 표의 Conc. 열에 농도값을 내림차순으로 채우고,
    (2) 표준곡선 회귀용 작은 요약표(R=농도, S=Log(농도), T=해당 Ct열의 평균)를 새로
    써서 기존 AVG 행 위치에 의존하지 않도록 한다. 반환값은 (REGRESSION_X_RANGE, REGRESSION_Y_RANGE)."""
    n = len(conc_values_desc)
    last_summary_row = REGRESSION_FIRST_ROW + n - 1
    regression_x_range = f"{REGRESSION_LOG_COL}{REGRESSION_FIRST_ROW}:{REGRESSION_LOG_COL}{last_summary_row}"
    regression_y_range = f"{REGRESSION_CTAVG_COL}{REGRESSION_FIRST_ROW}:{REGRESSION_CTAVG_COL}{last_summary_row}"

    for pathogen in pathogens:
        sheet_name = f"{TARGET_SHEET_PREFIX}{pathogen}"
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        pairs = detect_conc_ct_pairs(ws)
        try:
            assigned = assign_conc_ct_columns(pairs, n, sheet_name)
        except ValueError as e:
            logs.append(LogRow("", "", target_sheet=sheet_name, status="오류", message=str(e)))
            continue

        for i, (conc_col, ct_col) in enumerate(assigned):
            value = conc_values_desc[i]
            for row in range(FIRST_DATA_ROW, last_data_row + 1):
                ws[f"{conc_col}{row}"] = value

            summary_row = REGRESSION_FIRST_ROW + i
            ws[f"{REGRESSION_CONC_COL}{summary_row}"] = value
            ws[f"{REGRESSION_LOG_COL}{summary_row}"] = f"=LOG({REGRESSION_CONC_COL}{summary_row})"
            ws[f"{REGRESSION_CTAVG_COL}{summary_row}"] = (
                f"=AVERAGE({ct_col}{FIRST_DATA_ROW}:{ct_col}{last_data_row})"
            )
        logs.append(LogRow(
            "", "", target_sheet=sheet_name, status="농도요약작성",
            message=f"{n}개 농도 단계를 {REGRESSION_CONC_COL}{REGRESSION_FIRST_ROW}:"
                    f"{REGRESSION_CONC_COL}{last_summary_row}에 기록"
        ))

    return regression_x_range, regression_y_range


def build_ct_column_map(ws, conc_values_desc: list[float], sheet_name: str) -> dict[float, str]:
    """generate/update 모드: 이미 write_concentration_summary로 Conc.열을 채운 뒤,
    농도값 -> Ct열 매핑을 다시 계산한다(그 사이 코드가 안 바뀌면 detect_conc_ct_pairs 결과는
    동일하다)."""
    pairs = detect_conc_ct_pairs(ws)
    assigned = assign_conc_ct_columns(pairs, len(conc_values_desc), sheet_name)
    return {value: ct_col for value, (_, ct_col) in zip(conc_values_desc, assigned)}


def read_existing_conc_ct_pairs(ws) -> list[tuple[float, str]]:
    """recalc 모드: raw 폴더 없이, 이미 채워진 Conc. 열의 값을 직접 읽어서
    (농도값, Ct열) 목록을 복원한다. Conc. 열이 비어있는 쌍은 건너뛴다."""
    pairs = detect_conc_ct_pairs(ws)
    result: list[tuple[float, str]] = []
    for conc_col, ct_col in pairs:
        raw = ws[f"{conc_col}{FIRST_DATA_ROW}"].value
        value = clean_value(raw)
        if not isinstance(value, (int, float)):
            continue
        result.append((float(value), ct_col))
    return result


# ======================================================================
# Probit LoD (lod_automation_Canine-Ent.py 와 동일한 통계 로직)
# ======================================================================

def read_detection_counts(ws, conc_ct_pairs: list[tuple[float, str]],
                           last_data_row: int) -> list[tuple[float, int, int]]:
    """conc_ct_pairs(농도값, Ct열) 목록을 받아 각 농도별 (농도, 양성수, 시도수)를 구한다.
    빈칸이면 그 폴더는 아예 제외(시도수에서도 제외). 그 외에는 'UD'라는 글자가 포함된 값만
    음성으로 치고 나머지는 전부 양성으로 센다 - 타겟 시트의 Detect 수식(COUNTA -
    COUNTIF("*UD*"))과 동일한 기준이다(숫자 타입인지 여부로 판정하면 오타로 텍스트가 된 Ct
    값이 잘못 빠질 수 있다)."""
    result = []
    for conc, col in conc_ct_pairs:
        positive = 0
        total = 0
        for row in range(FIRST_DATA_ROW, last_data_row + 1):
            val = ws[f"{col}{row}"].value
            if val is None or (isinstance(val, str) and val.strip() == ""):
                continue
            total += 1
            if isinstance(val, str) and "ud" in val.lower():
                continue
            positive += 1
        result.append((conc, positive, total))
    return result


def fit_probit_lod(points: list[tuple[float, int, int]], p: float = LOD_TARGET_P) -> dict:
    """positive ~ log10(conc) 에 대한 probit GLM을 적합하고, MASS::dose.p와 동일한 델타법으로
    p(기본 0.95) 지점의 LogX 추정치와 95% 신뢰구간(하한/상한)을 구한다."""
    pts = [(c, pos, tot) for c, pos, tot in points if tot > 0]
    if len(pts) < 3:
        raise ValueError(f"유효한 농도 그룹이 {len(pts)}개뿐이라 probit 회귀를 할 수 없음(최소 3개 필요)")

    conc = np.array([c for c, _, _ in pts], dtype=float)
    positive = np.array([pos for _, pos, _ in pts], dtype=float)
    total = np.array([tot for _, _, tot in pts], dtype=float)
    log_conc = np.log10(conc)

    X = sm.add_constant(log_conc)
    endog = np.column_stack([positive, total - positive])
    model = sm.GLM(endog, X, family=sm.families.Binomial(link=sm.families.links.Probit()))
    fit_result = model.fit()
    b0, b1 = fit_result.params
    cov = fit_result.cov_params()
    var_b0, var_b1, cov01 = cov[0, 0], cov[1, 1], cov[0, 1]

    def dose_p(prob: float) -> tuple[float, float]:
        z = norm.ppf(prob)
        xp = (z - b0) / b1
        se = np.sqrt((var_b0 + xp ** 2 * var_b1 + 2 * xp * cov01) / b1 ** 2)
        return xp, se

    xp, se = dose_p(p)
    z975 = norm.ppf(0.975)

    return {
        "b0": b0, "b1": b1,
        "dose_p": dose_p,
        "lod_log": xp,
        "ci_low_log": xp - z975 * se,
        "ci_high_log": xp + z975 * se,
        "points_log": [(np.log10(c), pos / tot) for c, pos, tot in pts],
    }


def generate_probit_plot(pathogen: str, fit: dict, out_png_path: Path):
    """R 스크립트와 동일한 스타일(빨간 S커브, 95%선, LoD/CI 수직선, 신뢰구간 밴드)로 플롯을 그린다.
    모든 농도에서 100% 검출되는 등 완전분리(perfect separation)가 일어나면 LoD/CI가 수학적으로
    발산해 천문학적 숫자가 나올 수 있다 - 그 값을 그대로 그리면 matplotlib이 좌표 변환에서 죽으므로,
    선/점/글자 위치는 그래프 범위(X_MIN~X_MAX) 안으로 고정해서 그린다. 발산 여부와 무관하게 실제
    계산값은 그대로 AC22~24 셀에 기록되며, 이 함수는 시각화만 담당한다."""
    X_MIN, X_MAX = -1, 4
    b0, b1 = fit["b0"], fit["b1"]
    dose_p = fit["dose_p"]
    lod_log, ci_low, ci_high = fit["lod_log"], fit["ci_low_log"], fit["ci_high_log"]
    lod_log_plot = float(np.clip(lod_log, X_MIN, X_MAX))
    ci_low_plot = float(np.clip(ci_low, X_MIN, X_MAX))
    ci_high_plot = float(np.clip(ci_high, X_MIN, X_MAX))
    z975 = norm.ppf(0.975)

    fig, ax = plt.subplots(figsize=(7, 5))

    x = np.linspace(X_MIN, X_MAX, 400)
    ax.plot(x, norm.cdf(b0 + b1 * x), color="red", linewidth=1.2)

    for log_conc, prop in fit["points_log"]:
        ax.plot(log_conc, prop, "o", color="red", markersize=5)

    ax.axhline(LOD_TARGET_P, linestyle="--", color="black", linewidth=0.8)
    ax.axvline(lod_log_plot, linestyle="--", color="black", linewidth=0.8)
    ax.axvline(ci_low_plot, linestyle=":", color="gray", linewidth=0.8)
    ax.axvline(ci_high_plot, linestyle=":", color="gray", linewidth=0.8)
    ax.plot(lod_log_plot, LOD_TARGET_P, "s", color="blue", markersize=6)

    ax.text(0, LOD_TARGET_P - 0.02, "95%", fontsize=8)
    ax.text(ci_low_plot, 0.01, f"{ci_low:.2f}", fontsize=8, ha="left")
    ax.text(lod_log_plot, 0.01, f"{lod_log:.2f}", fontsize=8, ha="left")
    ax.text(ci_high_plot, 0.01, f"{ci_high:.2f}", fontsize=8, ha="left")

    band_x_low, band_y_low, band_x_high, band_y_high = [], [], [], []
    for prob in np.arange(0.001, 0.999, 0.005):
        xp, se = dose_p(prob)
        band_x_low.append(float(np.clip(xp - z975 * se, X_MIN, X_MAX)))
        band_y_low.append(prob)
        band_x_high.append(float(np.clip(xp + z975 * se, X_MIN, X_MAX)))
        band_y_high.append(prob)
    ax.plot(band_x_low, band_y_low, ".", color="red", markersize=1)
    ax.plot(band_x_high, band_y_high, ".", color="red", markersize=1)

    ax.set_xlim(X_MIN, X_MAX)
    ax.set_ylim(0, 1.02)
    ax.set_xlabel("log(concentration) (copies/test)")
    ax.set_ylabel("proportion")
    ax.set_title(f"IRON-qPCR_{pathogen}_LoD")

    fig.tight_layout()
    fig.savefig(out_png_path, dpi=120)
    plt.close(fig)


def embed_plot_image(ws, png_path: Path, anchor_cell: str):
    """이전에 이 스크립트가 넣은 이미지를 지우고 새 이미지를 anchor_cell 위치에 삽입한다."""
    ws._images = []
    img = XLImage(str(png_path))
    img.anchor = anchor_cell
    ws.add_image(img)


def write_probit_lod(wb: openpyxl.Workbook, pathogens: list[str],
                      conc_ct_lookup: dict[str, list[tuple[float, str]]],
                      last_data_row: int, logs: list[LogRow]) -> Path | None:
    """각 병원체 시트의 Ct 데이터로 probit LoD를 구해 AC22~24에 쓰고, 그래프를 삽입한다.
    conc_ct_lookup: 시트명 -> [(농도값, Ct열), ...] (generate/update는 호출부에서 만든 매핑을,
    recalc는 read_existing_conc_ct_pairs로 복원한 매핑을 넘긴다)"""
    if not PROBIT_LIBS_AVAILABLE:
        logs.append(LogRow("", "", status="오류",
                            message="probit 계산에 필요한 패키지(statsmodels/scipy/matplotlib)가 설치되지 않음 "
                                    "- pip install -r requirements.txt 로 설치 필요"))
        return None

    tmp_dir = Path(tempfile.mkdtemp(prefix="lod_probit_"))

    for pathogen in pathogens:
        sheet_name = f"{TARGET_SHEET_PREFIX}{pathogen}"
        if sheet_name not in wb.sheetnames:
            logs.append(LogRow("", "", target_sheet=sheet_name, status="오류",
                                message="LoD(Probit) 입력 대상 시트 없음"))
            continue
        ws = wb[sheet_name]
        conc_ct_pairs = conc_ct_lookup.get(sheet_name, [])
        if not conc_ct_pairs:
            logs.append(LogRow("", "", target_sheet=sheet_name, status="오류",
                                message="농도/Ct열 매핑을 찾을 수 없어 LoD 계산 건너뜀"))
            continue

        points = read_detection_counts(ws, conc_ct_pairs, last_data_row)
        try:
            fit = fit_probit_lod(points)
        except Exception as e:
            logs.append(LogRow("", "", target_sheet=sheet_name, status="오류",
                                message=f"probit 회귀 실패: {e}"))
            continue

        ws[LOD_LOGX_CELL] = round(fit["lod_log"], 4)
        ws[LOD_CI_LOW_CELL] = round(fit["ci_low_log"], 4)
        ws[LOD_CI_HIGH_CELL] = round(fit["ci_high_log"], 4)
        logs.append(LogRow(
            "", "", target_sheet=sheet_name, target_cell=LOD_LOGX_CELL,
            written_value=round(fit["lod_log"], 4), status="LoD계산",
            message=f"LogX(LoD)={fit['lod_log']:.4f}, CI=[{fit['ci_low_log']:.4f}, {fit['ci_high_log']:.4f}]",
        ))

        png_path = tmp_dir / f"{pathogen}_lod.png"
        try:
            generate_probit_plot(pathogen, fit, png_path)
            embed_plot_image(ws, png_path, LOD_PLOT_ANCHOR)
            logs.append(LogRow("", "", target_sheet=sheet_name, target_cell=LOD_PLOT_ANCHOR,
                                status="그래프삽입"))
        except Exception as e:
            logs.append(LogRow("", "", target_sheet=sheet_name, status="오류",
                                message=f"그래프 생성/삽입 실패: {e}"))

    return tmp_dir


# ======================================================================
# 하위 폴더(N_1) 탐색 - 기존과 동일
# ======================================================================

def find_final_result_files(sub_dir: Path) -> list[Path]:
    return sorted(
        p for p in sub_dir.iterdir()
        if p.is_file() and "finalresult" in p.name.lower()
    )


def find_subfolders(top_dir: Path, n: int) -> list[Path]:
    """폴더명이 정확히 'N_1'이거나 끝이 'N_1'로 끝나는 폴더를 찾는다.
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


def parse_folder_number(name: str) -> int | None:
    """폴더명 끝이 'N_1'이면 N을 반환. 아니면 None."""
    m = TRAILING_NUMBER_RE.search(name)
    return int(m.group(1)) if m else None


def check_all_rows_valid(src: SourceWorkbook,
                          check_rows: list[int]) -> tuple[bool, dict[int, str]]:
    """check_rows(Target_input.xlsx에서 유도된 검사행들)을 모두 확인한다. 전부 'Valid'여야
    이 폴더 전체를 사용할 수 있다. 하나라도 Invalid(또는 그 외 값)면 그 폴더는 모든 시트에서
    통째로 제외된다."""
    statuses: dict[int, str] = {}
    all_valid = True
    for row in check_rows:
        b_val = src.cell("B", row)
        status = str(b_val).strip() if b_val is not None else ""
        statuses[row] = status or "(빈 값)"
        if status.lower() != "valid":
            all_valid = False
    return all_valid, statuses


def apply_subfolder_count(n: int) -> int:
    """농도당 반복(하위 폴더) 개수로부터 데이터 입력 마지막 행을 계산한다."""
    return FIRST_DATA_ROW + n - 1


# ======================================================================
# generate
# ======================================================================

def process(root: Path, template_path: Path, out_path: Path, subfolder_count: int,
            target_input_path: Path) -> list[LogRow]:
    logs: list[LogRow] = []
    last_data_row = apply_subfolder_count(subfolder_count)

    pathogens, extraction_rules = load_target_config(target_input_path)
    logs.append(LogRow("", "", status="타겟매핑로드",
                        message=f"{target_input_path} 에서 병원체 {len(pathogens)}개, "
                                f"검사행 {len(extraction_rules)}개 로드"))

    conc_folders, skipped = discover_concentration_folders(root)
    for name in skipped:
        logs.append(LogRow(name, "", status="건너뜀(농도아님)",
                            message="폴더명에서 숫자를 추출할 수 없어 농도 폴더로 보지 않음"))
    if not conc_folders:
        raise ValueError(f"{root} 안에서 농도로 해석 가능한 폴더를 하나도 찾지 못했습니다.")
    conc_values_desc = [f.value for f in conc_folders]
    logs.append(LogRow("", "", status="농도탐지",
                        message="내림차순 농도: " + ", ".join(f"{f.name}({f.value})" for f in conc_folders)))

    wb = openpyxl.load_workbook(template_path)  # 서식/수식 보존을 위해 복사본에 직접 로드

    regression_x_range, regression_y_range = write_concentration_summary(
        wb, pathogens, conc_values_desc, last_data_row, logs)

    conc_ct_lookup: dict[str, list[tuple[float, str]]] = {}
    for pathogen in pathogens:
        sheet_name = f"{TARGET_SHEET_PREFIX}{pathogen}"
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        ct_col_map = build_ct_column_map(ws, conc_values_desc, sheet_name)
        conc_ct_lookup[sheet_name] = [(v, ct_col_map[v]) for v in conc_values_desc]

    check_rows = [row for row, _ in extraction_rules]

    for folder in conc_folders:
        ct_col_by_sheet = {
            sheet: dict(pairs) for sheet, pairs in conc_ct_lookup.items()
        }

        for n in range(1, subfolder_count + 1):
            sub_name = f"{n}_1"
            target_row = FIRST_DATA_ROW + (n - 1)

            sub_dirs = find_subfolders(folder.path, n)
            if len(sub_dirs) == 0:
                logs.append(LogRow(folder.name, sub_name, status="오류", message="하위 폴더를 찾을 수 없음"))
                continue
            if len(sub_dirs) > 1:
                names = ", ".join(d.name for d in sub_dirs)
                logs.append(LogRow(folder.name, sub_name, status="오류",
                                    message=f"'{sub_name}'로 끝나는 하위 폴더가 {len(sub_dirs)}개 발견됨: {names}"))
                continue
            sub_dir = sub_dirs[0]

            matches = find_final_result_files(sub_dir)
            if len(matches) == 0:
                logs.append(LogRow(folder.name, sub_name, status="오류", message="FinalResult 파일을 찾을 수 없음"))
                continue
            if len(matches) > 1:
                names = ", ".join(m.name for m in matches)
                logs.append(LogRow(folder.name, sub_name, status="오류",
                                    message=f"FinalResult 파일이 {len(matches)}개 발견됨: {names}"))
                continue

            src_path = matches[0]
            try:
                src = SourceWorkbook(src_path)
            except Exception as e:
                logs.append(LogRow(folder.name, sub_name, status="오류", message=f"파일 열기 실패: {e}"))
                continue

            all_valid, statuses = check_all_rows_valid(src, check_rows)

            if not all_valid:
                invalid_desc = ", ".join(f"B{r}={s}" for r, s in statuses.items() if s.lower() != "valid")
                logs.append(LogRow(
                    folder.name, sub_name, status="건너뜀(폴더 전체)",
                    message=f"{invalid_desc} - 하나라도 Invalid라 이 폴더는 모든 탭에서 {target_row}행 전체 제외",
                ))
                src.close()
                continue

            for row, extractions in extraction_rules:
                for col, pathogen in extractions:
                    raw_val = src.cell(col, row)
                    value = clean_value(raw_val)
                    target_sheet_name = f"{TARGET_SHEET_PREFIX}{pathogen}"

                    if target_sheet_name not in wb.sheetnames:
                        logs.append(LogRow(
                            folder.name, sub_name, check_row=f"B{row}", judgement="Valid",
                            source_cell=f"{col}{row}", raw_value=raw_val,
                            status="오류", message=f"타겟 시트 '{target_sheet_name}' 없음",
                        ))
                        continue

                    ct_col = ct_col_by_sheet.get(target_sheet_name, {}).get(folder.value)
                    if ct_col is None:
                        logs.append(LogRow(
                            folder.name, sub_name, check_row=f"B{row}", judgement="Valid",
                            source_cell=f"{col}{row}", raw_value=raw_val,
                            status="오류", message=f"'{target_sheet_name}'에서 농도 {folder.value}에 대응하는 Ct열을 찾지 못함",
                        ))
                        continue

                    ws = wb[target_sheet_name]
                    target_cell = f"{ct_col}{target_row}"
                    ws[target_cell] = value

                    logs.append(LogRow(
                        folder.name, sub_name, check_row=f"B{row}", judgement="Valid",
                        source_cell=f"{col}{row}", raw_value=raw_val, written_value=value,
                        target_sheet=target_sheet_name, target_cell=target_cell,
                        status="기록완료",
                    ))

            src.close()

    write_slope_intercept_formulas(wb, pathogens, regression_x_range, regression_y_range, logs)
    tmp_dir = write_probit_lod(wb, pathogens, conc_ct_lookup, last_data_row, logs)

    wb.save(out_path)
    if tmp_dir is not None:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    return logs


# ======================================================================
# update (재측정 반영)
# ======================================================================

def update_with_retest(existing_path: Path, retest_root: Path, out_path: Path,
                        subfolder_count: int, target_input_path: Path) -> list[LogRow]:
    """재측정 폴더(existing_path와 동일한 상위cp/N_1 구조, 재측정된 항목만 존재)의 값으로
    기존 결과 파일을 업데이트한다. 실제로 값이 바뀐 셀만 덮어쓰고 노란색으로 표시한다."""
    logs: list[LogRow] = []
    last_data_row = apply_subfolder_count(subfolder_count)

    pathogens, extraction_rules = load_target_config(target_input_path)
    check_rows = [row for row, _ in extraction_rules]

    wb = openpyxl.load_workbook(existing_path)

    conc_ct_lookup: dict[str, list[tuple[float, str]]] = {}
    for pathogen in pathogens:
        sheet_name = f"{TARGET_SHEET_PREFIX}{pathogen}"
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        conc_ct_lookup[sheet_name] = read_existing_conc_ct_pairs(ws)

    retest_folders, skipped = discover_concentration_folders(retest_root)
    for name in skipped:
        logs.append(LogRow(name, "", status="건너뜀(농도아님)",
                            message="폴더명에서 숫자를 추출할 수 없어 농도 폴더로 보지 않음"))

    for folder in retest_folders:
        ct_col_by_sheet: dict[str, str] = {}
        for sheet, pairs in conc_ct_lookup.items():
            for value, ct_col in pairs:
                if value == folder.value:
                    ct_col_by_sheet[sheet] = ct_col
                    break

        for sub_dir in sorted(p for p in folder.path.iterdir() if p.is_dir()):
            n = parse_folder_number(sub_dir.name)
            if n is None or not (1 <= n <= subfolder_count):
                logs.append(LogRow(folder.name, sub_dir.name, status="오류",
                                    message="폴더명에서 번호(N_1)를 인식할 수 없음"))
                continue

            sub_name = f"{n}_1"
            target_row = FIRST_DATA_ROW + (n - 1)

            matches = find_final_result_files(sub_dir)
            if len(matches) == 0:
                logs.append(LogRow(folder.name, sub_name, status="오류", message="FinalResult 파일을 찾을 수 없음"))
                continue
            if len(matches) > 1:
                names = ", ".join(m.name for m in matches)
                logs.append(LogRow(folder.name, sub_name, status="오류",
                                    message=f"FinalResult 파일이 {len(matches)}개 발견됨: {names}"))
                continue

            src_path = matches[0]
            try:
                src = SourceWorkbook(src_path)
            except Exception as e:
                logs.append(LogRow(folder.name, sub_name, status="오류", message=f"파일 열기 실패: {e}"))
                continue

            all_valid, statuses = check_all_rows_valid(src, check_rows)

            if not all_valid:
                invalid_desc = ", ".join(f"B{r}={s}" for r, s in statuses.items() if s.lower() != "valid")
                logs.append(LogRow(
                    folder.name, sub_name, status="건너뜀(폴더 전체, 기존값 유지)",
                    message=f"{invalid_desc} - 이 재측정 폴더는 모든 탭에서 사용하지 않고 기존값 유지",
                ))
                src.close()
                continue

            for row, extractions in extraction_rules:
                for col, pathogen in extractions:
                    raw_val = src.cell(col, row)
                    value = clean_value(raw_val)
                    target_sheet_name = f"{TARGET_SHEET_PREFIX}{pathogen}"

                    if target_sheet_name not in wb.sheetnames:
                        logs.append(LogRow(
                            folder.name, sub_name, check_row=f"B{row}", judgement="Valid",
                            source_cell=f"{col}{row}", raw_value=raw_val,
                            status="오류", message=f"타겟 시트 '{target_sheet_name}' 없음",
                        ))
                        continue

                    ct_col = ct_col_by_sheet.get(target_sheet_name)
                    if ct_col is None:
                        logs.append(LogRow(
                            folder.name, sub_name, check_row=f"B{row}", judgement="Valid",
                            source_cell=f"{col}{row}", raw_value=raw_val,
                            status="오류",
                            message=f"'{target_sheet_name}'의 기존 Conc.열에서 농도 {folder.value}를 찾지 못함",
                        ))
                        continue

                    ws = wb[target_sheet_name]
                    target_cell = f"{ct_col}{target_row}"
                    old_value = ws[target_cell].value

                    if old_value == value:
                        logs.append(LogRow(
                            folder.name, sub_name, check_row=f"B{row}", judgement="Valid",
                            source_cell=f"{col}{row}", raw_value=raw_val,
                            previous_value=old_value, written_value=value,
                            target_sheet=target_sheet_name, target_cell=target_cell,
                            status="동일값-변경없음",
                        ))
                    else:
                        ws[target_cell] = value
                        ws[target_cell].fill = RETEST_FILL
                        logs.append(LogRow(
                            folder.name, sub_name, check_row=f"B{row}", judgement="Valid",
                            source_cell=f"{col}{row}", raw_value=raw_val,
                            previous_value=old_value, written_value=value,
                            target_sheet=target_sheet_name, target_cell=target_cell,
                            status="변경됨(노란색 표시)",
                        ))

            src.close()

    regression_x_range = regression_y_range = None
    conc_all = {v for pairs in conc_ct_lookup.values() for v, _ in pairs}
    if conc_all:
        n = len(conc_all)
        last_summary_row = REGRESSION_FIRST_ROW + n - 1
        regression_x_range = f"{REGRESSION_LOG_COL}{REGRESSION_FIRST_ROW}:{REGRESSION_LOG_COL}{last_summary_row}"
        regression_y_range = f"{REGRESSION_CTAVG_COL}{REGRESSION_FIRST_ROW}:{REGRESSION_CTAVG_COL}{last_summary_row}"
        write_slope_intercept_formulas(wb, pathogens, regression_x_range, regression_y_range, logs)
    tmp_dir = write_probit_lod(wb, pathogens, conc_ct_lookup, last_data_row, logs)

    wb.save(out_path)
    if tmp_dir is not None:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    return logs


# ======================================================================
# recalc (기울기/상수/LoD만 재계산)
# ======================================================================

def recalc_slope_and_lod(input_path: Path, out_path: Path, subfolder_count: int) -> list[LogRow]:
    """Ct 데이터가 이미 채워진 결과 파일을 열어, 각 시트의 Conc. 열 값을 직접 읽어 농도/Ct열
    매핑을 복원하고, 기울기/상수와 LoD(그래프 포함)만 다시 계산한다. raw 폴더나
    Target_input.xlsx 없이, 파일 자체의 시트 목록을 병원체 목록으로 사용한다."""
    logs: list[LogRow] = []
    last_data_row = apply_subfolder_count(subfolder_count)

    wb = openpyxl.load_workbook(input_path)
    pathogens = list(wb.sheetnames)

    conc_ct_lookup: dict[str, list[tuple[float, str]]] = {}
    conc_all: set[float] = set()
    for sheet_name in pathogens:
        ws = wb[sheet_name]
        pairs = read_existing_conc_ct_pairs(ws)
        if not pairs:
            logs.append(LogRow("", "", target_sheet=sheet_name, status="오류",
                                message="Conc. 열에서 농도값을 읽지 못해 이 시트는 재계산에서 제외"))
            continue
        conc_ct_lookup[sheet_name] = pairs
        conc_all.update(v for v, _ in pairs)

    if conc_all:
        n = len(conc_all)
        last_summary_row = REGRESSION_FIRST_ROW + n - 1
        regression_x_range = f"{REGRESSION_LOG_COL}{REGRESSION_FIRST_ROW}:{REGRESSION_LOG_COL}{last_summary_row}"
        regression_y_range = f"{REGRESSION_CTAVG_COL}{REGRESSION_FIRST_ROW}:{REGRESSION_CTAVG_COL}{last_summary_row}"
        write_slope_intercept_formulas(wb, list(conc_ct_lookup.keys()), regression_x_range,
                                        regression_y_range, logs)
    tmp_dir = write_probit_lod(wb, list(conc_ct_lookup.keys()), conc_ct_lookup, last_data_row, logs)

    wb.save(out_path)
    if tmp_dir is not None:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    return logs


def write_log(logs: list[LogRow], log_path: Path):
    log_wb = openpyxl.Workbook()
    ws = log_wb.active
    ws.title = "처리내역"

    headers = ["상위폴더", "하위폴더", "검사행", "판정", "소스셀", "원본값",
               "기존기록값", "기록값", "타겟시트", "타겟셀", "상태", "메시지"]
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
    changed = sum(1 for l in logs if l.status == "변경됨(노란색 표시)")
    unchanged = sum(1 for l in logs if l.status == "동일값-변경없음")
    skipped = sum(1 for l in logs if l.status.startswith("건너뜀"))
    formulas = sum(1 for l in logs if l.status == "수식입력")
    errors = [l for l in logs if l.status == "오류"]
    warnings = [l for l in logs if l.status == "경고"]

    summary.append(["항목", "값"])
    summary.append(["총 로그 행 수", total])
    summary.append(["기록 완료(신규 생성)", written])
    summary.append(["변경됨(재측정, 노란색 표시)", changed])
    summary.append(["동일값-변경없음(재측정)", unchanged])
    summary.append(["Invalid/이상값/농도아님으로 건너뜀", skipped])
    summary.append(["기울기/상수 수식 입력", formulas])
    summary.append(["오류 건수", len(errors)])
    summary.append(["경고 건수", len(warnings)])
    summary.append([])
    summary.append(["오류 목록"])
    for l in errors:
        summary.append([f"{l.top_folder}/{l.sub_folder} {l.target_sheet}", l.message])
    summary.append([])
    summary.append(["경고 목록"])
    for l in warnings:
        summary.append([f"{l.top_folder}/{l.sub_folder} {l.check_row}", l.message])

    log_wb.save(log_path)


def resolve_out_and_log_paths(base_path: Path, out_arg: str | None, log_arg: str | None,
                               suffix_label: str) -> tuple[Path, Path]:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = Path(out_arg) if out_arg else base_path.with_name(
        f"{base_path.stem}_{suffix_label}_{timestamp}{base_path.suffix}")
    log_path = Path(log_arg) if log_arg else out_path.with_name(f"{out_path.stem}_로그.xlsx")
    return out_path, log_path


def print_summary(logs: list[LogRow], out_path: Path, log_path: Path):
    written = sum(1 for l in logs if l.status == "기록완료")
    changed = sum(1 for l in logs if l.status == "변경됨(노란색 표시)")
    unchanged = sum(1 for l in logs if l.status == "동일값-변경없음")
    skipped = sum(1 for l in logs if l.status.startswith("건너뜀"))
    errors = sum(1 for l in logs if l.status == "오류")
    warnings = sum(1 for l in logs if l.status == "경고")

    print(f"완료: 결과 파일 -> {out_path}")
    print(f"      로그 파일 -> {log_path}")
    print(f"      신규기록 {written}건 / 변경(재측정) {changed}건 / 동일값 {unchanged}건 / "
          f"건너뜀 {skipped}건 / 오류 {errors}건 / 경고 {warnings}건")
    if errors:
        print("오류가 발생했습니다. 로그 파일의 '요약' 시트를 확인하세요.")


def run_generate(args):
    root = Path(args.root)
    template_path = Path(args.template)
    target_input_path = Path(args.target_input) if args.target_input else root / TARGET_INPUT_FILENAME

    if not root.is_dir():
        sys.exit(f"오류: ROOT 폴더를 찾을 수 없습니다: {root}")
    if not template_path.is_file():
        sys.exit(f"오류: 템플릿 파일을 찾을 수 없습니다: {template_path}")

    out_path, log_path = resolve_out_and_log_paths(template_path, args.out, args.log, "결과")

    # 템플릿은 절대 직접 열어서 쓰지 않고, 항상 먼저 새 경로로 복사한 뒤 그 사본을 채운다.
    shutil.copy2(template_path, out_path)

    try:
        logs = process(root, out_path, out_path, args.subfolder_count, target_input_path)
    except (FileNotFoundError, ValueError) as e:
        sys.exit(f"오류: {e}")
    write_log(logs, log_path)
    print_summary(logs, out_path, log_path)


def run_update(args):
    existing_path = Path(args.existing)
    retest_root = Path(args.retest_root)
    target_input_path = (Path(args.target_input) if args.target_input
                          else retest_root / TARGET_INPUT_FILENAME)

    if not existing_path.is_file():
        sys.exit(f"오류: 기존 결과 파일을 찾을 수 없습니다: {existing_path}")
    if not retest_root.is_dir():
        sys.exit(f"오류: 재측정 폴더를 찾을 수 없습니다: {retest_root}")

    out_path, log_path = resolve_out_and_log_paths(existing_path, args.out, args.log, "재측정반영")

    # 기존 결과 파일도 직접 수정하지 않고, 새 경로로 복사한 뒤 그 사본에 반영한다.
    shutil.copy2(existing_path, out_path)

    try:
        logs = update_with_retest(out_path, retest_root, out_path, args.subfolder_count,
                                   target_input_path)
    except (FileNotFoundError, ValueError) as e:
        sys.exit(f"오류: {e}")
    write_log(logs, log_path)
    print_summary(logs, out_path, log_path)


def run_recalc(args):
    input_path = Path(args.input)

    if not input_path.is_file():
        sys.exit(f"오류: 입력 파일을 찾을 수 없습니다: {input_path}")

    out_path, log_path = resolve_out_and_log_paths(input_path, args.out, args.log, "재계산")

    # 입력 파일도 직접 수정하지 않고, 새 경로로 복사한 뒤 그 사본에 반영한다.
    shutil.copy2(input_path, out_path)

    logs = recalc_slope_and_lod(out_path, out_path, args.subfolder_count)
    write_log(logs, log_path)
    print_summary(logs, out_path, log_path)


def main():
    parser = argparse.ArgumentParser(description="LoD FinalResult 데이터를 타겟 엑셀로 자동 집계 (범용 버전)")
    subparsers = parser.add_subparsers(dest="command", required=True)

    gen = subparsers.add_parser("generate", help="ROOT 폴더 전체를 훑어 빈 템플릿에서 새 결과 파일 생성")
    gen.add_argument("--root", required=True,
                      help="농도 폴더들(예: 5000cp, 556cp...)과 Target_input.xlsx가 들어있는 경로")
    gen.add_argument("--template", required=True, help="타겟 LoD 템플릿 xlsx 경로 (읽기 전용, 수정되지 않음)")
    gen.add_argument("--target-input",
                      help=f"타겟 매핑 파일 경로 (기본: --root 안의 {TARGET_INPUT_FILENAME})")
    gen.add_argument("--out", help="결과를 저장할 새 파일 경로 (기본: 템플릿명_결과_타임스탬프.xlsx)")
    gen.add_argument("--log", help="처리 로그를 저장할 경로 (기본: 결과파일명_로그.xlsx)")
    gen.add_argument("--subfolder-count", type=int, default=16,
                      help="농도당 하위 폴더(반복) 개수 (기본 16)")
    gen.set_defaults(func=run_generate)

    upd = subparsers.add_parser("update", help="재측정 폴더의 값으로 기존 결과 파일을 업데이트 (바뀐 셀만 노란색 표시)")
    upd.add_argument("--existing", required=True, help="이전에 이 스크립트가 만든 결과 파일 경로 (읽기 전용, 수정되지 않음)")
    upd.add_argument("--retest-root", required=True,
                      help="재측정 raw 데이터 폴더 경로 (농도 폴더 구조, 재측정분만 있으면 됨)")
    upd.add_argument("--target-input",
                      help=f"타겟 매핑 파일 경로 (기본: --retest-root 안의 {TARGET_INPUT_FILENAME})")
    upd.add_argument("--out", help="결과를 저장할 새 파일 경로 (기본: 기존파일명_재측정반영_타임스탬프.xlsx)")
    upd.add_argument("--log", help="처리 로그를 저장할 경로 (기본: 결과파일명_로그.xlsx)")
    upd.add_argument("--subfolder-count", type=int, default=16,
                      help="농도당 하위 폴더(반복) 개수 (기본 16, 기존 결과 파일을 만들 때와 동일해야 함)")
    upd.set_defaults(func=run_update)

    rec = subparsers.add_parser("recalc", help="Ct 데이터가 이미 채워진 결과 파일에서 기울기/상수/LoD(그래프 포함)만 다시 계산")
    rec.add_argument("--input", required=True, help="Ct 데이터가 채워진 기존 결과 파일 경로 (읽기 전용, 수정되지 않음)")
    rec.add_argument("--out", help="결과를 저장할 새 파일 경로 (기본: 입력파일명_재계산_타임스탬프.xlsx)")
    rec.add_argument("--log", help="처리 로그를 저장할 경로 (기본: 결과파일명_로그.xlsx)")
    rec.add_argument("--subfolder-count", type=int, default=16,
                      help="입력 파일의 농도당 반복(데이터 행) 개수 (기본 16)")
    rec.set_defaults(func=run_recalc)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
