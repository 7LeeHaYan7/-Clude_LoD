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
import tempfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.drawing.image import Image as XLImage

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
PATHOGENS = ["CPA", "CPE", "Giardia", "CPV2", "Campylobacter", "Salmonella", "CECoV"]

# 기울기/상수: 각 병원체 시트의 표준곡선(Log(농도) vs 평균 Ct) 선형회귀
SLOPE_CELL = "AB18"
INTERCEPT_CELL = "AB19"
REGRESSION_Y_RANGE = "R6:R11"  # 평균 Ct
REGRESSION_X_RANGE = "Q6:Q11"  # Log(농도)

# LoD(Probit) 결과 입력 셀: AA22=LoD LogX, AA23=95% CI 하한 LogX, AA24=95% CI 상한 LogX
LOD_LOGX_CELL = "AA22"
LOD_CI_LOW_CELL = "AA23"
LOD_CI_HIGH_CELL = "AA24"
LOD_TARGET_P = 0.95
LOD_PLOT_ANCHOR = "N40"
LAST_DATA_ROW = FIRST_DATA_ROW + SUBFOLDER_COUNT - 1  # 21

RETEST_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

SUBFOLDER_RE = re.compile(r"^(\d+)_1$")
TRAILING_NUMBER_RE = re.compile(r"(\d+)_1$")


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


def parse_folder_number(name: str) -> int | None:
    """폴더명 끝이 'N_1'이면 N을 반환. 아니면 None."""
    m = TRAILING_NUMBER_RE.search(name)
    return int(m.group(1)) if m else None


def write_slope_intercept_formulas(wb: openpyxl.Workbook, logs: list[LogRow]):
    """각 병원체 시트에 표준곡선 기울기(AB18)/상수(AB19) 수식을 입력한다."""
    for pathogen in PATHOGENS:
        sheet_name = f"{TARGET_SHEET_PREFIX}{pathogen}"
        if sheet_name not in wb.sheetnames:
            logs.append(LogRow("", "", target_sheet=sheet_name, status="오류",
                                message="기울기/상수 입력 대상 시트 없음"))
            continue
        ws = wb[sheet_name]
        ws[SLOPE_CELL] = f"=SLOPE({REGRESSION_Y_RANGE},{REGRESSION_X_RANGE})"
        ws[INTERCEPT_CELL] = f"=INTERCEPT({REGRESSION_Y_RANGE},{REGRESSION_X_RANGE})"
        logs.append(LogRow("", "", target_sheet=sheet_name, target_cell=SLOPE_CELL,
                            status="수식입력", message=f"기울기=SLOPE({REGRESSION_Y_RANGE},{REGRESSION_X_RANGE})"))
        logs.append(LogRow("", "", target_sheet=sheet_name, target_cell=INTERCEPT_CELL,
                            status="수식입력", message=f"상수=INTERCEPT({REGRESSION_Y_RANGE},{REGRESSION_X_RANGE})"))


def parse_conc_value(top_folder_name: str) -> float:
    """'556cp' -> 556.0 처럼 상위 폴더명에서 숫자 농도값을 뽑는다."""
    return float(top_folder_name.replace("cp", ""))


def read_detection_counts(ws) -> list[tuple[float, int, int]]:
    """이미 채워진 병원체 시트의 Ct 열(D6:D21 등)에서 농도별 (농도, 양성수, 시도수)를 구한다.
    숫자 값이 있으면 양성(검출), 'UD' 텍스트면 음성, 빈칸이면 그 폴더는 아예 제외(시도수에서도 제외) -
    타겟 시트의 Detect/Detect(%) 행(37~38)과 동일한 로직이다."""
    result = []
    for top in TOP_FOLDERS_ORDER:
        conc = parse_conc_value(top)
        col = CT_COLUMN[top]
        positive = 0
        total = 0
        for row in range(FIRST_DATA_ROW, LAST_DATA_ROW + 1):
            val = ws[f"{col}{row}"].value
            if val is None or (isinstance(val, str) and val.strip() == ""):
                continue
            total += 1
            if isinstance(val, (int, float)):
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
    """R 스크립트와 동일한 스타일(빨간 S커브, 95%선, LoD/CI 수직선, 신뢰구간 밴드)로 플롯을 그린다."""
    b0, b1 = fit["b0"], fit["b1"]
    dose_p = fit["dose_p"]
    lod_log, ci_low, ci_high = fit["lod_log"], fit["ci_low_log"], fit["ci_high_log"]
    z975 = norm.ppf(0.975)

    fig, ax = plt.subplots(figsize=(7, 5))

    x = np.linspace(-1, 4, 400)
    ax.plot(x, norm.cdf(b0 + b1 * x), color="red", linewidth=1.2)

    for log_conc, prop in fit["points_log"]:
        ax.plot(log_conc, prop, "o", color="red", markersize=5)

    ax.axhline(LOD_TARGET_P, linestyle="--", color="black", linewidth=0.8)
    ax.axvline(lod_log, linestyle="--", color="black", linewidth=0.8)
    ax.axvline(ci_low, linestyle=":", color="gray", linewidth=0.8)
    ax.axvline(ci_high, linestyle=":", color="gray", linewidth=0.8)
    ax.plot(lod_log, LOD_TARGET_P, "s", color="blue", markersize=6)

    ax.text(0, LOD_TARGET_P - 0.02, "95%", fontsize=8)
    ax.text(ci_low, 0.01, f"{ci_low:.2f}", fontsize=8, ha="left")
    ax.text(lod_log, 0.01, f"{lod_log:.2f}", fontsize=8, ha="left")
    ax.text(ci_high, 0.01, f"{ci_high:.2f}", fontsize=8, ha="left")

    band_x_low, band_y_low, band_x_high, band_y_high = [], [], [], []
    for prob in np.arange(0.001, 0.999, 0.005):
        xp, se = dose_p(prob)
        band_x_low.append(xp - z975 * se)
        band_y_low.append(prob)
        band_x_high.append(xp + z975 * se)
        band_y_high.append(prob)
    ax.plot(band_x_low, band_y_low, ".", color="red", markersize=1)
    ax.plot(band_x_high, band_y_high, ".", color="red", markersize=1)

    ax.set_xlim(-1, 4)
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


def write_probit_lod(wb: openpyxl.Workbook, logs: list[LogRow]) -> Path | None:
    """각 병원체 시트의 Ct 데이터로 probit LoD를 구해 AA22~24에 쓰고, 그래프를 삽입한다."""
    if not PROBIT_LIBS_AVAILABLE:
        logs.append(LogRow("", "", status="오류",
                            message="probit 계산에 필요한 패키지(statsmodels/scipy/matplotlib)가 설치되지 않음 "
                                    "- pip install -r requirements.txt 로 설치 필요"))
        return None

    tmp_dir = Path(tempfile.mkdtemp(prefix="lod_probit_"))

    for pathogen in PATHOGENS:
        sheet_name = f"{TARGET_SHEET_PREFIX}{pathogen}"
        if sheet_name not in wb.sheetnames:
            logs.append(LogRow("", "", target_sheet=sheet_name, status="오류",
                                message="LoD(Probit) 입력 대상 시트 없음"))
            continue
        ws = wb[sheet_name]

        points = read_detection_counts(ws)
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


def check_all_rows_valid(src: SourceWorkbook) -> tuple[bool, dict[int, str]]:
    """B18~B21을 모두 확인한다. 넷 다 'Valid'여야 이 폴더 전체를 사용할 수 있다.
    하나라도 Invalid(또는 그 외 값)면 그 폴더는 모든 시트에서 통째로 제외된다."""
    statuses: dict[int, str] = {}
    all_valid = True
    for row, _ in EXTRACTION_RULES:
        b_val = src.cell("B", row)
        status = str(b_val).strip() if b_val is not None else ""
        statuses[row] = status or "(빈 값)"
        if status.lower() != "valid":
            all_valid = False
    return all_valid, statuses


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

            all_valid, statuses = check_all_rows_valid(src)

            if not all_valid:
                invalid_desc = ", ".join(f"B{r}={s}" for r, s in statuses.items() if s.lower() != "valid")
                logs.append(LogRow(
                    top, sub_name, status="건너뜀(폴더 전체)",
                    message=f"{invalid_desc} - 하나라도 Invalid라 이 폴더는 모든 탭에서 {target_row}행 전체 제외",
                ))
                src.close()
                continue

            for row, extractions in EXTRACTION_RULES:
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

            src.close()

    write_slope_intercept_formulas(wb, logs)
    tmp_dir = write_probit_lod(wb, logs)

    wb.save(out_path)
    if tmp_dir is not None:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    return logs


def update_with_retest(existing_path: Path, retest_root: Path, out_path: Path) -> list[LogRow]:
    """재측정 폴더(existing_path와 동일한 상위cp/N_1 구조, 재측정된 항목만 존재)의 값으로
    기존 결과 파일을 업데이트한다. 실제로 값이 바뀐 셀만 덮어쓰고 노란색으로 표시한다."""
    logs: list[LogRow] = []

    wb = openpyxl.load_workbook(existing_path)

    for top_dir in sorted(p for p in retest_root.iterdir() if p.is_dir()):
        top = top_dir.name
        if top not in CT_COLUMN:
            logs.append(LogRow(top, "", status="오류",
                                message="알 수 없는 상위 폴더명 (5000cp/1670cp/556cp/185cp/62cp/21cp 중 하나여야 함)"))
            continue
        ct_col = CT_COLUMN[top]

        for sub_dir in sorted(p for p in top_dir.iterdir() if p.is_dir()):
            n = parse_folder_number(sub_dir.name)
            if n is None or not (1 <= n <= SUBFOLDER_COUNT):
                logs.append(LogRow(top, sub_dir.name, status="오류",
                                    message="폴더명에서 번호(N_1)를 인식할 수 없음"))
                continue

            sub_name = f"{n}_1"
            target_row = FIRST_DATA_ROW + (n - 1)

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

            all_valid, statuses = check_all_rows_valid(src)

            if not all_valid:
                invalid_desc = ", ".join(f"B{r}={s}" for r, s in statuses.items() if s.lower() != "valid")
                logs.append(LogRow(
                    top, sub_name, status="건너뜀(폴더 전체, 기존값 유지)",
                    message=f"{invalid_desc} - 이 재측정 폴더는 모든 탭에서 사용하지 않고 기존값 유지",
                ))
                src.close()
                continue

            for row, extractions in EXTRACTION_RULES:
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
                    old_value = ws[target_cell].value

                    if old_value == value:
                        logs.append(LogRow(
                            top, sub_name, check_row=f"B{row}", judgement="Valid",
                            source_cell=f"{col}{row}", raw_value=raw_val,
                            previous_value=old_value, written_value=value,
                            target_sheet=target_sheet_name, target_cell=target_cell,
                            status="동일값-변경없음",
                        ))
                    else:
                        ws[target_cell] = value
                        ws[target_cell].fill = RETEST_FILL
                        logs.append(LogRow(
                            top, sub_name, check_row=f"B{row}", judgement="Valid",
                            source_cell=f"{col}{row}", raw_value=raw_val,
                            previous_value=old_value, written_value=value,
                            target_sheet=target_sheet_name, target_cell=target_cell,
                            status="변경됨(노란색 표시)",
                        ))

            src.close()

    write_slope_intercept_formulas(wb, logs)
    tmp_dir = write_probit_lod(wb, logs)

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
    summary.append(["Invalid/이상값으로 건너뜀", skipped])
    summary.append(["기울기/상수 수식 입력", formulas])
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

    if not root.is_dir():
        sys.exit(f"오류: ROOT 폴더를 찾을 수 없습니다: {root}")
    if not template_path.is_file():
        sys.exit(f"오류: 템플릿 파일을 찾을 수 없습니다: {template_path}")

    out_path, log_path = resolve_out_and_log_paths(template_path, args.out, args.log, "결과")

    # 템플릿은 절대 직접 열어서 쓰지 않고, 항상 먼저 새 경로로 복사한 뒤 그 사본을 채운다.
    shutil.copy2(template_path, out_path)

    logs = process(root, out_path, out_path)
    write_log(logs, log_path)
    print_summary(logs, out_path, log_path)


def run_update(args):
    existing_path = Path(args.existing)
    retest_root = Path(args.retest_root)

    if not existing_path.is_file():
        sys.exit(f"오류: 기존 결과 파일을 찾을 수 없습니다: {existing_path}")
    if not retest_root.is_dir():
        sys.exit(f"오류: 재측정 폴더를 찾을 수 없습니다: {retest_root}")

    out_path, log_path = resolve_out_and_log_paths(existing_path, args.out, args.log, "재측정반영")

    # 기존 결과 파일도 직접 수정하지 않고, 새 경로로 복사한 뒤 그 사본에 반영한다.
    shutil.copy2(existing_path, out_path)

    logs = update_with_retest(out_path, retest_root, out_path)
    write_log(logs, log_path)
    print_summary(logs, out_path, log_path)


def main():
    parser = argparse.ArgumentParser(description="LoD FinalResult 데이터를 타겟 엑셀로 자동 집계")
    subparsers = parser.add_subparsers(dest="command", required=True)

    gen = subparsers.add_parser("generate", help="1st 폴더 전체를 훑어 빈 템플릿에서 새 결과 파일 생성")
    gen.add_argument("--root", required=True, help="1st 폴더 경로 (상위 폴더 5000cp 등이 들어있는 경로)")
    gen.add_argument("--template", required=True, help="타겟 LoD 템플릿 xlsx 경로 (읽기 전용, 수정되지 않음)")
    gen.add_argument("--out", help="결과를 저장할 새 파일 경로 (기본: 템플릿명_결과_타임스탬프.xlsx)")
    gen.add_argument("--log", help="처리 로그를 저장할 경로 (기본: 결과파일명_로그.xlsx)")
    gen.set_defaults(func=run_generate)

    upd = subparsers.add_parser("update", help="재측정 폴더의 값으로 기존 결과 파일을 업데이트 (바뀐 셀만 노란색 표시)")
    upd.add_argument("--existing", required=True, help="이전에 이 스크립트가 만든 결과 파일 경로 (읽기 전용, 수정되지 않음)")
    upd.add_argument("--retest-root", required=True,
                      help="재측정 raw 데이터 폴더 경로 (5000cp/N_1/FinalResult 같은 동일 구조, 재측정분만 있으면 됨)")
    upd.add_argument("--out", help="결과를 저장할 새 파일 경로 (기본: 기존파일명_재측정반영_타임스탬프.xlsx)")
    upd.add_argument("--log", help="처리 로그를 저장할 경로 (기본: 결과파일명_로그.xlsx)")
    upd.set_defaults(func=run_update)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
