"""④ 산출자료 생성: 요약(다음 산출 때 전기 비교용), 기관별 엑셀, 검증 결과, 한글용 산출내역(docx)."""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .engine import FeeResult
from .validate import FAIL, OK, SKIP, WARN, Check
from .workbook import Aggregation

FONT = "맑은 고딕"
THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="DDE5F0")
TOTAL_FILL = PatternFill("solid", fgColor="F2F2F2")
RESULT_FILL = {OK: "E2EFDA", WARN: "FFF2CC", FAIL: "F8CBAD", SKIP: "EDEDED"}
NUM, RATIO, PCT = "#,##0", "0.00000000", "0.0%"


# ---------------------------------------------------------------------------
# 엑셀 공통
# ---------------------------------------------------------------------------
def _style_table(ws, top: int, left: int, n_rows: int, n_cols: int, formats: dict[int, str] | None = None,
                 total_row: bool = False) -> None:
    """top 행이 머리글인 표에 테두리·글꼴·숫자 서식을 적용한다."""
    formats = formats or {}
    for r in range(top, top + n_rows + 1):
        for c in range(left, left + n_cols):
            cell = ws.cell(r, c)
            cell.border = BORDER
            cell.font = Font(name=FONT, size=10, bold=(r == top or (total_row and r == top + n_rows)))
            if r == top:
                cell.fill = HEAD_FILL
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            else:
                if total_row and r == top + n_rows:
                    cell.fill = TOTAL_FILL
                if c - left in formats:
                    cell.number_format = formats[c - left]


def _write_df(ws, df: pd.DataFrame, top: int = 1, left: int = 1, formats: dict[str, str] | None = None,
              total_row: bool = False) -> None:
    formats = formats or {}
    for j, col in enumerate(df.columns):
        ws.cell(top, left + j, col)
    for i, row in enumerate(df.itertuples(index=False), start=1):
        for j, v in enumerate(row):
            ws.cell(top + i, left + j, None if pd.isna(v) else v)
    _style_table(ws, top, left, len(df), len(df.columns),
                 {j: formats[c] for j, c in enumerate(df.columns) if c in formats}, total_row)


def _autowidth(ws, min_w: int = 8, max_w: int = 40) -> None:
    widths: dict[int, int] = {}
    for row in ws.iter_rows():
        for c in row:
            if c.value is None:
                continue
            text = f"{c.value:,}" if isinstance(c.value, int) else str(c.value)
            w = sum(2 if ord(ch) > 0x2E80 else 1 for ch in text) + 2
            widths[c.column] = max(widths.get(c.column, 0), w)
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = max(min_w, min(max_w, w))


def _title(ws, text: str, sub: str | None = None) -> int:
    ws["A1"] = text
    ws["A1"].font = Font(name=FONT, size=14, bold=True)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = Font(name=FONT, size=9, color="666666")
    return 4


def _total_row(df: pd.DataFrame, label_col: str, sum_cols: list[str]) -> pd.DataFrame:
    total = {c: "" for c in df.columns}
    total[label_col] = "합계"
    for c in sum_cols:
        total[c] = df[c].sum()
    return pd.concat([df, pd.DataFrame([total])], ignore_index=True)


# ---------------------------------------------------------------------------
# 피벗 검토 파일
# ---------------------------------------------------------------------------
def write_pivot_review(agg: Aggregation, path: str) -> None:
    with pd.ExcelWriter(path) as xw:
        agg.grouped.to_excel(xw, sheet_name="항목별 합계")
        for sheet, pivot in agg.pivots:
            pivot.assign(총합계=pivot.sum(axis=1)).to_excel(xw, sheet_name=f"피벗_{sheet}"[:31])
        for name, item in agg.items.items():
            if item.get("compare_with"):
                other = item["compare_with"]
                cmp = agg.grouped[[name, other]].assign(차이=agg.grouped[name] - agg.grouped[other])
                cmp.to_excel(xw, sheet_name=f"비교_{name}"[:31])
        pd.DataFrame(
            [(n, i.get("column", "(입력 안 함)"), ", ".join(i["sources"]),
              "확정" if i.get("confirmed") else "임시(확인 필요)") for n, i in agg.items.items()],
            columns=["항목", "계산파일 열", "포함 로우데이터 열", "상태"],
        ).to_excel(xw, sheet_name="매핑 규칙", index=False)


# ---------------------------------------------------------------------------
# 산출결과 요약 (다음 산출 때 --prev 로 사용)
# ---------------------------------------------------------------------------
SUMMARY_FORMATS = {"지로": NUM, "CMS": NUM, "펌뱅킹": NUM, "계좌자동납부": NUM, "카드이동": NUM, "실적합계": NUM,
                   "실적비율": RATIO, "실적회비": NUM, "균등회비": NUM, "회비합계": NUM}


def write_summary(res: FeeResult, period: str, path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "기관별 산출결과"
    df = res.table.reset_index()
    _write_df(ws, df, formats={c: SUMMARY_FORMATS.get(c, NUM) for c in df.columns if c not in ("기관코드", "기관명")})
    ws.freeze_panes = "C2"
    _autowidth(ws)

    ws2 = wb.create_sheet("총액 내역")
    s = res.summary
    rows = [("산출 기간", period), ("생성 일시", datetime.now().strftime("%Y-%m-%d %H:%M")),
            ("회비 총액", s["회비총액"]), (f"자동송금분 ({s['자동송금비율']:.0%})", s["자동송금분"]),
            ("자동납부분", s["자동납부분"]), (f"균등회비 ({s['균등회비비율']:.0%})", s["균등회비총액"]),
            ("실적회비", s["실적회비총액"]), ("균등회비 나누는 기관 수", s["참가기관수"]),
            ("기관당 균등회비", s["기관당균등회비"]), ("실적 합계(건)", s["실적합계"])]
    ws2.append(["항목", "값"])
    for r in rows:
        ws2.append(list(r))
    _style_table(ws2, 1, 1, len(rows), 2, {1: NUM})
    _autowidth(ws2)
    wb.save(path)


# ---------------------------------------------------------------------------
# 검증 결과
# ---------------------------------------------------------------------------
def write_validation(checks: list[Check], prev_df: pd.DataFrame | None, cross_df: pd.DataFrame | None,
                     path: str, period: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "검증 요약"
    n = {k: sum(c.결과 == k for c in checks) for k in (OK, WARN, FAIL, SKIP)}
    top = _title(ws, f"{period} 회비 산출 검증 결과",
                 f"정상 {n[OK]} · 주의 {n[WARN]} · 오류 {n[FAIL]} · 생략 {n[SKIP]}  "
                 f"({datetime.now():%Y-%m-%d %H:%M})")
    df = pd.DataFrame([c.__dict__ for c in checks])
    _write_df(ws, df, top=top)
    for i, c in enumerate(checks, start=top + 1):
        ws.cell(i, 3).fill = PatternFill("solid", fgColor=RESULT_FILL[c.결과])
        ws.cell(i, 3).alignment = Alignment(horizontal="center")
        ws.cell(i, 4).alignment = Alignment(wrap_text=True, vertical="top")
    _autowidth(ws, max_w=60)
    ws.column_dimensions["D"].width = 90

    if prev_df is not None:
        ws2 = wb.create_sheet("전기 대비")
        fm = {c: (PCT if c.endswith("증감률") else NUM) for c in prev_df.columns if c not in ("기관코드", "기관명", "이상 여부")}
        _write_df(ws2, prev_df, formats=fm)
        flag_col = list(prev_df.columns).index("이상 여부") + 1
        for i, v in enumerate(prev_df["이상 여부"], start=2):
            if v:
                for c in range(1, len(prev_df.columns) + 1):
                    ws2.cell(i, c).fill = PatternFill("solid", fgColor=RESULT_FILL[WARN])
                ws2.cell(i, flag_col).font = Font(name=FONT, size=10, bold=True, color="C00000")
        ws2.freeze_panes = "C2"
        _autowidth(ws2)

    if cross_df is not None:
        ws3 = wb.create_sheet("교차검증")
        _write_df(ws3, cross_df, formats={"계산파일 회비합계": NUM, "엔진 회비합계": NUM})
        for i, v in enumerate(cross_df["결과"], start=2):
            ws3.cell(i, 5).fill = PatternFill("solid", fgColor=RESULT_FILL[OK if v == "일치" else FAIL])
        _autowidth(ws3)
    wb.save(path)


# ---------------------------------------------------------------------------
# 기관별 산출자료
# ---------------------------------------------------------------------------
def safe_filename(text: str) -> str:
    return re.sub(r'[\\/:*?"<>|\s]+', "_", text).strip("_")


def org_filename(code: str, name: str, period: str) -> str:
    return safe_filename(f"{code}_{name}_{period}_회비산출내역") + ".xlsx"


def write_org_files(cfg: dict, res: FeeResult, agg: Aggregation, period: str, out_dir: str) -> dict[str, Path]:
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    groups = cfg["summary"]["groups"]
    s = res.summary
    paths = {}
    for code, r in res.table.iterrows():
        wb = Workbook()
        ws = wb.active
        ws.title = "회비 산출내역"
        top = _title(ws, f"{period} 자동이체통합관리 회비 산출내역",
                     f"{code} {r['기관명']}  |  생성 {datetime.now():%Y-%m-%d}")

        ws.cell(top, 1, "1. 회비 내역").font = Font(name=FONT, size=11, bold=True)
        fee_df = pd.DataFrame([
            ("실적회비", int(r["실적회비"]), f"실적회비 총액 {s['실적회비총액']:,}원 × 실적 비율"),
            ("균등회비", int(r["균등회비"]), f"균등회비 총액 {s['균등회비총액']:,}원 ÷ {s['참가기관수']}개 기관"),
            ("회비 합계", int(r["회비합계"]), ""),
        ], columns=["구분", "금액(원)", "산출 근거"])
        _write_df(ws, fee_df, top=top + 1, formats={"금액(원)": NUM}, total_row=True)
        row = top + 1 + len(fee_df) + 2

        ws.cell(row, 1, "2. 실적 내역").font = Font(name=FONT, size=11, bold=True)
        detail = []
        g = agg.grouped.loc[code] if code in agg.grouped.index else None
        for grp, item_names in groups.items():
            for it in item_names:
                detail.append((grp, it, int(g[it]) if g is not None else 0))
        detail.append(("카드이동", "카드이동서비스", int(r["카드이동"])))
        det_df = pd.DataFrame(detail, columns=["서비스", "항목", "건수"])
        det_df = _total_row(det_df, "서비스", ["건수"])
        _write_df(ws, det_df, top=row + 1, formats={"건수": NUM}, total_row=True)
        row = row + 1 + len(det_df) + 2

        ws.cell(row, 1, "3. 실적 비율").font = Font(name=FONT, size=11, bold=True)
        ratio_df = pd.DataFrame([
            ("기관 실적 합계(건)", int(r["실적합계"])),
            ("전체 실적 합계(건)", s["실적합계"]),
            ("실적 비율", float(r["실적비율"])),
        ], columns=["구분", "값"])
        _write_df(ws, ratio_df, top=row + 1)
        for i in range(row + 2, row + 4):
            ws.cell(i, 2).number_format = NUM
        ws.cell(row + 4, 2).number_format = "0.000000%"
        row = row + 1 + len(ratio_df) + 2

        ws.cell(row, 1, "※ 실적 비율은 소수점 8자리 반올림 후 단수 조정(합계 1), "
                        "실적회비는 원 단위 반올림 후 단수 조정(합계 = 실적회비 총액)한 값입니다."
                ).font = Font(name=FONT, size=9, color="666666")
        _autowidth(ws, max_w=50)
        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 28
        ws.column_dimensions["C"].width = 42
        ws.page_setup.orientation = "portrait"
        ws.page_setup.fitToWidth = 1
        p = out / org_filename(code, r["기관명"], period)
        wb.save(p)
        paths[code] = p
    return paths


# ---------------------------------------------------------------------------
# 한글용 산출내역 (docx: 한글에서 바로 열거나 표를 복사해 붙여넣기)
# ---------------------------------------------------------------------------
def _set_font(run, size: float, bold: bool = False, color: RGBColor | None = None) -> None:
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.append(fonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia"):
        fonts.set(qn(attr), FONT)


def _shade(cell, hex_color: str) -> None:
    tcpr = cell._element.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hex_color)
    tcpr.append(shd)


def _doc_table(doc, header: list[str], rows: list[list], aligns: list[str], widths_cm: list[float],
               total_last: bool = False, size: float = 8.5) -> None:
    table = doc.add_table(rows=1 + len(rows), cols=len(header))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    al = {"l": WD_ALIGN_PARAGRAPH.LEFT, "c": WD_ALIGN_PARAGRAPH.CENTER, "r": WD_ALIGN_PARAGRAPH.RIGHT}
    for i, values in enumerate([header] + rows):
        for j, v in enumerate(values):
            cell = table.cell(i, j)
            cell.width = Cm(widths_cm[j])
            p = cell.paragraphs[0]
            p.alignment = al["c"] if i == 0 else al[aligns[j]]
            p.paragraph_format.space_after = Pt(0)
            head_or_total = i == 0 or (total_last and i == len(rows))
            _set_font(p.add_run(str(v)), size, bold=head_or_total)
            if i == 0:
                _shade(cell, "DDE5F0")
            elif total_last and i == len(rows):
                _shade(cell, "F2F2F2")


def _heading(doc, text: str, size: float = 12) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    _set_font(p.add_run(text), size, bold=True)


def write_hwp_doc(res: FeeResult, period: str, path: str) -> None:
    t, s = res.table, res.summary
    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
    for side in ("left_margin", "right_margin"):
        setattr(sec, side, Cm(1.8))
    sec.top_margin = sec.bottom_margin = Cm(1.5)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_font(p.add_run(f"{period} 자동이체통합관리 회비 산출내역"), 16, bold=True)

    _heading(doc, "□ 회비 총괄")
    _doc_table(doc, ["구분", "금액(원)", "비고"], [
        ["회비 총액", f"{s['회비총액']:,}", ""],
        ["자동송금분", f"{s['자동송금분']:,}", f"총액의 {s['자동송금비율']:.0%}"],
        ["자동납부분", f"{s['자동납부분']:,}", f"총액의 {1 - s['자동송금비율']:.0%}"],
        ["  균등회비", f"{s['균등회비총액']:,}", f"자동납부분의 {s['균등회비비율']:.0%}, "
                                                f"{s['참가기관수']}개 기관 × {s['기관당균등회비']:,}원"],
        ["  실적회비", f"{s['실적회비총액']:,}", "자동납부분 - 균등회비, 기관별 실적 비율로 배분"],
    ], ["l", "r", "l"], [4, 4.5, 12])

    _heading(doc, "□ 자동납부분 실적")
    groups = s["실적묶음"]
    num_cols = [*groups, "카드이동", "실적합계"]
    rows = [[c, r["기관명"], *(f"{r[k]:,}" for k in num_cols), f"{r['실적비율'] * 100:.6f}%"]
            for c, r in t.iterrows()]
    rows.append(["합계", "", *(f"{t[k].sum():,}" for k in num_cols), f"{t['실적비율'].sum() * 100:.6f}%"])
    w = 18.4 / (len(num_cols) + 1)
    _doc_table(doc, ["기관코드", "기관명", *groups, "카드이동", "총합계", "실적 비율"], rows,
               ["c", "l"] + ["r"] * (len(num_cols) + 1), [1.6, 3.6] + [w] * (len(num_cols) + 1), True)

    _heading(doc, "□ 자동납부분 기관별 회비")
    rows = [[c, r["기관명"], f"{r['실적비율'] * 100:.6f}%", f"{r['실적회비']:,}", f"{r['균등회비']:,}",
             f"{r['회비합계']:,}"] for c, r in t.iterrows()]
    rows.append(["합계", "", f"{t['실적비율'].sum() * 100:.6f}%", f"{t['실적회비'].sum():,}",
                 f"{t['균등회비'].sum():,}", f"{t['회비합계'].sum():,}"])
    _doc_table(doc, ["기관코드", "기관명", "실적 비율", "실적회비(원)", "균등회비(원)", "회비 합계(원)"],
               rows, ["c", "l", "r", "r", "r", "r"], [1.8, 4.2, 3.2, 4.0, 3.6, 4.0], True)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    _set_font(p.add_run("※ 실적 비율은 소수점 8자리 반올림 후 분담률이 큰 기관부터 0.000001%씩 단수 조정(합계 100%), "
                        "실적회비는 원 단위 반올림 후 실적이 큰 기관부터 1원씩 단수 조정한 값입니다."),
              8.5, color=RGBColor(0x66, 0x66, 0x66))
    doc.save(path)
