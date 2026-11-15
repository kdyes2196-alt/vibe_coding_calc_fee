"""① 실적 집계: 로우데이터 피벗 → 계산파일 양식에 값·수식 입력."""

from __future__ import annotations

import random
from dataclasses import dataclass

import pandas as pd

from .pivot import check_mapping, group_items, load_raw_sheet, pivot_by_org
from .xlsx_writer import Formula, XlsxPatcher, split_ref


class CalcFeeError(Exception):
    """사용자가 설정이나 입력 파일을 고쳐야 하는 오류."""


@dataclass
class Aggregation:
    pivots: list[tuple[str, pd.DataFrame]]   # (시트 이름, 피벗 표)
    grouped: pd.DataFrame                    # 기관코드 × 항목 합계
    items: dict                              # 항목 이름 → 설정
    raw_totals: dict[str, int]               # 시트별 로우데이터 값 총합 (검증용)


@dataclass
class TemplateInfo:
    orgs: pd.DataFrame          # index=기관코드, columns=[기관명, 실적행, 회비행]
    card: pd.Series             # 기관코드 → 카드이동서비스 실적
    template_remit_rate: float | None


def build_pivots(raw_path: str, cfg: dict, log=print) -> Aggregation:
    """설정의 모든 로우데이터 시트를 피벗하고, 항목별 합계를 하나의 표로 합친다."""
    pivots, groups, items_all, raw_totals = [], [], {}, {}
    for src in cfg["sources"]:
        try:
            raw = load_raw_sheet(raw_path, src["sheet"], src["key_column"])
        except ValueError as e:
            raise CalcFeeError(f"로우데이터에서 '{src['sheet']}' 시트를 읽지 못했습니다: {e}")
        if src["key_column"] not in raw.columns:
            raise CalcFeeError(f"'{src['sheet']}' 시트에 기관코드 열 '{src['key_column']}'이 없습니다.")
        pivot = pivot_by_org(raw, src["key_column"], src.get("ignore_columns", []), cfg.get("code_merge", {}))
        try:
            check_mapping(list(pivot.columns), src["items"])
        except ValueError as e:
            raise CalcFeeError(f"[{src['sheet']}] {e}")
        for name in src["items"]:
            if name in items_all:
                raise CalcFeeError(f"항목 이름 '{name}'이 여러 시트에 중복되어 있습니다.")
        items_all.update(src["items"])
        pivots.append((src["sheet"], pivot))
        groups.append(group_items(pivot, src["items"]))
        value_cols = [c for c in raw.columns if c != src["key_column"] and c not in src.get("ignore_columns", [])]
        raw_totals[src["sheet"]] = int(raw[value_cols].apply(pd.to_numeric, errors="coerce").fillna(0).to_numpy().sum())
        log(f"[피벗] '{src['sheet']}' {len(raw):,}행 → 기관 {len(pivot)}곳 × 값 열 {len(pivot.columns)}개")

    columns = [i["column"] for i in items_all.values() if "column" in i]
    dup = sorted({c for c in columns if columns.count(c) > 1})
    if dup:
        raise CalcFeeError(f"계산파일 열 {dup}에 여러 항목이 지정되어 있습니다.")

    grouped = pd.concat(groups, axis=1).fillna(0).astype("int64").sort_index()
    return Aggregation(pivots, grouped, items_all, raw_totals)


def dummy_values(fill: dict) -> dict[int, int]:
    """[[dummy_fill]] 설정의 행 번호 → 난수. 같은 seed 면 항상 같은 값."""
    rng = random.Random(fill.get("seed"))
    return {r: rng.randint(fill["min"], fill["max"]) for r in range(fill["first_row"], fill["last_row"] + 1)}


def round_adjust_formulas(ra: dict) -> dict[str, Formula]:
    """각 행 = base + 단수 조정분.

    N = (target - SUMPRODUCT(base_range)) / unit  (조정해야 할 단위 개수)
    N > 0 이면 rank_column 상위 N개 행에 +unit, N < 0 이면 상위 |N|개 행에서 -unit.
    순위가 같은 행은 위쪽 행이 먼저 (COUNTIF 로 동순위 구분).
    """
    col, kcol, unit = ra["column"], ra["rank_column"], ra["unit"]
    first, last = ra["first_row"], ra["last_row"]
    rng = f"${kcol}${first}:${kcol}${last}"
    n = f"ROUND(({ra['target']}-SUMPRODUCT({ra['base_range']}))/({unit}),0)"
    out = {}
    for r in range(first, last + 1):
        rank = f"(RANK({kcol}{r},{rng})+COUNTIF(${kcol}${first - 1}:{kcol}{r - 1},{kcol}{r}))"
        base = ra["base"].format(r=r)
        out[f"{col}{r}"] = Formula(f"{base}+SIGN({n})*({rank}<=ABS({n}))*({unit})")
    return out


def _sum_formula(col: str, first: int, last: int) -> Formula:
    return Formula(f"SUM({col}{first}:{col}{last})")


def fill_template(cfg: dict, agg: Aggregation, template: str, out_path: str, log=print) -> TemplateInfo:
    """계산파일 양식에 집계값·수식을 넣어 out_path 로 저장하고, 기관 목록 등 정보를 돌려준다."""
    tgt, fee = cfg["target"], cfg["fee"]
    placed = {n: i for n, i in agg.items.items() if "column" in i}
    wb = XlsxPatcher(template)
    try:
        # 기관 목록: 실적 시트와 회비 시트 양쪽에서 읽어 대조
        t_rows = range(tgt["first_row"], tgt["last_row"] + 1)
        f_rows = range(fee["first_row"], fee["last_row"] + 1)
        t_codes = wb.read_column(tgt["sheet"], tgt["code_column"], t_rows)
        f_codes = wb.read_column(fee["sheet"], fee["code_column"], f_rows)
        f_names = wb.read_column(fee["sheet"], fee["name_column"], f_rows)
        row_of = {code: r for r, code in t_codes.items()}
        fee_row_of = {code: r for r, code in f_codes.items()}
        if len(row_of) != len(t_codes) or len(fee_row_of) != len(f_codes):
            raise CalcFeeError("계산파일에 같은 기관코드가 두 번 이상 있습니다.")
        if set(row_of) != set(fee_row_of):
            raise CalcFeeError(
                f"'{tgt['sheet']}'와 '{fee['sheet']}'의 기관 목록이 다릅니다: "
                f"{sorted(set(row_of) ^ set(fee_row_of))}"
            )
        missing = sorted(set(agg.grouped.index) - set(row_of))
        if missing:
            raise CalcFeeError(
                f"계산파일 '{tgt['sheet']}'에 없는 기관코드가 로우데이터에 있습니다: {missing}\n"
                "계산파일에 행을 추가하거나 config 의 [code_merge] 에 합산 규칙을 넣어 주세요."
            )

        # 1. 실적: 항목 값 + 행 합계 수식 + 합계 행
        values: dict[str, int | Formula] = {}
        for code, r in row_of.items():
            for name, item in placed.items():
                values[f"{item['column']}{r}"] = int(agg.grouped.at[code, name]) if code in agg.grouped.index else 0
            for col, formula in cfg.get("row_formulas", {}).items():
                values[f"{col}{r}"] = Formula(formula.format(r=r))
        for col in cfg.get("totals", {}).get("columns", []):
            values[f"{col}{tgt['total_row']}"] = _sum_formula(col, tgt["first_row"], tgt["last_row"])
        wb.set_cells(tgt["sheet"], values)
        log(f"[입력] '{tgt['sheet']}' 기관 {len(row_of)}곳 × 항목 {len(placed)}개")

        # 임시 난수 (카드이동서비스 등)
        card_rows: dict[int, int] = {}
        for fill in cfg.get("dummy_fill", []):
            dv = dummy_values(fill)
            cells: dict[str, int | Formula] = {f"{fill['column']}{r}": v for r, v in dv.items()}
            if fill.get("total_cell"):
                cells[fill["total_cell"]] = _sum_formula(fill["column"], fill["first_row"], fill["last_row"])
            wb.set_cells(fill["sheet"], cells)
            if fill["sheet"] == fee["sheet"] and fill["column"] == fee["card_column"]:
                card_rows = dv
            log(f"[임시] '{fill['sheet']}' {fill['column']}{fill['first_row']}:{fill['column']}{fill['last_row']} "
                f"난수 {len(dv)}개 입력 (테스트용)")

        # 회비 총액 + 기타 고정값
        wb.set_cells(fee["sheet"], {fee["total_fee_cell"]: fee["total_fee"]})
        log(f"[고정값] '{fee['sheet']}' {fee['total_fee_cell']} 회비 총액 = {fee['total_fee']:,}원")
        for fv in cfg.get("fixed_values", []):
            wb.set_cells(fv["sheet"], {fv["cell"]: fv["value"]})
            log(f"[고정값] '{fv['sheet']}' {fv['cell']} = {fv['value']:,}")

        for ff in cfg.get("formula_fill", []):
            col, first, last = ff["column"], ff["first_row"], ff["last_row"]
            cells = {f"{col}{r}": Formula(ff["formula"].format(r=r)) for r in range(first, last + 1)}
            if ff.get("total_cell"):
                cells[ff["total_cell"]] = _sum_formula(col, first, last)
            wb.set_cells(ff["sheet"], cells)
            log(f"[수식] '{ff['sheet']}' {col}{first}:{col}{last} = {ff['formula']}")

        for ra in cfg.get("round_adjust", []):
            wb.set_cells(ra["sheet"], round_adjust_formulas(ra))
            log(f"[단수조정] '{ra['sheet']}' {ra['column']}{ra['first_row']}:{ra['column']}{ra['last_row']} "
                f"합계 = {ra['target']} (큰 기관부터 {ra['unit']}씩)")

        rc, rr = split_ref(fee["remit_rate_cell"])
        rate_text = wb.read_column(fee["sheet"], rc, range(rr, rr + 1))
        try:
            template_rate = float(rate_text[rr]) if rr in rate_text else None
        except ValueError:
            template_rate = None

        wb.request_full_recalc()
        wb.save(out_path)
    finally:
        wb.close()

    orgs = pd.DataFrame(
        {"기관명": [f_names.get(fee_row_of[c], "") for c in row_of],
         "실적행": list(row_of.values()),
         "회비행": [fee_row_of[c] for c in row_of]},
        index=pd.Index(list(row_of), name="기관코드"),
    ).sort_values("회비행")
    card = pd.Series({c: card_rows.get(r, 0) for c, r in fee_row_of.items()}, name="카드이동").reindex(orgs.index)
    return TemplateInfo(orgs=orgs, card=card, template_remit_rate=template_rate)
