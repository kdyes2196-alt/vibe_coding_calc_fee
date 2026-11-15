"""실행: python -m calc_fee --raw 로우데이터.xlsx --template 계산파일.xlsx --out 결과.xlsx"""

from __future__ import annotations

import argparse
import sys
import tomllib
from pathlib import Path

import pandas as pd

from .pivot import check_mapping, group_items, load_raw_sheet, pivot_by_org
from .xlsx_writer import Formula, XlsxPatcher

DEFAULT_CONFIG = Path(__file__).resolve().parent.parent / "config" / "mapping.toml"


def build_pivots(raw_path: str, cfg: dict) -> tuple[list[tuple[str, pd.DataFrame]], pd.DataFrame, dict]:
    """설정의 모든 로우데이터 시트를 피벗하고, 항목별 합계를 하나의 표로 합친다."""
    pivots, groups, items_all = [], [], {}
    for src in cfg["sources"]:
        raw = load_raw_sheet(raw_path, src["sheet"], src["key_column"])
        pivot = pivot_by_org(raw, src["key_column"], src.get("ignore_columns", []), cfg.get("code_merge", {}))
        try:
            check_mapping(list(pivot.columns), src["items"])
        except ValueError as e:
            raise SystemExit(f"[{src['sheet']}] {e}")
        for name in src["items"]:
            if name in items_all:
                raise SystemExit(f"항목 이름 '{name}'이 여러 시트에 중복되어 있습니다.")
        items_all.update(src["items"])
        pivots.append((src["sheet"], pivot))
        groups.append(group_items(pivot, src["items"]))
        print(f"[피벗] '{src['sheet']}' {len(raw):,}행 → 기관 {len(pivot)}곳 × 값 열 {len(pivot.columns)}개")

    columns = [i["column"] for i in items_all.values() if "column" in i]
    dup = sorted({c for c in columns if columns.count(c) > 1})
    if dup:
        raise SystemExit(f"계산파일 열 {dup}에 여러 항목이 지정되어 있습니다.")

    grouped = pd.concat(groups, axis=1).fillna(0).astype("int64").sort_index()
    return pivots, grouped, items_all


def compare_items(grouped: pd.DataFrame, items: dict) -> list[str]:
    """compare_with 가 지정된 항목이 비교 대상과 기관별로 같은지 확인한다."""
    notes = []
    for name, item in items.items():
        other = item.get("compare_with")
        if not other:
            continue
        diff = grouped[name] - grouped[other]
        bad = diff[diff != 0]
        if bad.empty:
            notes.append(f"[검증] '{name}' = '{other}' (모든 기관 일치)")
        else:
            notes.append(
                f"[검증] '{name}'과 '{other}'이 {len(bad)}개 기관에서 다릅니다 "
                f"(차이 합계 {int(bad.sum()):,}). 검토용 피벗 파일 '검증' 시트를 확인하세요."
            )
    return notes


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="로우데이터를 피벗해서 계산파일에 넣습니다.")
    ap.add_argument("--raw", required=True, help="로우데이터 엑셀 (예: 1_raw_filled.xlsx)")
    ap.add_argument("--template", required=True, help="계산파일 양식 (예: 2.계산파일.xlsx)")
    ap.add_argument("--out", required=True, help="결과 계산파일 저장 경로")
    ap.add_argument("--pivot-out", help="검토용 피벗 결과 저장 경로 (기본: 결과 파일 옆 *_피벗.xlsx)")
    ap.add_argument("--config", default=str(DEFAULT_CONFIG), help="매핑 설정 파일 (기본: config/mapping.toml)")
    args = ap.parse_args(argv)

    with open(args.config, "rb") as f:
        cfg = tomllib.load(f)
    tgt = cfg["target"]

    # 1) 로우데이터 시트별 피벗
    pivots, grouped, items = build_pivots(args.raw, cfg)
    placed = {n: i for n, i in items.items() if "column" in i}

    # 2) 계산파일의 기관코드 행 찾기
    wb = XlsxPatcher(args.template)
    rows = range(tgt["first_row"], tgt["last_row"] + 1)
    code_rows = wb.read_column(tgt["sheet"], tgt["code_column"], rows)
    row_of = {code: r for r, code in code_rows.items()}
    missing = sorted(set(grouped.index) - set(row_of))
    if missing:
        raise SystemExit(
            f"계산파일 '{tgt['sheet']}'에 없는 기관코드가 로우데이터에 있습니다: {missing}\n"
            "계산파일에 행을 추가하거나 config 의 [code_merge] 에 합산 규칙을 넣어 주세요."
        )

    # 3) 값 + 합계 수식 만들기 (로우데이터에 없는 기관은 0)
    values: dict[str, int | Formula] = {}
    for code, r in row_of.items():
        for name, item in placed.items():
            values[f"{item['column']}{r}"] = int(grouped.at[code, name]) if code in grouped.index else 0
        for col, formula in cfg.get("row_formulas", {}).items():
            values[f"{col}{r}"] = Formula(formula.format(r=r))
    t = tgt["total_row"]
    for col in cfg.get("totals", {}).get("columns", []):
        values[f"{col}{t}"] = Formula(f"SUM({col}{tgt['first_row']}:{col}{tgt['last_row']})")

    wb.set_cells(tgt["sheet"], values)
    wb.request_full_recalc()
    wb.save(args.out)
    wb.close()
    print(f"[입력] '{tgt['sheet']}' 기관 {len(row_of)}곳 × 항목 {len(placed)}개 → {args.out}")

    # 4) 검증 + 검토용 피벗 파일
    notes = compare_items(grouped, items)
    pivot_out = args.pivot_out or str(Path(args.out).with_name(Path(args.out).stem + "_피벗.xlsx"))
    with pd.ExcelWriter(pivot_out) as xw:
        grouped.to_excel(xw, sheet_name="항목별 합계")
        for sheet, pivot in pivots:
            pivot.assign(총합계=pivot.sum(axis=1)).to_excel(xw, sheet_name=f"피벗_{sheet}"[:31])
        for name, item in items.items():
            if item.get("compare_with"):
                other = item["compare_with"]
                cmp = grouped[[name, other]].assign(차이=grouped[name] - grouped[other])
                cmp.to_excel(xw, sheet_name="검증")
        pd.DataFrame(
            [(n, i.get("column", "(입력 안 함)"), ", ".join(i["sources"]),
              "확정" if i.get("confirmed") else "임시(확인 필요)")
             for n, i in items.items()],
            columns=["항목", "계산파일 열", "포함 로우데이터 열", "상태"],
        ).to_excel(xw, sheet_name="매핑 규칙", index=False)
    print(f"[검토] 피벗 결과 → {pivot_out}")
    for note in notes:
        print(note)

    unconfirmed = [n for n, i in items.items() if not i.get("confirmed")]
    if unconfirmed:
        print(f"[주의] 임시 규칙 항목: {', '.join(unconfirmed)} (config/mapping.toml 확인)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
