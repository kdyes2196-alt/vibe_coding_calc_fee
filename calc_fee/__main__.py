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
    src, tgt, items = cfg["source"], cfg["target"], cfg["items"]

    # 1) 피벗
    raw = load_raw_sheet(args.raw, src["sheet"], src["key_column"])
    pivot = pivot_by_org(raw, src["key_column"], src.get("ignore_columns", []), cfg.get("code_merge", {}))
    check_mapping(list(pivot.columns), items)
    grouped = group_items(pivot, items)
    print(f"[피벗] '{src['sheet']}' {len(raw):,}행 → 기관 {len(pivot)}곳 × 파일/전문 {len(pivot.columns)}종")

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
        for name, item in items.items():
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
    print(f"[입력] '{tgt['sheet']}' 기관 {len(row_of)}곳 × 항목 {len(items)}개 → {args.out}")

    # 4) 검토용 피벗 파일
    pivot_out = args.pivot_out or str(Path(args.out).with_name(Path(args.out).stem + "_피벗.xlsx"))
    with pd.ExcelWriter(pivot_out) as xw:
        pivot.assign(총합계=pivot.sum(axis=1)).to_excel(xw, sheet_name="피벗(파일전문별)")
        grouped.to_excel(xw, sheet_name="항목별 합계")
        pd.DataFrame(
            [(n, i["column"], ", ".join(i["sources"]), "확정" if i.get("confirmed") else "임시(확인 필요)")
             for n, i in items.items()],
            columns=["항목", "계산파일 열", "포함 파일/전문", "상태"],
        ).to_excel(xw, sheet_name="매핑 규칙", index=False)
    print(f"[검토] 피벗 결과 → {pivot_out}")

    unconfirmed = [n for n, i in items.items() if not i.get("confirmed")]
    if unconfirmed:
        print(f"[주의] 임시 규칙 항목: {', '.join(unconfirmed)} (config/mapping.toml 확인)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
