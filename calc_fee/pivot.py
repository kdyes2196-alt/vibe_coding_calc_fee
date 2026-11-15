"""로우데이터 시트를 기관코드별로 피벗(합계)하고, 매핑 규칙에 따라 계산파일 항목으로 묶는다."""

from __future__ import annotations

import pandas as pd


def load_raw_sheet(path: str, sheet: str, key_column: str) -> pd.DataFrame:
    """로우데이터 시트를 읽는다. 기관코드는 앞자리 0이 사라지지 않게 문자열로 읽는다."""
    df = pd.read_excel(path, sheet_name=sheet, dtype={key_column: str})
    # 헤더에 숫자(212000 등)가 섞여 있으므로 모두 문자열로 통일
    df.columns = [str(c).strip() for c in df.columns]
    df = df[df[key_column].notna()].copy()
    df[key_column] = df[key_column].astype(str).str.strip()
    return df


def pivot_by_org(
    df: pd.DataFrame, key_column: str, ignore_columns: list[str], code_merge: dict[str, str]
) -> pd.DataFrame:
    """피벗테이블: 행 = 기관코드, 값 = 파일/전문 열별 합계."""
    value_columns = [c for c in df.columns if c != key_column and c not in ignore_columns]
    values = df[value_columns].apply(pd.to_numeric, errors="coerce").fillna(0)
    codes = df[key_column].replace(code_merge)
    pivot = values.groupby(codes).sum().sort_index()
    pivot.index.name = key_column
    return pivot.round().astype("int64")


def check_mapping(pivot_columns: list[str], items: dict) -> None:
    """모든 파일/전문 열이 정확히 한 항목에만 들어가는지 확인한다."""
    used: dict[str, str] = {}
    problems = []
    for name, item in items.items():
        for src in item["sources"]:
            if src in used:
                problems.append(f"'{src}' 열이 '{used[src]}'와 '{name}'에 중복으로 들어가 있습니다.")
            used[src] = name
            if src not in pivot_columns:
                problems.append(f"'{name}' 항목의 '{src}' 열이 로우데이터에 없습니다.")
    for col in pivot_columns:
        if col not in used:
            problems.append(f"로우데이터의 '{col}' 열이 어떤 항목에도 들어가 있지 않습니다.")
    if problems:
        raise ValueError("매핑 설정을 확인해 주세요:\n  - " + "\n  - ".join(problems))


def group_items(pivot: pd.DataFrame, items: dict) -> pd.DataFrame:
    """피벗 결과를 계산파일 항목(지로 신규등록 등)별로 합산한다."""
    out = pd.DataFrame(index=pivot.index)
    for name, item in items.items():
        out[name] = pivot[item["sources"]].sum(axis=1)
    return out
