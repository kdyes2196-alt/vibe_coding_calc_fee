"""회비 계산 엔진: 계산파일 '2. 납부 회비 산출' 과 같은 순서로 회비를 따로 계산한다.

계산파일의 수식 결과와 이 엔진의 결과를 비교(교차검증)하면, 사람이 하던 2인 교차검증을 대신할 수 있다.
엑셀 ROUND 와 같게 하려고 Decimal 의 사사오입(ROUND_HALF_UP)을 쓴다.
"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal

import pandas as pd


def excel_round(x: Decimal, digits: int = 0) -> Decimal:
    return x.quantize(Decimal(1).scaleb(-digits), rounding=ROUND_HALF_UP)


def adjust_rounding(base: list[Decimal], target: Decimal, unit: Decimal, rank_key: list) -> list[Decimal]:
    """반올림 값 합계를 target 에 맞춘다. rank_key 가 큰 순서(같으면 앞쪽)부터 unit 씩 더하거나 뺀다."""
    n = int(excel_round((target - sum(base)) / unit))
    if n == 0:
        return list(base)
    order = sorted(range(len(base)), key=lambda i: (-rank_key[i], i))
    out = list(base)
    step = unit if n > 0 else -unit
    for i in order[: abs(n)]:
        out[i] += step
    return out


@dataclass
class FeeResult:
    table: pd.DataFrame       # 기관별 산출 결과
    summary: dict             # 총액 내역


# 결과 표 열: 기관명, [summary.groups 의 묶음들], 아래 열들
RESULT_TAIL = ["계좌자동납부", "카드이동", "실적합계", "실적비율", "실적회비", "균등회비", "회비합계"]


def compute_fees(cfg: dict, orgs: pd.DataFrame, grouped: pd.DataFrame, card: pd.Series) -> FeeResult:
    """orgs: index=기관코드(계산파일 순서), grouped: 기관코드 × 항목, card: 기관코드 → 카드이동 실적."""
    fee, groups = cfg["fee"], cfg["summary"]["groups"]
    g = grouped.reindex(orgs.index).fillna(0).astype("int64")

    t = pd.DataFrame(index=orgs.index)
    t["기관명"] = orgs["기관명"]
    for name, item_names in groups.items():
        t[name] = g[item_names].sum(axis=1)
    t["계좌자동납부"] = t[list(groups)].sum(axis=1)
    t["카드이동"] = card.reindex(orgs.index).fillna(0).astype("int64")
    t["실적합계"] = t["계좌자동납부"] + t["카드이동"]

    # 총액 배분 (계산파일 E2~G5)
    total = Decimal(fee["total_fee"])
    remit_rate = Decimal(str(fee["remit_rate"]))
    auto = excel_round(total * (Decimal(1) - remit_rate))                   # E3 자동납부분
    remit = total - auto                                                    # E2 자동송금분
    equal_before = auto * Decimal(str(fee["equal_rate"]))                   # G2 수정 전 균등
    n_orgs = Decimal(fee["n_orgs"])
    equal_total = excel_round(equal_before / n_orgs) * n_orgs               # G3 균등회비
    perf_total = auto - equal_total                                         # G4 실적회비

    # 실적 비율 (F열) + 단수 조정
    perf = [Decimal(int(v)) for v in t["실적합계"]]
    perf_sum = sum(perf)
    d = fee["ratio_decimals"]
    if perf_sum == 0:
        raise ValueError("실적 합계가 0 이라 비율을 계산할 수 없습니다.")
    ratio = adjust_rounding([excel_round(p / perf_sum, d) for p in perf], Decimal(1), Decimal(1).scaleb(-d), perf)
    # 실적회비 (G열) + 단수 조정
    perf_fee = adjust_rounding([excel_round(r * perf_total) for r in ratio], perf_total, Decimal(1), perf)
    equal_each = equal_total / n_orgs                                       # H열

    t["실적비율"] = [float(r) for r in ratio]
    t["실적회비"] = [int(x) for x in perf_fee]
    t["균등회비"] = int(equal_each)
    t["회비합계"] = t["실적회비"] + t["균등회비"]

    summary = {
        "회비총액": int(total), "자동송금분": int(remit), "자동납부분": int(auto),
        "자동송금비율": float(remit_rate), "균등회비총액": int(equal_total), "실적회비총액": int(perf_total),
        "균등회비비율": float(fee["equal_rate"]), "참가기관수": int(n_orgs), "기관당균등회비": int(equal_each),
        "실적합계": int(perf_sum), "실제기관수": len(t),
    }
    summary["실적묶음"] = list(groups)
    return FeeResult(table=t[["기관명", *groups, *RESULT_TAIL]], summary=summary)
