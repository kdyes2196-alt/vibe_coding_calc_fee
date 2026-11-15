"""③ 산출 결과 자동 검증.

1) 내부 정합성 : 로우데이터 합계 보존, 비율 합계 1, 회비 합계 일치, 설정값과 계산파일 일치 등
2) 전기 대비   : 지난번 산출결과_요약.xlsx 와 비교해 증감률이 큰 기관을 이상치로 표시
3) 교차검증    : LibreOffice 로 결과 계산파일의 수식을 실제로 계산한 값과 엔진 계산값을 기관별로 비교
                 (LibreOffice 가 없으면 '생략' 으로 표시)
"""

from __future__ import annotations

import glob
import os
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path

import openpyxl
import pandas as pd

from .engine import FeeResult
from .workbook import Aggregation, TemplateInfo

OK, WARN, FAIL, SKIP = "정상", "주의", "오류", "생략"


@dataclass
class Check:
    구분: str
    검증항목: str
    결과: str
    내용: str


# ---------------------------------------------------------------------------
# 1) 내부 정합성
# ---------------------------------------------------------------------------
def internal_checks(cfg: dict, agg: Aggregation, info: TemplateInfo, res: FeeResult) -> list[Check]:
    checks: list[Check] = []
    t, s = res.table, res.summary

    for sheet, pivot in agg.pivots:
        raw_total = agg.raw_totals[sheet]
        piv_total = int(pivot.to_numpy().sum())
        checks.append(Check("집계", f"로우데이터 합계 보존 ({sheet})", OK if raw_total == piv_total else FAIL,
                            f"로우데이터 {raw_total:,} / 피벗 {piv_total:,}"))

    placed = [n for n, i in agg.items.items() if "column" in i]
    in_book = int(agg.grouped.reindex(info.orgs.index).fillna(0)[placed].to_numpy().sum())
    all_placed = int(agg.grouped[placed].to_numpy().sum())
    checks.append(Check("집계", "계산파일 입력 누락 없음", OK if in_book == all_placed else FAIL,
                        f"집계 {all_placed:,} / 계산파일 입력 {in_book:,}"))

    for name, item in agg.items.items():
        other = item.get("compare_with")
        if other:
            diff = agg.grouped[name] - agg.grouped[other]
            n_bad = int((diff != 0).sum())
            checks.append(Check("집계", f"'{name}' = '{other}'", OK if n_bad == 0 else WARN,
                                "모든 기관 일치" if n_bad == 0 else
                                f"{n_bad}개 기관 불일치 (차이 합계 {int(diff.sum()):,}) → 피벗_검토.xlsx '비교' 시트"))

    ratio_sum = round(float(t["실적비율"].sum()), 10)
    checks.append(Check("회비", "실적 비율 합계 = 1", OK if ratio_sum == 1 else FAIL, f"합계 {ratio_sum}"))
    pf = int(t["실적회비"].sum())
    checks.append(Check("회비", "실적회비 합계 = 실적회비 총액", OK if pf == s["실적회비총액"] else FAIL,
                        f"{pf:,} / {s['실적회비총액']:,}"))
    tot = int(t["회비합계"].sum())
    checks.append(Check("회비", "기관별 회비 합계 = 자동납부분", OK if tot == s["자동납부분"] else FAIL,
                        f"{tot:,} / {s['자동납부분']:,}"))
    checks.append(Check("회비", "회비 총액 = 자동송금분 + 자동납부분",
                        OK if s["자동송금분"] + s["자동납부분"] == s["회비총액"] else FAIL,
                        f"{s['자동송금분']:,} + {s['자동납부분']:,} = {s['회비총액']:,}"))

    n_cfg, n_real = s["참가기관수"], s["실제기관수"]
    checks.append(Check("설정", "균등회비 나누는 기관 수 = 실제 기관 수", OK if n_cfg == n_real else WARN,
                        f"설정 {n_cfg}곳 / 계산파일 {n_real}곳"
                        + ("" if n_cfg == n_real else " → config [fee].n_orgs 와 계산파일 G3·H열 수식 확인")))
    tr = info.template_remit_rate
    checks.append(Check("설정", "자동송금 비율: 설정 = 계산파일",
                        SKIP if tr is None else (OK if abs(tr - s["자동송금비율"]) < 1e-12 else FAIL),
                        f"설정 {s['자동송금비율']} / 계산파일 {cfg['fee']['remit_rate_cell']} {tr}"))

    zero = t.index[t["실적합계"] == 0].tolist()
    checks.append(Check("회비", "실적이 0인 기관 없음", OK if not zero else WARN,
                        "없음" if not zero else f"{len(zero)}곳: " + ", ".join(f"{c} {t.at[c, '기관명']}" for c in zero)))

    no_raw = [c for c in t.index if c not in agg.grouped.index]
    if no_raw:
        checks.append(Check("집계", "로우데이터에 없는 기관 (계좌 실적 0 처리)", WARN,
                            f"{len(no_raw)}곳: " + ", ".join(f"{c} {t.at[c, '기관명']}" for c in no_raw)))

    if cfg.get("dummy_fill"):
        where = ", ".join(f"{f['sheet']} {f['column']}열" for f in cfg["dummy_fill"])
        checks.append(Check("설정", "임시 난수 데이터 사용", WARN, f"{where} (실제 데이터 아님, config [[dummy_fill]])"))
    return checks


# ---------------------------------------------------------------------------
# 2) 전기 대비
# ---------------------------------------------------------------------------
def load_summary(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="기관별 산출결과", dtype={"기관코드": str})
    return df.set_index("기관코드")


def compare_previous(cfg: dict, res: FeeResult, prev_path: str | None) -> tuple[list[Check], pd.DataFrame | None]:
    if not prev_path:
        return [Check("전기 대비", "전기 산출결과와 비교", SKIP, "--prev 로 지난 산출결과_요약.xlsx 를 지정하면 비교합니다")], None
    prev = load_summary(prev_path)
    th = cfg["validation"]["change_threshold"]
    min_amt = cfg["validation"]["min_amount_for_change"]
    cur = res.table
    codes = list(cur.index) + [c for c in prev.index if c not in cur.index]
    rows = []
    for c in codes:
        name = cur.at[c, "기관명"] if c in cur.index else prev.at[c, "기관명"]
        row = {"기관코드": c, "기관명": name}
        flags = []
        for col in ["실적합계", "회비합계"]:
            p = int(prev.at[c, col]) if c in prev.index else None
            n = int(cur.at[c, col]) if c in cur.index else None
            row[f"전기 {col}"], row[f"당기 {col}"] = p, n
            if p is not None and n is not None:
                row[f"{col} 증감"] = n - p
                rate = (n - p) / p if p else None
                row[f"{col} 증감률"] = rate
                if rate is not None and abs(rate) > th and (col != "회비합계" or p >= min_amt):
                    flags.append(f"{col} {rate:+.1%}")
        if c not in prev.index:
            flags.append("신규 기관")
        if c not in cur.index:
            flags.append("이번 산출에서 빠짐")
        row["이상 여부"] = ", ".join(flags)
        rows.append(row)
    df = pd.DataFrame(rows)
    flagged = df[df["이상 여부"] != ""]
    check = Check("전기 대비", f"증감률 ±{th:.0%} 초과 기관", OK if flagged.empty else WARN,
                  f"전기 파일: {Path(prev_path).name} / 이상 {len(flagged)}곳"
                  + ("" if flagged.empty else ": " + ", ".join(f"{r.기관코드} {r.기관명}" for r in flagged.head(10).itertuples())
                     + (" 외" if len(flagged) > 10 else "")))
    return [check], df


# ---------------------------------------------------------------------------
# 3) 교차검증 (LibreOffice 로 계산파일 수식을 실제로 계산)
# ---------------------------------------------------------------------------
def find_soffice() -> str | None:
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    for pattern in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                    "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        for p in glob.glob(pattern):
            return p
    return None


def recalc_with_libreoffice(xlsx: str, soffice: str, timeout: int = 180) -> str:
    """xlsx 를 LibreOffice 로 열어 수식을 계산한 복사본을 만들고 그 경로를 돌려준다."""
    out_dir = tempfile.mkdtemp(prefix="calc_fee_recalc_")
    profile_dir = tempfile.mkdtemp(prefix="calc_fee_lo_profile_")   # 사용 중인 LibreOffice 와 충돌 방지
    try:
        subprocess.run(
            [soffice, f"-env:UserInstallation={Path(profile_dir).as_uri()}", "--headless",
             "--convert-to", "xlsx", "--outdir", out_dir, os.path.abspath(xlsx)],
            check=True, capture_output=True, timeout=timeout,
        )
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)
    out = os.path.join(out_dir, Path(xlsx).name)
    if not os.path.exists(out):
        raise RuntimeError("LibreOffice 변환 결과 파일이 만들어지지 않았습니다.")
    return out


def crosscheck(cfg: dict, info: TemplateInfo, res: FeeResult, result_xlsx: str, mode: str = "auto"
               ) -> tuple[list[Check], pd.DataFrame | None]:
    if mode == "off":
        return [Check("교차검증", "계산파일 수식 결과 = 엔진 계산", SKIP, "--crosscheck off")], None
    soffice = find_soffice()
    if not soffice:
        msg = "LibreOffice 가 없어 생략 (설치하면 자동으로 교차검증합니다)"
        if mode == "on":
            raise RuntimeError(msg)
        return [Check("교차검증", "계산파일 수식 결과 = 엔진 계산", SKIP, msg)], None

    computed = recalc_with_libreoffice(result_xlsx, soffice)
    wb = openpyxl.load_workbook(computed, data_only=True)
    tgt, fee = cfg["target"], cfg["fee"]
    ws_t, ws_f = wb[tgt["sheet"]], wb[fee["sheet"]]
    # 엔진 결과 열 → 계산파일 회비 시트 열
    fee_cols = {"계좌자동납부": "C", "카드이동": "D", "실적합계": "E", "실적비율": "F",
                "실적회비": "G", "균등회비": "H", "회비합계": "I"}
    rows, n_bad = [], 0
    for code, r in res.table.iterrows():
        fr, tr = info.orgs.at[code, "회비행"], info.orgs.at[code, "실적행"]
        cells = {k: ws_f[f"{c}{fr}"].value for k, c in fee_cols.items()}
        cells["실적시트 계좌합계"] = ws_t[f"R{tr}"].value
        bad = []
        for k, v in cells.items():
            exp = r["계좌자동납부"] if k == "실적시트 계좌합계" else r[k]
            if not isinstance(v, (int, float)) or abs(v - exp) > (1e-12 if k == "실적비율" else 0.5):
                bad.append(f"{k}: 계산파일 {v} / 엔진 {exp}")
        n_bad += bool(bad)
        rows.append({"기관코드": code, "기관명": r["기관명"],
                     "계산파일 회비합계": cells["회비합계"], "엔진 회비합계": int(r["회비합계"]),
                     "결과": "일치" if not bad else "불일치", "내용": "; ".join(bad)})
    errors = [c.coordinate for ws in (ws_t, ws_f) for row in ws.iter_rows() for c in row
              # '1. 실적' U열(카드이동서비스)은 쓰지 않는 열이라 원래 있던 #REF! 를 제외
              if isinstance(c.value, str) and c.value.startswith("#")
              and not (ws is ws_t and c.column_letter == "U")]
    df = pd.DataFrame(rows)
    checks = [
        Check("교차검증", "계산파일 수식 결과 = 엔진 계산 (기관별)", OK if n_bad == 0 else FAIL,
              f"{len(rows)}곳 중 {len(rows) - n_bad}곳 일치"),
        Check("교차검증", "계산파일 수식 오류(#REF! 등) 없음", OK if not errors else FAIL,
              "없음" if not errors else f"{len(errors)}칸: {', '.join(errors[:10])}"),
    ]
    shutil.rmtree(os.path.dirname(computed), ignore_errors=True)
    return checks, df
