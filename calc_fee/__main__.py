"""회비 산출·검증 자동화 프로그램

  산출 : python -m calc_fee run  --raw 로우데이터.xlsx --template 계산파일.xlsx --period "2026년 3분기" --out-dir out/2026Q3
  발송 : python -m calc_fee mail --out-dir out/2026Q3            (미리보기만, 실제 발송 안 함)
         python -m calc_fee mail --out-dir out/2026Q3 --send     (실제 발송)
"""

from __future__ import annotations

import argparse
import json
import sys
import tomllib
from pathlib import Path

from . import reports, validate
from .engine import compute_fees
from .mailer import load_contacts, run_mail
from .validate import FAIL, WARN
from .workbook import CalcFeeError, build_pivots, fill_template

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_CONFIG = ROOT / "config" / "mapping.toml"
DEFAULT_MAIL = ROOT / "config" / "mail.toml"
DEFAULT_CONTACTS = ROOT / "config" / "contacts.csv"
INFO_FILE = "산출정보.json"

OUT_FILES = {
    "result": "계산파일_결과.xlsx",
    "pivot": "피벗_검토.xlsx",
    "summary": "산출결과_요약.xlsx",
    "validation": "검증결과.xlsx",
    "hwp": "산출내역_한글용.docx",
    "orgs": "기관별",
    "mail": "메일",
}


def load_toml(path: str | Path) -> dict:
    with open(path, "rb") as f:
        return tomllib.load(f)


# ---------------------------------------------------------------------------
def cmd_run(args) -> int:
    cfg = load_toml(args.config)
    out = Path(args.out_dir)
    out.mkdir(parents=True, exist_ok=True)
    paths = {k: out / v for k, v in OUT_FILES.items()}
    period = args.period

    print(f"■ ① 실적 집계 — {period}")
    agg = build_pivots(args.raw, cfg)
    info = fill_template(cfg, agg, args.template, str(paths["result"]))
    reports.write_pivot_review(agg, str(paths["pivot"]))

    print("■ 회비 계산")
    res = compute_fees(cfg, info.orgs, agg.grouped, info.card)
    s = res.summary
    print(f"  회비 총액 {s['회비총액']:,}원 = 자동송금분 {s['자동송금분']:,} + 자동납부분 {s['자동납부분']:,}"
          f" (균등 {s['균등회비총액']:,} / 실적 {s['실적회비총액']:,})")

    print("■ ③ 산출 결과 검증")
    checks = validate.internal_checks(cfg, agg, info, res)
    prev_checks, prev_df = validate.compare_previous(cfg, res, args.prev)
    try:
        cross_checks, cross_df = validate.crosscheck(cfg, info, res, str(paths["result"]), args.crosscheck)
    except Exception as e:  # LibreOffice 실행 실패 등: 교차검증만 실패로 기록하고 계속
        cross_checks = [validate.Check("교차검증", "계산파일 수식 결과 = 엔진 계산", FAIL, f"실행 실패: {e}")]
        cross_df = None
    checks += prev_checks + cross_checks
    reports.write_validation(checks, prev_df, cross_df, str(paths["validation"]), period)
    for c in checks:
        print(f"  [{c.결과}] {c.검증항목} — {c.내용}")

    print("■ ④ 산출자료 생성")
    reports.write_summary(res, period, str(paths["summary"]))
    reports.write_hwp_doc(res, period, str(paths["hwp"]))
    org_files = reports.write_org_files(cfg, res, agg, period, str(paths["orgs"]))

    n_fail = sum(c.결과 == FAIL for c in checks)
    n_warn = sum(c.결과 == WARN for c in checks)
    (out / INFO_FILE).write_text(json.dumps({
        "period": period, "totals": s, "fail_count": n_fail, "warn_count": n_warn,
        "org_files": {c: str(p.relative_to(out)) for c, p in org_files.items()},
    }, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"\n완료: {out}/")
    for k in ("result", "summary", "validation", "hwp", "pivot"):
        print(f"  - {OUT_FILES[k]}")
    print(f"  - {OUT_FILES['orgs']}/ ({len(org_files)}개 기관)")
    print(f"검증: 오류 {n_fail}건, 주의 {n_warn}건 → {OUT_FILES['validation']} 확인")
    return 1 if n_fail else 0


# ---------------------------------------------------------------------------
def cmd_mail(args) -> int:
    out = Path(args.out_dir)
    info_path = out / INFO_FILE
    if not info_path.exists():
        raise CalcFeeError(f"{info_path} 가 없습니다. 먼저 run 으로 산출을 실행하세요.")
    info = json.loads(info_path.read_text(encoding="utf-8"))
    if args.send and info["fail_count"] and not args.ignore_errors:
        raise CalcFeeError(f"검증 오류가 {info['fail_count']}건 있어 발송하지 않습니다. "
                           f"{OUT_FILES['validation']} 를 확인하세요. (그래도 보내려면 --ignore-errors)")

    mail_cfg = load_toml(args.mail_config)
    contacts = load_contacts(args.contacts)
    summary = validate.load_summary(str(out / OUT_FILES["summary"]))
    org_files = {c: out / p for c, p in info["org_files"].items()}
    only = [c.strip() for c in args.only.split(",")] if args.only else None

    if args.send:
        n = len(only or summary.index)
        print(f"실제 발송: {n}개 기관 대상, SMTP {mail_cfg['smtp']['host']}:{mail_cfg['smtp']['port']}")
        if not args.yes and input("발송하려면 '발송' 이라고 입력하세요: ").strip() != "발송":
            print("취소했습니다.")
            return 1

    df = run_mail(mail_cfg, contacts, summary, info["totals"], info["period"], org_files,
                  str(out / OUT_FILES["mail"]), send=args.send, only=only)
    counts = df["결과"].value_counts().to_dict()
    print("결과: " + ", ".join(f"{k} {v}건" for k, v in counts.items()))
    excluded = df[df["결과"] == "제외"]
    for r in excluded.itertuples():
        print(f"  제외: {r.기관코드} {r.기관명} — {r.내용}")
    where = out / OUT_FILES["mail"]
    print(f"{'발송결과' if args.send else '미리보기(.eml)·발송목록'}: {where}/")
    failed = counts.get("실패", 0) + counts.get("일부 실패", 0)
    return 1 if failed else 0


# ---------------------------------------------------------------------------
def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="python -m calc_fee", description="회비 산출·검증 자동화 프로그램")
    sub = ap.add_subparsers(dest="command", required=True)

    r = sub.add_parser("run", help="실적 집계 → 회비 계산 → 검증 → 산출자료 생성")
    r.add_argument("--raw", required=True, help="로우데이터 엑셀")
    r.add_argument("--template", required=True, help="계산파일 양식 (원본은 바뀌지 않음)")
    r.add_argument("--period", required=True, help='산출 기간 이름 (예: "2026년 3분기")')
    r.add_argument("--out-dir", required=True, help="결과를 저장할 폴더")
    r.add_argument("--prev", help="전기 산출결과_요약.xlsx (전기 대비 이상치 검증)")
    r.add_argument("--crosscheck", choices=["auto", "on", "off"], default="auto",
                   help="LibreOffice 교차검증 (auto: 설치되어 있으면 실행)")
    r.add_argument("--config", default=str(DEFAULT_CONFIG), help="산출 기준 설정 파일")
    r.set_defaults(func=cmd_run)

    m = sub.add_parser("mail", help="기관 담당자 앞 산출내역 메일 (기본: 미리보기)")
    m.add_argument("--out-dir", required=True, help="run 결과 폴더")
    m.add_argument("--send", action="store_true", help="실제로 발송 (없으면 .eml 미리보기만 생성)")
    m.add_argument("--only", help="일부 기관만 (기관코드 쉼표 구분, 예: 002,003)")
    m.add_argument("--yes", action="store_true", help="발송 확인 질문 생략")
    m.add_argument("--ignore-errors", action="store_true", help="검증 오류가 있어도 발송")
    m.add_argument("--contacts", default=str(DEFAULT_CONTACTS), help="담당자 목록 CSV")
    m.add_argument("--mail-config", default=str(DEFAULT_MAIL), help="메일 설정 파일")
    m.set_defaults(func=cmd_mail)

    args = ap.parse_args(argv)
    try:
        return args.func(args)
    except CalcFeeError as e:
        print(f"[오류] {e}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
