"""기관 담당자 앞 산출내역 메일 발송.

- 담당자 목록: config/contacts.csv (기관코드, 기관명, 담당자, 이메일, 참조)
- 메일 설정  : config/mail.toml (SMTP 서버, 보내는 사람, 제목·본문 양식)
- 비밀번호는 파일에 쓰지 않고 환경변수(기본 CALC_FEE_SMTP_PASSWORD)로만 받습니다.

기본은 '미리보기' 입니다: 실제로 보내지 않고 .eml 파일과 발송목록만 만듭니다.
--send 를 붙여야 실제로 발송합니다.
"""

from __future__ import annotations

import csv
import mimetypes
import os
import re
import smtplib
import ssl
from dataclasses import dataclass
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr, make_msgid
from pathlib import Path

import pandas as pd

EMAIL_RE = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"


@dataclass
class Contact:
    code: str
    name: str
    person: str
    to: list[str]
    cc: list[str]


def _split(addrs: str) -> list[str]:
    return [a.strip() for a in (addrs or "").replace(";", ",").split(",") if a.strip()]


def load_contacts(path: str) -> dict[str, Contact]:
    df = pd.read_csv(path, dtype=str, encoding="utf-8-sig").fillna("")
    need = {"기관코드", "기관명", "담당자", "이메일"}
    if not need <= set(df.columns):
        raise ValueError(f"담당자 목록 열이 부족합니다. 필요한 열: {sorted(need)} (+ 선택: 참조)")
    out = {}
    for r in df.itertuples(index=False):
        code = r.기관코드.strip()
        if not code:
            continue
        if code in out:
            raise ValueError(f"담당자 목록에 기관코드 {code} 가 두 번 있습니다. 여러 명이면 이메일 칸에 쉼표로 적어 주세요.")
        out[code] = Contact(code, r.기관명.strip(), r.담당자.strip(), _split(r.이메일),
                            _split(getattr(r, "참조", "")))
    return out


def _render(template: str, values: dict) -> str:
    try:
        return template.format(**values)
    except KeyError as e:
        raise ValueError(f"메일 양식에 알 수 없는 항목 {{{e.args[0]}}} 가 있습니다. "
                         f"쓸 수 있는 항목: {', '.join(sorted(values))}")


def build_message(mail_cfg: dict, contact: Contact, values: dict, attachment: Path) -> EmailMessage:
    m = mail_cfg["message"]
    msg = EmailMessage()
    msg["Subject"] = _render(m["subject"], values)
    msg["From"] = formataddr((mail_cfg["sender"].get("name", ""), mail_cfg["sender"]["address"]))
    msg["To"] = ", ".join(contact.to)
    cc = contact.cc + _split(",".join(m.get("always_cc", [])))
    if cc:
        msg["Cc"] = ", ".join(cc)
    if m.get("reply_to"):
        msg["Reply-To"] = m["reply_to"]
    msg["Message-ID"] = make_msgid(domain=mail_cfg["sender"]["address"].split("@")[-1])
    msg.set_content(_render(m["body"], values))
    ctype, _ = mimetypes.guess_type(attachment.name)
    maintype, subtype = (ctype or "application/octet-stream").split("/", 1)
    msg.add_attachment(attachment.read_bytes(), maintype=maintype, subtype=subtype, filename=attachment.name)
    return msg


def _connect(smtp_cfg: dict) -> smtplib.SMTP:
    host, port = smtp_cfg["host"], int(smtp_cfg["port"])
    security = smtp_cfg.get("security", "starttls")
    timeout = smtp_cfg.get("timeout", 30)
    if security == "ssl":
        server: smtplib.SMTP = smtplib.SMTP_SSL(host, port, timeout=timeout, context=ssl.create_default_context())
    else:
        server = smtplib.SMTP(host, port, timeout=timeout)
        if security == "starttls":
            server.starttls(context=ssl.create_default_context())
    user = smtp_cfg.get("username")
    if user:
        pw = os.environ.get(smtp_cfg.get("password_env", "CALC_FEE_SMTP_PASSWORD"))
        if not pw:
            server.quit()
            raise RuntimeError(f"SMTP 비밀번호 환경변수 {smtp_cfg.get('password_env', 'CALC_FEE_SMTP_PASSWORD')} 가 비어 있습니다.")
        server.login(user, pw)
    return server


def run_mail(mail_cfg: dict, contacts: dict[str, Contact], summary: pd.DataFrame, totals: dict, period: str,
             org_files: dict[str, Path], out_dir: str, send: bool = False, only: list[str] | None = None,
             log=print) -> pd.DataFrame:
    """기관별 메일을 만들고, send=True 면 발송한다. 결과 표(발송결과.csv 로도 저장)를 돌려준다."""
    out = Path(out_dir)
    preview = out / "미리보기"
    preview.mkdir(parents=True, exist_ok=True)
    targets = only or list(summary.index)
    unknown = [c for c in targets if c not in summary.index]
    if unknown:
        raise ValueError(f"산출결과에 없는 기관코드: {unknown}")

    rows, messages = [], []
    for code in targets:
        r = summary.loc[code]
        base = {"기관코드": code, "기관명": r["기관명"], "담당자": "", "받는사람": "", "참조": "",
                "첨부파일": "", "결과": "", "내용": ""}
        c = contacts.get(code)
        if c is None:
            rows.append({**base, "결과": "제외", "내용": "담당자 목록에 없음"})
            continue
        bad = [a for a in c.to + c.cc if not re.match(EMAIL_RE, a)]
        if not c.to or bad:
            rows.append({**base, "담당자": c.person, "결과": "제외",
                         "내용": "이메일 없음" if not c.to else f"이메일 형식 오류: {', '.join(bad)}"})
            continue
        att = org_files.get(code)
        if att is None or not att.exists():
            rows.append({**base, "담당자": c.person, "결과": "제외", "내용": "기관별 산출자료 파일 없음"})
            continue
        values = {"기간": period, "기관코드": code, "기관명": r["기관명"], "담당자": c.person or "담당자",
                  "실적회비": f"{int(r['실적회비']):,}", "균등회비": f"{int(r['균등회비']):,}",
                  "회비합계": f"{int(r['회비합계']):,}", "실적비율": f"{float(r['실적비율']) * 100:.6f}%",
                  "실적합계": f"{int(r['실적합계']):,}", "회비총액": f"{totals.get('회비총액', 0):,}"}
        msg = build_message(mail_cfg, c, values, att)
        eml = preview / f"{code}_{r['기관명']}.eml".replace("/", "_")
        eml.write_bytes(bytes(msg))
        base.update({"담당자": c.person, "받는사람": msg["To"], "참조": msg.get("Cc", ""), "첨부파일": att.name})
        rows.append(base)
        messages.append((len(rows) - 1, msg))

    if send and messages:
        smtp_cfg = mail_cfg["smtp"]
        server = _connect(smtp_cfg)
        try:
            for idx, msg in messages:
                try:
                    refused = server.send_message(msg)
                    rows[idx]["결과"] = "발송" if not refused else "일부 실패"
                    rows[idx]["내용"] = "" if not refused else f"거부된 주소: {', '.join(refused)}"
                except smtplib.SMTPException as e:
                    rows[idx]["결과"], rows[idx]["내용"] = "실패", str(e)
                log(f"  {rows[idx]['결과']}: {rows[idx]['기관코드']} {rows[idx]['기관명']} → {rows[idx]['받는사람']}")
        finally:
            try:
                server.quit()
            except smtplib.SMTPException:
                pass
    else:
        for idx, _ in messages:
            rows[idx]["결과"] = "미리보기"

    df = pd.DataFrame(rows)
    name = f"발송결과_{datetime.now():%Y%m%d_%H%M%S}.csv" if send else "발송목록_미리보기.csv"
    df.to_csv(out / name, index=False, encoding="utf-8-sig", quoting=csv.QUOTE_MINIMAL)
    return df
