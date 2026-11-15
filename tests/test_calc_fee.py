"""테스트 실행: python -m unittest discover -s tests -v"""

from __future__ import annotations

import socket
import socketserver
import tempfile
import threading
import tomllib
import unittest
from decimal import Decimal
from email import message_from_bytes, policy
from pathlib import Path

import pandas as pd

from calc_fee.engine import adjust_rounding, compute_fees, excel_round
from calc_fee.mailer import Contact, run_mail
from calc_fee.pivot import check_mapping, pivot_by_org

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"


# ---------------------------------------------------------------------------
class TestRounding(unittest.TestCase):
    def test_excel_round_half_up(self):
        self.assertEqual(excel_round(Decimal("2.5")), Decimal("3"))
        self.assertEqual(excel_round(Decimal("0.123456785"), 8), Decimal("0.12345679"))

    def test_adjust_adds_to_largest_first(self):
        base = [Decimal(10), Decimal(30), Decimal(20)]
        out = adjust_rounding(base, Decimal(62), Decimal(1), rank_key=[1, 3, 2])
        self.assertEqual(out, [Decimal(10), Decimal(31), Decimal(21)])

    def test_adjust_subtracts_when_over(self):
        out = adjust_rounding([Decimal(10), Decimal(30)], Decimal(39), Decimal(1), rank_key=[1, 3])
        self.assertEqual(out, [Decimal(10), Decimal(29)])

    def test_adjust_tie_goes_to_upper_row(self):
        out = adjust_rounding([Decimal(5), Decimal(5)], Decimal(11), Decimal(1), rank_key=[7, 7])
        self.assertEqual(out, [Decimal(6), Decimal(5)])


class TestPivot(unittest.TestCase):
    def test_pivot_merges_codes_and_sums(self):
        df = pd.DataFrame({"code": ["081", "005", "002"], "d": ["1", "1", "1"], "A": [1, 2, 3], "B": [10, 20, 30]})
        p = pivot_by_org(df, "code", ["d"], {"005": "081"})
        self.assertEqual(p.loc["081", "A"], 3)
        self.assertEqual(p.loc["081", "B"], 30)

    def test_mapping_detects_missing_and_duplicate(self):
        with self.assertRaises(ValueError):
            check_mapping(["A", "B"], {"x": {"sources": ["A"]}})
        with self.assertRaises(ValueError):
            check_mapping(["A"], {"x": {"sources": ["A"]}, "y": {"sources": ["A"]}})


class TestEngine(unittest.TestCase):
    def test_totals_are_exact(self):
        cfg = {
            "fee": {"total_fee": 1_000_000_007, "remit_rate": 0.03, "equal_rate": 0.25, "n_orgs": 3,
                    "ratio_decimals": 8},
            "summary": {"groups": {"지로": ["a"]}},
        }
        orgs = pd.DataFrame({"기관명": ["가", "나", "다"]}, index=["001", "002", "003"])
        grouped = pd.DataFrame({"a": [333, 333, 334]}, index=orgs.index)
        card = pd.Series([0, 1, 2], index=orgs.index)
        res = compute_fees(cfg, orgs, grouped, card)
        s, t = res.summary, res.table
        self.assertEqual(s["자동송금분"] + s["자동납부분"], s["회비총액"])
        self.assertAlmostEqual(t["실적비율"].sum(), 1.0, places=12)
        self.assertEqual(t["실적회비"].sum(), s["실적회비총액"])
        self.assertEqual(t["회비합계"].sum(), s["자동납부분"])


# ---------------------------------------------------------------------------
class _SMTPHandler(socketserver.StreamRequestHandler):
    """메일 발송 테스트용 최소 SMTP 서버 (이 컴퓨터 안에서만 동작, 외부로 나가지 않음)."""

    def handle(self):
        self.wfile.write(b"220 test\r\n")
        mail = {"rcpt": [], "data": b""}
        while True:
            line = self.rfile.readline()
            if not line:
                return
            cmd = line.strip().upper()
            if cmd.startswith(b"EHLO") or cmd.startswith(b"HELO"):
                self.wfile.write(b"250 ok\r\n")
            elif cmd.startswith(b"MAIL FROM"):
                mail = {"rcpt": [], "data": b""}
                self.wfile.write(b"250 ok\r\n")
            elif cmd.startswith(b"RCPT TO"):
                addr = line.split(b":", 1)[1].strip().strip(b"<>").decode()
                if addr.startswith("reject"):
                    self.wfile.write(b"550 no such user\r\n")
                else:
                    mail["rcpt"].append(addr)
                    self.wfile.write(b"250 ok\r\n")
            elif cmd == b"DATA":
                self.wfile.write(b"354 go\r\n")
                chunks = []
                while (l := self.rfile.readline()) not in (b".\r\n", b""):
                    chunks.append(l[1:] if l.startswith(b"..") else l)
                mail["data"] = b"".join(chunks)
                self.server.received.append(mail)
                self.wfile.write(b"250 queued\r\n")
            elif cmd == b"QUIT":
                self.wfile.write(b"221 bye\r\n")
                return
            else:
                self.wfile.write(b"250 ok\r\n")


class TestMailer(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.att = self.tmp / "002_산업은행.xlsx"
        self.att.write_bytes(b"dummy")
        self.summary = pd.DataFrame(
            {"기관명": ["산업은행", "기업은행", "국민은행"], "실적회비": [100, 200, 300], "균등회비": [10, 10, 10],
             "회비합계": [110, 210, 310], "실적비율": [0.1, 0.2, 0.7], "실적합계": [1, 2, 7]},
            index=pd.Index(["002", "003", "004"], name="기관코드"))
        self.contacts = {
            "002": Contact("002", "산업은행", "홍담당", ["a@example.com"], ["cc@example.com"]),
            "003": Contact("003", "기업은행", "김담당", ["reject@example.com"], []),
        }
        self.files = {"002": self.att, "003": self.att}
        with open(ROOT / "config" / "mail.toml", "rb") as f:
            self.cfg = tomllib.load(f)

    def test_preview_does_not_connect(self):
        self.cfg["smtp"]["host"] = "invalid.invalid"   # 접속하면 실패하므로, 미리보기는 접속하지 않아야 함
        df = run_mail(self.cfg, self.contacts, self.summary, {"회비총액": 1}, "2026년 3분기", self.files,
                      str(self.tmp / "mail"), send=False, log=lambda *a: None)
        self.assertEqual(df.set_index("기관코드").loc["002", "결과"], "미리보기")
        self.assertEqual(df.set_index("기관코드").loc["004", "결과"], "제외")   # 담당자 없음
        eml = (self.tmp / "mail" / "미리보기" / "002_산업은행.eml").read_bytes()
        msg = message_from_bytes(eml, policy=policy.default)
        self.assertIn("산업은행", msg["Subject"])
        self.assertIn("110원", msg.get_body(("plain",)).get_content())
        self.assertEqual([a.get_filename() for a in msg.iter_attachments()], ["002_산업은행.xlsx"])

    def test_send_through_local_smtp(self):
        server = socketserver.ThreadingTCPServer(("127.0.0.1", 0), _SMTPHandler)
        server.received = []
        threading.Thread(target=server.serve_forever, daemon=True).start()
        try:
            self.cfg["smtp"].update(host="127.0.0.1", port=server.server_address[1], security="none", username="")
            df = run_mail(self.cfg, self.contacts, self.summary, {"회비총액": 1}, "2026년 3분기", self.files,
                          str(self.tmp / "mail"), send=True, log=lambda *a: None).set_index("기관코드")
        finally:
            server.shutdown()
            server.server_close()
        self.assertEqual(df.loc["002", "결과"], "발송")
        self.assertEqual(df.loc["003", "결과"], "실패")       # 서버가 받는 사람을 거부
        self.assertEqual(df.loc["004", "결과"], "제외")
        self.assertEqual(len(server.received), 1)
        self.assertEqual(sorted(server.received[0]["rcpt"]), ["a@example.com", "cc@example.com"])
        self.assertTrue(list((self.tmp / "mail").glob("발송결과_*.csv")))


# ---------------------------------------------------------------------------
@unittest.skipUnless((DATA / "1_raw_filled.xlsx").exists() and (DATA / "2_계산파일.xlsx").exists(),
                     "data/ 에 샘플 로우데이터·계산파일이 있어야 실행")
class TestEndToEnd(unittest.TestCase):
    def test_run_produces_consistent_outputs(self):
        from calc_fee.__main__ import main
        out = Path(tempfile.mkdtemp()) / "out"
        code = main(["run", "--raw", str(DATA / "1_raw_filled.xlsx"), "--template", str(DATA / "2_계산파일.xlsx"),
                     "--period", "테스트", "--out-dir", str(out), "--crosscheck", "auto"])
        self.assertEqual(code, 0)
        summary = pd.read_excel(out / "산출결과_요약.xlsx", sheet_name="기관별 산출결과", dtype={"기관코드": str})
        self.assertEqual(len(summary), 64)
        self.assertEqual(summary["회비합계"].sum(), 1_422_008_360)
        self.assertEqual(len(list((out / "기관별").glob("*.xlsx"))), 64)
        self.assertTrue((out / "산출내역_한글용.docx").exists())
        val = pd.read_excel(out / "검증결과.xlsx", sheet_name="검증 요약", header=3)
        self.assertFalse((val["결과"] == "오류").any())


if __name__ == "__main__":
    unittest.main()
