"""Tests for the Excel / CSV Data Cleaner (unittest style; also runs under pytest).

Run from the repository root:  python -m unittest discover -s tests -v
Fixtures are small fixed files in tests/fixtures/ (rebuild: python tests/make_fixtures.py).
"""
import contextlib
import hashlib
import io
import sys
import tempfile
import time
import unittest
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from datacleaner import normalize as N  # noqa: E402
from datacleaner.pipeline import run  # noqa: E402

FIXTURES = Path(__file__).resolve().parent / "fixtures"


class CleanerTestCase(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.tmp = Path(self._tmp.name)

    def run_cleaner(self, *inputs, out=None, date_order=None):
        out = Path(out or self.tmp / "out")
        with contextlib.redirect_stdout(io.StringIO()):
            code = run([str(i) for i in inputs], str(out), date_order)
        return code, out

    @staticmethod
    def sheet(out, file, sheet=None):
        if file.endswith(".csv"):
            df = pd.read_csv(out / file, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        else:
            df = pd.read_excel(out / file, sheet_name=sheet, dtype=str, keep_default_na=False)
        return df.to_dict("records")

    def by_row(self, out):
        """{source_row: (status, record)} across all four output categories."""
        result = {}
        for status, file, sheet in [("clean", "cleaned_data.csv", None),
                                    ("invalid", "errors.xlsx", "invalid"),
                                    ("duplicate", "errors.xlsx", "duplicates"),
                                    ("needs_review", "needs_review.xlsx", "needs_review")]:
            for rec in self.sheet(out, file, sheet):
                result[int(rec["source_row"])] = (status, rec)
        return result

    def summary_text(self, out):
        return (out / "summary.txt").read_text(encoding="utf-8")


class TestPipeline(CleanerTestCase):
    def test_01_normal_csv(self):
        code, out = self.run_cleaner(FIXTURES / "normal.csv")
        self.assertEqual(code, 0)
        for name in ("cleaned_data.xlsx", "cleaned_data.csv", "errors.xlsx", "needs_review.xlsx",
                     "change_log.xlsx", "summary.xlsx", "summary.txt", "process.log"):
            self.assertTrue((out / name).exists(), name)
        rows = self.by_row(out)
        self.assertEqual({s for s, _ in rows.values()}, {"clean"})
        _, r2 = rows[3]
        self.assertEqual((r2["name"], r2["email"], r2["phone"], r2["date"], r2["company"]),
                         ("John Carter", "john.carter@example.com", "090-3333-4444",
                          "2024-01-06", "Acme Labs LLC"))
        self.assertEqual(r2["issues"], "missing_address")
        _, r3 = rows[4]
        self.assertEqual((r3["name"], r3["phone"], r3["address"], r3["date"], r3["company"]),
                         ("山田 太郎", "090-5555-6666", "東京都千代田区架空町1-2-3",
                          "2024-01-07", "株式会社ABC商事"))
        # summary arithmetic: Clean + Duplicates + Invalid + Needs review == Input rows
        self.assertIn("Input rows (入力件数): 3", self.summary_text(out))

    def test_02_empty_csv_zero_byte_and_header_only(self):
        for name in ("empty_zero.csv", "header_only.csv"):
            with self.subTest(name):
                code, out = self.run_cleaner(FIXTURES / name, out=self.tmp / name)
                self.assertEqual(code, 0)
                self.assertIn("Input rows (入力件数): 0", self.summary_text(out))
                self.assertEqual(self.sheet(out, "cleaned_data.csv"), [])
                self.assertIn("WARNING", (out / "process.log").read_text(encoding="utf-8"))

    def test_03_missing_required_column_exit_code(self):
        code, _ = self.run_cleaner(FIXTURES / "missing_required.csv")
        self.assertEqual(code, 2)
        # one good file + one bad file: the good one is processed, the bad one is reported
        code, out = self.run_cleaner(FIXTURES / "missing_required.csv", FIXTURES / "normal.csv",
                                     out=self.tmp / "mixed")
        self.assertEqual(code, 0)
        self.assertIn("Skipped files (スキップしたファイル): 1  [missing_required.csv]",
                      self.summary_text(out))
        self.assertEqual(run([str(self.tmp / "nothing_here")], str(self.tmp / "o")), 1)

    def test_04_exact_and_email_duplicates(self):
        code, out = self.run_cleaner(FIXTURES / "duplicates.csv")
        rows = self.by_row(out)
        self.assertEqual(rows[2][0], "clean")
        self.assertEqual((rows[3][0], rows[3][1]["reason"], rows[3][1]["duplicate_of"]),
                         ("duplicate", "duplicate_exact", "duplicates.csv:2"))
        self.assertEqual((rows[4][0], rows[4][1]["reason"]), ("duplicate", "duplicate_email"))
        # same name + company but another e-mail: kept for review, never deleted
        self.assertEqual((rows[5][0], rows[5][1]["reason"], rows[5][1]["duplicate_of"]),
                         ("needs_review", "possible_duplicate", "duplicates.csv:2"))

    def test_05_invalid_emails(self):
        code, out = self.run_cleaner(FIXTURES / "invalid_email.csv")
        rows = self.by_row(out)
        self.assertEqual((rows[2][0], rows[2][1]["reason"]), ("invalid", "invalid_email"))
        self.assertEqual((rows[3][0], rows[3][1]["reason"]), ("invalid", "invalid_email"))  # -example.com
        self.assertEqual((rows[4][0], rows[4][1]["reason"], rows[4][1]["suggestion"]),
                         ("needs_review", "email_domain_typo", "eve.grant@gmail.com"))
        self.assertEqual(rows[4][1]["email"], "eve.grant@gmial.com")  # value is not "fixed"
        self.assertEqual(rows[5][0], "clean")

    def test_07_missing_values(self):
        code, out = self.run_cleaner(FIXTURES / "missing_values.csv")
        rows = self.by_row(out)
        self.assertEqual(rows[2][0], "clean")
        self.assertEqual(rows[2][1]["issues"],
                         "missing_phone;missing_address;missing_date;missing_company")
        # no e-mail but a phone number -> still clean
        self.assertEqual((rows[3][0], rows[3][1]["issues"]), ("clean", "missing_email"))
        self.assertEqual((rows[4][0], rows[4][1]["reason"]), ("invalid", "no_contact"))
        self.assertEqual((rows[5][0], rows[5][1]["reason"]), ("invalid", "missing_name"))

    def test_08_ambiguous_and_iso_dates(self):
        code, out = self.run_cleaner(FIXTURES / "dates.csv")
        rows = self.by_row(out)
        status, rec = rows[2]
        self.assertEqual((status, rec["reason"], rec["date"], rec["suggestion"]),
                         ("needs_review", "ambiguous_date", "03/04/2024",
                          "2024-03-04 or 2024-04-03"))
        self.assertEqual(rows[3][1]["date"], "2024-03-25")           # 25/03/2024: day > 12
        for n in (4, 5, 6, 7):                                       # ISO 8601 / CRM styles
            self.assertEqual((rows[n][0], rows[n][1]["date"]), ("clean", "2024-01-05"), n)
        self.assertEqual((rows[8][0], rows[8][1]["reason"], rows[8][1]["date"]),
                         ("needs_review", "invalid_date", "2024-02-30"))
        for order, expected in (("mdy", "2024-03-04"), ("dmy", "2024-04-03")):
            _, out2 = self.run_cleaner(FIXTURES / "dates.csv", out=self.tmp / order, date_order=order)
            status, rec = self.by_row(out2)[2]
            self.assertEqual((status, rec["date"]), ("clean", expected), order)

    def test_09_phone_review_rules_in_pipeline(self):
        code, out = self.run_cleaner(FIXTURES / "phones.csv")
        rows = self.by_row(out)
        s, r = rows[2]   # leading 0 lost (10 digits): JP or US, so ask
        self.assertEqual((s, r["reason"], r["phone"]), ("needs_review", "phone_leading_zero", "9012345678"))
        self.assertEqual(r["suggestion"], "090-1234-5678 (JP, leading 0 lost?) or +1 901-234-5678 (US)")
        s, r = rows[3]   # leading 0 lost (9 digits)
        self.assertEqual((s, r["reason"], r["suggestion"]), ("needs_review", "phone_leading_zero", "03-1234-5678"))
        s, r = rows[4]   # extension: value untouched
        self.assertEqual((s, r["reason"], r["phone"]),
                         ("needs_review", "phone_needs_check", "090-4444-5555 ext. 12"))
        self.assertEqual((rows[5][0], rows[5][1]["phone"]), ("clean", "+442079460958"))   # (0) dropped
        s, r = rows[6]   # landline without separators: digits kept, soft issue only
        self.assertEqual((s, r["phone"], "phone_format_unverified" in r["issues"]),
                         ("clean", "0467123456", True))
        self.assertEqual((rows[7][0], rows[7][1]["phone"]), ("clean", "0467-12-3456"))    # +81 keeps separators
        self.assertEqual((rows[8][0], rows[8][1]["reason"]), ("invalid", "invalid_phone"))

    def test_10_excel_input(self):
        code, out = self.run_cleaner(FIXTURES / "excel_input.xlsx")
        self.assertEqual(code, 0)
        rows = self.by_row(out)
        s, r = rows[2]   # numeric cell 9012345678 (Excel dropped the 0) + a real date cell
        self.assertEqual((s, r["reason"], r["date"]), ("needs_review", "phone_leading_zero", "2024-03-01"))
        s, r = rows[3]   # Excel serial 45352
        self.assertEqual((s, r["date"], r["company"]), ("clean", "2024-03-01", "Orion Labs Inc."))
        self.assertEqual({r["source_file"] for _, r in rows.values()}, {"excel_input.xlsx"})

    def test_11_shift_jis_and_japanese(self):
        code, out = self.run_cleaner(FIXTURES / "sjis.csv")
        self.assertEqual(code, 0)
        self.assertIn("cp932", (out / "process.log").read_text(encoding="utf-8"))
        rows = self.by_row(out)
        _, a = rows[2]
        self.assertEqual((a["name"], a["phone"], a["address"], a["date"], a["company"]),
                         ("佐藤 花子", "090-1234-5678", "大阪府大阪市北区見本町1-2-3", "2024-01-05",
                          "ABC商事株式会社"))
        _, b = rows[3]
        self.assertEqual((b["name"], b["phone"], b["date"], b["company"]),
                         ("鈴木 一郎", "03-1234-5678", "2024-01-06", "サクラ食品株式会社"))

    def test_12_performance_1300_rows_under_5_seconds(self):
        lines = ["name,email,phone,address,date,company"]
        for i in range(1300):
            n = i - 40 if i % 10 == 0 and i >= 40 else i          # every 10th row repeats an earlier one
            lines.append(f"Test User {n},user{n}@example.com,090-{1000 + n:04d}-{n:04d},"
                         f"{n} Sample St,2024/1/{n % 28 + 1},Load Test Inc")
        path = self.tmp / "load.csv"
        path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        started = time.perf_counter()
        code, out = self.run_cleaner(path)
        elapsed = time.perf_counter() - started
        self.assertEqual(code, 0)
        self.assertLess(elapsed, 5.0)
        self.assertIn("Input rows (入力件数): 1300", self.summary_text(out))

    def test_13_same_input_same_output_and_inputs_untouched(self):
        inputs = [FIXTURES / n for n in ("normal.csv", "duplicates.csv", "excel_input.xlsx", "sjis.csv")]
        before = [hashlib.sha256(p.read_bytes()).hexdigest() for p in inputs]
        _, out1 = self.run_cleaner(*inputs, out=self.tmp / "r1")
        _, out2 = self.run_cleaner(*inputs, out=self.tmp / "r2")
        for name in ("cleaned_data.csv", "cleaned_data.xlsx", "errors.xlsx", "needs_review.xlsx",
                     "change_log.xlsx"):
            self.assertEqual((out1 / name).read_bytes(), (out2 / name).read_bytes(), name)
        self.assertEqual(before, [hashlib.sha256(p.read_bytes()).hexdigest() for p in inputs])

    def test_14_refuses_to_overwrite_an_input_file(self):
        out = self.tmp / "out"
        out.mkdir()
        target = out / "cleaned_data.csv"
        target.write_text("name,email\nAlice Baker,alice.baker@example.com\n", encoding="utf-8")
        before = target.read_bytes()
        code, _ = self.run_cleaner(target, out=out)
        self.assertEqual(code, 1)
        self.assertEqual(target.read_bytes(), before)


class TestNormalize(unittest.TestCase):
    def phone(self, value):
        return N.normalize_phone(value)

    def test_06_phone_table_from_spec(self):
        table = [
            ("09012345678", "090-1234-5678"), ("０９０－１２３４－５６７８", "090-1234-5678"),
            ("+81 90-1234-5678", "090-1234-5678"), ("0312345678", "03-1234-5678"),
            ("0451234567", None), ("0467-12-3456", "0467-12-3456"),   # None: see next test
            ("0120123456", "0120-123-456"), ("(415) 555-0132", "+1 415-555-0132"),
            ("1-415-555-0132", "+1 415-555-0132"), ("+44 20 7946 0958", "+442079460958"),
            ("090-1234-567", "INVALID"),
        ]
        for raw, expected in table:
            with self.subTest(raw):
                value, problem, _ = self.phone(raw)
                if expected == "INVALID":
                    self.assertEqual(problem, "invalid_phone")
                elif expected is None:
                    self.assertIn(problem, ("", "phone_format_unverified"))
                else:
                    self.assertEqual((value, problem), (expected, ""))

    def test_06b_phone_area_code_is_never_guessed(self):
        # 0451234567 has no separators: 045- or 0451-? -> digits stay, soft issue
        self.assertEqual(self.phone("0451234567"), ("0451234567", "phone_format_unverified", ""))
        self.assertEqual(self.phone("0467123456"), ("0467123456", "phone_format_unverified", ""))
        # separators given by the input are kept, also behind +81
        self.assertEqual(self.phone("+81 467-12-3456")[:2], ("0467-12-3456", ""))
        self.assertEqual(self.phone("+81-45-123-4567")[:2], ("045-123-4567", ""))
        self.assertEqual(self.phone("(0467)12-3456")[:2], ("0467-12-3456", ""))
        self.assertEqual(self.phone("+81 3-1234-5678")[:2], ("03-1234-5678", ""))

    def test_06c_phone_leading_zero_lost(self):
        self.assertEqual(self.phone("9012345678"),
                         ("9012345678", "phone_leading_zero",
                          "090-1234-5678 (JP, leading 0 lost?) or +1 901-234-5678 (US)"))
        self.assertEqual(self.phone("312345678"), ("312345678", "phone_leading_zero", "03-1234-5678"))
        self.assertEqual(self.phone("467123456"), ("467123456", "phone_leading_zero", "0467123456"))
        self.assertEqual(self.phone("4155550132")[:2], ("+1 415-555-0132", ""))  # other bare 10 digits: US
        self.assertEqual(self.phone("415-555-0132")[:2], ("+1 415-555-0132", ""))

    def test_06e_leading_zero_lost_with_japanese_separators(self):
        # US numbers are written 3-3-4, so 2-4-4 / 1-4-4 groups are a Japanese number minus its 0
        self.assertEqual(self.phone("90-1234-5678"),
                         ("90-1234-5678", "phone_leading_zero",
                          "090-1234-5678 (JP, leading 0 lost?) or +1 901-234-5678 (US)"))
        self.assertEqual(self.phone("3-1234-5678"), ("3-1234-5678", "phone_leading_zero", "03-1234-5678"))
        self.assertEqual(self.phone("90 1234 5678")[1], "phone_leading_zero")
        # unaffected: a leading 0 is present, 3-3-4 groups, or a 9-digit value that starts with 0
        self.assertEqual(self.phone("090-1234-5678")[:2], ("090-1234-5678", ""))
        self.assertEqual(self.phone("03-1234-5678")[:2], ("03-1234-5678", ""))
        self.assertEqual(self.phone("415-555-0132")[:2], ("+1 415-555-0132", ""))
        self.assertEqual(self.phone("012345678")[1], "invalid_phone")

    def test_06d_phone_uk_trunk_zero_extension_and_multiple(self):
        self.assertEqual(self.phone("+44 (0)20 7946 0958")[:2], ("+442079460958", ""))
        for raw in ("090-1234-5678 ext. 12", "03-1234-5678 内線 12", "090-1234-5678 x12",
                    "090-1234-5678 #12", "090-1234-5678/03-1234-5678",
                    "090-1234-5678、03-1234-5678", "090-1234-5678, 03-1234-5678"):
            with self.subTest(raw):
                self.assertEqual(self.phone(raw), (raw, "phone_needs_check", ""))
        self.assertEqual(self.phone("415,555,0132")[:2], ("+1 415-555-0132", ""))  # not a list of numbers
        for raw in ("12345", "abc-defg", "0000000000", "+81 123", "+1 415-555-013"):
            with self.subTest(raw):
                self.assertEqual(self.phone(raw)[1], "invalid_phone")

    def test_dates(self):
        for raw in ("2024-01-05T10:30:00Z", "2024-01-05T10:30:00.000+09:00", "5-Jan-2024",
                    "2024年1月5日(金)", "2024年1月5日（金曜日）", "2024/1/5 (Fri)", "2024.1.5",
                    "Jan 5, 2024", "5 January 2024", "20240105", "45296"):
            with self.subTest(raw):
                self.assertEqual(N.normalize_date(raw), ("2024-01-05", "", ""))
        self.assertEqual(N.normalize_date("2023/13/01")[1], "invalid_date")

    def test_company_and_email_details(self):
        self.assertEqual(N.normalize_company("(株) ABC商事"), "株式会社ABC商事")
        self.assertEqual(N.normalize_company("ABC商事 (株)"), "ABC商事株式会社")
        self.assertEqual(N.normalize_company("sakura foods co., ltd."), "Sakura Foods Co., Ltd.")
        self.assertFalse(N.is_valid_email("taro@-example.com"))
        self.assertFalse(N.is_valid_email("taro@example-.com"))
        self.assertTrue(N.is_valid_email("taro@ex-ample.com"))


if __name__ == "__main__":
    unittest.main()
