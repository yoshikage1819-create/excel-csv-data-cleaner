"""Compare output/ with input/ground_truth.json and write output/verification_report.md (SPEC 14-3).

Usage: python scripts/verify_against_ground_truth.py [--input input] [--output output]
Run main.py without --date-order first (the ground truth assumes ambiguous dates go to review).
"""
import argparse
import json
import sys
from pathlib import Path

import pandas as pd

FIELDS = ("name", "email", "phone", "address", "date", "company")
KIND_ORDER = ["duplicate_exact", "duplicate_email", "possible_duplicate", "invalid_email",
              "email_domain_typo", "ambiguous_date", "invalid_date", "invalid_phone",
              "missing_name", "no_contact", "phone_leading_zero", "soft_missing",
              "phone_format_unverified"]


def read_table(path, sheet=None):
    if str(path).endswith(".csv"):
        df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    else:
        df = pd.read_excel(path, sheet_name=sheet, dtype=str, keep_default_na=False)
    return df.to_dict("records")


def load_actual(out_dir):
    actual = {}

    def add(records, status):
        for r in records:
            key = (r["source_file"], int(float(r["source_row"])))
            reasons = [x for x in r.get("reason", "").split(";") if x]
            issues = [x for x in r.get("issues", "").split(";") if x]
            actual[key] = {"status": status, "reasons": reasons, "issues": issues,
                           "values": {c: r[c] for c in FIELDS},
                           "duplicate_of": r.get("duplicate_of", ""),
                           "suggestion": r.get("suggestion", "")}

    add(read_table(out_dir / "cleaned_data.csv"), "clean")
    add(read_table(out_dir / "errors.xlsx", "invalid"), "invalid")
    add(read_table(out_dir / "errors.xlsx", "duplicates"), "duplicate")
    add(read_table(out_dir / "needs_review.xlsx", "needs_review"), "needs_review")
    return actual


def detected(row, act):
    """The planted problem was found: expected status and reason (or soft issues) present."""
    if act is None or act["status"] != row["expected_status"]:
        return False
    if row["expected_status"] == "clean":
        return all(i in act["issues"] for i in row.get("expected_issues", []))
    return all(r in act["reasons"] for r in row["expected_reasons"])


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default="input")
    parser.add_argument("--output", default="output")
    args = parser.parse_args()
    in_dir, out_dir = Path(args.input), Path(args.output)
    truth = json.loads((in_dir / "ground_truth.json").read_text(encoding="utf-8"))
    actual = load_actual(out_dir)
    rows = truth["rows"]

    lines = ["# Verification report (照合レポート)", "",
             f"- ground truth: {len(rows)} rows, seed {truth['seed']}",
             f"- output rows found: {len(actual)}",
             f"- rows missing from the output: "
             f"{sum(1 for r in rows if (r['source_file'], r['source_row']) not in actual)}", ""]

    # 1. detection per planted kind
    lines += ["## 1. Planted errors: planted / detected as expected / rate (仕込んだ数 / 検出数 / 検出率)", "",
              "| kind | planted | detected | rate |", "| --- | ---: | ---: | ---: |"]
    rates = {}
    for kind in KIND_ORDER:
        subset = [r for r in rows if r["planted"] == kind]
        hit = sum(1 for r in subset if detected(r, actual.get((r["source_file"], r["source_row"]))))
        rates[kind] = hit / len(subset) if subset else 1.0
        lines.append(f"| {kind} | {len(subset)} | {hit} | {rates[kind]:.1%} |")
    dup_of = [r for r in rows if "duplicate_of" in r]
    dup_ok = sum(1 for r in dup_of
                 if actual.get((r["source_file"], r["source_row"]), {}).get("duplicate_of")
                 == r["duplicate_of"])
    sugg = [r for r in rows if "suggestion" in r]
    sugg_ok = sum(1 for r in sugg
                  if actual.get((r["source_file"], r["source_row"]), {}).get("suggestion")
                  == r["suggestion"])
    lines += ["", f"- `duplicate_of` correct: {dup_ok} / {len(dup_of)}",
              f"- `suggestion` correct (typo / ambiguous date): {sugg_ok} / {len(sugg)}", ""]

    # 2. rows with a different category than expected
    wrong = [r for r in rows
             if not detected(r, actual.get((r["source_file"], r["source_row"])))]
    lines += [f"## 2. Rows with unexpected category (期待と違う区分): {len(wrong)} rows"
              + (" (first 30 shown)" if len(wrong) > 30 else ""), ""]
    for r in wrong[:30]:
        act = actual.get((r["source_file"], r["source_row"]))
        got = f"{act['status']} {act['reasons']}" if act else "(missing)"
        lines.append(f"- {r['source_file']}:{r['source_row']} expected {r['expected_status']} "
                     f"{r['expected_reasons'] or r.get('expected_issues', [])} / got {got}")
    if not wrong:
        lines.append("- none")
    lines.append("")

    # 3. normalization accuracy on clean rows
    clean_rows = [r for r in rows if r["expected_status"] == "clean"]
    in_clean, exact, bad_cells = 0, 0, []
    for r in clean_rows:
        act = actual.get((r["source_file"], r["source_row"]))
        if not act or act["status"] != "clean":
            continue
        in_clean += 1
        diffs = [(c, act["values"][c], r["canonical"][c]) for c in FIELDS
                 if act["values"][c] != r["canonical"][c]]
        if diffs:
            bad_cells.append((r["source_file"], r["source_row"], diffs))
        else:
            exact += 1
    accuracy = exact / in_clean if in_clean else 0.0
    lines += ["## 3. Normalization accuracy on clean rows (正規化精度)", "",
              f"- expected-clean rows that were output as clean: {in_clean} / {len(clean_rows)}",
              f"- all 6 columns equal to canonical: {exact} / {in_clean} = {accuracy:.2%}", ""]
    for f, n, diffs in bad_cells[:10]:
        lines.append(f"- {f}:{n} " + "; ".join(f"{c}: got {g!r} want {w!r}" for c, g, w in diffs))
    lines.append("")

    # 4. false positives
    false_pos = [r for r in clean_rows
                 if (a := actual.get((r["source_file"], r["source_row"]))) is None
                 or a["status"] != "clean"]
    lines += [f"## 4. False positives: expected-clean rows flagged invalid / needs_review / duplicate "
              f"(誤検出): {len(false_pos)}", ""]
    for r in false_pos[:10]:
        a = actual.get((r["source_file"], r["source_row"]))
        lines.append(f"- {r['source_file']}:{r['source_row']} -> {a['status'] if a else 'missing'} "
                     f"{a['reasons'] if a else ''}")
    lines.append("")

    # 5. targets
    dup_rate = min(rates["duplicate_exact"], rates["duplicate_email"])
    targets = [("duplicate detection >= 95%", dup_rate >= .95),
               ("invalid email detection >= 95%", rates["invalid_email"] >= .95),
               ("false positives ~ 0", len(false_pos) == 0),
               ("normalization accuracy >= 98%", accuracy >= .98),
               ("every planted kind detected as expected (no unexpected category)", not wrong)]
    lines += ["## 5. Targets (目標)", ""] + [f"- {'PASS' if ok else 'FAIL'}: {t}" for t, ok in targets]
    report = "\n".join(lines) + "\n"
    (out_dir / "verification_report.md").write_text(report, encoding="utf-8")
    sys.stdout.reconfigure(errors="replace")
    print(report)
    return 0 if all(ok for _, ok in targets) else 1


if __name__ == "__main__":
    sys.exit(main())
