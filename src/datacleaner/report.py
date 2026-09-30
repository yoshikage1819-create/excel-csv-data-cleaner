"""Output files and processing summary (SPEC sections 9 and 10)."""
import io
import re
import unicodedata
import zipfile
from collections import Counter
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from .validate import COLUMNS

CLEANED_COLUMNS = [*COLUMNS, "issues", "source_file", "source_row"]
INVALID_COLUMNS = ["source_file", "source_row", "reason", *COLUMNS]
DUPLICATE_COLUMNS = ["source_file", "source_row", "reason", "duplicate_of", *COLUMNS]
REVIEW_COLUMNS = ["source_file", "source_row", "reason", "suggestion", "duplicate_of", *COLUMNS]
CHANGE_COLUMNS = ["source_file", "source_row", "column", "before", "after"]

OUTPUT_FILES = ("cleaned_data.xlsx", "cleaned_data.csv", "errors.xlsx", "needs_review.xlsx",
                "change_log.xlsx", "summary.xlsx", "summary.txt", "process.log")


# ------------------------------------------------------------------ xlsx
def _display_width(text):
    return sum(2 if unicodedata.east_asian_width(ch) in "WF" else 1 for ch in text)


def save_workbook(wb, path):
    """Save with fixed timestamps so identical content gives identical bytes."""
    wb.properties.created = wb.properties.modified = datetime(2000, 1, 1)
    buf = io.BytesIO()
    wb.save(buf)
    with zipfile.ZipFile(buf) as zin, \
            zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zi = zipfile.ZipInfo(info.filename, date_time=(2000, 1, 1, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            data = zin.read(info.filename)
            if info.filename == "docProps/core.xml":  # openpyxl stamps "modified" at save time
                data = re.sub(rb"(<dcterms:modified[^>]*>)[^<]*", rb"\g<1>2000-01-01T00:00:00Z", data)
            zout.writestr(zi, data)


def write_xlsx(path, sheets):
    """sheets: {sheet_name: (header, rows)}. Bold header, frozen top row,
    autofilter, column widths fitted to content (max 50)."""
    wb = Workbook()
    wb.remove(wb.active)
    for name, (header, rows) in sheets.items():
        ws = wb.create_sheet(name)
        for c, title in enumerate(header, start=1):
            cell = ws.cell(1, c, title)
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="D9E1F2")
        widths = [_display_width(h) for h in header]
        for r, row in enumerate(rows, start=2):
            for c, v in enumerate(row, start=1):
                cell = ws.cell(r, c, v)
                if isinstance(v, str):
                    if v.startswith("="):  # keep formula-like text as text
                        cell.data_type = "s"
                    widths[c - 1] = max(widths[c - 1], _display_width(v))
        for c, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(c)].width = min(50, w + 2)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    save_workbook(wb, path)


# --------------------------------------------------------------- summary
def _breakdown(rows, status, codes):
    counts = Counter(code for r in rows if r["status"] == status for code in r["reasons"])
    return " / ".join(f"{label} {counts[code]}" for label, code in codes)


def build_summary(rows, n_files, skipped, processed_at, elapsed):
    """List of (label, value) pairs shared by summary.txt and summary.xlsx."""
    by_status = Counter(r["status"] for r in rows)
    missing = Counter(code for r in rows if r["status"] == "clean" for code in r["issues"])
    dup = _breakdown(rows, "duplicate", [("exact", "duplicate_exact"), ("same email", "duplicate_email")])
    invalid = _breakdown(rows, "invalid", [
        ("invalid email", "invalid_email"), ("invalid phone", "invalid_phone"),
        ("missing name", "missing_name"), ("no contact", "no_contact")])
    review = _breakdown(rows, "needs_review", [
        ("ambiguous date", "ambiguous_date"), ("possible duplicate", "possible_duplicate"),
        ("email typo", "email_domain_typo"), ("invalid date", "invalid_date"),
        ("phone leading zero", "phone_leading_zero"), ("phone needs check", "phone_needs_check")])
    missing_text = " / ".join(
        f"{c} {missing['missing_' + c]}" for c in ("phone", "address", "date", "company", "email"))
    return [
        ("Processed at (処理日時)", processed_at.strftime("%Y-%m-%d %H:%M:%S")),
        ("Input files (入力ファイル数)", str(n_files)),
        ("Skipped files (スキップしたファイル)",
         f"{len(skipped)}" + (f"  [{', '.join(skipped)}]" if skipped else "")),
        ("Input rows (入力件数)", str(len(rows))),
        ("Clean rows (出力件数)", str(by_status["clean"])),
        ("Duplicates removed (削除した重複)", f"{by_status['duplicate']}  [{dup}]"),
        ("Invalid rows (不正データ)", f"{by_status['invalid']}  [{invalid}]"),
        ("Needs review (要確認)", f"{by_status['needs_review']}  [{review}]"),
        ("Missing values in clean rows (欠損)", missing_text),
        ("Clean rows with unverified phone format (電話の書式未確認)", str(missing["phone_format_unverified"])),
        ("Cells auto-corrected (自動修正したセル)", str(sum(len(r["changes"]) for r in rows))),
        ("Processing time (処理時間)", f"{elapsed:.1f} s"),
    ]


def summary_text(items):
    lines = ["Excel / CSV Data Cleaner - Processing Summary (処理サマリー)"]
    lines += [f"{label}: {value}" for label, value in items]
    return "\n".join(lines) + "\n"


# ---------------------------------------------------------------- writer
def _records(rows, columns):
    return [[r["reason_text"] if c == "reason" else r["suggestion_text"] if c == "suggestion"
             else r["issues_text"] if c == "issues" else r[c] for c in columns] for r in rows]


def write_reports(rows, summary_items, out_dir):
    out = Path(out_dir)
    for r in rows:
        r["reason_text"] = ";".join(r["reasons"])
        r["suggestion_text"] = "; ".join(r["suggestions"])
        r["issues_text"] = ";".join(r["issues"])
    clean = [r for r in rows if r["status"] == "clean"]
    invalid = [r for r in rows if r["status"] == "invalid"]
    duplicates = [r for r in rows if r["status"] == "duplicate"]
    review = [r for r in rows if r["status"] == "needs_review"]

    cleaned = _records(clean, CLEANED_COLUMNS)
    write_xlsx(out / "cleaned_data.xlsx", {"cleaned": (CLEANED_COLUMNS, cleaned)})
    pd.DataFrame(cleaned, columns=CLEANED_COLUMNS).to_csv(
        out / "cleaned_data.csv", index=False, encoding="utf-8-sig", lineterminator="\r\n")
    write_xlsx(out / "errors.xlsx", {
        "invalid": (INVALID_COLUMNS, _records(invalid, INVALID_COLUMNS)),
        "duplicates": (DUPLICATE_COLUMNS, _records(duplicates, DUPLICATE_COLUMNS)),
    })
    write_xlsx(out / "needs_review.xlsx", {
        "needs_review": (REVIEW_COLUMNS, _records(review, REVIEW_COLUMNS))})
    changes = [[r["source_file"], r["source_row"], c, before, after]
               for r in rows for c, before, after in r["changes"]]
    write_xlsx(out / "change_log.xlsx", {"changes": (CHANGE_COLUMNS, changes)})
    write_xlsx(out / "summary.xlsx", {
        "summary": (["Item (項目)", "Value (値)"], [list(item) for item in summary_items])})
    (out / "summary.txt").write_text(summary_text(summary_items), encoding="utf-8")
    return {"clean": len(clean), "invalid": len(invalid), "duplicate": len(duplicates),
            "needs_review": len(review)}
