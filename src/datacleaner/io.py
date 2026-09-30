"""Reading CSV / Excel files and mapping headers to the standard columns (SPEC section 5)."""
import io as _io
import json
import logging
import re
import unicodedata
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

log = logging.getLogger("datacleaner")

STANDARD_COLUMNS = ("name", "email", "phone", "address", "date", "company")
REQUIRED_COLUMNS = ("name", "email")
SUPPORTED_SUFFIXES = (".csv", ".xlsx")
ALIASES_PATH = Path(__file__).resolve().parents[2] / "config" / "column_aliases.json"


class SkipFile(Exception):
    """The file cannot be processed. missing_columns is True when the cause is
    a missing required column (exit code 2 if every file fails that way)."""

    def __init__(self, message, missing_columns=False):
        super().__init__(message)
        self.missing_columns = missing_columns


def normalize_header(text):
    text = unicodedata.normalize("NFKC", str(text)).lower()
    return re.sub(r"[\s_\-]+", "", text)


def load_alias_map():
    with open(ALIASES_PATH, encoding="utf-8") as f:
        aliases = json.load(f)
    return {normalize_header(a): col for col, names in aliases.items() for a in names}


def find_input_files(inputs):
    """Expand files/folders into a list ordered by file name."""
    found = []
    for item in inputs:
        p = Path(item)
        if p.is_dir():
            files = [f for f in p.iterdir()
                     if f.is_file() and f.suffix.lower() in SUPPORTED_SUFFIXES
                     and not f.name.startswith("~$")]
            if not files:
                log.warning("No .csv/.xlsx files in folder: %s", p)
            found.extend(files)
        elif p.is_file():
            if p.name.startswith("~$") or p.suffix.lower() not in SUPPORTED_SUFFIXES:
                log.warning("Skipping unsupported file: %s", p)
            else:
                found.append(p)
        else:
            log.error("Input not found: %s", p)
    unique = {str(f.resolve()).casefold(): f for f in found}
    return sorted(unique.values(), key=lambda f: (f.name.casefold(), f.name, str(f)))


def _read_csv_table(path):
    data = path.read_bytes()
    if not data:
        return [], "(empty file)"
    for encoding in ("utf-8-sig", "cp932"):
        try:
            text = data.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise SkipFile(f"{path.name}: cannot decode the file (tried utf-8-sig and cp932).")
    try:
        df = pd.read_csv(_io.StringIO(text), header=None, dtype=str,
                         keep_default_na=False, skip_blank_lines=False)
    except (pd.errors.ParserError, pd.errors.EmptyDataError) as e:
        raise SkipFile(f"{path.name}: cannot parse the CSV ({e}).") from e
    return df.values.tolist(), encoding


def _cell_to_text(v):
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _read_xlsx_table(path):
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        raise SkipFile(f"{path.name}: cannot open the Excel file ({e}).") from e
    try:
        rows = [[_cell_to_text(v) for v in row]
                for row in wb.worksheets[0].iter_rows(values_only=True)]
    finally:
        wb.close()
    return rows, f"xlsx (sheet: {wb.worksheets[0].title})"


def _is_blank(row):
    return not any(str(c).strip() for c in row)


def read_file(path, alias_map):
    """Return a list of {"source_file", "source_row", "raw": {column: str}}."""
    path = Path(path)
    table, encoding = (_read_xlsx_table(path) if path.suffix.lower() == ".xlsx"
                       else _read_csv_table(path))
    header_idx = next((i for i, r in enumerate(table) if not _is_blank(r)), None)
    if header_idx is None:
        log.warning("%s: empty file, treated as 0 rows.", path.name)
        return []

    mapping = {}
    for idx, header in enumerate(table[header_idx]):
        std = alias_map.get(normalize_header(header))
        if std and std not in mapping.values():
            mapping[idx] = std
        elif std:
            log.warning("%s: duplicate column '%s' for '%s' ignored.", path.name, header, std)
        elif str(header).strip():
            log.warning("%s: unknown column '%s' ignored.", path.name, header)
    mapped = set(mapping.values())
    missing = [c for c in REQUIRED_COLUMNS if c not in mapped]
    if missing:
        raise SkipFile(
            f"{path.name}: required column(s) missing: {', '.join(missing)}. "
            f"Columns found: {[str(h) for h in table[header_idx]]}. Skipping this file.",
            missing_columns=True)
    for col in STANDARD_COLUMNS:
        if col not in mapped:
            log.warning("%s: optional column '%s' not found, filled with empty values.",
                        path.name, col)

    rows, blank = [], 0
    for i in range(header_idx + 1, len(table)):
        cells = table[i]
        if _is_blank(cells):
            blank += 1
            continue
        raw = {c: "" for c in STANDARD_COLUMNS}
        for idx, col in mapping.items():
            raw[col] = str(cells[idx]) if idx < len(cells) else ""
        rows.append({"source_file": path.name, "source_row": i + 1, "raw": raw})

    log.info("%s: %d rows read (encoding: %s); columns: %s", path.name, len(rows), encoding,
             ", ".join(f"{table[header_idx][i]}->{c}" for i, c in mapping.items()))
    if blank:
        log.warning("%s: %d blank row(s) ignored.", path.name, blank)
    if not rows:
        log.warning("%s: no data rows (header only), treated as 0 rows.", path.name)
    return rows
