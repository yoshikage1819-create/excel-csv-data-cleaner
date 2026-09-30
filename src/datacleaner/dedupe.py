"""Duplicate detection (SPEC section 8)."""
from .validate import COLUMNS


def _ref(row):
    return f"{row['source_file']}:{row['source_row']}"


def mark_duplicates(rows):
    """`rows` must be in processing order (file name, then row number); the
    first row of a group is always kept. Sets row["dup"] and
    row["duplicate_of"], and appends possible_duplicate to row["review"].
    Invalid rows are ignored."""
    seen_exact, seen_email, seen_name_company, seen_phone = {}, {}, {}, {}
    for row in rows:
        row["dup"] = ""
        row["duplicate_of"] = ""
        if row["hard"]:
            continue

        first = seen_exact.get(tuple(row[c] for c in COLUMNS))
        if first:
            row["dup"], row["duplicate_of"] = "duplicate_exact", _ref(first)
            continue
        email = row["email"]
        first = seen_email.get(email) if email else None
        if first:
            row["dup"], row["duplicate_of"] = "duplicate_email", _ref(first)
            continue

        candidates = []
        if row["name"] and row["company"]:
            candidates.append(seen_name_company.get((row["name"], row["company"])))
        if row["phone"]:
            candidates.append(seen_phone.get(row["phone"]))
        for other in candidates:
            if other and other["email"] != email:
                row["review"].append("possible_duplicate")
                row["duplicate_of"] = _ref(other)
                break

        seen_exact[tuple(row[c] for c in COLUMNS)] = row
        if email:
            seen_email[email] = row
        if row["name"] and row["company"]:
            seen_name_company.setdefault((row["name"], row["company"]), row)
        if row["phone"]:
            seen_phone.setdefault(row["phone"], row)
