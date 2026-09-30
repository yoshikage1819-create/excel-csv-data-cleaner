"""Row validation and classification (SPEC section 7)."""
from . import normalize as N

COLUMNS = ("name", "email", "phone", "address", "date", "company")


def validate_row(row, date_order=None):
    """Normalize the six columns of `row` (from row["raw"]) and evaluate the
    per-row rules. Fills: the six columns, changes, hard, review, issues,
    suggestions."""
    raw = row["raw"]
    phone, phone_problem, phone_suggestion = N.normalize_phone(raw["phone"])
    date_value, date_problem, date_suggestion = N.normalize_date(raw["date"], date_order)
    values = {
        "name": N.normalize_name(raw["name"]),
        "email": N.normalize_email(raw["email"]),
        "phone": phone,
        "address": N.normalize_address(raw["address"]),
        "date": date_value,
        "company": N.normalize_company(raw["company"]),
    }
    row.update(values)
    row["changes"] = [(c, raw[c], values[c]) for c in COLUMNS if raw[c] != values[c]]

    hard, review, issues, suggestions = [], [], [], []
    if not values["name"]:
        hard.append("missing_name")
    email = values["email"]
    if not email:
        issues.append("missing_email")
    elif not N.is_valid_email(email):
        hard.append("invalid_email")
    else:
        suggestion = N.email_typo_suggestion(email)
        if suggestion:
            review.append("email_domain_typo")
            suggestions.append(suggestion)
    if not phone:
        issues.append("missing_phone")
    elif phone_problem == "invalid_phone":
        hard.append("invalid_phone")
    elif phone_problem == "phone_format_unverified":
        issues.append(phone_problem)
    elif phone_problem:  # phone_leading_zero / phone_needs_check
        review.append(phone_problem)
        if phone_suggestion:
            suggestions.append(phone_suggestion)
    if not email and not phone:
        hard.append("no_contact")
    if not values["address"]:
        issues.append("missing_address")
    if not date_value:
        issues.append("missing_date")
    elif date_problem:
        review.append(date_problem)
        if date_suggestion:
            suggestions.append(date_suggestion)
    if not values["company"]:
        issues.append("missing_company")

    row.update(hard=hard, review=review, issues=issues, suggestions=suggestions)


def assign_status(rows):
    """Priority: invalid > duplicate > needs_review > clean. Call after dedupe."""
    for row in rows:
        row["reasons"] = row["hard"] + ([row["dup"]] if row["dup"] else []) + row["review"]
        if row["hard"]:
            row["status"] = "invalid"
        elif row["dup"]:
            row["status"] = "duplicate"
        elif row["review"]:
            row["status"] = "needs_review"
        else:
            row["status"] = "clean"
