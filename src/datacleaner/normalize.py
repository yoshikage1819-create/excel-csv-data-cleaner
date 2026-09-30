"""Cell-level normalization rules (SPEC section 6). Pure functions, no I/O.

Principle: only convert what we are sure about. Anything that cannot be
converted is returned unchanged so that validate.py can flag it.
"""
import re
import unicodedata
from datetime import date, timedelta

EMPTY_TOKENS = {"", "nan", "none", "null", "n/a", "na", "-", "—"}

_CONTROL = re.compile(r"[\x00-\x08\x0e-\x1f\x7f]")
_SPACES = re.compile(r"\s+")


def clean_text(value):
    """6.1: NFKC, whitespace cleanup, empty-token handling."""
    s = unicodedata.normalize("NFKC", str(value))
    s = _CONTROL.sub("", s)
    s = _SPACES.sub(" ", s).strip()
    return "" if s.lower() in EMPTY_TOKENS else s


def _capitalize_words(s):
    return re.sub(r"[a-z]+", lambda m: m.group().capitalize(), s.lower())


# ---------------------------------------------------------------- name (6.2)
_ASCII_NAME = re.compile(r"[A-Za-z .'\-]+")


def normalize_name(value):
    s = clean_text(value)
    if (
        _ASCII_NAME.fullmatch(s)
        and re.search(r"[A-Za-z]", s)
        and (s == s.upper() or s == s.lower())
    ):
        return _capitalize_words(s)
    return s


# --------------------------------------------------------------- email (6.3)
_EMAIL = re.compile(r"[a-z0-9._%+-]+@[a-z0-9-]+(\.[a-z0-9-]+)*\.[a-z]{2,}")

EMAIL_DOMAIN_TYPOS = {
    "gmial.com": "gmail.com",
    "gmai.com": "gmail.com",
    "gamil.com": "gmail.com",
    "gmail.co": "gmail.com",
    "hotmial.com": "hotmail.com",
    "outlok.com": "outlook.com",
    "yahooo.co.jp": "yahoo.co.jp",
    "yaho.co.jp": "yahoo.co.jp",
    "yahooo.com": "yahoo.com",
}


def normalize_email(value):
    return clean_text(value).lower()


def is_valid_email(email):
    if not _EMAIL.fullmatch(email):
        return False
    local, _, domain = email.partition("@")
    if local.startswith(".") or local.endswith(".") or ".." in local:
        return False
    return not any(label.startswith("-") or label.endswith("-") for label in domain.split("."))


def email_typo_suggestion(email):
    """Corrected address if the domain is a known typo, else ''."""
    local, _, domain = email.rpartition("@")
    fixed = EMAIL_DOMAIN_TYPOS.get(domain)
    return f"{local}@{fixed}" if fixed else ""


# --------------------------------------------------------------- phone (6.4)
_JP_MOBILE = ("070", "080", "090", "050")
_EXTENSION = re.compile(r"(ext\.?|内線|#|(?<![a-z])x)\s*\d", re.I)
_TRUNK_ZERO = re.compile(r"^(\+\d{1,3})[\s.\-]*\(0\)")


def _jp_phone(d, groups):
    """d: digits in domestic form (leading 0). groups: separator-split groups of
    the original text, or None when they are unusable.
    Returns (value, problem) or None; problem is "" or "phone_format_unverified"."""
    n = len(d)
    if n < 2 or d[1] == "0":
        return None
    if d.startswith("0120") and n == 10:
        return f"{d[:4]}-{d[4:7]}-{d[7:]}", ""
    if d.startswith("0800") and n == 11:
        return f"{d[:4]}-{d[4:7]}-{d[7:]}", ""
    if d[:3] in _JP_MOBILE and n == 11:
        return f"{d[:3]}-{d[3:7]}-{d[7:]}", ""
    if d[:2] in ("03", "06") and n == 10:
        return f"{d[:2]}-{d[2:6]}-{d[6:]}", ""
    if n == 10 and d[:3] not in _JP_MOBILE and not d.startswith(("0120", "0800")):
        # The area-code length cannot be known from the digits alone: keep the
        # input's own separators, or leave the digits untouched.
        if groups and len(groups) == 3 and groups[0].startswith("0"):
            return "-".join(groups), ""
        return d, "phone_format_unverified"
    return None


def _jp_groups(groups):
    """Groups of a +81 / 81... number as domestic groups: ["81","467","12","3456"] -> ["0467","12","3456"]."""
    if len(groups) > 1 and groups[0] == "81":
        return ["0" + groups[1], *groups[2:]]
    return None


def _na_phone(d):
    return f"+1 {d[:3]}-{d[3:6]}-{d[6:]}"


def _looks_like_extension_or_multiple(s):
    if _EXTENSION.search(s):
        return True
    parts = [p for p in re.split(r"[/,、;]", s) if len(re.sub(r"\D", "", p)) >= 7]
    return len(parts) >= 2


def normalize_phone(value):
    """Return (normalized, problem, suggestion).

    problem: "" | "invalid_phone" | "phone_leading_zero" | "phone_needs_check" |
    "phone_format_unverified" (soft: the number is kept as bare digits).
    Unless the number could be converted, the cleaned original is returned.
    """
    s = clean_text(value)
    if not s:
        return "", "", ""
    if _looks_like_extension_or_multiple(s):
        return s, "phone_needs_check", ""
    s = _TRUNK_ZERO.sub(r"\1", s)
    digits = re.sub(r"\D", "", s)
    groups = [g for g in re.split(r"\D+", s) if g]

    # Excel may have dropped a leading 0: bare digits, or the separators of a Japanese number
    # (2-4-4 / 1-4-4; North American numbers are written 3-3-4).
    shape = [len(g) for g in groups]
    if (not s.startswith("+") and digits[:1] != "0"
            and (re.fullmatch(r"\d+", s) or shape in ([2, 4, 4], [1, 4, 4]))):
        if len(digits) == 10 and digits[:2] in ("50", "70", "80", "90") and shape != [1, 4, 4]:
            jp = _jp_phone("0" + digits, None)[0]
            return s, "phone_leading_zero", f"{jp} (JP, leading 0 lost?) or {_na_phone(digits)} (US)"
        if len(digits) == 9:
            return s, "phone_leading_zero", normalize_phone("0" + digits)[0]

    if s.startswith("+"):
        if digits.startswith("81"):
            out = _jp_phone("0" + digits[2:], _jp_groups(groups))
        elif digits.startswith("1"):
            out = (_na_phone(digits[1:]), "") if len(digits) == 11 else None
        elif 8 <= len(digits) <= 15:
            out = ("+" + digits, "")
        else:
            out = None
    elif digits.startswith("81") and len(digits) in (11, 12):
        out = _jp_phone("0" + digits[2:], _jp_groups(groups))
    elif digits.startswith("0"):
        out = _jp_phone(digits, groups)
    elif digits.startswith("1") and len(digits) == 11:
        out = (_na_phone(digits[1:]), "")
    elif len(digits) == 10:
        out = (_na_phone(digits), "")
    else:
        out = None
    return (out[0], out[1], "") if out else (s, "invalid_phone", "")


# ---------------------------------------------------------------- date (6.5)
_MONTHS = {}
for _i, _full in enumerate(
    ["january", "february", "march", "april", "may", "june", "july",
     "august", "september", "october", "november", "december"], start=1):
    _MONTHS[_full] = _i
    _MONTHS[_full[:3]] = _i
_MONTHS["sept"] = 9
_MON = "(" + "|".join(sorted(_MONTHS, key=len, reverse=True)) + r")\.?"

_YMD = re.compile(
    r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})"
    r"(?:[ T]\d{1,2}:\d{2}(?::\d{2})?(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?)?", re.I)
_WEEKDAY = re.compile(r"\s*\((?:[月火水木金土日]曜?日?|(?:mon|tue|wed|thu|fri|sat|sun)[a-z]*\.?)\)$", re.I)
_YMD_JP = re.compile(r"(\d{4})年(\d{1,2})月(\d{1,2})日")
_YMD_8 = re.compile(r"(\d{4})(\d{2})(\d{2})")
_MDY_TEXT = re.compile(_MON + r"\s+(\d{1,2}),?\s+(\d{4})", re.I)
_DMY_TEXT = re.compile(r"(\d{1,2})(?:\s+|-)" + _MON + r"(?:,?\s+|-)(\d{4})", re.I)
_SERIAL = re.compile(r"(\d{5})(?:\.\d+)?")
_A_B_YEAR = re.compile(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})")


def _iso(y, m, d):
    try:
        return date(y, m, d).isoformat()
    except ValueError:
        return None


def normalize_date(value, date_order=None):
    """Return (normalized, problem, suggestion).

    problem is "", "ambiguous_date" or "invalid_date"; in the latter two the
    returned value is the cleaned original text.
    """
    s = clean_text(value)
    if not s:
        return "", "", ""
    s = _WEEKDAY.sub("", s)

    def result(iso):
        return (iso, "", "") if iso else (s, "invalid_date", "")

    for pattern in (_YMD, _YMD_JP, _YMD_8):
        m = pattern.fullmatch(s)
        if m:
            return result(_iso(*(int(g) for g in m.groups())))
    m = _MDY_TEXT.fullmatch(s)
    if m:
        return result(_iso(int(m.group(3)), _MONTHS[m.group(1).lower()], int(m.group(2))))
    m = _DMY_TEXT.fullmatch(s)
    if m:
        return result(_iso(int(m.group(3)), _MONTHS[m.group(2).lower()], int(m.group(1))))
    m = _SERIAL.fullmatch(s)
    if m:
        n = int(m.group(1))
        if 20000 <= n <= 60000:
            return (date(1899, 12, 30) + timedelta(days=n)).isoformat(), "", ""
        return s, "invalid_date", ""
    m = _A_B_YEAR.fullmatch(s)
    if m:
        a, b, y = (int(g) for g in m.groups())
        if a > 12 and b <= 12:
            return result(_iso(y, b, a))
        if b > 12 and a <= 12:
            return result(_iso(y, a, b))
        if a > 12 and b > 12:
            return s, "invalid_date", ""
        if a == b:
            return result(_iso(y, a, b))
        if date_order == "mdy":
            return result(_iso(y, a, b))
        if date_order == "dmy":
            return result(_iso(y, b, a))
        return s, "ambiguous_date", f"{y:04d}-{a:02d}-{b:02d} or {y:04d}-{b:02d}-{a:02d}"
    return s, "invalid_date", ""


# ------------------------------------------------------------- address (6.6)
_ADDRESS_DASH = re.compile(
    "(?<=\\d)[\u2010\u2011\u2012\u2013\u2014\u2015\u2212\u30fc\uff70\uff0d](?=\\d)")


def normalize_address(value):
    return _ADDRESS_DASH.sub("-", clean_text(value))


# -------------------------------------------------------------- company (6.7)
_EN_FORM = re.compile(
    r"(?P<base>.+?)(?:\s*,\s*|\s+)(?P<form>co\.?\s*,?\s*ltd\.?|inc\.?|llc|corp\.?|ltd\.?)",
    re.I)
_EN_FORM_CANON = {"coltd": "Co., Ltd.", "inc": "Inc.", "llc": "LLC",
                  "corp": "Corp.", "ltd": "Ltd."}


def normalize_company(value):
    s = clean_text(value)
    for short, full in (("(株)", "株式会社"), ("(有)", "有限会社")):
        if s.startswith(short):
            s = full + s[len(short):].lstrip()
        if s.endswith(short):
            s = s[: -len(short)].rstrip() + full
    m = _EN_FORM.fullmatch(s)
    if m:
        base = m.group("base")
        if base.isascii() and re.search(r"[a-z]", base) and base == base.lower():
            base = _capitalize_words(base)
        form = _EN_FORM_CANON[re.sub(r"[\s.,]", "", m.group("form").lower())]
        s = f"{base} {form}"
    return s
