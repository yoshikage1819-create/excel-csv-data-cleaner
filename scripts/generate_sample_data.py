"""Generate fictional messy sample data (SPEC section 4).

Deterministic (random.Random(42)). All people, companies, addresses and
e-mail addresses are fictional; e-mail domains are example.com/.net/.org.

Output (in ./input): messy_customers.csv (cp932, 1000 rows),
crm_export_2.xlsx (300 rows), ground_truth.json.

Processing order is file name order, so crm_export_2.xlsx is processed before
messy_customers.csv. Every duplicate row therefore comes *after* the row it
duplicates in that order (cross-file duplicates: original in the xlsx,
duplicate in the CSV).
"""
import csv
import json
import random
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
from datacleaner.report import save_workbook  # noqa: E402  (deterministic xlsx bytes)
from openpyxl import Workbook  # noqa: E402

CSV_NAME, XLSX_NAME = "messy_customers.csv", "crm_export_2.xlsx"
CSV_ROWS, XLSX_ROWS = 1000, 300
CSV_HEADER = ["氏名", "メール", "電話番号", "住所", "登録日", "会社名"]
XLSX_HEADER = ["Full Name", "E-mail", "Tel", "Address", "Created", "Company"]
FIELDS = ("name", "email", "phone", "address", "date", "company")

# planted error counts (SPEC 4.2)
N_EXACT, N_EMAIL, N_POSSIBLE = 40, 30, 15
N_INVALID_EMAIL, N_TYPO, N_AMBIGUOUS, N_INVALID_DATE = 50, 10, 20, 5
N_INVALID_PHONE, N_MISSING_NAME, N_NO_CONTACT, N_SOFT_MISSING = 15, 10, 10, 100
N_UNVERIFIED_PHONE = 5  # clean rows whose 10-digit landline is kept as bare digits (R1-2)
# phone with the leading 0 lost, per file: 4 mobile (10 digits) + 1 landline (9 digits) (R1-1)
LEADING_ZERO_PER_FILE = 5

# ------------------------------------------------------------------ pools
JP_SURNAMES = [
    ("佐藤", "sato"), ("鈴木", "suzuki"), ("高橋", "takahashi"), ("田中", "tanaka"),
    ("伊藤", "ito"), ("渡辺", "watanabe"), ("山本", "yamamoto"), ("中村", "nakamura"),
    ("小林", "kobayashi"), ("加藤", "kato"), ("吉田", "yoshida"), ("山田", "yamada"),
    ("佐々木", "sasaki"), ("山口", "yamaguchi"), ("松本", "matsumoto"), ("井上", "inoue"),
    ("木村", "kimura"), ("林", "hayashi"), ("清水", "shimizu"), ("森", "mori"),
    ("池田", "ikeda"), ("橋本", "hashimoto"), ("石川", "ishikawa"), ("前田", "maeda"),
    ("藤田", "fujita"), ("岡田", "okada"), ("後藤", "goto"), ("長谷川", "hasegawa"),
    ("村上", "murakami"), ("近藤", "kondo"), ("石井", "ishii"), ("斎藤", "saito"),
    ("坂本", "sakamoto"), ("遠藤", "endo"), ("青木", "aoki"), ("藤井", "fujii"),
    ("西村", "nishimura"), ("福田", "fukuda"), ("太田", "ota"), ("三浦", "miura"),
]
JP_GIVEN = [
    ("太郎", "taro"), ("花子", "hanako"), ("一郎", "ichiro"), ("次郎", "jiro"),
    ("健太", "kenta"), ("翔太", "shota"), ("大輔", "daisuke"), ("拓也", "takuya"),
    ("直樹", "naoki"), ("誠", "makoto"), ("学", "manabu"), ("純", "jun"),
    ("徹", "toru"), ("優子", "yuko"), ("美咲", "misaki"), ("陽子", "yoko"),
    ("恵子", "keiko"), ("由美", "yumi"), ("麻衣", "mai"), ("愛", "ai"),
    ("彩", "aya"), ("裕子", "hiroko"), ("真理", "mari"), ("智子", "tomoko"),
    ("香織", "kaori"), ("千尋", "chihiro"), ("沙織", "saori"), ("悠", "yu"),
    ("蓮", "ren"), ("陸", "riku"), ("結衣", "yui"), ("茜", "akane"),
    ("葵", "aoi"), ("隼人", "hayato"), ("和也", "kazuya"), ("雅人", "masato"),
    ("剛", "tsuyoshi"), ("亮", "ryo"), ("萌", "moe"), ("颯太", "sota"),
]
EN_FIRST = [
    "James", "Mary", "Robert", "Linda", "Michael", "Susan", "William", "Karen", "David",
    "Nancy", "Richard", "Lisa", "Joseph", "Betty", "Thomas", "Helen", "Charles", "Sandra",
    "Daniel", "Donna", "Matthew", "Carol", "Anthony", "Ruth", "Mark", "Sharon", "Steven",
    "Laura", "Paul", "Emily", "Andrew", "Anna", "Kevin", "Rachel", "Brian", "Diane",
    "George", "Edward", "Julia", "Peter",
]
EN_LAST = [
    "Baker", "Carter", "Dawson", "Ellis", "Foster", "Grant", "Hayes", "Irwin", "Jensen",
    "Keller", "Lawson", "Nolan", "Owens", "Parker", "Reed", "Sutton", "Turner", "Underwood",
    "Vaughn", "Walsh", "Yates", "Abbott", "Barnes", "Chambers", "Dixon", "Eaton", "Garrett",
    "Holland", "Ingram", "Joyce", "Kemp", "Lowe", "Marsh", "Osborne", "Pratt", "Rowe",
    "Sanders", "Tucker", "Wells", "Young",
]
DOMAINS = ["example.com", "example.net", "example.org"]

EN_BRAND_1 = ["Acme", "Sakura", "Blue", "Green", "Silver", "Golden", "Northern", "Pacific",
              "Summit", "Harbor", "Maple", "Cedar", "Orion", "Nova", "Delta", "Apex",
              "Bright", "Clear", "Swift", "Prime"]
EN_BRAND_2 = ["Trading", "Foods", "Systems", "Logistics", "Design", "Works", "Partners",
              "Labs", "Supply", "Media", "Energy", "Craft", "Solutions", "Materials", "Tech"]
EN_FORMS = {  # canonical form -> accepted dirty spellings
    "Inc.": ["Inc", "Inc.", "INC."],
    "Co., Ltd.": ["Co.,Ltd", "Co., Ltd.", "Co. Ltd.", "Co Ltd", "CO., LTD."],
    "LLC": ["LLC", "llc", "Llc"],
    "Corp.": ["Corp", "Corp.", "CORP."],
}
JP_BRAND_1 = ["ABC", "XYZ", "サクラ", "ミライ", "東和", "青葉", "日の出", "大和", "北斗", "星野",
              "光洋", "富士見", "若葉", "緑川", "白鳥"]
JP_BRAND_2 = ["商事", "食品", "工業", "電機", "物産", "運輸", "製作所", "建設", "産業", "貿易",
              "化学", "情報システム"]

JP_PLACES = [("東京都", "千代田区"), ("東京都", "港区"), ("東京都", "新宿区"),
             ("大阪府", "大阪市北区"), ("愛知県", "名古屋市中区"), ("北海道", "札幌市中央区"),
             ("福岡県", "福岡市博多区"), ("神奈川県", "横浜市西区"), ("京都府", "京都市中京区"),
             ("宮城県", "仙台市青葉区")]
JP_TOWNS = ["架空町", "見本町", "試験町", "仮想町", "サンプル台", "ダミー通", "仮名町", "例示町",
            "テスト町", "模擬町"]
EN_STREETS = ["Sample St", "Example Ave", "Test Rd", "Demo Blvd", "Fiction Ln", "Mockup Way",
              "Placeholder Dr", "Sandbox Ct"]
EN_CITIES = [("Springfield", "IL"), ("Shelbyville", "IL"), ("Fairview", "OH"),
             ("Riverton", "WY"), ("Lakeside", "CA"), ("Oakdale", "MN"), ("Georgetown", "TX"),
             ("Franklin", "TN")]
# No area code starting with 50/70/80/90: a bare 10-digit number like that is ambiguous
# with a Japanese mobile number whose leading 0 was lost (REVIEW_1 R1-1).
NA_AREAS = [212, 312, 415, 617, 206, 303, 404, 512, 602, 714, 305, 213, 646, 619, 720]
AREA3 = ("045", "052", "092", "011", "075")
WEEKDAYS_JP = "月火水木金土日"
WEEKDAYS_EN = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
MONTHS_FULL = ["January", "February", "March", "April", "May", "June", "July", "August",
               "September", "October", "November", "December"]

EMPTY_LOOKS = ["", "", "N/A", "-", "n/a", "null", "none"]
TYPO_DOMAINS = [("gmial.com", "gmail.com"), ("gmai.com", "gmail.com"),
                ("hotmial.com", "hotmail.com"), ("yahooo.co.jp", "yahoo.co.jp"),
                ("outlok.com", "outlook.com")]
INVALID_DATES = ["2024-02-30", "2023/13/01", "2022-02-29", "2024-04-31", "2021/6/31"]
INVALID_PHONES = ["12345", "090-1234-567", "abc-defg", "0000000000", "+81 123"]

FULLWIDTH = {c: c + 0xFEE0 for c in range(0x21, 0x7F)}


def to_fullwidth(s):
    return s.translate(FULLWIDTH)


# ------------------------------------------------------------ record building
class Gen:
    def __init__(self, seed=42):
        self.rng = random.Random(seed)
        jp = [(s, g) for s in JP_SURNAMES for g in JP_GIVEN]
        en = [(f, l) for f in EN_FIRST for l in EN_LAST]
        self.rng.shuffle(jp)
        self.rng.shuffle(en)
        self.jp_names, self.en_names = iter(jp), iter(en)
        self.used_email, self.used_phone = set(), set()

    # -- canonical values
    def phone(self, lang, force=None):
        r = self.rng
        while True:
            x = r.random()
            if force:
                kind = force
            elif lang == "jp":
                kind = ("mobile" if x < .55 else "fixed" if x < .75 else "area3" if x < .87
                        else "ip" if x < .93 else "free" if x < .97 else "na" if x < .99 else "uk")
            else:
                kind = ("na" if x < .55 else "uk" if x < .80 else "mobile" if x < .90 else "fixed")
            if kind == "mobile":
                p = f"{r.choice(['090', '080', '070'])}-{r.randint(1000, 9999)}-{r.randint(0, 9999):04d}"
            elif kind == "ip":
                p = f"050-{r.randint(1000, 9999)}-{r.randint(0, 9999):04d}"
            elif kind == "fixed":
                p = f"{r.choice(['03', '06'])}-{r.randint(1000, 9999)}-{r.randint(0, 9999):04d}"
            elif kind == "area3":
                p = f"{r.choice(AREA3)}-{r.randint(200, 999)}-{r.randint(0, 9999):04d}"
            elif kind == "free":
                p = f"0120-{r.randint(100, 999)}-{r.randint(100, 999)}"
            elif kind == "na":
                p = f"+1 {r.choice(NA_AREAS)}-555-01{r.randint(0, 99):02d}"
            else:
                p = f"+442079460{r.randint(0, 999):03d}"
            if p not in self.used_phone:
                self.used_phone.add(p)
                return p

    def address(self, lang):
        r = self.rng
        if lang == "jp":
            pref, ward = r.choice(JP_PLACES)
            return f"{pref}{ward}{r.choice(JP_TOWNS)}{r.randint(1, 30)}-{r.randint(1, 20)}-{r.randint(1, 15)}"
        city, state = r.choice(EN_CITIES)
        return f"{r.randint(10, 999)} {r.choice(EN_STREETS)}, {city}, {state}"

    def date(self):
        return date(2019, 1, 1) + timedelta(days=self.rng.randint(0, 2556))

    def company(self, lang):
        r = self.rng
        jp_company = r.random() < (0.85 if lang == "jp" else 0.15)
        if jp_company:
            brand = r.choice(JP_BRAND_1) + r.choice(JP_BRAND_2)
            return ("株式会社" + brand) if r.random() < 0.65 else (brand + "株式会社")
        return f"{r.choice(EN_BRAND_1)} {r.choice(EN_BRAND_2)} {r.choice(list(EN_FORMS))}"

    def record(self):
        r = self.rng
        lang = "jp" if r.random() < 0.6 else "en"
        if lang == "jp":
            (sk, sr), (gk, gr) = next(self.jp_names)
            name, local = f"{sk} {gk}", f"{gr}.{sr}"
        else:
            first, last = next(self.en_names)
            name, local = f"{first} {last}", f"{first.lower()}.{last.lower()}"
        email = f"{local}@{r.choice(DOMAINS)}"
        while email in self.used_email:
            email = f"{local}{r.randint(1, 99)}@{r.choice(DOMAINS)}"
        self.used_email.add(email)
        return {"lang": lang, "name": name, "email": email, "phone": self.phone(lang),
                "address": self.address(lang), "date": self.date().isoformat(),
                "company": self.company(lang)}

    # -- dirt (only variants the normalizer must map back to the canonical form)
    def pad(self, v, jp):
        r = self.rng
        x = r.random()
        if x < 0.06:
            return f" {v} "
        if x < 0.10:
            return f"　{v}　" if jp else f"  {v}"
        if x < 0.13:
            return f"{v}  "
        return v

    def dirty_name(self, rec):
        r, x = self.rng, self.rng.random()
        n = rec["name"]
        if rec["lang"] == "en":
            n = n.upper() if x < .15 else n.lower() if x < .30 else n.replace(" ", "  ") if x < .40 else n
        else:
            n = n.replace(" ", "　") if x < .20 else n.replace(" ", "  ") if x < .30 else n
        return self.pad(n, rec["lang"] == "jp")

    def dirty_email(self, email):
        x = self.rng.random()
        local, domain = email.split("@")
        if x < .15:
            email = f"{local.title()}@{domain.capitalize()}"
        elif x < .22:
            email = email.upper()
        elif x < .35:
            email = to_fullwidth(email)
        return self.pad(email, False)

    def dirty_phone(self, p):
        r, x = self.rng, self.rng.random()
        if p.startswith("+1 "):
            a, b, c = p[3:].split("-")
            options = [p, f"({a}) {b}-{c}", f"{a}.{b}.{c}", f"1-{a}-{b}-{c}", f"{a}{b}{c}",
                       f"+1 ({a}) {b}-{c}", f"+1-{a}-{b}-{c}"]
        elif p.startswith("+44"):
            d = p[3:]
            options = [p, f"+44 {d[:2]} {d[2:6]} {d[6:]}", f"+44-{d[:2]}-{d[2:6]}-{d[6:]}",
                       f"+44 (0){d[:2]} {d[2:6]} {d[6:]}"]
        else:
            parts = p.split("-")
            intl = "-".join([parts[0][1:], *parts[1:]])
            options = [p, p, "".join(parts), " ".join(parts),
                       p.translate({ord(c): c2 for c, c2 in zip("0123456789-", "０１２３４５６７８９－")}),
                       f"({parts[0]}){'-'.join(parts[1:])}", f"+81 {intl}", f"+81-{intl}"]
            if parts[0] in AREA3:  # bare digits would (correctly) stay unformatted: not a "clean" dirt
                options.remove("".join(parts))
        return self.pad(r.choice(options) if x < .6 else p, False)

    def dirty_address(self, addr, lang, xlsx):
        x = self.rng.random()
        if lang == "jp":
            return self._jp_address_dirt(addr, x, xlsx)
        return self.pad(addr.replace(", ", ",  ") if x < .15 else addr, False)

    def _jp_address_dirt(self, addr, x, xlsx):
        i = next(k for k, ch in enumerate(addr) if ch.isdigit())
        head, tail = addr[:i], addr[i:]
        if x < .15:
            tail = tail.translate({ord(c): c2 for c, c2 in zip("0123456789-", "０１２３４５６７８９－")})
        elif x < .27:
            tail = tail.replace("-", "ー")
        elif x < .40:
            tail = tail.replace("-", "−" if xlsx else "－")  # U+2212 is not encodable in cp932
        return self.pad(head + tail, True)

    def dirty_company(self, comp):
        r, x = self.rng, self.rng.random()
        if comp.startswith("株式会社"):
            brand = comp[4:]
            options = [comp, comp, f"(株){brand}", f"㈱{to_fullwidth(brand)}", f"(株){to_fullwidth(brand)}"]
            c = r.choice(options) if x < .5 else comp
        elif comp.endswith("株式会社"):
            brand = comp[:-4]
            options = [comp, f"{brand}(株)", f"{to_fullwidth(brand)}(株)", f"{brand}㈱"]
            c = r.choice(options) if x < .5 else comp
        else:
            base, _, form = comp.rpartition(" ")
            if comp.endswith("Co., Ltd."):
                base, form = comp[:-len(" Co., Ltd.")], "Co., Ltd."
            if x < .5:
                base = base.lower() if r.random() < .5 else base
                c = base + r.choice([" ", ", "]) + r.choice(EN_FORMS[form])
            else:
                c = comp
        return self.pad(c, True)

    def dirty_date(self, iso, xlsx):
        r = self.rng
        y, m, d = int(iso[:4]), int(iso[5:7]), int(iso[8:])
        options = [iso, iso, f"{y}/{m}/{d}", f"{y}/{m:02d}/{d:02d}", f"{y}.{m}.{d}",
                   f"{y}年{m}月{d}日", f"{y}{m:02d}{d:02d}",
                   f"{MONTHS_FULL[m - 1][:3]} {d}, {y}", f"{d} {MONTHS_FULL[m - 1][:3]} {y}",
                   f"{MONTHS_FULL[m - 1]} {d}, {y}", f"{d} {MONTHS_FULL[m - 1]} {y}",
                   f"{y}-{m:02d}-{d:02d} 09:30:00", to_fullwidth(f"{y}/{m}/{d}")]
        wd = date(y, m, d).weekday()
        options += [f"{y}-{m:02d}-{d:02d}T10:30:00Z", f"{y}-{m:02d}-{d:02d}T10:30:00.000+09:00",
                    f"{d}-{MONTHS_FULL[m - 1][:3]}-{y}", f"{y}年{m}月{d}日({WEEKDAYS_JP[wd]})",
                    f"{y}/{m}/{d} ({WEEKDAYS_EN[wd]})"]
        if d > 12:
            options += [f"{d:02d}/{m:02d}/{y}", f"{d}.{m}.{y}", f"{m:02d}/{d:02d}/{y}", f"{m}-{d}-{y}"]
        if d == m:
            options.append(f"{d}/{m}/{y}")
        if xlsx:
            options += [datetime(y, m, d)] * 4 + [(date(y, m, d) - date(1899, 12, 30)).days]
        v = r.choice(options)
        return self.pad(v, False) if isinstance(v, str) else v

    def raw(self, rec, xlsx):
        return {
            "name": self.dirty_name(rec),
            "email": self.dirty_email(rec["email"]),
            "phone": self.dirty_phone(rec["phone"]) if rec["phone"] else "",
            "address": self.dirty_address(rec["address"], rec["lang"], xlsx) if rec["address"] else "",
            "date": self.dirty_date(rec["date"], xlsx) if rec["date"] else "",
            "company": self.dirty_company(rec["company"]) if rec["company"] else "",
        }


def invalid_email(local, domain, i):
    tld = domain.split(".")[-1]
    return [f"{local}@@{domain}", f"{local}.{domain}", f"{local}@{domain.split('.')[0]}",
            f"{local}@.{tld}", f"{local.replace('.', '..', 1)}@{domain}",
            f"{local}@{domain.replace('.', ',')}", f"{local.replace('.', ' ', 1)}@{domain}",
            f".{local}@{domain}", f"{local}.@{domain}"][i % 9]


def ambiguous_date(g):
    while True:
        d = g.date()
        if d.day <= 12 and d.day != d.month:
            break
    a, b = (d.month, d.day) if g.rng.random() < .5 else (d.day, d.month)
    sep = g.rng.choice("/-.")
    return f"{a:02d}{sep}{b:02d}{sep}{d.year}", f"{d.year}-{a:02d}-{b:02d} or {d.year}-{b:02d}-{a:02d}"


# ------------------------------------------------------------------- main
def build(seed=42):
    g = Gen(seed)
    r = g.rng
    n_independent = CSV_ROWS + XLSX_ROWS - N_EXACT - N_EMAIL - N_POSSIBLE
    kinds = (["invalid_email"] * N_INVALID_EMAIL + ["email_domain_typo"] * N_TYPO
             + ["ambiguous_date"] * N_AMBIGUOUS + ["invalid_date"] * N_INVALID_DATE
             + ["invalid_phone"] * N_INVALID_PHONE + ["missing_name"] * N_MISSING_NAME
             + ["no_contact"] * N_NO_CONTACT)
    special = [("phone_leading_zero", "xlsx")] * LEADING_ZERO_PER_FILE         + [("phone_leading_zero", "csv")] * LEADING_ZERO_PER_FILE
    n_regular = n_independent - len(special)
    kinds += ["clean"] * (n_regular - len(kinds))
    r.shuffle(kinds)
    n_xlsx_regular = XLSX_ROWS - 16 - LEADING_ZERO_PER_FILE  # 16 duplicate rows live in the xlsx
    files = ["xlsx"] * n_xlsx_regular + ["csv"] * (n_regular - n_xlsx_regular)
    r.shuffle(files)
    pairs = list(zip(kinds, files)) + special
    r.shuffle(pairs)

    entries, lz_count = [], {"xlsx": 0, "csv": 0}
    for kind, f in pairs:
        rec = g.record()
        e = {"kind": kind, "file": f, "rec": rec, "issues": [], "anchor": None, "expected": {}}
        if kind == "phone_leading_zero":  # 4 mobile + 1 landline per file
            e["variant"] = "fixed" if lz_count[f] == LEADING_ZERO_PER_FILE - 1 else "mobile"
            lz_count[f] += 1
            rec["phone"] = g.phone("jp", force=e["variant"])
        entries.append(e)
    clean = [e for e in entries if e["kind"] == "clean"]

    # soft-missing optional values (kept clean)
    fields = ["address", "date", "company", "phone"] * (N_SOFT_MISSING // 4)
    r.shuffle(fields)
    soft = r.sample(range(len(clean)), N_SOFT_MISSING)
    for idx, field in zip(soft, fields):
        clean[idx]["rec"][field] = ""
        clean[idx]["issues"] = [f"missing_{field}"]
        clean[idx]["soft"] = True

    # landlines with a 3-digit area code written as bare digits: kept as digits, soft issue
    area3 = [e for e in clean if not e.get("soft") and e["rec"]["phone"][:3] in AREA3
             and len(e["rec"]["phone"]) == 12]
    for e in r.sample(area3, N_UNVERIFIED_PHONE):
        e["rec"]["phone"] = e["rec"]["phone"].replace("-", "")
        e["issues"] = ["phone_format_unverified"]
        e["unverified"] = True

    # anchors: clean rows with every field present
    pools = {f: [e for e in clean if e["file"] == f and not e.get("soft") and not e.get("unverified")]
             for f in ("xlsx", "csv")}
    for pool in pools.values():
        r.shuffle(pool)
    dependents = []
    spec = ([("duplicate_exact", "xlsx", "xlsx")] * 8 + [("duplicate_exact", "csv", "csv")] * 32
            + [("duplicate_email", "xlsx", "xlsx")] * 6 + [("duplicate_email", "csv", "csv")] * 14
            + [("duplicate_email", "csv", "xlsx")] * 10
            + [("possible_name", "csv", "csv")] * 5 + [("possible_name", "xlsx", "xlsx")] * 2
            + [("possible_name", "csv", "xlsx")] * 3
            + [("possible_phone", "csv", "csv")] * 3 + [("possible_phone", "csv", "xlsx")] * 2)
    for kind, dep_file, anchor_file in spec:
        anchor = pools[anchor_file].pop()
        rec = dict(anchor["rec"])
        if kind == "duplicate_email":
            if r.random() < .5:
                while rec["address"] == anchor["rec"]["address"]:
                    rec["address"] = g.address(rec["lang"])
            else:
                while rec["date"] == anchor["rec"]["date"]:
                    rec["date"] = g.date().isoformat()
        elif kind.startswith("possible"):
            local = f"alt{len(g.used_email)}.{rec['email'].split('@')[0]}"
            rec["email"] = f"{local}@{r.choice(DOMAINS)}"
            g.used_email.add(rec["email"])
            if kind == "possible_name":
                rec["phone"] = g.phone(rec["lang"])
            else:  # same phone, different person and company
                (sk, sr), (gk, gr) = next(g.jp_names)
                rec["lang"], rec["name"] = "jp", f"{sk} {gk}"
                rec["company"] = g.company("jp")
            rec["address"], rec["date"] = g.address(rec["lang"]), g.date().isoformat()
        dependents.append({"kind": kind, "file": dep_file, "rec": rec, "issues": [],
                           "anchor": anchor, "expected": {}})

    # raw cells and expected results
    counters = {}
    for e in entries + dependents:
        xlsx = e["file"] == "xlsx"
        raw = g.raw(e["rec"], xlsx)
        kind = e["kind"]
        i = counters[kind] = counters.get(kind, -1) + 1
        local, domain = e["rec"]["email"].split("@")
        status, reasons, extra = "clean", [], {}
        if kind == "clean":
            pass
        elif kind in ("duplicate_exact", "duplicate_email"):
            status, reasons = "duplicate", [kind]
        elif kind.startswith("possible"):
            status, reasons = "needs_review", ["possible_duplicate"]
        elif kind == "invalid_email":
            raw["email"], status, reasons = invalid_email(local, domain, i), "invalid", [kind]
        elif kind == "email_domain_typo":
            typo, fixed = TYPO_DOMAINS[i % len(TYPO_DOMAINS)]
            raw["email"], status, reasons = f"{local}@{typo}", "needs_review", [kind]
            extra["suggestion"] = f"{local}@{fixed}"
        elif kind == "ambiguous_date":
            raw["date"], extra["suggestion"] = ambiguous_date(g)
            status, reasons = "needs_review", [kind]
        elif kind == "invalid_date":
            raw["date"], status, reasons = INVALID_DATES[i % len(INVALID_DATES)], "needs_review", [kind]
        elif kind == "invalid_phone":
            raw["phone"], status, reasons = INVALID_PHONES[i % len(INVALID_PHONES)], "invalid", [kind]
        elif kind == "missing_name":
            raw["name"], status, reasons = r.choice(["", "", "-", "N/A"]), "invalid", [kind]
        elif kind == "no_contact":
            raw["email"], raw["phone"] = r.choice(EMPTY_LOOKS), r.choice(EMPTY_LOOKS)
            status, reasons = "invalid", [kind]
        elif kind == "phone_leading_zero":
            canonical = e["rec"]["phone"]
            digits = canonical.replace("-", "")[1:]  # the leading 0 is lost
            raw["phone"] = int(digits) if xlsx else digits  # xlsx: a numeric cell
            status, reasons = "needs_review", [kind]
            extra["suggestion"] = (
                f"{canonical} (JP, leading 0 lost?) or +1 {digits[:3]}-{digits[3:6]}-{digits[6:]} (US)"
                if e["variant"] == "mobile" else canonical)
        if e.get("unverified"):
            digits = e["rec"]["phone"]
            raw["phone"] = r.choice([digits, to_fullwidth(digits), f" {digits} "])
        if e.get("soft"):  # soft-missing: blank looks
            field = e["issues"][0][len("missing_"):]
            raw[field] = r.choice(EMPTY_LOOKS)
        e["raw"], e["expected"] = raw, {"status": status, "reasons": reasons, **extra}

    # layout: independents shuffled, dependents inserted after their anchor
    layout = {"xlsx": [e for e in entries if e["file"] == "xlsx"],
              "csv": [e for e in entries if e["file"] == "csv"]}
    r.shuffle(dependents)
    for dep in dependents:
        rows = layout[dep["file"]]
        pos = r.randint(rows.index(dep["anchor"]) + 1, len(rows)) if dep["anchor"] in rows \
            else r.randint(0, len(rows))
        rows.insert(pos, dep)
    assert len(layout["csv"]) == CSV_ROWS and len(layout["xlsx"]) == XLSX_ROWS
    return layout


def cell_for_csv(v):
    assert isinstance(v, str), v
    return v


def write_all(layout, out_dir):
    out_dir.mkdir(parents=True, exist_ok=True)
    # CSV (Shift_JIS / cp932)
    with open(out_dir / CSV_NAME, "w", encoding="cp932", newline="") as f:
        w = csv.writer(f, lineterminator="\r\n")
        w.writerow(CSV_HEADER)
        for e in layout["csv"]:
            w.writerow([cell_for_csv(e["raw"][c]) for c in FIELDS])
    # XLSX
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(XLSX_HEADER)
    for e in layout["xlsx"]:
        ws.append([e["raw"][c] for c in FIELDS])
    for row in ws.iter_rows(min_row=2):
        if isinstance(row[4].value, datetime):
            row[4].number_format = "yyyy-mm-dd"
    save_workbook(wb, out_dir / XLSX_NAME)

    # ground truth
    file_names = {"xlsx": XLSX_NAME, "csv": CSV_NAME}
    where = {id(e): (file_names[k], i + 2) for k, rows in layout.items() for i, e in enumerate(rows)}
    truth_rows = []
    planted = dict.fromkeys(["duplicate_exact", "duplicate_email", "possible_duplicate",
                             "invalid_email", "email_domain_typo", "ambiguous_date",
                             "invalid_date", "invalid_phone", "missing_name", "no_contact",
                             "soft_missing", "phone_leading_zero", "phone_format_unverified"], 0)
    for k in ("xlsx", "csv"):
        for e in layout[k]:
            name, row_no = where[id(e)]
            exp = e["expected"]
            kind = "possible_duplicate" if e["kind"].startswith("possible") else e["kind"]
            item = {"source_file": name, "source_row": row_no, "planted": kind,
                    "expected_status": exp["status"], "expected_reasons": exp["reasons"]}
            if e["kind"] == "clean":
                item["planted"] = kind = ("soft_missing" if e.get("soft") else
                                          "phone_format_unverified" if e.get("unverified") else "clean")
                item["expected_issues"] = e["issues"]
                item["canonical"] = {c: e["rec"][c] for c in FIELDS}
            if e["anchor"] is not None:
                a_name, a_row = where[id(e["anchor"])]
                item["duplicate_of"] = f"{a_name}:{a_row}"
            if "suggestion" in exp:
                item["suggestion"] = exp["suggestion"]
            if kind in planted:
                planted[kind] += 1
            truth_rows.append(item)
    truth = {"seed": 42, "files": sorted(file_names.values()), "planted_counts": planted,
             "rows": truth_rows}
    with open(out_dir / "ground_truth.json", "w", encoding="utf-8") as f:
        json.dump(truth, f, ensure_ascii=False, indent=1)
    return planted


def main():
    out_dir = ROOT / "input"
    planted = write_all(build(42), out_dir)
    print(f"Generated sample data in {out_dir}")
    for k, v in planted.items():
        print(f"  {k}: {v}")


if __name__ == "__main__":
    main()
