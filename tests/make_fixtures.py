"""Regenerates the small fixed fixtures in tests/fixtures/ (already committed; run only to rebuild).

All data is fictional; e-mail addresses use example.com/.net/.org.
"""
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

OUT = Path(__file__).resolve().parent / "fixtures"
OUT.mkdir(exist_ok=True)
HEADER = "name,email,phone,address,date,company\n"


def write(name, text, encoding="utf-8"):
    (OUT / name).write_bytes(text.encode(encoding))


write("normal.csv", HEADER + (
    'Alice Baker,alice.baker@example.com,090-1111-2222,"1 Sample St, Springfield, IL",2024-01-05,Acme Trading Inc.\n'
    ' JOHN  CARTER ,John.Carter@Example.com,09033334444,,2024/1/6,acme labs llc\n'
    '山田　太郎,taro.yamada@example.net,+81 90-5555-6666,東京都千代田区架空町１－２－３,2024年1月7日,(株)ABC商事\n'))

(OUT / "empty_zero.csv").write_bytes(b"")
write("header_only.csv", HEADER)
write("missing_required.csv", "name,phone\nAlice Baker,090-1111-2222\n")

write("duplicates.csv", HEADER + (
    'Bob Dawson,bob.dawson@example.com,090-1000-0001,1 Sample St,2024-02-01,Blue Works Inc.\n'
    'BOB DAWSON,BOB.DAWSON@EXAMPLE.COM,09010000001,1 Sample St,2024/2/1,blue works inc\n'
    'Bob Dawson,bob.dawson@example.com,090-1000-0002,2 Sample St,2024-02-02,Blue Works Inc.\n'
    'Bob Dawson,bob.alt@example.com,090-1000-0003,3 Sample St,2024-02-03,Blue Works Inc.\n'))

write("invalid_email.csv", HEADER + (
    'Carl Ellis,carl@@example.com,090-2000-0001,1 Test Rd,2024-03-01,Acme Labs LLC\n'
    'Dana Foster,dana@-example.com,090-2000-0002,2 Test Rd,2024-03-02,Acme Labs LLC\n'
    'Eve Grant,eve.grant@gmial.com,090-2000-0003,3 Test Rd,2024-03-03,Acme Labs LLC\n'
    'Fay Hayes,fay.hayes@example.com,090-2000-0004,4 Test Rd,2024-03-04,Acme Labs LLC\n'))

write("missing_values.csv", HEADER + (
    'Gina Irwin,gina.irwin@example.com,,,,\n'
    'Hank Jensen,,090-3000-0002,5 Demo Blvd,2024-04-02,Delta Media Corp.\n'
    'Iris Keller,,,6 Demo Blvd,2024-04-03,Delta Media Corp.\n'
    ',jane.lawson@example.com,090-3000-0004,7 Demo Blvd,2024-04-04,Delta Media Corp.\n'))

write("dates.csv", HEADER + (
    'Kim Nolan,kim.nolan@example.com,090-4000-0001,,03/04/2024,\n'
    'Leo Owens,leo.owens@example.com,090-4000-0002,,25/03/2024,\n'
    'Mia Parker,mia.parker@example.com,090-4000-0003,,2024-01-05T10:30:00.000+09:00,\n'
    'Ned Reed,ned.reed@example.com,090-4000-0004,,2024-01-05T10:30:00Z,\n'
    'Ola Sutton,ola.sutton@example.com,090-4000-0005,,5-Jan-2024,\n'
    'Pam Turner,pam.turner@example.com,090-4000-0006,,2024年1月5日(金),\n'
    'Ray Walsh,ray.walsh@example.com,090-4000-0007,,2024-02-30,\n'))

write("phones.csv", HEADER + (
    'Sam Yates,sam.yates@example.com,9012345678,,,\n'
    'Tia Abbott,tia.abbott@example.com,312345678,,,\n'
    'Uma Barnes,uma.barnes@example.com,090-4444-5555 ext. 12,,,\n'
    'Vic Chambers,vic.chambers@example.com,+44 (0)20 7946 0958,,,\n'
    'Wes Dixon,wes.dixon@example.com,0467123456,,,\n'
    'Xia Eaton,xia.eaton@example.com,+81 467-12-3456,,,\n'
    'Yan Garrett,yan.garrett@example.com,12345,,,\n'))

write("sjis.csv", "氏名,メール,電話番号,住所,登録日,会社名\n"
      "佐藤　花子,hanako.sato@example.com,０９０－１２３４－５６７８,大阪府大阪市北区見本町１ー２ー３,2024/1/5,ＡＢＣ商事(株)\n"
      "鈴木 一郎,ichiro.suzuki@example.org,03-1234-5678,東京都港区試験町4-5-6,20240106,サクラ食品株式会社\n",
      encoding="cp932")

wb = Workbook()
ws = wb.active
ws.append(["Full Name", "E-mail", "Tel", "Address", "Created", "Company"])
ws.append(["Zed Holland", "zed.holland@example.com", 9012345678, "1 Mockup Way", datetime(2024, 3, 1), "Orion Labs Inc."])
ws.append(["Amy Ingram", "amy.ingram@example.com", "090-5000-0002", "2 Mockup Way", 45352, "Orion Labs Inc."])
wb.save(OUT / "excel_input.xlsx")
print("fixtures written to", OUT)
