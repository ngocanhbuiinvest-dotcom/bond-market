# -*- coding: utf-8 -*-
"""
Sinh dashboard TPDN VBMA (HTML tu chua) tu file CSV tong hop.
Chay:  python build_dashboard.py
Output: TPDN_VBMA_dashboard.html  (mo bang trinh duyet -> nut 'In / Xuat PDF')
"""
import csv, json, os, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
CSV = os.path.normpath(os.path.join(HERE, "..", "TPDN_VBMA", "TPDN_VBMA_tong_hop_thang.csv"))
OUT = os.path.join(HERE, "TPDN_VBMA_dashboard.html")
TPL = os.path.join(HERE, "template.html")

def num(x):
    """Blank/'-' -> None (JSON null) de dashboard hien '—' thay vi 0% sai lech."""
    x = (x or "").strip()
    if x in ("", "-"): return None
    try:
        return int(x)
    except ValueError:
        return float(x)

rows = []
with open(CSV, "r", encoding="utf-8-sig", newline="") as f:
    r = csv.reader(f)
    next(r)  # header
    for line in r:
        if not line or not line[0].strip():
            continue
        rows.append({
            "thang": line[0], "ky": line[1], "ngay": line[2],
            "rl": num(line[3]), "sodotRL": num(line[4]),
            "cc": num(line[5]), "sodotCC": num(line[6]),
            "tong": num(line[7]), "mom": num(line[8]), "yoy": num(line[9]),
            "lkRL": num(line[10]), "lkCC": num(line[11]), "lkTong": num(line[12]),
            "mualai": num(line[13]), "socham": num(line[14]), "goccham": num(line[15]),
            "gtgd": num(line[16]), "bq": num(line[17]),
            "momtc": num(line[18]), "yoytc": num(line[19]), "lkgtgd": num(line[20]),
        })

latest = rows[-1]
period = f"{rows[0]['ky']} – {rows[-1]['ky']}"
generated = datetime.date.today().strftime("%d/%m/%Y")

DATA_JSON = json.dumps(rows, ensure_ascii=False)

with open(TPL, "r", encoding="utf-8") as f:
    TEMPLATE = f.read()

html = TEMPLATE
html = html.replace("__DATA__", DATA_JSON)
html = html.replace("__PERIOD__", period)
html = html.replace("__GENERATED__", generated)
html = html.replace("__LATESTKY__", latest["ky"])

with open(OUT, "w", encoding="utf-8") as f:
    f.write(html)
print("Da tao:", OUT)
print("So thang:", len(rows), "| Moi nhat:", latest["ky"])
