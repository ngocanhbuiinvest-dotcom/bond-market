# -*- coding: utf-8 -*-
"""
Scraper PHÂN NHÓM NGÀNH ICB từ API VietCap (iq.vietcap.com.vn) — CÔNG KHAI, không cần đăng nhập.
Mục đích: lấy phân loại ngành ICB (Industry Classification Benchmark) 4 cấp cho DN niêm yết
HOSE/HNX/UPCOM để ĐỐI CHIẾU / BỔ SUNG cho bộ phân loại theo từ khóa (sector_map.py) — đặc biệt
cứu các TCPH bị rơi vào nhóm "Khác".

Hai endpoint REST (GET, JSON, không auth):
  1) /api/iq-insight-service/v2/company/search-bar?language=1  -> ~2,088 công ty + ICB Lv1..Lv4
  2) /api/iq-insight-service/v1/sectors/icb-codes              -> cây mã ngành ICB (vi/en/code/level)

LƯU Ý ĐỘ PHỦ: VietCap chỉ có DN NIÊM YẾT (mã cổ phiếu). TCPH phát hành TPDN nhiều là SPV chưa niêm
yết -> chỉ khớp ~25% số TCPH nhưng ~63% GIÁ TRỊ phát hành (các DN lớn/ngân hàng đều niêm yết).
Khớp bằng: (a) tiền tố MÃ CK trong tên TCPH ("ACB - Ngân hàng..."), (b) tên chuẩn hoá.

Xuất: vietcap_companies_raw.csv/json (phẳng icbLv1..4) + vietcap_icb_raw.csv/json
"""
import csv
import json
import sys
import time
from datetime import datetime

import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
VERIFY_SSL = True   # iq.vietcap.com.vn có cert hợp lệ

BASE = "https://iq.vietcap.com.vn/api/iq-insight-service"
COMPANY_URL = BASE + "/v2/company/search-bar?language=1"   # language=1 = tiếng Việt
ICB_URL = BASE + "/v1/sectors/icb-codes"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
    "Accept": "application/json",
    "Referer": "https://iq.vietcap.com.vn/",
}

COMP_COLS = ["code", "name", "shortName", "organCode", "floor", "comTypeCode", "isBank",
             "icb1_code", "icb1_name", "icb2_code", "icb2_name",
             "icb3_code", "icb3_name", "icb4_code", "icb4_name"]


def get_json(url):
    for attempt in range(3):
        try:
            r = requests.get(url, headers=HEADERS, timeout=45, verify=VERIFY_SSL)
            r.raise_for_status()
            j = r.json()
            if not j.get("successful", True):
                raise RuntimeError(f"API báo lỗi: {j.get('msg')}")
            return j.get("data", j)
        except Exception as e:
            print(f"  ! lỗi {url[-40:]} (lần {attempt+1}): {e}", file=sys.stderr)
            time.sleep(2)
    raise RuntimeError(f"Không lấy được {url}")


def flatten_company(c):
    def lv(key):
        d = c.get(key) or {}
        return d.get("code", ""), d.get("name", "")
    rec = {k: c.get(k, "") for k in ("code", "name", "shortName", "organCode",
                                     "floor", "comTypeCode", "isBank")}
    for i in range(1, 5):
        code, name = lv(f"icbLv{i}")
        rec[f"icb{i}_code"], rec[f"icb{i}_name"] = code, name
    return rec


def scrape():
    print("Lấy danh sách công ty + ICB...")
    comps = get_json(COMPANY_URL)
    comp_rows = [flatten_company(c) for c in comps]
    print(f"  -> {len(comp_rows)} công ty")
    print("Lấy cây mã ngành ICB...")
    icb = get_json(ICB_URL)
    print(f"  -> {len(icb)} mã ngành ICB")
    return comp_rows, icb


def save(comp_rows, icb):
    with open("vietcap_companies_raw.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=COMP_COLS, extrasaction="ignore")
        w.writeheader()
        w.writerows(comp_rows)
    with open("vietcap_companies_raw.json", "w", encoding="utf-8") as f:
        json.dump(comp_rows, f, ensure_ascii=False)
    with open("vietcap_icb_raw.json", "w", encoding="utf-8") as f:
        json.dump(icb, f, ensure_ascii=False)
    if icb:
        cols = list(icb[0].keys())
        with open("vietcap_icb_raw.csv", "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
            w.writeheader()
            w.writerows(icb)
    print(f"Đã lưu vietcap_companies_raw.csv/json ({len(comp_rows)}) + vietcap_icb_raw.csv/json ({len(icb)})")


if __name__ == "__main__":
    t0 = time.time()
    print(f"Bắt đầu lấy phân nhóm ngành ICB (VietCap) lúc {datetime.now():%Y-%m-%d %H:%M:%S}")
    comp_rows, icb = scrape()
    save(comp_rows, icb)
    print(f"Xong sau {time.time()-t0:.0f}s.")
