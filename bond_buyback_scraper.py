# -*- coding: utf-8 -*-
"""
Scraper CBTT KẾT QUẢ MUA LẠI TP trước hạn (trong nước) - CBIS / HNX
Tab "CBTT kết quả mua lại TP trước hạn" trên trang Thông tin phát hành.

Cùng cơ chế POST/CP-TOKEN như bond_issuance_scraper, khác:
  - Chỉ số tab trong arrCurrentPage/arrNumberRecord = 1 (REPURCHASE_TAB)
  - Bảng kết quả: #tbRepurchaseResult (18 cột, đã có sẵn giá trị theo mệnh giá)

Mỗi dòng = một ĐỢT mua lại của một mã TP (một mã có thể mua lại nhiều đợt).
Xuất: bond_buyback_raw.csv / .json
"""
import csv
import re
import sys
import time
import json
from datetime import datetime

import requests
import urllib3
from bs4 import BeautifulSoup

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
VERIFY_SSL = False

BASE = "https://cbonds.hnx.vn"
PAGE_URL = BASE + "/to-chuc-phat-hanh/thong-tin-phat-hanh"
SEARCH_URL = PAGE_URL + "/tim-kiem"
PAGE_SIZE = 100
N_TABS = 12
REPURCHASE_TAB = 1
DELAY = 0.8

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
    "X-Requested-With": "XMLHttpRequest",
    "Origin": BASE,
    "Referer": PAGE_URL,
}

COLUMNS = [
    "stt", "ngay_dang_tin", "ten_dn", "ma_tp", "menh_gia", "ky_han",
    "ngay_phat_hanh", "ngay_dao_han", "gt_phat_hanh", "gt_dang_luu_hanh",
    "gt_mua_lai", "sl_mua_lai", "gt_con_lai", "sl_con_lai",
    "ngay_mua_lai", "tinh_trang", "ghi_chu", "van_ban",
]
NUM_COLS = ["menh_gia", "gt_phat_hanh", "gt_dang_luu_hanh", "gt_mua_lai",
            "sl_mua_lai", "gt_con_lai", "sl_con_lai"]


def get_session_token():
    s = requests.Session()
    s.headers.update(HEADERS)
    r = s.get(PAGE_URL, timeout=30, verify=VERIFY_SSL)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")
    meta = soup.find("meta", attrs={"name": "__RequestVerificationToken"})
    if not meta or not meta.get("content"):
        raise RuntimeError("Không tìm thấy __RequestVerificationToken")
    return s, meta["content"]


def build_body(page):
    parts = [("searchKeys[]", "") for _ in range(4)]
    for i in range(N_TABS):
        parts.append(("arrCurrentPage[]", str(page) if i == REPURCHASE_TAB else "1"))
    for i in range(N_TABS):
        parts.append(("arrNumberRecord[]", str(PAGE_SIZE) if i == REPURCHASE_TAB else "1"))
    return parts


def to_number(s):
    if not s:
        return None
    s = s.replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def parse_table(html):
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table", id="tbRepurchaseResult")
    if not table or not table.find("tbody"):
        return []
    rows = []
    for tr in table.find("tbody").find_all("tr", recursive=False):
        tds = tr.find_all("td", recursive=False)
        if len(tds) < len(COLUMNS):
            continue
        rec = {k: td.get_text(strip=True) for k, td in zip(COLUMNS, tds)}
        m = re.search(r"ViewFileTTPH\(([\d.]+)", str(tds[-1]))
        rec["file_id"] = m.group(1) if m else ""
        rec.pop("van_ban", None)
        rows.append(rec)
    return rows


def scrape():
    session, token = get_session_token()
    ph = {"CP-TOKEN": token,
          "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"}
    all_rows, page = [], 1
    while True:
        for attempt in range(3):
            try:
                r = session.post(SEARCH_URL, data=build_body(page), headers=ph,
                                 timeout=45, verify=VERIFY_SSL)
                r.raise_for_status()
                break
            except Exception as e:
                print(f"  ! lỗi trang {page} (lần {attempt+1}): {e}", file=sys.stderr)
                time.sleep(2)
        else:
            page += 1
            continue
        if "application/json" in r.headers.get("Content-Type", "") or r.text.lstrip().startswith("{"):
            session, token = get_session_token()
            ph["CP-TOKEN"] = token
            continue
        rows = parse_table(r.text)
        if not rows:
            break
        all_rows.extend(rows)
        print(f"Trang {page}: +{len(rows)} (tổng {len(all_rows)})")
        if len(rows) < PAGE_SIZE:
            break
        page += 1
        time.sleep(DELAY)

    # khử trùng đợt mua lại: mã TP + ngày thực hiện mua lại + giá trị mua lại (+ngày đăng)
    seen, uniq = set(), []
    for r in all_rows:
        k = (r.get("ma_tp"), r.get("ngay_mua_lai"), r.get("gt_mua_lai"), r.get("ngay_dang_tin"))
        if k in seen:
            continue
        seen.add(k)
        for c in NUM_COLS:
            r[c + "_num"] = to_number(r.get(c))
        uniq.append(r)
    return uniq


def save(rows):
    if not rows:
        print("Không có dữ liệu.")
        return
    cols = [c for c in COLUMNS if c != "van_ban"] + ["file_id"] + [c + "_num" for c in NUM_COLS]
    with open("bond_buyback_raw.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
    with open("bond_buyback_raw.json", "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False)
    print(f"Đã lưu {len(rows)} đợt mua lại -> bond_buyback_raw.csv / .json")


if __name__ == "__main__":
    t0 = time.time()
    print(f"Bắt đầu crawl mua lại lúc {datetime.now():%Y-%m-%d %H:%M:%S}")
    rows = scrape()
    save(rows)
    print(f"Xong sau {time.time()-t0:.0f}s. Số đợt mua lại: {len(rows)}")
