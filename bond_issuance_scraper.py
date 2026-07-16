# -*- coding: utf-8 -*-
"""
Scraper Kết quả chào bán TPDN riêng lẻ TRONG NƯỚC - CBIS / HNX
Nguồn: https://cbonds.hnx.vn/to-chuc-phat-hanh/thong-tin-phat-hanh  (tab "Kết quả chào bán trong nước")

Cơ chế (đã reverse-engineer):
  1. GET trang phát hành -> lấy meta __RequestVerificationToken + cookie session
  2. POST /to-chuc-phat-hanh/thong-tin-phat-hanh/tim-kiem
       header  CP-TOKEN = token
       body    searchKeys[] = [comId, bondCode, issueFromDate, issueToDate]
               arrCurrentPage[]  (12 phần tử, index 0 = tab chào bán trong nước = số trang)
               arrNumberRecord[] (12 phần tử, index 0 = số bản ghi/trang)
  3. Parse bảng #tbReleaseResult

Xuất: bond_issuance_raw.csv  (dữ liệu thô, dùng lại cho phân tích/dashboard)
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

# Cert chain của cbonds.hnx.vn thiếu intermediate CA -> bỏ xác thực SSL (dữ liệu công khai, chỉ đọc)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
VERIFY_SSL = False

BASE = "https://cbonds.hnx.vn"
PAGE_URL = BASE + "/to-chuc-phat-hanh/thong-tin-phat-hanh"
SEARCH_URL = PAGE_URL + "/tim-kiem"
PAGE_SIZE = 100
N_TABS = 12          # số tab của form (arrCurrentPage/arrNumberRecord có 12 phần tử)
DOMESTIC_TAB = 0     # index tab "Kết quả chào bán trong nước"
DELAY = 0.8          # nghỉ giữa các request cho lịch sự với server

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
    "X-Requested-With": "XMLHttpRequest",
    "Origin": BASE,
    "Referer": PAGE_URL,
}

# thứ tự cột trong #tbReleaseResult (theo thead)
COLUMNS = [
    "stt", "ngay_dang_tin", "ten_dn", "ma_tp", "tien_te", "ky_han",
    "ngay_phat_hanh", "ngay_dao_han", "ky_han_con_lai", "khoi_luong",
    "menh_gia", "loai_hinh_tra_lai", "loai_lai_suat", "pt_thanh_toan_lai",
    "mua_lai_hoan_doi", "thi_truong", "lai_suat", "tinh_trang", "van_ban",
]


def get_session_token():
    s = requests.Session()
    s.headers.update(HEADERS)
    r = s.get(PAGE_URL, timeout=30, verify=VERIFY_SSL)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")
    meta = soup.find("meta", attrs={"name": "__RequestVerificationToken"})
    if not meta or not meta.get("content"):
        raise RuntimeError("Không tìm thấy __RequestVerificationToken")
    return s, meta["content"], soup


def build_body(page):
    parts = []
    for _ in range(4):                       # searchKeys rỗng: comId, bondCode, fromDate, toDate
        parts.append(("searchKeys[]", ""))
    for i in range(N_TABS):
        parts.append(("arrCurrentPage[]", str(page) if i == DOMESTIC_TAB else "1"))
    for i in range(N_TABS):
        parts.append(("arrNumberRecord[]", str(PAGE_SIZE) if i == DOMESTIC_TAB else "10"))
    return parts


def parse_release_table(html):
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table", id="tbReleaseResult")
    if not table:
        return []
    body = table.find("tbody")
    if not body:
        return []
    rows = []
    for tr in body.find_all("tr", recursive=False):
        tds = tr.find_all("td", recursive=False)
        if len(tds) < len(COLUMNS):
            continue
        rec = {}
        for key, td in zip(COLUMNS, tds):
            rec[key] = td.get_text(strip=True)
        # link file đính kèm (ViewFileTTPH(id, type))
        m = re.search(r"ViewFileTTPH\(([\d.]+)", str(tds[-1]))
        rec["file_id"] = m.group(1) if m else ""
        rec.pop("van_ban", None)
        rows.append(rec)
    return rows


def to_number(s):
    if not s:
        return None
    s = s.replace(",", "").replace("%", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def scrape():
    session, token, soup = get_session_token()
    post_headers = {"CP-TOKEN": token,
                    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"}

    # tổng số bản ghi (để biết số trang) — lấy từ trang đầu
    total = None
    for txt in soup.stripped_strings:
        m = re.search(r"Tổng số\s+([\d.,]+)\s+bản ghi", txt)
        if m:
            total = int(m.group(1).replace(".", "").replace(",", ""))
            break

    all_rows = []
    page = 1
    while True:
        body = build_body(page)
        for attempt in range(3):
            try:
                r = session.post(SEARCH_URL, data=body, headers=post_headers,
                                 timeout=45, verify=VERIFY_SSL)
                r.raise_for_status()
                break
            except Exception as e:
                print(f"  ! lỗi trang {page} (lần {attempt+1}): {e}", file=sys.stderr)
                time.sleep(2)
        else:
            print(f"  !! Bỏ qua trang {page} sau 3 lần thử", file=sys.stderr)
            page += 1
            continue

        # server có thể trả JSON khi hết phiên
        ct = r.headers.get("Content-Type", "")
        if "application/json" in ct or r.text.lstrip().startswith("{"):
            # thử làm mới token rồi lặp lại
            session, token, _ = get_session_token()
            post_headers["CP-TOKEN"] = token
            continue

        rows = parse_release_table(r.text)
        if not rows:
            break
        all_rows.extend(rows)
        print(f"Trang {page}: +{len(rows)} bản ghi (tổng {len(all_rows)}"
              + (f"/{total}" if total else "") + ")")

        if total and len(all_rows) >= total:
            break
        if len(rows) < PAGE_SIZE:
            break
        page += 1
        time.sleep(DELAY)

    # khử trùng lặp theo mã TP + ngày phát hành (STT reset mỗi trang)
    seen, unique = set(), []
    for r in all_rows:
        k = (r.get("ma_tp"), r.get("ngay_phat_hanh"), r.get("ngay_dang_tin"))
        if k in seen:
            continue
        seen.add(k)
        r["khoi_luong_num"] = to_number(r.get("khoi_luong"))
        r["menh_gia_num"] = to_number(r.get("menh_gia"))
        r["lai_suat_num"] = to_number(r.get("lai_suat"))
        if r["khoi_luong_num"] is not None and r["menh_gia_num"] is not None:
            r["gia_tri_phat_hanh"] = r["khoi_luong_num"] * r["menh_gia_num"]
        else:
            r["gia_tri_phat_hanh"] = None
        unique.append(r)
    return unique, total


def save_csv(rows, path="bond_issuance_raw.csv"):
    if not rows:
        print("Không có dữ liệu để lưu.")
        return
    cols = [c for c in COLUMNS if c != "van_ban"] + \
           ["file_id", "khoi_luong_num", "menh_gia_num", "lai_suat_num", "gia_tri_phat_hanh"]
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
    print(f"Đã lưu {len(rows)} bản ghi -> {path}")


if __name__ == "__main__":
    t0 = time.time()
    print(f"Bắt đầu crawl lúc {datetime.now():%Y-%m-%d %H:%M:%S}")
    rows, total = scrape()
    save_csv(rows)
    # lưu kèm JSON để dashboard dùng
    with open("bond_issuance_raw.json", "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False)
    print(f"Xong sau {time.time()-t0:.0f}s. Tổng site báo: {total}, thu được: {len(rows)}")
