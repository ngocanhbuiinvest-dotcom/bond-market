# -*- coding: utf-8 -*-
"""
Scraper DANH SÁCH TRÁI PHIẾU (cấp MÃ) - CBIS / HNX
Nguồn: https://cbonds.hnx.vn/to-chuc-phat-hanh/danh-sach-trai-phieu

KHÁC với `bond_issuance_scraper.py` (cấp ĐỢT CÔNG BỐ, dữ liệu chỉ từ 30/12/2020):
đây là danh bạ CẤP MÃ của TOÀN BỘ trái phiếu riêng lẻ HNX đang quản lý (~6.8k mã, có cả
mã phát hành từ 2019 trở về trước), và quan trọng nhất là có cột **KL còn lưu hành** do
CHÍNH HNX công bố -> dùng để đối chiếu chéo với dư nợ mà dự án tự tính
(dư nợ = phát hành − mua lại − đáo hạn).

Cơ chế khác các scraper cũ: POST **JSON** (không phải form-urlencoded)
  body {"SearchKeys":[bondCode,comId,releaseDate,depositId,currency,status,orgName,rankNote],
        "CurrentPage":n,"NumberRecordOnPage":100}
  header CP-TOKEN = meta __RequestVerificationToken (cùng quy ước với các trang khác).
⚠ Server CHẶN NumberRecordOnPage > 100 (truyền 500 vẫn trả 100) -> phải phân trang đủ.
VERIFY_SSL=False (cert site thiếu intermediate CA).

Xuất: bond_list_raw.csv / .json
"""
import csv
import json
import re
import sys
import time
from datetime import datetime

import requests
import urllib3
from bs4 import BeautifulSoup

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
VERIFY_SSL = False

BASE = "https://cbonds.hnx.vn"
PAGE_URL = BASE + "/to-chuc-phat-hanh/danh-sach-trai-phieu"
PAGE_SIZE = 100          # trần thực tế của server
DELAY = 0.25

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
    "X-Requested-With": "XMLHttpRequest",
    "Origin": BASE,
    "Referer": PAGE_URL,
}

COLUMNS = ["stt", "ma_tp", "tien_te", "ten_dn", "menh_gia", "ky_han", "ky_han_con_lai",
           "ngay_phat_hanh", "ngay_dao_han", "pt_tra_lai", "ky_han_tra_lai",
           "kl_phat_hanh", "kl_con_luu_hanh", "to_chuc_luu_ky", "lai_suat",
           "doi_tuong_chao_ban", "dv_xhtn", "kq_xhtn", "ngay_hieu_luc",
           "tai_lieu", "tc_ben_vung", "dv_danh_gia", "tinh_trang"]

NUM_COLS = ["menh_gia", "kl_phat_hanh", "kl_con_luu_hanh", "lai_suat"]


def get_session_token():
    s = requests.Session()
    s.headers.update(HEADERS)
    r = s.get(PAGE_URL, timeout=30, verify=VERIFY_SSL)
    r.raise_for_status()
    meta = BeautifulSoup(r.text, "html.parser").find(
        "meta", attrs={"name": "__RequestVerificationToken"})
    if not meta or not meta.get("content"):
        raise RuntimeError("Không tìm thấy __RequestVerificationToken")
    return s, meta["content"]


def parse_total(html):
    m = re.search(r"Tổng số\s*<b>([\d.,]+)</b>", html)
    return int(re.sub(r"[.,]", "", m.group(1))) if m else None


def parse_table(html):
    """Trả về list dòng; None nếu không thấy bảng (phiên hết hạn -> caller làm mới token)."""
    soup = BeautifulSoup(html, "html.parser")
    tb = soup.find("table")
    if not tb or not tb.find("tbody"):
        return None
    rows = []
    for tr in tb.find("tbody").find_all("tr", recursive=False):
        tds = [td.get_text(strip=True) for td in tr.find_all("td", recursive=False)]
        if len(tds) < len(COLUMNS):
            continue
        rows.append(dict(zip(COLUMNS, tds[:len(COLUMNS)])))
    return rows


def to_num(s):
    s = (s or "").replace(",", "").strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def scrape():
    session, token = get_session_token()
    hdr = {"CP-TOKEN": token, "Content-Type": "application/json;charset=utf-8"}

    def post(page):
        body = json.dumps({"SearchKeys": [""] * 8, "CurrentPage": page,
                           "NumberRecordOnPage": PAGE_SIZE})
        return session.post(PAGE_URL, data=body, headers=hdr, timeout=90, verify=VERIFY_SSL)

    r = post(1)
    total = parse_total(r.text)
    npages = (total + PAGE_SIZE - 1) // PAGE_SIZE if total else None
    print(f"Tổng số trái phiếu: {total} -> {npages} trang")
    all_rows = parse_table(r.text) or []
    print(f"trang 1: {len(all_rows)}")

    page = 2
    while npages is None or page <= npages:
        rows = None
        for attempt in range(3):
            try:
                rr = post(page)
                rr.raise_for_status()
                rows = parse_table(rr.text)
                if rows is None:                 # phiên hết hạn -> làm mới token
                    session, tk = get_session_token()
                    hdr["CP-TOKEN"] = tk
                    continue
                break
            except Exception as e:
                print(f"  ! lỗi trang {page} (lần {attempt+1}): {e}", file=sys.stderr)
                time.sleep(2)
        if not rows:
            break
        all_rows.extend(rows)
        if page % 10 == 0 or page == npages:
            print(f"trang {page}: luỹ kế {len(all_rows)}")
        if len(rows) < PAGE_SIZE:
            break
        page += 1
        time.sleep(DELAY)

    for x in all_rows:
        for c in NUM_COLS:
            x[c + "_num"] = to_num(x[c])
    return all_rows, total


def save(rows, csv_path="bond_list_raw.csv", json_path="bond_list_raw.json"):
    if not rows:
        print("Không có dữ liệu để lưu.")
        return
    cols = COLUMNS + [c + "_num" for c in NUM_COLS]
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False)
    print(f"Đã lưu {len(rows)} mã -> {csv_path} / {json_path}")


if __name__ == "__main__":
    t0 = time.time()
    rows, total = scrape()
    save(rows)
    n_dup = len(rows) - len({r["ma_tp"] for r in rows})
    print(f"Xong sau {time.time()-t0:.0f}s. Thu được {len(rows)}/{total} mã"
          f"{f' (trùng mã: {n_dup})' if n_dup else ''}.")
