# -*- coding: utf-8 -*-
"""
Scraper DANH MỤC trái phiếu đăng ký giao dịch (crosswalk) - CBIS / HNX
Nguồn: https://cbonds.hnx.vn/danh-muc-trai-phieu

Mục đích: bảng khóa nối giữa các hệ mã của site HNX. Mỗi dòng cho một trái phiếu ĐKGD với:
  - Mã trái phiếu CBTT      (dùng ở dữ liệu phát hành/mua lại)
  - Mã trái phiếu giao dịch (dùng ở dữ liệu GIAO DỊCH thứ cấp - bond_secondary_raw.csv)
  - Mã ISIN
  - Tên tổ chức phát hành   (đầy đủ)
  - mệnh giá, KL ĐKGD, trạng thái ĐKGD, ngày GD đầu/cuối, đối tượng GD

=> Dùng để MAP mã giao dịch thứ cấp về đúng TCPH/ngành (tab "Giao dịch thứ cấp" trước đây
   không map được vì mã giao dịch ≠ mã phát hành).

Cơ chế giống các scraper khác: GET lấy __RequestVerificationToken -> POST /danh-sach
  body: keySearch = 6 trường ngăn '|' (để trống = tất cả) + arrCurrentPage[] + arrNumberRecord[].
VERIFY_SSL=False (cert site thiếu intermediate CA).

Xuất: bond_catalog_raw.csv / .json
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
PAGE_URL = BASE + "/danh-muc-trai-phieu"
SEARCH_URL = PAGE_URL + "/danh-sach"
PAGE_SIZE = 100
DELAY = 0.3

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
    "X-Requested-With": "XMLHttpRequest",
    "Origin": BASE,
    "Referer": PAGE_URL,
}

COLUMNS = ["stt", "ma_cbtt", "ma_gd", "isin", "ten_tcph", "menh_gia",
           "kl_dkgd", "trang_thai", "ngay_gd_dau", "ngay_gd_cuoi", "doi_tuong"]


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
    key = "|".join([""] * 6)
    return [("keySearch", key),
            ("arrCurrentPage[]", str(page)),
            ("arrNumberRecord[]", str(PAGE_SIZE))]


def parse_total(html):
    m = re.search(r"Tổng số\s*<b>(\d+)</b>", html)
    return int(m.group(1)) if m else None


def parse_table(html):
    soup = BeautifulSoup(html, "html.parser")
    for t in soup.find_all("table"):
        thead = t.find("thead")
        heads = [th.get_text(strip=True) for th in thead.find_all("th")] if thead else []
        if "Mã trái phiếu giao dịch" in heads and "Tên tổ chức phát hành" in heads:
            body = t.find("tbody")
            if not body:
                return []
            rows = []
            for tr in body.find_all("tr", recursive=False):
                tds = [td.get_text(strip=True) for td in tr.find_all("td", recursive=False)]
                if len(tds) < 11:
                    continue
                rows.append(dict(zip(COLUMNS, tds[:11])))
            return rows
    return None


def scrape():
    session, token = get_session_token()
    post_headers = {"CP-TOKEN": token,
                    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"}
    # trang 1 để lấy tổng số
    r = session.post(SEARCH_URL, data=build_body(1), headers=post_headers,
                     timeout=60, verify=VERIFY_SSL)
    total = parse_total(r.text)
    npages = (total + PAGE_SIZE - 1) // PAGE_SIZE if total else None
    print(f"Tổng số mã ĐKGD: {total} -> {npages} trang")
    all_rows = parse_table(r.text) or []
    print(f"trang 1: {len(all_rows)}")

    page = 2
    while npages is None or page <= npages:
        rows = None
        for attempt in range(3):
            try:
                rr = session.post(SEARCH_URL, data=build_body(page), headers=post_headers,
                                  timeout=60, verify=VERIFY_SSL)
                rr.raise_for_status()
                rows = parse_table(rr.text)
                if rows is None:      # phiên hết hạn -> làm mới token
                    session, tok = get_session_token()
                    post_headers["CP-TOKEN"] = tok
                    continue
                break
            except Exception as e:
                print(f"  ! lỗi trang {page} (lần {attempt+1}): {e}", file=sys.stderr)
                time.sleep(2)
        if not rows:
            break
        all_rows.extend(rows)
        print(f"trang {page}: +{len(rows)} (luỹ kế {len(all_rows)})")
        if len(rows) < PAGE_SIZE:
            break
        page += 1
        time.sleep(DELAY)
    return all_rows


def save(rows, csv_path="bond_catalog_raw.csv", json_path="bond_catalog_raw.json"):
    if not rows:
        print("Không có dữ liệu để lưu.")
        return
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=COLUMNS, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False)
    print(f"Đã lưu {len(rows)} bản ghi -> {csv_path} / {json_path}")


if __name__ == "__main__":
    t0 = time.time()
    print(f"Bắt đầu crawl DANH MỤC ĐKGD lúc {datetime.now():%Y-%m-%d %H:%M:%S}")
    rows = scrape()
    save(rows)
    print(f"Xong sau {time.time()-t0:.0f}s. Tổng: {len(rows)} mã.")
