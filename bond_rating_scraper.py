# -*- coding: utf-8 -*-
"""
Scraper THÔNG TIN XẾP HẠNG TÍN NHIỆM (XHTN) TPDN - CBIS / HNX
Nguồn: https://cbonds.hnx.vn/danh-sach-thong-tin-xep-hang-tin-nhiem

Cơ chế (đã reverse-engineer, giống các scraper khác):
  1. GET trang -> lấy meta __RequestVerificationToken + cookie session
  2. POST /danh-sach-thong-tin-xep-hang-tin-nhiem/danh-sach
       header  CP-TOKEN = token
       body    keySearch = 'objecttype|issuer|bondcode|tradingcode|TrustRating' (để trống = '||||')
               arrCurrentPage[]  = số trang
               arrNumberRecord[] = số bản ghi/trang (tối đa 100)
  3. Parse bảng #tbOrgDeposit (9 cột)

Cột bảng: STT · Đơn vị XHTN · Đối tượng XHTN · Mã trái phiếu · TCPH · Kết quả XHTN gần nhất ·
          Hiệu lực từ ngày · Loại xếp hạng (Dài/Ngắn hạn) · File đính kèm
TCPH lấy kèm comId qua view_TT_TCPH('<comId>'); file đính kèm qua ViewFile('<refId>','800').

Xuất: bond_rating_raw.csv / .json
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
PAGE_URL = BASE + "/danh-sach-thong-tin-xep-hang-tin-nhiem"
SEARCH_URL = PAGE_URL + "/danh-sach"
PAGE_SIZE = 100
DELAY = 0.4

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
    "X-Requested-With": "XMLHttpRequest",
    "Origin": BASE,
    "Referer": PAGE_URL,
}

COLUMNS = ["stt", "don_vi_xhtn", "doi_tuong_xhtn", "ma_tp", "com_id", "ten_tcph",
           "ket_qua_xhtn", "hieu_luc_tu_ngay", "loai_xep_hang", "file_id"]


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
    # keySearch = objecttype|issuer|bondcode|tradingcode|TrustRating (để trống = tất cả)
    return [("keySearch", "||||"),
            ("arrCurrentPage[]", str(page)),
            ("arrNumberRecord[]", str(PAGE_SIZE))]


def parse_total(html):
    m = re.search(r"Tổng số\s*<b>(\d+)</b>", html) or re.search(r"Tổng số\s+([\d.,]+)\s+bản ghi", html)
    return int(m.group(1).replace(".", "").replace(",", "")) if m else None


def parse_table(html):
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table", id="tbOrgDeposit")
    if not table:
        return None
    body = table.find("tbody")
    if not body:
        return []
    rows = []
    for tr in body.find_all("tr", recursive=False):
        tds = tr.find_all("td", recursive=False)
        if len(tds) < 8:
            continue
        rec = {
            "stt": tds[0].get_text(strip=True),
            "don_vi_xhtn": tds[1].get_text(strip=True),
            "doi_tuong_xhtn": tds[2].get_text(strip=True),
            "ma_tp": tds[3].get_text(strip=True),
            "ten_tcph": tds[4].get_text(strip=True),
            "ket_qua_xhtn": tds[5].get_text(strip=True),
            "hieu_luc_tu_ngay": tds[6].get_text(strip=True),
            "loai_xep_hang": tds[7].get_text(strip=True),
        }
        # comId của TCPH: view_TT_TCPH('<comId>')
        m = re.search(r"view_TT_TCPH\('([^']*)'\)", str(tds[4]))
        rec["com_id"] = m.group(1) if m else ""
        # file đính kèm: ViewFile('<refId>', '800')
        fid = ""
        if len(tds) > 8:
            mf = re.search(r"ViewFile\('([^']+)'", str(tds[8]))
            fid = mf.group(1) if mf else ""
        rec["file_id"] = fid
        rows.append(rec)
    return rows


def scrape():
    session, token = get_session_token()
    post_headers = {"CP-TOKEN": token,
                    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"}
    r = session.post(SEARCH_URL, data=build_body(1), headers=post_headers,
                     timeout=60, verify=VERIFY_SSL)
    total = parse_total(r.text)
    npages = (total + PAGE_SIZE - 1) // PAGE_SIZE if total else None
    print(f"Tổng số kết quả XHTN: {total} -> {npages} trang")
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


def save(rows, csv_path="bond_rating_raw.csv", json_path="bond_rating_raw.json"):
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
    print(f"Bắt đầu crawl XẾP HẠNG TÍN NHIỆM lúc {datetime.now():%Y-%m-%d %H:%M:%S}")
    rows = scrape()
    save(rows)
    print(f"Xong sau {time.time()-t0:.0f}s. Tổng: {len(rows)} kết quả XHTN.")
