# -*- coding: utf-8 -*-
"""
Scraper TIN BẤT THƯỜNG (Công bố thông tin bất thường) TPDN riêng lẻ trong nước - CBIS / HNX
Nguồn: https://cbonds.hnx.vn/to-chuc-phat-hanh/tin-cong-bo  (tab "Tin bất thường")

Mục đích chính: dựng DANH SÁCH CHẬM TRẢ gốc/lãi trái phiếu. HNX bắt buộc TCPH công bố thông tin
bất thường khi chậm thanh toán gốc/lãi (NĐ 65/2022). Tin bất thường còn gồm nhiều loại khác (thay
đổi điều khoản, xử phạt thuế, tài sản đảm bảo...) -> ta quét TOÀN BỘ tab này rồi PHÂN LOẠI sự kiện
"chậm trả" ở khâu build_reports (theo từ khóa tiêu đề). Giữ đủ để bức tranh rủi ro không bị bỏ sót.

Cơ chế (đã reverse-engineer):
  1. GET trang -> lấy meta __RequestVerificationToken + cookie session
  2. POST /to-chuc-phat-hanh/tin-cong-bo-x
       header  CP-TOKEN = token
       body    keysSearch[]  = [title, dateFrom, dateTo, bondCode, comId, articleType, loaihinh]
               currentPages[] = [p_dinhky, p_?, p_batthuong, p_?]   (INDEX 2 = tab Tin bất thường)
               numberRecord[] = [n_dinhky, n_?, n_batthuong, n_?]   (INDEX 2 = tab Tin bất thường)
  3. Response = HTML gồm 4 bảng; ta parse #tbInconstant (Tin bất thường).

Cột #tbInconstant: STT · Ngày đăng tin · Tên doanh nghiệp · Mã TP liên quan · Tiêu đề tin ·
                   Ghi chú · Tình trạng · File đính kèm
loaihinh = '0' (trong nước). Tiêu đề kèm article_id qua showArticle('<id>'); file qua ViewFile.

Xuất: bond_latepay_raw.csv / .json  (toàn bộ Tin bất thường; cột is_late/is_cured phân loại sẵn)
"""
import csv
import re
import sys
import time
import json
import unicodedata
from datetime import datetime

import requests
import urllib3
from bs4 import BeautifulSoup

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
VERIFY_SSL = False

BASE = "https://cbonds.hnx.vn"
PAGE_URL = BASE + "/to-chuc-phat-hanh/tin-cong-bo"
SEARCH_URL = BASE + "/to-chuc-phat-hanh/tin-cong-bo-x"
INCONSTANT_IDX = 2   # index tab "Tin bất thường" trong currentPages[]/numberRecord[]
PAGE_SIZE = 100
DELAY = 0.5

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
    "X-Requested-With": "XMLHttpRequest",
    "Origin": BASE,
    "Referer": PAGE_URL,
}

COLUMNS = ["article_id", "ngay_dang_tin", "ten_dn", "ma_tp", "tieu_de",
           "ghi_chu", "tinh_trang", "file_id", "loai_su_kien"]


def strip_accent(s):
    return "".join(c for c in unicodedata.normalize("NFD", s or "")
                   if unicodedata.category(c) != "Mn").lower()


def classify_event(title):
    """Phân loại sự kiện chậm trả từ tiêu đề CBTT bất thường.
       -> 'cham_tra'      : công bố chậm thanh toán gốc/lãi (sự kiện vỡ/chậm)
          'khac_phuc'     : đã thanh toán sau thời gian bị chậm (khắc phục)
          ''              : không phải sự kiện chậm trả."""
    t = strip_accent(title)
    late = any(k in t for k in ["cham thanh toan", "cham tra", "khong the thanh toan",
                                "khong thanh toan duoc"])
    if not late:
        return ""
    cured = any(k in t for k in ["sau thoi gian bi cham", "sau khi cham",
                                 "sau thoi gian cham", "sau cham thanh toan"])
    return "khac_phuc" if cured else "cham_tra"


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
    # keysSearch: [title, dateFrom, dateTo, bondCode, comId, articleType, loaihinh]
    keys = ["", "", "", "", "", "", "0"]
    pages = ["1", "1", "1", "1"]
    recs = ["10", "10", "10", "10"]
    pages[INCONSTANT_IDX] = str(page)
    recs[INCONSTANT_IDX] = str(PAGE_SIZE)
    body = [("keysSearch[]", k) for k in keys]
    body += [("currentPages[]", p) for p in pages]
    body += [("numberRecord[]", n) for n in recs]
    return body


def parse_inconstant(html):
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table", id="tbInconstant")
    if not table:
        return None
    body = table.find("tbody")
    if not body:
        return []
    rows = []
    for tr in body.find_all("tr", recursive=False):
        tds = tr.find_all("td", recursive=False)
        if len(tds) < 7:
            continue
        title = tds[4].get_text(strip=True)
        rec = {
            "ngay_dang_tin": tds[1].get_text(strip=True),
            "ten_dn": tds[2].get_text(strip=True),
            "ma_tp": tds[3].get_text(strip=True),
            "tieu_de": title,
            "ghi_chu": tds[5].get_text(strip=True),
            "tinh_trang": tds[6].get_text(strip=True),
        }
        m = re.search(r"showArticle\('([^']+)'\)", str(tds[4]))
        rec["article_id"] = m.group(1) if m else ""
        fid = ""
        if len(tds) > 7:
            mf = re.search(r"ViewFile\('([^']+)'", str(tds[7]))
            fid = mf.group(1) if mf else ""
        rec["file_id"] = fid
        rec["loai_su_kien"] = classify_event(title)
        rows.append(rec)
    return rows


def scrape():
    session, token = get_session_token()
    post_headers = {"CP-TOKEN": token,
                    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"}
    all_rows, seen = [], set()
    page = 1
    while True:
        rows = None
        for attempt in range(3):
            try:
                r = session.post(SEARCH_URL, data=build_body(page), headers=post_headers,
                                 timeout=60, verify=VERIFY_SSL)
                r.raise_for_status()
                ct = r.headers.get("Content-Type", "")
                if "application/json" in ct or r.text.lstrip().startswith("{"):
                    session, tok = get_session_token()   # phiên hết hạn
                    post_headers["CP-TOKEN"] = tok
                    continue
                rows = parse_inconstant(r.text)
                if rows is None:
                    session, tok = get_session_token()
                    post_headers["CP-TOKEN"] = tok
                    continue
                break
            except Exception as e:
                print(f"  ! lỗi trang {page} (lần {attempt+1}): {e}", file=sys.stderr)
                time.sleep(2)
        if not rows:
            break
        new = [x for x in rows if x["article_id"] not in seen]
        for x in new:
            seen.add(x["article_id"])
        all_rows.extend(new)
        n_late = sum(1 for x in all_rows if x["loai_su_kien"])
        print(f"trang {page}: +{len(rows)} (mới {len(new)}, luỹ kế {len(all_rows)}, chậm trả {n_late})")
        if len(rows) < PAGE_SIZE or not new:
            break
        page += 1
        time.sleep(DELAY)
    return all_rows


def save(rows, csv_path="bond_latepay_raw.csv", json_path="bond_latepay_raw.json"):
    if not rows:
        print("Không có dữ liệu để lưu.")
        return
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=COLUMNS, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False)
    n_late = sum(1 for x in rows if x["loai_su_kien"] == "cham_tra")
    n_cured = sum(1 for x in rows if x["loai_su_kien"] == "khac_phuc")
    print(f"Đã lưu {len(rows)} tin bất thường -> {csv_path} / {json_path}")
    print(f"  Trong đó: {n_late} CBTT chậm trả · {n_cured} CBTT khắc phục (thanh toán sau chậm)")


if __name__ == "__main__":
    t0 = time.time()
    print(f"Bắt đầu crawl TIN BẤT THƯỜNG lúc {datetime.now():%Y-%m-%d %H:%M:%S}")
    rows = scrape()
    save(rows)
    print(f"Xong sau {time.time()-t0:.0f}s. Tổng: {len(rows)} tin.")
