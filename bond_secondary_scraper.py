# -*- coding: utf-8 -*-
"""
Scraper Giao dịch TPDN riêng lẻ THỨ CẤP (thị trường giao dịch) - CBIS / HNX
Nguồn: https://cbonds.hnx.vn/thong-ke-thi-truong  (tab "Thống kê giao dịch theo Mã trái phiếu")

Cơ chế (đã reverse-engineer):
  1. GET /thong-ke-thi-truong -> lấy meta __RequestVerificationToken + cookie session
  2. POST /thong-ke-thi-truong/danh-sach
       header  CP-TOKEN = token
       body    keySearch = 10 trường ngăn bởi '|':
                 txtFromDate|txtToDate|TK_Bond|slS_Period_Unit|slS_Period|txtS_Year|
                 txtFromDate_TT|txtToDate_TT|txtFromDate_bond|txtToDate_bond
               (tab "theo Mã TP" dùng 2 trường cuối = ngày GD từ/đến)
               arrCurrentPage[]  = [p0, p1, p2]  -> index 1 = trang tab "theo Mã TP"
               arrNumberRecord[] = [n0, n1, n2]  -> index 1 = số bản ghi/trang (tối đa 100)
  3. Parse bảng có thead chứa "Ngày giao dịch" + "Mã giao dịch"

QUAN TRỌNG - bài học (FIX 15/07/2026):
  - Bộ lọc ngày CHỈ hoạt động khi from == to (một ngày cụ thể). Nếu truyền KHOẢNG (from != to)
    server trả snapshot ngày mới nhất -> scraper cũ (cửa sổ tháng) gộp về đúng 1 ngày. => Quét THEO
    TỪNG NGÀY.
  - Mỗi ngày server liệt kê TOÀN BỘ mã đang niêm yết (kl=0 nếu không khớp lệnh), SẮP XẾP theo giá trị
    GIẢM DẦN. Mã có giao dịch (kl>0) nằm ở đầu -> chỉ cần phân trang tới khi gặp dòng giá trị 0 thì dừng.
  - Ngày không có total (=0) = cuối tuần / nghỉ lễ -> bỏ qua.
  - Data có sẵn từ 19/07/2023 (ngày mở thị trường GD TPDN riêng lẻ HNX).
  - VERIFY_SSL=False (cert site thiếu intermediate CA).

Xuất: bond_secondary_raw.csv / .json  (một dòng = một mã×ngày CÓ giao dịch)
"""
import csv
import re
import sys
import time
import json
import argparse
from datetime import datetime, date, timedelta

import requests
import urllib3
from bs4 import BeautifulSoup

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
VERIFY_SSL = False

BASE = "https://cbonds.hnx.vn"
PAGE_URL = BASE + "/thong-ke-thi-truong"
SEARCH_URL = PAGE_URL + "/danh-sach"
BOND_TAB = 1          # index tab "theo Mã trái phiếu"
PAGE_SIZE = 100       # server cap 100 bản ghi/trang
DEFAULT_START = "19/07/2023"   # ngày mở thị trường GD TPDN riêng lẻ HNX
DELAY = 0.4           # nghỉ giữa các request
OVERLAP_DAYS = 7      # incremental: quét lại 7 ngày gần nhất (chồng lấn) để bắt phiên dở/đính chính
CSV_PATH = "bond_secondary_raw.csv"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
    "X-Requested-With": "XMLHttpRequest",
    "Origin": BASE,
    "Referer": PAGE_URL,
}

COLUMNS = ["stt", "ngay_gd", "ma_tp", "khoi_luong", "gia_tri", "gia_cuoi"]


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


def build_key(day_ddmmyyyy):
    # txtFromDate|txtToDate|TK_Bond|Period_Unit|Period|Year|FromDate_TT|ToDate_TT|FromDate_bond|ToDate_bond
    return "|".join(["", "", "", "", "", "", "", "", day_ddmmyyyy, day_ddmmyyyy])


def build_body(key, page):
    parts = [("keySearch", key)]
    for i in range(3):
        parts.append(("arrCurrentPage[]", str(page) if i == BOND_TAB else "1"))
    for i in range(3):
        parts.append(("arrNumberRecord[]", str(PAGE_SIZE) if i == BOND_TAB else "10"))
    return parts


def parse_total(html):
    m = re.search(r"Tổng số\s*<b>(\d+)</b>", html)
    return int(m.group(1)) if m else None


def parse_bond_table(html):
    soup = BeautifulSoup(html, "html.parser")
    for t in soup.find_all("table"):
        thead = t.find("thead")
        heads = [th.get_text(strip=True) for th in thead.find_all("th")] if thead else []
        if "Ngày giao dịch" in heads and "Mã giao dịch" in heads:
            body = t.find("tbody")
            if not body:
                return []
            rows = []
            for tr in body.find_all("tr", recursive=False):
                tds = tr.find_all("td", recursive=False)
                if len(tds) < 6:
                    continue
                vals = [td.get_text(strip=True) for td in tds]
                rows.append(dict(zip(COLUMNS, vals[:6])))
            return rows
    return None   # không tìm thấy bảng (phiên lỗi/hết hạn)


def to_int(s):
    if not s:
        return None
    s = s.replace(",", "").replace(".", "").strip()
    try:
        return int(s)
    except ValueError:
        return None


def trading_days(start_ddmmyyyy, end_ddmmyyyy):
    d0 = datetime.strptime(start_ddmmyyyy, "%d/%m/%Y").date()
    d1 = datetime.strptime(end_ddmmyyyy, "%d/%m/%Y").date()
    cur = d0
    while cur <= d1:
        if cur.weekday() < 5:          # bỏ Thứ Bảy (5) / Chủ Nhật (6); ngày lễ lọc bằng total=0
            yield cur.strftime("%d/%m/%Y")
        cur += timedelta(days=1)


def fetch_day(session, post_headers, token_box, day):
    """Trả về danh sách row CÓ giao dịch (kl>0) trong 'day'. Dừng phân trang khi gặp dòng giá trị 0."""
    key = build_key(day)
    out = []
    page = 1
    while True:
        body = build_body(key, page)
        rows = None
        for attempt in range(3):
            try:
                r = session.post(SEARCH_URL, data=body, headers=post_headers,
                                 timeout=60, verify=VERIFY_SSL)
                r.raise_for_status()
                rows = parse_bond_table(r.text)
                if rows is None:      # phiên hết hạn -> làm mới token
                    session, tok = get_session_token()
                    token_box[0] = session
                    post_headers["CP-TOKEN"] = tok
                    continue
                break
            except Exception as e:
                print(f"  ! lỗi {day} trang {page} (lần {attempt+1}): {e}", file=sys.stderr)
                time.sleep(2)
        if not rows:
            break

        hit_zero = False
        for r in rows:
            kl = to_int(r.get("khoi_luong"))
            gt = to_int(r.get("gia_tri"))
            if not kl or not gt:          # sorted desc theo giá trị -> gặp 0 là hết phần có GD
                hit_zero = True
                break
            r["kl_num"] = kl
            r["gt_num"] = gt
            r["gia_cuoi_num"] = to_int(r.get("gia_cuoi"))
            r["gt_ty"] = gt / 1e9
            r.pop("stt", None)
            out.append(r)

        if hit_zero or len(rows) < PAGE_SIZE:
            break
        page += 1
        time.sleep(DELAY)
    return session, out


def scrape(start, end):
    session, token = get_session_token()
    post_headers = {"CP-TOKEN": token,
                    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"}
    token_box = [session]
    all_rows = []
    n_trading = 0
    for day in trading_days(start, end):
        session, rows = fetch_day(session, post_headers, token_box, day)
        post_headers = post_headers  # giữ header (CP-TOKEN có thể đã cập nhật trong fetch_day)
        if rows:
            n_trading += 1
            all_rows.extend(rows)
            gt = sum(r["gt_num"] for r in rows)
            print(f"{day}: {len(rows):3d} mã GD | {gt/1e9:10.1f} tỷ | luỹ kế {len(all_rows)}")
        time.sleep(DELAY)
    print(f"\nSố phiên có giao dịch: {n_trading}")
    return all_rows


def save(rows, csv_path="bond_secondary_raw.csv", json_path="bond_secondary_raw.json"):
    if not rows:
        print("Không có dữ liệu để lưu.")
        return
    cols = ["ngay_gd", "ma_tp", "khoi_luong", "gia_tri", "gia_cuoi",
            "kl_num", "gt_num", "gia_cuoi_num", "gt_ty"]
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False)
    print(f"Đã lưu {len(rows)} bản ghi -> {csv_path} / {json_path}")


def read_existing(path=CSV_PATH):
    """Đọc CSV thứ cấp hiện có -> (rows, max_date). Trả ([], None) nếu chưa có / rỗng."""
    import os
    if not os.path.exists(path):
        return [], None
    rows = list(csv.DictReader(open(path, encoding="utf-8-sig")))
    dates = []
    for r in rows:
        try:
            dates.append(datetime.strptime(r["ngay_gd"], "%d/%m/%Y").date())
        except (ValueError, KeyError):
            pass
    return rows, (max(dates) if dates else None)


def run_incremental(end):
    """GD thứ cấp bất biến với quá khứ -> chỉ quét từ (ngày cuối đã có − OVERLAP_DAYS) tới hôm nay,
       thay thế cửa sổ chồng lấn bằng dữ liệu mới rồi nối phiên mới. Không có CSV -> quét full."""
    existing, maxd = read_existing()
    if maxd is None:
        print("Chưa có dữ liệu cũ -> quét FULL từ", DEFAULT_START)
        return scrape(DEFAULT_START, end)
    start_d = maxd - timedelta(days=OVERLAP_DAYS)
    start = start_d.strftime("%d/%m/%Y")
    print(f"Incremental: dữ liệu cũ tới {maxd:%d/%m/%Y} -> quét lại từ {start} (chồng lấn {OVERLAP_DAYS}n) tới {end}")
    new_rows = scrape(start, end)
    # giữ các phiên CŨ trước cửa sổ chồng lấn; cửa sổ [start..] lấy hoàn toàn từ lần quét mới
    kept = [r for r in existing
            if _row_date(r) is not None and _row_date(r) < start_d]
    merged = kept + new_rows
    print(f"Merge: giữ {len(kept)} bản ghi cũ (< {start}) + {len(new_rows)} bản ghi mới = {len(merged)}")
    return merged


def _row_date(r):
    try:
        return datetime.strptime(r["ngay_gd"], "%d/%m/%Y").date()
    except (ValueError, KeyError):
        return None


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Scrape giao dịch TPDN riêng lẻ thứ cấp theo từng ngày (HNX)")
    ap.add_argument("--start", default=DEFAULT_START, help="dd/mm/yyyy (mặc định 19/07/2023)")
    ap.add_argument("--end", default=datetime.now().strftime("%d/%m/%Y"), help="dd/mm/yyyy (mặc định hôm nay)")
    ap.add_argument("--incremental", action="store_true",
                    help="chỉ quét phiên mới (từ ngày cuối đã có − chồng lấn) rồi merge vào CSV cũ")
    args = ap.parse_args()

    t0 = time.time()
    if args.incremental:
        print(f"Bắt đầu crawl GD thứ cấp INCREMENTAL lúc {datetime.now():%Y-%m-%d %H:%M:%S}")
        rows = run_incremental(args.end)
    else:
        print(f"Bắt đầu crawl GD thứ cấp THEO NGÀY {args.start} -> {args.end} lúc {datetime.now():%Y-%m-%d %H:%M:%S}")
        rows = scrape(args.start, args.end)
    save(rows)
    print(f"Xong sau {time.time()-t0:.0f}s. Tổng bản ghi (mã×ngày có GD): {len(rows)}")
