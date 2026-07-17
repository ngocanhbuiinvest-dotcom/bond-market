# -*- coding: utf-8 -*-
"""
Scraper TRÁI PHIẾU DOANH NGHIỆP RIÊNG LẺ - VSDC (Tổng công ty Lưu ký và Bù trừ CK Việt Nam)
Nguồn: https://www.vsd.vn/vi/ibl  ("Công cụ nợ và trái phiếu doanh nghiệp")

Khác biệt bản chất so với HNX (cbonds.hnx.vn):
  - VSD = ĐĂNG KÝ LƯU KÝ (số CK đã đăng ký tại VSDC, Giấy CNĐKCK).
  - HNX = ĐĂNG KÝ GIAO DỊCH + CBTT phát hành/mua lại.
  => Hai universe gần nhau nhưng KHÔNG trùng; đây chính là mục đích đối chiếu chéo.

Cơ chế crawl (giống pattern CP-TOKEN của HNX, nhưng tên header khác):
  GET /vi/ibl  -> meta[name="__VPToken"]  -> POST /vi/ibl/search (JSON) với header __VPToken.
  Body: {SearchKey, CurrentPage, RecordOnPage, OrderBy, OrderType}; trả về HTML mảnh (bảng).
  SearchKey = 11 trường ngăn '|':
     issuerId|code|stockType|sanGiaoDich|status|lang|kyHan|namPhatHanh|namDaoHan|tenTCPH|typeIsuStock
     stockType cố định "2,4" (theo JS trang); typeIsuStock=7 = TPDN riêng lẻ; lang=VI.
  LƯU Ý: server BỎ QUA RecordOnPage > 10 -> luôn 10 dòng/trang (~234 trang).

Hai giai đoạn:
  1) LIST   : duyệt trang -> mã CK, ISIN, tên TCĐKCK, kỳ hạn, ngày PH/ĐH, số lượng, tình trạng, id chi tiết.
  2) DETAIL : GET /vi/s-detail/<id> cho từng mã (đa luồng) -> MỆNH GIÁ, tổng giá trị ĐK, lãi suất,
              Giấy CNĐKCK, hình thức phát hành, và **mã CBTT** (trong ngoặc ở "Tên chứng khoán",
              vd "Trái phiếu ... (AAAH2124001)") = KHÓA NỐI sang dữ liệu phát hành HNX.
  Chạy --no-detail để bỏ giai đoạn 2 (nhanh, nhưng không có mệnh giá/giá trị).

Xuất: vsd_bond_raw.csv / .json
"""
import argparse
import csv
import json
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

import requests
import urllib3
from bs4 import BeautifulSoup

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
VERIFY_SSL = True

BASE = "https://www.vsd.vn"
PAGE_URL = BASE + "/vi/ibl"
SEARCH_URL = BASE + "/vi/ibl/search"
DETAIL_URL = BASE + "/vi/s-detail/{}"
PAGE_SIZE = 10          # server chốt cứng 10
DELAY = 0.15
DETAIL_WORKERS = 8
TYPE_TPDN_RIENG_LE = "7"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
    "X-Requested-With": "XMLHttpRequest",
    "Origin": BASE,
    "Referer": PAGE_URL,
}

LIST_COLUMNS = ["stt", "ma_ck", "isin", "ten_tcdkck", "ky_han",
                "ngay_phat_hanh", "ngay_dao_han", "so_luong", "tinh_trang"]
DETAIL_COLUMNS = ["ma_cbtt", "ten_ck", "menh_gia", "tong_so_dk", "gia_tri_dk",
                  "lai_suat", "hinh_thuc_ph", "noi_gd", "giay_cnck", "cach_tra_lai"]
COLUMNS = LIST_COLUMNS + DETAIL_COLUMNS + ["detail_id", "issuer_id"]


def get_session():
    s = requests.Session()
    s.headers.update(HEADERS)
    r = s.get(PAGE_URL, timeout=60, verify=VERIFY_SSL)
    r.raise_for_status()
    m = re.search(r'name="__VPToken"\s+content="([^"]+)"', r.text)
    if not m:
        raise RuntimeError("Không tìm thấy meta __VPToken")
    s.headers["__VPToken"] = m.group(1)
    return s


def build_key(type_isu=TYPE_TPDN_RIENG_LE):
    # issuerId|code|stockType|sanGiaoDich|status|lang|kyHan|namPhatHanh|namDaoHan|tenTCPH|typeIsuStock
    return "|".join(["0", "", "2,4", "", "", "VI", "", "", "", "", type_isu])


def parse_total(html):
    m = re.search(r"/\s*([\d.,]+)\s*bản ghi", html)
    return int(m.group(1).replace(".", "").replace(",", "")) if m else None


def parse_summary(html):
    """Dòng tổng kết cuối bảng của VSD (dùng để đối chiếu số ta tự cộng)."""
    txt = BeautifulSoup(html, "html.parser").get_text(" ", strip=True)
    out = {}
    m = re.search(r"Tổng\s+([\d.,]+)\s+mã chứng khoán\s*\(([\d.,]+)\s+mã chứng khoán đăng ký,"
                  r"\s*([\d.,]+)\s+mã chứng khoán hủy đăng ký\)", txt)
    if m:
        out["tong_ma"], out["ma_hieu_luc"], out["ma_huy"] = [
            int(x.replace(".", "").replace(",", "")) for x in m.groups()]
    m = re.search(r"Tổng số lượng chứng khoán đang lưu hành:\s*([\d.,]+)", txt)
    if m:
        out["sl_luu_hanh"] = int(m.group(1).replace(".", "").replace(",", ""))
    return out


def parse_list_table(html):
    soup = BeautifulSoup(html, "html.parser")
    t = soup.find("table", id="tblIsuStock")
    if not t:
        return None
    body = t.find("tbody")
    if not body:
        return []
    rows = []
    for tr in body.find_all("tr", recursive=False):
        tds = tr.find_all("td", recursive=False)
        if len(tds) < 9:
            continue
        vals = [td.get_text(strip=True) for td in tds[:9]]
        row = dict(zip(LIST_COLUMNS, vals))
        # id chi tiết ở href /s-detail/<id>; id TCPH ở href /id/<id>
        a = tds[1].find("a", href=True)
        m = re.search(r"/s-detail/(\d+)", a["href"]) if a else None
        row["detail_id"] = m.group(1) if m else ""
        a2 = tds[3].find("a", href=True)
        m2 = re.search(r"/id/(\d+)", a2["href"]) if a2 else None
        row["issuer_id"] = m2.group(1) if m2 else ""
        rows.append(row)
    return rows


def scrape_list(session):
    body = {"SearchKey": build_key(), "CurrentPage": 1,
            "RecordOnPage": PAGE_SIZE, "OrderBy": "", "OrderType": ""}
    r = session.post(SEARCH_URL, json=body, timeout=60, verify=VERIFY_SSL)
    r.raise_for_status()
    total = parse_total(r.text)
    summary = parse_summary(r.text)
    npages = (total + PAGE_SIZE - 1) // PAGE_SIZE if total else None
    print(f"Tổng số mã TPDN riêng lẻ (VSD): {total} -> {npages} trang")
    if summary:
        print(f"  VSD tự tổng kết: {summary}")
    rows = parse_list_table(r.text) or []

    page = 2
    while npages is None or page <= npages:
        got = None
        for attempt in range(3):
            try:
                body["CurrentPage"] = page
                rr = session.post(SEARCH_URL, json=body, timeout=60, verify=VERIFY_SSL)
                if rr.status_code in (400, 403):      # token hết hạn
                    session = get_session()
                    continue
                rr.raise_for_status()
                got = parse_list_table(rr.text)
                if got is None:
                    session = get_session()
                    continue
                break
            except Exception as e:
                print(f"  ! lỗi trang {page} (lần {attempt+1}): {e}", file=sys.stderr)
                time.sleep(2)
        if not got:
            break
        rows.extend(got)
        if page % 20 == 0 or page == npages:
            print(f"  trang {page}/{npages}: luỹ kế {len(rows)}")
        page += 1
        time.sleep(DELAY)
    return rows, summary


# ---------------------------------------------------------------- chi tiết
_NUM = re.compile(r"[\d.]+")


def _num(s):
    """'1.000.000.000 đồng' -> 1000000000 (VSD dùng '.' ngăn nghìn)."""
    if not s:
        return ""
    m = _NUM.search(s.replace(",", "."))
    if not m:
        return ""
    v = m.group(0).replace(".", "")
    return int(v) if v.isdigit() else ""


_CODE = r"[A-Z][A-Z0-9_.\-]{5,}"


def extract_ma_cbtt(ten_ck, ma_ck=""):
    """Bóc mã CBTT từ "Tên chứng khoán" của VSD. Tên có 4 dạng thực tế:
         1. 'Trái phiếu <tên DN> (BIDLH2331010)'  -> mã CBTT trong ngoặc cuối
         2. 'Trái phiếu VHMB2426004'              -> mã CBTT đứng trần
         3. 'Trái phiếu STA12601'                 -> chỉ là MÃ GD (== ma_ck) => KHÔNG phải mã CBTT
         4. 'Trái phiếu Ngân hàng TMCP ...'       -> không có mã
    """
    ten_ck = (ten_ck or "").strip()
    ma_ck = (ma_ck or "").strip().upper()
    m = re.search(r"\((%s)\)\s*$" % _CODE, ten_ck)
    if m and m.group(1).upper() != ma_ck:
        return m.group(1)
    m = re.match(r"^Trái phiếu\s+(%s)\s*$" % _CODE, ten_ck)
    if m and m.group(1).upper() != ma_ck:      # dạng 3: trùng mã GD -> bỏ
        return m.group(1)
    return ""


def parse_detail(html):
    soup = BeautifulSoup(html, "html.parser")
    pane = soup.find(id="Detail_TCPH_TTCK") or soup
    kv = {}
    for row in pane.find_all("div", class_="row"):
        cols = row.find_all("div", recursive=False)
        if len(cols) < 2:
            continue
        label = cols[0].get_text(" ", strip=True).rstrip(":").strip()
        value = cols[1].get_text(" ", strip=True)
        if label and label not in kv:
            kv[label] = value
    ten_ck = kv.get("Tên chứng khoán", "")
    return {
        "ma_cbtt": extract_ma_cbtt(ten_ck, kv.get("Mã chứng khoán", "")),
        "ten_ck": ten_ck,
        "menh_gia": _num(kv.get("Mệnh giá", "")),
        "tong_so_dk": _num(kv.get("Tổng số chứng khoán đăng ký", "")),
        "gia_tri_dk": _num(kv.get("Tổng giá trị chứng khoán đăng ký", "")),
        "lai_suat": kv.get("Lãi suất", ""),
        "hinh_thuc_ph": kv.get("Hình thức phát hành", ""),
        "noi_gd": kv.get("Nơi giao dịch (*)", "") or kv.get("Nơi giao dịch", ""),
        "giay_cnck": re.sub(r"\s+", " ", kv.get("Giấy chứng nhận ĐKCK", "")),
        "cach_tra_lai": kv.get("Cách thức trả lãi", ""),
    }


def fetch_detail(args):
    row, i, n = args
    did = row.get("detail_id")
    if not did:
        return row
    s = requests.Session()
    s.headers.update(HEADERS)
    for attempt in range(3):
        try:
            r = s.get(DETAIL_URL.format(did), timeout=60, verify=VERIFY_SSL)
            r.raise_for_status()
            row.update(parse_detail(r.text))
            return row
        except Exception as e:
            if attempt == 2:
                print(f"  ! chi tiết lỗi {row.get('ma_ck')} ({did}): {e}", file=sys.stderr)
            time.sleep(1.5)
    return row


def scrape_details(rows):
    print(f"Lấy chi tiết {len(rows)} mã ({DETAIL_WORKERS} luồng)...")
    t0 = time.time()
    out = []
    with ThreadPoolExecutor(max_workers=DETAIL_WORKERS) as ex:
        for i, row in enumerate(ex.map(fetch_detail,
                                       [(r, i, len(rows)) for i, r in enumerate(rows)]), 1):
            out.append(row)
            if i % 200 == 0:
                print(f"  {i}/{len(rows)} ({time.time()-t0:.0f}s)")
    ok = sum(1 for r in out if r.get("menh_gia") != "")
    print(f"  xong {len(out)} mã, có mệnh giá: {ok}, có mã CBTT: "
          f"{sum(1 for r in out if r.get('ma_cbtt'))} ({time.time()-t0:.0f}s)")
    return out


def save(rows, summary, csv_path="vsd_bond_raw.csv", json_path="vsd_bond_raw.json"):
    if not rows:
        print("Không có dữ liệu để lưu.")
        return
    for r in rows:
        for c in COLUMNS:
            r.setdefault(c, "")
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=COLUMNS, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump({"summary": summary, "rows": rows}, f, ensure_ascii=False)
    print(f"Đã lưu {len(rows)} bản ghi -> {csv_path} / {json_path}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-detail", action="store_true",
                    help="chỉ lấy danh sách, bỏ qua trang chi tiết (không có mệnh giá/giá trị)")
    a = ap.parse_args()

    t0 = time.time()
    print(f"Bắt đầu crawl VSD TPDN riêng lẻ lúc {datetime.now():%Y-%m-%d %H:%M:%S}")
    sess = get_session()
    rows, summary = scrape_list(sess)
    if not a.no_detail:
        rows = scrape_details(rows)
    save(rows, summary)
    print(f"Xong sau {time.time()-t0:.0f}s. Tổng: {len(rows)} mã.")
