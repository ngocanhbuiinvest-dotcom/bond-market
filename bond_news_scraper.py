# -*- coding: utf-8 -*-
"""
Scraper TIN KHÁC + TIN TỪ SỞ (Công bố thông tin) TPDN riêng lẻ - CBIS / HNX
Nguồn: https://cbonds.hnx.vn/to-chuc-phat-hanh/tin-cong-bo

Bổ sung cho `bond_latepay_scraper.py` (chỉ quét tab "Tin bất thường"). Khảo sát 22/07/2026
cho thấy trang có 4 tab và ta mới dùng 1:

  index 0 "Tin định kỳ"     6.882 tin - BCTC & báo cáo định kỳ của TCPH (chưa khai thác)
  index 1 "Tin khác"          725 tin - SCRAPER NÀY
  index 2 "Tin bất thường"  2.229 tin - bond_latepay_scraper.py
  index 3 "Tin từ sở"         105 tin - SCRAPER NÀY (bản tin thị trường tháng của HNX)

⚠ BẪY 1 — INDEX ≠ THỨ TỰ BẢNG TRONG DOM. currentPages[]/numberRecord[] đánh index theo
  (định kỳ, tin khác, bất thường, từ sở) = (0,1,2,3) nhưng bảng trong HTML trả về xếp theo
  (tbPeriodic, tbInconstant, tbOthers, tbOthers) -> index 1 nằm ở BẢNG THỨ 3.
  Ánh xạ đúng: {0:0, 2:1, 1:2, 3:3}. Đã kiểm bằng cách đặt numberRecord[i]=25 rồi đếm dòng.
⚠ BẪY 2 — HAI BẢNG CUỐI TRÙNG id="tbOthers" -> KHÔNG được tìm theo id, phải lấy theo
  THỨ TỰ find_all("table").
⚠ BẪY 3 — article_id nằm ở `href="javascript:showArticle('123')"`, KHÔNG phải ở onclick
  (khác với tab Tin bất thường). Lấy nhầm chỗ sẽ ra rỗng toàn bộ mà không báo lỗi.

File đính kèm: Tin khác  -> /view-file?refId=<id>&tableType=3
               Tin từ sở -> /view-file?refId=<id>&tableType=7001

Xuất: bond_news_raw.csv/.json (Tin khác) · bond_hnxnews_raw.csv/.json (Tin từ sở)
"""
import csv
import json
import re
import sys
import time
import unicodedata

import requests
import urllib3
from bs4 import BeautifulSoup

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
VERIFY_SSL = False

BASE = "https://cbonds.hnx.vn"
PAGE_URL = BASE + "/to-chuc-phat-hanh/tin-cong-bo"
SEARCH_URL = BASE + "/to-chuc-phat-hanh/tin-cong-bo-x"
OTHER_IDX, SO_IDX = 1, 3
TABLE_OF = {0: 0, 2: 1, 1: 2, 3: 3}     # index tab -> thứ tự bảng trong HTML
PAGE_SIZE = 100
DELAY = 0.4

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
    "X-Requested-With": "XMLHttpRequest",
    "Origin": BASE,
    "Referer": PAGE_URL,
}

COLS_OTHER = ["article_id", "ngay_dang_tin", "ten_dn", "ma_tp", "tieu_de",
              "ghi_chu", "tinh_trang", "loai_su_kien"]
COLS_SO = ["article_id", "ngay_dang_tin", "nguon", "tieu_de", "tinh_trang"]


def strip_accent(s):
    return "".join(c for c in unicodedata.normalize("NFD", s or "")
                   if unicodedata.category(c) != "Mn").lower()


def classify_news(title):
    """Phân loại sự kiện tab "Tin khác".

    Hai nhóm có giá trị nghiệp vụ cao nhất (user chốt 22/07/2026):

    (1) KHẮC PHỤC CHẬM TRẢ — doanh nghiệp công bố ĐÃ TRẢ ở tab này chứ không phải tab
        "Tin bất thường", nên dashboard trước đây chỉ thấy 5 lượt khắc phục trong khi
        thực tế có ~90. Hệ quả: 33 mã bị treo trạng thái "đang chậm trả" oan.
          'khac_phuc'  - hoàn thành thanh toán SAU KHI đã chậm
          'tt_bo_sung' - trả bổ sung phần còn thiếu (khắc phục MỘT PHẦN, chưa xong)
        Phân biệt hai loại này quan trọng: trả bổ sung KHÔNG đồng nghĩa hết chậm trả.

    (2) CẢNH BÁO SỚM — sự kiện đi TRƯỚC cả CBTT gia hạn (gia hạn vốn đã đi trước chậm trả):
          'xin_y_kien' - xin ý kiến/lấy ý kiến trái chủ, hội nghị người sở hữu trái phiếu
        Muốn gia hạn thì luật buộc phải lấy ý kiến người sở hữu trước -> tin này là mắt
        xích sớm nhất quan sát được của chuỗi tái cơ cấu nợ.

    Ngoài ra: 'tat_toan' (đáo hạn/không còn dư nợ - dùng để đối chiếu), 'ls_thuc_te'
    (lãi suất thả nổi thực tế kỳ này), 'ban_lai' (nhà đầu tư thực hiện quyền bán lại),
    'xu_phat' (bị xử phạt hành chính), 'mua_lai', '' (còn lại).
    """
    t = strip_accent(title)
    # --- cảnh báo sớm: xét TRƯỚC nhóm thanh toán vì tiêu đề hay lẫn chữ "trả lãi"
    if any(k in t for k in ["y kien trai chu", "y kien nguoi so huu", "y kien cua nguoi so huu",
                            "hoi nghi nguoi so huu", "dai hoi nguoi so huu",
                            "y kien bang van ban ve viec gia han"]):
        return "xin_y_kien"
    if any(k in t for k in ["xu phat vi pham hanh chinh", "quyet dinh xu phat"]):
        return "xu_phat"
    if "quyen ban lai" in t or "thuc hien quyen ban" in t:
        return "ban_lai"
    if "lai suat thuc te" in t or ("lai suat" in t and "ky tinh lai" in t):
        return "ls_thuc_te"
    # --- khắc phục sau chậm trả (phải xét TRƯỚC 'tat_toan': tiêu đề thường có cả hai ý)
    if any(k in t for k in ["sau khi cham", "sau cham", "sau thoi gian bi cham",
                            "sau thoi gian cham", "sau chi cham"]):
        return "khac_phuc"
    if "bo sung" in t and any(k in t for k in ["thanh toan", "goc", "lai"]):
        return "tt_bo_sung"
    if any(k in t for k in ["khong con du no", "het du no", "ket thuc du no",
                            "hoan thanh viec thanh toan", "hoan thanh thanh toan",
                            "day du nghia vu", "tat toan", "thanh toan trai phieu dao han",
                            "thanh toan goc lai trai phieu dao han"]):
        return "tat_toan"
    if "mua lai" in t:
        return "mua_lai"
    return ""


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


def _article_id(td):
    a = td.find("a")
    if not a:
        return ""
    m = re.search(r"showArticle(?:HNX)?\('([^']+)'\)", a.get("href", "") or "")
    return m.group(1) if m else ""


def fetch_page(session, hdr, idx, page):
    body = [("keysSearch[]", x) for x in ["", "", "", "", "", "", "0"]]
    pages, nums = [1] * 4, [1] * 4
    pages[idx], nums[idx] = page, PAGE_SIZE
    body += [("currentPages[]", str(p)) for p in pages]
    body += [("numberRecord[]", str(n)) for n in nums]
    r = session.post(SEARCH_URL, data=body, headers=hdr, timeout=90, verify=VERIFY_SSL)
    r.raise_for_status()
    tables = BeautifulSoup(r.text, "html.parser").find_all("table")
    if len(tables) <= TABLE_OF[idx]:
        return None                      # phiên hết hạn -> caller làm mới token
    return tables[TABLE_OF[idx]]


def parse_other(tb):
    rows = []
    body = tb.find("tbody")
    if not body:
        return rows
    for tr in body.find_all("tr", recursive=False):
        tds = tr.find_all("td", recursive=False)
        if len(tds) < 8:
            continue
        title = tds[4].get_text(strip=True)
        rows.append({"article_id": _article_id(tds[4]),
                     "ngay_dang_tin": tds[1].get_text(strip=True),
                     "ten_dn": tds[2].get_text(strip=True),
                     "ma_tp": tds[3].get_text(strip=True),
                     "tieu_de": title,
                     "ghi_chu": tds[5].get_text(strip=True),
                     "tinh_trang": tds[6].get_text(strip=True),
                     "loai_su_kien": classify_news(title)})
    return rows


def parse_so(tb):
    rows = []
    body = tb.find("tbody")
    if not body:
        return rows
    for tr in body.find_all("tr", recursive=False):
        tds = tr.find_all("td", recursive=False)
        if len(tds) < 6:
            continue
        rows.append({"article_id": _article_id(tds[3]),
                     "ngay_dang_tin": tds[1].get_text(strip=True),
                     "nguon": tds[2].get_text(strip=True),
                     "tieu_de": tds[3].get_text(strip=True),
                     "tinh_trang": tds[4].get_text(strip=True)})
    return rows


def scrape_tab(session, hdr, idx, parser, label):
    all_rows, page = [], 1
    while True:
        tb = None
        for attempt in range(3):
            try:
                tb = fetch_page(session, hdr, idx, page)
                if tb is None:
                    session, tk = get_session_token()
                    hdr["CP-TOKEN"] = tk
                    continue
                break
            except Exception as e:
                print(f"  ! {label} trang {page} (lần {attempt+1}): {e}", file=sys.stderr)
                time.sleep(2)
        if tb is None:
            break
        rows = parser(tb)
        if not rows:
            break
        all_rows.extend(rows)
        if len(rows) < PAGE_SIZE:
            break
        page += 1
        time.sleep(DELAY)
    print(f"{label}: {len(all_rows)} tin")
    return all_rows


def save(rows, cols, csv_path, json_path):
    if not rows:
        print(f"Không có dữ liệu -> bỏ qua {csv_path}")
        return
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False)
    print(f"Đã lưu {len(rows)} tin -> {csv_path} / {json_path}")


if __name__ == "__main__":
    t0 = time.time()
    session, token = get_session_token()
    hdr = {"CP-TOKEN": token,
           "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"}

    other = scrape_tab(session, hdr, OTHER_IDX, parse_other, "Tin khác")
    save(other, COLS_OTHER, "bond_news_raw.csv", "bond_news_raw.json")
    if other:
        from collections import Counter
        c = Counter(r["loai_su_kien"] for r in other)
        print("  " + " · ".join(f"{k or 'khác'} {v}" for k, v in c.most_common()))

    so = scrape_tab(session, hdr, SO_IDX, parse_so, "Tin từ sở")
    save(so, COLS_SO, "bond_hnxnews_raw.csv", "bond_hnxnews_raw.json")

    print(f"Xong sau {time.time()-t0:.0f}s.")
