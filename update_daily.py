# -*- coding: utf-8 -*-
"""
ORCHESTRATOR cập nhật hàng ngày — xử lý đúng bài toán "data cũ có thể ĐỔI TRẠNG THÁI".

Nguyên tắc (xem STATUS.md mục "Tự động cập nhật"):
  1) Nguồn HAY ĐỔI + RẺ  -> quét lại TOÀN BỘ (ghi đè). Dashboard tính lại mọi trạng thái dẫn xuất
     (dư nợ, chậm trả, hết hiệu lực) từ snapshot hiện tại -> thay đổi trên data cũ TỰ ĐỘNG được bắt.
  2) Nguồn ĐẮT + BẤT BIẾN (GD thứ cấp) -> incremental (chỉ quét phiên mới + chồng lấn).
  3) An toàn: (a) GUARD chống mất data — không giữ bản mới nếu số bản ghi < GUARD_RATIO × bản cũ
     (nghi quét lỗi) -> khôi phục bản .prev; (b) SNAPSHOT-DIFF ghi 'changes_log.csv' để lưu lịch sử
     đổi trạng thái (Hiệu lực→Hết hiệu lực, đang chậm→đã khắc phục), bản ghi mới & bản ghi biến mất.

Chạy: python update_daily.py            (đầy đủ: scrape -> diff -> rebuild dashboard)
      python update_daily.py --no-scrape (chỉ diff lại + rebuild, không gọi scraper)
Kết quả: cập nhật *_raw.csv, changes_log.csv, update_state.json, dashboard.html.
"""
import os
import sys
import csv
import json
import shutil
import subprocess
from datetime import datetime

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable
ENV = {**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"}
GUARD_RATIO = 0.85           # bản mới < 85% bản cũ -> nghi lỗi, khôi phục bản cũ
CHANGES_LOG = "changes_log.csv"
STATE_FILE = "update_state.json"
ALERTS_LOG = "alerts_new_latepay.csv"   # lịch sử cảnh báo mã chậm trả mới
NOTIFY_PS = "notify_alert.ps1"

# name, scraper argv, raw csv, key cols, status cols (theo dõi đổi trạng thái), cột nhãn, có diff không
SOURCES = [
    ("Phát hành",  ["bond_issuance_scraper.py"], "bond_issuance_raw.csv",
     ["ma_tp", "ngay_phat_hanh"], ["tinh_trang"], "ten_dn", True),
    ("Mua lại",    ["bond_buyback_scraper.py"], "bond_buyback_raw.csv",
     ["ma_tp", "ngay_mua_lai", "ngay_dang_tin"], [], "ten_dn", True),
    ("Danh mục",   ["bond_catalog_scraper.py"], "bond_catalog_raw.csv",
     ["ma_cbtt"], ["trang_thai"], "ten_tcph", True),
    ("Xếp hạng",   ["bond_rating_scraper.py"], "bond_rating_raw.csv",
     ["ten_tcph", "ma_tp", "loai_xep_hang"], ["ket_qua_xhtn", "hieu_luc_tu_ngay"], "ten_tcph", True),
    ("Chậm trả",   ["bond_latepay_scraper.py"], "bond_latepay_raw.csv",
     ["article_id"], ["tinh_trang", "loai_su_kien"], "ten_dn", True),
    ("Ngành ICB",  ["vietcap_sector_scraper.py"], "vietcap_companies_raw.csv",
     ["code"], [], "name", False),
    # VSD: LIST rẻ (~50s) nhưng trang CHI TIẾT đắt (~11') và gần như BẤT BIẾN -> incremental
    # (nguyên tắc 2, giống GD thứ cấp). Theo dõi `tinh_trang` để bắt mã VSD hủy lưu ký.
    ("Lưu ký VSD", ["vsd_bond_scraper.py", "--incremental"], "vsd_bond_raw.csv",
     ["ma_ck"], ["tinh_trang", "so_luong"], "ten_tcdkck", True),
    ("GD thứ cấp", ["bond_secondary_scraper.py", "--incremental"], "bond_secondary_raw.csv",
     ["ngay_gd", "ma_tp"], [], "ma_tp", False),
]


def run_scraper(argv):
    print(f"  → chạy {' '.join(argv)}")
    try:
        r = subprocess.run([PY] + argv, cwd=HERE, env=ENV, capture_output=True,
                           text=True, timeout=3600, encoding="utf-8", errors="replace")
        tail = "\n".join((r.stdout or "").strip().splitlines()[-2:])
        if tail:
            print("    " + tail.replace("\n", "\n    "))
        if r.returncode != 0:
            print(f"    ! exit {r.returncode}: {(r.stderr or '')[-300:]}", file=sys.stderr)
        return r.returncode == 0
    except Exception as e:
        print(f"    !! lỗi chạy scraper: {e}", file=sys.stderr)
        return False


def read_csv(path):
    if not os.path.exists(path):
        return None
    try:
        return pd.read_csv(path, dtype=str).fillna("")
    except Exception:
        return None


def _keymap(df, keycols, statuscols, labelcol):
    m = {}
    have_label = labelcol in df.columns
    cols = keycols + statuscols + ([labelcol] if have_label else [])
    if not all(c in df.columns for c in keycols):
        return m
    for r in df[cols].itertuples(index=False):
        d = dict(zip(cols, r))
        k = tuple(d[c] for c in keycols)
        if k in m:
            continue
        m[k] = ({c: d[c] for c in statuscols}, d.get(labelcol, ""))
    return m


def diff(name, prev, new, keycols, statuscols, labelcol, run_date):
    """So bản cũ vs mới -> danh sách thay đổi (new / removed / status_change)."""
    changes = []
    if prev is None or new is None:
        return changes
    pm = _keymap(prev, keycols, statuscols, labelcol)
    nm = _keymap(new, keycols, statuscols, labelcol)
    for k, (st, lb) in nm.items():
        if k not in pm:
            changes.append([run_date, name, "|".join(k), lb, "Mới", "", "", ""])
        else:
            pst = pm[k][0]
            for c in statuscols:
                if pst.get(c, "") != st.get(c, ""):
                    changes.append([run_date, name, "|".join(k), lb, "Đổi trạng thái",
                                    c, pst.get(c, ""), st.get(c, "")])
    for k, (st, lb) in pm.items():
        if k not in nm:
            changes.append([run_date, name, "|".join(k), lb, "Biến mất", "", "", ""])
    return changes


def _cham_codes(df):
    """{mã TP chậm trả -> (tên DN, tiêu đề, ngày)} từ các dòng loai_su_kien=='cham_tra'."""
    m = {}
    if df is None or "loai_su_kien" not in df.columns or "ma_tp" not in df.columns:
        return m
    sub = df[df["loai_su_kien"] == "cham_tra"]
    for r in sub.itertuples(index=False):
        for c in str(getattr(r, "ma_tp", "")).split(","):
            c = c.strip()
            if c and c not in m:
                m[c] = (getattr(r, "ten_dn", ""), getattr(r, "tieu_de", ""),
                        getattr(r, "ngay_dang_tin", ""))
    return m


def detect_new_latepay(prev_df, new_df):
    """Mã TP LẦN ĐẦU xuất hiện trong nhóm chậm trả (chưa từng chậm ở bản cũ)."""
    prev = set(_cham_codes(prev_df).keys())
    return [{"ma": c, "dn": v[0], "td": v[1], "d": v[2]}
            for c, v in _cham_codes(new_df).items() if c not in prev]


def append_alerts(alerts, run_date):
    path = os.path.join(HERE, ALERTS_LOG)
    exists = os.path.exists(path)
    with open(path, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if not exists:
            w.writerow(["run_date", "ma_tp", "ten_dn", "ngay_cbtt", "tieu_de"])
        for a in alerts:
            w.writerow([run_date, a["ma"], a["dn"], a["d"], a["td"]])


def notify(title, message):
    """Cảnh báo desktop (balloon) — best-effort, không chặn tiến trình."""
    ps = os.path.join(HERE, NOTIFY_PS)
    if not os.path.exists(ps):
        return
    try:
        subprocess.Popen(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                          "-File", ps, "-Title", title, "-Message", message], cwd=HERE)
    except Exception as e:
        print(f"  (notify lỗi: {e})")


def append_changes(changes):
    if not changes:
        return
    path = os.path.join(HERE, CHANGES_LOG)
    exists = os.path.exists(path)
    with open(path, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if not exists:
            w.writerow(["run_date", "source", "key", "label", "change_type", "field", "old", "new"])
        w.writerows(changes)


def main(do_scrape=True):
    run_dt = datetime.now()
    run_date = run_dt.strftime("%d/%m/%Y %H:%M")
    print(f"=== CẬP NHẬT HÀNG NGÀY {run_date} ===")
    all_changes, state, new_alerts = [], [], []

    for name, argv, csvf, keys, status, label, do_diff in SOURCES:
        raw = os.path.join(HERE, csvf)
        prev_path = raw + ".prev"
        prev_df = read_csv(raw)
        prev_n = 0 if prev_df is None else len(prev_df)
        if prev_df is not None:
            shutil.copyfile(raw, prev_path)          # sao lưu trước khi ghi đè

        ok, guard, err = True, "", ""
        if do_scrape:
            ok = run_scraper(argv)
        new_df = read_csv(raw)
        new_n = 0 if new_df is None else len(new_df)

        # GUARD: bản mới hụt bất thường -> khôi phục bản cũ (trừ lần đầu prev_n==0)
        if prev_n > 0 and new_n < GUARD_RATIO * prev_n:
            guard = f"KHÔI PHỤC (mới {new_n} < {GUARD_RATIO:.0%}×{prev_n})"
            print(f"  ⚠ {name}: {guard} — giữ bản cũ.")
            shutil.copyfile(prev_path, raw)
            new_df, new_n, ok = prev_df, prev_n, False
        elif not ok:
            err = "scraper lỗi"

        if do_diff and ok and prev_df is not None:
            ch = diff(name, prev_df, new_df, keys, status, label, run_date)
            all_changes += ch
            if ch:
                print(f"  • {name}: {len(ch)} thay đổi ghi nhận")

        # CẢNH BÁO: mã chậm trả MỚI (chưa từng chậm ở bản cũ)
        if name == "Chậm trả" and ok and prev_df is not None:
            new_alerts = detect_new_latepay(prev_df, new_df)

        if os.path.exists(prev_path):
            os.remove(prev_path)
        state.append({"name": name, "ok": ok, "count": new_n, "prev": prev_n,
                      "delta": new_n - prev_n, "guard": bool(guard), "note": guard or err})
        print(f"  {name}: {new_n:,} bản ghi ({new_n-prev_n:+,})")

    append_changes(all_changes)
    print(f"\nTổng thay đổi trạng thái/bản ghi ghi vào {CHANGES_LOG}: {len(all_changes)}")

    # CẢNH BÁO mã chậm trả mới -> ghi lịch sử + thông báo desktop
    if new_alerts:
        append_alerts(new_alerts, run_date)
        codes = ", ".join(a["ma"] for a in new_alerts[:8]) + (" …" if len(new_alerts) > 8 else "")
        msg = f"{len(new_alerts)} mã TPDN chậm trả MỚI: {codes}"
        print(f"  ⚠ CẢNH BÁO: {msg}")
        notify("⚠ TPDN có mã chậm trả mới", msg)
    else:
        print("  (không có mã chậm trả mới)")

    # Ghi trạng thái TRƯỚC khi rebuild để dashboard nhúng đúng số liệu lần chạy này.
    st_path = os.path.join(HERE, STATE_FILE)
    st = {"last_run": run_date, "rebuild_ok": True, "n_changes": len(all_changes),
          "n_new_latepay": len(new_alerts), "new_latepay": new_alerts, "sources": state}
    json.dump(st, open(st_path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    # rebuild dashboard (nhẹ, không dựng Excel — hợp môi trường C: đầy)
    print("Rebuild dashboard...")
    rebuild_ok = run_scraper(["_rebuild_sec.py"])
    if not rebuild_ok:                      # cập nhật lại cờ nếu rebuild lỗi
        st["rebuild_ok"] = False
        json.dump(st, open(st_path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"=== XONG. Trạng thái ghi ở {STATE_FILE} ===")


if __name__ == "__main__":
    main(do_scrape="--no-scrape" not in sys.argv)
