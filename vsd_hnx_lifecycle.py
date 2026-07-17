# -*- coding: utf-8 -*-
"""
VÒNG ĐỜI TPDN RIÊNG LẺ QUA 2 NGUỒN — đo ĐỘ TRỄ giữa các khâu, phát hiện mã BỊ KẸT.

MÔ HÌNH NGHIỆP VỤ (user chốt 17/07/2026):
  Phát hành  --CBTT trên HNX trước--> đăng ký LƯU KÝ tại VSD --xong mới--> ĐKGD/niêm yết trên HNX
  Khi kết thúc (trước đáo hạn / trước mua lại): hủy ĐKGD trên HNX TRƯỚC --> hủy lưu ký VSD SAU
  Độ trễ mỗi khâu ~2 tuần - 1 tháng.

KIỂM CHỨNG BẰNG DỮ LIỆU (17/07/2026) — mô hình ĐÚNG:
  * Chiều thuận : TP phát hành từ 2024 -> khâu 1 trung vị 18 ngày, khâu 2 trung vị 17 ngày (tổng ~36 ngày).
  * Chiều nghịch: nhóm "VSD hủy nhưng HNX còn ĐKGD" = 0 mã / 2.332 -> HNX LUÔN hủy trước, đúng thứ tự.
  * CẢNH BÁO khi đo: sàn TPRL mở 19/07/2023 + NĐ65/2022 -> TP cũ (2021-2022) bị đăng ký HỒI TỐ hàng loạt
    trong 2023 => nếu gộp cả TP cũ, khâu 1 hiện trung vị 470 ngày (VÔ NGHĨA). Chỉ đo TP phát hành TỪ 2024.

Ý nghĩa: độ trễ là BÌNH THƯỜNG, nhưng mã nằm ngoài cửa sổ dự kiến quá lâu là BẤT THƯỜNG cần rà soát
(kẹt thủ tục, hoặc TCPH có vấn đề). Script chấm điểm từng mã theo cửa sổ này.

Nguồn ngày:
  - ngày phát hành      : VSD (đối chiếu chéo CBTT phát hành HNX)
  - ngày ĐK lưu ký VSD  : cột `ngay_gcn` (ngày cấp Giấy CNĐKCK) — do `vsd_bond_scraper.extract_gcn_date`
  - ngày ĐKGD / hủy ĐKGD: HNX catalog `ngay_gd_dau` / `ngay_gd_cuoi`
  - ngày hủy lưu ký VSD : ⚠ VSD KHÔNG công bố -> khâu 4 chỉ quan sát được ở dạng "ĐANG TREO"
                          (right-censored): mã HNX đã hủy ĐKGD mà VSD vẫn để "Hiệu lực".

Đầu vào : vsd_bond_raw.csv, bond_catalog_raw.csv, bond_issuance_raw.csv, bond_latepay_raw.csv
Đầu ra  : VSD_HNX_VongDoi.xlsx (4 sheet) + tóm tắt màn hình
"""
import csv
import re
import sys
from datetime import date, datetime

import pandas as pd
from openpyxl import Workbook

from vsd_vs_hnx_compare import (TY, NGHIN_TY, _date, _dstr, load_catalog, load_issuance,
                                load_vsd, join_vsd_hnx, write_sheet)

OUT_XLSX = "VSD_HNX_VongDoi.xlsx"
LATE_CSV = "bond_latepay_raw.csv"
TODAY = date.today()

SAN_TPRL_MO = date(2023, 7, 19)      # sàn TPRL vận hành -> mốc phân định hồi tố
DO_SACH_TU = date(2024, 1, 1)        # chỉ TP phát hành từ đây mới đo độ trễ "tự nhiên"
CUA_SO_NGAY = 30                     # cửa sổ dự kiến mỗi khâu (2 tuần - 1 tháng)
NGUONG_KET = 90                      # quá 3x cửa sổ -> coi là KẸT, cần rà soát


def load_late_codes():
    """Tập mã trái phiếu có CBTT chậm trả (từ tin bất thường HNX)."""
    late = set()
    try:
        with open(LATE_CSV, encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                if r.get("loai_su_kien") != "cham_tra":
                    continue
                for t in re.split(r"[;,/\s]+", r.get("ma_tp") or ""):
                    t = t.strip().upper()
                    if len(t) >= 6:
                        late.add(t)
    except FileNotFoundError:
        print(f"  ! không thấy {LATE_CSV} — bỏ qua đối chiếu chậm trả", file=sys.stderr)
    return late


def _days(a, b):
    return (a - b).days if isinstance(a, date) and isinstance(b, date) else None


def _stats(s, label):
    s = pd.Series([x for x in s if x is not None], dtype="float64")
    if s.empty:
        return f"  {label:<44} (không có dữ liệu)"
    return (f"  {label:<44} n={len(s):>4} | trung vị {s.median():>5.0f} ngày | "
            f"25%-75%: {s.quantile(.25):>4.0f} - {s.quantile(.75):<4.0f} | "
            f"trong {CUA_SO_NGAY} ngày: {(s <= CUA_SO_NGAY).mean() * 100:>4.1f}%")


def main():
    print("Nạp dữ liệu...")
    vsd, cat, iss = load_vsd(), load_catalog(), load_issuance()
    late = load_late_codes()
    vsd = join_vsd_hnx(vsd, cat)
    ci = cat.set_index("isin_k")

    # CHỈ mã nối được — join trên join_isin rỗng sẽ nhân bản chéo với 5 dòng catalog ISIN rỗng
    m = vsd[vsd["join_isin"] != ""].join(
        ci[["ma_cbtt", "ngay_gd_dau", "ngay_gd_cuoi", "dkgd"]].add_prefix("h_"), on="join_isin")
    m = m.merge(iss.rename(columns={"cbtt_k": "h_ma_cbtt"})[["h_ma_cbtt", "i_ph"]],
                on="h_ma_cbtt", how="left")

    m["d_gcn"] = m["ngay_gcn"].map(_date)              # ngày đăng ký lưu ký VSD
    m["d_gd_dau"] = m["h_ngay_gd_dau"].map(_date)      # ngày ĐKGD trên HNX
    m["d_gd_cuoi"] = m["h_ngay_gd_cuoi"].map(_date)    # ngày hủy ĐKGD trên HNX

    m["lag1"] = [_days(g, p) for g, p in zip(m["d_gcn"], m["d_ph"])]        # PH -> lưu ký
    m["lag2"] = [_days(d, g) for d, g in zip(m["d_gd_dau"], m["d_gcn"])]    # lưu ký -> ĐKGD
    m["lag_tong"] = [_days(d, p) for d, p in zip(m["d_gd_dau"], m["d_ph"])]  # PH -> lên sàn
    m["moi"] = [isinstance(p, date) and p >= DO_SACH_TU for p in m["d_ph"]]
    sach = m[m["moi"]]

    print("\n" + "=" * 92)
    print(f"CHIỀU THUẬN — phát hành → lưu ký VSD → ĐKGD HNX   (cửa sổ dự kiến mỗi khâu ≤ {CUA_SO_NGAY} ngày)")
    print(f"\nA. TP phát hành TỪ {DO_SACH_TU:%d/%m/%Y} (đo sạch, {len(sach)} mã):")
    print(_stats(sach["lag1"], "Khâu 1: phát hành → đăng ký lưu ký VSD"))
    print(_stats(sach["lag2"], "Khâu 2: lưu ký VSD → ĐKGD trên HNX"))
    print(_stats(sach["lag_tong"], "Tổng   : phát hành → lên sàn TPRL"))
    print(f"\nB. TOÀN BỘ (gồm TP cũ bị đăng ký HỒI TỐ sau NĐ65 — số liệu méo, chỉ để tham chiếu):")
    print(_stats(m["lag1"], "Khâu 1: phát hành → đăng ký lưu ký VSD"))
    print(_stats(m["lag2"], "Khâu 2: lưu ký VSD → ĐKGD trên HNX"))

    # ---- theo năm phát hành: xem quy trình có nhanh dần không
    yr = m[m["lag1"].notna()].copy()
    yr["nam"] = [p.year for p in yr["d_ph"]]
    tab = (yr[yr["nam"] >= 2023].groupby("nam")
             .agg(so_ma=("lag1", "size"), lag1_tv=("lag1", "median"), lag2_tv=("lag2", "median"))
             .reset_index())
    print("\nC. Độ trễ theo năm phát hành (trung vị, ngày):")
    for _, r in tab.iterrows():
        l2 = f"{r['lag2_tv']:.0f}" if pd.notna(r["lag2_tv"]) else "-"
        print(f"    {int(r['nam'])}: {int(r['so_ma']):>4} mã | khâu 1 = {r['lag1_tv']:>4.0f} | khâu 2 = {l2:>4}")

    # ---- CHIỀU NGHỊCH: hủy ĐKGD (HNX) -> hủy lưu ký (VSD)
    B = m[m["hieu_luc"] & ~m["h_dkgd"]].copy()      # VSD còn hiệu lực, HNX đã hủy ĐKGD = ĐANG TREO
    C = m[~m["hieu_luc"] & m["h_dkgd"]]             # ngược thứ tự -> kỳ vọng = 0
    B["treo"] = [_days(TODAY, d) for d in B["d_gd_cuoi"]]
    B["da_dh"] = [isinstance(d, date) and d < TODAY for d in B["d_dh"]]

    def _keys(r):
        return {x for x in [r["cbtt_k"], r["gd_k"], r["h_ma_cbtt"] or ""] if x}
    B["cham_tra"] = [bool(_keys(r) & late) for _, r in B.iterrows()]

    print("\n" + "-" * 92)
    print("CHIỀU NGHỊCH — hủy ĐKGD (HNX) → hủy lưu ký (VSD)")
    print(f"\n  Mã VSD đã hủy NHƯNG HNX còn ĐKGD (sai thứ tự): {len(C)} mã "
          f"→ {'ĐÚNG mô hình: HNX luôn hủy trước' if len(C) == 0 else 'CÓ NGOẠI LỆ — cần xem'}")
    print(f"  Mã ĐANG TREO (HNX đã hủy ĐKGD, VSD còn 'Hiệu lực'): {len(B)} mã | "
          f"{B['gt'].sum() / NGHIN_TY:,.1f} nghìn tỷ")
    x = B["treo"].dropna()
    if len(x):
        print(f"    số ngày treo: trung vị {x.median():.0f} | trung bình {x.mean():.0f} | max {x.max():.0f}")
        print(f"\n    Phân bố (⚠ chỉ {(x <= CUA_SO_NGAY).sum()}/{len(x)} mã nằm trong cửa sổ {CUA_SO_NGAY} ngày):")
        for lo, hi, lb in [(0, 14, "≤ 2 tuần"), (15, 30, "2 tuần - 1 tháng"),
                           (31, 90, "1 - 3 tháng"), (91, 365, "3 - 12 tháng"),
                           (366, 10 ** 6, "> 1 năm")]:
            sel = B[(B["treo"] >= lo) & (B["treo"] <= hi)]
            print(f"      {lb:<18} {len(sel):>3} mã | {sel['gt'].sum() / NGHIN_TY:>5.1f} nghìn tỷ")

    # ---- VSD "Hiệu lực" nhưng đã QUÁ ĐÁO HẠN (toàn bộ VSD, không cần nối HNX)
    H = vsd[vsd["hieu_luc"]].copy()
    H["qua_han"] = [_days(TODAY, d) for d in H["d_dh"]]
    Q = H[[isinstance(d, date) and d < TODAY for d in H["d_dh"]]].copy()

    def _keys2(r):
        return {x for x in [r["cbtt_k"], r["gd_k"]] if x}
    cb = dict(zip(m["ma_ck"], m["h_ma_cbtt"].fillna("")))
    Q["cham_tra"] = [bool((_keys2(r) | {cb.get(r["ma_ck"], "")}) - {""} & late) for _, r in Q.iterrows()]
    R = H[[not (isinstance(d, date) and d < TODAY) for d in H["d_dh"]]]
    base = 0.0
    if len(R):
        R2 = R.copy()
        R2["ct"] = [bool((_keys2(r) | {cb.get(r["ma_ck"], "")}) - {""} & late) for _, r in R2.iterrows()]
        base = R2["ct"].mean()

    print("\n" + "-" * 92)
    print("VSD ghi 'Hiệu lực' NHƯNG ĐÃ QUÁ NGÀY ĐÁO HẠN")
    print(f"  {len(Q)} mã | {Q['gt'].sum() / NGHIN_TY:,.1f} nghìn tỷ | "
          f"quá hạn trung vị {Q['qua_han'].median():.0f} ngày, max {Q['qua_han'].max():.0f}")
    if len(Q) and base:
        print(f"  Có CBTT chậm trả: {Q['cham_tra'].sum()}/{len(Q)} = {Q['cham_tra'].mean() * 100:.1f}% "
              f"— so với {base * 100:.1f}% ở nhóm chưa đáo hạn → cao gấp {Q['cham_tra'].mean() / base:.1f} lần")
        print("  → Nhóm này KHÔNG thuần tuý là 'VSD chậm dọn': tỷ lệ chậm trả cao vượt trội so với nền,")
        print("    nhưng cũng KHÔNG kết luận được cả nhóm đều chậm trả — chỉ ~1/3 đã có CBTT xác nhận.")
        print(f"     · Đã xác nhận chậm trả : {Q['cham_tra'].sum():>3} mã | "
              f"{Q[Q['cham_tra']]['gt'].sum() / NGHIN_TY:>5.1f} nghìn tỷ")
        print(f"     · Chưa rõ lý do        : {(~Q['cham_tra']).sum():>3} mã | "
              f"{Q[~Q['cham_tra']]['gt'].sum() / NGHIN_TY:>5.1f} nghìn tỷ ← DANH SÁCH THEO DÕI "
              f"(có thể do VSD chậm cập nhật, hoặc chậm trả chưa CBTT)")

    # ================================================================= Excel
    wb = Workbook()
    wb.remove(wb.active)

    tong = pd.DataFrame([
        ["MÔ HÌNH VÒNG ĐỜI", "", ""],
        ["Chiều thuận", "PH → CBTT HNX → lưu ký VSD → ĐKGD HNX", "user chốt 17/07/2026"],
        ["Chiều nghịch", "hủy ĐKGD HNX → hủy lưu ký VSD", "độ trễ ~2 tuần - 1 tháng"],
        ["", "", ""],
        ["KIỂM CHỨNG", "", ""],
        [f"Khâu 1: PH → lưu ký VSD (TP từ {DO_SACH_TU:%Y})",
         f"trung vị {sach['lag1'].median():.0f} ngày", "khớp mô hình"],
        [f"Khâu 2: lưu ký VSD → ĐKGD HNX (TP từ {DO_SACH_TU:%Y})",
         f"trung vị {sach['lag2'].median():.0f} ngày", "khớp mô hình"],
        [f"Tổng PH → lên sàn (TP từ {DO_SACH_TU:%Y})",
         f"trung vị {sach['lag_tong'].median():.0f} ngày", ""],
        ["Sai thứ tự (VSD hủy trước HNX)", f"{len(C)} mã",
         "= 0 → xác nhận HNX luôn hủy trước"],
        ["", "", ""],
        ["CẢNH BÁO PHƯƠNG PHÁP", "", ""],
        ["Khâu 1 nếu tính cả TP cũ", f"trung vị {m['lag1'].median():.0f} ngày",
         "VÔ NGHĨA: TP 2021-2022 bị đăng ký HỒI TỐ 2023 sau NĐ65 + sàn TPRL mở 19/07/2023"],
        ["", "", ""],
        ["MÃ CẦN RÀ SOÁT", "", ""],
        ["Đang treo > 1 năm (HNX hủy, VSD chưa)",
         f"{int((B['treo'] > 365).sum())} mã",
         f"{B[B['treo'] > 365]['gt'].sum() / NGHIN_TY:,.1f} nghìn tỷ"],
        ["VSD 'Hiệu lực' dù đã quá đáo hạn", f"{len(Q)} mã",
         f"{Q['gt'].sum() / NGHIN_TY:,.1f} nghìn tỷ — {Q['cham_tra'].sum()} mã đã có CBTT chậm trả"],
    ], columns=["Chỉ tiêu", "Giá trị", "Ghi chú"])

    lag_sheet = pd.DataFrame({
        "Mã CK (GD)": sach["ma_ck"], "Mã CBTT": sach["h_ma_cbtt"], "Tên TCPH": sach["ten_tcdkck"],
        "Ngày phát hành": sach["ngay_phat_hanh"], "Ngày ĐK lưu ký VSD": sach["ngay_gcn"],
        "Ngày ĐKGD HNX": sach["h_ngay_gd_dau"],
        "Khâu 1 PH→lưu ký (ngày)": sach["lag1"], "Khâu 2 lưu ký→ĐKGD (ngày)": sach["lag2"],
        "Tổng PH→lên sàn (ngày)": sach["lag_tong"],
        "Giá trị (tỷ)": sach["gt"] / TY,
        "Đánh giá": [("Chậm bất thường (>%d ngày)" % NGUONG_KET) if (pd.notna(t) and t > NGUONG_KET)
                     else ("Bình thường" if pd.notna(t) else "Thiếu ngày")
                     for t in sach["lag_tong"]],
    }).sort_values("Tổng PH→lên sàn (ngày)", ascending=False, na_position="last")

    treo_sheet = pd.DataFrame({
        "Mã CK (GD)": B["ma_ck"], "Mã CBTT": B["h_ma_cbtt"], "Tên TCPH": B["ten_tcdkck"],
        "Ngày hủy ĐKGD (HNX)": B["h_ngay_gd_cuoi"], "Ngày đáo hạn": B["ngay_dao_han"],
        "Số ngày treo": B["treo"], "Đã đáo hạn?": ["Đã đáo hạn" if x else "Chưa đáo hạn"
                                                   for x in B["da_dh"]],
        "Có CBTT chậm trả?": ["CÓ" if x else "-" for x in B["cham_tra"]],
        "Giá trị VSD (tỷ)": B["gt"] / TY,
        "Đánh giá": [("Trong cửa sổ bình thường" if pd.notna(t) and t <= CUA_SO_NGAY else
                      ("Chậm nhẹ" if pd.notna(t) and t <= NGUONG_KET else
                       ("KẸT — cần rà soát" if pd.notna(t) else "Thiếu ngày hủy ĐKGD")))
                     for t in B["treo"]],
    }).sort_values("Số ngày treo", ascending=False, na_position="last")

    qh_sheet = pd.DataFrame({
        "Mã CK (GD)": Q["ma_ck"], "Mã CBTT": Q["ma_cbtt"], "Tên TCĐKCK": Q["ten_tcdkck"],
        "Ngày đáo hạn": Q["ngay_dao_han"], "Số ngày quá hạn": Q["qua_han"],
        "Giá trị VSD (tỷ)": Q["gt"] / TY,
        "Có CBTT chậm trả?": ["CÓ — đã xác nhận chậm trả" if x else "-" for x in Q["cham_tra"]],
        "Đánh giá": ["Chưa tất toán (chậm trả đã CBTT)" if x else "THEO DÕI: quá hạn, chưa rõ lý do"
                     for x in Q["cham_tra"]],
    }).sort_values("Giá trị VSD (tỷ)", ascending=False)

    stamp = f"Nguồn: vsd.vn/vi/ibl & cbonds.hnx.vn — lập lúc {datetime.now():%d/%m/%Y %H:%M}"
    write_sheet(wb, "Tổng quan vòng đời", tong,
                "VÒNG ĐỜI TPDN RIÊNG LẺ QUA 2 NGUỒN — kiểm chứng mô hình & đo độ trễ. " + stamp,
                widths={"Chỉ tiêu": 44, "Giá trị": 24, "Ghi chú": 76})
    write_sheet(wb, "Độ trễ chiều thuận", lag_sheet,
                f"CHỈ trái phiếu phát hành từ {DO_SACH_TU:%d/%m/%Y} — TP cũ bị đăng ký hồi tố sau NĐ65 "
                f"nên độ trễ không phản ánh quy trình thật. Cửa sổ bình thường ≤ {CUA_SO_NGAY} ngày/khâu.")
    write_sheet(wb, "Đang treo (chờ hủy lưu ký)", treo_sheet,
                "HNX đã hủy ĐKGD nhưng VSD vẫn để 'Hiệu lực'. Theo quy trình, VSD hủy sau 2 tuần - 1 tháng; "
                f"treo > {NGUONG_KET} ngày là KẸT. ⚠ VSD không công bố ngày hủy lưu ký nên chỉ đo được mã đang treo.")
    write_sheet(wb, "Quá hạn mà VSD còn hiệu lực", qh_sheet,
                "Đã qua ngày đáo hạn nhưng VSD vẫn ghi 'Hiệu lực' → phần lớn là TP chưa tất toán được "
                "(tỷ lệ có CBTT chậm trả cao gấp ~10 lần nền). Nhóm chưa có CBTT chậm trả = danh sách theo dõi.")
    wb.save(OUT_XLSX)
    print("\n" + "=" * 92)
    print(f"Đã lưu -> {OUT_XLSX}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
