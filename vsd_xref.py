# -*- coding: utf-8 -*-
"""
LỚP ĐỐI CHIẾU VSD cho pipeline HNX — module dùng chung (như `sector_map.py`).

NGUYÊN TẮC (user chốt 17/07/2026): **HNX là CƠ SỞ**. VSD chỉ là lớp đối chiếu phủ lên.
=> KHÔNG thêm mã VSD-only vào universe; không đổi bất kỳ số liệu nào của HNX.
   Mã VSD không nối được chỉ hiện là "Không có ở VSD".

Khóa nối: mã CBTT của HNX (`ma_tp`) -> VSD. VSD cho 3 khóa (mã CK = mã GD, mã CBTT trong tên CK, ISIN);
danh mục ĐKGD HNX (`bond_catalog_raw.csv`) là bảng crosswalk giữa mã CBTT <-> mã GD <-> ISIN.
⚠ TP phát hành gần đây có mã CBTT TRÙNG mã GD (ABB12501, P5332601) -> phải nhận cả hai.

Ý nghĩa đối chiếu — hai nguồn ĐO HAI THỨ KHÁC NHAU:
  HNX dư nợ  = giá trị phát hành − mua lại − đáo hạn (tính từ CBTT)
  VSD        = giá trị chứng khoán ĐĂNG KÝ LƯU KÝ của mã còn "Hiệu lực"
Chênh hợp lệ do ĐỘ TRỄ VÒNG ĐỜI (hủy ĐKGD trên HNX trước, hủy lưu ký VSD sau ~2 tuần-1 tháng,
thực tế trung vị ~232 ngày) => "Lệch" KHÔNG có nghĩa là sai, mà là cần đọc kèm trạng thái VSD.
"""
import csv
import os
import re

TY = 1_000_000_000
TOL = 100_000_000          # sai số bỏ qua: 0,1 tỷ (chênh nhỏ hơn coi như khớp)

VSD_CSV = "vsd_bond_raw.csv"
CAT_CSV = "bond_catalog_raw.csv"

# trạng thái đối chiếu (dùng chung cho Excel + dashboard)
KHOP = "Khớp"
LECH = "Lệch"
VSD_HUY = "VSD đã hủy ĐK"
NO_VSD = "Không có ở VSD"
# HNX đã tất toán (đáo hạn/mua lại hết) nhưng VSD chưa hủy lưu ký -> ĐỘ TRỄ VÒNG ĐỜI, KHÔNG phải sai số.
# Tách khỏi "Lệch" vì gộp chung sẽ gây hiểu nhầm: 436/466 mã lệch thuộc loại này (17/07/2026).
VSD_CHUA_HUY = "VSD chưa hủy lưu ký"

NGUON_VSD_HL = "Hiệu lực"
NGUON_VSD_HUY = "Hủy đăng ký"
NGUON_VSD_NONE = "Không có"


def _num(s):
    s = re.sub(r"[^\d]", "", str(s or ""))
    return int(s) if s else 0


def _here(fn):
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), fn)


def _read(fn):
    p = fn if os.path.exists(fn) else _here(fn)
    with open(p, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def load_xref(vsd_csv=VSD_CSV, cat_csv=CAT_CSV):
    """-> dict: mã CBTT (HNX) -> {gt, tinh_trang, ma_ck, isin, dh, hieu_luc}

    Trả {} nếu chưa có dữ liệu VSD => pipeline HNX chạy bình thường, chỉ không có cột đối chiếu.
    """
    try:
        vsd = _read(vsd_csv)
    except FileNotFoundError:
        return {}
    try:
        cat = _read(cat_csv)
    except FileNotFoundError:
        cat = []

    # crosswalk từ danh mục ĐKGD: ISIN -> mã CBTT, mã GD -> mã CBTT
    isin2cbtt, gd2cbtt = {}, {}
    for r in cat:
        cbtt = (r.get("ma_cbtt") or "").strip().upper()
        if not cbtt:
            continue
        isin = (r.get("isin") or "").strip().upper()
        gd = (r.get("ma_gd") or "").strip().upper()
        if isin:
            isin2cbtt.setdefault(isin, cbtt)
        if gd:
            gd2cbtt.setdefault(gd, cbtt)

    out = {}
    for r in vsd:
        ma_ck = (r.get("ma_ck") or "").strip().upper()
        isin = (r.get("isin") or "").strip().upper()
        cbtt = (r.get("ma_cbtt") or "").strip().upper()
        # thứ tự ưu tiên: mã CBTT VSD tự khai -> ISIN qua crosswalk -> mã GD qua crosswalk -> mã CK
        key = cbtt or isin2cbtt.get(isin) or gd2cbtt.get(ma_ck) or ma_ck
        if not key:
            continue
        gt = _num(r.get("gia_tri_dk"))
        if not gt:
            gt = _num(r.get("so_luong")) * _num(r.get("menh_gia"))
        hl = (r.get("tinh_trang") or "").strip().lower() == "hiệu lực"
        rec = {"gt": gt, "tinh_trang": (r.get("tinh_trang") or "").strip(),
               "ma_ck": ma_ck, "isin": isin, "dh": r.get("ngay_dao_han") or "",
               "hieu_luc": hl}
        # 1 mã CBTT có thể ứng nhiều mã VSD (hiếm) -> ưu tiên bản còn hiệu lực, rồi giá trị lớn hơn
        old = out.get(key)
        if old is None or (hl and not old["hieu_luc"]) or (hl == old["hieu_luc"] and gt > old["gt"]):
            out[key] = rec
    return out


def doi_chieu(ma_cbtt, gt_hnx, xref):
    """Đối chiếu 1 mã. gt_hnx: dư nợ theo HNX (VNĐ).
    -> (khop, nguon_vsd, gt_vsd_ty|None, chenh_ty|None)

    LƯU Ý ĐỌC SỐ: mã VSD đã hủy ĐK mà HNX còn dư nợ KHÔNG phải lỗi — VSD hủy lưu ký SAU;
    ngược lại VSD còn hiệu lực mà HNX hết dư nợ là do VSD chưa dọn (trung vị ~232 ngày).
    """
    r = xref.get((ma_cbtt or "").strip().upper()) if xref else None
    if not r:
        return NO_VSD, NGUON_VSD_NONE, None, None
    gt_vsd = r["gt"]
    chenh = (gt_hnx or 0) - gt_vsd
    if not r["hieu_luc"]:
        return VSD_HUY, NGUON_VSD_HUY, round(gt_vsd / TY, 1), round(chenh / TY, 1)
    if abs(chenh) <= TOL:
        khop = KHOP
    elif (gt_hnx or 0) <= 0 < gt_vsd:
        # HNX hết dư nợ (đáo hạn / mua lại hết) mà VSD còn hiệu lực = đang chờ hủy lưu ký.
        khop = VSD_CHUA_HUY
    else:
        khop = LECH
    return khop, NGUON_VSD_HL, round(gt_vsd / TY, 1), round(chenh / TY, 1)


def mo_ta_nguon(khop, gt_hnx_ty, gt_vsd_ty):
    """Cột "Nguồn" (user chốt 17/07/2026): TRÙNG -> chỉ ghi trạng thái khớp;
    KHÔNG trùng -> nêu RÕ khác biệt ngay trong ô, khỏi phải dò cột khác."""
    if khop == KHOP:
        return "Khớp (HNX = VSD)"
    if khop == LECH:
        return f"Lệch: HNX {gt_hnx_ty:,.1f} ≠ VSD {gt_vsd_ty:,.1f} (tỷ)"
    if khop == VSD_CHUA_HUY:
        return (f"Khác (độ trễ): HNX đã tất toán, VSD chưa hủy lưu ký "
                f"({gt_vsd_ty:,.1f} tỷ còn đăng ký)")
    if khop == VSD_HUY:
        return f"Khác: VSD đã hủy lưu ký (HNX còn {gt_hnx_ty:,.1f} tỷ)"
    return "Khác: chỉ có ở HNX, không có ở VSD"


if __name__ == "__main__":
    x = load_xref()
    print(f"xref: {len(x)} mã CBTT nối được sang VSD")
    hl = sum(1 for v in x.values() if v["hieu_luc"])
    print(f"  còn hiệu lực: {hl} | đã hủy ĐK: {len(x) - hl}")
