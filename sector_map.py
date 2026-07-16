# -*- coding: utf-8 -*-
"""
Phân loại NGÀNH tổ chức phát hành TPDN (dùng chung cho build_reports / _rebuild_sec).

Chiến lược 2 lớp (theo yêu cầu 15/07 v3 - bóc tách nhóm "Khác"):
  1) OVERRIDES: gán tay theo CHUỖI CON đặc trưng trong tên (khớp cả 2 dạng tên của site HNX:
     "VJC - CÔNG TY ... VIETJET" và "CÔNG TY CỔ PHẦN HÀNG KHÔNG VIETJET"). Ưu tiên cao nhất,
     đặt các mục CỤ THỂ trước mục CHUNG. Dùng cho DN lớn mà tên không lộ ngành (Khang Điền, HAGL,
     Tasco, Thành Thành Công, Sovico, Núi Pháo, Nam Long, các SPV BĐS/nghỉ dưỡng...).
  2) KEYWORDS: nếu không trúng override thì suy theo từ khóa ngành. Thứ tự kiểm tra QUAN TRỌNG:
     nhóm cụ thể (du lịch, năng lượng, sản xuất...) đặt TRƯỚC nhóm rộng (BĐS, thương mại, đa ngành).
  3) Không trúng gì -> "Khác".

Lưu ý ranh giới BĐS vs Xây dựng: tên có "bất động sản/địa ốc/nhà/đô thị/land/homes" -> BĐS;
chỉ có "xây dựng/hạ tầng" -> Xây dựng - Hạ tầng (chấp nhận một số SPV BĐS lọt sang Xây dựng;
các tên lớn đã được OVERRIDES nắn lại).

CẬP NHẬT 16/07/2026 - LỚP VIETCAP ICB (giữa OVERRIDES và KEYWORDS):
  Đối chiếu tên/mã CK của TCPH với danh mục ICB VietCap (`vietcap_companies_raw.csv`,
  scraper `vietcap_sector_scraper.py`). TCPH nào khớp (mã CK niêm yết) -> lấy ngành ICB (map về
  13 nhóm dự án qua ICB2_GROUP), CHÍNH XÁC hơn suy từ từ khóa. Phủ ~25% số TCPH nhưng ~63% GIÁ TRỊ
  (SPV chưa niêm yết không có -> rơi về KEYWORDS như cũ). Thứ tự ưu tiên:
  OVERRIDES (quyết định tay) -> VietCap ICB -> KEYWORDS. `classify_with_source()` trả kèm NGUỒN.
"""
import csv
import os
import re
import unicodedata

# Thứ tự nhóm chuẩn (dùng cho màu & sắp xếp trên dashboard). "Khác" luôn cuối.
GROUP_ORDER = [
    "Ngân hàng",
    "Bất động sản",
    "Du lịch - Nghỉ dưỡng",
    "Xây dựng - Hạ tầng",
    "Năng lượng",
    "Sản xuất - Công nghiệp",
    "Nông nghiệp - Thực phẩm",
    "Khai khoáng",
    "Hàng không - Vận tải",
    "Chứng khoán",
    "Tài chính - Bảo hiểm",
    "Thương mại - Dịch vụ",
    "Khác",
]

# (chuỗi con IN HOA có dấu, nhóm) - kiểm tra theo thứ tự, dừng ở mục đầu tiên khớp.
OVERRIDES = [
    # --- Hàng không - Vận tải ---
    ("VIETJET", "Hàng không - Vận tải"),
    # --- Sản xuất - Công nghiệp ---
    ("VINFAST", "Sản xuất - Công nghiệp"),
    ("Ô TÔ TRƯỜNG HẢI", "Sản xuất - Công nghiệp"),
    ("THACO", "Sản xuất - Công nghiệp"),
    ("GOLDSUN", "Sản xuất - Công nghiệp"),
    ("IN VÀ BAO BÌ", "Sản xuất - Công nghiệp"),
    # --- Nông nghiệp - Thực phẩm ---
    ("MASAN", "Nông nghiệp - Thực phẩm"),
    ("PAN FARM", "Nông nghiệp - Thực phẩm"),
    ("THÀNH THÀNH CÔNG", "Nông nghiệp - Thực phẩm"),   # TTC (mía đường)
    ("HOÀNG ANH GIA LAI", "Nông nghiệp - Thực phẩm"),  # HAGL (nông nghiệp)
    # --- Khai khoáng ---
    ("NÚI PHÁO", "Khai khoáng"),
    # --- Du lịch - Nghỉ dưỡng ---
    ("VINPEARL", "Du lịch - Nghỉ dưỡng"),
    ("DU LỊCH PHÚ QUỐC", "Du lịch - Nghỉ dưỡng"),
    ("SUNBAY", "Du lịch - Nghỉ dưỡng"),
    ("VẠN HƯƠNG", "Du lịch - Nghỉ dưỡng"),             # Đồi Rồng (Hải Phòng)
    ("SÀI GÒN - LÂM ĐỒNG", "Du lịch - Nghỉ dưỡng"),
    ("HƯNG THỊNH QUY NHƠN", "Du lịch - Nghỉ dưỡng"),
    ("BÀ NÀ", "Du lịch - Nghỉ dưỡng"),
    ("CÁP TREO", "Du lịch - Nghỉ dưỡng"),
    ("BÔNG SEN", "Du lịch - Nghỉ dưỡng"),              # Bông Sen Corp (khách sạn SG)
    ("CRYSTAL BAY", "Du lịch - Nghỉ dưỡng"),
    ("ĐẢO NGỌC XANH", "Du lịch - Nghỉ dưỡng"),
    ("SUMMER BEACH", "Du lịch - Nghỉ dưỡng"),
    ("MẶT TRỜI HẠ LONG", "Du lịch - Nghỉ dưỡng"),      # Sun World Hạ Long (không phải điện MT)
    ("ALLGREEN", "Du lịch - Nghỉ dưỡng"),              # Allgreen - Vượng Thành - Trùng Dương (Hồ Tràm)
    ("TOÀN HẢI VÂN", "Du lịch - Nghỉ dưỡng"),          # Toàn Hải Vân (nghỉ dưỡng)
    # --- Bất động sản (tên không lộ ngành) ---
    ("NAM LONG", "Bất động sản"),
    ("KHANG ĐIỀN", "Bất động sản"),
    ("HẢI PHÁT", "Bất động sản"),
    ("BECAMEX", "Bất động sản"),                        # BĐS khu công nghiệp
    ("GOLDEN HILL", "Bất động sản"),
    ("OSAKA GARDEN", "Bất động sản"),
    ("AQUA CITY", "Bất động sản"),
    ("TÂN THÀNH LONG AN", "Bất động sản"),
    ("MARINA CENTER", "Bất động sản"),
    ("HOÀNG PHÚ VƯƠNG", "Bất động sản"),
    ("EAGLE SIDE", "Bất động sản"),
    ("VẠN TRƯỜNG PHÁT", "Bất động sản"),
    ("SUNSHINE", "Bất động sản"),
    ("R&H", "Bất động sản"),                            # R&H Group (BĐS)
    ("NOVALAND", "Bất động sản"),
    ("NO VA", "Bất động sản"),
    ("HƯNG THỊNH", "Bất động sản"),                     # còn lại (không phải Quy Nhơn) -> BĐS
    ("THÁI SƠN - LONG AN", "Bất động sản"),
    ("THÀNH PHỐ AQUA", "Bất động sản"),
    ("THỜI ĐẠI MỚI T&T", "Bất động sản"),               # SPV BĐS nhóm T&T
    ("LUXURY LIVING", "Sản xuất - Công nghiệp"),        # nội thất
    # --- Xây dựng - Hạ tầng ---
    ("TASCO", "Xây dựng - Hạ tầng"),                    # BOT giao thông
    ("VINACONEX", "Xây dựng - Hạ tầng"),
    ("XUẤT NHẬP KHẨU VÀ XÂY DỰNG VIỆT NAM", "Xây dựng - Hạ tầng"),
    ("NAM QUANG", "Xây dựng - Hạ tầng"),               # phát triển hạ tầng
    # --- Năng lượng ---
    ("TRUNG NAM", "Năng lượng"),                        # Trungnam (điện gió/mặt trời)
    # --- Tài chính - Bảo hiểm ---
    ("F88", "Tài chính - Bảo hiểm"),
    # --- Các holding lớn: user 15/07 v4 chốt gộp vào BẤT ĐỘNG SẢN ---
    ("VINGROUP", "Bất động sản"),
    ("SOVICO", "Bất động sản"),
    ("I.P.A", "Bất động sản"),
]


# ---------------------------------------------------------------------------
# LỚP VIETCAP ICB: map ngành ICB (Lv2 + tinh chỉnh Lv3/Lv4) -> 13 nhóm dự án.
# ---------------------------------------------------------------------------
VC_CSV = "vietcap_companies_raw.csv"

# ICB Lv2 (tên tiếng Việt VietCap) -> nhóm dự án. 3 nhóm ICB2 cần tinh chỉnh sâu (xem _icb_to_group).
ICB2_GROUP = {
    "Ngân hàng": "Ngân hàng",
    "Bất động sản": "Bất động sản",
    "Bảo hiểm": "Tài chính - Bảo hiểm",
    "Dịch vụ tài chính": "Tài chính - Bảo hiểm",          # tinh chỉnh: chứng khoán -> Chứng khoán
    "Xây dựng và Vật liệu": "Xây dựng - Hạ tầng",
    "Hàng & Dịch vụ Công nghiệp": "Sản xuất - Công nghiệp",  # tinh chỉnh: vận tải -> Hàng không - Vận tải
    "Thực phẩm và đồ uống": "Nông nghiệp - Thực phẩm",
    "Hàng cá nhân & Gia dụng": "Sản xuất - Công nghiệp",
    "Ô tô và phụ tùng": "Sản xuất - Công nghiệp",
    "Tài nguyên Cơ bản": "Sản xuất - Công nghiệp",         # tinh chỉnh: khoáng sản -> Khai khoáng
    "Hóa chất": "Sản xuất - Công nghiệp",
    "Điện, nước & xăng dầu khí đốt": "Năng lượng",
    "Dầu khí": "Năng lượng",
    "Du lịch và Giải trí": "Du lịch - Nghỉ dưỡng",
    "Truyền thông": "Thương mại - Dịch vụ",
    "Bán lẻ": "Thương mại - Dịch vụ",
    "Công nghệ Thông tin": "Thương mại - Dịch vụ",         # dự án không có nhóm CNTT riêng
    "Y tế": "Sản xuất - Công nghiệp",                      # dược phẩm/thiết bị y tế (SX)
    "Viễn thông": "Thương mại - Dịch vụ",
}

_VC_CODE = {}   # mã CK -> (icb2_name, deep_text)  ; deep_text = icb3+icb4 (để tinh chỉnh)
_VC_NAME = {}   # tên chuẩn hoá -> (icb2_name, deep_text)
_VC_LOADED = False


def _strip_accent(s):
    return "".join(c for c in unicodedata.normalize("NFD", s or "")
                   if unicodedata.category(c) != "Mn").lower()


def _norm_name(s):
    s = _strip_accent(re.sub(r"^\s*[A-Z0-9]{3,4}\s*[-–]\s*", "", (s or "").strip()))
    s = re.sub(r"\b(cong ty|cp|co phan|tnhh|mtv|tap doan|ctcp|joint stock|company|jsc|"
               r"corporation|corp)\b", " ", s)
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _ticker(name):
    m = re.match(r"^([A-Z]{3}[A-Z0-9]?)\s*[-–]\s*", (name or "").strip())
    return m.group(1).upper() if m else None


def _load_vietcap():
    global _VC_LOADED
    if _VC_LOADED:
        return
    _VC_LOADED = True
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), VC_CSV)
    if not os.path.exists(path):
        return
    try:
        with open(path, encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                icb2 = (r.get("icb2_name") or "").strip()
                if not icb2:
                    continue
                deep = ((r.get("icb3_name") or "") + " " + (r.get("icb4_name") or "")).strip()
                code = (r.get("code") or "").strip().upper()
                if code:
                    _VC_CODE[code] = (icb2, deep)
                nm = _norm_name(r.get("name"))
                if nm:
                    _VC_NAME.setdefault(nm, (icb2, deep))
    except Exception:
        pass


def _icb_to_group(icb2, deep_text, name):
    base = ICB2_GROUP.get(icb2)
    if base is None:
        return None
    txt = (_strip_accent(deep_text) + " " + _strip_accent(name))
    if icb2 == "Dịch vụ tài chính":
        if "chung khoan" in txt or "securities" in txt:
            return "Chứng khoán"
    elif icb2 == "Hàng & Dịch vụ Công nghiệp":
        if any(k in txt for k in ["van tai", "hang khong", "cang", "logistic",
                                  "van chuyen", "tau bien", "airline", "hang hai"]):
            return "Hàng không - Vận tải"
    elif icb2 == "Tài nguyên Cơ bản":
        if any(k in txt for k in ["khoang", "khai thac", "mining", "mineral", "than "]):
            return "Khai khoáng"
    return base


def vietcap_group(name):
    """Ngành 13-nhóm theo ICB VietCap nếu TCPH khớp (mã CK hoặc tên); None nếu không khớp."""
    _load_vietcap()
    if not _VC_CODE and not _VC_NAME:
        return None
    tk = _ticker(name)
    if tk and tk in _VC_CODE:
        icb2, deep = _VC_CODE[tk]
        return _icb_to_group(icb2, deep, name)
    key = _norm_name(name)
    if key and key in _VC_NAME:
        icb2, deep = _VC_NAME[key]
        return _icb_to_group(icb2, deep, name)
    return None


def _keyword_group(s):
    # KEYWORDS (ngân hàng/CK trước vì rất rõ)
    if "NGÂN HÀNG" in s or "BANK" in s or re.search(r"\bTMCP\b", s):
        return "Ngân hàng"
    if "CHỨNG KHOÁN" in s or "SECURITIES" in s or "CTCK" in s:
        return "Chứng khoán"
    if any(k in s for k in ["HÀNG KHÔNG", "AVIATION", "AIRLINE",
                            "VẬN TẢI", "LOGISTIC", "CẢNG BIỂN", "TÀU BIỂN", "VẬN CHUYỂN"]):
        return "Hàng không - Vận tải"
    if any(k in s for k in ["KHOÁNG SẢN", "KHAI THÁC", "KHAI KHOÁNG", "MINERAL", "MINING"]):
        return "Khai khoáng"
    if any(k in s for k in ["NÔNG NGHIỆP", "NÔNG SẢN", "FARM", "THỰC PHẨM", "FOOD",
                            "CHĂN NUÔI", "THUỶ SẢN", "THỦY SẢN", "MÍA ĐƯỜNG", "CÀ PHÊ", "SỮA"]):
        return "Nông nghiệp - Thực phẩm"
    if any(k in s for k in ["DU LỊCH", "NGHỈ DƯỠNG", "RESORT", "KHÁCH SẠN", "HOTEL",
                            "GIẢI TRÍ", "GOLF", "SAFARI", "BEACH", "MARINA"]):
        return "Du lịch - Nghỉ dưỡng"
    if any(k in s for k in ["BẤT ĐỘNG SẢN", "ĐỊA ỐC", "LAND", "HOMES", "VINHOME",
                            "ĐÔ THỊ", "NHÀ Ở", "PROPERT", "REAL ESTATE", "REALTY",
                            "BĐS", "KINH DOANH NHÀ"]):
        return "Bất động sản"
    if any(k in s for k in ["ĐIỆN", "NĂNG LƯỢNG", "ENERGY", "SOLAR", "MẶT TRỜI",
                            "GIÓ", "THỦY ĐIỆN", "POWER", "PIN "]):
        return "Năng lượng"
    if any(k in s for k in ["SẢN XUẤT", "Ô TÔ", "THÉP", "XI MĂNG", "NHỰA", "BAO BÌ",
                            "DỆT", "MAY MẶC", "CÔNG NGHIỆP", "HOÁ CHẤT", "HÓA CHẤT",
                            "VẬT LIỆU", "CƠ KHÍ", "MANUFACTUR"]):
        return "Sản xuất - Công nghiệp"
    if any(k in s for k in ["XÂY DỰNG", "HẠ TẦNG", "GIAO THÔNG", "CONSTRUCTION",
                            "INFRA", "CẦU ĐƯỜNG", "THI CÔNG", "BOT"]):
        return "Xây dựng - Hạ tầng"
    if "TÀI CHÍNH" in s or "FINANCE" in s or "BẢO HIỂM" in s or "INSURANCE" in s:
        return "Tài chính - Bảo hiểm"
    if any(k in s for k in ["THƯƠNG MẠI", "DỊCH VỤ", "TRADING", "SERVICE", "XUẤT NHẬP KHẨU"]):
        return "Thương mại - Dịch vụ"
    # MẶC ĐỊNH cho DN tên thuần "Đầu tư/Phát triển/Tập đoàn" (không lộ ngành khác) = BẤT ĐỘNG SẢN
    # (theo yêu cầu user 15/07 v4: TPDN riêng lẻ phi ngân hàng ở VN đa phần là SPV bất động sản).
    # Các holding đa ngành THẬT (Vingroup, Sovico, I.P.A) đã được gán riêng ở OVERRIDES.
    if any(k in s for k in ["TẬP ĐOÀN", "HOLDING", "GROUP", "ĐẦU TƯ", "PHÁT TRIỂN",
                            "INVEST", "CAPITAL"]):
        return "Bất động sản"
    return "Khác"


def classify_with_source(name: str):
    """Trả (nhóm, nguồn). Ưu tiên: OVERRIDES (tay) -> VietCap ICB -> KEYWORDS.
       nguồn ∈ {'Override', 'VietCap ICB', 'Từ khóa', 'Khác'}."""
    s = (name or "").upper()
    if not s:
        return "Khác", "Khác"
    for key, grp in OVERRIDES:          # 1) quyết định gán tay (cao nhất)
        if key in s:
            return grp, "Override"
    vg = vietcap_group(name)            # 2) ICB VietCap (DN niêm yết khớp mã/tên)
    if vg:
        return vg, "VietCap ICB"
    g = _keyword_group(s)              # 3) suy theo từ khóa
    return g, ("Từ khóa" if g != "Khác" else "Khác")


def classify(name: str) -> str:
    return classify_with_source(name)[0]


def order_groups(present):
    """Trả về GROUP_ORDER đã lọc theo tập nhóm 'present' (giữ đúng thứ tự chuẩn),
    nối thêm nhóm lạ (nếu có) ở cuối trước 'Khác'."""
    present = set(present)
    out = [g for g in GROUP_ORDER if g in present]
    extra = [g for g in present if g not in GROUP_ORDER]
    if extra:
        # chèn nhóm lạ trước "Khác"
        if "Khác" in out:
            i = out.index("Khác")
            out = out[:i] + extra + out[i:]
        else:
            out += extra
    return out
