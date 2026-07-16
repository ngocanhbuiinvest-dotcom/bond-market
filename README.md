# Thống kê TPDN riêng lẻ trong nước (HNX / CBIS)

Pipeline thu thập & phân tích **kết quả chào bán trái phiếu doanh nghiệp riêng lẻ trong nước**
từ Chuyên trang TPDN của HNX: https://cbonds.hnx.vn/to-chuc-phat-hanh/thong-tin-phat-hanh

## Quy trình 3 bước

```
python bond_issuance_scraper.py     # 1) Crawl phát hành  -> bond_issuance_raw.csv / .json
python bond_buyback_scraper.py      # 2) Crawl mua lại     -> bond_buyback_raw.csv / .json
python build_reports.py             # 3) Phân tích -> Excel + dashboard.html
```

> `build_reports.py` tự động tính **Dư nợ đang lưu hành = Phát hành − Mua lại − Đáo hạn**
> nếu có `bond_buyback_raw.csv`; nếu chưa có sẽ chỉ dựng phần phát hành.

> Windows: nếu console báo lỗi Unicode, đặt trước khi chạy:
> `$env:PYTHONUTF8=1; $env:PYTHONIOENCODING="utf-8"`

## File đầu ra

| File | Nội dung |
|---|---|
| `bond_issuance_raw.csv` / `.json` | Dữ liệu thô 3.2k+ đợt phát hành (đã tách số, tính giá trị) |
| `bond_buyback_raw.csv` / `.json` | Dữ liệu thô 6.1k+ đợt mua lại trước hạn (giá trị phát hành/đang lưu hành/mua lại/còn lại) |
| `TPDN_PhatHanh_TrongNuoc.xlsx` | 9 sheet: Tổng quan, Theo năm, Theo tháng, Nhóm ngành, Kỳ hạn & Lãi suất, Top TCPH, Dữ liệu thô, **Dư nợ & Mua lại**, **Chi tiết dư nợ** (kèm biểu đồ) |
| `dashboard.html` | Dashboard tương tác, self-contained, **3 tab riêng**: (1) Phát hành mới, (2) Mua lại trước hạn, (3) Trái phiếu đang lưu hành (KPI cấu trúc như tab phát hành + maturity wall). Mỗi tab có bộ lọc + bảng sắp xếp riêng |
| `dashboard_data.json` | Dữ liệu tổng hợp dạng JSON |

Mở dashboard: nếu `file://` bị chặn thì chạy `python -m http.server 8777` rồi mở
`http://127.0.0.1:8777/dashboard.html`.

## Cơ chế lấy dữ liệu (tham khảo kỹ thuật)

- Trang render server-side, **không có API JSON/nút export**.
- Bảng nạp qua `POST /to-chuc-phat-hanh/thong-tin-phat-hanh/tim-kiem`, trả về **HTML**.
- Bắt buộc header `CP-TOKEN` = giá trị meta `__RequestVerificationToken` (lấy từ trang GET, cùng session/cookie).
- Body (jQuery-style): `searchKeys[]` = [comId, bondCode, fromDate, toDate];
  `arrCurrentPage[]` và `arrNumberRecord[]` (12 phần tử; index 0 = tab "chào bán trong nước").
- Cert SSL của site thiếu intermediate CA → scraper đặt `VERIFY_SSL=False`.

## Phân loại nhóm TCPH

Heuristic theo từ khóa trong tên doanh nghiệp: Ngân hàng, Chứng khoán, Bất động sản,
Năng lượng, Tài chính - Bảo hiểm, còn lại là "Khác". Có thể tinh chỉnh hàm `classify()`
trong `build_reports.py` (ví dụ bổ sung mapping ngành chuẩn theo mã DN).

## Mở rộng (định hướng tiếp theo)

- Thêm tab **mua lại/hoán đổi trước hạn** (đổi `DOMESTIC_TAB` + parse bảng tương ứng).
- Thêm nguồn **tin bất thường "chậm thanh toán gốc/lãi"** từ `/to-chuc-phat-hanh/tin-cong-bo`
  để theo dõi rủi ro TCPH.
- **Maturity wall**: thống kê đáo hạn theo tháng/quý từ cột `ngay_dao_han`.
- Tự động hoá cập nhật định kỳ (crawl incremental theo `Ngày đăng tin`).
