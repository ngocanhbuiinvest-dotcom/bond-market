# TPDN theo tháng — Nguồn chuẩn HNX (VBMA để cross-check)

## ⚠️ NGUYÊN TẮC SỐ LIỆU (quyết định user 15/07/2026)
**HNX = NGUỒN CHUẨN** (số thực tế, dùng snapshot cào mới nhất). **VBMA chỉ để cross-check** + cấp 2 chỉ tiêu HNX không có.
- 3 mức chính lấy từ HNX theo snapshot mới nhất: **phát hành riêng lẻ, mua lại trước hạn, GTGD thứ cấp**.
- Mua lại: dùng **TOÀN BỘ** bản ghi HNX (KHÔNG lọc theo ngày chốt). VBMA công bố mua lại thấp hơn thực tế ~40-50% vì VBMA không đính chính mua lại.
- HNX **chỉ có phát hành riêng lẻ**. Lấy từ VBMA (ghi rõ nguồn): **phát hành công chúng** (SSC) và **chậm trả** (số mã + gốc). `Tổng TPDN = HNX riêng lẻ + VBMA công chúng`.

## File dữ liệu
- **CHÍNH:** `TPDN_VBMA/TPDN_HNX_tong_hop_thang.{csv,xlsx}` — canonical HNX theo tháng, cửa sổ T7/2025–T6/2026. Đơn vị: tỷ VND.
  XLSX: sheet `HNX theo thang` (nhóm cột HNX / Bổ sung-VBMA / Cross-check) + `Ghi chu & Nguon`.
- **Cross-check chi tiết:** `TPDN_VBMA/TPDN_doi_chieu_VBMA_vs_HNX.xlsx` (4 sheet: PH RL / Mua lại / GD thứ cấp / Kết luận) + `doi_chieu_{1,2,3}_*.csv`.
- **Tham chiếu VBMA:** `TPDN_VBMA/TPDN_VBMA_tong_hop_thang.{csv,xlsx}` — số VBMA (đã đính chính), giữ để đối chiếu. **Cửa sổ T1/2021–T6/2026 (66 tháng)** — 18 tháng T1/2024→T6/2025 cào 16/07/2026 từ 19 báo cáo; 36 tháng 2021–2023 cào bổ sung cùng ngày từ 37 báo cáo (T1/2021→T1/2024), script `TPDN_VBMA/backfill_2021_2023.py` (+`_fix_xlsx.py`).
  **Phương pháp theo thời kỳ** (format báo cáo VBMA đổi): 2021 = công bố lần đầu (chưa có Phụ lục 1 TPDN); 2022 = đính chính theo cột "cùng kỳ" báo cáo 2023 (M+12; riêng T1–T2/2022 công bố lần đầu, T5/2022 dùng chênh lũy kế); 2023+ = đính chính chuẩn M+1. Mua lại có số tháng từ T1/2023 (2022 chỉ có lũy kế năm 210.830); chậm trả từ T10/2023; GTGD thứ cấp từ T3/2024. %YoY chỉ tính từ T3/2023, chặn khi gốc <500 tỷ (%YoY 2024 đã lấp nhờ gốc 2023). Mốc kiểm chứng (PL1 báo cáo T12/2023): 2022 RL 249.075/CC 21.237; 2023 RL 274.170/CC 37.070 — tổng 12 tháng của chuỗi thấp hơn mốc năm do CBTT muộn (đặc điểm dữ liệu). Kiểm chứng lũy kế 6T2025 = RL 240.664 / CC 27.904 / Tổng 268.568 (khớp mốc). Backup `_backup_20260716`, `_backup_20260716_truoc_backfill2021`.
- **Dashboard:** `dashboard/TPDN_VBMA_dashboard.html` — tự sinh từ CSV VBMA. Cập nhật: `python dashboard/build_dashboard.py` (đọc CSV → HTML). Chạy lại sau mỗi lần sửa CSV.
  Giao diện 3 tab (Tổng hợp / Riêng lẻ / Công chúng-niêm yết) + bộ lọc **phong cách Lighthouse** (navy #051A3A + gold #BB9C66, giống dashboard_ALL KQGD): thanh `.pbar` sticky viền trái gold, badge "TPDN · VBMA", nhãn "Kỳ:" + 2 dropdown: **Năm** (Tất cả năm/2021→) → **Kỳ** (Cả năm/Quý/Tháng, chỉ liệt kê kỳ có dữ liệu, khóa khi "Tất cả năm") + ghi chú kỳ đang xem. **Mặc định Toàn kỳ**, chart/KPI/bảng đổi theo bộ lọc. KPI Tổng hợp = Tổng/RL/CC/kỳ lọc; tab RL có chart số đợt + chậm trả; mỗi bảng có nút **Kết xuất Excel** (.xls theo kỳ lọc + thứ tự đang xem, số thô). Bảng mặc định **mới nhất trước**, thu gọn 12 dòng + nút "Xem thêm/Thu gọn", nút đảo chiều ↓↑; dòng Tổng kỳ lọc luôn hiện (tfoot). Ô trống trong CSV → JSON null → hiện "—" (không ép về 0).

## Dữ liệu HNX (raw, ở thư mục cha `3. Bond Market/`, scraper cập nhật hàng ngày)
- `bond_issuance_raw.csv` — phát hành **CHỈ riêng lẻ**. `ngay_phat_hanh`, `gia_tri_phat_hanh` (đơn vị **đồng**, /1e9 → tỷ), `ngay_dang_tin`.
- `bond_buyback_raw.csv` — mua lại. `ngay_mua_lai`, `gt_mua_lai_num` (đồng), `ngay_dang_tin`.
- `bond_secondary_raw.csv` — GD thứ cấp (từ 2023-07). `ngay_gd`, `gt_ty` (đã ra **tỷ**). Số phiên = số ngày GD; BQ/phiên = GTGD/số phiên.

## Kết quả cross-check HNX vs VBMA (12T, đã kiểm)
- **GD thứ cấp: khớp gần tuyệt đối** (10/12 tháng lệch 0,0%, cả số phiên & BQ/phiên) — validate pipeline. Không đính chính (tất toán trong ngày).
- **Phát hành RL:** HNX ≈ VBMA đã đính chính (sát nhau); chênh còn lại = CBTT muộn. Tháng mới nhất VBMA chưa đính chính nên lệch lớn (đó là kỳ vọng, dùng số HNX).
- **Mua lại:** HNX cao hơn VBMA ~60-110%/tháng. Lọc HNX theo `ngay_dang_tin` < ngày chốt báo cáo → khớp VBMA tuyệt đối 12/12 tháng ⇒ gap = CBTT muộn VBMA không bổ sung. **Dùng số HNX.**

## Ranh giới dữ liệu HNX (đã xác minh — KHÔNG phải lỗi scraper)
- **Q1/2025 phát hành riêng lẻ gần như bằng 0 là THẬT** (đóng băng hậu Tết + siết pháp lý): cả quý chỉ 2 mã (VDS 04/03, ACB 26/03); T1–T2/2025 = 0 mã. Bằng chứng: HNX H1‑2025 RL = 248.664 tỷ vs VBMA 6T2025 = 240.664 tỷ (+3,3%, đúng kiểu HNX nhỉnh hơn do CBTT muộn) ⇒ HNX đủ ở cấp tổng.
- Hệ quả: `%YoY` RL các tháng **T1/T2/T3/2026 để trống** vì **gốc 2025 ≈ 0** (chia cho ~0), không phải thiếu dữ liệu. `%MoM` chặn khi gốc <500 tỷ (tránh tỷ lệ vô nghĩa).
- `%MoM/%YoY` tính từ chuỗi HNX-vs-HNX (nhất quán). %YoY 2024 base khỏe (40-78 đợt/tháng), tin cậy.

## Quy trình cập nhật khi có snapshot HNX / báo cáo VBMA mới
1. Thêm/ghi đè tháng mới nhất từ HNX (3 mức chính, snapshot mới nhất).
2. Các tháng cũ tự "đính chính" vì luôn dùng snapshot mới nhất (HNX bắt trọn CBTT muộn).
3. Bổ sung công chúng + chậm trả từ VBMA (báo cáo tháng); tính lại Tổng TPDN, %MoM/%YoY.
4. Backup trước khi ghi. Đối chiếu lại cross-check để chắc pipeline không lệch.

## Nguồn VBMA (cho công chúng + chậm trả + cross-check)
Báo cáo tháng: https://vbma.org.vn/vi/reports/monthly (mới nhất trên cùng).
Link PDF: `/storage/reports/<Month><Year>/VBMA_BAO CAO TTTP THANG <m> <yyyy>.pdf` — thư mục theo **tháng công bố** (= tháng sau tháng báo cáo).
Số phát hành/GTGD ở **Phụ lục 1**; số đợt/mua lại/chậm trả ở trang 'Trái phiếu doanh nghiệp'.
Đính chính VBMA (tham khảo): giá trị phát hành tháng M = cột "tháng trước" trong Phụ lục 1 báo cáo M+1. Kiểm chứng tổng năm với biểu đồ trang 2 (2025 = 627.810 tỷ).
Base lũy kế 6T2025 (VBMA): RL 240.664 / CC 27.904 / Tổng 268.568.

## Tải PDF VBMA (SSL chặn tải trực tiếp)
WebFetch/curl/PowerShell đều lỗi SSL. Cách dùng được: mở vbma.org.vn bằng in-app browser →
`javascript_tool` fetch PDF same-origin → arrayBuffer → trích text bằng pdf.js import động từ cdnjs.
Tìm trang Phụ lục 1: normalize NFD + xoá dấu combining + collapse whitespace rồi match `/Phu luc 1/i`.
Chỉ trả text đã trích, đừng chuyển cả PDF (~700KB) qua context.
