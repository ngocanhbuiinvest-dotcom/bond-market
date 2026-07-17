# STATUS — Dự án TPDN riêng lẻ trong nước (bàn giao)

Cập nhật: 16/07/2026. Đọc file này trước khi làm tiếp.

## Mục tiêu
Thu thập & phân tích thị trường **TPDN riêng lẻ trong nước** từ Chuyên trang TPDN của HNX
(https://cbonds.hnx.vn/to-chuc-phat-hanh/thong-tin-phat-hanh): phát hành, mua lại trước hạn,
dư nợ đang lưu hành, giao dịch thứ cấp, **xếp hạng tín nhiệm**, **chậm trả gốc/lãi**.
(Chưa làm: hoán đổi trái phiếu, tự động cập nhật định kỳ.)

## Quy trình (Windows: đặt `PYTHONUTF8=1; PYTHONIOENCODING=utf-8`)
```
python bond_issuance_scraper.py   # phát hành     -> bond_issuance_raw.csv/json  (~3.2k đợt)
python bond_buyback_scraper.py    # mua lại       -> bond_buyback_raw.csv/json   (~6.1k đợt)
python bond_secondary_scraper.py  # GD thứ cấp    -> bond_secondary_raw.csv/json (~70k mã×ngày, THEO NGÀY)
python bond_catalog_scraper.py    # danh mục ĐKGD -> bond_catalog_raw.csv/json  (2,171 mã: crosswalk mã)
python bond_rating_scraper.py     # xếp hạng TN   -> bond_rating_raw.csv/json   (~148 kết quả XHTN)
python bond_latepay_scraper.py    # tin bất thường-> bond_latepay_raw.csv/json  (~2,2k tin; ~411 chậm trả)
python vietcap_sector_scraper.py  # ngành ICB     -> vietcap_companies_raw.csv/json (~2,088 DN niêm yết + ICB)
python build_reports.py           # -> Excel 9 sheet + dashboard.html (7 tab) + dashboard_data.json
# Chỉ dựng lại dashboard (bỏ Excel, nhẹ): python _rebuild_sec.py   (đã gồm XHTN + chậm trả)
# TỰ ĐỘNG CẬP NHẬT HÀNG NGÀY (làm tất cả trên + diff + rebuild): python update_daily.py
```

## Kỹ thuật crawl (đã reverse-engineer)
- `POST /to-chuc-phat-hanh/thong-tin-phat-hanh/tim-kiem` → trả **HTML** (không phải JSON).
- Header bắt buộc: `CP-TOKEN` = nội dung meta `__RequestVerificationToken` (lấy từ trang GET, cùng session/cookie).
- Body (jQuery-style): `searchKeys[]`=[comId,bondCode,fromDate,toDate];
  `arrCurrentPage[]` & `arrNumberRecord[]` = 12 phần tử; **index tab**: 0=phát hành (`tbReleaseResult`),
  1=mua lại (`tbRepurchaseResult`).
- `VERIFY_SSL=False` (cert site thiếu intermediate CA).
- Tài liệu đính kèm: `GET /view-file?refId=<file_id>&tableType=2`.

## Quyết định nghiệp vụ đã chốt
- **Giá trị phát hành = khối lượng × mệnh giá**.
- **Gộp trùng** đợt phát hành theo (mã TP + ngày phát hành), giữ bản đăng tin mới nhất → 3,108 đợt / 3,087 mã.
- **"Hết hiệu lực"** = bản CBTT kết quả bị vô hiệu (nghi hủy/thay thế). ĐÃ verify: tài liệu là bản
  kết quả phát hành thật, KHÔNG do đáo hạn (nhóm Hiệu lực cũng nhiều mã đã quá hạn). → **Giữ trong
  gross issuance + gắn cờ** (KPI + cột Tình trạng + bộ lọc). Không tự dán nhãn "bị hủy".
- **Dư nợ đang lưu hành = Phát hành − Mua lại trước hạn − Đáo hạn**, cấp mã TP: mã có mua lại → lấy
  "giá trị còn lại sau mua lại" của đợt mới nhất; chưa mua lại → giá trị phát hành; đã qua đáo hạn → 0.
- Số hiển thị **chuẩn quốc tế en-US** (phẩy = hàng nghìn, chấm = thập phân). Excel giữ **số thực**
  (hiển thị dấu tùy Regional Settings của Windows).

## FIX quan trọng 15/07 (universe dư nợ)
`compute_outstanding` trước chỉ dựng universe từ bảng phát hành (dữ liệu từ 30/12/2020) rồi LEFT-join
mua lại → ~90 mã phát hành 2019–2020 còn hạn bị mất khỏi tab Đang lưu hành (vd HAGLBOND16.26,
CII012029-G). Đã sửa: **universe = HỢP (phát hành ∪ mua lại)**; mã chỉ có ở mua lại dựng từ
`bond_buyback_raw.csv`; ngày đáo hạn = MAX(dh phát hành, dh mua lại); thêm cột **Nguồn**
(Phát hành / CBTT mua lại). Dư nợ hiện ≈ **1,137,366 tỷ / ~1,280 mã**; maturity wall đỉnh 2028.

## Con số hiện tại (tham chiếu)
Gross phát hành 2,531,167 tỷ · Mua lại (toàn bộ) 1,324,995 tỷ · Dư nợ đang lưu hành 1,137,366 tỷ.
Cơ cấu phát hành (taxonomy 13 nhóm, 15/07 v4): Ngân hàng 51.1%, BĐS 25.9%, Khác 2.8%, còn lại nhỏ.
Coupon BQ phát hành ~7.9%, đang lưu hành ~8.7%. (BĐS tăng từ ~14%→25.9% vì classifier mới + mặc định
"Đầu tư/Phát triển/Tập đoàn/Invest/Capital"=BĐS + gộp Vingroup/Sovico/IPA vào BĐS.)

## Phân loại NGÀNH — 13 nhóm (sector_map.py, 15/07 v4)
Trước: 6 nhóm, "Khác" ~53% giá trị GD thứ cấp (nhiều DN lớn bị bỏ sót). Nay `sector_map.py` (dùng chung
build_reports/_rebuild_sec) phân **14 nhóm** bằng 2 lớp: (1) **OVERRIDES** — gán tay theo chuỗi con tên
DN lớn (Vietjet→Hàng không, VinFast→Sản xuất, Nam Long/Becamex/Novaland/Hưng Thịnh→BĐS, HAGL/TTC/Masan→
Nông nghiệp-TP, Núi Pháo→Khai khoáng, Tasco/Vinaconex→Xây dựng, Trung Nam→Năng lượng, Vinpearl/Crystal
Bay/Bà Nà→Du lịch, Vingroup/Sovico/IPA→BĐS, F88→Tài chính); (2) **KEYWORDS** mở rộng (thứ tự: ngành
cụ thể trước BĐS/thương mại/đa ngành). Ranh giới: tên có BĐS/địa ốc/nhà/đô thị/land/homes→BĐS; chỉ có
xây dựng/hạ tầng→Xây dựng. **MẶC ĐỊNH (user 15/07 v4): tên thuần "Đầu tư/Phát triển/Tập đoàn/Invest/
Capital" không lộ ngành khác → BẤT ĐỘNG SẢN**. **Vingroup/Sovico/IPA cũng gộp vào BĐS** (user chốt) →
BỎ hẳn nhóm "Đầu tư-Đa ngành", còn 13 nhóm. `GROUP_ORDER` = thứ tự chuẩn (màu ổn định);
`order_groups(present)` hợp nhất nhóm 3 nguồn. **Khác còn 2.8% (PH) / 0.8% (GD)** — chỉ còn DN thật sự
khác ngành/quá mờ (CMC, 3C=CNTT, AAC, Vipico...); bóc thêm chỉ cần thêm dòng OVERRIDES trong
sector_map.py. COLORS template 8→14 màu. GD thứ cấp theo nhóm: NH 39.5% · BĐS 34.7% · Du lịch 6.1% · TM-DV 4.0%.

## Xếp hạng tín nhiệm & Chậm trả gốc/lãi (THÊM 16/07/2026)
Hai scraper + hai tab dashboard mới, tái dùng khung sẵn có.
- **XHTN** (`bond_rating_scraper.py`): `POST /danh-sach-thong-tin-xep-hang-tin-nhiem/danh-sach`,
  `keySearch='||||'` (5 trường: objecttype|issuer|bondcode|tradingcode|TrustRating), arrCurrentPage/
  arrNumberRecord (≤100/trang), parse `#tbOrgDeposit`. 148 kết quả · 109 TCPH · 6 đơn vị (Saigon
  Ratings/FiinRatings/VIS/S&I/Moody's/Thiên Minh). Hạng lẫn thang nội địa `vnXXX` + quốc tế + Moody's;
  file đính kèm `ViewFile('<refId>','800')`. comId TCPH từ `view_TT_TCPH('<id>')`. Ordering hạng: JS
  `ratingRank()` bỏ tiền tố vn, quy Moody's→S&P; `isInvestmentGrade` = ≥ BBB-.
- **Chậm trả** (`bond_latepay_scraper.py`): `POST /to-chuc-phat-hanh/tin-cong-bo-x`, body dạng
  `keysSearch[]`(7: title,from,to,bondCode,comId,articleType,loaihinh='0') + `currentPages[]`(4) +
  `numberRecord[]`(4). **INDEX 2 = tab "Tin bất thường"** (`#tbInconstant`) — quét TOÀN BỘ 2,221 tin
  rồi PHÂN LOẠI theo từ khóa tiêu đề (`classify_event`): `cham_tra` (chậm thanh toán/chậm trả/không thể
  thanh toán) vs `khac_phuc` (…sau thời gian bị chậm…). 411 lượt CBTT chậm trả · 116 TCPH · 246 mã TP ·
  3 khắc phục. **Chỉ có SỐ LƯỢT/số mã, KHÔNG có số tiền chậm trả** (nằm trong file PDF đính kèm) — UI ghi rõ.
- Payload JSON: `rating`={rows:[{dv,dvfull,**doi**,dn,ma,kq,hl,loai,nhom,file}],kpi{n,n_tcph,n_org,n_tp,n_dv}}
  (`doi`='org'|'tp' = đối tượng XHTN: **Tổ chức phát hành 145 vs Trái phiếu 3**) · `latepay`={rows:[{d,dn,ma,td,
  loai,nhom,file}],kpi}. `build_json(df,out,bb,sec,rating,latepay)`; loaders `load_rating()`/`load_latepay()`;
  `_rebuild_sec.py` đã gọi cả hai. Chart mới: `barCount()` (cột đếm, nhãn x xoay 40°).
- Tab **XHTN**: KPI + bar phân bố hạng + hbar thị phần đơn vị XHTN + hbar TCPH/ngành + **HAI bảng riêng**
  (① Xếp hạng Tổ chức phát hành — không cột Mã TP · ② Xếp hạng Trái phiếu — có Mã TP), search `fr_q` áp cả 2.
- Tab **Chậm trả**: KPI (6) + bar số lượt theo thời gian (Tháng/Quý/Năm + Phân ngành) + hbar theo ngành + top TCPH,
  rồi **DANH SÁCH CHẬM TRẢ gộp cấp mã TP** (`buildList` nổ mã theo dấu phẩy → 247 mã: Mã TP·TCPH·Nhóm·**Loại chậm
  (Gốc/Lãi/Cả hai)**·**Dư nợ chậm trả (tỷ)**·Số lượt CBTT·Chậm lần đầu·CBTT gần nhất·**Thời gian chậm (năm)**·Trạng
  thái). **Thời gian chậm** = từ chậm lần đầu → `new Date()` (mã còn chậm) / ngày khắc phục (mã đã khắc phục),
  quy năm (`yearsBetween`, /365.25). **CHỈ 1 bảng**
  (16/07 v3: BỎ bảng feed từng CBTT cho gọn). search `fl_q`. CSV: DanhSachChamTra. Tab **Tổng hợp** thêm **KPI thứ 5
  "Dư nợ chậm trả"** (#ks 4→5 cột; snapshot, chỉ theo lọc Nhóm; ≈101 nghìn tỷ = 8.9% dư nợ; đọc `DATA.latepay.dno`).
- **DƯ NỢ CHẬM TRẢ (16/07 v2, user chốt KHÔNG OCR file)**: khảo sát 20 file → chỉ ~20% PDF có text, 80% scan →
  OCR không đáng tin. Số tiền trong file là dòng tiền 1 kỳ, KHÔNG phải dư nợ. → **Dư nợ chậm trả = dư nợ gốc còn
  lại của mã** (giá trị phát hành − mua lại, KHÔNG trừ đáo hạn vì mã chậm chưa trả gốc), join từ dữ liệu phát
  hành+mua lại đã có: `outstanding_by_bond(df,bb)` → dict {ma:đồng}; khớp **212/246 mã** (34 mã cũ ngoài phạm vi
  = "n/a"). **Loại chậm gốc/lãi/cả hai bóc từ TIÊU ĐỀ** (`_late_type`: 225 cả hai·166 lãi·19 gốc). KPI: **Dư nợ
  chậm trả ≈ 101 nghìn tỷ = 8.9% tổng dư nợ** · **Dư nợ chậm GỐC (gốc/cả hai) ≈ 71 nghìn tỷ = 6.3%** (mẫu số =
  `out.outstanding_ty`). Payload `latepay.dno`={ma:tỷ}, rows thêm `lct`; KPI tính LẠI theo bộ lọc trong JS.
  `makeTable` thêm option `w` (độ rộng cột). Nâng cấp tương lai: OCR file lấy số tiền kỳ chậm (nếu cần).

## Phân nhóm ngành — LỚP VIETCAP ICB (THÊM 16/07/2026 v3, user chốt "override cho mã khớp")
`sector_map.py` giờ 3 lớp ưu tiên: **OVERRIDES (gán tay) → VietCap ICB → KEYWORDS**. `vietcap_sector_scraper.py`
gọi **API công khai VietCap** `iq.vietcap.com.vn/api/iq-insight-service` (không auth): `/v2/company/search-bar?language=1`
(2,088 DN niêm yết + ICB Lv1..Lv4) + `/v1/sectors/icb-codes` (177 mã ngành). TCPH khớp bằng **tiền tố mã CK**
trong tên ("ACB - Ngân hàng…" → code ACB) hoặc tên chuẩn hoá → lấy ngành ICB, map về 13 nhóm qua `ICB2_GROUP`
(key = ICB Lv2) + tinh chỉnh Lv3/Lv4 (Dịch vụ tài chính+chứng khoán→Chứng khoán · Hàng&DV Công nghiệp+vận tải→
Hàng không-Vận tải · Tài nguyên Cơ bản+khoáng→Khai khoáng). **Độ phủ: ~25% số TCPH nhưng ~58% GIÁ TRỊ** (SPV chưa
niêm yết không có → về KEYWORDS). Nguồn phân ngành theo giá trị: VietCap 1,476k tỷ · Từ khóa 555k · Override 433k ·
Khác 67k (2.6%). Sửa nhiều lỗi (BSI→Chứng khoán, Becamex/BCM→BĐS qua mã CK, Camimex→Nông nghiệp-TP, Transimex→
Vận tải). `classify_with_source()` trả (nhóm, nguồn); `load()` thêm `df.nhom_src`; rows payload thêm `nghn`; tab
**Phát hành** thêm cột **"Nguồn ngành"** (pill: VietCap ICB xanh ngọc · Override tím · Từ khóa). Chi tiết API: memory `vietcap-icb-api`.

## Tự động cập nhật hàng ngày (THÊM 16/07/2026 v4) — xử lý "data cũ ĐỔI TRẠNG THÁI"
**Vấn đề**: cập nhật incremental-theo-ngày sẽ SÓT thay đổi trên bản ghi cũ (vd Hiệu lực→Hết hiệu lực,
đang chậm→đã khắc phục, trạng thái ĐKGD). **Phát hiện**: mọi nguồn hay đổi đều RẺ để quét lại toàn bộ;
nguồn đắt duy nhất (GD thứ cấp ~28') lại BẤT BIẾN với quá khứ. → `update_daily.py` (orchestrator):
- Nguồn hay đổi + rẻ (phát hành/mua lại/danh mục/XHTN/chậm trả/ICB) → **full re-scrape** (subprocess),
  dashboard tính lại mọi trạng thái dẫn xuất → thay đổi tự động được bắt.
- **GD thứ cấp → incremental**: `bond_secondary_scraper.py --incremental` (đọc CSV cũ, quét lại từ ngày
  cuối − `OVERLAP_DAYS=7` chồng lấn, merge). Chỉ ~24s thay vì ~28'. Đã test: +121 phiên (14→16/07).
- **GUARD chống mất data**: sao lưu `*.csv.prev`; nếu bản mới < `GUARD_RATIO=0.85`×bản cũ → khôi phục
  bản cũ (nghi quét lỗi mạng), đánh dấu guard.
- **SNAPSHOT-DIFF → `changes_log.csv`**: so bản cũ↔mới theo key+cột trạng thái (SOURCES trong update_daily),
  ghi: Mới / **Đổi trạng thái** (field, cũ→mới) / Biến mất. Test đã bắt LPB12604 ĐKGD "Chờ giao dịch→Đăng
  ký giao dịch". `update_state.json` = trạng thái lần chạy (đếm + delta + guard). LƯU Ý: ghi state TRƯỚC rebuild.
- Lịch: `Dang_ky_lich_TPDN.ps1` đăng ký Task Scheduler `TPDN_CapNhatHangNgay` chạy `run_update.bat`
  **16:30 các NGÀY LÀM VIỆC (T2–T6)**, StartWhenAvailable (chạy bù nếu máy tắt). **ĐÃ ĐĂNG KÝ 16/07/2026**
  (State=Ready). Log: `update_log.txt`. LƯU Ý encoding: PowerShell 5.1 đọc `.ps1` KHÔNG BOM theo codepage
  ANSI → tiếng Việt thành mojibake → parse lỗi; file đã lưu lại **UTF-8 kèm BOM** (parse 0 lỗi). Nếu cần
  đăng ký lại từ máy khác: chạy PS1, hoặc `Register-ScheduledTask` inline. Gỡ: `Unregister-ScheduledTask
  -TaskName "TPDN_CapNhatHangNgay" -Confirm:$false`. Chạy thử: `Start-ScheduledTask -TaskName "TPDN_CapNhatHangNgay"`.
- **CẢNH BÁO "mã chậm trả MỚI"** (16/07 v5): `detect_new_latepay()` so tập mã đang chậm bản cũ↔mới, trả mã
  LẦN ĐẦU xuất hiện trong nhóm chậm trả. Khi có → (1) ghi lịch sử `alerts_new_latepay.csv`
  (run_date·ma_tp·ten_dn·ngay_cbtt·tieu_de), (2) **balloon khay hệ thống** qua `notify_alert.ps1`
  (System.Windows.Forms.NotifyIcon, best-effort — task nền có thể không hiện), (3) nhúng vào
  `update_state.json` (`n_new_latepay`, `new_latepay[]`). Đã test bằng 2 mã giả (TESTCH2401/2402) rồi khôi
  phục sạch (n_new_latepay=0).
- Dashboard **tab thứ 8 "Cập nhật & thay đổi"**: badge lần chạy cuối (`... · ⚠ N mã chậm trả mới`) + 7 thẻ
  đếm/delta mỗi nguồn + **bảng nhật ký thay đổi** (lọc nguồn/loại/search, xuất CSV). **Tab "Chậm trả" có
  banner cảnh báo** liệt kê mã chậm trả mới của lần cập nhật gần nhất (ẩn khi không có). Payload
  `updates`={state, changes[]} (`load_updates()`); banner/badge đọc `state.n_new_latepay`/`new_latepay`.

## Dashboard (dashboard.html, self-contained, mở offline; hoặc `python -m http.server`)
8 tab: **Tổng hợp · Phát hành mới · Mua lại trước hạn · Trái phiếu đang lưu hành · Giao dịch thứ cấp ·
Xếp hạng tín nhiệm · Chậm trả gốc/lãi · Cập nhật & thay đổi**. Mỗi tab có
bộ lọc riêng (năm/quý/tháng/nhóm/tình trạng) + KPI + biểu đồ + bảng chi tiết sắp xếp được + nút Xuất CSV.
KPI tab "Đang lưu hành" cấu trúc song song tab "Phát hành".
- Tab **Tổng hợp** (tinh chỉnh 15/07 v2 — charts-first): **4 KPI** (Phát hành · Mua lại · Ròng · Dư nợ;
  số đợt/số TCPH gộp dòng phụ, `#ks` = grid 4 cột) → **col12 biểu đồ CỘT hợp nhất "Phát hành · Mua lại ·
  Ròng theo thời gian"** (hàm `phmlCombo`: cột đôi PH xanh dương / ML xanh lá trên mốc 0 + đường cam =
  ròng; **GỘP** biểu đồ theo tháng & theo năm cũ, có điều khiển **Đơn vị thời gian** Ngày/Tuần/Tháng/Quý/Năm
  `cs_gran` + **Phân ngành** multi-select `cs_cgroup`; aggregator `aggDual(riRows,rbRows,gran)` +
  `periodKey`) → hàng col4 **donut dư nợ** + col8 **bảng tăng trưởng MoM/QoQ/YoY** (có bộ lọc **Đơn vị**
  `gr_unit` Tháng/Quý/Năm + **Kỳ** `gr_period`; mặc định = kỳ hoàn chỉnh gần nhất < data_date; unit=quý→MoM
  "–", unit=năm→MoM/QoQ "–"; helpers `serY`/`qAdd`/`mToQ`/`fillGrowthPeriod`) → col12 **Khoảng thời gian
  tùy chọn** (Từ–Đến ngày, 4 tiles). Combo & bảng chỉ thêm lọc riêng của card; vẫn kế thừa lọc trên
  cùng (năm/quý/tháng/nhóm). **donut** (hàm `donut`): giá trị (nghìn tỷ) + % hiện TRÊN biểu đồ (nhãn ngoài
  + leader line, lát ≥4%), legend chỉ còn tên ngành — áp dụng cho MỌI donut (dư nợ/phát hành/mua lại/thứ cấp).
  (Hàm cũ `dualSeries`/`dualBars`/`dualBarsNet`/`barSigned` còn định nghĩa nhưng không còn dùng ở tab này.)
- Thay "Top TCPH phát hành/mua lại" bằng: tab Phát hành = **lãi suất phát hành BQ theo ngành qua các năm**
  (multi-line %); tab Mua lại = **giá trị mua lại theo ngành qua các năm** (multi-line). Hàm `multiLine()`.
- rows/bb_rows có thêm trường `kl` (khối lượng = số trái phiếu) phục vụ thống kê khối lượng.
- Tab Phát hành & Đang lưu hành: biểu đồ **combo gộp** "Giá trị/Dư nợ & lãi suất" với 2 điều khiển
  riêng trên card: **Đơn vị thời gian (Ngày/Tuần/Tháng/Quý/Năm)** + **Phân ngành (multi-select checkbox)**.
  Cột cam = giá trị (tỷ VNĐ, trục trái rút gọn "k"), đường xanh = lãi suất BQ gia quyền (%, trục phải).
  Hàm: `aggCombo(rows,gran,dateOf→Date)` + `combo()` + component `multiSelect()` trong dashboard_template.html.
  Giá trị vẽ theo **tỷ** (không chia 1000) để đơn vị Ngày/Tuần vẫn có nghĩa.
- Tab **Giao dịch thứ cấp** (MỚI 15/07): thống kê GD trên thị trường thứ cấp HNX (bond_secondary_raw.csv).
  Lọc **Năm GD + Nhóm TCPH + search**. 6 KPI (tổng giá trị, BQ/phiên, phiên cao nhất, tổng KL, số mã
  từng GD, số phiên). Biểu đồ col12 **combo `secCombo`**: cột cam = giá trị GD (tỷ) · đường xanh = **số mã
  GD bình quân/phiên** (tỷ, đổi 15/07 v2 từ "số mã GD"; trục phải = BQ/phiên tỷ) — có Đơn vị thời gian +
  Phân ngành như tab Phát hành, aggregator riêng `secAgg` (thêm distinct `days`/kỳ → `bq=v/nd`).
  Donut giá trị theo nhóm + hbar Top 15 mã + bảng gộp cấp mã (`byBond`: **Mã giao dịch · Mã CBTT** ·
  TCPH · nhóm · giá trị · KL · số phiên GD · giá BQ = gt/kl đồng/TP). Payload: `sec.rows`=[{d,ma,gt(tỷ),kl}]
  70,794 dòng + `sec.meta`={ma:[tên DN,nhóm,**mã CBTT**]} + `sec.kpi`. `dashboard.html` ~7.5MB (mở offline OK).

## MAPPING mã giao dịch ↔ mã phát hành (FIX 15/07 v3) — QUAN TRỌNG
**Mã giao dịch thứ cấp ≠ mã phát hành/CBTT.** GD thứ cấp dùng mã niêm yết HNX (vd `VJC12101`,`VDI12101`),
còn phát hành/mua lại dùng mã CBTT (vd `VJCH2101`,`VDICH2128001`). Join trực tiếp `ma_tp` chỉ khớp 239/929
mã → 690 mã thành "(không rõ)". **Nguồn khóa nối chuẩn = Danh mục ĐKGD** `https://cbonds.hnx.vn/danh-muc-trai-phieu`
(`bond_catalog_scraper.py` → `bond_catalog_raw.csv`, 2,171 mã), có đủ **Mã CBTT · Mã giao dịch · ISIN · Tên
TCPH · trạng thái ĐKGD · ngày GD đầu/cuối**. `load_secondary` map `ma_gd → ten_tcph/ma_cbtt` (phủ **100%**
929 mã, 0 mã "(không rõ)"), nhóm ngành do `classify(ten_tcph)`. Catalog `ma_cbtt` khớp 2,165/2,171 mã phát
hành → cũng bắc cầu sang dữ liệu phát hành nếu cần. Phân bổ GD: NH 1,389 + Khác 1,352 + BĐS 650 nghìn tỷ
(nhóm "Khác" lớn do classifier chỉ nhận diện theo từ khoá tên — DN sản xuất/thương mại rơi vào Khác, đúng
hành vi như các tab khác, không phải lỗi map).

## Giao dịch thứ cấp — kỹ thuật crawl (bond_secondary_scraper.py)
Nguồn: `POST /thong-ke-thi-truong/danh-sach`, tab "theo Mã TP" (index 1 trong arrCurrentPage/arrNumberRecord).
`keySearch` = 10 trường ngăn '|', **2 trường cuối = ngày GD từ/đến**. **BÀI HỌC**: bộ lọc ngày CHỈ đúng khi
`from==to` (một ngày); truyền KHOẢNG (from≠to) → server trả snapshot ngày mới nhất → scraper cũ (cửa sổ
tháng) gộp nhầm về đúng 1 ngày. → **quét THEO TỪNG NGÀY** (bỏ T7/CN; ngày lễ = total 0 → bỏ). Mỗi ngày
server liệt kê TOÀN BỘ mã (kl=0 nếu không khớp), sắp xếp **giá trị GIẢM DẦN** → chỉ phân trang tới khi gặp
dòng giá trị 0 thì dừng. Data từ **19/07/2023** (ngày mở thị trường). Chạy ~1690s. 740 phiên, 929 mã thực GD.

## Con số GD thứ cấp (19/07/2023 – 14/07/2026)
Tổng giá trị GD **3,518,268 tỷ** · 740 phiên · BQ **4,754 tỷ/phiên** · phiên cao nhất 20,333 tỷ (29/06/2026).
Thanh khoản tăng đều theo năm: 2023 (½ năm) 1,948 → 2024 4,337 → 2025 5,608 → 2026 6,354 tỷ/phiên BQ.

## File chính
`bond_issuance_scraper.py`, `bond_buyback_scraper.py`, `bond_secondary_scraper.py`,
`bond_catalog_scraper.py` (crosswalk mã), `bond_rating_scraper.py` (XHTN),
`bond_latepay_scraper.py` (tin bất thường / chậm trả), `vietcap_sector_scraper.py` (ngành ICB VietCap),
`sector_map.py` (phân loại 13 ngành, 3 lớp OVERRIDES→VietCap ICB→từ khóa — dùng chung),
`update_daily.py` (orchestrator tự động cập nhật + diff + cảnh báo chậm trả mới + rebuild),
`run_update.bat` + `Dang_ky_lich_TPDN.ps1` (lịch 16:30 T2–T6), `notify_alert.ps1` (balloon cảnh báo desktop),
`changes_log.csv` (nhật ký thay đổi) + `alerts_new_latepay.csv` (lịch sử mã chậm trả mới) + `update_state.json` (trạng thái lần chạy),
`build_reports.py`, `_rebuild_sec.py` (dựng nhanh dashboard),
`dashboard_template.html` (nguồn — sửa ở đây rồi chạy build_reports để nhúng data ra `dashboard.html`),
`TPDN_PhatHanh_TrongNuoc.xlsx`, `README.md`.

## Hướng tiếp theo (user quan tâm)
1. ~~Chậm trả gốc/lãi + dư nợ chậm trả~~ ✅ ĐÃ LÀM 16/07 v2 (loại chậm từ tiêu đề · dư nợ chậm join dữ liệu ·
   % trên tổng dư nợ). Nâng cấp CHỈ nếu user cần số tiền kỳ chậm chính xác: OCR file PDF scan (80% file là
   scan → view-file `owa.hnx.vn/ftp/.../*.pdf`; refId từ `view-file?refId=<id>&tableType=3`). User đã chọn
   KHÔNG OCR (16/07). File thật: GET `/view-file` trả HTML bảng file → link `owa.hnx.vn/ftp/CBIS/ArticleFileAttach/*.pdf`.
2. ~~Tự động cập nhật định kỳ~~ ✅ ĐÃ LÀM 16/07 v4 (`update_daily.py` + lịch + nhật ký thay đổi); v5: **lịch đã
   đăng ký 16:30 T2–T6** + **cảnh báo mã chậm trả mới** (banner dashboard + balloon desktop + `alerts_new_latepay.csv`).
   Còn: hoán đổi trái phiếu (menu có "CBTT trước/kết quả hoán đổi TP"); cảnh báo **qua email** khi có mã chậm trả
   mới (hiện mới có balloon desktop — best-effort, task chạy nền có thể không hiện).
3. **Khai thác VBMA vào tab Tổng hợp** (đã phân tích, user hoãn 15/07): VBMA có 4 chiều HNX-scraper
   thiếu — ưu tiên (a) **chậm trả** (số mã + gốc chậm trả theo tháng, `VBMA/TPDN_VBMA/..._thang.csv`
   cột `So ma cham tra`/`Goc cham tra`), (b) **GTGD thứ cấp/thanh khoản** (`GTGD thu cap`/`BQ/phien`);
   phụ: (c) phát hành công chúng RL vs CC, (d) đối chiếu số đính chính HNX↔VBMA. Cách làm: đọc CSV VBMA
   trong `build_reports.py` → nhúng vào `dashboard_data.json` → thêm dải "Bối cảnh thị trường (VBMA)"
   cuối tab Tổng hợp + link sang `VBMA/dashboard/TPDN_VBMA_dashboard.html`.
   (Lưu ý: scraper thứ cấp HNX `bond_secondary_scraper.py` mới chỉ chạy thử 1 tháng 06/2026 → chưa dùng được.)

## Môi trường (15/07)
Ổ **C: đầy 100%** → `build_reports.py` lỗi `OSError: No space left` khi openpyxl ghi Excel (dùng TMP trên C:).
Dữ liệu không đổi nên chỉ cần rebuild `dashboard.html` (nhúng thẳng `dashboard_data.json` vào template,
ghi ra E:, không qua TMP). Muốn chạy lại full pipeline/Excel: dọn ổ C: hoặc đặt `TMP=E:\...\tmp` trước khi chạy.


---

# NGUỒN THỨ 2: VSD (VSDC) — ĐỐI CHIẾU CHÉO VỚI HNX (17/07/2026)

## Scraper `vsd_bond_scraper.py` -> `vsd_bond_raw.csv` / `.json`
Nguồn: `https://www.vsd.vn/vi/ibl` ("Công cụ nợ và trái phiếu doanh nghiệp"). **2.332 mã TPDN riêng lẻ**.
Cơ chế GIỐNG pattern HNX nhưng khác tên header: GET `/vi/ibl` -> `meta[name="__VPToken"]`
-> POST `/vi/ibl/search` (JSON) kèm header `__VPToken`. Trả HTML mảnh. `VERIFY_SSL=True` (VSD cert OK, khác HNX).
Body: `{SearchKey, CurrentPage, RecordOnPage, OrderBy, OrderType}`; SearchKey = **11 trường** ngăn `|`:
`issuerId|code|stockType|sanGiaoDich|status|lang|kyHan|namPhatHanh|namDaoHan|tenTCPH|typeIsuStock`
- `stockType` cố định `"2,4"` (theo JS trang); `lang=VI`; **`typeIsuStock=7` = TPDN riêng lẻ** (1=TPCP, 5=TPDN ra công chúng...).
- **BẪY: server BỎ QUA `RecordOnPage`>10** -> luôn 10 dòng/trang -> 234 trang (đừng phí công thử page size lớn).

2 giai đoạn: (1) LIST 234 trang; (2) **DETAIL** `GET /vi/s-detail/<id>` từng mã, 8 luồng, ~12 phút.
Detail cho thứ list KHÔNG có: **mệnh giá, tổng giá trị ĐK, lãi suất, cách trả lãi, Giấy CNĐKCK, và mã CBTT**.
Chạy `--no-detail` để bỏ giai đoạn 2 (nhanh nhưng mất mệnh giá/giá trị).

**Bóc mã CBTT** (`extract_ma_cbtt`, hàm dùng chung — đừng viết lại inline): tên CK VSD có 4 dạng:
`... (BIDLH2331010)` trong ngoặc | `Trái phiếu VHMB2426004` trần | `Trái phiếu STA12601` = **chỉ mã GD, KHÔNG phải CBTT**
| chỉ tên DN (không mã). Được **1.596/2.332** mã có mã CBTT riêng.

**KIỂM CHỨNG PARSER**: chân trang VSD tự công bố tổng -> `parse_summary()` lưu vào JSON.
Đã đối chiếu KHỚP TUYỆT ĐỐI: 2.332 mã / 1.595 hiệu lực / 737 hủy / 2.257.150.836 TP.
⚠ Nhãn VSD ghi "Tổng số lượng chứng khoán **đang lưu hành**" nhưng thực chất = tổng **TẤT CẢ** (gồm cả mã đã hủy);
chỉ mã hiệu lực = 742.759.208. Đừng trích nhãn này ra báo cáo.

## Đối chiếu `vsd_vs_hnx_compare.py` -> `VSD_vs_HNX_DoiChieu.xlsx` (8 sheet)
Hai nguồn ĐO HAI THỨ KHÁC NHAU -> mục tiêu là **giải thích** chênh, không ép khớp 100%:
VSD = **đăng ký lưu ký** tại VSDC | HNX catalog = **đăng ký giao dịch** sàn TPRL | HNX issuance = CBTT phát hành.

Khóa nối: **ISIN -> mã CBTT -> mã GD** (2.168/2.332 khớp, 100% qua ISIN).
⚠ **BẪY ĐÃ MẮC PHẢI**: TP phát hành gần đây có **mã CBTT TRÙNG mã GD** (ABB12501, P5332601) — nếu chỉ nối bằng
`ma_cbtt` đã bóc (bỏ dạng 3) thì báo nhầm 1.570 mã "HNX có, VSD không có"; đúng ra phải thử **CẢ `cbtt_k` LẪN `gd_k`**
-> còn 855.

### KẾT QUẢ (17/07/2026)
**Mức độ đồng thuận rất cao**: trên **1.400 mã cả hai nguồn đều nói còn hiệu lực**, **1.390 mã (99,3%) khớp
CHÍNH XÁC tới từng đồng**, mệnh giá lệch **0 mã**; tổng 1.229,7 vs 1.226,6 nghìn tỷ (**+0,26%**).
=> Hai hệ thống độc lập xác nhận lẫn nhau; dữ liệu HNX của dự án đáng tin.

**CẦU NỐI** giải thích trọn vẹn chênh 1.394,0 (VSD hiệu lực) vs 1.235,3 (HNX ĐKGD) nghìn tỷ — A+B+D = VSD; A+C+E = HNX:
| | Nhóm | Số mã | Nghìn tỷ |
|-|------|------:|---------:|
|A| Cả 2 còn hiệu lực | 1.400 | 1.229,7 (HNX 1.226,6) |
|B| (−) VSD hiệu lực, HNX đã hủy ĐKGD | 103 | 70,2 |
|D| (−) Chỉ có ở VSD | 92 | 94,1 |
|C| (+) VSD hủy, HNX còn ĐKGD | 0 | 0,0 |
|E| (+) Chỉ có ở HNX, đang ĐKGD | 5 | 8,7 |
**Bản chất**: hủy ĐKGD (rời sàn) ≠ hủy lưu ký -> VSD giữ, HNX cho KL ĐKGD = 0. Nhóm D = TP mới 2026 chưa kịp
ĐKGD (30 mã) + TP cũ 2006–2021 trước thời kỳ ĐKGD.

**855 mã CBTT HNX không thấy ở VSD** — phần lớn là BÌNH THƯỜNG: 674 đã đáo hạn (VSD gỡ khỏi danh sách)
+ 140 đã mua lại hết trước hạn. **Chỉ 41 mã / 43,4 nghìn tỷ là khoảng trống thật** (chưa đáo hạn & còn dư nợ):
2 nhóm — (a) mới phát hành 06/2026, độ trễ đăng ký lưu ký (STB12602–05, ASP12601);
(b) **TCPH có vấn đề**: BSECH2126003 Bông Sen 4.800 tỷ, NTDCH2227001 Nova Thảo Điền 2.300 tỷ, VTPCH2126003/4/5.

**BẤT THƯỜNG NGUỒN đáng chú ý**: `VHA32105` (Thương mại - Quảng cáo - Xây dựng Việt Hân) — VSD ghi
**420.860 TP = 4.208,6 tỷ**, HNX ghi **42.086 TP = 420,86 tỷ**, mà CBTT phát hành chỉ **500 tỷ**
-> **VSD lệch 10× và vượt xa lượng phát hành** => nghiêng về lỗi dữ liệu phía VSD. Đây là 1 trong 10 mã ở
sheet "Lệch SL-GT (cùng hiệu lực)" — sheet này CHỈ gồm mã cả 2 nguồn còn hiệu lực nên chênh = bất thường THẬT.

### Sheet trong `VSD_vs_HNX_DoiChieu.xlsx`
`Tổng quan` · `Cầu nối chênh lệch` · `Lệch SL-GT (cùng hiệu lực)` (10) · `Lệch trạng thái` (103) ·
`Lệch ngày PH-ĐH` (78, mã nhiều đợt lệch là bình thường do CBTT gộp) · `Chỉ có ở VSD` (164) ·
`Chỉ có ở HNX (ĐKGD)` (5) · `CBTT HNX thiếu ở VSD` (855, lọc cột "Phân loại").

### Chạy lại
```
python vsd_bond_scraper.py          # ~12 phút (list 234 trang + detail 2.332 mã)
python vsd_vs_hnx_compare.py        # cần bond_catalog/issuance/buyback_raw.csv sẵn
```
CHƯA nối vào `update_daily.py` và CHƯA lên dashboard — cần user quyết có đưa thành tab thứ 9 không.


## VÒNG ĐỜI 2 NGUỒN — `vsd_hnx_lifecycle.py` -> `VSD_HNX_VongDoi.xlsx` (17/07/2026)

**MÔ HÌNH NGHIỆP VỤ (user chốt 17/07/2026)** — nền tảng để đọc mọi chênh lệch VSD↔HNX:
> Phát hành → **CBTT trên HNX trước** → đăng ký **lưu ký VSD** → xong mới quay lại **ĐKGD/niêm yết HNX**.
> Khi kết thúc (thường trước đáo hạn / trước mua lại): **hủy ĐKGD trên HNX TRƯỚC** → **hủy lưu ký VSD SAU**.
> Độ trễ mỗi khâu ~**2 tuần - 1 tháng**.
=> Chênh lệch giữa 2 nguồn phần lớn KHÔNG phải sai số, mà là **độ trễ vòng đời**. Mã nằm ngoài cửa sổ
quá lâu mới là bất thường. Đây là lăng kính đúng để đọc `VSD_vs_HNX_DoiChieu.xlsx`.

### Mô hình được dữ liệu XÁC NHẬN
| Khâu | TP phát hành từ 2024 (n=1.090) |
|------|-------------------------------|
| 1. PH → đăng ký lưu ký VSD | trung vị **15 ngày** (IQR 13–21; 87,1% ≤ 30 ngày) |
| 2. lưu ký VSD → ĐKGD HNX | trung vị **16 ngày** (IQR 13–22; 87,4% ≤ 30 ngày) |
| Tổng PH → lên sàn TPRL | trung vị **34 ngày** |

**Quy trình NHANH DẦN theo năm PH** (khâu 1 / khâu 2, trung vị ngày):
2023: 29/25 · 2024: 18/17 · 2025: 14/17 · 2026: **15/13**.

**Chiều nghịch xác nhận thứ tự**: mã "VSD đã hủy nhưng HNX còn ĐKGD" = **0/2.332** → HNX LUÔN hủy trước.
(Trước khi biết mô hình, nhóm C=0 là kết quả chưa lý giải được; giờ nó là bằng chứng khẳng định.)

⚠ **BẪY ĐO LƯỜNG QUAN TRỌNG**: nếu gộp cả TP cũ, khâu 1 trung vị = **27 ngày nhưng IQR 15–528**
(TP 2021-2022 bị đăng ký **HỒI TỐ** 2023 sau NĐ65 + sàn TPRL mở 19/07/2023). **Chỉ đo TP phát hành TỪ 2024**
(hằng số `DO_SACH_TU`). Đừng trích số "toàn bộ" ra báo cáo.

### Ngày lưu ký lấy ở đâu
`ngay_gcn` = ngày cấp Giấy CNĐKCK (trang chi tiết VSD) = **ngày đăng ký lưu ký**; VSD **KHÔNG công bố
ngày HỦY lưu ký** (tab 2 "Thông tin đăng ký chứng khoán" chỉ ghi các lần ĐK, không ghi hủy)
=> khâu 4 chỉ đo được dạng **"đang treo" (right-censored)**.
🐛 **BUG ĐÃ SỬA**: VSD viết ngày kiểu **`25/7/2025` (tháng 1 chữ số)** và 3 cách diễn đạt
("cấp lần đầu ngày" / "cấp ngày" / "ngày") -> regex `\d{2}/\d{2}/\d{4}` chỉ bóc được 1.231/2.332.
`extract_gcn_date()` neo vào chữ "ngày" + `\d{1,2}` -> **2.282/2.332** (n mẫu khâu 1 nhảy 174 -> 1.090).

### Phát hiện: mã NGOÀI cửa sổ (2 nhóm cần rà soát)
1. **ĐANG TREO** (HNX đã hủy ĐKGD, VSD còn "Hiệu lực"): 103 mã / 70,2 nghìn tỷ. **Chỉ 12/94 mã trong cửa
   sổ 30 ngày**; trung vị treo **232 ngày**, max 942 (NCR12101 GD cuối 18/12/2023). >1 năm: 37 mã/23,0 nghìn tỷ.
   => Độ trễ hủy lưu ký THỰC TẾ dài hơn nhiều mức 2 tuần-1 tháng.
2. **VSD "Hiệu lực" dù ĐÃ QUÁ ĐÁO HẠN**: 101 mã / 63,6 nghìn tỷ, quá hạn trung vị **494 ngày** (max 964).
   **Tỷ lệ có CBTT chậm trả = 30,7% vs nền 3,2% -> cao gấp 9,6 lần** (đối chiếu `bond_latepay_raw.csv`).
   Tách: 31 mã/19,1 nghìn tỷ đã xác nhận chậm trả · **70 mã/44,5 nghìn tỷ CHƯA rõ lý do = DANH SÁCH THEO DÕI**
   (VSD chậm cập nhật HOẶC chậm trả chưa CBTT — vd GHI12101 Golden Hill 5.760 tỷ ĐH 15/04/2024).
   ⚠ Diễn đạt cho đúng: KHÔNG kết luận cả nhóm là chậm trả (chỉ ~1/3 có CBTT xác nhận).

### Hệ quả: VSD "Hiệu lực" KHÔNG phải thước đo dư nợ
`1.394,0` nghìn tỷ (VSD hiệu lực) **gồm 63,6 nghìn tỷ đã quá đáo hạn**; loại ra còn `1.330,5`.
So: HNX đang ĐKGD `1.235,3` | dự án (PH − mua lại − đáo hạn) `~1.137`. **Vẫn dùng cách đo của dự án**;
VSD dùng để ĐỐI CHIẾU và làm TÍN HIỆU RỦI RO, không thay thế.

🐛 **BUG PANDAS ĐÃ MẮC — CẢNH BÁO**: `cat` có **5 dòng ISIN RỖNG**; `vj.join(ci, on='join_isin')` khi vj
còn 164 dòng `join_isin=''` -> **nhân bản chéo** (2.332 -> 2.988 dòng, thừa 656) làm 101 mã quá hạn phình
thành 237 và gán nhầm mã CBTT (SNL12101 -> VIC12510!). **LUÔN lọc `join_isin != ''` TRƯỚC khi join.**
`vsd_vs_hnx_compare.py` và `vsd_hnx_lifecycle.py` đã lọc đúng; script thăm dò ad-hoc thì phải nhớ.

### Chạy
```
python vsd_hnx_lifecycle.py     # cần vsd_bond_raw.csv (có cột ngay_gcn) + bond_catalog/issuance/latepay_raw.csv
```
