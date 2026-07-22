# -*- coding: utf-8 -*-
"""Rebuild dashboard_data.json + dashboard.html (KHÔNG dựng lại Excel) - dùng khi chỉ cập nhật
   dữ liệu (giao dịch thứ cấp / XHTN / chậm trả...). Tái sử dụng logic của build_reports."""
import os
import build_reports as B

df = B.load()
bb = B.load_buyback() if os.path.exists(B.BBRAW) else None
out = B.compute_outstanding(df, bb) if bb is not None else None
sec = B.load_secondary(df, bb)
rating = B.load_rating()
dno_map = B.outstanding_by_bond(df, bb)
news = B.load_news()
latepay = B.load_latepay(dno_map, out["outstanding_ty"] if out is not None else None, news)
giahan = B.load_giahan(out)
B.build_json(df, out, bb, sec, rating, latepay, giahan)
print("Rebuild dashboard xong (json + html).")
