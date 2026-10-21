"""Gom event log ServiceNow (UCI #498) thành 1 dòng / incident.

Input : incident_event_log.csv  (141,712 event, 36 cột)
Output: incidents_clean.csv     (1 dòng / incident, trạng thái cuối cùng)

Chạy: python src/prepare_data.py data/raw/incident_event_log.zip data/processed/incidents_clean.csv
(pandas đọc thẳng file .zip chứa 1 CSV, không cần giải nén)
"""
import sys

import pandas as pd

src, dst = sys.argv[1], sys.argv[2]
df = pd.read_csv(src, na_values="?", low_memory=False)
for c in ("opened_at", "resolved_at", "closed_at", "sys_updated_at"):
    df[c] = pd.to_datetime(df[c], format="%d/%m/%Y %H:%M", errors="coerce")

# Mỗi incident có nhiều event -> lấy event cuối cùng (sys_mod_count lớn nhất) làm trạng thái chốt
df = df.sort_values(["number", "sys_mod_count", "sys_updated_at"])
history = df.groupby("number").agg(
    so_event=("incident_state", "size"),
    tung_cho_user=("incident_state", lambda s: (s == "Awaiting User Info").any()),
)
last = df.groupby("number").tail(1).set_index("number").join(history)

out = pd.DataFrame({
    "Mã incident": last.index,
    "Ngày mở": last["opened_at"].values,
    "Mức ưu tiên": last["priority"].values,
    "Danh mục": last["category"].values,
    "Nhóm xử lý": last["assignment_group"].values,
    "Kênh tiếp nhận": last["contact_type"].values,
    "Số lần chuyển nhóm": last["reassignment_count"].values,
    "Số lần mở lại": last["reopen_count"].values,
    "Từng chờ user bổ sung": last["tung_cho_user"].map({True: "Có", False: "Không"}).values,
    "Kết quả SLA": last["made_sla"].map({True: "Đạt", False: "Không đạt"}).values,
    "Ngày giải quyết": last["resolved_at"].values,
    "Ngày đóng": last["closed_at"].values,
    "Số event": last["so_event"].values,
})
out.to_csv(dst, index=False, encoding="utf-8-sig")
assert out["Mã incident"].is_unique and len(out) == df["number"].nunique()
print(f"{len(df):,} event -> {len(out):,} incident")
