"""
電商平台進銷存系統 — 首頁 / 儀表板
平台：蝦皮 ｜ 露天 ｜ 官網 (EasyStore)
"""
import streamlit as st
import pandas as pd

st.set_page_config(page_title="電商平台進銷存系統", page_icon="📊", layout="wide")

from utils.data_manager import load_orders, load_storage, load_compare_table, load_daily_report

st.title("📊 電商平台進銷存系統")
st.caption("蝦皮 ｜ 露天 ｜ 官網 (EasyStore)")
st.markdown("---")

# ── 快速指標 ─────────────────────────────────────────────────
orders = load_orders()
daily  = load_daily_report()
compare = load_compare_table()
storage = load_storage()

c1, c2, c3, c4 = st.columns(4)

with c1:
    n = orders["訂單編號"].nunique() if not orders.empty and "訂單編號" in orders.columns else 0
    st.metric("總訂單數", f"{n:,}")
with c2:
    rev = int(daily["營業額"].sum()) if not daily.empty and "營業額" in daily.columns else 0
    st.metric("總營業額", f"${rev:,}")
with c3:
    profit = int(daily["淨利"].sum()) if not daily.empty and "淨利" in daily.columns else 0
    st.metric("總淨利", f"${profit:,}")
with c4:
    unmatched = 0
    if not compare.empty and "貨號" in compare.columns:
        unmatched = int(compare["貨號"].isna().sum() + (compare["貨號"] == "").sum())
    st.metric("未匹配商品", unmatched)

# ── 各平台訂單統計 ───────────────────────────────────────────
if not orders.empty and "平台" in orders.columns:
    st.markdown("### 各平台訂單量")
    plat_counts = orders.groupby("平台")["訂單編號"].nunique().reset_index()
    plat_counts.columns = ["平台", "訂單數"]
    cols = st.columns(len(plat_counts))
    for idx, (i, row) in enumerate(plat_counts.iterrows()):
        with cols[idx]:
            st.metric(row["平台"], f"{row['訂單數']:,} 筆")

# ── 近期訂單 ─────────────────────────────────────────────────
if not daily.empty:
    st.markdown("### 近期日報表（最新 20 筆）")
    recent = daily.sort_values("日期", ascending=False).head(20)
    st.dataframe(recent, use_container_width=True, hide_index=True)
else:
    st.info("👈 尚未有報表資料，請先至側邊欄「📥 匯入資料」頁面上傳平台訂單")