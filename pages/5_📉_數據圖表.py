"""數據圖表（需求 6：圖表商品清單優化）"""
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="數據圖表", page_icon="📉", layout="wide")

from utils.data_manager import load_daily_report

st.title("📉 數據圖表")

daily = load_daily_report()
if daily.empty:
    st.info("請先至「日報表」頁面產生報表")
    st.stop()

daily["日期"] = pd.to_datetime(daily["日期"], errors="coerce")
daily["月份"] = daily["日期"].dt.month

# ══════════════════════════════════════════════════════════════
# 1. 各平台月度訂單量
# ══════════════════════════════════════════════════════════════
st.markdown("### 📊 各平台月度訂單量")

# 排除非正常訂單
normal = daily[~daily["出貨狀態"].isin(["!取消!", "!退貨!"])]

plat_monthly = (
    normal.groupby(["月份", "平台"])["訂單編號"]
    .count()
    .reset_index(name="訂單數")
)

fig1 = px.bar(
    plat_monthly, x="月份", y="訂單數", color="平台",
    barmode="group", text_auto=True,
)
fig1.update_layout(xaxis=dict(dtick=1))
st.plotly_chart(fig1, use_container_width=True)

# ── 平均月訂單量 ─────────────────────────────────────────────
months_with_data = plat_monthly.groupby("月份")["訂單數"].sum()
active_months = (months_with_data > 0).sum()
if active_months > 0:
    avg_orders = int(months_with_data.sum() / active_months)
    st.metric("年度月平均訂單量", f"{avg_orders} 筆/月")

# ══════════════════════════════════════════════════════════════
# 2. 月度營業額 vs 淨利
# ══════════════════════════════════════════════════════════════
st.markdown("### 💰 月度營業額 vs 淨利")

monthly_fin = daily.groupby("月份").agg(
    營業額=("營業額", "sum"),
    淨利=("淨利", "sum"),
).reset_index()

fig2 = go.Figure()
fig2.add_trace(go.Bar(name="營業額", x=monthly_fin["月份"], y=monthly_fin["營業額"]))
fig2.add_trace(go.Bar(name="淨利", x=monthly_fin["月份"], y=monthly_fin["淨利"]))
fig2.update_layout(barmode="group", xaxis=dict(dtick=1), yaxis_title="NT$")
st.plotly_chart(fig2, use_container_width=True)

# ══════════════════════════════════════════════════════════════
# 3. 平台營收佔比
# ══════════════════════════════════════════════════════════════
st.markdown("### 🥧 平台營收佔比")
plat_rev = normal.groupby("平台")["營業額"].sum().reset_index()
if not plat_rev.empty and plat_rev["營業額"].sum() > 0:
    fig3 = px.pie(plat_rev, values="營業額", names="平台", hole=0.4)
    st.plotly_chart(fig3, use_container_width=True)
else:
    st.info("無營收資料")

# ══════════════════════════════════════════════════════════════
# 4. 日度趨勢
# ══════════════════════════════════════════════════════════════
st.markdown("### 📅 日度營業額趨勢")
daily_trend = daily.groupby("日期")["營業額"].sum().reset_index()
if not daily_trend.empty:
    fig4 = px.line(daily_trend, x="日期", y="營業額", markers=True)
    fig4.update_layout(yaxis_title="NT$")
    st.plotly_chart(fig4, use_container_width=True)
