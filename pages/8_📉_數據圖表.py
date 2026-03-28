"""數據圖表"""
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="數據圖表", page_icon="📉", layout="wide")

from utils.data_manager import load_daily_report, load_monthly_report

st.title("📉 數據圖表")

daily = load_daily_report()
if daily.empty:
    st.info("請先至「日報表」頁面產生報表")
    st.stop()

daily["日期"] = pd.to_datetime(daily["日期"], errors="coerce")
daily["月份"] = daily["日期"].dt.month

# ══════════════════════════════════════════════════════════════
# 1. 各平台月度訂單量（對應 VBA DataVisualization 蝦皮/露天/Y拍欄位）
# ══════════════════════════════════════════════════════════════
st.markdown("### 📊 各平台月度訂單量")

# 排除取消/退貨（對應 VBA 條件 出貨狀態="" ）
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
st.plotly_chart(fig1, width="stretch")

# ── 年度月平均訂單量（對應 VBA G欄 average monthly order count）──
months_with_data = plat_monthly.groupby("月份")["訂單數"].sum()
active_months = (months_with_data > 0).sum()
if active_months > 0:
    avg_orders = int(months_with_data.sum() / active_months)
    st.metric("年度月平均訂單量", f"{avg_orders} 筆/月")

# ══════════════════════════════════════════════════════════════
# 2. 月度營業額 vs 淨利
#    淨利優先使用月報表（含固定月費，對應 VBA DataVisualization F欄）
#    若月報表尚未產生則 fallback 至日報表累計淨利
# ══════════════════════════════════════════════════════════════
st.markdown("### 💰 月度營業額 vs 淨利")

monthly_rev = daily.groupby("月份").agg(
    營業額=("營業額", "sum"),
).reset_index()

monthly_rep = load_monthly_report()
if not monthly_rep.empty and "月份" in monthly_rep.columns and "淨利" in monthly_rep.columns:
    rep_净利 = (
        pd.to_numeric(monthly_rep["月份"], errors="coerce")
        .pipe(lambda s: monthly_rep.assign(月份=s))
        .groupby("月份")["淨利"]
        .sum()
        .reset_index()
    )
    monthly_fin = monthly_rev.merge(rep_净利, on="月份", how="left")
    monthly_fin["淨利"] = pd.to_numeric(monthly_fin["淨利"], errors="coerce").fillna(0).astype(int)
    monthly_fin["淨利來源"] = "月報表（含固定月費）"
else:
    daily_net = daily.groupby("月份")["淨利"].sum().reset_index()
    monthly_fin = monthly_rev.merge(daily_net, on="月份", how="left")
    monthly_fin["淨利"] = pd.to_numeric(monthly_fin["淨利"], errors="coerce").fillna(0).astype(int)
    monthly_fin["淨利來源"] = "日報表（僅含變動成本）"

淨利來源 = monthly_fin["淨利來源"].iloc[0] if not monthly_fin.empty else ""
st.caption(f"淨利資料來源：{淨利來源}")

fig2 = go.Figure()
fig2.add_trace(go.Bar(name="營業額", x=monthly_fin["月份"], y=monthly_fin["營業額"]))
fig2.add_trace(go.Bar(name="淨利", x=monthly_fin["月份"], y=monthly_fin["淨利"]))
fig2.update_layout(barmode="group", xaxis=dict(dtick=1), yaxis_title="NT$")
st.plotly_chart(fig2, width="stretch")

# ══════════════════════════════════════════════════════════════
# 3. 平台營收佔比
# ══════════════════════════════════════════════════════════════
st.markdown("### 🥧 平台營收佔比")
plat_rev = normal.groupby("平台")["營業額"].sum().reset_index()
if not plat_rev.empty and plat_rev["營業額"].sum() > 0:
    fig3 = px.pie(plat_rev, values="營業額", names="平台", hole=0.4)
    st.plotly_chart(fig3, width="stretch")
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
    st.plotly_chart(fig4, width="stretch")
