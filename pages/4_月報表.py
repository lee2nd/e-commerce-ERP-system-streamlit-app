"""月報表"""
import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="月報表", page_icon="📈", layout="wide")

from utils.data_manager import load_daily_report
from utils.calculators import generate_monthly_report

st.title("📈 月報表")

daily = load_daily_report()
if daily.empty:
    st.info("請先至「日報表」頁面產生報表")
    st.stop()

monthly = generate_monthly_report(daily)

if monthly.empty:
    st.info("無月份資料")
    st.stop()

# ── 表格 ─────────────────────────────────────────────────────
st.dataframe(
    monthly,
    width="stretch",
    hide_index=True,
    column_config={
        "月份": st.column_config.NumberColumn("月份", format="%d 月"),
        "營業額": st.column_config.NumberColumn(format="$%d"),
        "成本": st.column_config.NumberColumn(format="$%d"),
        "賣家折扣": st.column_config.NumberColumn(format="$%d"),
        "運費折抵": st.column_config.NumberColumn(format="$%d"),
        "成交手續費": st.column_config.NumberColumn(format="$%d"),
        "金流服務費": st.column_config.NumberColumn(format="$%d"),
        "淨利": st.column_config.NumberColumn(format="$%d"),
    },
)

# ── 圖表 ─────────────────────────────────────────────────────
st.markdown("### 月度趨勢")
chart_data = monthly.melt(
    id_vars="月份",
    value_vars=["營業額", "成本", "淨利"],
    var_name="指標",
    value_name="金額",
)
fig1 = px.bar(
    chart_data, x="月份", y="金額", color="指標",
    barmode="group", text_auto=True,
    labels={"金額": "NT$", "月份": "月份"},
)
fig1.update_layout(xaxis=dict(dtick=1))
st.plotly_chart(fig1, width="stretch")

# 訂單數
st.markdown("### 月度訂單數")
fig2 = px.bar(
    monthly, x="月份", y="訂單數", text_auto=True,
    labels={"訂單數": "筆"},
)
fig2.update_layout(xaxis=dict(dtick=1))
st.plotly_chart(fig2, width="stretch")

# 下載
csv = monthly.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
st.download_button("📥 下載月報表 CSV", csv, "monthly_report.csv", "text/csv")
