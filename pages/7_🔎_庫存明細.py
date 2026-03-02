"""庫存明細（需求 4：包含主貨號 / 貨號）"""
import streamlit as st
import pandas as pd

st.set_page_config(page_title="庫存明細", page_icon="🔎", layout="wide")

from utils.data_manager import load_storage, load_delivery
from utils.calculators import generate_inventory

st.title("🔎 庫存明細")

storage  = load_storage()
delivery = load_delivery()

if storage.empty:
    st.info("請先至「匯入資料」頁面新增入庫資料")
    st.stop()

inventory = generate_inventory(storage, delivery)

if inventory.empty:
    st.warning("無法產生庫存明細（入庫資料可能缺少必要欄位）")
    st.stop()

# ── 摘要 ─────────────────────────────────────────────────────
c1, c2, c3 = st.columns(3)
c1.metric("商品種類", f"{len(inventory)} 項")
c2.metric("庫存充足", int((inventory["現有庫存"] > 0).sum()))
c3.metric("庫存不足（≤0）", int((inventory["現有庫存"] <= 0).sum()))

# ── 篩選 ─────────────────────────────────────────────────────
search = st.text_input("🔍 搜尋（商品名稱 / 貨號 / 主貨號）")
view = inventory.copy()
if search:
    mask = (
        view["商品名稱"].str.contains(search, case=False, na=False)
        | view["貨號"].str.contains(search, case=False, na=False)
        | view["主貨號"].str.contains(search, case=False, na=False)
    )
    view = view[mask]

low_stock = st.checkbox("僅顯示庫存不足（≤0）")
if low_stock:
    view = view[view["現有庫存"] <= 0]

# ── 表格 ─────────────────────────────────────────────────────
st.dataframe(
    view,
    use_container_width=True,
    hide_index=True,
    column_config={
        "進貨金額": st.column_config.NumberColumn(format="$%d"),
        "銷售金額": st.column_config.NumberColumn(format="$%d"),
        "平均成本": st.column_config.NumberColumn(format="$%.1f"),
    },
)

# 下載
csv = view.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
st.download_button("📥 下載庫存明細 CSV", csv, "inventory.csv", "text/csv")
