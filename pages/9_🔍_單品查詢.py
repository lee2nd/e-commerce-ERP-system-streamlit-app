"""單品查詢"""
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="單品查詢", page_icon="🔍", layout="wide")

from utils.data_manager import load_delivery, load_storage

st.title("🔍 單品查詢")

delivery = load_delivery()
storage  = load_storage()

if delivery.empty:
    st.info("請先至「日報表」頁面產生報表")
    st.stop()

# 排除 TBD
valid = delivery[
    delivery["商品名稱"].notna()
    & (delivery["商品名稱"] != "")
    & (~delivery["商品名稱"].str.contains("TBD", case=False, na=False))
].copy()

# ── 建立 (貨號)商品名稱 組合 key（對應 VBA SingleItemSearch 格式）──
_NULL_LIKE = frozenset({"nan", "none", "nat", "<na>", ""})


def _sku_clean(v) -> str:
    s = str(v).strip() if pd.notna(v) else ""
    return "" if s.lower() in _NULL_LIKE else s


valid["商品key"] = valid.apply(
    lambda r: f"({_sku_clean(r['貨號'])}){r['商品名稱']}"
              if _sku_clean(r["貨號"]) else r["商品名稱"],
    axis=1,
)


def _extract_name(label: str) -> str:
    """從 '(貨號)商品名稱' 提取商品名稱（對應 VBA InStr/Right 解析）"""
    if label.startswith("(") and ")" in label:
        return label[label.index(")") + 1:]
    return label


# ── 商品選單 ─────────────────────────────────────────────────
product_list = sorted(valid["商品key"].unique().tolist())
selected = st.selectbox("選擇商品", product_list)

if not selected:
    st.stop()

item_name = _extract_name(selected)
item_data = valid[valid["商品key"] == selected].copy()
item_data["日期"] = pd.to_datetime(item_data["日期"], errors="coerce")
item_data["月份"] = item_data["日期"].dt.month

# ── 平均成本（對應 VBA VLOOKUP 庫存明細!單位成本）────────────
avg_cost = 0.0
if not storage.empty and "商品名稱" in storage.columns:
    stg_match = storage[storage["商品名稱"] == item_name]
    if not stg_match.empty:
        avg_cost = pd.to_numeric(stg_match["單位成本"], errors="coerce").mean()
        if pd.isna(avg_cost):
            avg_cost = 0.0

# ── 摘要 ─────────────────────────────────────────────────────
c1, c2, c3, c4 = st.columns(4)
c1.metric("總銷售數量", f"{int(item_data['數量'].sum()):,}")
c2.metric("總銷售金額", f"${int(item_data['金額'].sum()):,}")
c3.metric("平均成本", f"${avg_cost:,.1f}")
c4.metric(
    "估計總淨利",
    f"${int(item_data['金額'].sum() - item_data['數量'].sum() * avg_cost):,}",
)

# ── 月度銷售趨勢（對應 VBA SingleItemVsualization B/C欄）─────
st.markdown("### 月度銷售趨勢")

monthly = item_data.groupby("月份").agg(
    銷售金額=("金額", "sum"),
    銷售數量=("數量", "sum"),
).reset_index()

if not monthly.empty:
    monthly["估計淨利"] = monthly["銷售金額"] - monthly["銷售數量"] * avg_cost

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="銷售金額", x=monthly["月份"], y=monthly["銷售金額"],
        text=monthly["銷售金額"].apply(lambda x: f"${x:,.0f}"),
        textposition="outside",
    ))
    fig.add_trace(go.Bar(
        name="估計淨利", x=monthly["月份"], y=monthly["估計淨利"],
        text=monthly["估計淨利"].apply(lambda x: f"${x:,.0f}"),
        textposition="outside",
    ))
    fig.update_layout(barmode="group", xaxis=dict(dtick=1), yaxis_title="NT$")
    st.plotly_chart(fig, width="stretch")

# ── 規格分布（對應 VBA SpecRanking）─────────────────────────
st.markdown("### 規格銷量分布")

spec = (
    item_data.groupby("規格")
    .agg(銷售數量=("數量", "sum"), 銷售金額=("金額", "sum"))
    .sort_values(by="銷售數量", ascending=False) # type: ignore
    .reset_index()
)

if not spec.empty and len(spec) > 1:
    fig2 = px.pie(spec, values="銷售數量", names="規格", hole=0.3)
    st.plotly_chart(fig2, width="stretch")

st.dataframe(spec, width="stretch", hide_index=True)

# ── 平台分布 ─────────────────────────────────────────────────
st.markdown("### 各平台銷量")

plat = (
    item_data.groupby("平台")
    .agg(銷售數量=("數量", "sum"), 銷售金額=("金額", "sum"))
    .sort_values(by="銷售數量", ascending=False) # type: ignore
    .reset_index()
)
st.dataframe(plat, width="stretch", hide_index=True)

# ── 明細 ─────────────────────────────────────────────────────
with st.expander("📋 銷售明細"):
    st.dataframe(
        item_data[["日期", "規格", "貨號", "數量", "單價", "金額", "平台"]],
        width="stretch",
        hide_index=True,
    )
