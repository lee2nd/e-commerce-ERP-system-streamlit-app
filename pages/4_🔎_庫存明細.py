import streamlit as st
from utils.data_manager import (
    load_storage, load_delivery,
    load_inventory_details, save_inventory_details,
    clear_inventory_details,
)
from utils.calculators import generate_inventory_details

st.set_page_config(page_title="庫存明細", page_icon="🔎", layout="wide")
st.title("🔎 庫存明細")

if st.button("🔄 更新庫存明細", type="primary"):
    storage  = load_storage()
    delivery = load_delivery()
    if storage.empty:
        st.warning("請先至「匯入資料」頁面新增入庫資料")
    else:
        result = generate_inventory_details(storage, delivery)
        save_inventory_details(result)
        st.success(f"✅ 庫存明細已更新！共 {len(result)} 筆")
        st.rerun()

inventory = load_inventory_details()

if inventory.empty:
    st.info("尚無庫存明細，請點擊上方「🔄 更新庫存明細」按鈕產生")
    st.stop()

# 摘要
c1, c2, c3, c4 = st.columns(4)
c1.metric("商品種類", f"{len(inventory)}")
c2.metric("庫存充足", int((inventory["現有庫存"] > 0).sum()))
c3.metric("庫存不足（≤0）", int((inventory["現有庫存"] <= 0).sum()))
cost_mismatch = int(
    (inventory["平均成本(庫存明細)"] != inventory["平均成本(入庫)"]).sum()
)
c4.metric("成本不一致 ⚠️", cost_mismatch)

# 篩選
col_s, col_f = st.columns([3, 1])
search = col_s.text_input("🔍 搜尋（主貨號 / 貨號 / 名稱）")
low_stock = col_f.checkbox("僅顯示庫存不足（≤0）")

view = inventory.copy()
if search:
    mask = (
        view["名稱"].astype(str).str.contains(search, case=False, na=False)
        | view["貨號"].astype(str).str.contains(search, case=False, na=False)
        | view["主貨號"].astype(str).str.contains(search, case=False, na=False)
    )
    view = view[mask]
if low_stock:
    view = view[view["現有庫存"] <= 0]

# 表格（成本不一致行標黃）
def _highlight_cost_mismatch(row):
    if row["平均成本(庫存明細)"] != row["平均成本(入庫)"]:
        return ["background-color: #fff3cd"] * len(row)
    return [""] * len(row)

styled = view.style.apply(_highlight_cost_mismatch, axis=1).format({
    "進貨合計": "${:,.0f}",
    "銷售合計": "${:,.0f}",
    "平均成本(庫存明細)": "${:.1f}",
    "平均成本(入庫)":     "${:.1f}",
})

st.dataframe(styled, width='stretch', hide_index=True)

if cost_mismatch:
    st.caption("⚠️ 黃底列表示「平均成本(庫存明細)」與「平均成本(入庫)」數字不同，請確認入庫資料是否有誤")

# ── 清0 庫存明細 ───────────────────────────────────────────
st.markdown("---")
if st.button("🗑️ 清除庫存明細", key="clear_inv_btn"):
    st.session_state["confirm_clear_inv"] = True
if st.session_state.get("confirm_clear_inv"):
    st.warning("⚠️ 確定要清除所有庫存明細資料嗎？此操作無法復原！")
    _c1, _c2 = st.columns(2)
    if _c1.button("✅ 確認清除", key="confirm_clear_inv_yes", type="primary"):
        with st.spinner("清除中…"):
            clear_inventory_details()
        st.session_state.pop("confirm_clear_inv", None)
        st.success("✅ 庫存明細已清除")
        st.rerun()
    if _c2.button("❌ 取消", key="confirm_clear_inv_no"):
        st.session_state.pop("confirm_clear_inv", None)
        st.rerun()

