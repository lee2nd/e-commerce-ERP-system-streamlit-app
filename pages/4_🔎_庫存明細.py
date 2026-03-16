import streamlit as st
import pandas as pd
from utils.data_manager import (
    load_storage, load_delivery,
    load_compare_table,
    load_inventory_details, save_inventory_details,
    clear_inventory_details,
)
from utils.calculators import generate_inventory_details

SIZE_ORDER = ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL"]
SIZE_RANK = {s: i for i, s in enumerate(SIZE_ORDER)}


def extract_size(spec: str) -> str:
    """Extract the first matching size token from a spec string."""
    spec_text = "" if pd.isna(spec) else str(spec)
    for size in reversed(SIZE_ORDER):  # match 2XL before XL
        if size in spec_text:
            return size
    return ""


def size_sort_key(spec: str) -> int:
    size = extract_size(spec)
    return SIZE_RANK.get(size, 999)

st.set_page_config(page_title="庫存明細", page_icon="🔎", layout="wide")
st.title("🔎 庫存明細")

if st.button("🔄 更新庫存明細", type="primary"):
    storage  = load_storage()
    delivery = load_delivery()
    compare  = load_compare_table()
    if storage.empty:
        st.warning("請先至「匯入資料」頁面新增入庫資料")
    else:
        result = generate_inventory_details(storage, delivery)
        # 過濾掉對照表內未匹配的項目（入庫品名為空或"未匹配"）
        if not compare.empty and "貨號" in compare.columns and "入庫品名" in compare.columns:
            matched_skus = set(
                compare.loc[
                    ~compare["入庫品名"].fillna("").astype(str).isin(["", "未匹配"]),
                    "貨號"
                ].astype(str).str.strip()
            )
            result = result[result["貨號"].astype(str).str.strip().isin(matched_skus)]
        save_inventory_details(result)
        st.success(f"✅ 庫存明細已更新！共 {len(result)} 筆")
        st.rerun()

inventory = load_inventory_details()

if inventory.empty:
    st.info("尚無庫存明細，請點擊上方「🔄 更新庫存明細」按鈕產生")
    st.stop()

# 摘要
c1, c2, c3 = st.columns(3)
c1.metric("商品種類", f"{len(inventory)}")
c2.metric("庫存充足", int((inventory["現有庫存"] > 0).sum()))
c3.metric("庫存不足（≤0）", int((inventory["現有庫存"] <= 0).sum()))

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

if "主貨號" in view.columns and "規格" in view.columns:
    view = view.sort_values(
        by=["主貨號", "規格"],
        key=lambda col: col.map(size_sort_key) if col.name == "規格" else col,
    )
    view = view.reset_index(drop=True)

st.dataframe(view, width='stretch', hide_index=True)


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

