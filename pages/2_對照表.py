"""
對照表管理（需求 4：新增主貨號 / 貨號，需求 5：依貨號自動匹配）
"""
import streamlit as st
import pandas as pd

st.set_page_config(page_title="對照表管理", page_icon="📋", layout="wide")

from utils.data_manager import (
    load_compare_table, save_compare_table,
    load_storage, load_orders,
)
from utils.calculators import auto_match_compare_table

st.title("📋 對照表管理")

compare = load_compare_table()
storage = load_storage()

# ── 統計 ─────────────────────────────────────────────────────
if not compare.empty:
    total = len(compare)
    matched = int((compare["貨號"].notna() & (compare["貨號"] != "") & (compare["貨號"] != "nan")).sum())
    c1, c2, c3 = st.columns(3)
    c1.metric("總商品數", total)
    c2.metric("已匹配", matched)
    c3.metric("未匹配", total - matched)

# ── 重新自動匹配 ─────────────────────────────────────────────
col_a, col_b = st.columns(2)
with col_a:
    if st.button("🔄 重新自動匹配貨號", type="primary"):
        orders = load_orders()
        if orders.empty:
            st.warning("無訂單資料可匹配")
        else:
            updated = auto_match_compare_table(orders, storage, compare)
            save_compare_table(updated)
            st.success("自動匹配完成！")
            st.rerun()

with col_b:
    if st.button("🗑️ 清空對照表"):
        save_compare_table(pd.DataFrame())
        st.rerun()

st.markdown("---")

# ── 可編輯表格 ───────────────────────────────────────────────
if not compare.empty:
    st.subheader("編輯對照表")
    st.caption("可直接在表格中修改「主貨號」或「貨號」，完成後按下方按鈕儲存")

    # 提供入庫貨號清單做為參考
    sku_options = []
    if not storage.empty and "貨號" in storage.columns:
        sku_options = sorted(storage["貨號"].dropna().unique().tolist())

    # 篩選
    filter_opt = st.radio(
        "篩選", ["全部", "未匹配", "已匹配"],
        horizontal=True, key="cmp_filter",
    )
    view = compare.copy()
    if filter_opt == "未匹配":
        view = view[(view["貨號"].isna()) | (view["貨號"] == "") | (view["貨號"] == "nan")]
    elif filter_opt == "已匹配":
        view = view[(view["貨號"].notna()) & (view["貨號"] != "") & (view["貨號"] != "nan")]

    if view.empty:
        st.info("沒有符合條件的項目")
    else:
        edited = st.data_editor(
            view,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "平台商品名稱": st.column_config.TextColumn("平台商品名稱", disabled=True, width="large"),
                "平台": st.column_config.TextColumn("平台", disabled=True, width="small"),
                "主貨號": st.column_config.TextColumn("主貨號", width="medium"),
                "貨號": st.column_config.TextColumn("貨號", width="medium"),
            },
            key="compare_editor",
        )

        if st.button("💾 儲存變更", type="primary"):
            # 把編輯後的資料合併回完整的 compare table
            unchanged = compare[~compare["平台商品名稱"].isin(view["平台商品名稱"])]
            saved = pd.concat([unchanged, edited], ignore_index=True)
            saved = saved.drop_duplicates("平台商品名稱")
            save_compare_table(saved)
            st.success("已儲存！")
            st.rerun()

    # 入庫貨號參考
    if sku_options:
        with st.expander("📦 入庫貨號參考清單"):
            ref = storage[["主貨號", "商品名稱", "規格", "貨號"]].drop_duplicates("貨號")
            st.dataframe(ref, use_container_width=True, hide_index=True)
else:
    st.info("對照表為空，請先至「匯入資料」頁面上傳訂單")
