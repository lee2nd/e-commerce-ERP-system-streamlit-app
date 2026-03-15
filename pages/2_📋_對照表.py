import streamlit as st
import pandas as pd
import io
from utils.data_manager import (
    load_compare_table, save_compare_table,
    load_storage,
    load_platform_orders,
    clear_compare_table,
)
from utils.parsers import parse_shopee, parse_ruten, parse_easystore
from utils.calculators import auto_match_compare_table

st.set_page_config(page_title="對照表", page_icon="📋", layout="wide")
st.title("📋 對照表")

compare = load_compare_table()
storage = load_storage()

# 統計
if not compare.empty:
    total = len(compare)
    matched = int(
        (compare["貨號"].fillna("").astype(str).str.strip() != "").sum()
    )
    c1, c2, c3 = st.columns(3)
    c1.metric("總商品數", total)
    c2.metric("已匹配", matched)
    c3.metric("未匹配", total - matched)

if st.button("🔄 重新掃描訂單（新增未匹配項目）", type="primary"):
    parts = []
    for plat, parser in [("蝦皮", parse_shopee), ("露天", parse_ruten), ("官網", parse_easystore)]:
        raw = load_platform_orders(plat)
        if not raw.empty:
            buf = io.BytesIO()
            raw.to_excel(buf, index=False, engine="openpyxl")
            buf.seek(0)
            buf.name = f"{plat}.xlsx"
            try:
                parts.append(parser(buf))
            except Exception:
                pass
    if parts:
        orders = pd.concat(parts, ignore_index=True)
    else:
        orders = pd.DataFrame()
    if orders.empty:
        st.warning("無訂單資料可匹配，請先至首頁匯入平台訂單")
    else:
        updated = auto_match_compare_table(orders, storage, compare)
        # 依平台排序
        _plat_order = {"蝦皮": 0, "露天": 1, "官網": 2}
        updated["_sort"] = updated["平台"].map(_plat_order).fillna(9)
        updated = updated.sort_values(["_sort", "平台商品名稱"]).drop(columns="_sort").reset_index(drop=True)
        save_compare_table(updated)
        st.success(f"掃描完成！對照表共 {len(updated)} 筆")
        st.rerun()

# 顯示對照表
if not compare.empty:
    filter_opt = st.radio(
        "篩選狀態", ["全部", "未匹配", "已匹配"],
        horizontal=True, key="cmp_filter",
        label_visibility="collapsed",
    )

    for col in ["入庫品名", "主貨號", "貨號", "平台商品名稱"]:
        if col in compare.columns:
            compare[col] = compare[col].fillna("").astype(str)

    view = compare.copy()
    if filter_opt == "未匹配":
        view = view[view["貨號"].str.strip() == ""]
    elif filter_opt == "已匹配":
        view = view[view["貨號"].str.strip() != ""]

    if view.empty:
        st.info("沒有符合條件的項目")
    else:
        _PLAT_COLORS = {"蝦皮": "#FF6B35", "露天": "#4A90D9", "官網": "#2ECC71"}

        def _highlight_platform(row):
            color = _PLAT_COLORS.get(row["平台"], "")
            if color:
                return [f"background-color: {color}20; color: {color}" if c == "平台"
                        else f"background-color: {color}10" for c in row.index]
            return [""] * len(row)

        styled_view = view[["平台", "平台商品名稱", "貨號", "主貨號", "入庫品名"]].style.apply(
            _highlight_platform, axis=1
        )
        st.dataframe(styled_view, width='stretch', hide_index=True)

else:
    st.info("對照表為空，請先至首頁「匯入平台訂單」上傳訂單資料")

# ── 清0 對照表 ───────────────────────────────────────────
st.markdown("---")
if st.button("🗑️ 清除對照表", key="clear_cmp_btn"):
    st.session_state["confirm_clear_cmp"] = True
if st.session_state.get("confirm_clear_cmp"):
    st.warning("⚠️ 確定要清除對照表所有資料嗎？此操作無法復原！")
    _c1, _c2 = st.columns(2)
    if _c1.button("✅ 確認清除", key="confirm_clear_cmp_yes", type="primary"):
        with st.spinner("清除中…"):
            clear_compare_table()
        st.session_state.pop("confirm_clear_cmp", None)
        st.success("✅ 對照表已清除")
        st.rerun()
    if _c2.button("❌ 取消", key="confirm_clear_cmp_no"):
        st.session_state.pop("confirm_clear_cmp", None)
        st.rerun()
