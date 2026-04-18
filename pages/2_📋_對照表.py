import streamlit as st
import pandas as pd
import io
from datetime import datetime, timezone, timedelta
from utils.data_manager import (
    load_compare_table, save_compare_table,
    load_storage,
    load_platform_orders,
    clear_compare_table,
    load_combo_sku,
    load_custom_orders,
)
from utils.parsers import parse_shopee, parse_ruten, parse_easystore, parse_mo
from utils.calculators import auto_match_compare_table
from utils.styles import apply_global_styles

st.set_page_config(page_title="對照表", page_icon="📋", layout="wide")

apply_global_styles()
st.title("📋 對照表")

TZ_TAIPEI = timezone(timedelta(hours=8))

compare = load_compare_table()
storage = load_storage()
combo = load_combo_sku()

# 建立 組合貨號 → 原料名稱說明 lookup（用於在對照表顯示組合內容）
combo_detail_map: dict[str, str] = {}
if not combo.empty and not storage.empty:
    name_col = "商品名稱" if "商品名稱" in storage.columns else "名稱"
    stg_name_lookup = (
        storage.drop_duplicates("貨號")
        .set_index("貨號")
        .apply(
            lambda r: f"{r.get(name_col, '')} {r.get('規格', '')}".strip(),
            axis=1,
        )
        .to_dict()
    )
    for combo_code, grp in combo.groupby("組合貨號"):
        parts = []
        for _, comp in grp.iterrows():
            mat_sku = str(comp["原料貨號"]).strip()
            mat_qty = int(comp["原料數量"])
            mat_name = stg_name_lookup.get(mat_sku, mat_sku)
            parts.append(f"{mat_name} ×{mat_qty}")
        combo_detail_map[str(combo_code)] = " + ".join(parts)

# 統計
if not compare.empty:
    total = len(compare)
    matched = int(
        (~compare["入庫品名"].fillna("").astype(str).str.strip().isin(["", "未匹配"])).sum()
    )
    c1, c2, c3 = st.columns(3)
    c1.metric("總商品數", total)
    c2.metric("已匹配", matched)
    c3.metric("未匹配", total - matched)

if st.button("🔄 重新掃描訂單", type="primary"):
    parts = []
    for plat, parser in [("蝦皮", parse_shopee), ("露天", parse_ruten), ("官網", parse_easystore), ("MO店", parse_mo)]:
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
    # 自建訂單：轉換為統一格式加入掃描
    _cust_raw = load_custom_orders()
    if not _cust_raw.empty and "貨號" in _cust_raw.columns:
        _cust_orders = pd.DataFrame()
        _cust_orders["訂單編號"] = _cust_raw["訂單編號"].astype(str)
        _cust_orders["日期"] = pd.to_datetime(_cust_raw["日期"], errors="coerce")
        _cust_orders["平台"] = "其他"
        _cust_orders["平台商品名稱"] = _cust_raw["貨號"].fillna("").astype(str).str.strip()
        _cust_orders["貨號"] = _cust_raw["貨號"].fillna("").astype(str).str.strip()
        _cust_orders["數量"] = (
            pd.to_numeric(_cust_raw["數量"], errors="coerce").fillna(0).astype(int)
            if "數量" in _cust_raw.columns else 0
        )
        _cust_orders["單價"] = (
            pd.to_numeric(_cust_raw["單價"], errors="coerce").fillna(0)
            if "單價" in _cust_raw.columns else 0.0
        )        
        _cust_orders["金額"] = _cust_orders["數量"] * _cust_orders["單價"]
        _cust_orders["賣家折扣"] = 0
        _cust_orders["訂單狀態"] = _cust_raw.get("訂單狀態", "已完成")
        parts.append(_cust_orders)
    if parts:
        orders = pd.concat(parts, ignore_index=True)
    else:
        orders = pd.DataFrame()
    if orders.empty:
        st.warning("無訂單資料可匹配，請先至首頁匯入平台訂單或自建訂單")
    else:
        updated = auto_match_compare_table(orders, storage, compare, load_combo_sku())
        # 依平台排序
        _plat_order = {"蝦皮": 0, "露天": 1, "官網": 2, "MO店": 3, "其他": 4}
        updated["_sort"] = updated["平台"].map(_plat_order).fillna(9)
        updated = updated.sort_values(["_sort", "平台商品名稱"]).drop(columns="_sort").reset_index(drop=True)
        save_compare_table(updated)
        st.session_state["compare_saved_at"] = datetime.now(tz=TZ_TAIPEI).strftime("%Y-%m-%d %H:%M:%S")
        st.success(f"掃描完成！對照表共 {len(updated)} 筆")
        st.rerun()

_compare_ts = st.session_state.get("compare_saved_at")
if _compare_ts:
    st.caption(f"🕐 最後掃描儲存：{_compare_ts}")

# 顯示對照表
if not compare.empty:
    filter_opt = st.radio(
        "篩選狀態", ["全部", "未匹配", "已匹配", "組合貨號"],
        horizontal=True, key="cmp_filter",
        label_visibility="collapsed",
    )

    for col in ["入庫品名", "主貨號", "貨號", "平台商品名稱"]:
        if col in compare.columns:
            compare[col] = compare[col].fillna("").astype(str)

    view = compare.copy()
    if filter_opt == "未匹配":
        view = view[view["入庫品名"].astype(str) == "未匹配"]
    elif filter_opt == "已匹配":
        view = view[~view["入庫品名"].astype(str).isin(["", "未匹配"])]
    elif filter_opt == "組合貨號":
        view = view[view["入庫品名"].astype(str).str.startswith("組合:")]

    if view.empty:
        st.info("沒有符合條件的項目")
    else:
        _PLAT_COLORS = {"蝦皮": "#FF6B35", "露天": "#4A90D9", "官網": "#2ECC71", "MO店": "#AB63FA", "其他": "#F39C12"}

        # 建立顯示用的 入庫品名 欄位：組合商品顯示 [組合] 並附上原料名稱
        def _format_stg_name(row):
            val = str(row.get("入庫品名", ""))
            if val.startswith("組合:"):
                combo_code = str(row.get("貨號", "")).strip()
                detail = combo_detail_map.get(combo_code, "")
                if detail:
                    return f"[組合] {detail}"
                return "[組合]"
            return val

        view = view.copy()
        view["入庫品名_顯示"] = view.apply(_format_stg_name, axis=1)

        def _highlight_platform(row):
            color = _PLAT_COLORS.get(row["平台"], "")
            if color:
                return [f"background-color: {color}20; color: {color}" if c == "平台"
                        else f"background-color: {color}10" for c in row.index]
            return [""] * len(row)

        def _highlight_unmatched(val):
            return "color: red" if val == "未匹配" else ""

        def _highlight_combo(val):
            return "color: #8B5CF6; font-weight: bold" if str(val).startswith("[組合]") else ""

        display_view = view[["平台", "平台商品名稱", "貨號", "主貨號", "入庫品名_顯示"]].rename(
            columns={"入庫品名_顯示": "入庫品名"}
        )
        # 其他平台的平台商品名稱顯示為空白
        display_view.loc[view["平台"] == "其他", "平台商品名稱"] = ""

        _cmp_page_size = 500
        _cmp_total = len(display_view)
        _cmp_total_pages = max(1, (_cmp_total - 1) // _cmp_page_size + 1)
        _cmp_dl_col, _cmp_pg_col = st.columns([1, 3])
        _cmp_csv = display_view.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        _cmp_dl_col.download_button("⬇️ 下載對照表", data=_cmp_csv, file_name="對照表.csv", mime="text/csv", key="dl_cmp")
        if _cmp_total_pages > 1:
            _cmp_page = _cmp_pg_col.selectbox(
                "頁碼", list(range(1, _cmp_total_pages + 1)),
                format_func=lambda x: f"{x}/{_cmp_total_pages} 頁",
                key="cmp_page",
            )
        else:
            _cmp_page = 1
        _cmp_start = (_cmp_page - 1) * _cmp_page_size
        display_slice = display_view.iloc[_cmp_start:_cmp_start + _cmp_page_size].copy()

        styled_view = (
            display_slice
            .style.apply(_highlight_platform, axis=1)
            .map(_highlight_unmatched, subset=["入庫品名"])
            .map(_highlight_combo, subset=["入庫品名"])
        )
        st.dataframe(styled_view, width='stretch', hide_index=True)
        st.caption(f"第 {_cmp_page} 頁 / 共 {_cmp_total_pages} 頁（{_cmp_total:,} 筆）")

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
