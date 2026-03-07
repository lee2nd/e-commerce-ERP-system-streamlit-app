"""
對照表管理 — 平台商品名稱 ↔ 入庫品名 對照，自動帶出貨號
"""
import streamlit as st
import pandas as pd

st.set_page_config(page_title="對照表管理", page_icon="📋", layout="wide")

from utils.data_manager import (
    load_compare_table, save_compare_table,
    load_storage,
    load_platform_orders,
)
from utils.parsers import parse_shopee, parse_ruten, parse_easystore, read_file_flexible
from utils.calculators import auto_match_compare_table

st.title("📋 對照表管理")

compare = load_compare_table()
storage = load_storage()

# ── 從入庫.xlsx 建立「入庫品名」下拉清單與貨號映射 ──────────────
stg_mapping: dict[str, dict] = {}   # 入庫品名 → {貨號, 主貨號}
if not storage.empty:
    for _, row in storage.drop_duplicates(subset=["名稱", "規格"] if "名稱" in storage.columns
                                          else ["商品名稱", "規格"]).iterrows():
        name = str(row.get("商品名稱", row.get("名稱", ""))).strip()
        spec = str(row.get("規格", "")).strip()
        if name:
            label = f"{name}[{spec}]" if spec else name
            stg_mapping[label] = {
                "貨號": str(row.get("貨號", "")),
                "主貨號": str(row.get("主貨號", "")),
            }
stg_options = [""] + sorted(stg_mapping.keys())

# ── 確保 入庫品名 欄位存在 ───────────────────────────────────
if not compare.empty and "入庫品名" not in compare.columns:
    compare["入庫品名"] = ""

# ── 統計 ─────────────────────────────────────────────────────
if not compare.empty:
    total = len(compare)
    matched = int(
        (compare["入庫品名"].fillna("").astype(str) != "").sum()
    )
    c1, c2, c3 = st.columns(3)
    c1.metric("總商品數", total)
    c2.metric("已匹配", matched)
    c3.metric("未匹配", total - matched)

# ── 重新自動匹配 / 清空 ─────────────────────────────────────

if st.button("🔄 重新掃描訂單（新增未匹配項目）", type="primary"):
    import io
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
        save_compare_table(updated)
        st.success(f"掃描完成！對照表共 {len(updated)} 筆")
        st.rerun()

# ── 可編輯對照表 ─────────────────────────────────────────────
if not compare.empty:
    st.subheader("對照表")
    st.caption(
        "在「入庫品名」欄用下拉選單選擇對應的入庫商品，按下方「✅ 確認並自動帶出貨號」會自動填入貨號與主貨號"
    )

    # 篩選
    filter_opt = st.radio(
        "篩選", ["全部", "未匹配", "已匹配"],
        horizontal=True, key="cmp_filter",
    )
    view = compare.copy()
    if filter_opt == "未匹配":
        view = view[(view["入庫品名"].isna()) | (view["入庫品名"].astype(str) == "")]
    elif filter_opt == "已匹配":
        view = view[(view["入庫品名"].notna()) & (view["入庫品名"].astype(str) != "")]

    if view.empty:
        st.info("沒有符合條件的項目")
    else:
        edited = st.data_editor(
            view,
            use_container_width=True,
            hide_index=True,
            column_config={
                "平台商品名稱": st.column_config.TextColumn(
                    "Orders 品名", disabled=True, width="large",
                ),
                "平台": st.column_config.TextColumn("平台", disabled=True, width="small"),
                "入庫品名": st.column_config.SelectboxColumn(
                    "入庫品名", options=stg_options, width="large",
                ),
                "貨號": st.column_config.TextColumn("貨號", disabled=True, width="medium"),
                "主貨號": st.column_config.TextColumn("主貨號", disabled=True, width="medium"),
            },
            column_order=["平台商品名稱", "平台", "入庫品名", "貨號", "主貨號"],
            key="compare_editor",
        )

        if st.button("✅ 確認並自動帶出貨號", type="primary"):
            # 根據選擇的入庫品名自動填入貨號、主貨號
            for idx in edited.index:
                stg_name = str(edited.at[idx, "入庫品名"]).strip()
                if stg_name and stg_name in stg_mapping:
                    edited.at[idx, "貨號"] = stg_mapping[stg_name]["貨號"]
                    edited.at[idx, "主貨號"] = stg_mapping[stg_name]["主貨號"]

            # 合併回完整對照表
            view_keys = set(view["平台商品名稱"].astype(str) + "||" + view["平台"].astype(str))
            unchanged = compare[
                ~(compare["平台商品名稱"].astype(str) + "||" + compare["平台"].astype(str)).isin(view_keys)
            ]
            saved = pd.concat([unchanged, edited], ignore_index=True)
            saved = saved.drop_duplicates(subset=["平台商品名稱", "平台"])
            save_compare_table(saved)
            st.success("已儲存！貨號已自動帶出")
            st.rerun()

else:
    st.info("對照表為空，請先至首頁「匯入平台訂單」上傳訂單資料")
