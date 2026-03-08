import streamlit as st
import pandas as pd
import io
from utils.data_manager import (
    load_compare_table, save_compare_table,
    load_storage,
    load_platform_orders,
)
from utils.parsers import parse_shopee, parse_ruten, parse_easystore
from utils.calculators import auto_match_compare_table

st.set_page_config(page_title="對照表", page_icon="📋", layout="wide")
st.title("📋 對照表")

compare = load_compare_table()
storage = load_storage()

# 從入庫.xlsx 建立「入庫品名」下拉清單與貨號映射
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

# 統計
if not compare.empty:
    total = len(compare)
    matched = int(
        (compare["入庫品名"].fillna("").astype(str) != "").sum()
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

# user 編輯對照表
if not compare.empty:
    st.caption(
        "在「入庫品名」欄用下拉選單選擇對應的入庫商品，確認入庫品名都填寫完，按下方「✅ 確認並自動帶出貨號」會自動填入貨號與主貨號"
    )
    filter_opt = st.radio(
            "篩選狀態", ["全部", "未匹配", "已匹配"], 
            horizontal=True, key="cmp_filter",
            label_visibility="collapsed",
        )

    # 在複製給 view 之前，強制將這些文字欄位轉為字串型態，避免 float 轉換錯誤
    cols_to_str = ["入庫品名", "主貨號", "貨號", "平台商品名稱"]
    for col in cols_to_str:
        if col in compare.columns:
            # fillna("") 把 NaN 換成空字串，astype(str) 強制轉換為字串
            compare[col] = compare[col].fillna("").astype(str)
                
    view = compare.copy()
    if filter_opt == "未匹配":
        view = view[(view["入庫品名"].isna()) | (view["入庫品名"].astype(str) == "")]
    elif filter_opt == "已匹配":
        view = view[(view["入庫品名"].notna()) & (view["入庫品名"].astype(str) != "")]

    if view.empty:
        st.info("沒有符合條件的項目")
    else:
        # 平台顏色標註
        _PLAT_COLORS = {"蝦皮": "#FF6B35", "露天": "#4A90D9", "官網": "#2ECC71"}

        def _highlight_platform(row):
            color = _PLAT_COLORS.get(row["平台"], "")
            if color:
                return [f"background-color: {color}20; color: {color}" if c == "平台"
                        else f"background-color: {color}10" for c in row.index]
            return [""] * len(row)

        styled_view = view.style.apply(_highlight_platform, axis=1)

        # 💡 解法 1：使用 st.form 包裝編輯器與按鈕，防止狀態中途遺失
        with st.form("compare_editor_form"):
            edited = st.data_editor(
                styled_view,
                width="stretch",
                hide_index=True,
                column_config={
                    "平台": st.column_config.TextColumn("平台", disabled=True, width="small"),
                    "平台商品名稱": st.column_config.TextColumn(
                        "平台商品名稱", disabled=True, width="large",
                    ),
                    "入庫品名": st.column_config.SelectboxColumn(
                        "入庫品名", options=stg_options, width="large",
                    ),
                    "主貨號": st.column_config.TextColumn("主貨號", disabled=True, width="medium"),
                    "貨號": st.column_config.TextColumn("貨號", disabled=True, width="medium"),
                },
                column_order=["平台", "平台商品名稱", "入庫品名", "主貨號", "貨號"],
                key="compare_editor",
            )

            # 使用 form_submit_button 替代原有的 button
            submitted = st.form_submit_button("✅ 儲存入庫品名 → 自動帶出貨號", type="primary")

            if submitted:
                # 💡 解法 2：直接依據 Index 更新 compare 表，避免使用 concat 導致欄位遺失
                compare.loc[edited.index, "入庫品名"] = edited["入庫品名"]

                # 根據已儲存的入庫品名自動填入貨號、主貨號
                for idx in edited.index:
                    raw_val = compare.at[idx, "入庫品名"]
                    stg_name = "" if pd.isna(raw_val) else str(raw_val).strip()
                    
                    if stg_name and stg_name in stg_mapping:
                        compare.at[idx, "貨號"] = stg_mapping[stg_name]["貨號"]
                        compare.at[idx, "主貨號"] = stg_mapping[stg_name]["主貨號"]
                    else:
                        # (可選) 如果使用者清空了入庫品名，順便把貨號也清空
                        compare.at[idx, "貨號"] = ""
                        compare.at[idx, "主貨號"] = ""

                # 儲存結果並重新整理畫面
                # 記得把這行的註解拿掉！
                save_compare_table(compare)
                st.success("已儲存！貨號已自動帶出")
                st.rerun()

else:
    st.info("對照表為空，請先至首頁「匯入平台訂單」上傳訂單資料")
