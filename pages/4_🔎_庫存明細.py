import streamlit as st
import pandas as pd
from utils.data_manager import (
    load_storage, load_delivery,
    load_compare_table,
    load_inventory_details, save_inventory_details,
    clear_inventory_details,
    load_combo_sku,
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


def extract_color(spec: str) -> str:
    """Return the spec string with the size token removed (i.e. the colour part)."""
    spec_text = "" if pd.isna(spec) else str(spec)
    for size in reversed(SIZE_ORDER):  # remove 2XL before XL to avoid partial replacement
        spec_text = spec_text.replace(size, "")
    return spec_text.strip()

st.set_page_config(page_title="庫存明細", page_icon="🔎", layout="wide")
st.title("🔎 庫存明細")

if st.button("🔄 更新庫存明細", type="primary"):
    storage  = load_storage()
    delivery = load_delivery()
    compare  = load_compare_table()
    combo    = load_combo_sku()
    if storage.empty:
        st.warning("請先至「匯入資料」頁面新增入庫資料")
    else:
        result = generate_inventory_details(storage, delivery, combo)
        # 過濾掉對照表內未匹配的項目（入庫品名為空或"未匹配"）
        if not compare.empty and "貨號" in compare.columns and "入庫品名" in compare.columns:
            matched_skus = set(
                compare.loc[
                    ~compare["入庫品名"].fillna("").astype(str).isin(["", "未匹配"]),
                    "貨號"
                ].astype(str).str.strip()
            )
            # 組合貨號的原料 SKU 也須納入庫存明細
            if not combo.empty and "組合貨號" in combo.columns and "原料貨號" in combo.columns:
                combo_codes_matched = matched_skus & set(combo["組合貨號"].astype(str).str.strip())
                if combo_codes_matched:
                    component_skus = set(
                        combo.loc[
                            combo["組合貨號"].astype(str).str.strip().isin(combo_codes_matched),
                            "原料貨號"
                        ].astype(str).str.strip()
                    )
                    matched_skus |= component_skus
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
    view = (
        view
        .assign(
            _color=view["規格"].map(extract_color),
            _size_rank=view["規格"].map(size_sort_key),
        )
        .sort_values(by=["主貨號", "_color", "_size_rank"])
        .drop(columns=["_color", "_size_rank"])
    )
    view = view.reset_index(drop=True)

_inv_page_size = 500
_inv_total = len(view)
_inv_total_pages = max(1, (_inv_total - 1) // _inv_page_size + 1)
_inv_dl_col, _inv_pg_col = st.columns([1, 3])
_inv_csv = view.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
_inv_dl_col.download_button("⬇️ 下載庫存明細", data=_inv_csv, file_name="庫存明細.csv", mime="text/csv", key="dl_inv")
if _inv_total_pages > 1:
    _inv_page = _inv_pg_col.selectbox(
        "頁碼", list(range(1, _inv_total_pages + 1)),
        format_func=lambda x: f"{x}/{_inv_total_pages} 頁",
        key="inv_page",
    )
else:
    _inv_page = 1
_inv_start = (_inv_page - 1) * _inv_page_size
st.dataframe(view.iloc[_inv_start:_inv_start + _inv_page_size], width='stretch', hide_index=True)
st.caption(f"第 {_inv_page} 頁 / 共 {_inv_total_pages} 頁（{_inv_total:,} 筆）")


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

