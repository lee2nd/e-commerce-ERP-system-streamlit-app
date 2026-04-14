import streamlit as st
import pandas as pd
from datetime import datetime, timezone, timedelta
from utils.styles import apply_global_styles
from utils.data_manager import (
    load_platform_orders,
    load_compare_table,
    load_storage,
    load_delivery,
    save_delivery,
    clear_delivery,
    load_combo_sku,
)

st.set_page_config(page_title="導出出庫", page_icon="📦", layout="wide")

apply_global_styles()
st.title("📦 導出出庫")

TZ_TAIPEI = timezone(timedelta(hours=8))

# 載入對照表與入庫資料，建立映射
compare = load_compare_table()
storage = load_storage()
combo = load_combo_sku()

# 從入庫建立 貨號 → {名稱, 規格, 主貨號, 單位成本} lookup
stg_sku_lookup: dict[str, dict] = {}
if not storage.empty:
    name_col = "商品名稱" if "商品名稱" in storage.columns else "名稱"
    cost_col = "單位成本" if "單位成本" in storage.columns else "單價"
    for _, row in storage.drop_duplicates(subset=["貨號"]).iterrows():
        sku = str(row.get("貨號", "")).strip()
        if sku:
            stg_sku_lookup[sku] = {
                "名稱": str(row.get(name_col, "")).strip(),
                "規格": str(row.get("規格", "")).strip(),
                "主貨號": str(row.get("主貨號", "")).strip(),
                "單位成本": float(row.get(cost_col, 0) or 0),
            }

# 從入庫建立 入庫品名 → {貨號, 主貨號} 映射（保留向後兼容）
stg_mapping: dict[str, dict] = {}
if not storage.empty:
    name_col = "商品名稱" if "商品名稱" in storage.columns else "名稱"
    for _, row in storage.drop_duplicates(subset=[name_col, "規格"]).iterrows():
        name = str(row.get(name_col, "")).strip()
        spec = str(row.get("規格", "")).strip()
        if name:
            label = f"{name}[{spec}]" if spec else name
            stg_mapping[label] = {
                "貨號": str(row.get("貨號", "")),
                "主貨號": str(row.get("主貨號", "")),
            }

# 組合貨號 → stg_mapping 加入組合品名
if not combo.empty:
    for combo_code in combo["組合貨號"].unique():
        sub = combo[combo["組合貨號"] == combo_code]
        parts = " + ".join(
            f"{r['原料貨號']}×{int(r['原料數量'])}" for _, r in sub.iterrows()
        )
        label = f"組合:{parts}"
        stg_mapping[label] = {
            "貨號": combo_code,
            "主貨號": combo_code.split("-")[0] if "-" in combo_code else combo_code,
        }

# 從對照表建立 (平台商品名稱, 平台) → 入庫品名 映射
compare_mapping: dict[tuple[str, str], str] = {}
# (平台商品名稱, 平台) → {貨號, 主貨號}
compare_sku_mapping: dict[tuple[str, str], dict] = {}
_NL = frozenset({"nan", "none", "nat", "<na>"})

def _cs(v) -> str:
    """Convert to clean string, treating NaN/None-like as empty."""
    s = str(v).strip()
    return "" if s.lower() in _NL else s

if not compare.empty:
    for _, row in compare.iterrows():
        key = (_cs(row.get("平台商品名稱", "")), _cs(row.get("平台", "")))
        stg_name = _cs(row.get("入庫品名", ""))
        if stg_name:
            compare_mapping[key] = stg_name
        compare_sku_mapping[key] = {
            "貨號": _cs(row.get("貨號", "")),
            "主貨號": _cs(row.get("主貨號", "")),
        }


def _n(val, default=0):
    """Safe numeric conversion"""
    try:
        v = pd.to_numeric(val, errors="coerce")
        return float(v) if pd.notna(v) else default
    except Exception:
        return default


def _filter_shopee(df: pd.DataFrame) -> pd.DataFrame:
    """跑皮：過濾完全退貨的商品列，保留有效成交的列"""
    if df.empty:
        return df
    mask = pd.Series(True, index=df.index)
    if "不成立原因" in df.columns:
        # 有不成立原因且不含「遺失」→ 跳過
        mask &= (
            df["不成立原因"].fillna("").astype(str).str.strip() == ""
        ) | df["不成立原因"].fillna("").astype(str).str.contains("遺失", na=False)
    if "退貨 / 退款狀態" in df.columns and "數量" in df.columns and "退貨數量" in df.columns:
        ret_stat = df["退貨 / 退款狀態"].fillna("").astype(str).str.strip()
        qty = pd.to_numeric(df["數量"], errors="coerce").fillna(0)
        ret_qty = pd.to_numeric(df["退貨數量"], errors="coerce").fillna(0)
        effective_qty = (qty - ret_qty).clip(lower=0)
        # 有退貨狀態 且 有效數量 == 0 → 跳過（完全退貨）
        fully_returned = (ret_stat != "") & (effective_qty == 0)
        mask &= ~fully_returned
    elif "退貨 / 退款狀態" in df.columns:
        # 沒有退貨數量欄→ 原邀輯
        mask &= df["退貨 / 退款狀態"].fillna("").astype(str).str.strip() == ""
    return df[mask].copy()


def _filter_ruten(df: pd.DataFrame) -> pd.DataFrame:
    """露天：交易狀況 有 '退貨' 字眼就跳過，訂單狀態 含 '取消' 也跳過"""
    if df.empty:
        return df
    mask = pd.Series(True, index=df.index)
    if "交易狀況" in df.columns:
        mask &= ~df["交易狀況"].fillna("").astype(str).str.contains("退貨", na=False)
    if "訂單狀態" in df.columns:
        mask &= ~df["訂單狀態"].fillna("").astype(str).str.contains("取消", na=False)
    return df[mask].copy()


def _filter_easystore(df: pd.DataFrame) -> pd.DataFrame:
    """官網：出貨前取消 (Fulfillment Service空+Restocked) 或 取消訂購 → 跳過不出庫"""
    if df.empty:
        return df.copy()
    _df = df.copy()
    # forward-fill Order Name first, then group-by ffill other order-level fields
    if "Order Name" in _df.columns:
        _df["Order Name"] = _df["Order Name"].ffill()
        _ff_cols = [c for c in ["Remark", "Fulfillment Service", "Fulfillment Status"] if c in _df.columns]
        if _ff_cols:
            _df[_ff_cols] = _df.groupby("Order Name")[_ff_cols].ffill()
    mask = pd.Series(True, index=_df.index)
    # 出貨前取消：Fulfillment Service 空 且 Fulfillment Status == Restocked
    if "Fulfillment Service" in _df.columns and "Fulfillment Status" in _df.columns:
        is_prestocked = (
            _df["Fulfillment Service"].fillna("").astype(str).str.strip() == ""
        ) & (_df["Fulfillment Status"].fillna("").astype(str) == "Restocked")
        mask &= ~is_prestocked
    # 取消訂購（未取貨）→ 不出庫
    if "Remark" in _df.columns:
        mask &= _df["Remark"].fillna("").astype(str).str.strip() != "取消訂購"
    return _df[mask].copy()


def _build_platform_key(row: pd.Series, platform: str) -> str:
    """根據平台組出對照表的 key (平台商品名稱)"""
    if platform == "蝦皮":
        name = str(row.get("商品名稱", "")).strip()
        spec = str(row.get("商品選項名稱", "")).strip()
        return f"{name}::{spec}"
    elif platform == "露天":
        name = str(row.get("商品名稱", "")).strip()
        spec1 = str(row.get("規格", "")).strip()
        spec2 = str(row.get("項目", "")).strip()
        return f"{name}::{spec1}::{spec2}"
    elif platform == "官網":
        name = str(row.get("Item Name", "")).strip()
        variant = str(row.get("Item Variant", "")).strip()
        return f"{name}::{variant}"
    return ""


def _get_order_data(row: pd.Series, platform: str) -> dict:
    """從訂單取得數量、單價、日期"""
    ret_qty = 0  # 僅蝦皮部份退貨時使用
    if platform == "蝦皮":
        qty = row.get("數量", 0)
        # 部份退貨：使用有效數量
        ret_qty_raw = row.get("退貨數量", 0)
        ret_qty = int(_n(ret_qty_raw)) if pd.notna(ret_qty_raw) else 0
        price = row.get("商品活動價格") if pd.notna(row.get("商品活動價格")) else row.get("商品原價", 0)
        date = str(row.get("訂單成立日期", ""))[:10]
    elif platform == "露天":
        qty = row.get("數量", 0)
        price = row.get("單價", 0)
        date = str(row.get("結帳時間", ""))[:10]
    elif platform == "官網":
        qty = row.get("Quantity", 0)
        price = row.get("Item Price", 0)
        date = str(row.get("Date", ""))[:10]
    else:
        qty, price, date = 0, 0, ""
    
    # 確保數值
    try:
        qty = int(qty) if pd.notna(qty) else 0
    except (ValueError, TypeError):
        qty = 0
    if platform == "蝦皮":
        qty = max(0, qty - ret_qty)
    try:
        price = float(price) if pd.notna(price) else 0
    except (ValueError, TypeError):
        price = 0
    
    if platform == "官網":
        order_no = str(row.get("Order Name", "")).strip()
    else:
        order_no = str(row.get("訂單編號", "")).strip()
    return {"訂單編號": order_no, "數量": qty, "單價": price, "日期": date}


def _get_row_sku(row: pd.Series, platform: str) -> str:
    """取得訂單行的原始 SKU，用於驗證是否在入庫中"""
    _nl = {"nan", "none", "nat", "<na>"}
    if platform == "蝦皮":
        raw = str(row.get("商品選項貨號", "")).strip() or str(row.get("主商品貨號", "")).strip()
    elif platform == "露天":
        raw = str(row.get("賣家自用料號", "")).strip()
    elif platform == "官網":
        raw = str(row.get("Item SKU", "")).strip()
    else:
        raw = ""
    return "" if raw.lower() in _nl else raw


def generate_delivery() -> pd.DataFrame:
    """產生出庫資料"""
    records = []
    
    platform_config = [
        ("蝦皮", _filter_shopee),
        ("露天", _filter_ruten),
        ("官網", _filter_easystore),
    ]
    
    for platform, filter_func in platform_config:
        raw = load_platform_orders(platform)
        if raw.empty:
            continue
        
        # 過濾無效訂單
        filtered = filter_func(raw)
        
        for _, row in filtered.iterrows():
            # 組出對照表 key
            plat_key = _build_platform_key(row, platform)
            if not plat_key:
                continue
            
            # 查對照表取得入庫品名
            stg_name = compare_mapping.get((plat_key, platform), "")
            # 入庫品名為空：對照表內完全沒有這筆訂單記錄，視為未匹配
            if not stg_name:
                stg_name = "未匹配"

            if stg_name == "未匹配":
                # 入庫品名為未匹配：用平台原始資料
                sku_info = compare_sku_mapping.get((plat_key, platform), {})
                sku = sku_info.get("貨號", "")
                main_sku = sku_info.get("主貨號", "")
                if platform == "蝦皮":
                    prod_name = str(row.get("商品名稱", "")).strip()
                    prod_spec = str(row.get("商品選項名稱", "")).strip()
                elif platform == "露天":
                    prod_name = str(row.get("商品名稱", "")).strip()
                    s1 = str(row.get("規格", "")).strip()
                    s2 = str(row.get("項目", "")).strip()
                    prod_spec = f"{s1}::{s2}".strip("::") if s1 or s2 else ""
                elif platform == "官網":
                    prod_name = str(row.get("Item Name", "")).strip()
                    prod_spec = str(row.get("Item Variant", "")).strip()
                else:
                    prod_name = plat_key
                    prod_spec = ""
                is_unmatched = True
            elif stg_name.startswith("組合:"):
                # 組合貨號：展開為原料明細，每個原料獨立一筆出庫記錄
                sku_info = compare_sku_mapping.get((plat_key, platform), {})
                combo_code = sku_info.get("貨號", "")
                if not combo_code:
                    fb = stg_mapping.get(stg_name, {})
                    combo_code = fb.get("貨號", "")

                order_data = _get_order_data(row, platform)
                order_qty = order_data["數量"]

                if not combo.empty and combo_code:
                    components = combo[combo["組合貨號"].astype(str).str.strip() == combo_code]
                    for _, comp in components.iterrows():
                        mat_sku = str(comp["原料貨號"]).strip()
                        mat_qty = int(comp["原料數量"])
                        total_mat_qty = order_qty * mat_qty
                        comp_info = stg_sku_lookup.get(mat_sku, {})
                        comp_name = comp_info.get("名稱", mat_sku)
                        comp_spec = comp_info.get("規格", "")
                        comp_main = comp_info.get("主貨號", mat_sku.split("-")[0] if "-" in mat_sku else mat_sku)
                        comp_cost = comp_info.get("單位成本", 0)
                        records.append({
                            "訂單編號": order_data["訂單編號"],
                            "主貨號": comp_main,
                            "貨號": mat_sku,
                            "名稱": comp_name,
                            "規格": comp_spec,
                            "出庫數量": total_mat_qty,
                            "單價": round(comp_cost, 2),
                            "金額": round(comp_cost * total_mat_qty, 2),
                            "出庫日期": order_data["日期"],
                            "匹配狀態": "已匹配",
                            "平台": platform,
                        })
                continue  # skip the normal records.append below
            else:
                # 一般匹配流程
                if "[" in stg_name and stg_name.endswith("]"):
                    parts = stg_name.rsplit("[", 1)
                    prod_name = parts[0]
                    prod_spec = parts[1].rstrip("]")
                else:
                    prod_name = stg_name
                    prod_spec = ""
                sku_info = stg_mapping.get(stg_name, {})
                sku = sku_info.get("貨號", "")
                main_sku = sku_info.get("主貨號", "")
                is_unmatched = not bool(sku_info)

                # 進一步驗證：對照表雖顯示已匹配（名稱找到），但訂單原始 SKU 不在入庫中
                # → 視為未匹配，與日報表邏輯保持一致
                if not is_unmatched:
                    raw_order_sku = _get_row_sku(row, platform)
                    if raw_order_sku and raw_order_sku not in stg_sku_lookup:
                        _combo_codes = set(combo["組合貨號"].astype(str).str.strip()) if not combo.empty else set()
                        if raw_order_sku not in _combo_codes:
                            if platform == "蝦皮":
                                prod_name = str(row.get("商品名稱", "")).strip()
                                prod_spec = str(row.get("商品選項名稱", "")).strip()
                            elif platform == "露天":
                                prod_name = str(row.get("商品名稱", "")).strip()
                                _s1 = str(row.get("規格", "")).strip()
                                _s2 = str(row.get("項目", "")).strip()
                                prod_spec = f"{_s1}::{_s2}".strip("::") if _s1 or _s2 else ""
                            elif platform == "官網":
                                prod_name = str(row.get("Item Name", "")).strip()
                                prod_spec = str(row.get("Item Variant", "")).strip()
                            sku = raw_order_sku
                            main_sku = raw_order_sku.split("-")[0] if "-" in raw_order_sku else raw_order_sku
                            is_unmatched = True

            # 取得訂單資料
            order_data = _get_order_data(row, platform)

            _price = round(order_data["單價"], 2)
            records.append({
                "訂單編號": order_data["訂單編號"],
                "主貨號": main_sku,
                "貨號": sku,
                "名稱": prod_name,
                "規格": prod_spec,
                "出庫數量": order_data["數量"],
                "單價": _price,
                "金額": round(order_data["數量"] * _price, 2),
                "出庫日期": order_data["日期"],
                "匹配狀態": "未匹配" if is_unmatched else "已匹配",
                "平台": platform,
            })
    
    if records:
        df = pd.DataFrame(records)
        # 統一日期格式後依日期排序（兼容 YYYY/MM/DD 與 YYYY-MM-DD）
        df["出庫日期"] = df["出庫日期"].apply(
            lambda v: pd.to_datetime(v, dayfirst=False, errors="coerce").strftime("%Y-%m-%d")
            if pd.notna(pd.to_datetime(v, dayfirst=False, errors="coerce")) else ""
        )
        df = df.sort_values("出庫日期").reset_index(drop=True)
        return df
    return pd.DataFrame()


# 統計現有出庫資料
existing_delivery = load_delivery()
if not existing_delivery.empty:

    colors = {"蝦皮": "#FF6B35", "露天": "#4A90D9", "官網": "#2ECC71"}

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("總筆數", len(existing_delivery))
    c2.metric("總金額", f"${existing_delivery['金額'].sum():,.0f}")

    # 定義一個共用的顯示函數，方便重複使用
    def custom_metric(label, value, color, column):
        column.markdown(
            f"""
            <div style="border-left: 3px solid {color}; padding-left: 10px; border-radius: 4px;">
                <p style="font-size: 14px; color: #606060; margin-bottom: 0px;">{label}</p>
                <p style="font-size: 24px; font-weight: bold; color: {color}; margin-top: -5px;">{value}</p>
            </div>
            """,
            unsafe_allow_html=True
        )
    # 3. 蝦皮
    shopee_count = len(existing_delivery[existing_delivery["平台"] == "蝦皮"])
    custom_metric("蝦皮", shopee_count, colors["蝦皮"], c3)
    # 4. 露天
    ruten_count = len(existing_delivery[existing_delivery["平台"] == "露天"])
    custom_metric("露天", ruten_count, colors["露天"], c4)
    # 5. 官網
    official_count = len(existing_delivery[existing_delivery["平台"] == "官網"])
    custom_metric("官網", official_count, colors["官網"], c5)


# 顯示對照表匹配狀態
if compare.empty:
    st.warning("⚠️ 對照表為空，請先至「對照表管理」頁面建立商品對照")
else:
    unmatched_count = (compare["入庫品名"].fillna("").astype(str) == "未匹配").sum()
    total_count = len(compare)
    if unmatched_count > 0:
        st.info(f"📋 對照表：{unmatched_count}/{total_count} 筆未匹配入庫")

# 導出出庫按鈕
if st.button("🚀 導出出庫", type="primary"):
    with st.spinner("正在產生出庫資料..."):
        new_delivery = generate_delivery()
    
    if new_delivery.empty:
        st.warning("無法產生出庫資料，請確認：\n1. 已匯入平台訂單\n2. 對照表已建立商品對照")
    else:
        # 每次重建，全欄位去重
        combined = new_delivery.drop_duplicates(keep="last").reset_index(drop=True)
        save_delivery(combined)
        st.session_state["delivery_saved_at"] = datetime.now(tz=TZ_TAIPEI).strftime("%Y-%m-%d %H:%M:%S")
        st.success(f"✅ 出庫資料已產生！共 {len(combined)} 筆")
        st.rerun()

_delivery_ts = st.session_state.get("delivery_saved_at")
if _delivery_ts:
    st.caption(f"🕐 最後導出：{_delivery_ts}")

# 顯示出庫資料
st.subheader("📋 出庫資料")
delivery = load_delivery()
if delivery.empty:
    st.info("尚無出庫資料，請點擊上方「導出出庫」按鈕產生")
else:
    # 篩選器
    _f1, _f2, _f3 = st.columns([2, 2, 3])
    with _f1:
        match_filter = st.radio("匹配狀態", options=["全部", "已匹配", "未匹配"], horizontal=True)
    with _f2:
        plat_options = sorted(delivery["平台"].dropna().unique().tolist()) if "平台" in delivery.columns else []
        plat_filter = st.multiselect("平台", options=plat_options, default=plat_options)
    with _f3:
        if "出庫日期" in delivery.columns:
            _dates = pd.to_datetime(delivery["出庫日期"], errors="coerce").dropna()
            _min_date = _dates.min().date() if not _dates.empty else None
            _max_date = _dates.max().date() if not _dates.empty else None
            date_range = st.date_input("日期範圍", value=(_min_date, _max_date) if _min_date else [], key="dlv_date_range")
        else:
            date_range = []

    view_dlv = delivery.copy()
    if "匹配狀態" in view_dlv.columns:
        if match_filter == "未匹配":
            view_dlv = view_dlv[view_dlv["匹配狀態"] == "未匹配"]
        elif match_filter == "已匹配":
            view_dlv = view_dlv[view_dlv["匹配狀態"] == "已匹配"]
    if "平台" in view_dlv.columns:
        view_dlv = view_dlv[view_dlv["平台"].isin(plat_filter)]
    if "出庫日期" in view_dlv.columns and isinstance(date_range, (list, tuple)) and len(date_range) == 2:
        _start, _end = pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1])
        _dt = pd.to_datetime(view_dlv["出庫日期"], errors="coerce")
        view_dlv = view_dlv[(_dt >= _start) & (_dt <= _end)]
    # 欄位順序：平台移至最後
    _cols = [c for c in view_dlv.columns if c != "平台"] + (["平台"] if "平台" in view_dlv.columns else [])
    view_dlv = view_dlv[_cols]

    # 平台顏色標註
    _PLAT_COLORS = {"蝦皮": "#FF6B35", "露天": "#4A90D9", "官網": "#2ECC71"}

    def _highlight_row(row):
        color = _PLAT_COLORS.get(row.get("平台", ""), "")
        result = []
        for c in row.index:
            result.append(f"background-color: {color}20; color: {color}" if c == "平台"
                            else f"background-color: {color}10")
        return result

    _dlv_page_size = 500
    _dlv_total = len(view_dlv)
    _dlv_total_pages = max(1, (_dlv_total - 1) // _dlv_page_size + 1)
    _dlv_dl_col, _dlv_pg_col = st.columns([1, 3])
    _dlv_csv = view_dlv.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    _dlv_dl_col.download_button("⬇️ 下載出庫資料", data=_dlv_csv, file_name="出庫資料.csv", mime="text/csv", key="dl_dlv")
    if _dlv_total_pages > 1:
        _dlv_page = _dlv_pg_col.selectbox(
            "頁碼", list(range(1, _dlv_total_pages + 1)),
            format_func=lambda x: f"{x}/{_dlv_total_pages} 頁",
            key="dlv_page",
        )
    else:
        _dlv_page = 1
    _dlv_start = (_dlv_page - 1) * _dlv_page_size
    view_dlv_slice = view_dlv.iloc[_dlv_start:_dlv_start + _dlv_page_size].copy()

    styled = view_dlv_slice.style.apply(_highlight_row, axis=1)
    _money_cfg = {
        c: st.column_config.NumberColumn(format="$%.2f")
        for c in ["單價", "金額"] if c in view_dlv_slice.columns
    }
    st.dataframe(styled, width='stretch', hide_index=True, column_config=_money_cfg)
    st.caption(f"第 {_dlv_page} 頁 / 共 {_dlv_total_pages} 頁（{_dlv_total:,} 筆）")

# ── 清0 出庫資料 ───────────────────────────────────────────
st.markdown("---")
if st.button("🗑️ 清除出庫資料", key="clear_dlv_btn"):
    st.session_state["confirm_clear_dlv"] = True
if st.session_state.get("confirm_clear_dlv"):
    st.warning("⚠️ 確定要清除所有出庫資料嗎？此操作無法復原！")
    _c1, _c2 = st.columns(2)
    if _c1.button("✅ 確認清除", key="confirm_clear_dlv_yes", type="primary"):
        with st.spinner("清除中…"):
            clear_delivery()
        st.session_state.pop("confirm_clear_dlv", None)
        st.success("✅ 出庫資料已清除")
        st.rerun()
    if _c2.button("❌ 取消", key="confirm_clear_dlv_no"):
        st.session_state.pop("confirm_clear_dlv", None)
        st.rerun()
