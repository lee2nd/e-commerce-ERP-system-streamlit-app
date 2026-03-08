import streamlit as st
import pandas as pd
from utils.data_manager import (
    load_platform_orders,
    load_compare_table,
    load_storage,
    load_delivery,
    save_delivery,
)

st.set_page_config(page_title="產生出庫", page_icon="📦", layout="wide")
st.title("📦 產生出庫")

# 載入對照表與入庫資料，建立映射
compare = load_compare_table()
storage = load_storage()

# 從入庫建立 入庫品名 → {貨號, 主貨號} 映射
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

# 從對照表建立 (平台商品名稱, 平台) → 入庫品名 映射
compare_mapping: dict[tuple[str, str], str] = {}
if not compare.empty:
    for _, row in compare.iterrows():
        key = (str(row.get("平台商品名稱", "")).strip(), str(row.get("平台", "")).strip())
        stg_name = str(row.get("入庫品名", "")).strip()
        if stg_name:
            compare_mapping[key] = stg_name


def _filter_shopee(df: pd.DataFrame) -> pd.DataFrame:
    """蝦皮：不成立原因、退貨 / 退款狀態 欄位有內容就跳過"""
    if df.empty:
        return df
    mask = pd.Series(True, index=df.index)
    if "不成立原因" in df.columns:
        mask &= df["不成立原因"].fillna("").astype(str).str.strip() == ""
    if "退貨 / 退款狀態" in df.columns:
        mask &= df["退貨 / 退款狀態"].fillna("").astype(str).str.strip() == ""
    return df[mask].copy()


def _filter_ruten(df: pd.DataFrame) -> pd.DataFrame:
    """露天：交易狀況 有 '退貨' 字眼就跳過"""
    if df.empty:
        return df
    if "交易狀況" not in df.columns:
        return df
    mask = ~df["交易狀況"].fillna("").astype(str).str.contains("退貨", na=False)
    return df[mask].copy()


def _filter_easystore(df: pd.DataFrame) -> pd.DataFrame:
    """官網：目前不做特殊過濾"""
    return df.copy() if not df.empty else df


def _build_platform_key(row: pd.Series, platform: str) -> str:
    """根據平台組出對照表的 key (平台商品名稱)"""
    if platform == "蝦皮":
        name = str(row.get("商品名稱", "")).strip()
        spec = str(row.get("商品選項名稱", "")).strip()
        return f"{name}::{spec}" if spec else name
    elif platform == "露天":
        name = str(row.get("商品名稱", "")).strip()
        spec1 = str(row.get("規格", "")).strip()
        spec2 = str(row.get("項目", "")).strip()
        return f"{name}::{spec1}::{spec2}" if spec1 or spec2 else name
    elif platform == "官網":
        name = str(row.get("Item Name", "")).strip()
        variant = str(row.get("Item Variant", "")).strip()
        return f"{name}::{variant}" if variant else name
    return ""


def _get_order_data(row: pd.Series, platform: str) -> dict:
    """從訂單取得數量、單價、日期"""
    if platform == "蝦皮":
        qty = row.get("數量", 0)
        # 優先用商品活動價格，沒有則用商品原價
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
    try:
        price = float(price) if pd.notna(price) else 0
    except (ValueError, TypeError):
        price = 0
    
    return {"數量": qty, "單價": price, "日期": date}


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
            if not stg_name:
                continue  # 未匹配則跳過
            
            # 從入庫品名拆出名稱與規格
            if "[" in stg_name and stg_name.endswith("]"):
                parts = stg_name.rsplit("[", 1)
                prod_name = parts[0]
                prod_spec = parts[1].rstrip("]")
            else:
                prod_name = stg_name
                prod_spec = ""
            
            # 查入庫表取得貨號、主貨號
            sku_info = stg_mapping.get(stg_name, {})
            sku = sku_info.get("貨號", "")
            main_sku = sku_info.get("主貨號", "")
            
            # 取得訂單資料
            order_data = _get_order_data(row, platform)
            
            records.append({
                "主貨號": main_sku,
                "名稱": prod_name,
                "規格": prod_spec,
                "數量": order_data["數量"],
                "單價": order_data["單價"],
                "金額": order_data["數量"] * order_data["單價"],
                "日期": order_data["日期"],
                "平台": platform,
                "貨號": sku,
            })
    
    if records:
        df = pd.DataFrame(records)
        # 依日期排序
        df = df.sort_values("日期").reset_index(drop=True)
        return df
    return pd.DataFrame()


# 統計現有出庫資料
existing_delivery = load_delivery()
if not existing_delivery.empty:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("總筆數", len(existing_delivery))
    c2.metric("總金額", f"${existing_delivery['金額'].sum():,.0f}")
    c3.metric("蝦皮", len(existing_delivery[existing_delivery["平台"] == "蝦皮"]))
    c4.metric("露天/官網", len(existing_delivery[existing_delivery["平台"] != "蝦皮"]))

# 顯示對照表匹配狀態
if compare.empty:
    st.warning("⚠️ 對照表為空，請先至「對照表管理」頁面建立商品對照")
else:
    matched_count = (compare["入庫品名"].fillna("").astype(str) != "").sum()
    total_count = len(compare)
    if matched_count < total_count:
        st.info(f"📋 對照表：{matched_count}/{total_count} 筆已匹配，未匹配的商品將不會出現在出庫資料中")

# 導出出庫按鈕
if st.button("🚀 導出出庫", type="primary"):
    with st.spinner("正在產生出庫資料..."):
        new_delivery = generate_delivery()
    
    if new_delivery.empty:
        st.warning("無法產生出庫資料，請確認：\n1. 已匯入平台訂單\n2. 對照表已建立商品對照")
    else:
        # 與現有資料合併，去重
        if not existing_delivery.empty:
            combined = pd.concat([existing_delivery, new_delivery], ignore_index=True)
            # 全欄位去重
            combined = combined.drop_duplicates(keep="last").reset_index(drop=True)
        else:
            combined = new_delivery
        
        save_delivery(combined)
        st.success(f"✅ 出庫資料已產生！共 {len(combined)} 筆")
        st.rerun()

# 顯示出庫資料
st.subheader("📋 出庫資料")
delivery = load_delivery()
if delivery.empty:
    st.info("尚無出庫資料，請點擊上方「導出出庫」按鈕產生")
else:
    # 平台顏色標註
    _PLAT_COLORS = {"蝦皮": "#FF6B35", "露天": "#4A90D9", "官網": "#2ECC71"}
    
    def _highlight_platform(row):
        color = _PLAT_COLORS.get(row.get("平台", ""), "")
        if color:
            return [f"background-color: {color}20; color: {color}" if c == "平台"
                    else f"background-color: {color}10" for c in row.index]
        return [""] * len(row)
    
    styled = delivery.style.apply(_highlight_platform, axis=1)
    st.dataframe(styled, use_container_width=True, hide_index=True)
