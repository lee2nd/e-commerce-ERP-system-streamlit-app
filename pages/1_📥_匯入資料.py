"""匯入資料：上傳平台訂單 / 入庫資料"""
import streamlit as st
import pandas as pd

st.set_page_config(page_title="匯入資料", page_icon="📥", layout="wide")

from utils.data_manager import (
    load_orders, append_orders, save_orders,
    load_storage, save_storage,
    load_compare_table, save_compare_table,
)
from utils.parsers import parse_shopee, parse_ruten, parse_easystore, read_file_flexible
from utils.calculators import auto_match_compare_table

st.title("📥 匯入資料")

tab_order, tab_storage = st.tabs(["📦 匯入平台訂單", "📥 匯入 / 管理入庫資料"])

# ══════════════════════════════════════════════════════════════
# Tab 1 – 匯入平台訂單
# ══════════════════════════════════════════════════════════════
with tab_order:
    platform = st.selectbox("選擇平台", ["蝦皮", "露天", "官網 (EasyStore)"])
    uploaded = st.file_uploader(
        "上傳訂單檔案（.xlsx / .xls / .csv）",
        type=["xlsx", "xls", "csv"],
        key="order_upload",
    )

    if uploaded and st.button("🚀 開始匯入", type="primary"):
        try:
            with st.spinner("解析中…"):
                if "蝦皮" in platform:
                    new = parse_shopee(uploaded)
                elif "露天" in platform:
                    new = parse_ruten(uploaded)
                else:
                    new = parse_easystore(uploaded)

            if new.empty:
                st.warning("未解析到任何訂單")
            else:
                append_orders(new)

                # 自動更新對照表
                storage = load_storage()
                compare = load_compare_table()
                updated = auto_match_compare_table(new, storage, compare)
                save_compare_table(updated)

                st.success(f"✅ 成功匯入 **{len(new)}** 筆訂單")
                st.dataframe(new.head(30), use_container_width=True, hide_index=True)

                if not updated.empty:
                    matched = int((updated["貨號"].notna() & (updated["貨號"] != "")).sum())
                    st.info(f"對照表：{matched} / {len(updated)} 商品已自動匹配貨號")
        except Exception as e:
            st.error(f"匯入失敗：{e}")
            st.info("💡 蝦皮資料若讀取失敗，請嘗試從賣家中心匯出 **CSV** 格式")

    st.markdown("---")
    st.subheader("目前已匯入的訂單")
    orders = load_orders()
    if not orders.empty:
        cols = st.columns(3)
        for i, plat in enumerate(["蝦皮", "露天", "官網"]):
            with cols[i]:
                cnt = len(orders[orders["平台"] == plat])
                st.metric(plat, f"{cnt} 筆")
        st.dataframe(orders, use_container_width=True, hide_index=True)

        if st.button("🗑️ 清空所有訂單"):
            save_orders(pd.DataFrame())
            st.rerun()
    else:
        st.info("尚未匯入任何訂單")

# ══════════════════════════════════════════════════════════════
# Tab 2 – 入庫管理
# ══════════════════════════════════════════════════════════════
with tab_storage:
    st.subheader("上傳入庫檔案")
    st.markdown(
        "欄位格式：`主貨號, 商品名稱, 規格, 貨號, 數量, 單位成本`（總金額 / 入庫日期可選）"
    )

    stg_file = st.file_uploader(
        "上傳入庫檔案（.xlsx / .csv）",
        type=["xlsx", "xls", "csv"],
        key="stg_upload",
    )
    if stg_file and st.button("匯入入庫", type="primary"):
        try:
            new_stg = read_file_flexible(stg_file)
            required = ["主貨號", "商品名稱", "規格", "貨號", "數量", "單位成本"]
            missing = [c for c in required if c not in new_stg.columns]
            if missing:
                st.error(f"缺少必要欄位：{', '.join(missing)}")
            else:
                if "總金額" not in new_stg.columns:
                    new_stg["總金額"] = new_stg["數量"] * new_stg["單位成本"]
                if "入庫日期" not in new_stg.columns:
                    new_stg["入庫日期"] = pd.Timestamp.now().strftime("%Y-%m-%d")
                existing = load_storage()
                combined = pd.concat([existing, new_stg], ignore_index=True).drop_duplicates()
                save_storage(combined)
                st.success(f"✅ 匯入 {len(new_stg)} 筆入庫資料")
        except Exception as e:
            st.error(f"匯入失敗：{e}")

    # 手動新增
    st.markdown("---")
    st.subheader("手動新增入庫")
    with st.form("add_stg", clear_on_submit=True):
        r1 = st.columns(4)
        main_sku = r1[0].text_input("主貨號")
        name     = r1[1].text_input("商品名稱")
        spec     = r1[2].text_input("規格")
        sku      = r1[3].text_input("貨號")
        r2 = st.columns(3)
        qty       = r2[0].number_input("數量", min_value=1, value=1)
        unit_cost = r2[1].number_input("單位成本", min_value=0.0, step=0.1)
        stg_date  = r2[2].date_input("入庫日期")
        if st.form_submit_button("➕ 新增"):
            if not all([main_sku, name, sku]):
                st.error("請至少填寫主貨號、商品名稱、貨號")
            else:
                row = pd.DataFrame([{
                    "主貨號": main_sku, "商品名稱": name, "規格": spec,
                    "貨號": sku, "數量": qty, "單位成本": unit_cost,
                    "總金額": qty * unit_cost, "入庫日期": str(stg_date),
                }])
                existing = load_storage()
                save_storage(pd.concat([existing, row], ignore_index=True))
                st.success("新增成功")
                st.rerun()

    st.markdown("---")
    st.subheader("目前入庫資料")
    storage = load_storage()
    if not storage.empty:
        st.dataframe(storage, use_container_width=True, hide_index=True)
        if st.button("🗑️ 清空入庫"):
            save_storage(pd.DataFrame())
            st.rerun()
    else:
        st.info("尚未有入庫資料")
