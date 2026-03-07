"""首頁：匯入資料（入庫 / 平台訂單）"""
import streamlit as st
import pandas as pd
from pathlib import Path

st.set_page_config(page_title="電商平台進銷存系統", page_icon="📊", layout="wide")

from utils.data_manager import (
    load_orders, append_orders, save_orders,
    load_storage, save_storage,
    load_compare_table, save_compare_table,
    load_platform_orders, append_platform_orders,
)
from utils.parsers import parse_shopee, parse_ruten, parse_easystore, read_file_flexible
from utils.calculators import auto_match_compare_table

# 調整元件的樣式
st.markdown("""
    <style>
    [data-testid="stFileUploader"] {
        max-width: 150px;
    }
    [data-testid="stFileUploaderDropzoneInstructions"] {
        display: none;
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <style>
    [data-testid="stSelectbox"] {
        max-width: 150px;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📊 電商平台 ERP & 報表系統")
st.caption("蝦皮 ｜ 露天 ｜ 官網 (EasyStore) ｜ MOMO")
st.markdown("---")

# ══════════════════════════════════════════════════════════════
# Tab 切換
# ══════════════════════════════════════════════════════════════
tab_storage, tab_order = st.tabs(["📥 匯入入庫資料", "📦 匯入平台訂單"])

# ══════════════════════════════════════════════════════════════
# Tab 1 – 匯入入庫資料
# ══════════════════════════════════════════════════════════════
TEMPLATE_PATH = Path(__file__).resolve().parent / "data" / "入庫.xlsx"

with tab_storage:

    st.subheader("目前入庫資料 (後台使用)")
    storage = load_storage()
    if not storage.empty:
        st.dataframe(storage, use_container_width=True, hide_index=True)
    else:
        st.info("尚未有入庫資料")
            
    st.subheader("手動 EXCEL 新增入庫")
    st.markdown(
        "請先下載 **入庫.xlsx** ，在 Excel 中填寫完最新資料後再上傳匯入，請勿刪到舊的資料，不然會比對不到。"
    )
    if TEMPLATE_PATH.exists():
        with open(TEMPLATE_PATH, "rb") as f:
            st.download_button(
                label="⬇️ 下載入庫",
                data=f,
                file_name="入庫.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            )
    else:
        st.warning("找不到範本檔案 data/入庫.xlsx")

    stg_file = st.file_uploader(
        "⬆ 上傳入庫",
        type=["xlsx", "xls", "csv"],
        key="stg_upload",
    )

    if stg_file:
        try:
            # 讀取上傳的檔案
            raw_df = read_file_flexible(stg_file)
            st.markdown("#### 預覽上傳的資料")
            st.dataframe(raw_df, use_container_width=True, hide_index=True)

            if st.button("🚀 確認匯入入庫資料", type="primary"):
                # 檢查必要欄位（入庫.xlsx 格式）
                required_cols = ["主貨號", "貨號", "名稱", "規格", "入庫數量", "單價", "金額", "入庫日期"]
                missing = [c for c in required_cols if c not in raw_df.columns]
                if missing:
                    st.error(f"缺少必要欄位：{', '.join(missing)}")
                else:
                    # 對應欄位到系統格式
                    col_map = {
                        "名稱": "商品名稱",
                        "入庫數量": "數量",
                        "單價": "單位成本",
                        "金額": "總金額",
                    }
                    new_stg = raw_df.rename(columns=col_map).copy()

                    # 確保數值型態
                    new_stg["數量"] = pd.to_numeric(new_stg["數量"], errors="coerce").fillna(0).astype(int)
                    new_stg["單位成本"] = pd.to_numeric(new_stg["單位成本"], errors="coerce").fillna(0)
                    new_stg["總金額"] = pd.to_numeric(new_stg["總金額"], errors="coerce").fillna(
                        new_stg["數量"] * new_stg["單位成本"]
                    )

                    # 入庫日期
                    new_stg["入庫日期"] = pd.to_datetime(
                        new_stg["入庫日期"], errors="coerce"
                    ).dt.strftime("%Y-%m-%d")
                    new_stg["入庫日期"] = new_stg["入庫日期"].fillna(
                        pd.Timestamp.now().strftime("%Y-%m-%d")
                    )

                    # 保留系統需要的欄位順序
                    keep_cols = ["主貨號", "貨號", "商品名稱", "規格", "數量", "單位成本", "總金額", "入庫日期"]
                    new_stg = new_stg[keep_cols]

                    # 整份取代（使用者應在 Excel 中保留舊資料 + 新增資料後上傳）
                    save_storage(new_stg)

                    st.success(f"✅ 成功匯入 **{len(new_stg)}** 筆入庫資料，已儲存至 data/入庫.xlsx")
        except Exception as e:
            st.error(f"匯入失敗：{e}")

    # 手動新增
    st.markdown("---")
    st.subheader("網頁新增入庫")
    with st.form("add_stg", clear_on_submit=True):
        r1 = st.columns(4)
        main_sku = r1[0].text_input("主貨號")
        sku      = r1[1].text_input("貨號")
        name     = r1[2].text_input("商品名稱")
        spec     = r1[3].text_input("規格")
        r2 = st.columns(4)
        qty       = r2[0].number_input("數量", min_value=1, value=1)
        unit_cost = r2[1].number_input("單位成本", min_value=0.0, step=0.1)
        stg_date  = r2[2].date_input("入庫日期")
        if st.form_submit_button("➕ 新增"):
            if not sku:
                sku = f"{main_sku}-{spec}" if spec else main_sku
            total = qty * unit_cost
            row = pd.DataFrame([{
                "主貨號": main_sku, "貨號": sku, "商品名稱": name,
                "規格": spec, "數量": qty, "單位成本": unit_cost,
                "總金額": total, "入庫日期": str(stg_date),
            }])
            existing = load_storage()
            save_storage(pd.concat([existing, row], ignore_index=True))
            st.session_state["stg_success"] = True 
            st.rerun()
            
    if st.session_state.pop("stg_success", False):
        st.success("新增成功")            

# ══════════════════════════════════════════════════════════════
# Tab 2 – 匯入平台訂單
# ══════════════════════════════════════════════════════════════
with tab_order:
    platform = st.selectbox("選擇平台", ["蝦皮", "露天", "官網 (EasyStore)", "MOMO"], index=0)
    uploaded = st.file_uploader(
        "上傳訂單檔案",
        type=["xlsx", "xls", "csv"],
        key="order_upload",
    )

    # 平台名稱 → 檔案名稱對應
    _PLAT_FILE = {"蝦皮": "蝦皮", "露天": "露天", "官網": "官網", "MOMO": "MOMO"}

    if uploaded and st.button("🚀 開始匯入訂單", type="primary"):
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

                # 寫入平台專屬 xlsx（累積 + 去重）
                plat_key = new["平台"].iloc[0] if "平台" in new.columns else ""
                plat_file = _PLAT_FILE.get(plat_key, plat_key)
                if plat_file:
                    append_platform_orders(new, plat_file)

                # 自動更新對照表
                stg = load_storage()
                compare = load_compare_table()
                updated = auto_match_compare_table(new, stg, compare)
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
    st.subheader("各平台累積訂單")
    for plat_name, plat_file in [("\U0001f6d2 蝦皮", "蝦皮"), ("\U0001f3ea 露天", "露天"), ("\U0001f310 官網", "官網")]:
        pdf = load_platform_orders(plat_file)
        st.markdown(f"**{plat_name}**（{len(pdf)} 筆）")
        if not pdf.empty:
            st.dataframe(pdf, use_container_width=True, hide_index=True)
        else:
            st.info(f"尚未匯入{plat_file}訂單")