import streamlit as st
import pandas as pd
from pathlib import Path

from utils.data_manager import (
    DATA_DIR,
    load_storage, save_storage,
    load_compare_table, save_compare_table,
    load_platform_orders, append_platform_orders,
    clear_storage, clear_platform_orders,
)
from utils.parsers import parse_shopee, parse_ruten, parse_easystore, read_file_flexible
from utils.calculators import auto_match_compare_table


def _to_arrow_safe_display_df(df: pd.DataFrame) -> pd.DataFrame:
    """Convert mixed-type object columns to strings to avoid Arrow serialization errors."""
    if df.empty:
        return df

    safe_df = df.copy()
    for col in safe_df.columns:
        if pd.api.types.is_object_dtype(safe_df[col]):
            safe_df[col] = safe_df[col].map(lambda v: "" if pd.isna(v) else str(v))
    return safe_df

# 調整元件的樣式
st.markdown("""
    <style>
    [data-testid="stFileUploader"] {
        max-width: 150px;
    }
    [data-testid="stFileUploaderDropzoneInstructions"] {
        display: none;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <style>
    [data-testid="stSelectbox"] {
        max-width: 150px;
    }
    </style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# Tab 切換
# ══════════════════════════════════════════════════════════════
tab_storage, tab_order = st.tabs(["📥 匯入入庫資料", "📦 匯入平台訂單"])

# ══════════════════════════════════════════════════════════════
# Tab 1 – 匯入入庫資料
# ══════════════════════════════════════════════════════════════
TEMPLATE_PATH = DATA_DIR / "入庫.xlsx"

with tab_storage:

    if st.session_state.pop("stg_upload_success", None) is not None:
        st.success("✅ 成功匯入入庫資料")

    st.subheader("目前入庫資料 (後台使用)")
    storage = load_storage()
    if not storage.empty:
        st.dataframe(storage, width="stretch", hide_index=True)
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
            st.dataframe(raw_df, width="stretch", hide_index=True)

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

                    # 相同 主貨號/貨號/商品名稱/規格/單位成本/入庫日期 → 合併數量與總金額
                    group_keys = ["主貨號", "貨號", "商品名稱", "規格", "單位成本", "入庫日期"]
                    new_stg = (
                        new_stg.groupby(group_keys, dropna=False, sort=False)
                        .agg(數量=("數量", "sum"), 總金額=("總金額", "sum"))
                        .reset_index()[keep_cols]
                    )

                    # 整份取代（使用者應在 Excel 中保留舊資料 + 新增資料後上傳）
                    save_storage(new_stg)
                    st.session_state["stg_upload_success"] = len(new_stg)
                    st.rerun()
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
            combined = pd.concat([existing, row], ignore_index=True)
            # 相同 主貨號/貨號/商品名稱/規格/單位成本/入庫日期 → 合併數量與總金額
            _keep = ["主貨號", "貨號", "商品名稱", "規格", "數量", "單位成本", "總金額", "入庫日期"]
            _grp  = ["主貨號", "貨號", "商品名稱", "規格", "單位成本", "入庫日期"]
            combined = (
                combined.groupby(_grp, dropna=False, sort=False)
                .agg(數量=("數量", "sum"), 總金額=("總金額", "sum"))
                .reset_index()[_keep]
            )
            save_storage(combined)
            st.session_state["stg_success"] = True 
            st.rerun()
            
    if st.session_state.pop("stg_success", False):
        st.success("新增成功")

    # 手動刪除
    st.markdown("---")
    st.subheader("網頁刪除入庫")
    with st.form("del_stg", clear_on_submit=True):
        d1 = st.columns(2)
        del_sku  = d1[0].text_input("貨號")
        del_date = d1[1].date_input("入庫日期")
        if st.form_submit_button("🗑️ 刪除"):
            if not del_sku:
                st.warning("請輸入貨號")
            else:
                existing = load_storage()
                del_date_str = str(del_date)
                mask = ~(
                    (existing["貨號"].astype(str) == del_sku) &
                    (existing["入庫日期"].astype(str) == del_date_str)
                )
                updated = existing[mask].reset_index(drop=True)
                if len(updated) == len(existing):
                    st.session_state["stg_del_notfound"] = True
                else:
                    save_storage(updated)
                    st.session_state["stg_del_success"] = True
                st.rerun()

    if st.session_state.pop("stg_del_success", False):
        st.success("刪除成功")
    if st.session_state.pop("stg_del_notfound", False):
        st.warning("找不到符合的資料，請確認貨號與入庫日期是否正確")

    # ── 清0 入庫資料 ──────────────────────────────────────────
    st.markdown("---")
    if st.button("🗑️ 清除入庫資料", key="clear_stg_btn"):
        st.session_state["confirm_clear_stg"] = True
    if st.session_state.get("confirm_clear_stg"):
        st.warning("⚠️ 確定要清除所有入庫資料嗎？此操作無法復原！")
        _c1, _c2 = st.columns(2)
        if _c1.button("✅ 確認清除", key="confirm_clear_stg_yes", type="primary"):
            with st.spinner("清除中…"):
                clear_storage()
            st.session_state.pop("confirm_clear_stg", None)
            st.success("✅ 入庫資料已清除")
            st.rerun()
        if _c2.button("❌ 取消", key="confirm_clear_stg_no"):
            st.session_state.pop("confirm_clear_stg", None)
            st.rerun()
            
# ══════════════════════════════════════════════════════════════
# Tab 2 – 匯入平台訂單
# ══════════════════════════════════════════════════════════════
with tab_order:
    if st.session_state.pop("order_upload_success", None) is not None:
        st.success("✅ 訂單匯入成功")

    platform = st.selectbox("選擇平台", ["蝦皮", "露天", "官網 (EasyStore)", "MO店"], index=0)
    uploaded = st.file_uploader(
        "上傳訂單檔案",
        type=["xlsx", "xls", "csv"],
        key="order_upload",
    )

    # 平台名稱 → 檔案名稱對應
    _PLAT_FILE = {"蝦皮": "蝦皮", "露天": "露天", "官網": "官網", "MO店": "MO店"}

    # 各平台特徵欄位數（用來驗證上傳檔案是否與所選平台一致）
    _PLAT_COL_COUNT = {
        "蝦皮": 57,
        "露天": 26,
        "官網": 77,
    }

    def _check_platform_columns(col_count, selected_platform):
        """回傳 (通過, 錯誤訊息)。檢查上傳檔案欄位數是否符合所選平台。"""
        expected = _PLAT_COL_COUNT.get(selected_platform)
        if expected and col_count != expected:
            # 嘗試辨識實際平台
            for pname, pcnt in _PLAT_COL_COUNT.items():
                if col_count == pcnt:
                    return False, f"檔案欄位數不符！您選擇了 **{selected_platform}**（應為 {expected} 欄），但檔案有 **{col_count}** 欄，看起來是 **{pname}** 的格式。"
            return False, f"檔案欄位數不符！您選擇了 **{selected_platform}**（應為 {expected} 欄），但檔案有 **{col_count}** 欄，請確認上傳正確的檔案。"
        return True, ""

    if uploaded and st.button("🚀 開始匯入訂單", type="primary"):
        try:
            # 先讀取原始檔案做欄位檢查
            raw_preview = read_file_flexible(uploaded)
            uploaded.seek(0)  # 重置指標供後續 parser 使用

            plat_short = "蝦皮" if "蝦皮" in platform else ("露天" if "露天" in platform else "官網")
            col_ok, col_msg = _check_platform_columns(len(raw_preview.columns), plat_short)
            if not col_ok:
                st.error(col_msg)
                st.stop()

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
                # 寫入平台專屬 xlsx（累積原始資料 + 去重）
                plat_file = _PLAT_FILE.get(plat_short, plat_short)
                if plat_file:
                    append_platform_orders(raw_preview, plat_file)

                # 自動更新對照表
                stg = load_storage()
                compare = load_compare_table()
                updated = auto_match_compare_table(new, stg, compare)
                save_compare_table(updated)
                st.session_state["order_upload_success"] = len(new)
                st.rerun()

        except Exception as e:
            st.error(f"匯入失敗：{e}")

    st.markdown("---")
    st.subheader("各平台累積訂單")
    for plat_name, plat_file in [("\U0001f6d2 蝦皮", "蝦皮"), ("\U0001f3ea 露天", "露天"), ("\U0001f310 官網", "官網")]:
        pdf = load_platform_orders(plat_file)
        st.markdown(f"**{plat_name}**（{len(pdf)} 筆）")
        if not pdf.empty:
            st.dataframe(_to_arrow_safe_display_df(pdf), width="stretch", hide_index=True)
        else:
            st.info(f"尚未匯入{plat_file}訂單")

    # ── 清0 平台訂單 ───────────────────────────────────────────
    st.markdown("---")
    if st.button("🗑️ 清除所有平台訂單", key="clear_orders_btn"):
        st.session_state["confirm_clear_orders"] = True
    if st.session_state.get("confirm_clear_orders"):
        st.warning("⚠️ 確定要清除所有平台訂單（蝦皮、露天、官網）嗎？此操作無法復原！")
        _c1, _c2 = st.columns(2)
        if _c1.button("✅ 確認清除", key="confirm_clear_orders_yes", type="primary"):
            with st.spinner("清除中…"):
                for _plat in ["蝦皮", "露天", "官網"]:
                    clear_platform_orders(_plat)
            st.session_state.pop("confirm_clear_orders", None)
            st.success("✅ 所有平台訂單已清除")
            st.rerun()
        if _c2.button("❌ 取消", key="confirm_clear_orders_no"):
            st.session_state.pop("confirm_clear_orders", None)
            st.rerun()