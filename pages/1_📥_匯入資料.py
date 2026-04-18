import streamlit as st
import pandas as pd
from datetime import datetime, timezone, timedelta

from utils.data_manager import (
    load_storage, save_storage,
    load_platform_orders, append_platform_orders,
    clear_storage, clear_platform_orders,
    load_combo_sku, save_combo_sku, clear_combo_sku,
    load_custom_orders, save_custom_orders, clear_custom_orders,
    _CUSTOM_ORDER_COLS,
)
from utils.parsers import parse_shopee, parse_ruten, parse_easystore, read_file_flexible
from utils.styles import apply_global_styles

TZ_TAIPEI = timezone(timedelta(hours=8))
apply_global_styles()


@st.cache_data
def _to_arrow_safe_display_df(df: pd.DataFrame) -> pd.DataFrame:
    """Convert mixed-type object columns to strings to avoid Arrow serialization errors."""
    if df.empty:
        return df

    safe_df = df.copy()
    for col in safe_df.columns:
        if pd.api.types.is_object_dtype(safe_df[col]):
            safe_df[col] = safe_df[col].map(lambda v: "" if pd.isna(v) else str(v))
    return safe_df


@st.cache_data
def _df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    """Cache CSV serialization so it only runs when data changes."""
    return _to_arrow_safe_display_df(df).to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")

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
# Tab 1 – 匯入入庫資料
# ══════════════════════════════════════════════════════════════
@st.fragment
def _render_storage_tab():

    _stg_result = st.session_state.pop("stg_upload_success", None)
    if _stg_result is not None:
        if _stg_result > 0:
            st.success(f"✅ 成功匯入入庫資料，新增 {_stg_result} 筆")
        else:
            st.info("✅ 匯入完成，本次檔案內的資料皆已存在（無新增）")
    _stg_upload_ts = st.session_state.get("stg_upload_saved_at")
    if _stg_upload_ts:
        st.caption(f"🕐 最後匯入：{_stg_upload_ts}")

    st.subheader("目前入庫資料 (後台使用)")
    storage = load_storage()
    if not storage.empty:
        _stg_page_size = 500
        _stg_total = len(storage)
        _stg_total_pages = max(1, (_stg_total - 1) // _stg_page_size + 1)
        _stg_dl_col, _stg_pg_col = st.columns([1, 3])
        _stg_dl_col.download_button("⬇️ 下載全部入庫資料", data=_df_to_csv_bytes(storage), file_name="入庫資料.csv", mime="text/csv", key="dl_stg_list")
        if _stg_total_pages > 1:
            _stg_page = _stg_pg_col.selectbox(
                "頁碼", list(range(1, _stg_total_pages + 1)),
                format_func=lambda x: f"{x}/{_stg_total_pages} 頁",
                key="stg_list_page",
            )
        else:
            _stg_page = 1
        _stg_start = (_stg_page - 1) * _stg_page_size
        st.dataframe(storage.iloc[_stg_start:_stg_start + _stg_page_size], width='stretch', hide_index=True)
        st.caption(f"第 {_stg_page} 頁 / 共 {_stg_total_pages} 頁（{_stg_total:,} 筆）")
    else:
        st.info("尚未有入庫資料")
            
    st.subheader("手動 EXCEL 新增入庫")

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

            # ── 防呆：同貨號但品名／規格不同 ──────────────────────────
            _dup_rows = []
            if "貨號" in raw_df.columns and "名稱" in raw_df.columns and "規格" in raw_df.columns:
                # 1. 檔案內部：同貨號 + 不同 名稱/規格
                _file_map = raw_df[["貨號", "名稱", "規格"]].astype(str).drop_duplicates()
                _dup_in_file = _file_map[_file_map.duplicated(subset=["貨號"], keep=False)].copy()
                if not _dup_in_file.empty:
                    _dup_in_file.insert(0, "來源", "檔案內部重複")
                    _dup_rows.append(_dup_in_file.rename(columns={"名稱": "商品名稱"}))

                # 2. 與現有入庫資料衝突：同貨號 + 不同 商品名稱/規格
                _existing_stg = load_storage()
                if not _existing_stg.empty:
                    _upload_map = (
                        raw_df[["貨號", "名稱", "規格"]].astype(str).drop_duplicates()
                        .rename(columns={"名稱": "商品名稱"})
                    )
                    _exist_map = _existing_stg[["貨號", "商品名稱", "規格"]].astype(str).drop_duplicates()
                    _merged = _upload_map.merge(_exist_map, on="貨號", suffixes=("_新", "_舊"))
                    _conflict = _merged[
                        (_merged["商品名稱_新"] != _merged["商品名稱_舊"]) |
                        (_merged["規格_新"] != _merged["規格_舊"])
                    ].copy()
                    if not _conflict.empty:
                        _conflict.insert(0, "來源", "與現有資料衝突")
                        _dup_rows.append(_conflict)

            _has_dup_error = bool(_dup_rows)
            if _has_dup_error:
                _all_dups = pd.concat(_dup_rows, ignore_index=True)
                st.error("⚠️ 發現貨號衝突！以下貨號相同但品名／規格不同，請修正後再匯入。")
                st.dataframe(_all_dups, width="stretch", hide_index=True)

            if not _has_dup_error and st.button("🚀 確認匯入入庫資料", type="primary"):
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
                        pd.Timestamp.now(tz="Asia/Taipei").strftime("%Y-%m-%d")
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

                    # 保留歷史：PK=(貨號+入庫日期) 已存在則保留舊，否則新增
                    existing_stg = load_storage()
                    if not existing_stg.empty:
                        existing_keys = set(
                            existing_stg["貨號"].astype(str).str.cat(existing_stg["入庫日期"].astype(str), sep="||")
                        )
                        new_keys = new_stg["貨號"].astype(str).str.cat(new_stg["入庫日期"].astype(str), sep="||")
                        truly_new = new_stg[~new_keys.isin(existing_keys)]
                        merged_stg = pd.concat([existing_stg, truly_new], ignore_index=True)
                        added = len(truly_new)
                    else:
                        merged_stg = new_stg
                        added = len(new_stg)
                    save_storage(merged_stg)
                    st.session_state["stg_upload_success"] = added
                    st.session_state["stg_upload_saved_at"] = datetime.now(tz=TZ_TAIPEI).strftime("%Y-%m-%d %H:%M:%S")
                    st.rerun(scope="fragment")
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
            existing = load_storage()
            # ── 防呆：同貨號但品名／規格不同 ──────────────────────────
            _match = existing[existing["貨號"].astype(str) == str(sku)]
            _form_conflict = _match[
                (_match["商品名稱"].astype(str) != str(name)) |
                (_match["規格"].astype(str) != str(spec))
            ]
            if not _form_conflict.empty:
                _conflict_records = (
                    _form_conflict[["貨號", "商品名稱", "規格"]]
                    .drop_duplicates()
                    .to_dict("records")
                )
                st.session_state["stg_dup_conflict"] = _conflict_records
                st.rerun(scope="fragment")
            else:
                total = qty * unit_cost
                row = pd.DataFrame([{
                    "主貨號": main_sku, "貨號": sku, "商品名稱": name,
                    "規格": spec, "數量": qty, "單位成本": unit_cost,
                    "總金額": total, "入庫日期": str(stg_date),
                }])
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
                st.session_state["stg_add_saved_at"] = datetime.now(tz=TZ_TAIPEI).strftime("%Y-%m-%d %H:%M:%S")
                st.rerun(scope="fragment")

    if "stg_dup_conflict" in st.session_state:
        _conflict_data = st.session_state.pop("stg_dup_conflict")
        st.error("⚠️ 貨號衝突！此貨號已存在但品名／規格不同，請確認後再新增。")
        st.dataframe(pd.DataFrame(_conflict_data), width="stretch", hide_index=True)

    if st.session_state.pop("stg_success", False):
        st.success("新增成功")
    _stg_add_ts = st.session_state.get("stg_add_saved_at")
    if _stg_add_ts:
        st.caption(f"🕐 最後新增：{_stg_add_ts}")

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
                    st.session_state["stg_del_saved_at"] = datetime.now(tz=TZ_TAIPEI).strftime("%Y-%m-%d %H:%M:%S")
                st.rerun(scope="fragment")

    if st.session_state.pop("stg_del_success", False):
        st.success("刪除成功")
    _stg_del_ts = st.session_state.get("stg_del_saved_at")
    if _stg_del_ts:
        st.caption(f"🕐 最後刪除：{_stg_del_ts}")
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
            st.rerun(scope="fragment")
        if _c2.button("❌ 取消", key="confirm_clear_stg_no"):
            st.session_state.pop("confirm_clear_stg", None)
            st.rerun(scope="fragment")


# ══════════════════════════════════════════════════════════════
# Tab 2 – 匯入平台訂單
# ══════════════════════════════════════════════════════════════
@st.fragment
def _render_order_tab():
    if st.session_state.pop("order_upload_success", None) is not None:
        st.success("✅ 訂單匯入成功")
    _order_ts = st.session_state.get("order_saved_at")
    if _order_ts:
        st.caption(f"🕐 最後匯入：{_order_ts}")

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

                st.session_state["order_upload_success"] = len(new)
                st.session_state["order_saved_at"] = datetime.now(tz=TZ_TAIPEI).strftime("%Y-%m-%d %H:%M:%S")
                st.rerun(scope="fragment")

        except Exception as e:
            st.error(f"匯入失敗：{e}")

    st.markdown("---")
    st.subheader("各平台累積訂單")
    for plat_name, plat_file in [("\U0001f6d2 蝦皮", "蝦皮"), ("\U0001f3ea 露天", "露天"), ("\U0001f310 官網", "官網")]:
        pdf = load_platform_orders(plat_file)
        with st.expander(f"**{plat_name}**（{len(pdf)} 筆）", expanded=False):
            if not pdf.empty:
                _p_page_size = 500
                _p_total = len(pdf)
                _p_total_pages = max(1, (_p_total - 1) // _p_page_size + 1)
                _p_dl_col, _p_pg_col = st.columns([1, 3])
                _p_dl_col.download_button(f"⬇️ 下載{plat_file}全部訂單", data=_df_to_csv_bytes(pdf), file_name=f"{plat_file}訂單.csv", mime="text/csv", key=f"dl_ord_{plat_file}")
                if _p_total_pages > 1:
                    _p_page = _p_pg_col.selectbox(
                        "頁碼", list(range(1, _p_total_pages + 1)),
                        format_func=lambda x: f"{x}/{_p_total_pages} 頁",
                        key=f"ord_page_{plat_file}",
                    )
                else:
                    _p_page = 1
                _p_start = (_p_page - 1) * _p_page_size
                _pdf_slice = _to_arrow_safe_display_df(pdf).iloc[_p_start:_p_start + _p_page_size]
                st.dataframe(_pdf_slice, width='stretch', hide_index=True)
                st.caption(f"第 {_p_page} 頁 / 共 {_p_total_pages} 頁（{_p_total:,} 筆）")
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
            st.rerun(scope="fragment")
        if _c2.button("❌ 取消", key="confirm_clear_orders_no"):
            st.session_state.pop("confirm_clear_orders", None)
            st.rerun(scope="fragment")


# ══════════════════════════════════════════════════════════════
# Tab 3 – 匯入自建訂單
# ══════════════════════════════════════════════════════════════
@st.fragment
def _render_custom_tab():
    _cust_result = st.session_state.pop("custom_upload_success", None)
    if _cust_result is not None:
        if _cust_result > 0:
            st.success(f"✅ 成功匯入自建訂單，新增 {_cust_result} 筆")
        else:
            st.info("✅ 匯入完成，本次檔案內的資料皆已存在（無新增）")
    _cust_upload_ts = st.session_state.get("custom_upload_saved_at")
    if _cust_upload_ts:
        st.caption(f"🕐 最後匯入：{_cust_upload_ts}")

    st.subheader("目前自建訂單")
    custom_orders = load_custom_orders()
    if not custom_orders.empty:
        _cust_page_size = 500
        _cust_total = len(custom_orders)
        _cust_total_pages = max(1, (_cust_total - 1) // _cust_page_size + 1)
        _cust_dl_col, _cust_pg_col = st.columns([1, 3])
        _cust_xlsx_buf = __import__("io").BytesIO()
        custom_orders.to_excel(_cust_xlsx_buf, index=False, engine="openpyxl")
        _cust_dl_col.download_button("⬇️ 下載自建訂單", data=_cust_xlsx_buf.getvalue(), file_name="自建訂單.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_custom_list")
        if _cust_total_pages > 1:
            _cust_page = _cust_pg_col.selectbox(
                "頁碼", list(range(1, _cust_total_pages + 1)),
                format_func=lambda x: f"{x}/{_cust_total_pages} 頁",
                key="custom_list_page",
            )
        else:
            _cust_page = 1
        _cust_start = (_cust_page - 1) * _cust_page_size
        st.dataframe(
            _to_arrow_safe_display_df(custom_orders).iloc[_cust_start:_cust_start + _cust_page_size],
            width='stretch', hide_index=True,
        )
        st.caption("欄位說明：")
        st.caption("小計＝數量×單價")
        st.caption("費用小記＝折扣優惠＋實際運費－買家支付運費＋未取貨/退貨運費＋其他費用")
        st.caption("訂單總金額＝小計－費用小記")
        st.caption(f"第 {_cust_page} 頁 / 共 {_cust_total_pages} 頁（{_cust_total:,} 筆）")
    else:
        st.info("尚未有自建訂單")

    # ── Excel 匯入 ──
    st.markdown("---")
    st.subheader("Excel 匯入自建訂單")
    cust_file = st.file_uploader(
        "⬆ 上傳自建訂單",
        type=["xlsx", "xls", "csv"],
        key="custom_upload",
    )

    if cust_file:
        try:
            raw_cust = read_file_flexible(cust_file)
            st.markdown("#### 預覽上傳的資料")
            st.dataframe(raw_cust, width="stretch", hide_index=True)

            required_cols = ["日期", "平台名稱", "訂單編號", "訂單狀態", "貨號", "數量", "單價"]
            missing = [c for c in required_cols if c not in raw_cust.columns]
            if missing:
                st.error(f"缺少必要欄位：{', '.join(missing)}")
            elif st.button("🚀 確認匯入自建訂單", type="primary"):
                new_cust = raw_cust.copy()
                # 確保數值欄位型態
                new_cust["數量"] = pd.to_numeric(new_cust["數量"], errors="coerce").fillna(0).astype(int)
                new_cust["單價"] = pd.to_numeric(new_cust["單價"], errors="coerce").fillna(0)
                new_cust["小計"] = new_cust["數量"] * new_cust["單價"]
                for _fc in ["折扣優惠", "買家支付運費", "實際運費", "未取貨/退貨運費", "其他費用"]:
                    if _fc not in new_cust.columns:
                        new_cust[_fc] = 0
                    new_cust[_fc] = pd.to_numeric(new_cust[_fc], errors="coerce").fillna(0)
                new_cust["費用小記"] = (
                    new_cust["折扣優惠"] + new_cust["實際運費"] - new_cust["買家支付運費"]
                    + new_cust["未取貨/退貨運費"] + new_cust["其他費用"]
                )
                new_cust["訂單總金額"] = new_cust["小計"] - new_cust["費用小記"]
                new_cust["日期"] = new_cust["日期"].replace("NaT", pd.Timestamp.now(tz="Asia/Taipei").strftime("%Y-%m-%d"))
                for _sc in ["平台名稱", "訂單編號", "訂單狀態", "貨號"]:
                    if _sc in new_cust.columns:
                        new_cust[_sc] = new_cust[_sc].astype(str).str.strip()
                # 補齊可能缺少的文字欄位
                for _tc in ["買家姓名", "買家帳號"]:
                    if _tc not in new_cust.columns:
                        new_cust[_tc] = ""
                # 只保留標準欄位
                new_cust = new_cust[[c for c in _CUSTOM_ORDER_COLS if c in new_cust.columns]]

                # 去重合併
                existing_cust = load_custom_orders()
                if not existing_cust.empty:
                    existing_keys = set(
                        existing_cust["訂單編號"].astype(str) + "||"
                        + existing_cust["貨號"].astype(str) + "||"
                        + existing_cust["日期"].astype(str)
                    )
                    new_keys = (
                        new_cust["訂單編號"].astype(str) + "||"
                        + new_cust["貨號"].astype(str) + "||"
                        + new_cust["日期"].astype(str)
                    )
                    truly_new = new_cust[~new_keys.isin(existing_keys)]
                    merged = pd.concat([existing_cust, truly_new], ignore_index=True)
                    added = len(truly_new)
                else:
                    merged = new_cust
                    added = len(new_cust)
                save_custom_orders(merged)
                st.session_state["custom_upload_success"] = added
                st.session_state["custom_upload_saved_at"] = datetime.now(tz=TZ_TAIPEI).strftime("%Y-%m-%d %H:%M:%S")
                st.rerun(scope="fragment")
        except Exception as e:
            st.error(f"匯入失敗：{e}")

    # ── 網頁新增自建訂單 ──
    st.markdown("---")
    st.subheader("網頁新增自建訂單")

    _STATUS_OPTIONS = ["已完成", "退貨", "損失賠償", "未取貨"]

    default_row = pd.DataFrame([{
        "日期": pd.Timestamp.now().normalize(),
        "平台名稱": "其他",
        "訂單編號": "",
        "訂單狀態": "已完成",
        "買家姓名": "",
        "買家帳號": "",
        "貨號": "",
        "數量": 1,
        "單價": 0,
        "折扣優惠": 0,
        "買家支付運費": 0,
        "實際運費": 0,
        "未取貨/退貨運費": 0,
        "其他費用": 0,
    }])

    _cust_editor_cfg = {
        "日期": st.column_config.DateColumn("日期", format="YYYY-MM-DD", required=True),
        "平台名稱": st.column_config.TextColumn("平台名稱", required=True),
        "訂單編號": st.column_config.TextColumn("訂單編號", required=True),
        "訂單狀態": st.column_config.SelectboxColumn("訂單狀態", options=_STATUS_OPTIONS, required=True),
        "買家姓名": st.column_config.TextColumn("買家姓名"),
        "買家帳號": st.column_config.TextColumn("買家帳號"),
        "貨號": st.column_config.TextColumn("貨號", required=True),
        "數量": st.column_config.NumberColumn("數量", min_value=1, step=1, required=True),
        "單價": st.column_config.NumberColumn("單價", min_value=0, step=1, required=True),
        "折扣優惠": st.column_config.NumberColumn("折扣優惠", min_value=0, step=1),
        "買家支付運費": st.column_config.NumberColumn("買家支付運費", min_value=0, step=1),
        "實際運費": st.column_config.NumberColumn("實際運費", min_value=0, step=1),
        "未取貨/退貨運費": st.column_config.NumberColumn("未取貨/退貨運費", min_value=0, step=1),
        "其他費用": st.column_config.NumberColumn("其他費用", min_value=0, step=1),
    }

    edited_custom = st.data_editor(
        default_row,
        num_rows="dynamic",
        width='stretch',
        hide_index=True,
        key="custom_order_editor",
        column_config=_cust_editor_cfg,
    )

    if st.button("💾 儲存自建訂單", type="primary", key="save_custom_order"):
        valid = edited_custom[
            edited_custom["訂單編號"].fillna("").astype(str).str.strip() != ""
        ].copy()
        if valid.empty:
            st.error("請至少填寫一筆訂單（訂單編號不可為空）")
        else:
            valid["數量"] = pd.to_numeric(valid["數量"], errors="coerce").fillna(0).astype(int)
            valid["單價"] = pd.to_numeric(valid["單價"], errors="coerce").fillna(0)
            valid["小計"] = valid["數量"] * valid["單價"]
            for _fc in ["折扣優惠", "買家支付運費", "實際運費", "未取貨/退貨運費", "其他費用"]:
                valid[_fc] = pd.to_numeric(valid[_fc], errors="coerce").fillna(0)
            valid["費用小記"] = (
                valid["折扣優惠"] + valid["實際運費"] - valid["買家支付運費"]
                + valid["未取貨/退貨運費"] + valid["其他費用"]
            )
            valid["訂單總金額"] = valid["小計"] - valid["費用小記"]
            valid["日期"] = pd.to_datetime(valid["日期"], errors="coerce").dt.strftime("%Y-%m-%d")
            valid["日期"] = valid["日期"].fillna(pd.Timestamp.now(tz="Asia/Taipei").strftime("%Y-%m-%d"))
            for _sc in ["平台名稱", "訂單編號", "訂單狀態", "貨號", "買家姓名", "買家帳號"]:
                if _sc in valid.columns:
                    valid[_sc] = valid[_sc].fillna("").astype(str).str.strip()
            valid = valid[[c for c in _CUSTOM_ORDER_COLS if c in valid.columns]]

            existing_cust = load_custom_orders()
            merged = pd.concat([existing_cust, valid], ignore_index=True) if not existing_cust.empty else valid
            save_custom_orders(merged)
            st.session_state["custom_add_success"] = True
            st.session_state["custom_add_saved_at"] = datetime.now(tz=TZ_TAIPEI).strftime("%Y-%m-%d %H:%M:%S")
            st.rerun(scope="fragment")

    if st.session_state.pop("custom_add_success", False):
        st.success("✅ 自建訂單新增成功")
    _cust_add_ts = st.session_state.get("custom_add_saved_at")
    if _cust_add_ts:
        st.caption(f"🕐 最後新增：{_cust_add_ts}")

    # ── 刪除自建訂單 ──
    st.markdown("---")
    st.subheader("刪除自建訂單")
    with st.form("del_custom", clear_on_submit=True):
        d1 = st.columns(2)
        del_cust_oid = d1[0].text_input("訂單編號")
        del_cust_date = d1[1].date_input("日期")
        if st.form_submit_button("🗑️ 刪除"):
            if not del_cust_oid:
                st.warning("請輸入訂單編號")
            else:
                existing_cust = load_custom_orders()
                del_date_str = str(del_cust_date)
                mask = ~(
                    (existing_cust["訂單編號"].astype(str) == del_cust_oid) &
                    (existing_cust["日期"].astype(str) == del_date_str)
                )
                updated = existing_cust[mask].reset_index(drop=True)
                if len(updated) == len(existing_cust):
                    st.session_state["custom_del_notfound"] = True
                else:
                    save_custom_orders(updated)
                    st.session_state["custom_del_success"] = True
                    st.session_state["custom_del_saved_at"] = datetime.now(tz=TZ_TAIPEI).strftime("%Y-%m-%d %H:%M:%S")
                st.rerun(scope="fragment")

    if st.session_state.pop("custom_del_success", False):
        st.success("刪除成功")
    _cust_del_ts = st.session_state.get("custom_del_saved_at")
    if _cust_del_ts:
        st.caption(f"🕐 最後刪除：{_cust_del_ts}")
    if st.session_state.pop("custom_del_notfound", False):
        st.warning("找不到符合的資料，請確認訂單編號與日期是否正確")

    # ── 清除所有自建訂單 ──
    st.markdown("---")
    if st.button("🗑️ 清除所有自建訂單", key="clear_custom_btn"):
        st.session_state["confirm_clear_custom"] = True
    if st.session_state.get("confirm_clear_custom"):
        st.warning("⚠️ 確定要清除所有自建訂單嗎？此操作無法復原！")
        _c1, _c2 = st.columns(2)
        if _c1.button("✅ 確認清除", key="confirm_clear_custom_yes", type="primary"):
            with st.spinner("清除中…"):
                clear_custom_orders()
            st.session_state.pop("confirm_clear_custom", None)
            st.success("✅ 自建訂單已清除")
            st.rerun(scope="fragment")
        if _c2.button("❌ 取消", key="confirm_clear_custom_no"):
            st.session_state.pop("confirm_clear_custom", None)
            st.rerun(scope="fragment")


# ══════════════════════════════════════════════════════════════
# Tab 4 – 組合貨號
# ══════════════════════════════════════════════════════════════
@st.fragment
def _render_combo_tab():
    st.subheader("🔗 組合貨號管理")
    st.markdown(
        "組合商品由多組原料貨號組成，例如：`YGK03002-RDL30-RR-DIY` = `YGJ03002-RDL × 30` + `YGK03002-RR × 1`\n\n"
        "出庫時系統會自動扣除各原料貨號的庫存數量。"
    )

    # ── 顯示現有組合貨號 ──
    combo_df = load_combo_sku()
    if not combo_df.empty:
        st.markdown("#### 目前組合貨號")
        st.dataframe(combo_df, width="stretch", hide_index=True)
    else:
        st.info("尚未建立組合貨號")

    if st.session_state.pop("combo_add_success", False):
        st.success("✅ 組合貨號新增成功")
    if st.session_state.pop("combo_del_success", False):
        st.success("✅ 組合貨號刪除成功")
    if st.session_state.pop("combo_del_notfound", False):
        st.warning("找不到該組合貨號")

    # ── Excel 匯入組合貨號 ──
    st.markdown("---")
    st.subheader("Excel 匯入組合貨號")

    _combo_upload_result = st.session_state.pop("combo_upload_success", None)
    if _combo_upload_result is not None:
        if _combo_upload_result > 0:
            st.success(f"✅ 成功匯入組合貨號，新增 {_combo_upload_result} 組")
        else:
            st.info("✅ 匯入完成，本次檔案內的組合貨號皆已存在（無新增）")

    combo_file = st.file_uploader(
        "⬆ 上傳組合貨號",
        type=["xlsx", "xls", "csv"],
        key="combo_upload",
    )

    if combo_file:
        try:
            raw_combo = read_file_flexible(combo_file)
            st.markdown("#### 預覽上傳的資料")
            st.dataframe(raw_combo, width="stretch", hide_index=True)

            required_cols = ["組合貨號", "原料貨號", "原料數量"]
            missing = [c for c in required_cols if c not in raw_combo.columns]
            if missing:
                st.error(f"缺少必要欄位：{', '.join(missing)}")
            elif st.button("🚀 確認匯入組合貨號", type="primary"):
                new_combo = raw_combo[required_cols].copy()
                new_combo["組合貨號"] = new_combo["組合貨號"].astype(str).str.strip()
                new_combo["原料貨號"] = new_combo["原料貨號"].astype(str).str.strip()
                new_combo["原料數量"] = pd.to_numeric(new_combo["原料數量"], errors="coerce").fillna(1).astype(int)
                new_combo = new_combo[new_combo["組合貨號"] != ""]

                existing = load_combo_sku()
                if not existing.empty:
                    # 以組合貨號為單位：新檔案中出現的組合貨號覆蓋舊的
                    new_codes = set(new_combo["組合貨號"].unique())
                    old_codes = set(existing["組合貨號"].unique())
                    added_count = len(new_codes - old_codes)
                    existing = existing[~existing["組合貨號"].isin(new_codes)]
                    combined = pd.concat([existing, new_combo], ignore_index=True)
                else:
                    added_count = len(new_combo["組合貨號"].unique())
                    combined = new_combo

                save_combo_sku(combined)
                st.session_state["combo_upload_success"] = added_count
                st.rerun(scope="fragment")
        except Exception as e:
            st.error(f"匯入失敗：{e}")

    # ── 新增組合貨號 ──
    st.markdown("---")
    st.subheader("新增組合貨號")

    # 使用 st.form 包裝，配合 data_editor 動態表格
    with st.form("new_combo_form", clear_on_submit=True):
        combo_code_input = st.text_input("組合貨號")
        
        st.markdown("**原料列表（可直接在表格下方點擊 ➕ 新增列，或選取左側勾選框刪除）：**")
        
        # 預設給定一列空白讓使用者填寫
        default_materials = pd.DataFrame([{"原料貨號": "", "原料數量": 1}])
        
        # 使用 data_editor 達到客戶端動態新增/刪除，不會觸發 rerun
        edited_materials = st.data_editor(
            default_materials,
            num_rows="dynamic", # 允許動態新增/刪除資料列
            width='stretch',
            hide_index=True,
            column_config={
                "原料貨號": st.column_config.TextColumn("原料貨號", required=True),
                "原料數量": st.column_config.NumberColumn("原料數量", min_value=1, step=1, required=True),
            }
        )
        
        submitted = st.form_submit_button("💾 儲存組合貨號", type="primary")

    if submitted:
        if not combo_code_input.strip():
            st.error("請輸入組合貨號")
        else:
            # 清理資料：過濾掉未填寫的空行
            valid_materials = edited_materials[edited_materials["原料貨號"].fillna("").str.strip() != ""]
            
            if valid_materials.empty:
                st.error("請至少輸入一筆原料貨號")
            else:
                # 建立要新增的 DataFrame
                new_rows = pd.DataFrame({
                    "組合貨號": combo_code_input.strip(),
                    "原料貨號": valid_materials["原料貨號"].str.strip(),
                    "原料數量": valid_materials["原料數量"].astype(int)
                })
                
                existing = load_combo_sku()
                if not existing.empty:
                    # 若已存在相同組合貨號，先移除舊的，以新的覆蓋
                    existing = existing[existing["組合貨號"] != combo_code_input.strip()]
                    
                combined = pd.concat([existing, new_rows], ignore_index=True)
                save_combo_sku(combined)
                
                st.session_state["combo_add_success"] = True
                st.rerun(scope="fragment")

    # ── 刪除組合貨號 ──
    st.markdown("---")
    st.subheader("刪除組合貨號")
    with st.form("del_combo", clear_on_submit=True):
        del_combo_code = st.text_input("要刪除的組合貨號")
        if st.form_submit_button("🗑️ 刪除"):
            if not del_combo_code:
                st.warning("請輸入組合貨號")
            else:
                existing = load_combo_sku()
                if existing.empty or del_combo_code not in existing["組合貨號"].values:
                    st.session_state["combo_del_notfound"] = True
                else:
                    updated = existing[existing["組合貨號"] != del_combo_code].reset_index(drop=True)
                    save_combo_sku(updated)
                    st.session_state["combo_del_success"] = True
                st.rerun(scope="fragment")

    # ── 清除所有組合貨號 ──
    st.markdown("---")
    if st.button("🗑️ 清除所有組合貨號", key="clear_combo_btn"):
        st.session_state["confirm_clear_combo"] = True
    if st.session_state.get("confirm_clear_combo"):
        st.warning("⚠️ 確定要清除所有組合貨號資料嗎？此操作無法復原！")
        _c1, _c2 = st.columns(2)
        if _c1.button("✅ 確認清除", key="confirm_clear_combo_yes", type="primary"):
            with st.spinner("清除中…"):
                clear_combo_sku()
            st.session_state.pop("confirm_clear_combo", None)
            st.success("✅ 組合貨號已清除")
            st.rerun(scope="fragment")
        if _c2.button("❌ 取消", key="confirm_clear_combo_no"):
            st.session_state.pop("confirm_clear_combo", None)
            st.rerun(scope="fragment")


# ══════════════════════════════════════════════════════════════
# Tab 切換
# ══════════════════════════════════════════════════════════════
tab_storage, tab_order, tab_custom, tab_combo = st.tabs(["📥 匯入入庫資料", "📦 匯入平台訂單", "📝 匯入自建訂單", "🔗 組合貨號"])

with tab_storage:
    _render_storage_tab()
with tab_order:
    _render_order_tab()
with tab_custom:
    _render_custom_tab()
with tab_combo:
    _render_combo_tab()
