"""日報表"""
import streamlit as st
import pandas as pd
from utils.data_manager import (
    load_platform_orders,
    load_compare_table, load_storage,
    load_daily_report, save_daily_report, clear_daily_report,
    load_combo_sku,
)
from utils.calculators import generate_daily_report

st.set_page_config(page_title="日報表", page_icon="📊", layout="wide")
st.title("📊 日報表")

settings = {
    "ruten_7_11": 60,
    "ruten_family": 60,
    "ruten_ok": 60,
    "ruten_laerfu": 50,
    "ruten_post": 65,
    "ruten_default_shipping": 65,
    "easystore_shipping": 65,
}

# ── 產生報表 ──────────────────────────────────────────────────
if st.button("🔄 重新產生日報表", type="primary"):
    shopee_raw    = load_platform_orders("蝦皮")
    ruten_raw     = load_platform_orders("露天")
    easystore_raw = load_platform_orders("官網")

    if shopee_raw.empty and ruten_raw.empty and easystore_raw.empty:
        st.warning("無訂單資料，請先至「匯入資料」頁面匯入平台訂單")
    else:
        compare = load_compare_table()
        storage = load_storage()
        with st.spinner("計算中…"):
            daily = generate_daily_report(
                shopee_raw, ruten_raw, easystore_raw,
                compare, storage, settings, load_combo_sku(),
            )
        if daily.empty:
            st.warning("計算結果為空，請確認資料是否正確")
        else:
            new_df = daily.drop(columns=["_unmatched"], errors="ignore")
            # 現有日報表已有的訂單保持不變，只新增從未出現過的訂單
            existing_daily = load_daily_report()
            if not existing_daily.empty and "訂單編號" in existing_daily.columns and "日期" in existing_daily.columns:
                # 部份退貨會產生兩筆相同訂單編號+日期（一筆已完成、一筆退貨），需加入訂單狀態區分
                _stat_ex = existing_daily["訂單狀態"].astype(str) if "訂單狀態" in existing_daily.columns else ""
                _stat_new = new_df["訂單狀態"].astype(str) if "訂單狀態" in new_df.columns else ""
                existing_keys = set(
                    existing_daily["訂單編號"].astype(str) + "||" + existing_daily["日期"].astype(str) + "||" + _stat_ex
                )
                new_keys = new_df["訂單編號"].astype(str) + "||" + new_df["日期"].astype(str) + "||" + _stat_new
                truly_new = new_df[~new_keys.isin(existing_keys)]
                merged = pd.concat([existing_daily, truly_new], ignore_index=True)
                save_daily_report(merged)
                added = len(truly_new)
                st.success(f"✅ 日報表已更新！新增 {added} 筆，保留 {len(existing_daily)} 筆既有記錄")
            else:
                save_daily_report(new_df)
                st.success(f"✅ 日報表已產生，共 {len(new_df)} 筆")
            st.rerun()

# ── 顯示報表 ──────────────────────────────────────────────────
daily = load_daily_report()

if daily.empty:
    st.info("尚未產生日報表，請按上方按鈕產生")
    st.stop()

daily["日期"] = pd.to_datetime(daily["日期"], errors="coerce")

# 舊資料回填備註「未匹配」
if "備註" not in daily.columns:
    daily["備註"] = ""
daily["備註"] = daily["備註"].fillna("").astype(str)
# 舊資料相容：備註為空且商品名稱也為空 → 代表當時沒匹配到，補標「未匹配」
if "商品名稱" in daily.columns:
    _need_fill = (
        (daily["備註"].fillna("").astype(str).str.strip() == "") &
        (daily["商品名稱"].fillna("").astype(str).str.strip() == "")
    )
    if _need_fill.any():
        daily.loc[_need_fill, "備註"] = "未匹配"
        save_daily_report(daily.drop(columns=["_unmatched"], errors="ignore"))

# 舊資料修正：退貨/未取貨且商品名稱有值，代表 SKU 有匹配，把誤標的「未匹配」清除
if "訂單狀態" in daily.columns and "商品名稱" in daily.columns:
    _fix_wrong = (
        daily["訂單狀態"].isin(["退貨", "未取貨"]) &
        daily["備註"].astype(str).str.contains("未匹配", na=False) &
        (daily["商品名稱"].fillna("").astype(str).str.strip() != "")
    )
    if _fix_wrong.any():
        daily.loc[_fix_wrong, "備註"] = ""
        save_daily_report(daily.drop(columns=["_unmatched"], errors="ignore"))

# 篩選器
st.markdown("---")
cf1, cf2, cf3, cf4 = st.columns(4)
plat_opts = sorted(daily["平台"].dropna().unique())
plat_filter = cf1.multiselect("平台", options=plat_opts, default=plat_opts)

status_opts = sorted(daily["訂單狀態"].dropna().unique())
status_filter = cf2.multiselect("訂單狀態", options=status_opts, default=status_opts)

min_d = daily["日期"].min()
max_d = daily["日期"].max()
if pd.notna(min_d) and pd.notna(max_d):
    date_range = cf3.date_input(
        "日期範圍",
        value=(min_d.date(), max_d.date()),
        min_value=min_d.date(), max_value=max_d.date(),
    )
else:
    date_range = None

unmatched_filter = cf4.radio("是否未匹配", options=["全部", "未匹配", "已匹配"], horizontal=True)

view = daily.copy()
view = view[view["平台"].isin(plat_filter)]
view = view[view["訂單狀態"].isin(status_filter)]
if date_range and len(date_range) == 2:
    view = view[
        (view["日期"].dt.date >= date_range[0]) &
        (view["日期"].dt.date <= date_range[1])
    ]
if unmatched_filter == "未匹配":
    view = view[view["備註"].astype(str).str.contains("未匹配", na=False)]
elif unmatched_filter == "已匹配":
    view = view[~view["備註"].astype(str).str.contains("未匹配", na=False)]

# 摘要
st.markdown("### 摘要")
s1, s2, s3, s4 = st.columns(4)
s1.metric("筆數", f"{len(view):,}")
s2.metric("訂單金額", f"${int(view['訂單金額'].sum()):,}" if "訂單金額" in view.columns else "—")
s3.metric("總成本", f"${int(view['總成本'].sum()):,}" if "總成本" in view.columns else "—")
s4.metric("淨利", f"${int(view['淨利'].sum()):,}" if "淨利" in view.columns else "—")

# 顏色
_PLAT_COLORS = {"蝦皮": "#FF6B35", "露天": "#4A90D9", "官網": "#2ECC71"}

money_cols = [c for c in ["訂單金額", "折扣優惠", "買家支付運費", "平台補助運費",
                           "實際運費支出", "物流處理費（運費差額）", "未取貨/退貨運費",
                           "成交手續費", "其他服務費", "金流與系統處理費",
                           "發票處理費", "其他費用", "商品成本", "總成本", "淨利"]
              if c in view.columns]

st.markdown("### 明細")

# 準備 editor 用 DataFrame（移除內部欄位，保留原始 index 供回寫）
view_edit = view.drop(columns=["_unmatched"], errors="ignore").copy()

col_cfg: dict = {}
# 識別欄位：唯讀
for _c in ["訂單編號", "平台"]:
    if _c in view_edit.columns:
        col_cfg[_c] = st.column_config.TextColumn(disabled=True)
# 日期欄位
if "日期" in view_edit.columns:
    col_cfg["日期"] = st.column_config.DateColumn("日期", format="YYYY-MM-DD")
# 訂單狀態：下拉選單
if "訂單狀態" in view_edit.columns:
    col_cfg["訂單狀態"] = st.column_config.SelectboxColumn(
        options=["已完成", "退貨", "未取貨", "遺失賠償"]
    )
# 金額欄位
for _c in money_cols:
    col_cfg[_c] = st.column_config.NumberColumn(format="$%d")

edited_df = st.data_editor(
    view_edit,
    key="daily_editor",
    width='stretch',
    hide_index=True,
    num_rows="fixed",
    height=500,
    column_config=col_cfg,
)

if st.button("💾 儲存修改", key="save_daily_edit"):
    # 以原始 index 把編輯後的列寫回完整 daily（非篩選部分保持不變）
    base = daily.drop(columns=["_unmatched"], errors="ignore").copy()
    save_edit = edited_df.copy()
    save_edit["日期"] = pd.to_datetime(save_edit["日期"], errors="coerce")
    for col in save_edit.columns:
        if col in base.columns:
            base.loc[save_edit.index, col] = save_edit[col].values

    # 重算被編輯列的總成本與淨利
    _cost_cols = ["商品成本", "折扣優惠", "未取貨/退貨運費",
                  "物流處理費（運費差額）",
                  "成交手續費", "其他服務費", "金流與系統處理費",
                  "發票處理費", "其他費用"]
    for _c in _cost_cols:
        if _c not in base.columns:
            base[_c] = 0
    idx = save_edit.index

    # 重算物流處理費（運費差額）：依平台公式
    if "物流處理費（運費差額）" in base.columns and "平台" in base.columns:
        for i in idx:
            plat = base.at[i, "平台"]
            actual = float(base.at[i, "實際運費支出"] or 0) if "實際運費支出" in base.columns else 0 # pyright: ignore[reportArgumentType]
            buyer = float(base.at[i, "買家支付運費"] or 0) if "買家支付運費" in base.columns else 0 # pyright: ignore[reportArgumentType]
            if plat == "露天":
                base.at[i, "物流處理費（運費差額）"] = max(0, actual - buyer)
            elif plat == "官網":
                base.at[i, "物流處理費（運費差額）"] = max(0, buyer - actual)

    base.loc[idx, "總成本"] = sum(
        base.loc[idx, _c].fillna(0) for _c in _cost_cols
    )
    base.loc[idx, "淨利"] = (
        base.loc[idx, "訂單金額"].fillna(0) - base.loc[idx, "總成本"]
    )

    save_daily_report(base)
    st.success("✅ 修改已儲存，總成本與淨利已重新計算")
    st.rerun()

# ── 清0 日報表 ─────────────────────────────────────────────
st.markdown("---")
if st.button("🗑️ 清除日報表", key="clear_daily_btn"):
    st.session_state["confirm_clear_daily"] = True
if st.session_state.get("confirm_clear_daily"):
    st.warning("⚠️ 確定要清除所有日報表資料嗎？此操作無法復原！")
    _c1, _c2 = st.columns(2)
    if _c1.button("✅ 確認清除", key="confirm_clear_daily_yes", type="primary"):
        with st.spinner("清除中…"):
            clear_daily_report()
        st.session_state.pop("confirm_clear_daily", None)
        st.success("✅ 日報表已清除")
        st.rerun()
    if _c2.button("❌ 取消", key="confirm_clear_daily_no"):
        st.session_state.pop("confirm_clear_daily", None)
        st.rerun()
