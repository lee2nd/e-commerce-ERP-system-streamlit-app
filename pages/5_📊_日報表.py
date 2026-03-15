"""日報表"""
import streamlit as st
import pandas as pd
from utils.data_manager import (
    load_platform_orders,
    load_compare_table, load_storage,
    load_daily_report, save_daily_report, clear_daily_report,
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
                compare, storage, settings,
            )
        if daily.empty:
            st.warning("計算結果為空，請確認資料是否正確")
        else:
            save_daily_report(daily.drop(columns=["_unmatched"], errors="ignore"))
            st.success(f"✅ 日報表已產生，共 {len(daily)} 筆")
            st.rerun()

# ── 顯示報表 ──────────────────────────────────────────────────
daily = load_daily_report()

if daily.empty:
    st.info("尚未產生日報表，請按上方按鈕產生")
    st.stop()

daily["日期"] = pd.to_datetime(daily["日期"], errors="coerce")

# 篩選器
st.markdown("---")
cf1, cf2, cf3 = st.columns(3)
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

view = daily.copy()
view = view[view["平台"].isin(plat_filter)]
view = view[view["訂單狀態"].isin(status_filter)]
if date_range and len(date_range) == 2:
    view = view[
        (view["日期"].dt.date >= date_range[0]) &
        (view["日期"].dt.date <= date_range[1])
    ]

# 摘要
st.markdown("### 摘要")
s1, s2, s3, s4 = st.columns(4)
s1.metric("筆數", f"{len(view):,}")
s2.metric("訂單金額", f"${int(view['訂單金額'].sum()):,}" if "訂單金額" in view.columns else "—")
s3.metric("總成本", f"${int(view['總成本'].sum()):,}" if "總成本" in view.columns else "—")
s4.metric("淨利", f"${int(view['淨利'].sum()):,}" if "淨利" in view.columns else "—")

# 顏色
_PLAT_COLORS = {"蝦皮": "#FF6B35", "露天": "#4A90D9", "官網": "#2ECC71"}
_PLAT_BG     = {"蝦皮": "#FFF0EB", "露天": "#EAF3FB", "官網": "#E9F9F1"}
_STATUS_COLORS = {"退貨": "#e74c3c", "未取貨": "#e67e22", "遺失賠償": "#9b59b6"}

def _highlight(row):
    status_color = _STATUS_COLORS.get(row.get("訂單狀態", ""), "")
    plat_color   = _PLAT_COLORS.get(row.get("平台", ""), "")
    plat_bg      = _PLAT_BG.get(row.get("平台", ""), "")
    styles = []
    for c in row.index:
        bg = f"background-color: {plat_bg}; " if plat_bg else ""
        if c == "平台" and plat_color:
            styles.append(f"{bg}color: {plat_color}; font-weight: bold")
        elif status_color and c == "訂單狀態":
            styles.append(f"{bg}color: {status_color}; font-weight: bold")
        else:
            styles.append(bg.rstrip("; "))
    return styles

# 表格 — drop internal column, format date
display = view.drop(columns=["_unmatched"], errors="ignore").copy()
display["日期"] = display["日期"].apply(
    lambda d: f"{d.year}年{d.month}月{d.day}日" if pd.notna(d) else ""
)

money_cols = [c for c in ["訂單金額", "折扣優惠", "買家支付運費", "平台補助運費",
                           "實際運費支出", "物流處理費（運費差額）", "未取貨/退貨運費",
                           "成交手續費", "其他服務費", "金流與系統處理費",
                           "發票處理費", "其他費用", "商品成本", "總成本", "淨利"]
              if c in display.columns]

col_cfg = {c: st.column_config.NumberColumn(format="$%d") for c in money_cols}

st.markdown("### 明細")
st.dataframe(
    display.style.apply(_highlight, axis=1),
    width='stretch',
    hide_index=True,
    column_config=col_cfg,
)

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
