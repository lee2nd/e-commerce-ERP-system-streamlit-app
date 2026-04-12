"""月報表"""
import streamlit as st
import pandas as pd
from datetime import datetime, timezone, timedelta
from utils.data_manager import (
    load_daily_report,
    load_monthly_report, save_monthly_report, clear_monthly_report,
)
from utils.calculators import compute_monthly_auto_from_daily

TZ_TAIPEI = timezone(timedelta(hours=8))
st.set_page_config(page_title="月報表", page_icon="📈", layout="wide")
from utils.styles import apply_global_styles
apply_global_styles()
st.title("📈 月報表")

# ── 欄位定義 ──────────────────────────────────────────────────
_FIXED_DEFAULTS: dict[str, int] = {
    "官網月費": 1231,
    "會計記帳費": 1900,
    "會費＆勞健保費": 4107,
    "營登租金": 1155,
}

AUTO_COLS = [
    "營業額", "商品成本", "折扣優惠", "手續費",
    "未取貨/退貨運費", "物流處理費（運費差額）",
]
MANUAL_COLS = [
    "其他費用（一）", "其他費用（二）", "薪資費用", "委製費用",
    "蝦皮廣告費用", "露天廣告費用", "MO店廣告費用", "官網廣告費用", "其他廣告費用",
    "耗材費用", "營業稅額", "官網月費", "會計記帳費", "會費＆勞健保費", "營登租金",
]
DERIVED_COLS = ["廣告費用合計", "總成本", "淨利"]
ALL_COLS = ["年份", "月份"] + AUTO_COLS + MANUAL_COLS + DERIVED_COLS + ["備註"]


def _compute_derived(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    def _n(col: str) -> pd.Series:
        if col not in df.columns:
            return pd.Series(0, index=df.index, dtype="float64")
        return pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["廣告費用合計"] = (
        _n("蝦皮廣告費用") + _n("露天廣告費用") + _n("MO店廣告費用")
        + _n("官網廣告費用") + _n("其他廣告費用")
    )
    df["總成本"] = (
        _n("商品成本") + _n("折扣優惠") + _n("手續費")
        + _n("未取貨/退貨運費") + _n("物流處理費（運費差額）")
        + _n("其他費用（一）") + _n("其他費用（二）")
        + _n("薪資費用") + _n("委製費用") + _n("廣告費用合計")
        + _n("耗材費用") + _n("營業稅額")
        + _n("官網月費") + _n("會計記帳費") + _n("會費＆勞健保費")
        + _n("營登租金")
    )
    df["淨利"] = _n("營業額") - df["總成本"]
    for c in DERIVED_COLS:
        df[c] = df[c].round(0).astype(int)
    return df


# ── 更新月報表 ─────────────────────────────────────────────────
if st.button("🔄 更新月報表（從日報表重算）", type="primary"):
    daily = load_daily_report()
    if daily.empty:
        st.warning("請先至「日報表」頁面產生日報表")
    else:
        auto_df = compute_monthly_auto_from_daily(daily)
        if auto_df.empty:
            st.warning("日報表計算結果為空")
        else:
            existing = load_monthly_report()
            merged = auto_df.copy()

            # 填入 manual 欄位：優先從現有月報表保留手動輸入值
            if not existing.empty:
                existing["年份"] = pd.to_numeric(existing["年份"], errors="coerce").fillna(0).astype(int)
                existing["月份"] = pd.to_numeric(existing["月份"], errors="coerce").fillna(0).astype(int)
                ex_lookup = existing.set_index(["年份", "月份"])
                for col in MANUAL_COLS:
                    if col in ex_lookup.columns:
                        merged[col] = merged.apply(
                            lambda r: ex_lookup.at[(int(r["年份"]), int(r["月份"])), col]
                            if (int(r["年份"]), int(r["月份"])) in ex_lookup.index else _FIXED_DEFAULTS.get(col, 0),
                            axis=1,
                        )
                    else:
                        merged[col] = _FIXED_DEFAULTS.get(col, 0)
                # 保留備註
                if "備註" in ex_lookup.columns:
                    merged["備註"] = merged.apply(
                        lambda r: ex_lookup.at[(int(r["年份"]), int(r["月份"])), "備註"]
                        if (int(r["年份"]), int(r["月份"])) in ex_lookup.index else "",
                        axis=1,
                    )
                else:
                    merged["備註"] = ""
            else:
                for col in MANUAL_COLS:
                    merged[col] = _FIXED_DEFAULTS.get(col, 0)
                merged["備註"] = ""

            # 確保所有欄位存在
            for c in ALL_COLS:
                if c not in merged.columns:
                    merged[c] = "" if c == "備註" else _FIXED_DEFAULTS.get(c, 0)

            merged = _compute_derived(merged)
            merged["備註"] = merged["備註"].fillna("").astype(str)
            merged = merged[ALL_COLS].sort_values(["年份", "月份"]).reset_index(drop=True)
            save_monthly_report(merged)
            st.success(f"✅ 月報表已更新，共 {len(merged)} 筆")
            st.rerun()

# ── 載入月報表 ─────────────────────────────────────────────────
monthly = load_monthly_report()

if monthly.empty:
    st.info("尚未產生月報表，請按上方「🔄 更新月報表」按鈕")
    st.stop()

# 確保欄位完整 & 數值型態
for c in ALL_COLS:
    if c not in monthly.columns:
        monthly[c] = "" if c == "備註" else _FIXED_DEFAULTS.get(c, 0)
for c in AUTO_COLS + MANUAL_COLS + DERIVED_COLS:
    monthly[c] = pd.to_numeric(monthly[c], errors="coerce").fillna(0).astype(int)
monthly["備註"] = monthly["備註"].fillna("").astype(str)
monthly = monthly[ALL_COLS].copy()

# ── 摘要 ───────────────────────────────────────────────────────
c1, c2, c3, c4 = st.columns(4)
c1.metric("月份數", len(monthly))
c2.metric("累計營業額", f"${monthly['營業額'].sum():,}")
c3.metric("累計總成本", f"${monthly['總成本'].sum():,}")
c4.metric("累計淨利", f"${monthly['淨利'].sum():,}")

# ── 下載 ──────────────────────────────────────────────────────
st.markdown("### 月報表明細")

_csv = monthly.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
st.download_button("⬇️ 下載月報表", data=_csv, file_name="月報表.csv", mime="text/csv", key="dl_monthly")
view_slice = monthly.copy()

# ── Column config ──────────────────────────────────────────────
_col_cfg: dict = {}
# 唯讀欄位（自動計算）
for c in ["年份", "月份"]:
    _col_cfg[c] = st.column_config.NumberColumn(c, disabled=True, format="%d")
for c in AUTO_COLS:
    _col_cfg[c] = st.column_config.NumberColumn(c, disabled=True, format="$%d")
for c in DERIVED_COLS:
    _col_cfg[c] = st.column_config.NumberColumn(c, disabled=True, format="$%d")
# 可編輯欄位（手動輸入）
for c in MANUAL_COLS:
    _col_cfg[c] = st.column_config.NumberColumn(c, format="$%d")
# 備註欄位（可編輯文字）
_col_cfg["備註"] = st.column_config.TextColumn("備註")

# ── Data editor ────────────────────────────────────────────────
edited_df = st.data_editor(
    view_slice,
    key="monthly_editor",
    width="stretch",
    hide_index=True,
    num_rows="fixed",
    height=500,
    column_config=_col_cfg,
)

if st.button("💾 儲存修改", key="save_monthly_edit"):
    base = monthly.copy()
    save_edit = edited_df.copy()
    for col in save_edit.columns:
        if col in base.columns:
            base.loc[save_edit.index, col] = save_edit[col].values
    base = _compute_derived(base)
    base = base[ALL_COLS]
    save_monthly_report(base)
    st.session_state["monthly_saved_at"] = datetime.now(tz=TZ_TAIPEI).strftime("%Y-%m-%d %H:%M:%S")
    st.success("✅ 修改已儲存，廣告費用合計、總成本與淨利已重算")
    st.rerun()

_monthly_ts = st.session_state.get("monthly_saved_at")
if _monthly_ts:
    st.caption(f"🕐 最後儲存：{_monthly_ts}")

# ── 清0 月報表 ─────────────────────────────────────────────────
st.markdown("---")
if st.button("🗑️ 清除月報表", key="clear_monthly_btn"):
    st.session_state["confirm_clear_monthly"] = True
if st.session_state.get("confirm_clear_monthly"):
    st.warning("⚠️ 確定要清除所有月報表資料嗎？此操作無法復原！")
    _c1, _c2 = st.columns(2)
    if _c1.button("✅ 確認清除", key="confirm_clear_monthly_yes", type="primary"):
        with st.spinner("清除中…"):
            clear_monthly_report()
        st.session_state.pop("confirm_clear_monthly", None)
        st.success("✅ 月報表已清除")
        st.rerun()
    if _c2.button("❌ 取消", key="confirm_clear_monthly_no"):
        st.session_state.pop("confirm_clear_monthly", None)
        st.rerun()
