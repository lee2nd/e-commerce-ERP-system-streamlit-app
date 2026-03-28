"""單品銷售查詢 & 銷售排行"""
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px

st.set_page_config(page_title="單品銷售查詢 & 銷售排行", page_icon="🏆", layout="wide")

from utils.data_manager import load_delivery, load_inventory_details

st.title("🏆 單品銷售查詢 & 銷售排行")

# ── 載入資料 ───────────────────────────────────────────────────
delivery_raw = load_delivery()
inventory    = load_inventory_details()

_NULL_LIKE = frozenset({"nan", "none", "nat", "<na>", "", "未匹配", "tbd"})

def _clean(v) -> str:
    s = str(v).strip() if pd.notna(v) else ""
    return "" if s.lower() in _NULL_LIKE else s

# 只取已匹配（貨號非空且非 TBD）
if not delivery_raw.empty:
    delivery = delivery_raw[
        delivery_raw["貨號"].apply(_clean).astype(bool)
        & ~delivery_raw["名稱"].fillna("").str.upper().str.contains("TBD")
    ].copy()
    delivery["_日期"] = pd.to_datetime(delivery["出庫日期"], errors="coerce")
    delivery["年"] = delivery["_日期"].dt.year.astype("Int64")
    delivery["月"] = delivery["_日期"].dt.month.astype("Int64")
else:
    delivery = pd.DataFrame()

MONTHS_LABEL = [f"{m}月" for m in range(1, 13)]

tab_single, tab_rank = st.tabs(["🔍 單品銷售查詢", "📊 銷售排行"])


# ══════════════════════════════════════════════════════════════
# Tab 1 — 單品銷售查詢
# ══════════════════════════════════════════════════════════════
with tab_single:

    if delivery.empty:
        st.info("尚無已匹配的出庫資料，請先完成「導出出庫」並確認對照表已匹配")
        st.stop()

    # ── 建立 distinct (主貨號 + 商品名稱) 下拉選單 ────────────
    def _item_key(row) -> str:
        main = _clean(row.get("主貨號", ""))
        name = _clean(row.get("名稱", ""))
        return f"({main}){name}" if main else name

    delivery["_item_key"] = delivery.apply(_item_key, axis=1)
    item_options = sorted(
        delivery["_item_key"].dropna().unique().tolist()
    )
    item_options = [k for k in item_options if k]

    if not item_options:
        st.warning("無有效商品資料")
        st.stop()

    # 年份選單
    all_years = sorted(
        delivery["年"].dropna().astype(int).unique().tolist(), reverse=True
    )
    col_sel1, col_sel2 = st.columns([3, 1])
    with col_sel1:
        selected_item = st.selectbox("選擇商品（主貨號）商品名稱", item_options, key="si_item")
    with col_sel2:
        selected_year = st.selectbox("年份", all_years, key="si_year")

    # 取出該商品、該年份的出庫資料
    mask = (delivery["_item_key"] == selected_item) & (delivery["年"] == selected_year)
    df_item = delivery[mask].copy()

    # 主貨號底下的各貨號
    sub_skus = df_item["貨號"].apply(_clean)
    sub_skus = sub_skus[sub_skus != ""].unique().tolist()

    # ── 月度彙總 ─────────────────────────────────────────────
    monthly = pd.DataFrame({"月": range(1, 13)})

    # 總營業額（出庫金額）
    if "金額" in df_item.columns:
        rev_m = (
            df_item.groupby("月")["金額"].sum().reset_index()
            .rename(columns={"金額": "總營業額"})
        )
    else:
        rev_m = pd.DataFrame({"月": range(1, 13), "總營業額": 0})
    monthly = monthly.merge(rev_m, on="月", how="left")
    monthly["總營業額"] = monthly["總營業額"].fillna(0).astype(int)

    qty_col = "出庫數量"
    if not inventory.empty and "貨號" in inventory.columns and "平均成本" in inventory.columns:
        cost_map = inventory.set_index("貨號")["平均成本"].to_dict()
        df_item["_單位成本"] = df_item["貨號"].apply(_clean).map(cost_map).fillna(0)

        # Safely get qty as a Series, defaulting to 1 if column is missing
        if qty_col in df_item.columns:
            qty_series = pd.to_numeric(df_item[qty_col], errors="coerce").fillna(0)
        else:
            qty_series = pd.Series(1, index=df_item.index)

        df_item["_成本"] = df_item["_單位成本"] * qty_series
        cost_m = df_item.groupby("月")["_成本"].sum().reset_index().rename(columns={"_成本": "總成本"})
        monthly = monthly.merge(cost_m, on="月", how="left")
    else:
        monthly["總成本"] = 0
    monthly["總成本"] = monthly["總成本"].fillna(0).round(0).astype(int)
    monthly["總淨利額"] = monthly["總營業額"] - monthly["總成本"]

    # ── 摘要指標 ─────────────────────────────────────────────
    st.markdown(f"#### {selected_year} 年 — {selected_item}")
    mc1, mc2, mc3 = st.columns(3)
    mc1.metric("總營業額", f"${int(monthly['總營業額'].sum()):,}")
    mc2.metric("總成本",   f"${int(monthly['總成本'].sum()):,}")
    mc3.metric("總淨利額", f"${int(monthly['總淨利額'].sum()):,}")

    # ── 年度銷售表 ────────────────────────────────────────────
    st.markdown("##### 年度銷售表")
    display_table = monthly.copy()
    display_table.insert(0, "月份", MONTHS_LABEL)
    display_table = display_table.drop(columns=["月"])
    # 合計列
    total_row: dict[str, object] = {
        "月份": "合計",
        "總營業額": int(display_table["總營業額"].sum()),
        "總成本": int(display_table["總成本"].sum()),
        "總淨利額": int(display_table["總淨利額"].sum()),
    }
    display_table = pd.concat([display_table, pd.DataFrame([total_row])], ignore_index=True)
    st.dataframe(display_table, hide_index=True, width='stretch')

    # ── 圖表 ─────────────────────────────────────────────────
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        name="總營業額", x=MONTHS_LABEL, y=monthly["總營業額"].tolist(),
        mode="lines+markers+text",
        line=dict(color="#2E8B57", width=2),
        text=["$" + f"{v:,}" if v != 0 else "" for v in monthly["總營業額"]],
        textposition="top center",
    ))
    fig.add_trace(go.Scatter(
        name="總淨利額", x=MONTHS_LABEL, y=monthly["總淨利額"].tolist(),
        mode="lines+markers+text",
        line=dict(color="#FF8C00", width=2),
        text=["$" + f"{v:,}" if v != 0 else "" for v in monthly["總淨利額"]],
        textposition="bottom center",
    ))
    fig.update_layout(
        title=f"{selected_year} 年 — {selected_item}",
        height=400, hovermode="x unified",
        yaxis=dict(title="金額 ($)", tickformat="$,d"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    st.plotly_chart(fig, width='stretch')

    # ── 各貨號（規格）銷量排行 ────────────────────────────────
    st.markdown("##### 各單品貨號銷量排行")
    _has_spec = "規格" in df_item.columns
    _grp_cols = ["貨號", "規格"] if _has_spec else ["貨號"]
    if sub_skus and qty_col in df_item.columns:
        sku_rank = (
            df_item.groupby(_grp_cols)
            .agg(
                銷售數量=(qty_col, "sum"),
                銷售金額=("金額", "sum"),
            )
            .reset_index()
            .sort_values("銷售數量", ascending=False)
            .reset_index(drop=True)
        )
        if not _has_spec:
            sku_rank["規格"] = ""
        sku_rank["規格"] = sku_rank["規格"].apply(_clean)
        sku_rank["排名"] = range(1, len(sku_rank) + 1)

        if not sku_rank.empty:
            fig_sku = px.bar(
                sku_rank,
                x="貨號", y="銷售數量",
                color="規格" if sku_rank["規格"].any() else None,
                text="銷售數量",
                title="各單品貨號銷售數量",
                height=380,
            )
            fig_sku.update_traces(textposition="outside")
            fig_sku.update_layout(yaxis_title="銷售數量")
            st.plotly_chart(fig_sku, width='stretch')
            st.dataframe(sku_rank[["排名", "貨號", "規格", "銷售數量", "銷售金額"]], hide_index=True, width='stretch')
    else:
        st.info("無細項貨號資料")


# ══════════════════════════════════════════════════════════════
# Tab 2 — 銷售排行
# ══════════════════════════════════════════════════════════════
with tab_rank:

    if delivery.empty:
        st.info("尚無已匹配的出庫資料")
        st.stop()

    st.markdown("#### 篩選條件")
    rank_col1, rank_col2 = st.columns(2)

    all_years_rank = sorted(
        delivery["年"].dropna().astype(int).unique().tolist(), reverse=True
    )
    all_months_rank = list(range(1, 13))

    with rank_col1:
        sel_years = st.multiselect(
            "年份（可多選）", options=all_years_rank,
            default=[all_years_rank[0]] if all_years_rank else [],
            key="rank_years",
        )
    with rank_col2:
        sel_months = st.multiselect(
            "月份（可多選，不選則全年）", options=all_months_rank,
            default=[],
            format_func=lambda m: f"{m}月",
            key="rank_months",
        )

    if not sel_years:
        st.warning("請選擇至少一個年份")
        st.stop()

    # 篩選
    mask_r = delivery["年"].isin(sel_years)
    if sel_months:
        mask_r = mask_r & delivery["月"].isin(sel_months)
    df_rank = delivery[mask_r].copy()

    if df_rank.empty:
        st.warning("所選條件無資料")
        st.stop()

    # 建立 主貨號商品 key
    df_rank["_item_key"] = df_rank.apply(_item_key, axis=1)

    # 彙總：各主貨號 銷售金額 & 數量
    qty_col_r = "出庫數量"
    rank_agg = (
        df_rank.groupby("_item_key")
        .agg(
            銷售金額=("金額", "sum"),
            銷售數量=(qty_col_r, "sum"),
        )
        .reset_index()
        .rename(columns={"_item_key": "商品"})
    )

    top_n = st.slider("顯示前 N 名", 5, 50, 20, key="rank_topn")

    # ── 銷售金額排行 ──────────────────────────────────────────
    st.markdown("#### 💰 銷售金額排行")
    rank_by_rev = rank_agg.sort_values("銷售金額", ascending=False).head(top_n).reset_index(drop=True)
    rank_by_rev["排名"] = range(1, len(rank_by_rev) + 1)

    if not rank_by_rev.empty:
        fig_rev = px.bar(
            rank_by_rev,
            x="銷售金額", y="商品",
            orientation="h",
            text="銷售金額",
            title=f"銷售金額前 {top_n} 名",
            height=max(380, len(rank_by_rev) * 28),
            color="銷售金額",
            color_continuous_scale="Greens",
        )
        fig_rev.update_traces(texttemplate="$%{text:,}", textposition="outside")
        fig_rev.update_layout(
            yaxis=dict(autorange="reversed"),
            coloraxis_showscale=False,
            xaxis_title="銷售金額 ($)",
        )
        st.plotly_chart(fig_rev, width='stretch')
        st.dataframe(
            rank_by_rev[["排名", "商品", "銷售金額", "銷售數量"]].assign(
                銷售金額=rank_by_rev["銷售金額"].apply(lambda v: f"${int(v):,}")
            ),
            hide_index=True, width='stretch',
        )

    st.markdown("---")

    # ── 銷售數量排行 ──────────────────────────────────────────
    st.markdown("#### 📦 銷售數量排行")
    rank_by_qty = rank_agg.sort_values("銷售數量", ascending=False).head(top_n).reset_index(drop=True)
    rank_by_qty["排名"] = range(1, len(rank_by_qty) + 1)

    if not rank_by_qty.empty:
        fig_qty = px.bar(
            rank_by_qty,
            x="銷售數量", y="商品",
            orientation="h",
            text="銷售數量",
            title=f"銷售數量前 {top_n} 名",
            height=max(380, len(rank_by_qty) * 28),
            color="銷售數量",
            color_continuous_scale="Blues",
        )
        fig_qty.update_traces(textposition="outside")
        fig_qty.update_layout(
            yaxis=dict(autorange="reversed"),
            coloraxis_showscale=False,
            xaxis_title="銷售數量（件）",
        )
        st.plotly_chart(fig_qty, width='stretch')
        st.dataframe(
            rank_by_qty[["排名", "商品", "銷售數量", "銷售金額"]].assign(
                銷售金額=rank_by_qty["銷售金額"].apply(lambda v: f"${int(v):,}")
            ),
            hide_index=True, width='stretch',
        )
