"""數據圖表"""
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

st.set_page_config(page_title="數據圖表", page_icon="📉", layout="wide")

from utils.data_manager import load_daily_report

st.info("⚠️ 請確認日、月報表都修改完畢")
st.title("📉 數據圖表")

daily = load_daily_report()
if daily.empty:
    st.info("尚無日報表資料，請先至「日報表」頁面產生報表")
    st.stop()

# ── 前處理 ─────────────────────────────────────────────────────
daily["日期"] = pd.to_datetime(daily["日期"], errors="coerce")
daily = daily.dropna(subset=["日期"])
daily["年"] = daily["日期"].dt.year.astype(int)
daily["月"] = daily["日期"].dt.month.astype(int)

if daily.empty:
    st.warning("日報表中無可解析的日期資料")
    st.stop()

# ── 年份選擇 ───────────────────────────────────────────────────
all_years = sorted(daily["年"].unique().tolist(), reverse=True)
selected_years = st.multiselect(
    "選擇年份（可多選進行跨年比較）",
    options=all_years,
    default=[all_years[0]] if all_years else [],
)
if not selected_years:
    st.warning("請選擇至少一個年份")
    st.stop()

# ── 常數 ──────────────────────────────────────────────────────
PLATFORMS = ["蝦皮", "露天", "官網", "MO店"]
PLAT_COLORS = {
    "蝦皮": "#636EFA",
    "露天": "#EF553B",
    "官網": "#00CC96",
    "MO店": "#AB63FA",
}
MONTHS_LABEL = [f"{m}月" for m in range(1, 13)]


# ══════════════════════════════════════════════════════════════
# 計算函式
# ══════════════════════════════════════════════════════════════

def _build_year_data(df_year: pd.DataFrame) -> pd.DataFrame:
    """
    回傳 12 列（1-12 月）的統計 DataFrame：
    - 各平台訂單量（已完成 distinct 訂單編號）
    - 總營業額（已完成 訂單金額加總）
    - 總淨利額（全部狀態 淨利加總）
    - _total（所有平台訂單量合計）
    - 年度平均訂單量（_total.sum / 12）
    """
    result = pd.DataFrame({"月": range(1, 13)})

    completed = df_year[df_year["訂單狀態"] == "已完成"].copy()

    # ── 各平台訂單量（已完成 distinct 訂單編號）──────────────
    for plat in PLATFORMS:
        cnt = (
            completed[completed["平台"] == plat]
            .groupby("月")["訂單編號"]
            .nunique()
            .reset_index()
            .rename(columns={"訂單編號": plat})
        )
        result = result.merge(cnt, on="月", how="left")
        result[plat] = result[plat].fillna(0).astype(int)

    result["_total"] = result[PLATFORMS].sum(axis=1)

    # ── 總營業額（已完成 訂單金額加總）──────────────────────
    if "訂單金額" in completed.columns:
        rev = (
            completed.groupby("月")["訂單金額"]
            .sum()
            .reset_index()
            .rename(columns={"訂單金額": "總營業額"})
        )
    else:
        rev = pd.DataFrame({"月": range(1, 13), "總營業額": 0})
    result = result.merge(rev, on="月", how="left")
    result["總營業額"] = result["總營業額"].fillna(0).astype(int)

    # ── 總淨利額（全部狀態 淨利加總）────────────────────────
    if "淨利" in df_year.columns:
        pft = (
            df_year.groupby("月")["淨利"]
            .sum()
            .reset_index()
            .rename(columns={"淨利": "總淨利額"})
        )
    else:
        pft = pd.DataFrame({"月": range(1, 13), "總淨利額": 0})
    result = result.merge(pft, on="月", how="left")
    result["總淨利額"] = result["總淨利額"].fillna(0).astype(int)

    # ── 年度平均訂單量 ─────────────────────────────────────
    result["年度平均訂單量"] = round(result["_total"].sum() / 12, 1)

    return result


def _make_chart(data: pd.DataFrame, year: int) -> go.Figure:
    avg = float(data["年度平均訂單量"].iloc[0])
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    # Bars – 各平台訂單量（左軸）
    for plat in PLATFORMS:
        ys = data[plat].tolist()
        texts = [str(v) if v > 0 else "" for v in ys]
        fig.add_trace(
            go.Bar(
                name=plat,
                x=MONTHS_LABEL,
                y=ys,
                marker_color=PLAT_COLORS[plat],
                text=texts,
                textposition="outside",
            ),
            secondary_y=False,
        )

    # Line – 總營業額（右軸）
    ys_rev = data["總營業額"].tolist()
    fig.add_trace(
        go.Scatter(
            name="總營業額",
            x=MONTHS_LABEL,
            y=ys_rev,
            mode="lines+markers+text",
            line=dict(color="red", width=2),
            marker=dict(size=6),
            text=["$" + f"{v:,}" if v != 0 else "" for v in ys_rev],
            textposition="top center",
        ),
        secondary_y=True,
    )

    # Line – 總淨利額（右軸）
    ys_pft = data["總淨利額"].tolist()
    fig.add_trace(
        go.Scatter(
            name="總淨利額",
            x=MONTHS_LABEL,
            y=ys_pft,
            mode="lines+markers+text",
            line=dict(color="#00B4D8", width=2),
            marker=dict(size=6),
            text=["$" + f"{v:,}" if v != 0 else "" for v in ys_pft],
            textposition="bottom center",
        ),
        secondary_y=True,
    )

    # Horizontal line – 年度平均訂單量（左軸）
    fig.add_trace(
        go.Scatter(
            name=f"年度平均訂單量（{avg:.1f}）",
            x=MONTHS_LABEL,
            y=[avg] * 12,
            mode="lines",
            line=dict(color="#2E8B57", width=2, dash="dot"),
        ),
        secondary_y=False,
    )

    fig.update_layout(
        title=dict(text=f"{year} 年度銷售圖表", font=dict(size=18)),
        barmode="group",
        height=520,
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        yaxis=dict(title="訂單量", rangemode="tozero", tickformat="d"),
        yaxis2=dict(title="金額 ($)", rangemode="tozero", tickformat="$,d"),
    )
    return fig


def _show_metrics(df_year: pd.DataFrame, data: pd.DataFrame):
    total_rev = int(data["總營業額"].sum())
    total_pft = int(data["總淨利額"].sum())
    avg_month = float(data["年度平均訂單量"].iloc[0])

    # 退貨率：退貨訂單 / 全部非取消訂單
    not_cancelled = df_year[~df_year["訂單狀態"].isin(["已取消"])]
    all_orders_n = not_cancelled["訂單編號"].nunique()
    return_orders_n = df_year[
        df_year["訂單狀態"].str.contains("退貨", na=False)
    ]["訂單編號"].nunique()
    return_rate = (return_orders_n / all_orders_n * 100) if all_orders_n > 0 else 0.0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("總營業額", f"${total_rev:,}")
    c2.metric("總淨利額", f"${total_pft:,}")
    c3.metric("年度平均訂單量 / 月", f"{avg_month:.1f}")
    c4.metric("退貨率", f"{return_rate:.1f}%", f"{return_orders_n} / {all_orders_n} 筆")


def _show_table(data: pd.DataFrame):
    table = data.copy()
    table.insert(0, "月份", MONTHS_LABEL)
    table = table.drop(columns=["月", "_total"])
    table = table.rename(columns={p: f"{p}訂單量" for p in PLATFORMS})
    # 末列加總
    total_row: dict[str, object] = {"月份": "合計"}
    for col in table.columns[1:]:
        if col == "年度平均訂單量":
            total_row[col] = None
        else:
            total_row[col] = int(table[col].sum())
    table = pd.concat([table, pd.DataFrame([total_row])], ignore_index=True)
    st.dataframe(table, hide_index=True, width='stretch', height=500)


# ══════════════════════════════════════════════════════════════
# 主渲染
# ══════════════════════════════════════════════════════════════

if len(selected_years) == 1:
    year = selected_years[0]
    df_y = daily[daily["年"] == year].copy()
    data = _build_year_data(df_y)

    _show_metrics(df_y, data)
    st.plotly_chart(_make_chart(data, year), width='stretch')

    st.markdown("### 年度銷售表")
    _show_table(data)

else:
    # 多年份：每年一個 Tab + 最後一個跨年比較 Tab
    tab_labels = [f"{y} 年" for y in selected_years] + ["📊 跨年比較"]
    tabs = st.tabs(tab_labels)

    year_data: dict[int, pd.DataFrame] = {}
    for i, year in enumerate(selected_years):
        df_y = daily[daily["年"] == year].copy()
        data = _build_year_data(df_y)
        year_data[year] = data

        with tabs[i]:
            _show_metrics(df_y, data)
            st.plotly_chart(_make_chart(data, year), width='stretch')
            st.markdown("### 年度銷售表")
            _show_table(data)

    # 跨年比較 Tab
    with tabs[-1]:
        st.markdown("### 📈 跨年 總營業額比較")
        fig_rev = go.Figure()
        for year in selected_years:
            fig_rev.add_trace(go.Scatter(
                name=str(year),
                x=MONTHS_LABEL,
                y=year_data[year]["總營業額"].tolist(),
                mode="lines+markers",
                marker=dict(size=6),
            ))
        fig_rev.update_layout(
            height=380, hovermode="x unified",
            yaxis=dict(title="金額 ($)", tickformat="$,d"),
        )
        st.plotly_chart(fig_rev, width='stretch')

        st.markdown("### 📈 跨年 總淨利額比較")
        fig_pft = go.Figure()
        for year in selected_years:
            fig_pft.add_trace(go.Scatter(
                name=str(year),
                x=MONTHS_LABEL,
                y=year_data[year]["總淨利額"].tolist(),
                mode="lines+markers",
                marker=dict(size=6),
            ))
        fig_pft.update_layout(
            height=380, hovermode="x unified",
            yaxis=dict(title="金額 ($)", tickformat="$,d"),
        )
        st.plotly_chart(fig_pft, width='stretch')

        st.markdown("### 📦 跨年 總訂單量比較")
        fig_ord = go.Figure()
        for year in selected_years:
            ys = year_data[year]["_total"].tolist()
            fig_ord.add_trace(go.Bar(
                name=str(year),
                x=MONTHS_LABEL,
                y=ys,
                text=[str(v) if v > 0 else "" for v in ys],
                textposition="outside",
            ))
        fig_ord.update_layout(
            barmode="group", height=380, hovermode="x unified",
            yaxis=dict(title="訂單量", rangemode="tozero", tickformat="d"),
        )
        st.plotly_chart(fig_ord, width='stretch')

        # 跨年摘要表
        st.markdown("### 年度摘要比較")
        summary_rows = []
        for year in selected_years:
            df_y = daily[daily["年"] == year]
            d = year_data[year]
            not_cancelled = df_y[~df_y["訂單狀態"].isin(["已取消"])]
            all_n = not_cancelled["訂單編號"].nunique()
            ret_n = df_y[df_y["訂單狀態"].str.contains("退貨", na=False)]["訂單編號"].nunique()
            summary_rows.append({
                "年份": year,
                "總營業額": f"${int(d['總營業額'].sum()):,}",
                "總淨利額": f"${int(d['總淨利額'].sum()):,}",
                "年均訂單量/月": f"{float(d['年度平均訂單量'].iloc[0]):.1f}",
                "退貨率": f"{(ret_n / all_n * 100) if all_n > 0 else 0:.1f}% ({ret_n}/{all_n})",
            })
        st.dataframe(pd.DataFrame(summary_rows), hide_index=True, width='stretch', height=500)
