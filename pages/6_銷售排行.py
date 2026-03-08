# """銷售排行"""
# import streamlit as st
# import pandas as pd
# import plotly.express as px

# st.set_page_config(page_title="銷售排行", page_icon="🏆", layout="wide")

# from utils.data_manager import load_delivery

# st.title("🏆 銷售排行")

# delivery = load_delivery()
# if delivery.empty:
#     st.info("請先至「日報表」頁面產生報表（會同時產生出庫資料）")
#     st.stop()

# # ══════════════════════════════════════════════════════════════
# # 商品銷售金額排行
# # ══════════════════════════════════════════════════════════════
# st.markdown("### 💰 商品銷售金額排行")

# top_n = st.slider("顯示前 N 名", 5, 50, 20)

# # 排除 TBD
# valid = delivery[
#     delivery["商品名稱"].notna()
#     & (delivery["商品名稱"] != "")
#     & (~delivery["商品名稱"].str.contains("TBD", case=False, na=False))
# ]

# if valid.empty:
#     st.warning("無有效銷售資料")
#     st.stop()

# sales_rank = (
#     valid.groupby("商品名稱")
#     .agg(銷售金額=("金額", "sum"), 銷售數量=("數量", "sum"))
#     .sort_values("銷售金額", ascending=False)
#     .head(top_n)
#     .reset_index()
# )
# sales_rank["排名"] = range(1, len(sales_rank) + 1)

# fig = px.bar(
#     sales_rank,
#     y="商品名稱", x="銷售金額",
#     orientation="h",
#     text="銷售金額",
#     color="銷售金額",
#     color_continuous_scale="Blues",
# )
# fig.update_layout(yaxis=dict(autorange="reversed"), height=max(400, top_n * 28))
# fig.update_traces(texttemplate="$%{text:,.0f}", textposition="outside")
# st.plotly_chart(fig, width="stretch")

# # 表格
# st.dataframe(
#     sales_rank[["排名", "商品名稱", "銷售金額", "銷售數量"]],
#     width="stretch",
#     hide_index=True,
#     column_config={
#         "銷售金額": st.column_config.NumberColumn(format="$%d"),
#     },
# )

# # ══════════════════════════════════════════════════════════════
# # 規格排行（for selected product）
# # ══════════════════════════════════════════════════════════════
# st.markdown("---")
# st.markdown("### 📦 單品規格銷量排行")

# product_list = sorted(valid["商品名稱"].unique().tolist())
# selected = st.selectbox("選擇商品", product_list)

# if selected:
#     spec_data = (
#         valid[valid["商品名稱"] == selected]
#         .groupby("規格")
#         .agg(銷售數量=("數量", "sum"), 銷售金額=("金額", "sum"))
#         .sort_values("銷售數量", ascending=False)
#         .reset_index()
#     )

#     if not spec_data.empty:
#         fig2 = px.bar(
#             spec_data, x="規格", y="銷售數量",
#             text_auto=True, color="銷售數量",
#             color_continuous_scale="Oranges",
#         )
#         st.plotly_chart(fig2, width="stretch")
#         st.dataframe(spec_data, width="stretch", hide_index=True)
#     else:
#         st.info("此商品無規格資料")
