# """日報表（需求 3：統一合併、需求 7：刪除促銷組合標籤）"""
# import streamlit as st
# import pandas as pd

# st.set_page_config(page_title="日報表", page_icon="📊", layout="wide")

# from utils.data_manager import (
#     load_orders, load_compare_table, load_storage,
#     load_daily_report, save_daily_report, load_settings,
#     load_delivery, save_delivery,
# )
# from utils.calculators import generate_daily_report, generate_delivery

# st.title("📊 日報表")

# # ── 產生報表 ─────────────────────────────────────────────────
# if st.button("🔄 重新產生日報表", type="primary"):
#     orders  = load_orders()
#     compare = load_compare_table()
#     storage = load_storage()
#     settings = load_settings()

#     if orders.empty:
#         st.warning("無訂單資料")
#     else:
#         with st.spinner("計算中…"):
#             daily = generate_daily_report(orders, compare, storage, settings)
#             save_daily_report(daily)

#             delivery = generate_delivery(orders, compare, storage)
#             save_delivery(delivery)

#         st.success(f"✅ 日報表已產生，共 {len(daily)} 筆")
#         st.rerun()

# # ── 顯示報表 ─────────────────────────────────────────────────
# daily = load_daily_report()

# if daily.empty:
#     st.info("尚未產生日報表，請按上方按鈕產生")
#     st.stop()

# daily["日期"] = pd.to_datetime(daily["日期"], errors="coerce")

# # 篩選器
# st.markdown("---")
# col_f1, col_f2, col_f3 = st.columns(3)
# with col_f1:
#     plat_filter = st.multiselect(
#         "平台", options=sorted(daily["平台"].unique()), default=sorted(daily["平台"].unique())
#     )
# with col_f2:
#     status_options = sorted(daily["出貨狀態"].unique())
#     status_filter = st.multiselect("出貨狀態", options=status_options, default=status_options)
# with col_f3:
#     min_date = daily["日期"].min()
#     max_date = daily["日期"].max()
#     if pd.notna(min_date) and pd.notna(max_date):
#         date_range = st.date_input(
#             "日期範圍",
#             value=(min_date.date(), max_date.date()),
#             min_value=min_date.date(),
#             max_value=max_date.date(),
#         )
#     else:
#         date_range = None

# view = daily.copy()
# view = view[view["平台"].isin(plat_filter)]
# view = view[view["出貨狀態"].isin(status_filter)]
# if date_range and len(date_range) == 2:
#     view = view[
#         (view["日期"].dt.date >= date_range[0]) & (view["日期"].dt.date <= date_range[1])
#     ]

# # 摘要
# st.markdown("### 摘要")
# s1, s2, s3, s4 = st.columns(4)
# s1.metric("筆數", f"{len(view):,}")
# s2.metric("營業額", f"${int(view['營業額'].sum()):,}")
# s3.metric("成本", f"${int(view['成本'].sum()):,}")
# s4.metric("淨利", f"${int(view['淨利'].sum()):,}")

# # 表格
# st.markdown("### 明細")
# st.dataframe(
#     view.sort_values("日期", ascending=False),
#     width="stretch",
#     hide_index=True,
#     column_config={
#         "營業額": st.column_config.NumberColumn(format="$%d"),
#         "賣家折扣": st.column_config.NumberColumn(format="$%d"),
#         "運費折抵": st.column_config.NumberColumn(format="$%d"),
#         "成交手續費": st.column_config.NumberColumn(format="$%d"),
#         "金流服務費": st.column_config.NumberColumn(format="$%d"),
#         "成本": st.column_config.NumberColumn(format="$%d"),
#         "淨利": st.column_config.NumberColumn(format="$%d"),
#     },
# )

# # 下載
# csv = view.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
# st.download_button("📥 下載日報表 CSV", csv, "daily_report.csv", "text/csv")
