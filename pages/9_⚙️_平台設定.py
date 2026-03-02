"""平台設定（費率 / 免運門檻）"""
import streamlit as st

st.set_page_config(page_title="平台設定", page_icon="⚙️", layout="wide")

from utils.data_manager import load_settings, save_settings, clear_all_data

st.title("⚙️ 平台設定")

settings = load_settings()

st.markdown("### 費率設定")
st.caption("各平台手續費率與免運門檻，會影響日報表淨利計算")

for plat in ("蝦皮", "露天", "官網"):
    st.markdown(f"#### {plat}")
    cols = st.columns(4)
    settings[f"{plat}_成交手續費率"] = cols[0].number_input(
        f"{plat} 成交手續費率",
        min_value=0.0, max_value=1.0, step=0.001,
        value=float(settings.get(f"{plat}_成交手續費率", 0)),
        format="%.3f",
        key=f"{plat}_fee",
    )
    settings[f"{plat}_金流服務費率"] = cols[1].number_input(
        f"{plat} 金流服務費率",
        min_value=0.0, max_value=1.0, step=0.001,
        value=float(settings.get(f"{plat}_金流服務費率", 0)),
        format="%.3f",
        key=f"{plat}_pay",
    )
    settings[f"{plat}_免運門檻"] = cols[2].number_input(
        f"{plat} 免運門檻 (NT$)",
        min_value=0, step=1,
        value=int(settings.get(f"{plat}_免運門檻", 0)),
        key=f"{plat}_threshold",
    )
    settings[f"{plat}_運費折抵金額"] = cols[3].number_input(
        f"{plat} 運費折抵金額 (NT$)",
        min_value=0, step=1,
        value=int(settings.get(f"{plat}_運費折抵金額", 0)),
        key=f"{plat}_ship",
    )

if st.button("💾 儲存設定", type="primary"):
    save_settings(settings)
    st.success("已儲存！下次產生日報表時會套用新費率。")

# ── 資料管理 ─────────────────────────────────────────────────
st.markdown("---")
st.markdown("### 🗂️ 資料管理")
st.warning("以下操作不可復原")

if st.button("🗑️ 清空所有資料（訂單、入庫、報表…）"):
    clear_all_data()
    st.success("所有資料已清空")
    st.rerun()
