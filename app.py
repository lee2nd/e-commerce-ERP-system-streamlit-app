import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="My Demo App", page_icon="🚀")

st.title("🚀 Hello, Streamlit!")
st.write("這是一個部署在 Streamlit Community Cloud 的 Demo App")

# 互動元件
name = st.text_input("你的名字是？", placeholder="輸入名字...")
if name:
    st.success(f"哈囉，{name}！👋")

# 簡單圖表
st.subheader("📊 隨機折線圖")
chart_data = pd.DataFrame(
    np.random.randn(20, 3),
    columns=["A", "B", "C"]
)
st.line_chart(chart_data)

# Slider
st.subheader("🎚️ 互動 Slider")
x = st.slider("選擇一個數字", 0, 100, 50)
st.write(f"你選了：**{x}**，它的平方是：**{x**2}**")