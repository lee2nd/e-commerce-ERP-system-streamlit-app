"""全域 UI 樣式注入工具。
在每個頁面的 st.set_page_config() 之後呼叫 apply_global_styles() 即可套用。
"""
import streamlit as st


def apply_global_styles() -> None:
    st.markdown("""
<style>
/* ── 基礎排版 ─────────────────────────────────────────────── */
html, body {
    font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont,
                 'PingFang TC', 'Microsoft JhengHei', sans-serif !important;
}

.main .block-container {
    padding-top: 1.5rem !important;
    padding-bottom: 3rem !important;
}

/* ── 標題 ─────────────────────────────────────────────────── */
h1 {
    font-size: 1.75rem !important;
    font-weight: 700 !important;
    padding-bottom: 0.6rem !important;
    border-bottom: 3px solid #2563EB !important;
    margin-bottom: 1.25rem !important;
}
h2 {
    font-size: 1.3rem !important;
    font-weight: 600 !important;
}
h3 {
    font-size: 1.05rem !important;
    font-weight: 600 !important;
    color: #334155 !important;
}

/* ── 指標卡片 (Metric) ──────────────────────────────────────── */
[data-testid="metric-container"] {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 1rem 1.25rem !important;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    transition: box-shadow 0.2s ease;
}
[data-testid="metric-container"]:hover {
    box-shadow: 0 4px 12px rgba(0,0,0,0.09);
}
[data-testid="stMetricLabel"] > div {
    font-size: 0.78rem !important;
    font-weight: 600 !important;
    color: #64748b !important;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}
[data-testid="stMetricValue"] > div {
    font-size: 1.55rem !important;
    font-weight: 700 !important;
}

/* ── 按鈕 ─────────────────────────────────────────────────── */
.stButton > button, [data-testid="stDownloadButton"] > button {
    border-radius: 8px !important;
    font-weight: 500 !important;
    transition: all 0.18s ease !important;
    letter-spacing: 0.01em;
}
.stButton > button:hover, [data-testid="stDownloadButton"] > button:hover {
    transform: translateY(-1px);
    box-shadow: 0 4px 14px rgba(37,99,235,0.22) !important;
}
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #2563EB 0%, #1d4ed8 100%) !important;
    border: none !important;
    box-shadow: 0 2px 6px rgba(37,99,235,0.30) !important;
}

/* ── Tabs ────────────────────────────────────────────────── */
.stTabs [data-baseweb="tab-list"] {
    gap: 4px;
    background: #f1f5f9;
    border-radius: 12px;
    padding: 4px 6px;
    border: none !important;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 8px !important;
    padding: 8px 20px !important;
    font-weight: 500 !important;
    color: #64748b !important;
    background: transparent !important;
    border: none !important;
    transition: all 0.15s ease !important;
}
.stTabs [data-baseweb="tab"]:hover {
    color: #2563EB !important;
    background: rgba(255,255,255,0.75) !important;
}
.stTabs [aria-selected="true"] {
    background: #ffffff !important;
    color: #2563EB !important;
    box-shadow: 0 1px 5px rgba(0,0,0,0.10) !important;
    font-weight: 600 !important;
}
.stTabs [data-baseweb="tab-highlight"] {
    display: none !important;
}
.stTabs [data-baseweb="tab-border"] {
    display: none !important;
}

/* ── Expander ────────────────────────────────────────────── */
[data-testid="stExpander"] {
    border: 1px solid #e2e8f0 !important;
    border-radius: 10px !important;
    overflow: hidden !important;
    margin-bottom: 6px;
    transition: box-shadow 0.2s ease;
}
[data-testid="stExpander"]:hover {
    box-shadow: 0 2px 8px rgba(0,0,0,0.07);
}
[data-testid="stExpander"] summary {
    font-weight: 500 !important;
    background: #f8fafc !important;
}

/* ── Alert / Info / Warning / Error ─────────────────────── */
[data-testid="stAlert"] {
    border-radius: 10px !important;
    border-left-width: 4px !important;
}

/* ── DataFrame ──────────────────────────────────────────── */
[data-testid="stDataFrame"] {
    border-radius: 10px !important;
    overflow: hidden !important;
    border: 1px solid #e2e8f0 !important;
}

/* ── Select / Input 圓角 ─────────────────────────────────── */
[data-testid="stSelectbox"] > div > div > div,
[data-testid="stMultiSelect"] > div > div,
.stTextInput > div > div > input,
.stNumberInput > div > div > input {
    border-radius: 8px !important;
}

/* ── 側邊欄 ─────────────────────────────────────────────── */
[data-testid="stSidebar"] {
    border-right: 1px solid #e2e8f0 !important;
}

/* ── 分隔線 ─────────────────────────────────────────────── */
hr {
    border: none !important;
    border-top: 1px solid #e2e8f0 !important;
    margin: 1.25rem 0 !important;
}

/* ── Caption 小字 ────────────────────────────────────────── */
[data-testid="stCaptionContainer"] p {
    color: #94a3b8 !important;
    font-size: 0.78rem !important;
}

/* ── 捲軸美化 ────────────────────────────────────────────── */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: #f1f5f9; border-radius: 4px; }
::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
</style>
""", unsafe_allow_html=True)
