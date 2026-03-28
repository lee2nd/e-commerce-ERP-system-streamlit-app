import io
import zipfile
import streamlit as st
from datetime import datetime, timezone, timedelta

TZ_TAIPEI = timezone(timedelta(hours=8))

def _get_rw_funcs():
    """Lazy import of read_raw_bytes / save_raw_bytes to avoid module-load crash
    when Streamlit Cloud has a cached older data_manager without these symbols."""
    try:
        from utils.data_manager import read_raw_bytes, save_raw_bytes
        return read_raw_bytes, save_raw_bytes
    except ImportError:
        return None, None

st.set_page_config(page_title="電商平台進銷存系統", page_icon="📊", layout="wide")
st.title("📊 電商平台 ERP & 報表系統")
st.caption("蝦皮 ｜ 露天 ｜ 官網 (EasyStore) ｜ MO店")

st.markdown("---")
st.subheader("🗂️ 資料備份與還原")

read_raw_bytes, save_raw_bytes = _get_rw_funcs()
if read_raw_bytes is None or save_raw_bytes is None:
    st.error(
        "⚠️ 備份 / 還原功能尚未就緒（部署版本不符，請稍候自動重新部署或手動重新整理頁面）。"
    )
    st.stop()
assert read_raw_bytes is not None
assert save_raw_bytes is not None

# 各檔案的顯示名稱與備註
_FILE_META: list[tuple[str, str, str]] = [
    ("入庫.xlsx",    "入庫資料",    "原有合併上傳功能請至【📥 匯入資料】頁面"),
    ("出庫.xlsx",    "出庫資料",    "導出出庫的結果"),
    ("對照表.xlsx",  "對照表",      "平台商品 ↔ 入庫品名對照"),
    ("庫存明細.xlsx","庫存明細",    "各商品現有庫存彙總"),
    ("日報表.xlsx",  "日報表",      "每日訂單收益明細"),
    ("月報表.xlsx",  "月報表",      "每月報表彙總"),
    ("組合貨號.xlsx","組合貨號",    "組合商品原料清單"),
    ("蝦皮.xlsx",    "蝦皮訂單",    "蝦皮累積原始訂單（原有合併上傳請至【📥 匯入資料】頁面）"),
    ("露天.xlsx",    "露天訂單",    "露天累積原始訂單（原有合併上傳請至【📥 匯入資料】頁面）"),
    ("官網.xlsx",    "官網訂單",    "官網累積原始訂單（原有合併上傳請至【📥 匯入資料】頁面）"),
]

# ── 一鍵下載 ZIP ──────────────────────────────────────────────
st.markdown("#### ⬇️ 一鍵下載全部資料")

if st.button("🔄 產生備份 ZIP", key="gen_zip"):
    with st.spinner("讀取所有檔案中…"):
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for fname, *_ in _FILE_META:
                raw = read_raw_bytes(fname)
                if raw:
                    zf.writestr(fname, raw)
        buf.seek(0)
        st.session_state["_zip_bytes"] = buf.read()
        st.session_state["_zip_ts"] = datetime.now(tz=TZ_TAIPEI).strftime("%Y%m%d_%H%M%S")

_zip_bytes = st.session_state.get("_zip_bytes")
_zip_ts    = st.session_state.get("_zip_ts", "backup")
if _zip_bytes:
    st.download_button(
        label="⬇️ 下載備份 ZIP",
        data=_zip_bytes,
        file_name=f"erp_backup_{_zip_ts}.zip",
        mime="application/zip",
        key="dl_zip",
    )
    st.caption(f"ZIP 產生時間：{_zip_ts}")

st.markdown("---")

# ── 逐一備份 & 全覆蓋上傳 ─────────────────────────────────────
st.markdown("#### 📂 逐一備份 & 全覆蓋上傳")
st.warning("⚠️ 全覆蓋上傳會**直接取代**現有資料，請務必先備份再操作。")

for fname, display_name, note in _FILE_META:
    with st.expander(f"📄 {display_name}（{fname}）"):
        st.caption(note)

        # ── 下載該檔案 ────────────────────────────────────────
        if st.button(f"📥 載入「{display_name}」下載連結", key=f"load_{fname}"):
            with st.spinner("讀取中…"):
                raw = read_raw_bytes(fname)
            if raw:
                st.session_state[f"_dl_{fname}"] = raw
            else:
                st.warning("目前無資料")

        _dl_data = st.session_state.get(f"_dl_{fname}")
        if _dl_data:
            st.download_button(
                label=f"⬇️ 下載 {display_name}",
                data=_dl_data,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{fname}",
            )

        st.markdown("---")

        # ── 全覆蓋上傳 ────────────────────────────────────────
        uploaded = st.file_uploader(
            f"全覆蓋上傳「{display_name}」",
            type=["xlsx"],
            key=f"up_{fname}",
            help="上傳的 xlsx 會完全取代現有資料（不合併，直接覆蓋）",
        )
        if uploaded:
            st.warning(f"⚠️ 確認後將以上傳檔案完全取代「{display_name}」，操作無法復原！")
            if st.button(f"✅ 確認全覆蓋「{display_name}」", key=f"confirm_{fname}", type="primary"):
                file_bytes = uploaded.read()
                try:
                    save_raw_bytes(fname, file_bytes)
                    ts = datetime.now(tz=TZ_TAIPEI).strftime("%Y-%m-%d %H:%M:%S")
                    st.session_state[f"_ow_ts_{fname}"] = ts
                    # 讓下載快取失效
                    st.session_state.pop(f"_dl_{fname}", None)
                    st.session_state.pop("_zip_bytes", None)
                    st.success(f"✅ 「{display_name}」已全覆蓋上傳！")
                    st.rerun()
                except Exception as e:
                    st.error(f"上傳失敗：{e}")

        _ow_ts = st.session_state.get(f"_ow_ts_{fname}")
        if _ow_ts:
            st.caption(f"🕐 最後覆蓋：{_ow_ts}")