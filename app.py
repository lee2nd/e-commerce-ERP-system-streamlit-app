import io
import zipfile
import psutil
import os
import streamlit as st
from streamlit_autorefresh import st_autorefresh
from datetime import datetime, timezone, timedelta
from utils.styles import apply_global_styles
from utils.data_manager import delete_all_data, restore_from_zip

TZ_TAIPEI = timezone(timedelta(hours=8))

def _get_rw_funcs():
    """Lazy import of read_raw_bytes / save_raw_bytes."""
    try:
        from utils.data_manager import read_raw_bytes, save_raw_bytes
        return read_raw_bytes, save_raw_bytes
    except ImportError:
        return None, None

st.set_page_config(page_title="電商平台進銷存系統", page_icon="📊", layout="wide")


apply_global_styles()

# 每 3 分鐘發送 keep-alive 訊號，防止 Hugging Face Spaces 閒置斷線
st_autorefresh(interval=3 * 60 * 1000, key="keep_alive")

st.title("📊 電商平台 ERP & 報表系統")
st.caption("蝦皮 ｜ 露天 ｜ 官網 (EasyStore) ｜ MO店")

try:
    # ── 記憶體監控 ────────────────────────────────────────────────
    _proc = psutil.Process(os.getpid())
    _mem = _proc.memory_info()
    _mem_mb = _mem.rss / (1024 * 1024)
    _mem_pct = _proc.memory_percent()

    # Hugging Face Spaces 免費方案記憶體上限約 16 GB
    _CLOUD_LIMIT_MB = 16 * 1024
    _cloud_pct = min(100.0, _mem_mb / _CLOUD_LIMIT_MB * 100)

    _mcol1, _mcol2, _mcol3 = st.columns(3)
    _mcol1.metric("本程序 RSS", f"{_mem_mb:.0f} MB")
    _mcol2.metric("佔系統記憶體", f"{_mem_pct:.1f}%")
    _mcol3.metric("佔 HF 上限 (16 GB)", f"{_cloud_pct:.1f}%")

    if _cloud_pct > 80:
        st.warning(f"⚠️ 記憶體使用量已達 HF 上限的 {_cloud_pct:.0f}%，建議清除不需要的資料或重啟應用")
    elif _cloud_pct > 60:
        st.info(f"ℹ️ 記憶體使用量 {_cloud_pct:.0f}%")
except Exception:
    pass  # Silently skip memory monitoring if unavailable

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
    ("自建訂單.xlsx","自建訂單",    "自建訂單（其他平台）累積記錄（原有新增請至【📥 匯入資料】頁面）"),
]

# ── 一鍵下載 ZIP ──────────────────────────────────────────────
st.markdown("#### ⬇️ 一鍵下載全部資料")

if st.button("🔄 產生備份 ZIP", key="gen_zip"):
    # 清除舊的個別下載快取，釋放記憶體
    for _f, *_ in _FILE_META:
        st.session_state.pop(f"_dl_{_f}", None)
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

# ── 一鍵上傳 ZIP 還原 ────────────────────────────────────────
st.markdown("#### ⬆️ 一鍵上傳 ZIP 還原")
st.info("上傳先前下載的備份 ZIP，會將 ZIP 內所有 .xlsx 檔案一次還原（覆蓋現有資料）。")

_uploaded_zip = st.file_uploader(
    "選擇備份 ZIP 檔",
    type=["zip"],
    key="upload_restore_zip",
)
if _uploaded_zip:
    st.warning("⚠️ 確認後將以 ZIP 內的檔案**全面覆蓋**現有資料，操作無法復原！")
    if st.button("✅ 確認上傳還原", key="confirm_restore_zip", type="primary"):
        try:
            zip_data = _uploaded_zip.read()
            restored = restore_from_zip(zip_data)
            if restored:
                # 清除舊的下載快取
                for _f, *_ in _FILE_META:
                    st.session_state.pop(f"_dl_{_f}", None)
                st.session_state.pop("_zip_bytes", None)
                ts = datetime.now(tz=TZ_TAIPEI).strftime("%Y-%m-%d %H:%M:%S")
                st.session_state["_toast_restore"] = f"✅ 已還原 {len(restored)} 個檔案：{', '.join(restored)}（{ts}）"
                st.rerun()
            else:
                st.warning("ZIP 內沒有找到任何 .xlsx 檔案")
        except Exception as e:
            st.error(f"還原失敗：{e}")

_t = st.session_state.pop("_toast_restore", None)
if _t:
    st.success(_t)

st.markdown("---")

# ── 一鍵全刪 ─────────────────────────────────────────────────
st.markdown("#### 🗑️ 一鍵刪除全部資料")
st.error("⚠️ 此操作會**永久刪除**所有資料檔案，無法復原！請務必先下載備份 ZIP。")

_col_del1, _col_del2 = st.columns([1, 3])
with _col_del1:
    _confirm_text = st.text_input(
        "請輸入「確認刪除」以啟用按鈕",
        key="delete_all_confirm_text",
        placeholder="確認刪除",
    )
with _col_del2:
    st.markdown("")  # 佔位對齊
    if st.button(
        "🗑️ 刪除全部資料",
        key="btn_delete_all",
        type="primary",
        disabled=(_confirm_text != "確認刪除"),
    ):
        with st.spinner("刪除中…"):
            deleted = delete_all_data()
        # 清除所有下載快取
        for _f, *_ in _FILE_META:
            st.session_state.pop(f"_dl_{_f}", None)
        st.session_state.pop("_zip_bytes", None)
        if deleted:
            st.session_state["_toast_delete"] = ("success", f"✅ 已刪除 {len(deleted)} 個檔案：{', '.join(deleted)}")
        else:
            st.session_state["_toast_delete"] = ("info", "目前沒有資料檔案需要刪除")
        st.rerun()

_t = st.session_state.pop("_toast_delete", None)
if _t:
    (st.success if _t[0] == "success" else st.info)(_t[1])

st.markdown("---")

# ── 逐一備份 & 全覆蓋上傳 ─────────────────────────────────────
st.markdown("#### 📂 逐一備份 & 全覆蓋上傳")
st.warning("⚠️ 全覆蓋上傳會**直接取代**現有資料，請務必先備份再操作。")

for fname, display_name, note in _FILE_META:
    with st.expander(f"📄 {display_name}（{fname}）"):
        st.caption(note)

        # ── 下載該檔案 ────────────────────────────────────────
        if st.button(f"📥 載入「{display_name}」下載連結", key=f"load_{fname}"):
            # 清除其他檔案的下載快取 & ZIP 快取，避免累積記憶體
            for _f, *_ in _FILE_META:
                if _f != fname:
                    st.session_state.pop(f"_dl_{_f}", None)
            st.session_state.pop("_zip_bytes", None)
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
                    st.session_state[f"_toast_ow_{fname}"] = f"✅ 「{display_name}」已全覆蓋上傳！（{ts}）"
                    st.rerun()
                except Exception as e:
                    st.error(f"上傳失敗：{e}")

        _ow_ts = st.session_state.get(f"_ow_ts_{fname}")
        if _ow_ts:
            st.caption(f"🕐 最後覆蓋：{_ow_ts}")

    _t = st.session_state.pop(f"_toast_ow_{fname}", None)
    if _t:
        st.success(_t)