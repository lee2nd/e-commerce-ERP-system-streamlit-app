"""
資料持久化層 — 透過 GitHub API 讀寫 Excel，統一管理 data/ 目錄。
支援本地開發（直接讀寫檔案）與 Streamlit Cloud（透過 GitHub API）。

Streamlit Cloud Secrets 需設定：
[github]
token = "ghp_xxxxxxxxxxxx"       # GitHub Personal Access Token (repo 權限)
owner = "lee2nd"                  # GitHub 帳號
repo  = "e-commerce-ERP-system-streamlit-app"
branch = "main"                   # 寫入的分支
"""

import os
import io
import base64
import time
import requests
import pandas as pd
from pathlib import Path

# ── 環境判斷 ────────────────────────────────────────────────
_ROOT = Path(__file__).resolve().parent.parent

def _is_cloud() -> bool:
    """判斷是否在 Streamlit Cloud 環境執行。"""
    # Streamlit Cloud 會注入這個環境變數；或 data_dev/ 不存在時也視為雲端
    return bool(os.environ.get("STREAMLIT_CLOUD")) or not (_ROOT / "data_dev").exists()

if _is_cloud():
    DATA_DIR = _ROOT / "data"
else:
    DATA_DIR = _ROOT / "data_dev"
DATA_DIR.mkdir(exist_ok=True)


# ══════════════════════════════════════════════════════════════
# GitHub API 工具函式（僅雲端使用）
# ══════════════════════════════════════════════════════════════

def _gh_config() -> dict:
    """從 Streamlit Secrets 讀取 GitHub 設定。"""
    import streamlit as st
    cfg = st.secrets["github"]
    return {
        "token":  cfg["token"],
        "owner":  cfg["owner"],
        "repo":   cfg["repo"],
        "branch": cfg.get("branch", "main"),
    }


def _gh_headers(token: str) -> dict:
    return {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json",
    }


def _gh_read_excel(filename: str) -> pd.DataFrame:
    """
    從 GitHub Contents API 下載 Excel 檔並回傳 DataFrame。
    使用 Contents API 而非 raw URL，以確保讀取最新版本（避免快取延遲）。
    """
    cfg = _gh_config()
    headers = _gh_headers(cfg["token"])
    api_url = (
        f"https://api.github.com/repos/{cfg['owner']}/{cfg['repo']}"
        f"/contents/data/{filename}"
    )
    resp = requests.get(
        api_url, headers=headers,
        params={"ref": cfg["branch"]},
        timeout=15,
    )
    if resp.status_code == 404:
        return pd.DataFrame()
    resp.raise_for_status()
    data = resp.json()
    # 檔案 < 1 MB：Contents API 直接回傳 base64 content
    if data.get("content"):
        content_bytes = base64.b64decode(data["content"])
    else:
        # 檔案較大時，改用 download_url 下載
        dl_resp = requests.get(data["download_url"], headers=headers, timeout=15)
        dl_resp.raise_for_status()
        content_bytes = dl_resp.content
    return pd.read_excel(io.BytesIO(content_bytes), engine="openpyxl")


def _gh_write_excel(df: pd.DataFrame, filename: str, commit_msg: str):
    """
    將 DataFrame 寫成 Excel 並透過 GitHub API commit 到 data/{filename}。
    若檔案已存在會先取得 sha 再更新；不存在則新建。
    """
    cfg = _gh_config()
    headers = _gh_headers(cfg["token"])
    api_base = f"https://api.github.com/repos/{cfg['owner']}/{cfg['repo']}/contents/data/{filename}"

    # 1. 先查目前的 sha（更新檔案時必填）
    sha = None
    get_resp = requests.get(api_base, headers=headers, params={"ref": cfg["branch"]}, timeout=10)
    if get_resp.status_code == 200:
        sha = get_resp.json().get("sha")
    elif get_resp.status_code != 404:
        get_resp.raise_for_status()

    # 2. DataFrame → Excel bytes → base64
    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl")
    b64_content = base64.b64encode(buf.getvalue()).decode("utf-8")

    # 3. PUT 請求
    if commit_msg is None:
        commit_msg = f"chore: update {filename} via Streamlit app"
    payload = {
        "message": commit_msg,
        "content": b64_content,
        "branch":  cfg["branch"],
    }
    if sha:
        payload["sha"] = sha

    put_resp = requests.put(api_base, headers=headers, json=payload, timeout=20)
    put_resp.raise_for_status()


# ══════════════════════════════════════════════════════════════
# 通用讀寫（CSV，本地用）
# ══════════════════════════════════════════════════════════════

def _load_csv(name: str, **kwargs) -> pd.DataFrame:
    path = DATA_DIR / f"{name}.csv"
    if path.exists() and path.stat().st_size > 0:
        return pd.read_csv(path, encoding="utf-8-sig", **kwargs)
    return pd.DataFrame()


def _save_csv(df: pd.DataFrame, name: str):
    path = DATA_DIR / f"{name}.csv"
    df.to_csv(path, index=False, encoding="utf-8-sig")


# ══════════════════════════════════════════════════════════════
# Excel 讀寫（本地 / 雲端自動切換）
# ══════════════════════════════════════════════════════════════

def _load_excel(filename: str) -> pd.DataFrame:
    """讀取 Excel：本地直接讀檔，雲端透過 GitHub API。"""
    if _is_cloud():
        # 寫入後的快取：確保 rerun 後能立即讀到最新資料（避免 API 延遲）
        import streamlit as st
        cache_key = f"_df_cache_{filename}"
        if cache_key in st.session_state:
            return st.session_state.pop(cache_key)
        try:
            return _gh_read_excel(filename)
        except Exception:
            return pd.DataFrame()
    else:
        path = DATA_DIR / filename
        if path.exists() and path.stat().st_size > 0:
            try:
                return pd.read_excel(path, engine="openpyxl")
            except Exception:
                return pd.DataFrame()
        return pd.DataFrame()


def _save_excel(df: pd.DataFrame, filename: str, commit_msg: str):
    """寫入 Excel：本地直接寫檔，雲端透過 GitHub API commit。"""
    if _is_cloud():
        _gh_write_excel(df, filename, commit_msg)
        # 寫入後暫存到 session_state，讓 rerun 後讀取不受 API 延遲影響
        import streamlit as st
        st.session_state[f"_df_cache_{filename}"] = df.copy()
    else:
        path = DATA_DIR / filename
        df.to_excel(path, index=False, engine="openpyxl")


# ══════════════════════════════════════════════════════════════
# 訂單（CSV）
# ══════════════════════════════════════════════════════════════

def load_orders() -> pd.DataFrame:
    return _load_csv("orders")


def save_orders(df: pd.DataFrame):
    _save_csv(df, "orders")


def append_orders(new_df: pd.DataFrame) -> pd.DataFrame:
    existing = load_orders()
    if existing.empty:
        combined = new_df.copy()
    else:
        combined = pd.concat([existing, new_df], ignore_index=True)
        combined = combined.drop_duplicates(keep="last")
    save_orders(combined)
    return combined


# ══════════════════════════════════════════════════════════════
# 入庫（入庫.xlsx）
# ══════════════════════════════════════════════════════════════

_STORAGE_COL_MAP = {
    "名稱":   "商品名稱",
    "入庫數量": "數量",
    "單價":   "單位成本",
    "金額":   "總金額",
}
_STORAGE_COL_MAP_REV = {v: k for k, v in _STORAGE_COL_MAP.items()}


def load_storage() -> pd.DataFrame:
    df = _load_excel("入庫.xlsx")
    if not df.empty:
        df = df.rename(columns=_STORAGE_COL_MAP)
    return df


def save_storage(df: pd.DataFrame):
    dedup_cols = ["貨號", "規格", "數量", "單位成本", "入庫日期"]
    existing_cols = [c for c in dedup_cols if c in df.columns]
    if existing_cols:
        df = df.drop_duplicates(subset=existing_cols, keep="last").reset_index(drop=True)
    out = df.rename(columns=_STORAGE_COL_MAP_REV)
    _save_excel(out, "入庫.xlsx", "chore: update 入庫.xlsx")


# ══════════════════════════════════════════════════════════════
# 各平台訂單（{platform_name}.xlsx）
# ══════════════════════════════════════════════════════════════

def load_platform_orders(platform_name: str) -> pd.DataFrame:
    return _load_excel(f"{platform_name}.xlsx")


def append_platform_orders(new_df: pd.DataFrame, platform_name: str) -> pd.DataFrame:
    existing = load_platform_orders(platform_name)
    if existing.empty:
        combined = new_df.copy()
    else:
        combined = pd.concat([existing, new_df], ignore_index=True)
    combined = combined.drop_duplicates(keep="last").reset_index(drop=True)
    _save_excel(combined, f"{platform_name}.xlsx", f"chore: update {platform_name}.xlsx")
    return combined


# ══════════════════════════════════════════════════════════════
# 對照表（對照表.xlsx）
# ══════════════════════════════════════════════════════════════

def load_compare_table() -> pd.DataFrame:
    return _load_excel("對照表.xlsx")


def save_compare_table(df: pd.DataFrame):
    _save_excel(df, "對照表.xlsx", "chore: update 對照表.xlsx")


# ══════════════════════════════════════════════════════════════
# 日報表（CSV）
# ══════════════════════════════════════════════════════════════

def load_daily_report() -> pd.DataFrame:
    return _load_csv("daily_report")


def save_daily_report(df: pd.DataFrame):
    _save_csv(df, "daily_report")


# ══════════════════════════════════════════════════════════════
# 出庫（出庫.xlsx）
# ══════════════════════════════════════════════════════════════

def load_delivery() -> pd.DataFrame:
    return _load_excel("出庫.xlsx")


def save_delivery(df: pd.DataFrame):
    _save_excel(df, "出庫.xlsx", "chore: update 出庫.xlsx")


# ══════════════════════════════════════════════════════════════
# 庫存明細（庫存明細.xlsx）
# ══════════════════════════════════════════════════════════════

def load_inventory_details() -> pd.DataFrame:
    return _load_excel("庫存明細.xlsx")


def save_inventory_details(df: pd.DataFrame):
    _save_excel(df, "庫存明細.xlsx", "chore: update 庫存明細.xlsx")


# ══════════════════════════════════════════════════════════════
# 資料清0（保留欄位結構，清除所有資料列）
# ══════════════════════════════════════════════════════════════

def clear_storage():
    """清空入庫資料，保留欄位結構。"""
    cols = ["主貨號", "貨號", "名稱", "規格", "入庫數量", "單價", "金額", "入庫日期"]
    existing = _load_excel("入庫.xlsx")
    if not existing.empty:
        cols = list(existing.columns)
    _save_excel(pd.DataFrame(columns=cols), "入庫.xlsx", "chore: clear 入庫.xlsx")


def clear_platform_orders(platform_name: str):
    """清空指定平台訂單，保留欄位結構。"""
    fname = f"{platform_name}.xlsx"
    existing = _load_excel(fname)
    cols = list(existing.columns) if not existing.empty else []
    _save_excel(pd.DataFrame(columns=cols), fname, f"chore: clear {fname}")


def clear_compare_table():
    """清空對照表，保留欄位結構。"""
    cols = ["平台商品名稱", "平台", "入庫品名", "貨號", "主貨號"]
    existing = _load_excel("對照表.xlsx")
    if not existing.empty:
        cols = list(existing.columns)
    _save_excel(pd.DataFrame(columns=cols), "對照表.xlsx", "chore: clear 對照表.xlsx")


def clear_delivery():
    """清空出庫資料，保留欄位結構。"""
    cols = ["主貨號", "貨號", "名稱", "規格", "出庫數量", "單價", "金額", "出庫日期", "平台"]
    existing = _load_excel("出庫.xlsx")
    if not existing.empty:
        cols = list(existing.columns)
    _save_excel(pd.DataFrame(columns=cols), "出庫.xlsx", "chore: clear 出庫.xlsx")


def clear_inventory_details():
    """清空庫存明細，保留欄位結構。"""
    cols = ["主貨號", "貨號", "名稱", "規格", "進貨數量", "進貨合計",
            "銷售數量", "銷售合計", "現有庫存", "平均成本(庫存明細)", "平均成本(入庫)"]
    existing = _load_excel("庫存明細.xlsx")
    if not existing.empty:
        cols = list(existing.columns)
    _save_excel(pd.DataFrame(columns=cols), "庫存明細.xlsx", "chore: clear 庫存明細.xlsx")