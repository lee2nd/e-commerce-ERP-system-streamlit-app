"""
資料持久化層 — 透過 GitHub API 讀寫 Excel，統一管理 data/ 目錄。
支援本地開發（直接讀寫檔案）與 Hugging Face Spaces（透過 GitHub API）。

Hugging Face Spaces 環境變數需設定：
GITHUB_TOKEN = "ghp_xxxxxxxxxxxx"   # GitHub Personal Access Token (repo 權限)
GITHUB_OWNER = "lee2nd"             # GitHub 帳號
GITHUB_REPO  = "e-commerce-ERP-system-streamlit-app"
GITHUB_BRANCH = "main"              # 寫入的分支（選填，預設 main）
"""

import os
import io
import base64
import requests
import pandas as pd
import streamlit as st
from pathlib import Path

# ── 環境判斷 ────────────────────────────────────────────────
_ROOT = Path(__file__).resolve().parent.parent

def _is_cloud() -> bool:
    """判斷是否在雲端環境（Hugging Face Spaces）執行。"""
    return bool(os.environ.get("SPACE_ID")) or not (_ROOT / "data_dev").exists()

if _is_cloud():
    DATA_DIR = _ROOT / "data"
else:
    DATA_DIR = _ROOT / "data_dev"
DATA_DIR.mkdir(exist_ok=True)


# ══════════════════════════════════════════════════════════════
# GitHub API 工具函式（僅雲端使用）
# ══════════════════════════════════════════════════════════════

def _gh_config() -> dict:
    """從環境變數讀取 GitHub 設定（Hugging Face Spaces）。"""
    try:
        return {
            "token":  os.environ["GITHUB_TOKEN"],
            "owner":  os.environ["GITHUB_OWNER"],
            "repo":   os.environ["GITHUB_REPO"],
            "branch": os.environ.get("GITHUB_BRANCH", "main"),
        }
    except KeyError as e:
        raise RuntimeError(
            f"雲端模式需要設定環境變數 {e}。"
            "若為本地開發，請確認 data_dev/ 資料夾存在。"
        ) from None


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


def _gh_write_raw_bytes(filename: str, file_bytes: bytes, commit_msg: str):
    """將原始 bytes 透過 GitHub API 直接 commit 到 data/{filename}。"""
    cfg = _gh_config()
    headers = _gh_headers(cfg["token"])
    api_base = (
        f"https://api.github.com/repos/{cfg['owner']}/{cfg['repo']}"
        f"/contents/data/{filename}"
    )
    sha = None
    get_resp = requests.get(api_base, headers=headers, params={"ref": cfg["branch"]}, timeout=10)
    if get_resp.status_code == 200:
        sha = get_resp.json().get("sha")
    elif get_resp.status_code != 404:
        get_resp.raise_for_status()
    b64_content = base64.b64encode(file_bytes).decode("utf-8")
    payload: dict = {"message": commit_msg, "content": b64_content, "branch": cfg["branch"]}
    if sha:
        payload["sha"] = sha
    put_resp = requests.put(api_base, headers=headers, json=payload, timeout=20)
    put_resp.raise_for_status()


# ── dtype 優化 ──────────────────────────────────────────────

def _optimize_dtypes(df: pd.DataFrame) -> pd.DataFrame:
    """壓縮 DataFrame 記憶體用量：數值 downcast。"""
    if df.empty:
        return df
    for col in df.select_dtypes(include=["int64"]).columns:
        df[col] = pd.to_numeric(df[col], downcast="integer")
    for col in df.select_dtypes(include=["float64"]).columns:
        df[col] = pd.to_numeric(df[col], downcast="float")
    return df


# ══════════════════════════════════════════════════════════════
# Excel 讀寫（本地 / 雲端自動切換）
# ══════════════════════════════════════════════════════════════

def _load_excel(filename: str) -> pd.DataFrame:
    """讀取 Excel：本地直接讀檔，雲端透過 GitHub API。"""
    if _is_cloud():
        # 寫入後的快取：確保 rerun 後能立即讀到最新資料（避免 API 延遲）
        cache_key = f"_df_cache_{filename}"
        if cache_key in st.session_state:
            return st.session_state.pop(cache_key)
        try:
            return _optimize_dtypes(_gh_read_excel(filename))
        except Exception as e:
            st.warning(f"Failed to load {filename}: {e}")
            return pd.DataFrame()
    else:
        path = DATA_DIR / filename
        if path.exists() and path.stat().st_size > 0:
            try:
                return _optimize_dtypes(pd.read_excel(path, engine="openpyxl"))
            except Exception as e:
                st.warning(f"Failed to load {filename}: {e}")
                return pd.DataFrame()
        return pd.DataFrame()


def _save_excel(df: pd.DataFrame, filename: str, commit_msg: str):
    """寫入 Excel：本地直接寫檔，雲端透過 GitHub API commit。"""
    if _is_cloud():
        _gh_write_excel(df, filename, commit_msg)
        # 寫入後暫存到 session_state，讓 rerun 後讀取不受 API 延遲影響
        st.session_state[f"_df_cache_{filename}"] = df.copy()
    else:
        path = DATA_DIR / filename
        df.to_excel(path, index=False, engine="openpyxl")


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


@st.cache_data(ttl=300)
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
    load_storage.clear()


# ══════════════════════════════════════════════════════════════
# 各平台訂單（{platform_name}.xlsx）
# ══════════════════════════════════════════════════════════════

@st.cache_data(ttl=300)
def load_platform_orders(platform_name: str) -> pd.DataFrame:
    return _load_excel(f"{platform_name}.xlsx")


def append_platform_orders(new_df: pd.DataFrame, platform_name: str) -> pd.DataFrame:
    # ── 檔案內部去重：同列完全相同時，數量欄 sum 合併 ──
    _qty_col = next((c for c in new_df.columns if c in ("數量", "Quantity")), None)
    if _qty_col and not new_df.empty:
        group_cols = [c for c in new_df.columns if c != _qty_col]
        new_df = (
            new_df.groupby(group_cols, dropna=False, sort=False)
            .agg({_qty_col: "sum"})
            .reset_index()[new_df.columns.tolist()]
        )

    existing = load_platform_orders(platform_name)
    if existing.empty:
        combined = new_df.copy()
    else:
        combined = pd.concat([existing, new_df], ignore_index=True)
    combined = combined.drop_duplicates(keep="last").reset_index(drop=True)
    _save_excel(combined, f"{platform_name}.xlsx", f"chore: update {platform_name}.xlsx")
    load_platform_orders.clear()
    return combined


# ══════════════════════════════════════════════════════════════
# 對照表（對照表.xlsx）
# ══════════════════════════════════════════════════════════════

@st.cache_data(ttl=300)
def load_compare_table() -> pd.DataFrame:
    return _load_excel("對照表.xlsx")


def save_compare_table(df: pd.DataFrame):
    _save_excel(df, "對照表.xlsx", "chore: update 對照表.xlsx")
    load_compare_table.clear()


# ══════════════════════════════════════════════════════════════
# 日報表（daily_report.xlsx）
# ══════════════════════════════════════════════════════════════

@st.cache_data(ttl=300)
def load_daily_report() -> pd.DataFrame:
    return _load_excel("日報表.xlsx")


def save_daily_report(df: pd.DataFrame):
    _save_excel(df, "日報表.xlsx", "chore: update 日報表.xlsx")
    load_daily_report.clear()


def clear_daily_report():
    """清空日報表，保留欄位結構。"""
    cols = ["日期", "訂單編號", "訂單狀態", "商品名稱", "貨號",
            "訂單金額", "折扣優惠", "買家支付運費", "平台補助運費",
            "實際運費支出", "物流處理費（運費差額）", "未取貨/退貨運費",
            "成交手續費", "其他服務費", "金流與系統處理費",
            "發票處理費", "其他費用", "商品成本", "總成本", "淨利", "備註", "平台"]
    existing = _load_excel("日報表.xlsx")
    if not existing.empty:
        cols = list(existing.columns)
    _save_excel(pd.DataFrame(columns=cols), "日報表.xlsx", "chore: clear 日報表.xlsx")
    load_daily_report.clear()


# ══════════════════════════════════════════════════════════════
# 月報表（月報表.xlsx）
# ══════════════════════════════════════════════════════════════

@st.cache_data(ttl=300)
def load_monthly_report() -> pd.DataFrame:
    return _load_excel("月報表.xlsx")


def save_monthly_report(df: pd.DataFrame):
    _save_excel(df, "月報表.xlsx", "chore: update 月報表.xlsx")
    load_monthly_report.clear()


def clear_monthly_report():
    """清空月報表，保留欄位結構。"""
    existing = _load_excel("月報表.xlsx")
    cols = list(existing.columns) if not existing.empty else ["年份", "月份"]
    _save_excel(pd.DataFrame(columns=cols), "月報表.xlsx", "chore: clear 月報表.xlsx")
    load_monthly_report.clear()


# ══════════════════════════════════════════════════════════════
# 出庫（出庫.xlsx）
# ══════════════════════════════════════════════════════════════

@st.cache_data(ttl=300)
def load_delivery() -> pd.DataFrame:
    return _load_excel("出庫.xlsx")


def save_delivery(df: pd.DataFrame):
    _save_excel(df, "出庫.xlsx", "chore: update 出庫.xlsx")
    load_delivery.clear()


# ══════════════════════════════════════════════════════════════
# 庫存明細（庫存明細.xlsx）
# ══════════════════════════════════════════════════════════════

@st.cache_data(ttl=300)
def load_inventory_details() -> pd.DataFrame:
    return _load_excel("庫存明細.xlsx")


def save_inventory_details(df: pd.DataFrame):
    _save_excel(df, "庫存明細.xlsx", "chore: update 庫存明細.xlsx")
    load_inventory_details.clear()


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
    load_storage.clear()


def clear_platform_orders(platform_name: str):
    """清空指定平台訂單，保留欄位結構。"""
    fname = f"{platform_name}.xlsx"
    existing = _load_excel(fname)
    cols = list(existing.columns) if not existing.empty else []
    _save_excel(pd.DataFrame(columns=cols), fname, f"chore: clear {fname}")
    load_platform_orders.clear()


def clear_compare_table():
    """清空對照表，保留欄位結構。"""
    cols = ["平台商品名稱", "平台", "入庫品名", "貨號", "主貨號"]
    existing = _load_excel("對照表.xlsx")
    if not existing.empty:
        cols = list(existing.columns)
    _save_excel(pd.DataFrame(columns=cols), "對照表.xlsx", "chore: clear 對照表.xlsx")
    load_compare_table.clear()


def clear_delivery():
    """清空出庫資料，保留欄位結構。"""
    cols = ["主貨號", "貨號", "名稱", "規格", "出庫數量", "單價", "金額", "出庫日期", "平台"]
    existing = _load_excel("出庫.xlsx")
    if not existing.empty:
        cols = list(existing.columns)
    _save_excel(pd.DataFrame(columns=cols), "出庫.xlsx", "chore: clear 出庫.xlsx")
    load_delivery.clear()


def clear_inventory_details():
    """清空庫存明細，保留欄位結構。"""
    cols = ["主貨號", "貨號", "名稱", "規格", "進貨數量", "進貨合計",
            "銷售數量", "銷售合計", "現有庫存", "平均成本"]
    existing = _load_excel("庫存明細.xlsx")
    if not existing.empty:
        cols = list(existing.columns)
    _save_excel(pd.DataFrame(columns=cols), "庫存明細.xlsx", "chore: clear 庫存明細.xlsx")
    load_inventory_details.clear()


# ══════════════════════════════════════════════════════════════
# 組合貨號（組合貨號.xlsx）
# ══════════════════════════════════════════════════════════════

@st.cache_data(ttl=300)
def load_combo_sku() -> pd.DataFrame:
    return _load_excel("組合貨號.xlsx")


def save_combo_sku(df: pd.DataFrame):
    _save_excel(df, "組合貨號.xlsx", "chore: update 組合貨號.xlsx")
    load_combo_sku.clear()


def clear_combo_sku():
    """清空組合貨號，保留欄位結構。"""
    cols = ["組合貨號", "原料貨號", "原料數量"]
    existing = _load_excel("組合貨號.xlsx")
    if not existing.empty:
        cols = list(existing.columns)
    _save_excel(pd.DataFrame(columns=cols), "組合貨號.xlsx", "chore: clear 組合貨號.xlsx")
    load_combo_sku.clear()


# ══════════════════════════════════════════════════════════════
# 自建訂單（自建訂單.xlsx）
# ══════════════════════════════════════════════════════════════

_CUSTOM_ORDER_COLS = [
    "日期", "平台名稱", "訂單編號", "訂單狀態", "買家姓名", "買家帳號",
    "貨號", "數量", "單價", "小計", "折扣優惠", "買家支付運費",
    "實際運費", "未取貨/退貨運費", "其他費用", "費用小記", "訂單總金額",
]


@st.cache_data(ttl=300)
def load_custom_orders() -> pd.DataFrame:
    return _load_excel("自建訂單.xlsx")


def save_custom_orders(df: pd.DataFrame):
    _save_excel(df, "自建訂單.xlsx", "chore: update 自建訂單.xlsx")
    load_custom_orders.clear()


def clear_custom_orders():
    """清空自建訂單，保留欄位結構。"""
    existing = _load_excel("自建訂單.xlsx")
    cols = list(existing.columns) if not existing.empty else _CUSTOM_ORDER_COLS
    _save_excel(pd.DataFrame(columns=cols), "自建訂單.xlsx", "chore: clear 自建訂單.xlsx")
    load_custom_orders.clear()


# ══════════════════════════════════════════════════════════════
# 原始 bytes 讀寫（備份 / 全覆蓋還原用）
# ══════════════════════════════════════════════════════════════

def read_raw_bytes(filename: str) -> bytes | None:
    """讀取 DATA_DIR 內指定檔案的原始 bytes；不存在時回傳 None。"""
    if _is_cloud():
        cfg = _gh_config()
        headers = _gh_headers(cfg["token"])
        api_url = (
            f"https://api.github.com/repos/{cfg['owner']}/{cfg['repo']}"
            f"/contents/data/{filename}"
        )
        resp = requests.get(api_url, headers=headers, params={"ref": cfg["branch"]}, timeout=15)
        if resp.status_code == 404:
            return None
        resp.raise_for_status()
        data = resp.json()
        if data.get("content"):
            return base64.b64decode(data["content"])
        dl_resp = requests.get(data["download_url"], headers=headers, timeout=15)
        dl_resp.raise_for_status()
        return dl_resp.content
    else:
        path = DATA_DIR / filename
        return path.read_bytes() if path.exists() else None


def _clear_file_cache(filename: str):
    """清除指定檔案對應的 st.cache_data 快取。"""
    if filename == "入庫.xlsx":
        load_storage.clear()
    elif filename == "出庫.xlsx":
        load_delivery.clear()
    elif filename == "對照表.xlsx":
        load_compare_table.clear()
    elif filename == "庫存明細.xlsx":
        load_inventory_details.clear()
    elif filename == "月報表.xlsx":
        load_monthly_report.clear()
    elif filename == "日報表.xlsx":
        load_daily_report.clear()
    elif filename == "組合貨號.xlsx":
        load_combo_sku.clear()
    elif filename in ("蝦皮.xlsx", "露天.xlsx", "官網.xlsx"):
        load_platform_orders.clear()


def save_raw_bytes(filename: str, file_bytes: bytes):
    """以原始 bytes 全覆蓋指定檔案，並清除相關快取。"""
    if _is_cloud():
        _gh_write_raw_bytes(filename, file_bytes, f"chore: overwrite {filename} via Streamlit restore")
        # 讓下一次 rerun 讀取到最新資料
        df = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")
        st.session_state[f"_df_cache_{filename}"] = df.copy()
    else:
        (DATA_DIR / filename).write_bytes(file_bytes)
    _clear_file_cache(filename)


def _clear_all_caches():
    """清除所有 st.cache_data 快取。"""
    load_storage.clear()
    load_delivery.clear()
    load_compare_table.clear()
    load_inventory_details.clear()
    load_monthly_report.clear()
    load_daily_report.clear()
    load_combo_sku.clear()
    load_platform_orders.clear()


def _gh_delete_file(filename: str):
    """透過 GitHub API 刪除 data/{filename}。"""
    cfg = _gh_config()
    headers = _gh_headers(cfg["token"])
    api_url = (
        f"https://api.github.com/repos/{cfg['owner']}/{cfg['repo']}"
        f"/contents/data/{filename}"
    )
    get_resp = requests.get(api_url, headers=headers, params={"ref": cfg["branch"]}, timeout=10)
    if get_resp.status_code == 404:
        return  # 檔案不存在，略過
    get_resp.raise_for_status()
    sha = get_resp.json()["sha"]
    payload = {
        "message": f"chore: delete {filename} via Streamlit reset",
        "sha": sha,
        "branch": cfg["branch"],
    }
    del_resp = requests.delete(api_url, headers=headers, json=payload, timeout=20)
    del_resp.raise_for_status()


def delete_all_data():
    """刪除 DATA_DIR 下所有 .xlsx 檔案（本地直接刪檔，雲端透過 GitHub API）。"""
    files = [f for f, *_ in [
        ("入庫.xlsx",), ("出庫.xlsx",), ("對照表.xlsx",), ("庫存明細.xlsx",),
        ("日報表.xlsx",), ("月報表.xlsx",), ("組合貨號.xlsx",),
        ("蝦皮.xlsx",), ("露天.xlsx",), ("官網.xlsx",),
    ]]
    deleted = []
    if _is_cloud():
        for fname in files:
            try:
                _gh_delete_file(fname)
                deleted.append(fname)
            except Exception:
                pass  # 刪除失敗的檔案略過
    else:
        for fname in files:
            path = DATA_DIR / fname
            if path.exists():
                path.unlink()
                deleted.append(fname)
    _clear_all_caches()
    # 同步清除 session_state 中的 df_cache
    for fname in files:
        st.session_state.pop(f"_df_cache_{fname}", None)
    return deleted


def restore_from_zip(zip_bytes: bytes) -> list[str]:
    """從 ZIP 檔還原所有 .xlsx 檔案到 DATA_DIR，回傳已還原的檔名清單。"""
    import zipfile as _zipfile
    restored: list[str] = []
    with _zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as zf:
        for name in zf.namelist():
            # 只處理 .xlsx，忽略子目錄或其他檔案
            if not name.endswith(".xlsx"):
                continue
            # 取得純檔名（忽略 ZIP 內的路徑）
            basename = name.split("/")[-1].split("\\")[-1]
            if not basename:
                continue
            file_bytes = zf.read(name)
            save_raw_bytes(basename, file_bytes)
            restored.append(basename)
    _clear_all_caches()
    return restored