"""
資料持久化層 — 內部以 CSV 格式儲存，對外介面維持 .xlsx 邏輯名稱不變。
支援本地開發（直接讀寫檔案）與 Hugging Face Spaces（透過 R2 S3-compatible API）。

Hugging Face Spaces 環境變數需設定：
R2_ACCESS_KEY_ID     = "..."   # Cloudflare R2 API Token Access Key ID
R2_SECRET_ACCESS_KEY = "..."   # Cloudflare R2 API Token Secret Access Key
"""

import os
import io
import logging
import boto3
import requests
import pandas as pd
import streamlit as st
from botocore.config import Config as BotoConfig
from functools import lru_cache
from pathlib import Path
from boto3.s3.transfer import TransferConfig
import zipfile as _zipfile

_log = logging.getLogger(__name__)

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
# Cloudflare R2 工具函式（僅雲端使用）
# ══════════════════════════════════════════════════════════════

_R2_BUCKET = "lee2nd-erp"
_R2_ENDPOINT = "https://3adce09e7050ac922cce36b5480d0bc7.r2.cloudflarestorage.com"
_R2_PUBLIC_BASE = "https://pub-848c9489895e448793d8f949ea5ce84c.r2.dev"

@lru_cache(maxsize=1)
def _r2_client():
    """建立（或重用）boto3 S3 client，指向 Cloudflare R2。lru_cache 內建 lock，避免多執行緒競態。"""
    return boto3.client(
        "s3",
        endpoint_url=_R2_ENDPOINT,
        aws_access_key_id=os.environ["R2_ACCESS_KEY_ID"],
        aws_secret_access_key=os.environ["R2_SECRET_ACCESS_KEY"],
        region_name="auto",
        config=BotoConfig(
            connect_timeout=30,
            read_timeout=120,
            retries={"max_attempts": 3, "mode": "adaptive"},
        ),
    )


def _r2_read_bytes(filename: str) -> bytes | None:
    """以公開 HTTP GET 從 R2 讀取檔案，回傳 bytes；404 回傳 None。"""
    url = f"{_R2_PUBLIC_BASE}/{filename}"
    resp = requests.get(url, timeout=(10, 120))
    if resp.status_code == 404:
        return None
    resp.raise_for_status()
    return resp.content


# 10 MB multipart threshold — 5MB chunk 對大檔 round trip 過多，改 10MB
_MULTIPART_THRESHOLD = 10 * 1024 * 1024
_MULTIPART_CHUNKSIZE = 10 * 1024 * 1024
_TRANSFER_CFG = TransferConfig(
    multipart_threshold=_MULTIPART_THRESHOLD,
    multipart_chunksize=_MULTIPART_CHUNKSIZE,
)


def _r2_write_bytes(filename: str, file_bytes: bytes):
    """將 bytes 上傳至 R2 bucket；TransferConfig 自動決定是否分片上傳。"""
    size = len(file_bytes)
    _r2_client().upload_fileobj(
        Fileobj=io.BytesIO(file_bytes),
        Bucket=_R2_BUCKET,
        Key=filename,
        Config=_TRANSFER_CFG,
    )
    _log.info("R2 upload OK: %s (%d bytes)", filename, size)


def _r2_delete_file(filename: str):
    """從 R2 bucket 刪除指定檔案（不存在時靜默略過）。"""
    _r2_client().delete_object(Bucket=_R2_BUCKET, Key=filename)


# ── CSV 內部格式工具 ─────────────────────────────────────────

def _csv_name(xlsx_name: str) -> str:
    """將邏輯檔名 (xxx.xlsx) 轉為內部 CSV 檔名。"""
    return xlsx_name.rsplit(".", 1)[0] + ".csv"


# ══════════════════════════════════════════════════════════════
# 資料讀寫（內部 CSV，對外邏輯名仍為 .xlsx）
# ══════════════════════════════════════════════════════════════

def _load_excel(filename: str) -> pd.DataFrame:
    """讀取資料（內部 CSV 格式）。"""
    csv_name = _csv_name(filename)

    if _is_cloud():
        cache_key = f"_df_cache_{filename}"
        if cache_key in st.session_state:
            return st.session_state.pop(cache_key)
        try:
            raw = _r2_read_bytes(csv_name)
            if raw is None:
                return pd.DataFrame()
            return pd.read_csv(io.BytesIO(raw), encoding="utf-8-sig", low_memory=False)
        except Exception as e:
            st.warning(f"Failed to load {filename}: {e}")
            return pd.DataFrame()
    else:
        csv_path = DATA_DIR / csv_name
        if csv_path.exists() and csv_path.stat().st_size > 0:
            try:
                return pd.read_csv(csv_path, encoding="utf-8-sig", low_memory=False)
            except Exception as e:
                st.warning(f"Failed to load {filename}: {e}")
                return pd.DataFrame()
        return pd.DataFrame()


def _save_excel(df: pd.DataFrame, filename: str):
    """寫入資料：以 CSV 格式儲存。"""
    csv_name = _csv_name(filename)
    if _is_cloud():
        try:
            buf = io.BytesIO()
            df.to_csv(buf, index=False, encoding="utf-8-sig")
            csv_bytes = buf.getvalue()
            del buf
        except Exception as e:
            _log.error("CSV serialization failed for %s: %s", csv_name, e)
            st.error(f"⚠️ 資料轉換為 CSV 時失敗：{e}")
            raise
        try:
            _r2_write_bytes(csv_name, csv_bytes)
        except Exception as e:
            _log.error("R2 upload failed for %s: %s", csv_name, e)
            st.error(f"⚠️ 雲端儲存失敗（{csv_name}）：{e}")
            raise
        finally:
            del csv_bytes
        st.session_state[f"_df_cache_{filename}"] = df.copy()
    else:
        csv_path = DATA_DIR / csv_name
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")


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
    _save_excel(out, "入庫.xlsx")
    load_storage.clear() # type: ignore


# ══════════════════════════════════════════════════════════════
# 各平台訂單（{platform_name}.xlsx）
# ══════════════════════════════════════════════════════════════

@st.cache_data(ttl=300)
def load_platform_orders(platform_name: str) -> pd.DataFrame:
    return _load_excel(f"{platform_name}.xlsx")


def append_platform_orders(new_df: pd.DataFrame, platform_name: str) -> pd.DataFrame:
    existing = load_platform_orders(platform_name)
    if existing.empty:
        combined = new_df.drop_duplicates(keep="last").reset_index(drop=True)
    else:
        combined = pd.concat([existing, new_df], ignore_index=True)
        del existing
        combined = combined.drop_duplicates(keep="last").reset_index(drop=True)
    _save_excel(combined, f"{platform_name}.xlsx")
    load_platform_orders.clear()
    return combined


# ══════════════════════════════════════════════════════════════
# 對照表（對照表.xlsx）
# ══════════════════════════════════════════════════════════════

@st.cache_data(ttl=300)
def load_compare_table() -> pd.DataFrame:
    return _load_excel("對照表.xlsx")


def save_compare_table(df: pd.DataFrame):
    _save_excel(df, "對照表.xlsx")
    load_compare_table.clear()


# ══════════════════════════════════════════════════════════════
# 日報表（daily_report.xlsx）
# ══════════════════════════════════════════════════════════════

@st.cache_data(ttl=300)
def load_daily_report() -> pd.DataFrame:
    return _load_excel("日報表.xlsx")


def save_daily_report(df: pd.DataFrame):
    _save_excel(df, "日報表.xlsx")
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
    _save_excel(pd.DataFrame(columns=cols), "日報表.xlsx")
    load_daily_report.clear()


# ══════════════════════════════════════════════════════════════
# 月報表（月報表.xlsx）
# ══════════════════════════════════════════════════════════════

@st.cache_data(ttl=300)
def load_monthly_report() -> pd.DataFrame:
    return _load_excel("月報表.xlsx")


def save_monthly_report(df: pd.DataFrame):
    _save_excel(df, "月報表.xlsx")
    load_monthly_report.clear()


def clear_monthly_report():
    """清空月報表，保留欄位結構。"""
    existing = _load_excel("月報表.xlsx")
    cols = list(existing.columns) if not existing.empty else ["年份", "月份"]
    _save_excel(pd.DataFrame(columns=cols), "月報表.xlsx")
    load_monthly_report.clear()


# ══════════════════════════════════════════════════════════════
# 出庫（出庫.xlsx）
# ══════════════════════════════════════════════════════════════

@st.cache_data(ttl=300)
def load_delivery() -> pd.DataFrame:
    return _load_excel("出庫.xlsx")


def save_delivery(df: pd.DataFrame):
    _save_excel(df, "出庫.xlsx")
    load_delivery.clear()


# ══════════════════════════════════════════════════════════════
# 庫存明細（庫存明細.xlsx）
# ══════════════════════════════════════════════════════════════

@st.cache_data(ttl=300)
def load_inventory_details() -> pd.DataFrame:
    return _load_excel("庫存明細.xlsx")


def save_inventory_details(df: pd.DataFrame):
    _save_excel(df, "庫存明細.xlsx")
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
    _save_excel(pd.DataFrame(columns=cols), "入庫.xlsx")
    load_storage.clear()


def clear_platform_orders(platform_name: str):
    """清空指定平台訂單，保留欄位結構。"""
    fname = f"{platform_name}.xlsx"
    existing = _load_excel(fname)
    cols = list(existing.columns) if not existing.empty else []
    _save_excel(pd.DataFrame(columns=cols), fname)
    load_platform_orders.clear()


def clear_compare_table():
    """清空對照表，保留欄位結構。"""
    cols = ["平台商品名稱", "平台", "入庫品名", "貨號", "主貨號"]
    existing = _load_excel("對照表.xlsx")
    if not existing.empty:
        cols = list(existing.columns)
    _save_excel(pd.DataFrame(columns=cols), "對照表.xlsx")
    load_compare_table.clear()


def clear_delivery():
    """清空出庫資料，保留欄位結構。"""
    cols = ["主貨號", "貨號", "名稱", "規格", "出庫數量", "單價", "金額", "出庫日期", "平台"]
    existing = _load_excel("出庫.xlsx")
    if not existing.empty:
        cols = list(existing.columns)
    _save_excel(pd.DataFrame(columns=cols), "出庫.xlsx")
    load_delivery.clear()


def clear_inventory_details():
    """清空庫存明細，保留欄位結構。"""
    cols = ["主貨號", "貨號", "名稱", "規格", "進貨數量", "進貨合計",
            "銷售數量", "銷售合計", "現有庫存", "平均成本"]
    existing = _load_excel("庫存明細.xlsx")
    if not existing.empty:
        cols = list(existing.columns)
    _save_excel(pd.DataFrame(columns=cols), "庫存明細.xlsx")
    load_inventory_details.clear()


# ══════════════════════════════════════════════════════════════
# 組合貨號（組合貨號.xlsx）
# ══════════════════════════════════════════════════════════════

@st.cache_data(ttl=300)
def load_combo_sku() -> pd.DataFrame:
    return _load_excel("組合貨號.xlsx")


def save_combo_sku(df: pd.DataFrame):
    _save_excel(df, "組合貨號.xlsx")
    load_combo_sku.clear()


def clear_combo_sku():
    """清空組合貨號，保留欄位結構。"""
    cols = ["組合貨號", "原料貨號", "原料數量"]
    existing = _load_excel("組合貨號.xlsx")
    if not existing.empty:
        cols = list(existing.columns)
    _save_excel(pd.DataFrame(columns=cols), "組合貨號.xlsx")
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
    df = _load_excel("自建訂單.xlsx")
    if not df.empty and "日期" in df.columns:
        df["日期"] = pd.to_datetime(df["日期"], errors="coerce").dt.strftime("%Y-%m-%d")
        df["日期"] = df["日期"].fillna("")
    return df


def save_custom_orders(df: pd.DataFrame):
    out = df.copy()
    if "日期" in out.columns:
        out["日期"] = pd.to_datetime(out["日期"], errors="coerce").dt.strftime("%Y-%m-%d")
        out["日期"] = out["日期"].fillna("")
    _save_excel(out, "自建訂單.xlsx")
    load_custom_orders.clear()


def clear_custom_orders():
    """清空自建訂單，保留欄位結構。"""
    existing = _load_excel("自建訂單.xlsx")
    cols = list(existing.columns) if not existing.empty else _CUSTOM_ORDER_COLS
    _save_excel(pd.DataFrame(columns=cols), "自建訂單.xlsx")
    load_custom_orders.clear()


# ══════════════════════════════════════════════════════════════
# 原始 bytes 讀寫（備份 / 全覆蓋還原用）
# ══════════════════════════════════════════════════════════════

def read_raw_csv_bytes(filename: str) -> bytes | None:
    """讀取指定檔案，直接回傳 CSV 原始 bytes（不轉 Excel，速度快）。"""
    csv_name = _csv_name(filename)
    try:
        if _is_cloud():
            return _r2_read_bytes(csv_name)
        else:
            csv_path = DATA_DIR / csv_name
            if not csv_path.exists() or csv_path.stat().st_size == 0:
                return None
            return csv_path.read_bytes()
    except Exception:
        return None


def _clear_file_cache(filename: str):
    """清除指定檔案對應的 st.cache_data 快取（接受任意副檔名）。"""
    base = filename.rsplit(".", 1)[0]
    _MAP = {
        "入庫": load_storage, "出庫": load_delivery, "對照表": load_compare_table,
        "庫存明細": load_inventory_details, "月報表": load_monthly_report,
        "日報表": load_daily_report, "組合貨號": load_combo_sku,
        "蝦皮": load_platform_orders, "露天": load_platform_orders,
        "官網": load_platform_orders, "MO店": load_platform_orders,
        "自建訂單": load_custom_orders,
    }
    fn = _MAP.get(base)
    if fn:
        fn.clear()


def save_raw_bytes(filename: str, file_bytes: bytes, cache_key: str | None = None):
    """以原始 bytes 全覆蓋指定槽位（接收 xlsx/xls/csv → 統一轉存為 CSV），並清除相關快取。"""
    ext = filename.rsplit(".", 1)[-1].lower()
    csv_name = _csv_name(cache_key or filename)
    if ext in ("xlsx", "xls"):
        try:
            df = pd.read_excel(io.BytesIO(file_bytes), engine="calamine")
            buf = io.BytesIO()
            df.to_csv(buf, index=False, encoding="utf-8-sig")
            data = buf.getvalue()
            del buf, df
        except Exception:
            data = file_bytes
            csv_name = filename
        if _is_cloud():
            _r2_write_bytes(csv_name, data)
        else:
            (DATA_DIR / csv_name).write_bytes(data)
    else:
        # csv 或其他格式：直接寫入
        if _is_cloud():
            _r2_write_bytes(csv_name, file_bytes)
        else:
            (DATA_DIR / csv_name).write_bytes(file_bytes)
    _clear_file_cache(cache_key or filename)


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
    load_custom_orders.clear()


def delete_all_data():
    """刪除 DATA_DIR 下所有資料檔（CSV）。"""
    _BASES = ["入庫", "出庫", "對照表", "庫存明細", "日報表", "月報表",
              "組合貨號", "蝦皮", "露天", "官網", "MO店", "自建訂單"]
    deleted = []
    if _is_cloud():
        for base in _BASES:
            try:
                _r2_delete_file(f"{base}.csv")
            except Exception:
                pass
            deleted.append(f"{base}.csv")
    else:
        for base in _BASES:
            path = DATA_DIR / f"{base}.csv"
            if path.exists():
                path.unlink()
            deleted.append(f"{base}.csv")
    _clear_all_caches()
    for base in _BASES:
        st.session_state.pop(f"_df_cache_{base}.xlsx", None)
    return deleted


def restore_from_zip(zip_bytes: bytes) -> list[str]:
    """從 ZIP 檔還原所有 .xlsx / .csv 檔案到 DATA_DIR，回傳已還原的檔名清單。"""
    restored: list[str] = []
    with _zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as zf:
        for name in zf.namelist():
            basename = name.split("/")[-1].split("\\")[-1]
            if not basename:
                continue
            file_bytes = zf.read(name)
            if basename.endswith(".csv"):
                # 直接寫入 CSV，不需轉換
                if _is_cloud():
                    _r2_write_bytes(basename, file_bytes)
                    xlsx_name = basename.rsplit(".", 1)[0] + ".xlsx"
                    st.session_state.pop(f"_df_cache_{xlsx_name}", None)
                else:
                    (DATA_DIR / basename).write_bytes(file_bytes)
                restored.append(basename.rsplit(".", 1)[0] + ".csv")
            elif basename.endswith(".parquet"):
                # 向後相容：還原舊版 parquet 備份，轉換為 CSV
                try:
                    df = pd.read_parquet(io.BytesIO(file_bytes))
                    csv_name = basename.rsplit(".", 1)[0] + ".csv"
                    buf = io.BytesIO()
                    df.to_csv(buf, index=False, encoding="utf-8-sig")
                    csv_bytes = buf.getvalue()
                    if _is_cloud():
                        _r2_write_bytes(csv_name, csv_bytes)
                        xlsx_name = basename.rsplit(".", 1)[0] + ".xlsx"
                        st.session_state.pop(f"_df_cache_{xlsx_name}", None)
                    else:
                        (DATA_DIR / csv_name).write_bytes(csv_bytes)
                    restored.append(basename.rsplit(".", 1)[0] + ".csv")
                except Exception as e:
                    _log.warning("Failed to convert parquet file %s: %s", basename, e)
            elif basename.endswith(".xlsx"):
                save_raw_bytes(basename, file_bytes)
                restored.append(basename.rsplit(".", 1)[0] + ".csv")
    _clear_all_caches()
    return restored