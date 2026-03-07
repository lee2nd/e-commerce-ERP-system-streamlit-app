"""
資料持久化層 — 讀寫 CSV / JSON，統一管理 data/ 目錄。
"""
import pandas as pd
import json
from pathlib import Path

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
DATA_DIR.mkdir(exist_ok=True)

# ── 通用讀寫 ────────────────────────────────────────────────
def _load(name: str, **kwargs) -> pd.DataFrame:
    path = DATA_DIR / f"{name}.csv"
    if path.exists() and path.stat().st_size > 0:
        return pd.read_csv(path, encoding="utf-8-sig", **kwargs)
    return pd.DataFrame()

def _save(df: pd.DataFrame, name: str):
    path = DATA_DIR / f"{name}.csv"
    df.to_csv(path, index=False, encoding="utf-8-sig")

# ── 訂單 ────────────────────────────────────────────────────
def load_orders() -> pd.DataFrame:
    return _load("orders")

def save_orders(df: pd.DataFrame):
    _save(df, "orders")

def append_orders(new_df: pd.DataFrame) -> pd.DataFrame:
    existing = load_orders()
    if existing.empty:
        combined = new_df.copy()
    else:
        combined = pd.concat([existing, new_df], ignore_index=True)
        combined = combined.drop_duplicates(
            subset=["訂單編號", "平台商品名稱", "數量", "單價"],
            keep="last",
        )
    save_orders(combined)
    return combined

# ── 入庫 ────────────────────────────────────────────────────
def load_storage() -> pd.DataFrame:
    return _load("storage")

def save_storage(df: pd.DataFrame):
    _save(df, "storage")

# ── 對照表 ──────────────────────────────────────────────────
def load_compare_table() -> pd.DataFrame:
    return _load("compare_table")

def save_compare_table(df: pd.DataFrame):
    _save(df, "compare_table")

# ── 日報表 ──────────────────────────────────────────────────
def load_daily_report() -> pd.DataFrame:
    return _load("daily_report")

def save_daily_report(df: pd.DataFrame):
    _save(df, "daily_report")

# ── 出庫 ────────────────────────────────────────────────────
def load_delivery() -> pd.DataFrame:
    return _load("delivery")

def save_delivery(df: pd.DataFrame):
    _save(df, "delivery")