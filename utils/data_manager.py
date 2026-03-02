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

# ── 平台設定 ────────────────────────────────────────────────
DEFAULT_SETTINGS: dict = {
    "蝦皮_成交手續費率": 0.065,
    "蝦皮_金流服務費率": 0.02,
    "蝦皮_免運門檻": 0,
    "蝦皮_運費折抵金額": 0,
    "露天_成交手續費率": 0.02,
    "露天_金流服務費率": 0.01,
    "露天_免運門檻": 0,
    "露天_運費折抵金額": 0,
    "官網_成交手續費率": 0.025,
    "官網_金流服務費率": 0.0,
    "官網_免運門檻": 0,
    "官網_運費折抵金額": 0,
}

def load_settings() -> dict:
    path = DATA_DIR / "settings.json"
    settings = DEFAULT_SETTINGS.copy()
    if path.exists():
        with open(path, "r", encoding="utf-8") as f:
            settings.update(json.load(f))
    return settings

def save_settings(settings: dict):
    path = DATA_DIR / "settings.json"
    with open(path, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)

# ── 清空所有資料 ────────────────────────────────────────────
def clear_all_data():
    for f in DATA_DIR.glob("*.csv"):
        f.unlink()
    json_path = DATA_DIR / "settings.json"
    if json_path.exists():
        json_path.unlink()
