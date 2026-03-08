"""
資料持久化層 — 讀寫 CSV / JSON，統一管理 data/ 目錄。
"""
import pandas as pd
from pathlib import Path

DATA_DIR = Path(__file__).resolve().parent.parent / "data_dev"
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
        combined = combined.drop_duplicates(keep="last")
    save_orders(combined)
    return combined

# ── 入庫（讀寫 入庫.xlsx）────────────────────────────────────
# xlsx 欄位 → 系統欄位的映射
_STORAGE_COL_MAP = {
    "名稱": "商品名稱",
    "入庫數量": "數量",
    "單價": "單位成本",
    "金額": "總金額",
}
# 反向映射（系統 → xlsx）
_STORAGE_COL_MAP_REV = {v: k for k, v in _STORAGE_COL_MAP.items()}

def load_storage() -> pd.DataFrame:
    """從 data/入庫.xlsx 讀取入庫資料，欄位自動映射為系統格式。"""
    path = DATA_DIR / "入庫.xlsx"
    if path.exists() and path.stat().st_size > 0:
        try:
            df = pd.read_excel(path, engine="openpyxl")
            df = df.rename(columns=_STORAGE_COL_MAP)
            return df
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()

def save_storage(df: pd.DataFrame):
    """將入庫資料寫回 data/入庫.xlsx，自動去重後欄位映射回 xlsx 格式。"""
    path = DATA_DIR / "入庫.xlsx"
    # 以「貨號 + 規格 + 數量 + 單位成本 + 入庫日期」去重，保留最後一筆
    dedup_cols = ["貨號", "規格", "數量", "單位成本", "入庫日期"]
    existing_cols = [c for c in dedup_cols if c in df.columns]
    if existing_cols:
        df = df.drop_duplicates(subset=existing_cols, keep="last").reset_index(drop=True)
    out = df.rename(columns=_STORAGE_COL_MAP_REV)
    out.to_excel(path, index=False, engine="openpyxl")

# ── 各平台訂單 xlsx ─────────────────────────────────────────
def load_platform_orders(platform_name: str) -> pd.DataFrame:
    """從 data/{platform_name}.xlsx 讀取該平台累積訂單。"""
    path = DATA_DIR / f"{platform_name}.xlsx"
    if path.exists() and path.stat().st_size > 0:
        try:
            return pd.read_excel(path, engine="openpyxl")
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()

def append_platform_orders(new_df: pd.DataFrame, platform_name: str) -> pd.DataFrame:
    """將新訂單追加至 data/{platform_name}.xlsx，全欄位去重。"""
    existing = load_platform_orders(platform_name)
    if existing.empty:
        combined = new_df.copy()
    else:
        combined = pd.concat([existing, new_df], ignore_index=True)
    combined = combined.drop_duplicates(keep="last")
    combined = combined.reset_index(drop=True)
    path = DATA_DIR / f"{platform_name}.xlsx"
    combined.to_excel(path, index=False, engine="openpyxl")
    return combined

# ── 對照表（讀寫 對照表.xlsx）──────────────────────────────────────────
def load_compare_table() -> pd.DataFrame:
    path = DATA_DIR / "對照表.xlsx"
    if path.exists() and path.stat().st_size > 0:
        try:
            return pd.read_excel(path, engine="openpyxl")
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()

def save_compare_table(df: pd.DataFrame):
    path = DATA_DIR / "對照表.xlsx"
    df.to_excel(path, index=False, engine="openpyxl")

# ── 日報表 ──────────────────────────────────────────────────
def load_daily_report() -> pd.DataFrame:
    return _load("daily_report")

def save_daily_report(df: pd.DataFrame):
    _save(df, "daily_report")

# ── 出庫（讀寫 出庫.xlsx）────────────────────────────────────
def load_delivery() -> pd.DataFrame:
    """從 data/出庫.xlsx 讀取出庫資料。"""
    path = DATA_DIR / "出庫.xlsx"
    if path.exists() and path.stat().st_size > 0:
        try:
            return pd.read_excel(path, engine="openpyxl")
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()

def save_delivery(df: pd.DataFrame):
    """將出庫資料寫入 data/出庫.xlsx。"""
    path = DATA_DIR / "出庫.xlsx"
    df.to_excel(path, index=False, engine="openpyxl")