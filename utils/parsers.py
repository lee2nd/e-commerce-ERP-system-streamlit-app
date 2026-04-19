"""
解析各電商平台匯出的訂單檔案，輸出統一欄位格式的 DataFrame。

統一欄位：
  訂單編號, 日期, 平台, 平台商品名稱, 貨號, 數量, 單價, 金額, 賣家折扣, 訂單狀態
"""
import io
import pandas as pd


# ── 通用檔案讀取（支援多種格式）─────────────────────────────
def read_file_flexible(file) -> pd.DataFrame:
    """嘗試用多種引擎讀取上傳的 Excel / CSV 檔案。"""
    content: bytes = file.read()
    file.seek(0)
    name = getattr(file, "name", "").lower()

    # CSV
    if name.endswith(".csv"):
        for enc in ("utf-8-sig", "utf-8", "big5", "cp950", "gbk"):
            try:
                return pd.read_csv(io.BytesIO(content), encoding=enc)
            except Exception:
                continue

    # Excel — 依序嘗試各引擎（calamine 最快，優先使用）
    for engine in ("calamine", "openpyxl", "xlrd"):
        try:
            return pd.read_excel(io.BytesIO(content), engine=engine)
        except Exception:
            continue

    # 最後嘗試當 CSV 讀
    for enc in ("utf-8-sig", "utf-8", "big5", "cp950", "gbk", "latin1"):
        try:
            return pd.read_csv(io.BytesIO(content), encoding=enc)
        except Exception:
            continue

    raise ValueError(
        "無法讀取此檔案，請嘗試另存為 CSV（UTF-8）或 XLSX 格式後重新上傳。"
    )


# ── 欄位比對輔助 ────────────────────────────────────────────
def _find_col(df: pd.DataFrame, patterns: list[str]) -> str | None:
    """用 pattern 清單（大小寫不敏感、部分比對）找出最匹配的欄位名。"""
    for pat in patterns:
        for col in df.columns:
            if pat.lower() in str(col).lower():
                return col
    return None


# ── 蝦皮 ────────────────────────────────────────────────────
def parse_shopee(file_or_df) -> pd.DataFrame:
    df = file_or_df if isinstance(file_or_df, pd.DataFrame) else read_file_flexible(file_or_df)

    # 自動偵測欄位（支援不同年版匯出格式）
    c_oid   = _find_col(df, ["訂單編號", "Order ID", "order_sn"]) or df.columns[0]
    c_stat  = _find_col(df, ["訂單狀態", "Order Status"]) or df.columns[1]
    c_ret   = _find_col(df, ["退貨", "退款", "Return", "Refund"])
    c_date  = _find_col(df, ["建立日期", "訂單建立日期", "出貨日期", "Create Date"]) or df.columns[5]
    c_prod  = _find_col(df, ["商品名稱", "Product Name"]) or (df.columns[21] if len(df.columns) > 21 else df.columns[-5])
    c_var   = _find_col(df, ["選項名稱", "商品選項", "Variation"]) or (df.columns[22] if len(df.columns) > 22 else None)
    c_orig  = _find_col(df, ["原始價格", "Original Price"]) or (df.columns[23] if len(df.columns) > 23 else None)
    c_deal  = _find_col(df, ["成交價格", "Deal Price"])
    c_qty   = _find_col(df, ["數量", "Quantity"]) or (df.columns[27] if len(df.columns) > 27 else df.columns[-2])
    c_disc  = _find_col(df, ["賣家折扣", "Seller Discount", "賣家優惠券", "賣家折扣活動"])
    c_sku   = _find_col(df, ["商品選項貨號", "商品SKU", "SKU Reference", "商品貨號", "sku_reference"])

    out = pd.DataFrame()
    out["訂單編號"] = df[c_oid].astype(str).str.strip()
    out["日期"]     = pd.to_datetime(df[c_date].astype(str).str[:10], errors="coerce")
    out["平台"]     = "蝦皮"

    prod = df[c_prod].fillna("").astype(str)
    var  = df[c_var].fillna("").astype(str) if c_var else ""
    out["平台商品名稱"] = prod + "::" + var

    out["貨號"] = df[c_sku].astype(str).str.strip() if c_sku else ""
    out["數量"] = pd.to_numeric(df[c_qty], errors="coerce").fillna(0).astype(int)

    if c_deal:
        deal = pd.to_numeric(df[c_deal], errors="coerce")
        orig = pd.to_numeric(df[c_orig], errors="coerce") if c_orig else 0
        out["單價"] = deal.fillna(orig).fillna(0)
    elif c_orig:
        out["單價"] = pd.to_numeric(df[c_orig], errors="coerce").fillna(0)
    else:
        out["單價"] = 0

    out["金額"] = out["數量"] * out["單價"]
    out["賣家折扣"] = pd.to_numeric(df[c_disc], errors="coerce").fillna(0).abs() if c_disc else 0

    # 訂單狀態
    out["訂單狀態"] = "正常"
    stat = df[c_stat].fillna("").astype(str)
    out.loc[stat.str.contains("取消|cancel", case=False), "訂單狀態"] = "已取消"
    out.loc[stat.str.contains("退貨|退款|return|refund", case=False), "訂單狀態"] = "退貨"
    if c_ret:
        ret = df[c_ret].fillna("").astype(str)
        out.loc[ret.str.contains("退貨|退款|return|refund", case=False), "訂單狀態"] = "退貨"

    out = out[out["訂單編號"].notna() & ~out["訂單編號"].isin(["", "nan"])]
    return out.reset_index(drop=True)


# ── 露天 ────────────────────────────────────────────────────
def parse_ruten(file_or_df) -> pd.DataFrame:
    df = file_or_df if isinstance(file_or_df, pd.DataFrame) else read_file_flexible(file_or_df)

    out = pd.DataFrame()
    out["訂單編號"] = df["訂單編號"].astype(str).str.strip()
    out["日期"]     = pd.to_datetime(df["結帳時間"].astype(str).str[:10], errors="coerce")
    out["平台"]     = "露天"

    name = df["商品名稱"].fillna("").astype(str)
    spec = df["規格"].fillna("").astype(str)
    item = df["項目"].fillna("").astype(str)
    out["平台商品名稱"] = name + "::" + spec + "::" + item

    if "賣家自用料號" in df.columns:
        _sku = df["賣家自用料號"].fillna("").astype(str).str.strip()
        out["貨號"] = _sku.where(~_sku.str.lower().isin(["nan", "none", "nat", "<na>"]), "")
    else:
        out["貨號"] = ""
    out["數量"] = pd.to_numeric(df["數量"], errors="coerce").fillna(0).astype(int)
    out["單價"] = pd.to_numeric(df["單價"], errors="coerce").fillna(0)
    out["金額"] = out["數量"] * out["單價"]

    s_disc = pd.to_numeric(df.get("賣家折扣碼金額", 0), errors="coerce").fillna(0).abs() # type: ignore
    p_disc = pd.to_numeric(df.get("露天折扣碼金額", 0), errors="coerce").fillna(0).abs() # type: ignore
    out["賣家折扣"] = s_disc + p_disc

    out["訂單狀態"] = "正常"
    if "交易狀況" in df.columns:
        tx = df["交易狀況"].fillna("").astype(str)
        out.loc[tx.str.contains("已領退貨|退貨", case=False), "訂單狀態"] = "退貨"
    if "訂單狀態" in df.columns:
        st_ = df["訂單狀態"].fillna("").astype(str)
        out.loc[st_.str.contains("取消", case=False), "訂單狀態"] = "已取消"

    out = out[out["訂單編號"].notna() & ~out["訂單編號"].isin(["", "nan"])]
    return out.reset_index(drop=True)


# ── 官網 EasyStore ──────────────────────────────────────────
def parse_easystore(file_or_df) -> pd.DataFrame:
    df = file_or_df if isinstance(file_or_df, pd.DataFrame) else read_file_flexible(file_or_df)

    # EasyStore 同一訂單多品項時，訂單層欄位只在首列有值 → forward-fill
    order_cols = [
        "Order Name", "Order Number", "Date", "Channel",
        "Subtotal", "Shipping Fee", "Order Discount",
        "Total Amount", "Financial Status", "Fulfillment Status",
        "Refunded Amount",
    ]
    for c in order_cols:
        if c in df.columns:
            df[c] = df[c].ffill()

    # 過濾空品項列
    if "Item Name" in df.columns:
        df = df[df["Item Name"].notna() & (df["Item Name"] != "")]

    out = pd.DataFrame()
    out["訂單編號"] = df["Order Name"].astype(str).str.strip()
    out["日期"]     = pd.to_datetime(df["Date"].astype(str).str[:10], errors="coerce")
    out["平台"]     = "官網"

    item_name = df["Item Name"].fillna("").astype(str)
    item_var  = df["Item Variant"].fillna("").astype(str) if "Item Variant" in df.columns else ""
    out["平台商品名稱"] = item_name + "::" + item_var

    out["貨號"] = df["Item SKU"].astype(str).str.strip() if "Item SKU" in df.columns else ""
    out["數量"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0).astype(int)
    out["單價"] = pd.to_numeric(df["Item Price"], errors="coerce").fillna(0)
    out["金額"] = out["數量"] * out["單價"]
    out["賣家折扣"] = pd.to_numeric(df.get("Order Discount", 0), errors="coerce").fillna(0).abs() # type: ignore

    out["訂單狀態"] = "正常"
    if "Financial Status" in df.columns:
        fs = df["Financial Status"].fillna("").astype(str)
        out.loc[fs.str.contains("Refund|退款", case=False, na=False), "訂單狀態"] = "退貨"
    if "Fulfillment Status" in df.columns:
        ffs = df["Fulfillment Status"].fillna("").astype(str)
        out.loc[ffs.str.contains("Cancel|取消", case=False, na=False), "訂單狀態"] = "已取消"

    out = out[out["訂單編號"].notna() & ~out["訂單編號"].isin(["", "nan"])]
    return out.reset_index(drop=True)


# ── MO店 ─────────────────────────────────────────────────────
def parse_mo(file_or_df) -> pd.DataFrame:
    df = file_or_df if isinstance(file_or_df, pd.DataFrame) else read_file_flexible(file_or_df)

    out = pd.DataFrame()
    out["訂單編號"] = df["訂單編號"].astype(str).str.strip()
    out["日期"]     = pd.to_datetime(df["轉單日"].astype(str).str[:10], errors="coerce")
    out["平台"]     = "MO店"

    name = df["商品名稱"].fillna("").astype(str)
    out["平台商品名稱"] = name

    out["貨號"] = df["商品原廠編號"].astype(str).str.strip() if "商品原廠編號" in df.columns else ""
    out["數量"] = pd.to_numeric(df["數量"], errors="coerce").fillna(0).astype(int)
    out["單價"] = pd.to_numeric(df["商品售價"], errors="coerce").fillna(0)
    out["金額"] = out["數量"] * out["單價"]
    out["賣家折扣"] = 0

    # 訂單狀態
    order_stat = df["訂單狀態"].fillna("").astype(str)
    ret_reason = df["銷退原因"].fillna("").astype(str) if "銷退原因" in df.columns else pd.Series("", index=df.index)

    out["訂單狀態"] = "正常"
    out.loc[order_stat == "配送結束", "訂單狀態"] = "已完成"
    out.loc[order_stat == "已回收", "訂單狀態"] = "退貨"
    out.loc[order_stat == "配送異常", "訂單狀態"] = "未取貨"
    out.loc[ret_reason == "配送異常結案", "訂單狀態"] = "未取貨"
    out.loc[order_stat == "取消訂單", "訂單狀態"] = "已取消"

    out = out[out["訂單編號"].notna() & ~out["訂單編號"].isin(["", "nan"])]
    return out.reset_index(drop=True)
