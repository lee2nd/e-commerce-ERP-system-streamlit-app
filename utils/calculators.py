"""
業務邏輯計算：對照表匹配、日報表 / 月報表 / 出庫 / 庫存明細產生。
"""
import pandas as pd


# ══════════════════════════════════════════════════════════════
# 對照表自動匹配（需求 5：依照貨號自動匹配商品）
# ══════════════════════════════════════════════════════════════
def auto_match_compare_table(
    orders_df: pd.DataFrame,
    storage_df: pd.DataFrame | None = None,
    existing_compare_df: pd.DataFrame | None = None,
) -> pd.DataFrame:
    if orders_df.empty:
        return existing_compare_df if existing_compare_df is not None else pd.DataFrame()

    result_cols = ["平台商品名稱", "平台", "入庫品名", "貨號", "主貨號"]

    # ── 從訂單取 貨號：同一個(平台商品名稱, 平台) 取第一個非空 SKU ──
    sku_src = (
        orders_df[["平台商品名稱", "平台", "貨號"]].copy()
        .assign(貨號=lambda d: d["貨號"].astype(str).str.strip())
    )
    sku_src = (
        sku_src[sku_src["貨號"].str.len() > 0]
        .replace("nan", "")
        .query('貨號 != ""')
        .drop_duplicates(subset=["平台商品名稱", "平台"])
        .rename(columns={"貨號": "_sku"})
    )

    # ── 從入庫建立 貨號 → 入庫品名 的映射 ──
    stg_name_map: dict[str, str] = {}
    if storage_df is not None and not storage_df.empty and "貨號" in storage_df.columns:
        nm = storage_df[["貨號", "商品名稱", "規格"]].drop_duplicates("貨號").copy()
        nm["貨號"] = nm["貨號"].astype(str).str.strip()
        nm["規格"] = nm["規格"].fillna("").astype(str)
        nm["商品名稱"] = nm["商品名稱"].fillna("").astype(str)
        nm["_label"] = nm.apply(
            lambda r: f"{r['商品名稱']}[{r['規格']}]" if r["規格"] else r["商品名稱"], axis=1
        )
        stg_name_map = nm.set_index("貨號")["_label"].to_dict()

    def _enrich(df: pd.DataFrame) -> pd.DataFrame:        
        df = df.merge(sku_src, on=["平台商品名稱", "平台"], how="left")
        df["貨號"] = df["_sku"].fillna("").astype(str)
        df.drop(columns=["_sku"], errors="ignore", inplace=True)
        df["主貨號"] = df["貨號"].apply(lambda s: s.split("-")[0] if s else "")
        df["入庫品名"] = df["貨號"].map(stg_name_map).fillna("")
        return df

    all_prods = (
        orders_df[["平台商品名稱", "平台"]]
        .drop_duplicates()
        .copy()
    )

    if existing_compare_df is not None and not existing_compare_df.empty:
        for c in result_cols:
            if c not in existing_compare_df.columns:
                existing_compare_df[c] = ""
        # 更新現有列的 貨號/主貨號/入庫品名
        existing = _enrich(existing_compare_df[["平台商品名稱", "平台"]].copy())
        # 新增尚未存在的項目
        known = set(
            existing_compare_df["平台商品名稱"].astype(str)
            + "||" + existing_compare_df["平台"].astype(str)
        )
        new_only = all_prods[
            ~(all_prods["平台商品名稱"].astype(str)
              + "||" + all_prods["平台"].astype(str)).isin(known)
        ].copy()
        new_only = _enrich(new_only)
        result = pd.concat([existing[result_cols], new_only[result_cols]], ignore_index=True)
    else:
        result = _enrich(all_prods)[result_cols].copy()

    return result.drop_duplicates(subset=["平台商品名稱", "平台"]).reset_index(drop=True)


# ══════════════════════════════════════════════════════════════
# 出庫紀錄
# ══════════════════════════════════════════════════════════════
def generate_delivery(
    orders_df: pd.DataFrame,
    compare_df: pd.DataFrame,
    storage_df: pd.DataFrame,
) -> pd.DataFrame:
    if orders_df.empty:
        return pd.DataFrame()

    df = orders_df[~orders_df["訂單狀態"].isin(["已取消"])].copy()

    # 合併對照表取得貨號 / 主貨號
    if not compare_df.empty and "平台商品名稱" in compare_df.columns:
        m = compare_df[["平台商品名稱", "貨號", "主貨號"]].drop_duplicates("平台商品名稱")
        df = df.merge(m, on="平台商品名稱", how="left", suffixes=("_ord", ""))
        if "貨號_ord" in df.columns:
            df["貨號"] = df["貨號"].where(
                df["貨號"].notna() & (df["貨號"] != ""), df["貨號_ord"]
            )
            df.drop(columns=["貨號_ord"], errors="ignore", inplace=True)

    # 合併入庫取得商品名稱 / 規格
    if not storage_df.empty and "貨號" in storage_df.columns:
        si = storage_df[["貨號", "主貨號", "商品名稱", "規格"]].drop_duplicates("貨號")
        df = df.merge(si, on="貨號", how="left", suffixes=("", "_stg"))
        for c in ("主貨號", "商品名稱"):
            sc = f"{c}_stg"
            if sc in df.columns:
                df[c] = df[sc].where(df[sc].notna() & (df[sc] != ""), df.get(c, ""))
                df.drop(columns=[sc], errors="ignore", inplace=True)

    cols_map = {
        "主貨號": df.get("主貨號", ""),
        "商品名稱": df.get("商品名稱", df["平台商品名稱"]),
        "規格": df.get("規格", ""),
        "貨號": df.get("貨號", ""),
        "數量": df["數量"],
        "單價": df["單價"],
        "金額": df["金額"],
        "日期": df["日期"],
        "平台": df["平台"],
    }
    delivery = pd.DataFrame(cols_map)
    delivery = delivery.drop_duplicates().sort_values("日期")
    return delivery.reset_index(drop=True)


# ══════════════════════════════════════════════════════════════
# 日報表（需求 3 合併 A/B、需求 7 刪除促銷組合標籤）
# ══════════════════════════════════════════════════════════════
def generate_daily_report(
    orders_df: pd.DataFrame,
    compare_df: pd.DataFrame,
    storage_df: pd.DataFrame,
    settings: dict,
) -> pd.DataFrame:
    if orders_df.empty:
        return pd.DataFrame()

    df = orders_df.copy()

    # ---- 1. 合併貨號 ------------------------------------------------
    if not compare_df.empty and "平台商品名稱" in compare_df.columns:
        m = compare_df[["平台商品名稱", "貨號"]].drop_duplicates("平台商品名稱")
        df = df.merge(
            m.rename(columns={"貨號": "_sku_cmp"}),
            on="平台商品名稱", how="left",
        )
        df["_sku"] = df["_sku_cmp"].where(
            df["_sku_cmp"].notna() & (~df["_sku_cmp"].isin(["", "nan"])),
            df["貨號"],
        )
        df.drop(columns=["_sku_cmp"], inplace=True)
    else:
        df["_sku"] = df["貨號"]

    # ---- 2. 合併成本 ------------------------------------------------
    if not storage_df.empty and "貨號" in storage_df.columns:
        avg = storage_df.groupby("貨號")["單位成本"].mean().reset_index()
        avg.columns = ["_sku", "平均成本"]
        df = df.merge(avg, on="_sku", how="left")
    else:
        df["平均成本"] = 0

    df["平均成本"] = pd.to_numeric(df["平均成本"], errors="coerce").fillna(0)
    df["_item_cost"] = df["數量"] * df["平均成本"]

    # ---- 3. 入庫名稱 lookup -----------------------------------------
    name_lookup: dict = {}
    if not storage_df.empty and "貨號" in storage_df.columns:
        nm = storage_df[["貨號", "商品名稱", "規格"]].drop_duplicates("貨號")
        nm["_name"] = nm["商品名稱"].fillna("") + "[" + nm["規格"].fillna("") + "]"
        name_lookup = nm.set_index("貨號")["_name"].to_dict()
    df["_storage_name"] = df["_sku"].map(name_lookup).fillna("")

    # ---- 4. 按訂單彙總 ----------------------------------------------
    grouped = df.groupby("訂單編號", sort=False)

    agg = grouped.agg(
        日期=("日期", "first"),
        平台=("平台", "first"),
        訂單狀態=("訂單狀態", "first"),
        營業額=("金額", "sum"),
        成本=("_item_cost", "sum"),
        賣家折扣=("賣家折扣", "first"),
    ).reset_index()

    # 貨號明細
    sku_detail = grouped.apply(
        lambda g: ";".join(
            f"{r['_sku']}({int(r['數量'])})"
            for _, r in g.iterrows()
            if pd.notna(r["_sku"]) and str(r["_sku"]) not in ("", "nan")
        ),
    ).reset_index(name="貨號明細")
    agg = agg.merge(sku_detail, on="訂單編號", how="left")

    # 入庫名稱
    name_detail = grouped.apply(
        lambda g: ",".join(sorted({
            n for n in g["_storage_name"] if n and n != "[]"
        })),
    ).reset_index(name="入庫名稱")
    agg = agg.merge(name_detail, on="訂單編號", how="left")

    # ---- 5. 特殊狀態 ------------------------------------------------
    cancel = agg["訂單狀態"].str.contains("取消", case=False, na=False)
    ret    = agg["訂單狀態"].str.contains("退貨|退款", case=False, na=False)

    agg.loc[cancel, ["營業額", "成本"]] = 0
    agg.loc[ret, "營業額"] = 0

    # ---- 6. 費用計算 ------------------------------------------------
    agg["賣家折扣"] = pd.to_numeric(agg["賣家折扣"], errors="coerce").fillna(0).abs()
    net = (agg["營業額"] - agg["賣家折扣"]).clip(lower=0)

    agg["運費折抵"]   = 0.0
    agg["成交手續費"] = 0.0
    agg["金流服務費"] = 0.0

    for plat in ("蝦皮", "露天", "官網"):
        mask = agg["平台"] == plat
        if not mask.any():
            continue
        fee  = settings.get(f"{plat}_成交手續費率", 0)
        pay  = settings.get(f"{plat}_金流服務費率", 0)
        thres = settings.get(f"{plat}_免運門檻", 0)
        ship  = settings.get(f"{plat}_運費折抵金額", 0)

        agg.loc[mask, "成交手續費"] = (net[mask] * fee).round(0)
        agg.loc[mask, "金流服務費"] = (net[mask] * pay).round(0)

        if thres > 0 and ship > 0:
            eligible = mask & (agg["營業額"] >= thres) & ~cancel & ~ret
            agg.loc[eligible, "運費折抵"] = ship

    agg.loc[cancel | ret, ["成交手續費", "金流服務費", "運費折抵"]] = 0

    # ---- 7. 淨利 ----------------------------------------------------
    agg["淨利"] = (
        agg["營業額"] - agg["賣家折扣"] - agg["運費折抵"]
        - agg["成交手續費"] - agg["金流服務費"] - agg["成本"]
    ).round(0)

    # ---- 8. 狀態標記 ------------------------------------------------
    agg["出貨狀態"] = ""
    agg.loc[cancel, "出貨狀態"] = "!取消!"
    agg.loc[ret, "出貨狀態"]    = "!退貨!"
    normal_no_cost = ~cancel & ~ret & ((agg["成本"] == 0) | agg["成本"].isna())
    agg.loc[normal_no_cost, "出貨狀態"] = "!未匹配!"

    # ---- 9. 排序 & 輸出 ---------------------------------------------
    agg = agg.sort_values("日期").reset_index(drop=True)

    cols = [
        "日期", "訂單編號", "入庫名稱", "貨號明細",
        "營業額", "賣家折扣", "運費折抵", "成交手續費", "金流服務費",
        "成本", "淨利", "出貨狀態", "平台",
    ]
    for c in cols:
        if c not in agg.columns:
            agg[c] = ""
    return agg[cols]


# ══════════════════════════════════════════════════════════════
# 月報表
# ══════════════════════════════════════════════════════════════
def generate_monthly_report(daily_df: pd.DataFrame) -> pd.DataFrame:
    if daily_df.empty:
        return pd.DataFrame()

    df = daily_df.copy()
    df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
    df["月份"] = df["日期"].dt.month

    monthly = df.groupby("月份").agg(
        營業額=("營業額", "sum"),
        成本=("成本", "sum"),
        賣家折扣=("賣家折扣", "sum"),
        運費折抵=("運費折抵", "sum"),
        成交手續費=("成交手續費", "sum"),
        金流服務費=("金流服務費", "sum"),
        淨利=("淨利", "sum"),
        訂單數=("訂單編號", "count"),
    ).reset_index()

    for c in ["營業額", "成本", "賣家折扣", "運費折抵", "成交手續費", "金流服務費", "淨利"]:
        monthly[c] = monthly[c].round(0).astype(int)

    return monthly.sort_values("月份").reset_index(drop=True)


# ══════════════════════════════════════════════════════════════
# 庫存明細
# ══════════════════════════════════════════════════════════════
def generate_inventory(
    storage_df: pd.DataFrame,
    delivery_df: pd.DataFrame,
) -> pd.DataFrame:
    if storage_df.empty:
        return pd.DataFrame()

    keys = ["主貨號", "商品名稱", "規格", "貨號"]
    for c in keys:
        if c not in storage_df.columns:
            return pd.DataFrame()

    ss = storage_df.groupby(keys).agg(
        進貨數量=("數量", "sum"),
        進貨金額=("總金額", "sum"),
        平均成本=("單位成本", "mean"),
    ).reset_index()

    if not delivery_df.empty and "貨號" in delivery_df.columns:
        ds = delivery_df.groupby("貨號").agg(
            銷售數量=("數量", "sum"),
            銷售金額=("金額", "sum"),
        ).reset_index()
        result = ss.merge(ds, on="貨號", how="left")
    else:
        result = ss.copy()
        result["銷售數量"] = 0
        result["銷售金額"] = 0

    result["銷售數量"] = result["銷售數量"].fillna(0).astype(int)
    result["銷售金額"] = result["銷售金額"].fillna(0)
    result["現有庫存"] = result["進貨數量"] - result["銷售數量"]
    result["平均成本"] = result["平均成本"].round(1)

    return result.reset_index(drop=True)


# ══════════════════════════════════════════════════════════════
# 庫存明細（對應 VBA InventoryDetails）
# ══════════════════════════════════════════════════════════════
def generate_inventory_details(
    storage_df: pd.DataFrame,
    delivery_df: pd.DataFrame,
) -> pd.DataFrame:
    """依入庫/出庫 xlsx 欄位產生庫存明細，對應 VBA InventoryDetails.vba 邏輯。"""
    if storage_df.empty:
        return pd.DataFrame()

    # 相容 load_storage() 轉換後的欄位名稱
    name_col   = "商品名稱" if "商品名稱" in storage_df.columns else "名稱"
    qty_col    = "數量"     if "數量"     in storage_df.columns else "入庫數量"
    amount_col = "總金額"   if "總金額"   in storage_df.columns else "金額"
    cost_col   = "單位成本" if "單位成本" in storage_df.columns else "單價"

    # 以貨號為 key 彙總入庫
    stg = (
        storage_df
        .groupby(["主貨號", "貨號", name_col, "規格"], dropna=False)
        .agg(
            進貨數量=(qty_col,    "sum"),
            進貨合計=(amount_col, "sum"),
            **{"平均成本": (cost_col, "mean")},
        )
        .reset_index()
        .rename(columns={name_col: "名稱"})
    )

    # 以貨號為 key 彙總出庫
    if not delivery_df.empty and "貨號" in delivery_df.columns:
        qty_d = "出庫數量" if "出庫數量" in delivery_df.columns else "數量"
        dly = (
            delivery_df
            .groupby("貨號", dropna=False)
            .agg(
                銷售數量=(qty_d,  "sum"),
                銷售合計=("金額", "sum"),
            )
            .reset_index()
        )
        result = stg.merge(dly, on="貨號", how="left")
    else:
        result = stg.copy()
        result["銷售數量"] = 0
        result["銷售合計"] = 0.0

    result["銷售數量"] = result["銷售數量"].fillna(0).astype(int)
    result["銷售合計"] = result["銷售合計"].fillna(0.0)
    result["現有庫存"] = result["進貨數量"] - result["銷售數量"]
    result["平均成本"] = (result["進貨合計"] / result["進貨數量"]).round(1)

    col_order = [
        "主貨號", "貨號", "名稱", "規格",
        "進貨數量", "進貨合計",
        "銷售數量", "銷售合計",
        "現有庫存",
        "平均成本",
    ]
    return result[col_order].reset_index(drop=True)
