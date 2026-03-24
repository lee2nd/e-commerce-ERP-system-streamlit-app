"""
業務邏輯計算：對照表匹配、日報表 / 月報表 / 出庫 / 庫存明細產生。
"""
import pandas as pd
from typing import cast
pd.set_option('future.no_silent_downcasting', True)

# ══════════════════════════════════════════════════════════════
# 對照表自動匹配（需求 5：依照貨號自動匹配商品）
# ══════════════════════════════════════════════════════════════
def auto_match_compare_table(
    orders_df: pd.DataFrame,
    storage_df: pd.DataFrame | None = None,
    existing_compare_df: pd.DataFrame | None = None,
    combo_df: pd.DataFrame | None = None,
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
    stg_name_map: dict = {}
    if storage_df is not None and not storage_df.empty and "貨號" in storage_df.columns:
        nm = storage_df[["貨號", "商品名稱", "規格"]].drop_duplicates("貨號").copy()
        nm["貨號"] = nm["貨號"].astype(str).str.strip()
        nm["規格"] = nm["規格"].fillna("").astype(str)
        nm["商品名稱"] = nm["商品名稱"].fillna("").astype(str)
        nm["_label"] = nm.apply(
            lambda r: f"{r['商品名稱']}[{r['規格']}]" if r["規格"] else r["商品名稱"], axis=1
        )
        stg_name_map = nm.set_index("貨號")["_label"].to_dict()

    # ── 組合貨號 → 入庫品名 的映射 ──
    if combo_df is not None and not combo_df.empty:
        for combo_code in combo_df["組合貨號"].unique():
            if combo_code not in stg_name_map:
                sub = combo_df[combo_df["組合貨號"] == combo_code]
                parts = " + ".join(
                    f"{r['原料貨號']}×{int(r['原料數量'])}" for _, r in sub.iterrows()
                )
                stg_name_map[combo_code] = f"組合:{parts}"

    def _enrich(df: pd.DataFrame) -> pd.DataFrame:        
        df = df.merge(sku_src, on=["平台商品名稱", "平台"], how="left")
        df["貨號"] = df["_sku"].fillna("").astype(str)
        df.drop(columns=["_sku"], errors="ignore", inplace=True)
        df["主貨號"] = df["貨號"].apply(lambda s: s.split("-")[0] if s else "")
        df["入庫品名"] = df["貨號"].map(stg_name_map)
        df["入庫品名"] = df.apply(
            lambda r: r["入庫品名"] if pd.notna(r["入庫品名"]) else ("未匹配" if r["貨號"] else ""),
            axis=1,
        )
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
# 日報表
# ══════════════════════════════════════════════════════════════

def _build_stg_lookup(storage_df: pd.DataFrame, combo_df: pd.DataFrame | None = None) -> dict:
    """貨號 → {名稱, 規格, 主貨號, 成本} using average cost"""
    if storage_df.empty:
        return {}
    avg_cost = storage_df.groupby("貨號")["單位成本"].mean()
    lookup = {}
    for _, r in storage_df.drop_duplicates("貨號", keep="last").iterrows():
        sku = str(r.get("貨號", "")).strip()
        if sku and sku not in ("nan", ""):
            lookup[sku] = {
                "名稱": str(r.get("商品名稱", "")).strip(),
                "規格": str(r.get("規格", "")).strip(),
                "主貨號": str(r.get("主貨號", "")).strip(),
                "成本": float(avg_cost.get(sku, 0) or 0),
            }
    # 組合貨號：成本 = 各原料成本 × 數量 之總和
    if combo_df is not None and not combo_df.empty:
        for combo_code in combo_df["組合貨號"].unique():
            sub = combo_df[combo_df["組合貨號"] == combo_code]
            total_cost = sum(
                lookup.get(str(r["原料貨號"]).strip(), {}).get("成本", 0) * int(r["原料數量"])
                for _, r in sub.iterrows()
            )
            parts = " + ".join(
                f"{r['原料貨號']}×{int(r['原料數量'])}" for _, r in sub.iterrows()
            )
            lookup[combo_code] = {
                "名稱": f"組合:{parts}",
                "規格": "",
                "主貨號": combo_code.split("-")[0] if "-" in combo_code else combo_code,
                "成本": total_cost,
            }
    return lookup


def _n(val, default=0):
    """Safe numeric conversion"""
    try:
        v = pd.to_numeric(val, errors="coerce")
        return float(v) if pd.notna(v) else default
    except Exception:
        return default


def _s(val) -> str:
    v = str(val).strip()
    return "" if v in ("nan", "None") else v


def _build_item_strings(items: list[dict]) -> tuple[str, str]:
    """Build 商品名稱 and 貨號 summary strings from item list"""
    names, skus = [], []
    seen_names = set()
    for it in items:
        n, spec, sku, qty = it["名稱"], it["規格"], it["sku"], it["qty"]
        label = f"{n}[{spec}]" if spec else n
        if label and label not in seen_names:
            seen_names.add(label)
            names.append(label)
        if sku:
            skus.append(f"{sku}({qty})")
    return ", ".join(names), "; ".join(skus)


def _ruten_logistics(shipping_method: str, settings: dict) -> float:
    sm = shipping_method.lower()
    if "萊爾富" in sm:
        return float(settings.get("ruten_laerfu", 50))
    if "郵局" in sm or "郵政" in sm:
        return float(settings.get("ruten_post", 65))
    if "ok" in sm:
        return float(settings.get("ruten_ok", 60))
    if "全家" in sm:
        return float(settings.get("ruten_family", 60))
    if "7-11" in sm:
        return float(settings.get("ruten_7_11", 60))
    return float(settings.get("ruten_default_shipping", 65))


def _expand_items_for_combo(sku: str, qty: int, stg: dict, combo_df) -> list[dict]:
    """If sku is a combo code, return component items; otherwise return single item."""
    if combo_df is not None and not combo_df.empty:
        combo_set = set(combo_df["組合貨號"].astype(str).str.strip())
        if sku in combo_set:
            components = combo_df[combo_df["組合貨號"].astype(str).str.strip() == sku]
            items = []
            for _, comp in components.iterrows():
                mat_sku = str(comp["原料貨號"]).strip()
                mat_qty = int(comp["原料數量"])
                comp_stg = stg.get(mat_sku, {})
                items.append({
                    "名稱": comp_stg.get("名稱", mat_sku),
                    "規格": comp_stg.get("規格", ""),
                    "sku": mat_sku,
                    "qty": qty * mat_qty,
                })
            return items
    info = stg.get(sku, {})
    return [{"名稱": info.get("名稱", ""), "規格": info.get("規格", ""), "sku": sku, "qty": qty}]


def _process_shopee(df: pd.DataFrame, stg: dict, combo_df=None) -> list[dict]:
    if df.empty:
        return []
    records = []
    for _, row in df.iterrows():
        not_established = _s(row.get("不成立原因", ""))
        # 不成立原因有值但不含「遺失」→ 取消訂單，跳過
        if not_established and "遺失" not in not_established:
            continue

        oid = _s(row.get("訂單編號", ""))
        if not oid:
            continue

        ret_stat = _s(row.get("退貨 / 退款狀態", ""))
        if not_established and "遺失" in not_established:
            _row_status = "遺失賠償"
        elif ret_stat:
            _row_status = "退貨"
        else:
            _row_status = "已完成"

        sku = _s(row.get("商品選項貨號", "")) or _s(row.get("主商品貨號", ""))
        qty = int(_n(row.get("數量", 0)))
        # 退貨數量：計算實際有效數量
        return_qty = int(_n(row.get("退貨數量", 0)))
        effective_qty = max(0, qty - return_qty)

        raw_act = row.get("商品活動價格")
        act_p_raw = pd.to_numeric(str(raw_act), errors="coerce") if raw_act is not None else None
        act_p = float(act_p_raw) if act_p_raw is not None and pd.notna(act_p_raw) else None
        orig_p = _n(row.get("商品原價", 0))
        price = act_p if act_p is not None else orig_p

        stg_info = stg.get(sku, {})

        # items_eff：有效數量（用於部份退貨的金額與庫存計算）
        items_eff = _expand_items_for_combo(sku, effective_qty, stg, combo_df) if effective_qty > 0 else []
        if items_eff and len(items_eff) == 1 and not stg_info:
            items_eff[0]["名稱"] = items_eff[0]["名稱"] or _s(row.get("商品名稱", ""))
            items_eff[0]["規格"] = items_eff[0]["規格"] or _s(row.get("商品選項名稱", ""))

        # items_orig：原始數量（整單退貨時仍顯示商品名稱以供識別）
        items_orig = _expand_items_for_combo(sku, qty, stg, combo_df)
        if len(items_orig) == 1 and not stg_info:
            items_orig[0]["名稱"] = items_orig[0]["名稱"] or _s(row.get("商品名稱", ""))
            items_orig[0]["規格"] = items_orig[0]["規格"] or _s(row.get("商品選項名稱", ""))

        records.append({
            "_oid": oid, "_date": _s(row.get("訂單成立日期", ""))[:10],
            "_row_status": _row_status,
            "_has_ret": bool(ret_stat),        # 此列是否有退貨狀態
            "_effective_qty": effective_qty,   # 有效成交數量
            "_items_eff": items_eff,           # 有效數量商品（部份退貨）
            "_items_orig": items_orig,         # 原始数量商品（整單退貨顯示用）
            "_line_amt": price * effective_qty,
            "_item_cost": stg_info.get("成本", 0) * effective_qty if stg_info else 0,
            "_matched": bool(stg_info),
            # order-level (take first per order)
            "_coupon": abs(_n(row.get("賣場優惠券", 0))),
            "_buyer_ship": _n(row.get("買家支付運費", 0)),
            "_plat_ship": _n(row.get("蝦皮補助運費", 0)),
            "_return_ship": abs(_n(row.get("退貨運費", 0))),
            "_tx_fee": abs(_n(row.get("成交手續費", 0))),
            "_svc_fee": abs(_n(row.get("其他服務費", 0))),
            "_pay_fee": abs(_n(row.get("金流與系統處理費", 0))),
        })

    result = []
    from itertools import groupby as _groupby
    records.sort(key=lambda x: x["_oid"])
    for oid, grp in _groupby(records, key=lambda x: x["_oid"]):
        rows = list(grp)
        f = rows[0]

        any_ret = any(r["_has_ret"] for r in rows)
        # 整單退貨：所有商品列都有退貨且有效數量均為 0
        is_full_ret = any_ret and all(r["_effective_qty"] == 0 for r in rows)
        # 部份退貨：有部分商品退貨但仍有有效數量
        is_partial_ret = any_ret and not is_full_ret

        if any_ret:
            status = "退貨"
        else:
            status = f["_row_status"]  # 已完成 or 遺失賠償

        ret_ship = f["_return_ship"] if any_ret else 0
        buyer_ship = f["_buyer_ship"]
        plat_ship = f["_plat_ship"]

        if is_full_ret:
            # 整單退貨：金額與商品成本歸零，只計退貨運費
            total_amt = 0
            total_cost_item = 0
            coupon = 0
            tx_fee = 0
            svc_fee = 0
            pay_fee = 0
            display_items = [it for r in rows for it in r["_items_orig"]]
            is_matched = any(r["_matched"] for r in rows)
        else:
            # 已完成 / 遺失賠償 / 部份退貨：以有效數量計算
            total_amt = sum(r["_line_amt"] for r in rows)
            total_cost_item = sum(r["_item_cost"] for r in rows)
            coupon = f["_coupon"]
            tx_fee = f["_tx_fee"]
            svc_fee = f["_svc_fee"]
            pay_fee = f["_pay_fee"]
            if is_partial_ret:
                # 只顯示有效成交的商品
                display_items = [it for r in rows for it in r["_items_eff"]]
                is_matched = any(r["_matched"] and r["_effective_qty"] > 0 for r in rows)
            else:
                display_items = [it for r in rows for it in r["_items_orig"]]
                is_matched = any(r["_matched"] for r in rows)

        # 總成本：商品成本＋折扣優惠＋退貨運費＋成交手續費＋其他服務費＋金流與系統處理費＋發票處理費＋其他費用 + 物流處理費（運費差額）
        total_cost = total_cost_item + coupon + ret_ship + tx_fee + svc_fee + pay_fee
        profit = total_amt - total_cost

        item_name_str, sku_str = _build_item_strings(display_items)
        result.append({
            "日期": f["_date"], "訂單編號": oid, "訂單狀態": status,
            "商品名稱": item_name_str, "貨號": sku_str,
            "訂單金額": round(total_amt, 0),
            "折扣優惠": round(coupon, 0),
            "買家支付運費": round(buyer_ship, 0),
            "平台補助運費": round(plat_ship, 0),
            "實際運費支出": round(buyer_ship + plat_ship, 0),
            "物流處理費（運費差額）": 0,
            "未取貨/退貨運費": round(ret_ship, 0),
            "成交手續費": round(tx_fee, 0),
            "其他服務費": round(svc_fee, 0),
            "金流與系統處理費": round(pay_fee, 0),
            "發票處理費": 0, "其他費用": 0,
            "商品成本": round(total_cost_item, 0),
            "總成本": round(total_cost, 0),
            "淨利": round(profit, 0),
            "備註": "" if is_matched else "未匹配", "平台": "蝦皮",
        })
    return result


def _process_ruten(df: pd.DataFrame, stg: dict, settings: dict, combo_df=None) -> list[dict]:
    if df.empty:
        return []
    records = []
    for _, row in df.iterrows():
        oid = _s(row.get("訂單編號", ""))
        if not oid:
            continue

        tx_stat = _s(row.get("交易狀況", ""))
        if "已領取退貨" in tx_stat:
            status = "未取貨"
        elif "取消" in tx_stat:
            continue  # skip cancelled
        else:
            status = "已完成"

        sku = _s(row.get("賣家自用料號", ""))
        qty = int(_n(row.get("數量", 0)))
        price = _n(row.get("單價", 0))

        stg_info = stg.get(sku, {})
        item_cost = stg_info.get("成本", 0) * qty if stg_info else 0
        items = _expand_items_for_combo(sku, qty, stg, combo_df)
        if len(items) == 1 and not stg_info:
            item_spec_raw = _s(row.get("規格", "")) + ("::" + _s(row.get("項目", "")) if _s(row.get("項目", "")) else "")
            items[0]["名稱"] = items[0]["名稱"] or _s(row.get("商品名稱", ""))
            items[0]["規格"] = items[0]["規格"] or item_spec_raw

        # 成交手續費：單價 × 數量 × 3%，四捨五入，最低1，最高400
        tx_fee_line = max(1, min(round(price * qty * 0.03), 400))
        # 其他服務費：單價 × 數量 × 5%，四捨五入，最低1，最高300
        svc_fee_line = max(1, min(round(price * qty * 0.05), 300))

        ship_method = _s(row.get("運送方式", "")) or _s(row.get("付款方式", ""))
        actual_ship = _ruten_logistics(ship_method, settings)
        buyer_ship = _n(row.get("運費", 0))
        ruten_disc = abs(_n(row.get("露天折扣碼金額", 0)))
        checkout_total = _n(row.get("結帳總金額", 0))
        # 金流與系統處理費：(結帳總金額 + 露天折扣碼金額) × 1.5%，四捨五入，最低1，無上限
        pay_fee = max(1, round((checkout_total + ruten_disc) * 0.015))

        records.append({
            "_oid": oid, "_date": _s(row.get("結帳時間", ""))[:10].replace("/", "-"),
            "_status": status,
            "_items": items,
            "_line_amt": price * qty,
            "_item_cost": item_cost,
            "_matched": bool(stg_info),
            "_coupon": abs(_n(row.get("賣家折扣碼金額", 0))),
            "_buyer_ship": buyer_ship,
            "_actual_ship": actual_ship,
            "_ship_method": ship_method,
            "_tx_fee": tx_fee_line,
            "_svc_fee": svc_fee_line,
            "_pay_fee": pay_fee,
        })

    result = []
    records.sort(key=lambda x: x["_oid"])
    from itertools import groupby as _groupby
    for oid, grp in _groupby(records, key=lambda x: x["_oid"]):
        rows = list(grp)
        f = rows[0]
        status = f["_status"]
        is_ret = status == "未取貨"

        total_amt = sum(r["_line_amt"] for r in rows) if not is_ret else 0
        total_cost_item = sum(r["_item_cost"] for r in rows) if not is_ret else 0

        coupon = f["_coupon"] if not is_ret else 0  # 未取貨不計折扣優惠
        buyer_ship = f["_buyer_ship"]
        actual_ship = f["_actual_ship"]
        plat_ship = max(0, actual_ship - buyer_ship)
        logistics_diff = abs(buyer_ship - actual_ship)

        ret_ship = actual_ship if is_ret else 0
        tx_fee = sum(r["_tx_fee"] for r in rows) if not is_ret else 0
        svc_fee = sum(r["_svc_fee"] for r in rows) if not is_ret else 0
        pay_fee = f["_pay_fee"] if not is_ret else 0

        # 總成本：商品成本＋折扣優惠＋未取貨/退貨運費＋成交手續費＋其他服務費＋金流與系統處理費＋發票處理費＋其他費用 + 物流處理費（運費差額）
        # （未取貨：只計實際運費）
        total_cost = total_cost_item + coupon + ret_ship + tx_fee + svc_fee + pay_fee + logistics_diff
        profit = total_amt - total_cost

        item_name_str, sku_str = _build_item_strings([it for r in rows for it in r["_items"]])
        result.append({
            "日期": f["_date"], "訂單編號": oid, "訂單狀態": status,
            "商品名稱": item_name_str, "貨號": sku_str,
            "訂單金額": round(total_amt, 0),
            "折扣優惠": round(coupon, 0),
            "買家支付運費": round(buyer_ship, 0),
            "平台補助運費": round(plat_ship, 0),
            "實際運費支出": round(actual_ship, 0),
            "物流處理費（運費差額）": round(logistics_diff, 0),
            "未取貨/退貨運費": round(ret_ship, 0),
            "成交手續費": round(tx_fee, 0),
            "其他服務費": round(svc_fee, 0),
            "金流與系統處理費": round(pay_fee, 0),
            "發票處理費": 0, "其他費用": 0,
            "商品成本": round(total_cost_item, 0),
            "總成本": round(total_cost, 0),
            "淨利": round(profit, 0),
            "備註": "" if any(r["_matched"] for r in rows) else "未匹配", "平台": "露天",
        })
    return result


def _process_easystore(df: pd.DataFrame, stg: dict, settings: dict, combo_df=None) -> list[dict]:
    if df.empty:
        return []
    actual_ship = float(settings.get("easystore_shipping", 65))

    # forward-fill order-level columns
    order_cols = ["Order Name", "Date", "Subtotal", "Shipping Fee", "Order Discount",
                  "Credit Used", "Financial Status", "Remark",
                  "Fulfillment Service", "Fulfillment Status"]
    
    df = df.copy()
    
    # 1. 先把 Order Name 填滿，以便後續可以正確 groupby
    if "Order Name" in df.columns:
        df["Order Name"] = df["Order Name"].ffill()
        
        # 2. 找出確實存在於 df 中且不是 Order Name 的欄位
        valid_cols = [c for c in order_cols if c in df.columns and c != "Order Name"]
        
        if valid_cols:
            # 3. 一次性對所有目標欄位進行 groupby + ffill (效能更好，且不用 lambda)
            df[valid_cols] = df.groupby("Order Name")[valid_cols].ffill()
    # 預先計算每筆訂單的最後一列 Transaction Status
    _tx_col = next((c for c in ("Transaction status", "Transaction Status") if c in df.columns), None)
    last_tx_status: dict[str, str] = {}

    if _tx_col and "Order Name" in df.columns:
        _tx_df = df[["Order Name", _tx_col]].copy()
        _tx_df[_tx_col] = _tx_df[_tx_col].fillna("").astype(str).str.strip()
        _tx_df = _tx_df[_tx_df[_tx_col] != ""]
        
        if not _tx_df.empty:
            # Use cast to silence the type checker
            raw_dict = _tx_df.groupby("Order Name")[_tx_col].last().to_dict()
            last_tx_status = cast(dict[str, str], raw_dict)

    if "Item Name" in df.columns:
        df = df[df["Item Name"].notna() & (df["Item Name"].astype(str).str.strip() != "")]

    records = []
    for _, row in df.iterrows():
        oid = _s(row.get("Order Name", ""))
        if not oid:
            continue

        fulfill_svc = _s(row.get("Fulfillment Service", ""))
        fulfill_stat = _s(row.get("Fulfillment Status", ""))

        # 未出貨，取消訂單：Fulfillment Service 為空 且 Fulfillment Status = "Restocked"/"Unfulfilled" → 跳過
        if not fulfill_svc and fulfill_stat in ("Restocked", "Unfulfilled"):
            continue

        refunded_amt = _n(row.get("Refunded Amount", 0))
        last_tx = last_tx_status.get(oid, "")

        # 已出貨，退貨：Fulfillment Service 不為空 且 Refunded Amount != 0
        if fulfill_svc and refunded_amt != 0:
            status = "退貨"
        # 已出貨，未取貨：Fulfillment Service 不為空 且 (Fulfillment Status = "Restocked"/"Unfulfilled" 或 最後一筆 Transaction Status = "Pending")
        elif fulfill_svc and (fulfill_stat in ("Restocked", "Unfulfilled") or last_tx == "Pending"):
            status = "未取貨"
        else:
            status = "已完成"

        sku = _s(row.get("Item SKU", ""))
        qty = int(_n(row.get("Quantity", 0)))
        price = _n(row.get("Item Price", 0))

        stg_info = stg.get(sku, {})
        item_cost = stg_info.get("成本", 0) * qty if stg_info else 0
        items = _expand_items_for_combo(sku, qty, stg, combo_df)
        if len(items) == 1 and not stg_info:
            items[0]["名稱"] = items[0]["名稱"] or _s(row.get("Item Name", ""))
            items[0]["規格"] = items[0]["規格"] or _s(row.get("Item Variant", ""))

        subtotal = _n(row.get("Subtotal", 0))
        order_disc = abs(_n(row.get("Order Discount", 0)))
        credit = abs(_n(row.get("Credit Used", 0)))
        buyer_ship = _n(row.get("Shipping Fee", 0))

        records.append({
            "_oid": oid, "_date": _s(row.get("Date", ""))[:10],
            "_status": status,
            "_items": items,
            "_line_amt": price * qty,
            "_item_cost": item_cost,
            "_matched": bool(stg_info),
            "_subtotal": subtotal,
            "_order_disc": order_disc,
            "_credit": credit,
            "_buyer_ship": buyer_ship,
        })

    result = []
    records.sort(key=lambda x: x["_oid"])
    from itertools import groupby as _groupby
    for oid, grp in _groupby(records, key=lambda x: x["_oid"]):
        rows = list(grp)
        f = rows[0]
        status = f["_status"]
        is_ret = status in ("未取貨", "退貨")
        is_nontaken = status == "未取貨"

        total_amt = f["_subtotal"] if not is_ret else 0
        total_cost_item = sum(r["_item_cost"] for r in rows) if not is_ret else 0

        coupon = (f["_order_disc"] + f["_credit"]) if not is_nontaken else 0  # 未取貨不計折扣優惠
        buyer_ship = f["_buyer_ship"]
        logistics_diff = abs(buyer_ship - actual_ship)
        ret_ship = actual_ship if is_ret else 0

        # 總成本：商品成本＋折扣優惠＋未取貨/退貨運費＋成交手續費＋其他服務費＋金流與系統處理費＋發票處理費＋其他費用 + 物流處理費（運費差額）
        # （未取貨：只計實際運費）
        total_cost = total_cost_item + coupon + ret_ship + logistics_diff
        profit = total_amt - total_cost

        item_name_str, sku_str = _build_item_strings([it for r in rows for it in r["_items"]])
        result.append({
            "日期": f["_date"], "訂單編號": oid, "訂單狀態": status,
            "商品名稱": item_name_str, "貨號": sku_str,
            "訂單金額": round(total_amt, 0),
            "折扣優惠": round(coupon, 0),
            "買家支付運費": round(buyer_ship, 0),
            "平台補助運費": 0,
            "實際運費支出": round(actual_ship, 0),
            "物流處理費（運費差額）": round(logistics_diff, 0),
            "未取貨/退貨運費": round(ret_ship, 0),
            "成交手續費": 0, "其他服務費": 0,
            "金流與系統處理費": 0, "發票處理費": 0, "其他費用": 0,
            "商品成本": round(total_cost_item, 0),
            "總成本": round(total_cost, 0),
            "淨利": round(profit, 0),
            "備註": "" if any(r["_matched"] for r in rows) else "未匹配", "平台": "官網",
        })
    return result


def generate_daily_report(
    shopee_raw: pd.DataFrame,
    ruten_raw: pd.DataFrame,
    easystore_raw: pd.DataFrame,
    compare_df: pd.DataFrame,
    storage_df: pd.DataFrame,
    settings: dict,
    combo_df: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """
    Generate daily report from raw platform DataFrames.
    settings keys:
      ruten_7_11, ruten_family, ruten_ok, ruten_laerfu, ruten_post,
      ruten_default_shipping, easystore_shipping
    """
    stg = _build_stg_lookup(storage_df, combo_df)

    all_records = (
        _process_shopee(shopee_raw, stg, combo_df)
        + _process_ruten(ruten_raw, stg, settings, combo_df)
        + _process_easystore(easystore_raw, stg, settings, combo_df)
    )
    if not all_records:
        return pd.DataFrame()

    df = pd.DataFrame(all_records)
    df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
    df = df.sort_values(["日期", "平台", "訂單編號"]).reset_index(drop=True)

    # Mark unmatched orders (商品名稱 is empty and 商品成本 == 0)
    df["_unmatched"] = (df["商品名稱"].fillna("") == "") & (df["商品成本"] == 0)
    # (keep _unmatched as internal; page can use it for colouring)

    return df


# ══════════════════════════════════════════════════════════════
# 月報表
# ══════════════════════════════════════════════════════════════
def generate_monthly_report(daily_df: pd.DataFrame) -> pd.DataFrame:
    if daily_df.empty:
        return pd.DataFrame()

    df = daily_df.copy()
    df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
    df["月份"] = df["日期"].dt.to_period("M").astype(str)

    # Map new column names → monthly aggregation
    _col = lambda name: name if name in df.columns else None
    agg_dict = {"訂單數": ("訂單編號", "count")}
    for col in ["訂單金額", "折扣優惠", "未取貨/退貨運費", "成交手續費", "其他服務費",
                "金流與系統處理費", "商品成本", "總成本", "淨利",
                # fallback old names
                "營業額", "成本", "賣家折扣", "運費折抵", "金流服務費"]:
        if col in df.columns:
            agg_dict[col] = (col, "sum")

    monthly = df.groupby("月份").agg(**agg_dict).reset_index()
    for c in list(monthly.columns):
        if c not in ("月份",):
            monthly[c] = pd.to_numeric(monthly[c], errors="coerce").fillna(0).round(0).astype(int)

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
def _expand_combo_delivery(delivery_df: pd.DataFrame, combo_df: pd.DataFrame) -> pd.DataFrame:
    """將出庫中的組合貨號展開為原料貨號明細，非組合貨號的記錄保持不變。"""
    if delivery_df.empty or combo_df.empty:
        return delivery_df

    combo_set = set(combo_df["組合貨號"].astype(str).str.strip().unique())
    qty_col = "出庫數量" if "出庫數量" in delivery_df.columns else "數量"

    normal_rows = delivery_df[~delivery_df["貨號"].astype(str).str.strip().isin(combo_set)]
    combo_rows = delivery_df[delivery_df["貨號"].astype(str).str.strip().isin(combo_set)]

    if combo_rows.empty:
        return delivery_df

    expanded = []
    for _, row in combo_rows.iterrows():
        combo_code = str(row["貨號"]).strip()
        order_qty = int(row[qty_col]) if pd.notna(row[qty_col]) else 0
        components = combo_df[combo_df["組合貨號"].astype(str).str.strip() == combo_code]
        for _, comp in components.iterrows():
            mat_sku = str(comp["原料貨號"]).strip()
            mat_qty = int(comp["原料數量"])
            new_row = row.copy()
            new_row["貨號"] = mat_sku
            new_row[qty_col] = order_qty * mat_qty
            new_row["金額"] = 0  # 組合原料不計金額（由組合本身計）
            # 更新 主貨號 為原料自身的主貨號
            if "主貨號" in new_row.index:
                new_row["主貨號"] = mat_sku.split("-")[0] if "-" in mat_sku else mat_sku
            expanded.append(new_row)

    if expanded:
        expanded_df = pd.DataFrame(expanded)
        return pd.concat([normal_rows, expanded_df], ignore_index=True)
    return normal_rows.reset_index(drop=True)


def generate_inventory_details(
    storage_df: pd.DataFrame,
    delivery_df: pd.DataFrame,
    combo_df: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """依入庫/出庫 xlsx 欄位產生庫存明細，對應 VBA InventoryDetails.vba 邏輯。"""
    if storage_df.empty:
        return pd.DataFrame()

    # 組合貨號展開：將出庫中的組合貨號拆為原料貨號
    if combo_df is not None and not combo_df.empty:
        delivery_df = _expand_combo_delivery(delivery_df, combo_df)

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
