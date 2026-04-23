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

    # ── 從訂單收集所有唯一 (平台商品名稱, 平台, 貨號) 組合（不先做 per-product 去重）──
    _null_like = {"nan", "none", "nat", "<na>", "none", "NaN", "None"}
    sku_all = (
        orders_df[["平台商品名稱", "平台", "貨號"]].copy()
        .assign(貨號=lambda d: d["貨號"].astype(str).str.strip())
    )
    sku_all = (
        sku_all[sku_all["貨號"].str.len() > 0]
        .assign(貨號=lambda d: d["貨號"].where(~d["貨號"].str.lower().isin({s.lower() for s in _null_like}), ""))
        .query('貨號 != ""')
        .drop_duplicates(subset=["平台商品名稱", "平台", "貨號"])
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

    # ── 為每個 (平台商品名稱, 平台) 選出代表 SKU ──
    # 若同一商品有多個 SKU（跨不同訂單），優先選「已在入庫中」的 SKU；
    # 若無任何 SKU 匹配，取第一個（字母序最小，確保結果穩定）。
    _sku_records = []
    for (prod_name, plat), grp in sku_all.groupby(["平台商品名稱", "平台"], sort=True):
        skus = sorted(grp["貨號"].tolist())
        matched_skus = [s for s in skus if s in stg_name_map]
        chosen = matched_skus[0] if matched_skus else skus[0]
        _sku_records.append({"平台商品名稱": prod_name, "平台": plat, "_sku": chosen})
    sku_src = pd.DataFrame(_sku_records) if _sku_records else pd.DataFrame(columns=["平台商品名稱", "平台", "_sku"])

    def _enrich(df: pd.DataFrame) -> pd.DataFrame:
        df = df.merge(sku_src, on=["平台商品名稱", "平台"], how="left")
        df["貨號"] = df["_sku"].fillna("").astype(str)
        df.drop(columns=["_sku"], errors="ignore", inplace=True)
        df["主貨號"] = df["貨號"].apply(lambda s: s.split("-")[0] if s else "")
        df["入庫品名"] = df["貨號"].map(stg_name_map)
        df["入庫品名"] = df.apply(
            lambda r: r["入庫品名"] if pd.notna(r["入庫品名"]) else "未匹配",
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
        # 先對現有對照表去重，避免歷史重複列影響結果
        existing_compare_df = existing_compare_df.drop_duplicates(
            subset=["平台商品名稱", "平台"]
        ).reset_index(drop=True)

        # 已匹配的列保留原貨號，不重新掃描；只對「未匹配」的列嘗試重新匹配
        _matched_mask = ~existing_compare_df["入庫品名"].astype(str).isin(["", "未匹配"])
        existing_matched = existing_compare_df[_matched_mask][result_cols].copy()
        existing_unmatched = _enrich(
            existing_compare_df[~_matched_mask][["平台商品名稱", "平台"]].copy()
        )

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
        result = pd.concat(
            [existing_matched, existing_unmatched[result_cols], new_only[result_cols]],
            ignore_index=True,
        )
    else:
        result = _enrich(all_prods)[result_cols].copy()

    return result.drop_duplicates(subset=["平台商品名稱", "平台"]).reset_index(drop=True)


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

        # items_ret：退回數量（部份退貨退貨行顯示用）
        return_qty_val = max(0, qty - effective_qty)
        items_ret = _expand_items_for_combo(sku, return_qty_val, stg, combo_df) if return_qty_val > 0 else []
        if items_ret and len(items_ret) == 1 and not stg_info:
            items_ret[0]["名稱"] = items_ret[0]["名稱"] or _s(row.get("商品名稱", ""))
            items_ret[0]["規格"] = items_ret[0]["規格"] or _s(row.get("商品選項名稱", ""))

        records.append({
            "_oid": oid, "_date": _s(row.get("訂單成立日期", ""))[:10],
            "_row_status": _row_status,
            "_has_ret": bool(ret_stat),        # 此列是否有退貨狀態
            "_effective_qty": effective_qty,   # 有效成交數量
            "_items_eff": items_eff,           # 有效數量商品（部份退貨保留行）
            "_items_orig": items_orig,         # 原始数量商品（整單退貨顯示用）
            "_items_ret": items_ret,           # 退回數量商品（部份退貨退貨行）
            "_line_amt": price * effective_qty,
            "_item_cost": stg_info.get("成本", 0) * effective_qty if stg_info else 0,
            "_matched": bool(stg_info),
            # order-level (take first per order)
            # 折扣優惠 = 賣家負擔優惠券（新）/ 賣場優惠券（舊）+ 賣家負擔蝦幣回饋券（新）/ 賣家蝦幣回饋券（舊）
            "_coupon": abs(_n(row.get("賣家負擔優惠券") or row.get("賣場優惠券", 0)))
                     + abs(_n(row.get("賣家負擔蝦幣回饋券") or row.get("賣家蝦幣回饋券", 0))),
            "_buyer_ship": _n(row.get("買家支付運費", 0)),
            "_plat_ship": _n(row.get("蝦皮補助運費", 0)),
            "_return_ship": abs(_n(row.get("退貨運費", 0))),
            "_tx_fee": abs(_n(row.get("成交手續費", 0))),
            "_svc_fee": abs(_n(row.get("其他服務費", 0))),
            "_pay_fee": abs(_n(row.get("金流與系統處理費", 0))),
        })

    def _shopee_row(date, oid, status, items, order_amt, coupon, buyer_ship, plat_ship,
                    ret_ship, tx_fee, svc_fee, pay_fee, cost_item, matched, remark_extra=""):
        total_cost = cost_item + coupon + ret_ship + tx_fee + svc_fee + pay_fee
        profit = order_amt - total_cost
        name_str, sku_str = _build_item_strings(items)
        note_parts = (["未匹配"] if not matched else []) + ([remark_extra] if remark_extra else [])
        return {
            "日期": date, "訂單編號": oid, "訂單狀態": status,
            "商品名稱": name_str, "貨號": sku_str,
            "訂單金額": round(order_amt, 0),
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
            "商品成本": round(cost_item, 0),
            "總成本": round(total_cost, 0),
            "淨利": round(profit, 0),
            "備註": "、".join(note_parts), "平台": "蝦皮",
        }

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

        buyer_ship = f["_buyer_ship"]
        plat_ship = f["_plat_ship"]
        ret_ship = f["_return_ship"] if any_ret else 0

        if is_partial_ret:
            # 部份退貨：拆成兩行——已完成（保留商品）+ 退貨（退回商品）
            retained_items = []
            for r in rows:
                if not r["_has_ret"]:
                    retained_items.extend(r["_items_orig"])
                elif r["_effective_qty"] > 0:
                    retained_items.extend(r["_items_eff"])
            returned_items = [it for r in rows if r["_has_ret"] for it in r["_items_ret"]]
            comp_amt = sum(r["_line_amt"] for r in rows)
            comp_cost_item = sum(r["_item_cost"] for r in rows)
            coupon = f["_coupon"]
            tx_fee = f["_tx_fee"]
            svc_fee = f["_svc_fee"]
            pay_fee = f["_pay_fee"]
            comp_matched = all(r["_matched"] for r in rows if not r["_has_ret"] or r["_effective_qty"] > 0)
            ret_matched = all(r["_matched"] for r in rows if r["_has_ret"])
            # 已完成行：保留商品、一般費用歸此行，退貨運費為 0
            result.append(_shopee_row(
                f["_date"], oid, "已完成", retained_items,
                comp_amt, coupon, buyer_ship, plat_ship,
                0, tx_fee, svc_fee, pay_fee, comp_cost_item, comp_matched, "部份退貨",
            ))
            # 退貨行：退回商品、費用均為 0，只計退貨運費
            result.append(_shopee_row(
                f["_date"], oid, "退貨", returned_items,
                0, 0, 0, 0,
                ret_ship, 0, 0, 0, 0, ret_matched, "部份退貨",
            ))
        elif is_full_ret:
            # 整單退貨：金額與商品成本歸零，只計退貨運費
            display_items = [it for r in rows for it in r["_items_orig"]]
            is_matched = all(r["_matched"] for r in rows)
            result.append(_shopee_row(
                f["_date"], oid, "退貨", display_items,
                0, 0, buyer_ship, plat_ship,
                ret_ship, 0, 0, 0, 0, is_matched,
            ))
        else:
            # 已完成 / 遺失賠償
            status = f["_row_status"]
            total_amt = sum(r["_line_amt"] for r in rows)
            total_cost_item = sum(r["_item_cost"] for r in rows)
            coupon = f["_coupon"]
            tx_fee = f["_tx_fee"]
            svc_fee = f["_svc_fee"]
            pay_fee = f["_pay_fee"]
            display_items = [it for r in rows for it in r["_items_orig"]]
            is_matched = all(r["_matched"] for r in rows)
            result.append(_shopee_row(
                f["_date"], oid, status, display_items,
                total_amt, coupon, buyer_ship, plat_ship,
                0, tx_fee, svc_fee, pay_fee, total_cost_item, is_matched,
            ))
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
        order_stat = _s(row.get("訂單狀態", ""))
        if "已領取退貨" in tx_stat:
            status = "未取貨"
        elif "取消" in tx_stat or "取消" in order_stat:
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
        logistics_diff = -max(0, buyer_ship - actual_ship)

        ret_ship = actual_ship if is_ret else 0
        tx_fee = sum(r["_tx_fee"] for r in rows) if not is_ret else 0
        svc_fee = sum(r["_svc_fee"] for r in rows) if not is_ret else 0
        pay_fee = f["_pay_fee"] if not is_ret else 0

        # 未取貨：買家支付運費、平台補助運費、實際運費支出、物流處理費歸零，只保留退貨運費
        out_buyer_ship = 0 if is_ret else buyer_ship
        out_plat_ship = 0 if is_ret else plat_ship
        out_actual_ship = 0 if is_ret else actual_ship
        out_logistics_diff = 0 if is_ret else logistics_diff

        # 總成本：商品成本＋折扣優惠＋未取貨/退貨運費＋成交手續費＋其他服務費＋金流與系統處理費＋發票處理費＋其他費用 + 物流處理費（運費差額）
        # （未取貨：只計實際運費）
        total_cost = total_cost_item + coupon + ret_ship + tx_fee + svc_fee + pay_fee + out_logistics_diff
        profit = total_amt - total_cost

        item_name_str, sku_str = _build_item_strings([it for r in rows for it in r["_items"]])
        result.append({
            "日期": f["_date"], "訂單編號": oid, "訂單狀態": status,
            "商品名稱": item_name_str, "貨號": sku_str,
            "訂單金額": round(total_amt, 0),
            "折扣優惠": round(coupon, 0),
            "買家支付運費": round(out_buyer_ship, 0),
            "平台補助運費": round(out_plat_ship, 0),
            "實際運費支出": round(out_actual_ship, 0),
            "物流處理費（運費差額）": round(out_logistics_diff, 0),
            "未取貨/退貨運費": round(ret_ship, 0),
            "成交手續費": round(tx_fee, 0),
            "其他服務費": round(svc_fee, 0),
            "金流與系統處理費": round(pay_fee, 0),
            "發票處理費": 0, "其他費用": 0,
            "商品成本": round(total_cost_item, 0),
            "總成本": round(total_cost, 0),
            "淨利": round(profit, 0),
            "備註": "" if all(r["_matched"] for r in rows) else "未匹配", "平台": "露天",
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
        logistics_diff = actual_ship - buyer_ship
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
            "備註": "" if all(r["_matched"] for r in rows) else "未匹配", "平台": "官網",
        })
    return result


def _process_custom(df: pd.DataFrame, stg: dict, combo_df=None) -> list[dict]:
    """Process custom (self-built) orders into daily report rows."""
    if df.empty:
        return []
    records = []
    for _, row in df.iterrows():
        oid = _s(row.get("訂單編號", ""))
        if not oid:
            continue

        status = _s(row.get("訂單狀態", "已完成"))
        # 跳過已取消
        if status in ("已取消", "取消"):
            continue

        sku = _s(row.get("貨號", ""))
        qty = int(_n(row.get("數量", 0)))
        price = _n(row.get("單價", 0))

        stg_info = stg.get(sku, {})
        item_cost = stg_info.get("成本", 0) * qty if stg_info else 0
        items = _expand_items_for_combo(sku, qty, stg, combo_df)

        records.append({
            "_oid": oid,
            "_date": _s(row.get("日期", ""))[:10].replace("/", "-"),
            "_status": status,
            "_items": items,
            "_line_amt": price * qty,
            "_item_cost": item_cost,
            "_matched": bool(stg_info),
            "_coupon": abs(_n(row.get("折扣優惠", 0))),
            "_buyer_ship": _n(row.get("買家支付運費", 0)),
            "_actual_ship": _n(row.get("實際運費", 0)),
            "_return_ship": abs(_n(row.get("未取貨/退貨運費", 0))),
            "_other_fee": abs(_n(row.get("其他費用", 0))),
        })

    result = []
    records.sort(key=lambda x: x["_oid"])
    from itertools import groupby as _groupby
    for oid, grp in _groupby(records, key=lambda x: x["_oid"]):
        rows = list(grp)
        f = rows[0]
        status = f["_status"]
        is_ret = status in ("退貨", "未取貨")

        total_amt = sum(r["_line_amt"] for r in rows) if not is_ret else 0
        total_cost_item = sum(r["_item_cost"] for r in rows) if not is_ret else 0

        coupon = f["_coupon"] if not is_ret else 0
        buyer_ship = f["_buyer_ship"]
        actual_ship = f["_actual_ship"]
        logistics_diff = actual_ship - buyer_ship
        ret_ship = f["_return_ship"]
        other_fee = f["_other_fee"]

        if is_ret:
            buyer_ship = 0
            actual_ship = 0
            logistics_diff = 0

        total_cost = total_cost_item + coupon + logistics_diff + ret_ship + other_fee
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
            "金流與系統處理費": 0, "發票處理費": 0,
            "其他費用": round(other_fee, 0),
            "商品成本": round(total_cost_item, 0),
            "總成本": round(total_cost, 0),
            "淨利": round(profit, 0),
            "備註": "" if all(r["_matched"] for r in rows) else "未匹配",
            "平台": "其他",
        })
    return result


def _process_mo(df: pd.DataFrame, stg: dict, combo_df=None) -> list[dict]:
    """Process MO店 raw orders into daily report rows."""
    if df.empty:
        return []
    records = []
    for _, row in df.iterrows():
        oid = _s(row.get("訂單編號", ""))
        if not oid:
            continue

        order_stat = _s(row.get("訂單狀態", ""))
        ret_reason = _s(row.get("銷退原因", ""))

        # 取消訂單不列入日報表
        if order_stat == "取消訂單":
            continue

        if ret_reason == "配送異常結案":
            status = "未取貨"
        elif order_stat == "已回收":
            status = "退貨"
        elif order_stat == "配送異常":
            status = "未取貨"
        elif order_stat == "配送結束":
            status = "已完成"
        else:
            status = "已完成"

        sku = _s(row.get("商品原廠編號", ""))
        qty = int(_n(row.get("數量", 0)))
        price = _n(row.get("商品售價", 0))

        stg_info = stg.get(sku, {})
        item_cost = stg_info.get("成本", 0) * qty if stg_info else 0
        items = _expand_items_for_combo(sku, qty, stg, combo_df)
        if len(items) == 1 and not stg_info:
            items[0]["名稱"] = items[0]["名稱"] or _s(row.get("商品名稱", ""))

        coupon = (abs(_n(row.get("單品折價券(商品自折)", 0)))
                  + abs(_n(row.get("行銷活動促銷(商品自折)", 0)))
                  + abs(_n(row.get("單店抵用券(商品自折)", 0))))
        buyer_ship = _n(row.get("客人支付運費", 0))
        plat_ship = abs(_n(row.get("商品滿額免運費", 0)))
        actual_ship = abs(_n(row.get("預估平台代扣運費(鑑賞期後:平台代扣運費)", 0)))
        tx_fee = abs(_n(row.get("成交手續費", 0)))
        svc_fee_pre = abs(_n(row.get("預購商品服務費", 0)))
        svc_fee_hidden = abs(_n(row.get("物流隱碼服務費", 0)))
        svc_fee_activity = abs(_n(row.get("活動服務費", 0)))
        other_svc = svc_fee_pre + svc_fee_hidden + svc_fee_activity
        pay_fee = abs(_n(row.get("金流與系統處理費", 0)))
        invoice_fee = abs(_n(row.get("發票處理費", 0)))

        # 未取貨/退貨運費：已回收 AND 配送異常 → 活動服務費 + 訂單進帳金額
        ret_ship = 0.0
        if order_stat in ("已回收", "配送異常"):
            ret_ship = abs(_n(row.get("活動服務費", 0))) + abs(_n(row.get("訂單進帳金額", 0)))

        records.append({
            "_oid": oid,
            "_date": _s(row.get("轉單日", ""))[:10].replace("/", "-"),
            "_status": status,
            "_items": items,
            "_line_amt": price * qty,
            "_item_cost": item_cost,
            "_matched": bool(stg_info),
            "_coupon": coupon,
            "_buyer_ship": buyer_ship,
            "_plat_ship": plat_ship,
            "_actual_ship": actual_ship,
            "_ret_ship": ret_ship,
            "_tx_fee": tx_fee,
            "_other_svc": other_svc,
            "_pay_fee": pay_fee,
            "_invoice_fee": invoice_fee,
        })

    result = []
    records.sort(key=lambda x: x["_oid"])
    from itertools import groupby as _groupby
    for oid, grp in _groupby(records, key=lambda x: x["_oid"]):
        rows = list(grp)
        f = rows[0]
        status = f["_status"]
        is_ret = status in ("退貨", "未取貨")

        total_amt = sum(r["_line_amt"] for r in rows) if not is_ret else 0
        total_cost_item = sum(r["_item_cost"] for r in rows) if not is_ret else 0

        coupon = sum(r["_coupon"] for r in rows) if not is_ret else 0
        buyer_ship = f["_buyer_ship"]
        plat_ship = f["_plat_ship"]
        actual_ship = f["_actual_ship"]
        logistics_diff = buyer_ship - actual_ship
        ret_ship = sum(r["_ret_ship"] for r in rows)
        tx_fee = sum(r["_tx_fee"] for r in rows) if not is_ret else 0
        other_svc = sum(r["_other_svc"] for r in rows) if not is_ret else 0
        pay_fee = f["_pay_fee"] if not is_ret else 0
        invoice_fee = f["_invoice_fee"] if not is_ret else 0

        if is_ret:
            buyer_ship = 0
            plat_ship = 0
            actual_ship = 0
            logistics_diff = 0

        total_cost = (total_cost_item + coupon + logistics_diff + ret_ship
                      + tx_fee + other_svc + pay_fee + invoice_fee)
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
            "其他服務費": round(other_svc, 0),
            "金流與系統處理費": round(pay_fee, 0),
            "發票處理費": round(invoice_fee, 0),
            "其他費用": 0,
            "商品成本": round(total_cost_item, 0),
            "總成本": round(total_cost, 0),
            "淨利": round(profit, 0),
            "備註": "" if all(r["_matched"] for r in rows) else "未匹配",
            "平台": "MO店",
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
    custom_raw: pd.DataFrame | None = None,
    mo_raw: pd.DataFrame | None = None,
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
        + (_process_custom(custom_raw, stg, combo_df) if custom_raw is not None else [])
        + (_process_mo(mo_raw, stg, combo_df) if mo_raw is not None else [])
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

def compute_monthly_auto_from_daily(daily_df: pd.DataFrame) -> pd.DataFrame:
    """從日報表計算月報表的自動欄位（按年份 + 月份匯總）。"""
    if daily_df.empty:
        return pd.DataFrame()

    df = daily_df.copy()
    df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
    df = df.dropna(subset=["日期"])
    if df.empty:
        return pd.DataFrame()

    df["年份"] = df["日期"].dt.year.astype(int)
    df["月份"] = df["日期"].dt.month.astype(int)

    def _s(col: str) -> pd.Series:
        return pd.to_numeric(df[col], errors="coerce").fillna(0) if col in df.columns else pd.Series(0, index=df.index)

    df["_手續費"] = _s("成交手續費") + _s("其他服務費") + _s("金流與系統處理費") + _s("發票處理費")

    agg = df.groupby(["年份", "月份"]).agg(
        營業額=("訂單金額", "sum") if "訂單金額" in df.columns else ("年份", "count"),
        商品成本=("商品成本", "sum") if "商品成本" in df.columns else ("年份", "count"),
        折扣優惠=("折扣優惠", "sum") if "折扣優惠" in df.columns else ("年份", "count"),
        手續費=("_手續費", "sum"),
    ).reset_index()

    # Columns that might not exist
    for src, dest in [("未取貨/退貨運費", "未取貨/退貨運費"),
                      ("物流處理費（運費差額）", "物流處理費（運費差額）"),
                      ("其他費用", "其他費用（一）")]:
        if src in df.columns:
            sub = df.groupby(["年份", "月份"])[src].sum().reset_index().rename(columns={src: dest})
            agg = agg.merge(sub, on=["年份", "月份"], how="left")
        else:
            agg[dest] = 0

    # Reset wrong columns that may have been named incorrectly above
    if "訂單金額" not in df.columns:
        agg["營業額"] = 0
    if "商品成本" not in df.columns:
        agg["商品成本"] = 0
    if "折扣優惠" not in df.columns:
        agg["折扣優惠"] = 0

    num_cols = ["營業額", "商品成本", "折扣優惠", "手續費",
                "未取貨/退貨運費", "物流處理費（運費差額）", "其他費用（一）"]
    for c in num_cols:
        if c not in agg.columns:
            agg[c] = 0
        agg[c] = pd.to_numeric(agg[c], errors="coerce").fillna(0).round(0).astype(int)

    return agg.sort_values(["年份", "月份"]).reset_index(drop=True)



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
