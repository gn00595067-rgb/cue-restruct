"""
回饋贈檔模組 (Rebate Bonus Slots)
依執行天數與預算門檻判定回饋%，以實作價換算檔次；報表僅 rate(Net) 顯示分區定價，Package-cost 顯示「回饋」。
"""

import math
from config import (
    REFERENCE_STD_SPOTS,
    REGIONS_ORDER,
    REBATE_RULES_NAT,
    REBATE_RULES_REGION,
)
from utils import get_sec_factor, calculate_schedule


def _days_bucket(days):
    """執行天數對應門檻：2個月內=60, 3個月內=91, 6個月內=181, 以上=None。"""
    if days <= 60:
        return 60
    if days <= 91:
        return 91
    if days <= 181:
        return 181
    return None


def _best_rebate_nat_rad(budget, days):
    """全省全家廣播：依 115 年度表，回傳 (全家%, 家樂福%)，家樂福% 為全家達標時可選回饋家樂福的%。"""
    bucket = _days_bucket(days)
    for days_max, rad_lo, rad_hi, cf_lo, cf_hi, pct_rad, pct_cf in REBATE_RULES_NAT:
        if days_max is not None and bucket is not None and bucket > days_max:
            continue
        if rad_lo is not None and budget < rad_lo:
            continue
        if rad_hi is not None and budget >= rad_hi:
            continue
        return (pct_rad, pct_cf)
    return None


def _best_rebate_nat_cf(budget, days):
    """全省家樂福：依 115 年度表，回傳家樂福達標時的回饋%。"""
    bucket = _days_bucket(days)
    for days_max, rad_lo, rad_hi, cf_lo, cf_hi, pct_rad, pct_cf in REBATE_RULES_NAT:
        if days_max is not None and bucket is not None and bucket > days_max:
            continue
        if cf_lo is not None and budget < cf_lo:
            continue
        if cf_hi is not None and budget >= cf_hi:
            continue
        return pct_cf
    return None


def _best_rebate_nat(media, budget, days):
    """全省：依 115 年度表，全家回傳全家%、家樂福回傳家樂福%。保留給相容用。"""
    if media == "全家廣播":
        t = _best_rebate_nat_rad(budget, days)
        return t[0] if t is not None else None  # 全家% (0 也算達標，可選家樂福)
    return _best_rebate_nat_cf(budget, days)


def _best_rebate_region(media, region, budget, days):
    """單區（北區/桃竹苗/中區/雲嘉南）：取符合條件中回饋%最高。"""
    if media != "全家廣播":
        return None
    group = "北區" if region == "北區" else "其他"
    bucket = _days_bucket(days)
    best = 0
    for m, g, days_max, low, high, pct in REBATE_RULES_REGION:
        if m != media or g != group:
            continue
        if days_max is not None and bucket is not None and bucket > days_max:
            continue
        if low is not None and budget < low:
            continue
        if high is not None and budget > high:
            continue
        if pct > best:
            best = pct
    return best if best > 0 else None


def _build_rebate_pct_map(config, total_budget, days):
    """同平台、同區域加總金額後決定回饋%。115年度表：全家達標存 (全家%, 家樂福可選%)，家樂福達標存 家樂福%。"""
    rebate_pct_map = {}
    cfg_rad = config.get("全家廣播")
    if cfg_rad:
        m_budget_rad = total_budget * (cfg_rad["share"] / 100.0)
        sec_shares_rad = cfg_rad.get("sec_shares") or {}
        total_nat_rad = sum(m_budget_rad * (pct / 100.0) for _sec, pct in sec_shares_rad.items())
        # 全省全家達標：僅當 is_national=True 時才比對全省門檻；單區時不設全省
        is_nat_rad = cfg_rad.get("is_national", False)
        if is_nat_rad:
            nat_rad_pair = _best_rebate_nat_rad(total_nat_rad, days)
            if nat_rad_pair is not None:
                pct_rad, pct_cf = nat_rad_pair
                rebate_pct_map[("全家廣播", "全省")] = pct_rad
                rebate_pct_map[("全家廣播", "全省_家樂福")] = pct_cf  # 全家達標時可選回饋家樂福的%
        regs_rad = cfg_rad.get("regions", [])
        region_shares_rad = cfg_rad.get("region_shares") or {}
        for r in ("北區", "桃竹苗", "中區", "雲嘉南"):
            if is_nat_rad and region_shares_rad and r in region_shares_rad:
                min_ratio = min(region_shares_rad.values())
                total_r = sum(
                    m_budget_rad * (pct / 100.0) * (region_shares_rad[r] - min_ratio) / 100.0
                    for _sec, pct in sec_shares_rad.items()
                )
            else:
                # 單區（無自訂區域比例）或全省：無 region_shares 時，單區預算 = 該區所佔全家預算
                total_ratio = sum(region_shares_rad.get(x, 0) for x in regs_rad)
                if total_ratio > 0:
                    total_r = sum(
                        m_budget_rad * (pct / 100.0) * (region_shares_rad.get(r, 0) / total_ratio)
                        for _sec, pct in sec_shares_rad.items()
                    )
                else:
                    # 單區且未設區域比例：全家預算全算在選定區域，多區則均分
                    n_regs = len([x for x in regs_rad if x in ("北區", "桃竹苗", "中區", "雲嘉南")]) or 1
                    share_r = (m_budget_rad / n_regs) if r in regs_rad else 0
                    total_r = sum(
                        (share_r * (pct / 100.0)) for _sec, pct in sec_shares_rad.items()
                    ) if share_r > 0 else 0
            pct_r = _best_rebate_region("全家廣播", r, total_r, days)
            if pct_r is not None:
                rebate_pct_map[("全家廣播", r)] = pct_r
    cfg_cf = config.get("家樂福")
    if cfg_cf:
        m_budget_cf = total_budget * (cfg_cf["share"] / 100.0)
        sec_shares_cf = cfg_cf.get("sec_shares") or {}
        total_nat_cf = sum(m_budget_cf * (pct / 100.0) for _sec, pct in sec_shares_cf.items())
        pct_cf = _best_rebate_nat_cf(total_nat_cf, days)
        if pct_cf is not None:
            rebate_pct_map[("家樂福", "全省")] = pct_cf
    return rebate_pct_map


def get_rebate_qualified_platforms(config, total_budget, active_days):
    """回傳有達標回饋門檻的平台集合（'全家廣播'、'家樂福'）。"""
    m = _build_rebate_pct_map(config, total_budget, active_days)
    return set(media for (media, _scope) in m)


def get_rebate_qualification_detail(config, total_budget, active_days):
    """
    回傳達標細節，供 UI 顯示回饋選項。
    回傳 dict:
      nat_rad: bool 全省全家廣播達標
      nat_cf: bool 全省家樂福達標
      region_rad: list 有達標的單區全家廣播區域 (e.g. ["北區", "桃竹苗"])
    """
    m = _build_rebate_pct_map(config, total_budget, active_days)
    nat_rad = ("全家廣播", "全省") in m
    nat_cf = ("家樂福", "全省") in m
    region_rad = [r for r in ("北區", "桃竹苗", "中區", "雲嘉南") if ("全家廣播", r) in m]
    return {"nat_rad": nat_rad, "nat_cf": nat_cf, "region_rad": region_rad}


def compute_rebate_rows(config, total_budget, active_days, rows, pricing_db, sec_factors, store_counts_num, regions_order, rebate_platform_choice=None, rebate_nat_destination=None, rebate_region_destination=None, apply_nat_rad=True, apply_nat_cf=True, apply_region_rad=True):
    """
    依 config 與現有 rows 計算回饋贈檔列。
    - rebate_platform_choice: 保留相容，等同 rebate_nat_destination。
    - rebate_nat_destination: 全省全家達標時回饋到 "全家廣播" 或 "家樂福"；僅在 apply_nat_rad=True 時有效。
    - rebate_region_destination: 單區全家達標時回饋顯示區域；僅在 apply_region_rad=True 時有效。
    - apply_nat_rad: 是否套用「全省全家達標」回饋（可選回饋全家或家樂福）。
    - apply_nat_cf: 是否套用「家樂福達標」回饋（家樂福預算×%→家樂福）；可與全家→家樂福併存。
    - apply_region_rad: 是否套用「單區全家達標」回饋。
    回傳要插入的 (insert_after_index, rebate_row) 列表。
    若呼叫端需要運算邏輯紀錄，可改為解構：inserts, rebate_logs = compute_rebate_rows(...)，rebate_logs 為回饋檔次計算明細。
    """
    result = []
    rebate_logs = []  # 回饋檔次運算明細，供運算邏輯面板顯示
    days = active_days

    rebate_pct_map = _build_rebate_pct_map(config, total_budget, days)

    # 兩平台皆達標時的舊參數相容
    if rebate_platform_choice in ("全家廣播", "家樂福") and rebate_nat_destination is None:
        rebate_nat_destination = rebate_platform_choice

    # 全省全家達標且業務選「回饋到家樂福」：用全家預算×「家樂福%」換算家樂福檔次（僅當 apply_nat_rad 且選家樂福）
    nat_rad_to_cf_budget = None  # (rebate_budget, rebate_pct, sec, n_days) 或 None
    if apply_nat_rad and rebate_nat_destination == "家樂福" and ("全家廣播", "全省") in rebate_pct_map:
        cfg_rad = config.get("全家廣播")
        if cfg_rad and cfg_rad.get("is_national"):
            m_budget_rad = total_budget * (cfg_rad["share"] / 100.0)
            sec_shares_rad = cfg_rad.get("sec_shares") or {}
            total_nat_rad = sum(m_budget_rad * (pct / 100.0) for _sec, pct in sec_shares_rad.items())
            pct_cf = rebate_pct_map.get(("全家廣播", "全省_家樂福"))
            if pct_cf is None:
                pct_cf = rebate_pct_map.get(("全家廣播", "全省"), 0)
            rebate_budget = total_nat_rad * pct_cf / 100.0
            sec_cf = None
            cfg_cf = config.get("家樂福")
            if cfg_cf and cfg_cf.get("sec_shares"):
                sec_cf = next(iter(cfg_cf["sec_shares"].keys()), None)
            if sec_cf is None:
                for r in rows:
                    if r.get("media") == "家樂福" and r.get("region") == "全省量販":
                        sec_cf = r.get("seconds")
                        break
            if sec_cf is None and sec_shares_rad:
                sec_cf = next(iter(sec_shares_rad.keys()), None)
            if sec_cf is not None and rebate_budget > 0:
                db_cf = pricing_db.get("家樂福", {})
                base_std = db_cf.get("量販_全省", {}).get("Std_Spots", 1)
                base_net = db_cf.get("量販_全省", {}).get("Net", 0)
                if isinstance(base_net, (list, tuple)):
                    base_net = base_net[1] if len(base_net) > 1 else base_net[0]
                factor_cf = get_sec_factor("家樂福", sec_cf, sec_factors)
                unit_net_cf = (base_net / base_std) * factor_cf if base_std else 0
                if unit_net_cf > 0:
                    n_days = active_days
                    nat_rad_to_cf_budget = (rebate_budget, pct_cf, sec_cf, n_days)

    i = 0
    while i < len(rows):
        r = rows[i]
        media = r["media"]
        sec = r["seconds"]
        region = r["region"]
        is_pkg = r.get("is_pkg_member", False)
        base_spots = r.get("spots", 0)
        base_schedule = list(r.get("schedule", []))
        daypart = r.get("daypart", "")
        if base_spots <= 0 and base_schedule:
            base_spots = sum(x for x in base_schedule if isinstance(x, (int, float)) and x > 0)
        n_days = len([x for x in base_schedule if isinstance(x, (int, float)) and x > 0]) or active_days
        if n_days <= 0:
            n_days = active_days

        if media in ["全家廣播", "新鮮視"]:
            cfg = config.get(media)
            if not cfg:
                i += 1
                continue
            m_budget = total_budget * (cfg["share"] / 100.0)
            sec_pct = (cfg.get("sec_shares") or {}).get(sec, 0)
            s_budget = m_budget * (sec_pct / 100.0) if sec_pct else 0

            if cfg.get("is_national") and is_pkg:
                # 全省全家達標：僅在 apply_nat_rad 時產出；若選回饋到家樂福則不產全家回饋列（改由家樂福區塊產出）
                if not apply_nat_rad:
                    group_end = i
                    while group_end + 1 < len(rows) and rows[group_end + 1].get("is_pkg_member") and rows[group_end + 1]["media"] == media and rows[group_end + 1]["seconds"] == sec:
                        group_end += 1
                    i = group_end
                    i += 1
                    continue
                if rebate_nat_destination == "家樂福":
                    group_end = i
                    while group_end + 1 < len(rows) and rows[group_end + 1].get("is_pkg_member") and rows[group_end + 1]["media"] == media and rows[group_end + 1]["seconds"] == sec:
                        group_end += 1
                    i = group_end
                    i += 1
                    continue
                rebate_pct = rebate_pct_map.get((media, "全省"))
                if rebate_pct is not None and base_spots > 0:
                    rebate_spots = max(0, int(round(base_spots * rebate_pct / 100)))
                    if rebate_spots % 2 != 0:
                        rebate_spots += 1
                    if rebate_spots > 0:
                        rebate_sch = calculate_schedule(rebate_spots, n_days)
                        db = pricing_db.get(media, {})
                        std_spots_ref = db.get("Std_Spots", 1)
                        factor = get_sec_factor(media, sec, sec_factors)
                        display_regs = regions_order
                        group_end = i
                        while group_end + 1 < len(rows) and rows[group_end + 1].get("is_pkg_member") and rows[group_end + 1]["media"] == media and rows[group_end + 1]["seconds"] == sec:
                            group_end += 1
                        for rr in display_regs:
                            list_price_region = db.get(rr)
                            if list_price_region is not None:
                                if isinstance(list_price_region, (list, tuple)):
                                    list_price_region = list_price_region[0]
                                unit_rate = int((list_price_region / std_spots_ref) * factor) if std_spots_ref else 0
                                rate_display = unit_rate * rebate_spots
                            else:
                                rate_display = 0
                            rebate_row = {
                                "media": media,
                                "region": rr,
                                "program_num": store_counts_num.get(f"新鮮視_{rr}" if media == "新鮮視" else rr, 0),
                                "daypart": daypart,
                                "seconds": sec,
                                "spots": rebate_spots,
                                "schedule": rebate_sch,
                                "rate_display": rate_display,
                                "pkg_display": "回饋",
                                "is_pkg_member": False,
                                "nat_pkg_display": 0,
                                "is_rebate": True,
                                "rebate_type": "全省",
                                "rebate_pct": rebate_pct,
                            }
                            result.append((group_end, rebate_row))
                        rebate_logs.append({
                            "media": media,
                            "region": "全省",
                            "rebate_type": "全省全家達標→全家",
                            "rebate_pct": rebate_pct,
                            "base_spots": base_spots,
                            "rebate_spots": rebate_spots,
                            "formula": f"基準檔次 × 回饋% = {base_spots} × {rebate_pct}% = {base_spots * rebate_pct / 100:.1f} → 取偶數 → {rebate_spots} 檔",
                        })
                        i = group_end
                elif not is_pkg and region in ("北區", "桃竹苗", "中區", "雲嘉南"):
                    # 單區全家達標：僅在 apply_region_rad 時產出；回饋可顯示在任一區（不必與購買區相同）
                    if not apply_region_rad:
                        i += 1
                        continue
                    # 先以購買區查%；若無（邊界情況）則取任一站有達標的%
                    rebate_pct = rebate_pct_map.get((media, region))
                    if rebate_pct is None:
                        for _r in ("北區", "桃竹苗", "中區", "雲嘉南"):
                            if (media, _r) in rebate_pct_map:
                                rebate_pct = rebate_pct_map[(media, _r)]
                                break
                    if rebate_pct is not None and base_spots > 0:
                        rebate_spots = max(0, int(round(base_spots * rebate_pct / 100)))
                        if rebate_spots % 2 != 0:
                            rebate_spots += 1
                        if rebate_spots == 0 and rebate_pct > 0:
                            rebate_spots = 2  # 至少 2 檔（偶數），確保單區回饋列會產出
                        if rebate_spots > 0:
                            rebate_sch = calculate_schedule(rebate_spots, n_days)
                            db = pricing_db.get(media, {})
                            std_spots_ref = db.get("Std_Spots", 1)
                            factor = get_sec_factor(media, sec, sec_factors)
                            # 回饋可選顯示在 北區/桃竹苗/中區/雲嘉南 任一區，不必與購買區相同
                            display_region = rebate_region_destination if rebate_region_destination in ("北區", "桃竹苗", "中區", "雲嘉南") else region
                            list_price_region = db.get(display_region)
                            if list_price_region is not None:
                                if isinstance(list_price_region, (list, tuple)):
                                    list_price_region = list_price_region[0]
                                unit_rate = int((list_price_region / std_spots_ref) * factor) if std_spots_ref else 0
                                rate_display = unit_rate * rebate_spots
                            else:
                                rate_display = 0
                            rebate_row = {
                                "media": media,
                                "region": display_region,
                                "program_num": store_counts_num.get(f"新鮮視_{display_region}" if media == "新鮮視" else display_region, 0),
                                "daypart": daypart,
                                "seconds": sec,
                                "spots": rebate_spots,
                                "schedule": rebate_sch,
                                "rate_display": rate_display,
                                "pkg_display": "回饋",
                                "is_pkg_member": False,
                                "nat_pkg_display": 0,
                                "is_rebate": True,
                                "rebate_type": "單區",
                                "rebate_pct": rebate_pct,
                            }
                            result.append((i, rebate_row))
                            display_region = rebate_region_destination if rebate_region_destination in ("北區", "桃竹苗", "中區", "雲嘉南") else region
                            rebate_logs.append({
                                "media": media,
                                "region": display_region,
                                "rebate_type": "單區全家達標",
                                "rebate_pct": rebate_pct,
                                "base_spots": base_spots,
                                "rebate_spots": rebate_spots,
                                "formula": f"基準檔次 × 回饋% = {base_spots} × {rebate_pct}% = {base_spots * rebate_pct / 100:.1f} → 取偶數 → {rebate_spots} 檔（顯示於 {display_region}）",
                            })
        elif media == "家樂福":
            cfg = config.get("家樂福")
            if not cfg or region != "全省量販":
                i += 1
                continue
            # 全省全家達標且選「回饋到家樂福」：用全家預算換算的家樂福回饋列，插在第一筆家樂福 row 之後
            if nat_rad_to_cf_budget is not None:
                rebate_budget, pct_nat, sec_cf, nd = nat_rad_to_cf_budget
                db_cf = pricing_db.get("家樂福", {})
                base_std = db_cf.get("量販_全省", {}).get("Std_Spots", 1)
                base_list = db_cf.get("量販_全省", {}).get("List", 0)
                base_net = db_cf.get("量販_全省", {}).get("Net", 0)
                if isinstance(base_net, (list, tuple)):
                    base_net = base_net[1] if len(base_net) > 1 else base_net[0]
                factor_cf = get_sec_factor("家樂福", sec_cf, sec_factors)
                unit_net_cf = (base_net / base_std) * factor_cf if base_std else 0
                if unit_net_cf > 0:
                    rebate_spots = max(0, int(round(rebate_budget / unit_net_cf)))
                    if rebate_spots % 2 != 0:
                        rebate_spots += 1
                    if rebate_spots > 0:
                        rebate_sch = calculate_schedule(rebate_spots, nd)
                        unit_rate = int((base_list / base_std) * factor_cf) if base_std else 0
                        rate_display = unit_rate * rebate_spots
                        cf_rebate_row = {
                            "media": "家樂福",
                            "region": "全省量販",
                            "program_num": store_counts_num.get("家樂福_量販", 0),
                            "daypart": db_cf.get("量販_全省", {}).get("Day_Part", ""),
                            "seconds": sec_cf,
                            "spots": rebate_spots,
                            "schedule": rebate_sch,
                            "rate_display": rate_display,
                            "pkg_display": "回饋",
                            "is_pkg_member": False,
                            "nat_pkg_display": 0,
                            "is_rebate": True,
                            "rebate_type": "全省",
                            "rebate_pct": pct_nat,
                            "rebate_from_nat_rad": True,
                        }
                        result.append((i, cf_rebate_row))
                        rebate_logs.append({
                            "media": "家樂福",
                            "region": "全省量販",
                            "rebate_type": "全省全家達標→家樂福",
                            "rebate_pct": pct_nat,
                            "rebate_budget": rebate_budget,
                            "unit_net_cf": unit_net_cf,
                            "rebate_spots": rebate_spots,
                            "formula": f"全省全家預算 × 家樂福% = 回饋預算 {rebate_budget:,.0f}；回饋預算 ÷ 家樂福單檔成本 {unit_net_cf:.2f} = {rebate_budget / unit_net_cf:.1f} → 取偶數 → {rebate_spots} 檔",
                        })
                nat_rad_to_cf_budget = None  # 只產一列
            # 家樂福達標回饋：僅在 apply_nat_cf 時產出（可與「全家預算→家樂福」併存）
            rebate_pct = rebate_pct_map.get(("家樂福", "全省"))
            if apply_nat_cf and rebate_pct is not None and base_spots > 0:
                rebate_spots = max(0, int(round(base_spots * rebate_pct / 100)))
                if rebate_spots % 2 != 0:
                    rebate_spots += 1
                if rebate_spots > 0:
                    rebate_sch = calculate_schedule(rebate_spots, n_days)
                    db = pricing_db.get("家樂福", {})
                    base_std = db.get("量販_全省", {}).get("Std_Spots", 1)
                    base_list = db.get("量販_全省", {}).get("List", 0)
                    factor = get_sec_factor("家樂福", sec, sec_factors)
                    unit_rate = int((base_list / base_std) * factor) if base_std else 0
                    rate_display = unit_rate * rebate_spots
                    rebate_row = {
                        "media": "家樂福",
                        "region": "全省量販",
                        "program_num": store_counts_num.get("家樂福_量販", 0),
                        "daypart": db.get("量販_全省", {}).get("Day_Part", ""),
                        "seconds": sec,
                        "spots": rebate_spots,
                        "schedule": rebate_sch,
                        "rate_display": rate_display,
                        "pkg_display": "回饋",
                        "is_pkg_member": False,
                        "nat_pkg_display": 0,
                        "is_rebate": True,
                        "rebate_type": "全省",
                        "rebate_pct": rebate_pct,
                    }
                    result.append((i, rebate_row))
                    rebate_logs.append({
                        "media": "家樂福",
                        "region": "全省量販",
                        "rebate_type": "家樂福達標",
                        "rebate_pct": rebate_pct,
                        "base_spots": base_spots,
                        "rebate_spots": rebate_spots,
                        "formula": f"基準檔次 × 回饋% = {base_spots} × {rebate_pct}% = {base_spots * rebate_pct / 100:.1f} → 取偶數 → {rebate_spots} 檔",
                    })
        i += 1

    # 後備：若應套用單區回饋但迴圈內未產出任何單區列，則強制產出一列（純單區或全省加重皆適用）
    if apply_region_rad and rebate_region_destination in ("北區", "桃竹苗", "中區", "雲嘉南"):
        has_region_rebate = any(entry[1].get("rebate_type") == "單區" for entry in result)
        if not has_region_rebate:
            cfg_rad = config.get("全家廣播")
            # 單區達標在 map 內即可（純單區 is_national=False，或全省加重 is_national=True 且有 region_shares 致單區達標）
            has_region_in_map = any(("全家廣播", _r) in rebate_pct_map for _r in ("北區", "桃竹苗", "中區", "雲嘉南"))
            if cfg_rad and has_region_in_map:
                rebate_pct = None
                for _r in ("北區", "桃竹苗", "中區", "雲嘉南"):
                    if ("全家廣播", _r) in rebate_pct_map:
                        rebate_pct = rebate_pct_map[("全家廣播", _r)]
                        break
                if rebate_pct is not None:
                    # 找最後一筆全家廣播 row 作為插入位置與參考（單區時為單區 row，全省加重時為全省 pkg row）
                    insert_after = None
                    ref_row = None
                    for idx, row in enumerate(rows):
                        if row.get("media") == "全家廣播":
                            insert_after = idx
                            ref_row = row
                    if insert_after is not None and ref_row is not None:
                        base_schedule = list(ref_row.get("schedule", []))
                        base_spots = ref_row.get("spots", 0) or sum(x for x in base_schedule if isinstance(x, (int, float)) and x > 0)
                        n_days = len([x for x in base_schedule if isinstance(x, (int, float)) and x > 0]) or active_days
                        if n_days <= 0:
                            n_days = active_days
                        if base_spots <= 0:
                            base_spots = 2
                        rebate_spots = max(2, int(round(base_spots * rebate_pct / 100)))
                        if rebate_spots % 2 != 0:
                            rebate_spots += 1
                        rebate_sch = calculate_schedule(rebate_spots, n_days)
                        media = ref_row["media"]
                        sec = ref_row["seconds"]
                        daypart = ref_row.get("daypart", "")
                        display_region = rebate_region_destination
                        db = pricing_db.get(media, {})
                        std_spots_ref = db.get("Std_Spots", 1)
                        factor = get_sec_factor(media, sec, sec_factors)
                        list_price_region = db.get(display_region)
                        if list_price_region is not None:
                            if isinstance(list_price_region, (list, tuple)):
                                list_price_region = list_price_region[0]
                            unit_rate = int((list_price_region / std_spots_ref) * factor) if std_spots_ref else 0
                            rate_display = unit_rate * rebate_spots
                        else:
                            rate_display = 0
                        rebate_row = {
                            "media": media,
                            "region": display_region,
                            "program_num": store_counts_num.get(f"新鮮視_{display_region}" if media == "新鮮視" else display_region, 0),
                            "daypart": daypart,
                            "seconds": sec,
                            "spots": rebate_spots,
                            "schedule": rebate_sch,
                            "rate_display": rate_display,
                            "pkg_display": "回饋",
                            "is_pkg_member": False,
                            "nat_pkg_display": 0,
                            "is_rebate": True,
                            "rebate_type": "單區",
                            "rebate_pct": rebate_pct,
                        }
                        result.append((insert_after, rebate_row))
                        rebate_logs.append({
                            "media": media,
                            "region": display_region,
                            "rebate_type": "單區全家達標（後備）",
                            "rebate_pct": rebate_pct,
                            "base_spots": base_spots,
                            "rebate_spots": rebate_spots,
                            "formula": f"基準檔次 × 回饋% = {base_spots} × {rebate_pct}% → 取偶數 → {rebate_spots} 檔（顯示於 {display_region}）",
                        })

    # 僅買全家廣播且選回饋家樂福：rows 裡沒有家樂福，上面迴圈不會消費 nat_rad_to_cf_budget，在此產出並插在最後
    if nat_rad_to_cf_budget is not None:
        rebate_budget, pct_nat, sec_cf, nd = nat_rad_to_cf_budget
        db_cf = pricing_db.get("家樂福", {})
        base_std = db_cf.get("量販_全省", {}).get("Std_Spots", 1)
        base_list = db_cf.get("量販_全省", {}).get("List", 0)
        base_net = db_cf.get("量販_全省", {}).get("Net", 0)
        if isinstance(base_net, (list, tuple)):
            base_net = base_net[1] if len(base_net) > 1 else base_net[0]
        factor_cf = get_sec_factor("家樂福", sec_cf, sec_factors)
        unit_net_cf = (base_net / base_std) * factor_cf if base_std else 0
        if unit_net_cf > 0:
            rebate_spots = max(0, int(round(rebate_budget / unit_net_cf)))
            if rebate_spots % 2 != 0:
                rebate_spots += 1
            if rebate_spots > 0:
                rebate_sch = calculate_schedule(rebate_spots, nd)
                unit_rate = int((base_list / base_std) * factor_cf) if base_std else 0
                rate_display = unit_rate * rebate_spots
                cf_rebate_row = {
                    "media": "家樂福",
                    "region": "全省量販",
                    "program_num": store_counts_num.get("家樂福_量販", 0),
                    "daypart": db_cf.get("量販_全省", {}).get("Day_Part", ""),
                    "seconds": sec_cf,
                    "spots": rebate_spots,
                    "schedule": rebate_sch,
                    "rate_display": rate_display,
                    "pkg_display": "回饋",
                    "is_pkg_member": False,
                    "nat_pkg_display": 0,
                    "is_rebate": True,
                    "rebate_type": "全省",
                    "rebate_pct": pct_nat,
                    "rebate_from_nat_rad": True,
                }
                result.append((len(rows) - 1, cf_rebate_row))
                rebate_logs.append({
                    "media": "家樂福",
                    "region": "全省量販",
                    "rebate_type": "全省全家達標→家樂福（僅買全家）",
                    "rebate_pct": pct_nat,
                    "rebate_budget": rebate_budget,
                    "unit_net_cf": unit_net_cf,
                    "rebate_spots": rebate_spots,
                    "formula": f"全省全家預算 × 家樂福% = 回饋預算 {rebate_budget:,.0f}；回饋預算 ÷ 家樂福單檔成本 {unit_net_cf:.2f} = {rebate_budget / unit_net_cf:.1f} → 取偶數 → {rebate_spots} 檔",
                })

    return (result, rebate_logs)


def compute_bonus_rebate_rows(config, total_budget, active_days, rows, bonus_pct, pricing_db, sec_factors, store_counts_num, regions_order):
    """
    主管加贈回饋：不論全省或分區，一律以 bonus_pct % 依實作價換算檔次，產出額外回饋列。
    回傳格式同 compute_rebate_rows：(insert_after_index, row) 列表。
    與門檻回饋獨立，可並存。
    """
    if not bonus_pct or bonus_pct <= 0:
        return []
    result = []
    days = active_days

    i = 0
    while i < len(rows):
        r = rows[i]
        media = r["media"]
        sec = r["seconds"]
        region = r["region"]
        is_pkg = r.get("is_pkg_member", False)
        base_spots = r.get("spots", 0)
        base_schedule = list(r.get("schedule", []))
        daypart = r.get("daypart", "")
        if not daypart and media in ["全家廣播", "新鮮視"]:
            daypart = pricing_db.get(media, {}).get("Day_Part", "")
        n_days = len([x for x in base_schedule if isinstance(x, (int, float)) and x > 0]) or active_days
        if n_days <= 0:
            n_days = active_days

        if media in ["全家廣播", "新鮮視"]:
            cfg = config.get(media)
            if not cfg:
                i += 1
                continue
            if cfg.get("is_national") and is_pkg:
                # 全省：六列額外回饋
                if base_spots > 0:
                    rebate_spots = max(0, int(round(base_spots * bonus_pct / 100)))
                    if rebate_spots % 2 != 0:
                        rebate_spots += 1
                    if rebate_spots > 0:
                        rebate_sch = calculate_schedule(rebate_spots, n_days)
                        db = pricing_db.get(media, {})
                        std_spots_ref = db.get("Std_Spots", 1)
                        factor = get_sec_factor(media, sec, sec_factors)
                        display_regs = regions_order
                        group_end = i
                        while group_end + 1 < len(rows) and rows[group_end + 1].get("is_pkg_member") and rows[group_end + 1]["media"] == media and rows[group_end + 1]["seconds"] == sec:
                            group_end += 1
                        for rr in display_regs:
                            list_price_region = db.get(rr)
                            if list_price_region is not None:
                                if isinstance(list_price_region, (list, tuple)):
                                    list_price_region = list_price_region[0]
                                unit_rate = int((list_price_region / std_spots_ref) * factor) if std_spots_ref else 0
                                rate_display = unit_rate * rebate_spots
                            else:
                                rate_display = 0
                            bonus_row = {
                                "media": media,
                                "region": rr,
                                "program_num": store_counts_num.get(f"新鮮視_{rr}" if media == "新鮮視" else rr, 0),
                                "daypart": daypart,
                                "seconds": sec,
                                "spots": rebate_spots,
                                "schedule": rebate_sch,
                                "rate_display": rate_display,
                                "pkg_display": "額外回饋",
                                "is_pkg_member": False,
                                "nat_pkg_display": 0,
                                "is_rebate": True,
                                "rebate_type": "全省",
                                "rebate_pct": bonus_pct,
                                "is_bonus_rebate": True,
                            }
                            result.append((group_end, bonus_row))
                        i = group_end
            elif not is_pkg and region in ("北區", "桃竹苗", "中區", "雲嘉南"):
                # 單區：一列額外回饋
                if base_spots > 0:
                    rebate_spots = max(0, int(round(base_spots * bonus_pct / 100)))
                    if rebate_spots % 2 != 0:
                        rebate_spots += 1
                    if rebate_spots > 0:
                        rebate_sch = calculate_schedule(rebate_spots, n_days)
                        db = pricing_db.get(media, {})
                        std_spots_ref = db.get("Std_Spots", 1)
                        factor = get_sec_factor(media, sec, sec_factors)
                        list_price_region = db.get(region)
                        if list_price_region is not None:
                            if isinstance(list_price_region, (list, tuple)):
                                list_price_region = list_price_region[0]
                            unit_rate = int((list_price_region / std_spots_ref) * factor) if std_spots_ref else 0
                            rate_display = unit_rate * rebate_spots
                        else:
                            rate_display = 0
                        bonus_row = {
                            "media": media,
                            "region": region,
                            "program_num": store_counts_num.get(f"新鮮視_{region}" if media == "新鮮視" else region, 0),
                            "daypart": daypart,
                            "seconds": sec,
                            "spots": rebate_spots,
                            "schedule": rebate_sch,
                            "rate_display": rate_display,
                            "pkg_display": "額外回饋",
                            "is_pkg_member": False,
                            "nat_pkg_display": 0,
                            "is_rebate": True,
                            "rebate_type": "單區",
                            "rebate_pct": bonus_pct,
                            "is_bonus_rebate": True,
                        }
                        result.append((i, bonus_row))
        elif media == "家樂福":
            cfg = config.get("家樂福")
            if not cfg or region != "全省量販":
                i += 1
                continue
            if base_spots > 0:
                rebate_spots = max(0, int(round(base_spots * bonus_pct / 100)))
                if rebate_spots % 2 != 0:
                    rebate_spots += 1
                if rebate_spots > 0:
                    rebate_sch = calculate_schedule(rebate_spots, n_days)
                    db = pricing_db.get("家樂福", {})
                    base_std = db.get("量販_全省", {}).get("Std_Spots", 1)
                    base_list = db.get("量販_全省", {}).get("List", 0)
                    factor = get_sec_factor("家樂福", sec, sec_factors)
                    unit_rate = int((base_list / base_std) * factor) if base_std else 0
                    rate_display = unit_rate * rebate_spots
                    bonus_row = {
                        "media": "家樂福",
                        "region": "全省量販",
                        "program_num": store_counts_num.get("家樂福_量販", 0),
                        "daypart": db.get("量販_全省", {}).get("Day_Part", ""),
                        "seconds": sec,
                        "spots": rebate_spots,
                        "schedule": rebate_sch,
                        "rate_display": rate_display,
                        "pkg_display": "額外回饋",
                        "is_pkg_member": False,
                        "nat_pkg_display": 0,
                        "is_rebate": True,
                        "rebate_type": "全省",
                        "rebate_pct": bonus_pct,
                        "is_bonus_rebate": True,
                    }
                    result.append((i, bonus_row))
        i += 1

    return result


def _net_price_from_db_entry(db_entry):
    """從定價表項目取得實作價 (Net)，用於主管回饋預算換算檔次。"""
    if db_entry is None:
        return 0
    if isinstance(db_entry, (list, tuple)):
        return db_entry[1] if len(db_entry) > 1 else (db_entry[0] if db_entry else 0)
    return db_entry.get("Net", db_entry.get("List", 0))


def compute_bonus_rebate_rows_from_allocation(bonus_config, rebate_budget, active_days, rows, pricing_db, sec_factors, store_counts_num, regions_order):
    """
    主管回饋（客製化分配）：依 rebate_budget 與 bonus_config 分配至各平台／區域／秒數，
    以實作價換算檔次，產出額外回饋列並插入 CUE 表。
    bonus_config 格式與主 config 類似但無 region_shares（無自訂區域比例）：
      {"全家廣播": {"is_national": True, "regions": ["全省"], "sec_shares": {15: 50, 20: 50}, "share": 40}, ...}
    回傳 (inserts, bonus_rebate_logs)：inserts 為 (insert_after_index, row) 列表；
    bonus_rebate_logs 為運算邏輯明細，供 render_logic_panel 顯示。
    """
    if not bonus_config or rebate_budget <= 0:
        return [], []
    result = []
    bonus_rebate_logs = []
    days = active_days

    def _last_idx_for_media_sec(media, sec, is_national):
        """找該 media+sec 在 rows 中最後一列的 index（全省為該 media+sec 的 package 組最後一列）。"""
        cand = -1
        for i, r in enumerate(rows):
            if r.get("is_rebate"):
                continue
            if r["media"] != media or r["seconds"] != sec:
                continue
            if is_national and r.get("is_pkg_member"):
                cand = i
            elif not is_national and not r.get("is_pkg_member") and r.get("region") in regions_order:
                cand = i
            else:
                cand = i
        return cand

    def _last_idx_cf(sec):
        for i in range(len(rows) - 1, -1, -1):
            r = rows[i]
            if r.get("is_rebate"):
                continue
            if r["media"] == "家樂福" and r["seconds"] == sec:
                return i
        return len(rows) - 1

    def _daypart_from_rows(media, sec=None):
        for r in rows:
            if r.get("is_rebate"):
                continue
            if r["media"] == media and (sec is None or r["seconds"] == sec):
                return r.get("daypart", "")
        return ""

    for media, cfg in bonus_config.items():
        if not cfg:
            continue
        platform_share = cfg.get("share", 0) or 0
        if platform_share <= 0:
            continue
        platform_budget = rebate_budget * (platform_share / 100.0)
        sec_shares = cfg.get("sec_shares") or {}
        if not sec_shares:
            continue
        is_national = cfg.get("is_national", True)
        regions = cfg.get("regions", ["全省"])
        if is_national or "全省" in regions:
            display_regions = list(regions_order)
        else:
            display_regions = [r for r in regions if r in regions_order]
        if not display_regions:
            display_regions = list(regions_order)

        if media == "家樂福":
            db = pricing_db.get("家樂福", {})
            base_info = db.get("量販_全省", {}) or {}
            base_std = base_info.get("Std_Spots", 1)
            base_net = base_info.get("Net", base_info.get("List", 0))
            if isinstance(base_net, (list, tuple)):
                base_net = base_net[1] if len(base_net) > 1 else base_net[0]
            daypart = base_info.get("Day_Part", "")
            for sec, sec_pct in sec_shares.items():
                if sec_pct <= 0:
                    continue
                sec_budget = platform_budget * (sec_pct / 100.0)
                factor = get_sec_factor("家樂福", sec, sec_factors)
                unit_rate = int((base_net / base_std) * factor) if base_std else 0
                if unit_rate <= 0:
                    continue
                spots = max(0, int(round(sec_budget / unit_rate)))
                if spots % 2 != 0:
                    spots += 1
                if spots <= 0:
                    continue
                unit_cost_net = (base_net / base_std) * factor if base_std else 0
                bonus_rebate_logs.append({
                    "media": "家樂福",
                    "region": "全省量販",
                    "seconds": sec,
                    "rebate_budget_total": rebate_budget,
                    "platform_share_pct": platform_share,
                    "platform_budget": platform_budget,
                    "sec_share_pct": sec_pct,
                    "sec_budget": sec_budget,
                    "unit_cost_net": unit_cost_net,
                    "spots": spots,
                    "formula": f"主管回饋金額 ${rebate_budget:,.0f} × 家樂福 {platform_share}% = ${platform_budget:,.0f}；× {sec}秒 {sec_pct}% = ${sec_budget:,.0f}；÷ 單檔成本 ${unit_cost_net:,.2f} = {sec_budget/unit_cost_net:.2f} → 取偶數 → {spots} 檔",
                })
                rebate_sch = calculate_schedule(spots, days)
                rate_display = unit_rate * spots
                bonus_row = {
                    "media": "家樂福",
                    "region": "全省量販",
                    "program_num": store_counts_num.get("家樂福_量販", 0),
                    "daypart": daypart,
                    "seconds": sec,
                    "spots": spots,
                    "schedule": rebate_sch,
                    "rate_display": rate_display,
                    "pkg_display": "額外回饋",
                    "is_pkg_member": False,
                    "nat_pkg_display": 0,
                    "is_rebate": True,
                    "rebate_type": "全省",
                    "rebate_pct": None,
                    "is_bonus_rebate": True,
                }
                idx = _last_idx_cf(sec)
                result.append((idx, bonus_row))
            continue

        db = pricing_db.get(media, {})
        if not db:
            continue
        std_spots_ref = db.get("Std_Spots", 1)
        # 額外回饋 day-part 比照一般列：依平台從定價表取 Day_Part，沒有再從既有 rows 取
        daypart = db.get("Day_Part", "") or _daypart_from_rows(media)

        for sec, sec_pct in sec_shares.items():
            if sec_pct <= 0:
                continue
            sec_budget = platform_budget * (sec_pct / 100.0)
            factor = get_sec_factor(media, sec, sec_factors)

            if is_national:
                list_price_region = db.get("全省")
                net_val = _net_price_from_db_entry(list_price_region)
                if std_spots_ref:
                    unit_rate = int((net_val / std_spots_ref) * factor)
                else:
                    unit_rate = 0
                if unit_rate <= 0:
                    continue
                spots = max(0, int(round(sec_budget / unit_rate)))
                if spots % 2 != 0:
                    spots += 1
                if spots <= 0:
                    continue
                rebate_sch = calculate_schedule(spots, days)
                unit_cost_net = (net_val / std_spots_ref) * factor if std_spots_ref else 0
                for rr in display_regions:
                    if rr == display_regions[0]:
                        bonus_rebate_logs.append({
                            "media": media,
                            "region": "全省（6區各 {} 檔）".format(spots),
                            "seconds": sec,
                            "rebate_budget_total": rebate_budget,
                            "platform_share_pct": platform_share,
                            "platform_budget": platform_budget,
                            "sec_share_pct": sec_pct,
                            "sec_budget": sec_budget,
                            "unit_cost_net": unit_cost_net,
                            "spots": spots,
                            "formula": f"主管回饋金額 ${rebate_budget:,.0f} × {media} {platform_share}% = ${platform_budget:,.0f}；× {sec}秒 {sec_pct}% = ${sec_budget:,.0f}；÷ 單檔成本(實作價) ${unit_cost_net:,.2f} = {sec_budget/unit_cost_net:.2f} → 取偶數 → {spots} 檔（全省 6 區各 {spots} 檔）",
                        })
                    list_price_r = db.get(rr)
                    if list_price_r is not None:
                        if isinstance(list_price_r, (list, tuple)):
                            list_price_r = list_price_r[0]
                        rate_display = int((list_price_r / std_spots_ref) * factor) * spots if std_spots_ref else 0
                    else:
                        rate_display = 0
                    bonus_row = {
                        "media": media,
                        "region": rr,
                        "program_num": store_counts_num.get(f"新鮮視_{rr}" if media == "新鮮視" else rr, 0),
                        "daypart": daypart,
                        "seconds": sec,
                        "spots": spots,
                        "schedule": rebate_sch,
                        "rate_display": rate_display,
                        "pkg_display": "額外回饋",
                        "is_pkg_member": False,
                        "nat_pkg_display": 0,
                        "is_rebate": True,
                        "rebate_type": "全省",
                        "rebate_pct": None,
                        "is_bonus_rebate": True,
                    }
                    idx = _last_idx_for_media_sec(media, sec, True)
                    if idx < 0:
                        idx = len(rows) - 1
                    result.append((idx, bonus_row))
            else:
                for rr in display_regions:
                    list_price_r = db.get(rr)
                    net_val = _net_price_from_db_entry(list_price_r)
                    if std_spots_ref:
                        unit_rate = int((net_val / std_spots_ref) * factor)
                    else:
                        unit_rate = 0
                    if unit_rate <= 0:
                        continue
                    region_budget = sec_budget / len(display_regions)
                    spots = max(0, int(round(region_budget / unit_rate)))
                    if spots % 2 != 0:
                        spots += 1
                    if spots <= 0:
                        continue
                    rebate_sch = calculate_schedule(spots, days)
                    unit_cost_net = (net_val / std_spots_ref) * factor if std_spots_ref else 0
                    bonus_rebate_logs.append({
                        "media": media,
                        "region": rr,
                        "seconds": sec,
                        "rebate_budget_total": rebate_budget,
                        "platform_share_pct": platform_share,
                        "platform_budget": platform_budget,
                        "sec_share_pct": sec_pct,
                        "sec_budget": sec_budget,
                        "region_budget": region_budget,
                        "unit_cost_net": unit_cost_net,
                        "spots": spots,
                        "formula": f"主管回饋金額 ${rebate_budget:,.0f} × {media} {platform_share}% = ${platform_budget:,.0f}；× {sec}秒 {sec_pct}% = ${sec_budget:,.0f}；÷ {len(display_regions)} 區 = ${region_budget:,.0f}／區；${region_budget:,.0f} ÷ 單檔成本 ${unit_cost_net:,.2f} = {region_budget/unit_cost_net:.2f} → 取偶數 → {spots} 檔（{rr}）",
                    })
                    lp = list_price_r
                    if lp is not None:
                        lp = lp[0] if isinstance(lp, (list, tuple)) else lp
                        rate_display = int((lp / std_spots_ref) * factor) * spots if std_spots_ref else 0
                    else:
                        rate_display = 0
                    bonus_row = {
                        "media": media,
                        "region": rr,
                        "program_num": store_counts_num.get(f"新鮮視_{rr}" if media == "新鮮視" else rr, 0),
                        "daypart": daypart,
                        "seconds": sec,
                        "spots": spots,
                        "schedule": rebate_sch,
                        "rate_display": rate_display,
                        "pkg_display": "額外回饋",
                        "is_pkg_member": False,
                        "nat_pkg_display": 0,
                        "is_rebate": True,
                        "rebate_type": "單區",
                        "rebate_pct": None,
                        "is_bonus_rebate": True,
                    }
                    idx = _last_idx_for_media_sec(media, sec, False)
                    if idx < 0:
                        idx = len(rows) - 1
                    result.append((idx, bonus_row))

    return (result, bonus_rebate_logs)


def get_row_groups(rows, regions_order):
    """
    從 rows 取得「同平台、同秒數、同區域」群組，供業務加贈檔次 UI 使用。
    回傳: list of dict {
        "key": (media, seconds, region_key),
        "last_index": int,
        "first_row": row_dict,
        "display_regions": list of str (該群組要顯示的區域列，全省為 6 區、單區/家樂福為 1 區),
    }
    """
    groups = []
    i = 0
    while i < len(rows):
        r = rows[i]
        media = r.get("media", "")
        sec = r.get("seconds", 0)
        if r.get("is_pkg_member"):
            # 全省：同一 media+sec 的連續 is_pkg_member 為一組
            j = i
            while j < len(rows) and rows[j].get("is_pkg_member") and rows[j].get("media") == media and rows[j].get("seconds") == sec:
                j += 1
            display_regions = [rows[k].get("region") for k in range(i, j)]
            groups.append({
                "key": (media, sec, "全省"),
                "last_index": j - 1,
                "first_row": rows[i],
                "display_regions": display_regions or regions_order[:6],
            })
            i = j
        else:
            groups.append({
                "key": (media, sec, r.get("region", "")),
                "last_index": i,
                "first_row": r,
                "display_regions": [r.get("region", "")],
            })
            i += 1
    return groups


def compute_custom_bonus_rows(rows, custom_bonus_config, campaign_start, campaign_end, pricing_db, sec_factors, store_counts_num, regions_order):
    """
    依業務設定的加贈檔次（每組可選：是否加贈、每日檔次、加贈日期區間）產出要插入的列。
    custom_bonus_config: dict, key = (media, seconds, region_key) -> {
        "enabled": bool,
        "spots_per_day": int,
        "date_start": date,
        "date_end": date,
    }
    campaign_start, campaign_end: 走期起迄 (date)。
    回傳: (inserts, [])，inserts 為 (insert_after_index, row) 列表。
    """
    from datetime import timedelta
    result = []
    groups = get_row_groups(rows, regions_order)
    days_count = (campaign_end - campaign_start).days + 1
    for g in groups:
        key = g["key"]
        cfg = custom_bonus_config.get(key)
        if not cfg or not cfg.get("enabled") or not cfg.get("spots_per_day"):
            continue
        spots_per_day = max(0, int(cfg.get("spots_per_day", 0)))
        b_start = cfg.get("date_start")
        b_end = cfg.get("date_end")
        if not b_start or not b_end:
            continue
        # 限制在走期內
        b_start = max(b_start, campaign_start)
        b_end = min(b_end, campaign_end)
        if b_start > b_end:
            continue
        # 建 schedule：僅在 [b_start, b_end] 內有檔次
        schedule = []
        for d in range(days_count):
            day_date = campaign_start + timedelta(days=d)
            schedule.append(spots_per_day if b_start <= day_date <= b_end else 0)
        total_spots = sum(schedule)
        if total_spots == 0:
            continue
        first = g["first_row"]
        media, sec, region_key = key
        daypart = first.get("daypart", "")
        display_regions = g["display_regions"]
        # 定價用於 rate_display；加贈列 Package 顯示「加贈」；Std_Spots 缺值時以 REFERENCE_STD_SPOTS 為 fallback
        db = pricing_db.get(media, {})
        factor = get_sec_factor(media, sec, sec_factors)
        cf_region_key = {"全省量販": "量販_全省", "全省超市": "超市_全省"}
        for disp_r in display_regions:
            list_price = None
            std_spots_ref = None
            if isinstance(db, dict):
                if media == "家樂福":
                    db_key = cf_region_key.get(disp_r, disp_r)
                    ent = db.get(db_key)
                    if isinstance(ent, dict):
                        list_price = ent.get("List")
                        std_spots_ref = ent.get("Std_Spots")
                    if std_spots_ref is None:
                        std_spots_ref = REFERENCE_STD_SPOTS.get("家樂福", {}).get(disp_r) or REFERENCE_STD_SPOTS.get("家樂福", {}).get(db_key)
                else:
                    ent = db.get(disp_r) or db.get("全省")
                    if isinstance(ent, (list, tuple)):
                        list_price = ent[0] if len(ent) else None
                    else:
                        list_price = ent
                    region_std = db.get("_Region_Std_Spots") or {}
                    std_spots_ref = region_std.get(disp_r) or db.get("Std_Spots") or std_spots_ref
                    if std_spots_ref is None:
                        std_spots_ref = REFERENCE_STD_SPOTS.get(media, {}).get(disp_r) or REFERENCE_STD_SPOTS.get(media, {}).get("全省")
            if std_spots_ref is None:
                std_spots_ref = 480  # 未知媒體最後防呆
            unit_rate = int((list_price / std_spots_ref) * factor) if list_price and std_spots_ref else 0
            rate_display = unit_rate * total_spots
            bonus_row = {
                "media": media,
                "region": disp_r,
                "program_num": "加贈檔次",
                "daypart": daypart,
                "seconds": sec,
                "spots": total_spots,
                "schedule": schedule[:],
                "rate_display": rate_display,
                "pkg_display": "加贈",
                "is_pkg_member": False,
                "nat_pkg_display": 0,
                "is_rebate": False,
                "is_custom_bonus": True,
            }
            result.append((g["last_index"], bonus_row))
    # 同 last_index 可能多列（全省 6 區），已依 display_regions 逐一 append
    return (result, [])


def merge_rebate_into_rows(rows, rebate_inserts):
    """將回饋列依 insert_after_index 由小到大插入（同一 index 可多列）。"""
    if not rebate_inserts:
        return rows
    by_idx = {}
    for idx, r in rebate_inserts:
        by_idx.setdefault(idx, []).append(r)
    out = []
    for i, r in enumerate(rows):
        out.append(r)
        if i in by_idx:
            out.extend(by_idx[i])
    return out


def get_rebate_summary_text(rebate_inserts, config=None, total_budget=None, active_days=None, rebate_nat_destination=None, qual=None, apply_nat_rad=None, apply_nat_cf=None, apply_region_rad=None, rebate_region_destination=None):
    """
    從 rebate_inserts 與（可選）達標資訊彙總「本次可回饋」文字。
    qual / apply_* / rebate_* 為選填，供 UI 顯示達標項目與套用狀態。
    回傳例如：「全家廣播 全省 10%（可選回饋至全省家樂福 15%）；家樂福 全省 8%；全家廣播 桃竹苗 8%」
    """
    m = None
    if config is not None and total_budget is not None and active_days is not None:
        m = _build_rebate_pct_map(config, total_budget, active_days)

    if not rebate_inserts:
        # 無套用回饋時：若有達標資訊則列出可回饋項目
        if m is None:
            return ""
        parts = []
        if ("全家廣播", "全省") in m:
            pct_rad = m[("全家廣播", "全省")]
            pct_cf = m.get(("全家廣播", "全省_家樂福"), pct_rad)
            parts.append(f"全家廣播 全省 {pct_rad}%（可選回饋至全省家樂福 {pct_cf}%）")
        if ("家樂福", "全省") in m:
            parts.append(f"家樂福 全省 {m[('家樂福', '全省')]}%")
        for r in ("北區", "桃竹苗", "中區", "雲嘉南"):
            if ("全家廣播", r) in m:
                parts.append(f"全家廣播 {r} {m[('全家廣播', r)]}%")
        return "；".join(parts) if parts else ""

    seen = {}
    for _idx, r in rebate_inserts:
        media = r.get("media", "")
        pct = r.get("rebate_pct")
        if pct is None:
            continue
        if r.get("rebate_type") == "單區":
            scope_key = f"{media}|單區|{r.get('region', '')}"
            label = f"{media} {r.get('region', '')} {pct}%"
        else:
            scope_key = f"{media}|全省"
            label = f"{media} 全省 {pct}%"
        if scope_key not in seen or seen[scope_key][1] < pct:
            seen[scope_key] = (label, pct)

    nat_rad_pct = None
    nat_rad_cf_pct = None
    if m is not None:
        if ("全家廣播", "全省") in m:
            nat_rad_pct = m[("全家廣播", "全省")]
        if ("全家廣播", "全省_家樂福") in m:
            nat_rad_cf_pct = m[("全家廣播", "全省_家樂福")]
        if nat_rad_pct is not None and nat_rad_cf_pct is None:
            cfg_rad = config.get("全家廣播")
            if cfg_rad and cfg_rad.get("sec_shares"):
                m_budget_rad = total_budget * (cfg_rad["share"] / 100.0)
                total_nat_rad = sum(m_budget_rad * (pct / 100.0) for _sec, pct in cfg_rad["sec_shares"].items())
                pair = _best_rebate_nat_rad(total_nat_rad, active_days)
                if pair is not None:
                    nat_rad_cf_pct = pair[1]

    parts = []
    for scope_key in sorted(seen.keys()):
        label, pct = seen[scope_key]
        if scope_key == "全家廣播|全省" and nat_rad_pct is not None:
            cf_pct = nat_rad_cf_pct if nat_rad_cf_pct is not None else pct
            label = f"{label}（可選回饋至全省家樂福 {cf_pct}%）"
        parts.append(label)
    if nat_rad_pct is not None:
        if "全家廣播|全省" not in seen:
            cf_pct = nat_rad_cf_pct if nat_rad_cf_pct is not None else nat_rad_pct
            parts.insert(0, f"全家廣播 全省 {nat_rad_pct}%（可選回饋至全省家樂福 {cf_pct}%）")
        elif "家樂福|全省" not in seen and nat_rad_cf_pct is not None:
            parts.append(f"全省家樂福 {nat_rad_cf_pct}%（全家達標可選）")
    return "；".join(parts)
