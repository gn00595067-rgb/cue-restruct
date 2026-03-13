"""
年約／季約細 CUE 模組 (Annual / Quarterly Contract -> Fixed-Wave CUE)
將已知總檔次與實收金額分配至各波段走期，每波段產出獨立 Excel/PDF。
"""

import math
from datetime import datetime, timedelta
from utils import calculate_schedule, get_sec_factor


def round_to_even(x):
    """將數值進位為最接近的偶數整數；小數或奇數時進位。"""
    n = int(math.ceil(float(x)))
    if n % 2 != 0:
        n += 1
    return n


def distribute_by_wave_days(total_spots, total_net, waves):
    """
    依各波段天數比例，將總檔次與總實收分配至各波。
    waves: list of (start_date, end_date) (date 或 datetime)
    回傳: list of (spots, net) 每個波一組，spots/net 為偶數/整數。
    """
    if not waves:
        return []
    total_days = 0
    days_per_wave = []
    for s, e in waves:
        d = (e - s).days + 1
        total_days += d
        days_per_wave.append(d)
    if total_days <= 0:
        return [(0, 0)] * len(waves)
    result = []
    spots_used = 0
    net_used = 0
    for i, days in enumerate(days_per_wave):
        ratio = days / total_days
        if i < len(waves) - 1:
            sp = round_to_even(total_spots * ratio)
            nt = int(round(total_net * ratio))
            spots_used += sp
            net_used += nt
        else:
            sp = round_to_even(total_spots - spots_used)
            nt = int(total_net - net_used)
        result.append((max(0, sp), max(0, nt)))
    return result


def build_wave_rows(combos, wave_start, wave_end, wave_spots_list, wave_net_list, pricing_db, sec_factors, store_counts_num, regions_order=None):
    """
    為單一波段產出 CUE 表所需的 rows（與 calculator 產出結構一致）。
    combos: list of dict {"media": str, "region": str, "seconds": int}
    wave_spots_list: list of int，與 combos 對應，該波每組合的檔次
    wave_net_list: list of int，與 combos 對應，該波每組合的實收(未稅)
    regions_order: 六區順序，用於全省時展開為六列；若為 None 則使用預設六區。
    """
    if regions_order is None:
        regions_order = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]
    rows = []
    days_count = (wave_end - wave_start).days + 1
    if days_count <= 0:
        return rows
    for idx, combo in enumerate(combos):
        media = combo["media"]
        region = combo["region"]
        seconds = combo["seconds"]
        spots = wave_spots_list[idx] if idx < len(wave_spots_list) else 0
        net_val = wave_net_list[idx] if idx < len(wave_net_list) else 0
        if spots <= 0:
            continue
        spots = round_to_even(spots)
        factor = get_sec_factor(media, seconds, sec_factors)
        if media in ["全家廣播", "新鮮視"] and media in pricing_db:
            db = pricing_db[media]
            std_spots_ref = db["Std_Spots"]
            if region == "全省":
                # 仿照一般 CUE：全省時展開為六區各一列
                calc_std = std_spots_ref
                sch = calculate_schedule(spots, days_count)
                daypart = db.get("Day_Part", "")
                for r in regions_order:
                    if r not in db:
                        continue
                    list_price_region = db[r][0]
                    unit_rate_display = (list_price_region / calc_std) * factor
                    total_rate_display = int(unit_rate_display * spots)
                    program_num = store_counts_num.get(f"新鮮視_{r}" if media == "新鮮視" else r, 0)
                    rows.append({
                        "media": media,
                        "region": r,
                        "program_num": program_num,
                        "daypart": daypart,
                        "seconds": seconds,
                        "spots": spots,
                        "schedule": sch,
                        "rate_display": total_rate_display,
                        "pkg_display": total_rate_display,
                        "is_pkg_member": True,
                        "nat_pkg_display": net_val,
                    })
                continue
            if region not in db:
                continue
            list_price = db[region][0]
            net_price = db[region][1]
            # 依 Pricing 表各區 Std_Spots（新鮮視不再 *2）
            calc_std = (db.get("_Region_Std_Spots") or {}).get(region) or std_spots_ref
            unit_list = (list_price / calc_std) * factor
            unit_net = (net_price / calc_std) * factor
            rate_display = int(unit_list * spots)
            pkg_display = net_val if net_val > 0 else int(unit_net * spots)
            sch = calculate_schedule(spots, days_count)
            program_num = store_counts_num.get(f"新鮮視_{region}" if media == "新鮮視" else region, 0)
            daypart = db.get("Day_Part", "")
            rows.append({
                "media": media,
                "region": region,
                "program_num": program_num,
                "daypart": daypart,
                "seconds": seconds,
                "spots": spots,
                "schedule": sch,
                "rate_display": rate_display,
                "pkg_display": pkg_display,
                "is_pkg_member": False,
                "nat_pkg_display": 0,
            })
        elif media == "家樂福" and media in pricing_db:
            db = pricing_db["家樂福"]
            if "量販_全省" not in db:
                continue
            base_std = db["量販_全省"]["Std_Spots"]
            base_list = db["量販_全省"]["List"]
            base_net = db["量販_全省"]["Net"]
            unit_list = (base_list / base_std) * factor
            unit_net = (base_net / base_std) * factor
            rate_display = int(unit_list * spots)
            pkg_display = net_val if net_val > 0 else int(unit_net * spots)
            sch = calculate_schedule(spots, days_count)
            rows.append({
                "media": media,
                "region": "全省量販",
                "program_num": store_counts_num.get("家樂福_量販", 0),
                "daypart": db["量販_全省"]["Day_Part"],
                "seconds": seconds,
                "spots": spots,
                "schedule": sch,
                "rate_display": rate_display,
                "pkg_display": pkg_display,
                "is_pkg_member": False,
            })
    return rows
