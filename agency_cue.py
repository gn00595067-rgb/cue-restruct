# -*- coding: utf-8 -*-
"""
代理商 CUE 計算模組 (Agency Cue Calculator)

負責三家配合代理商（2008傳媒 / 佳聖 / 凱絡）專用 CUE 表的純計算：
實作價（賣給客戶的 Net）、牌價（表上定價 List）、每日檔次分配、
凌晨補償四方案、專案回饋加贈、費用區結構。

本模組**不 import streamlit**，可獨立單元測試。
所有檔次數字取整一律採四捨五入（round half up）。

平台只有兩個、地區只有全台：
- 全家企頻：全家店內廣播，主時段 07:00-23:00
- 萬家福：原家樂福，萬家福=量販(09-23)、樂家康=超市(00-24)

業務規則來源與驗證數字見開發規格 §2、§4。
"""

from decimal import Decimal, ROUND_HALF_UP
from datetime import date, timedelta
import json
import math
import re

import config

# =============================================================================
# 平台與凌晨補償方案常數
# =============================================================================
PLATFORM_FAMILY = "全家企頻"
PLATFORM_WJF = "萬家福"

# 凌晨補償方案（製表時四選一）
COMP_MOVE50 = "原50%檔次併入主時段"   # 預設，2008 現行做法
COMP_PLAN1 = "方案一：15%媒體價值換全家檔次"
COMP_PLAN2 = "方案二：20%媒體價值換萬家福檔次"
COMP_NONE = "無補償"

# 顯示順序：原50%放最下面（仍為預設，UI 以 index 指向它）
COMP_OPTIONS = [COMP_PLAN1, COMP_PLAN2, COMP_NONE, COMP_MOVE50]

# 2008 萬家福專屬：固定 10% 回饋檔直接折進每日排檔（不另立回饋列），實收不變。
# 其他兩家（佳聖 / 凱絡）照原規則（回饋另立列）。
WJF_2008_REBATE_PCT = 10.0

# 自帶專案回饋（凌晨時數轉換）：主時段以外的時數，每小時折 30 檔（30秒基準），
# 再依實際秒數換算（檔 = 時數 × 30 × 30 / 秒）。三家共用同一算法。
# 例：主時段 07:00-23:00（16時）→ off = 8 時；15秒 = 8×30×2 = 480 檔。
AUTO_REBATE_SPOTS_PER_HOUR_30S = 30
AUTO_REBATE_DEFAULT_OFFHOURS = 8

# row.kind
KIND_MAIN = "main"
KIND_COMP = "comp"
KIND_SUPER = "super"
KIND_REBATE = "rebate"
KIND_SUPER_REBATE = "super_rebate"

NET_REBATE = "專案回饋"
NET_ON_MAG = "計價於量販"
CARAT_REBATE_LABEL = "聲活回饋"   # 凱絡回饋列在表上顯示的名稱（內部仍記 NET_REBATE）

# 列標籤（媒體 / 地區 / 時段）依代理商查表，取自各家原始範例檔
AGENCY_LABELS = {
    "2008傳媒": {
        "family": ("全省【全家便利商店】\n(企業頻道節目播放)", "全省", "07:00-23:00"),
        "mag":    ("全省【萬家福】\n連鎖量販店", "全省", "0900-2300"),
        "super":  ("全省【樂家康】\n連鎖超市店", "全省", "0000-2400"),
    },
    "佳聖": {
        "family": ("全家便利商店店鋪廣播", "全台", "07-23"),
        "mag":    ("萬家福", "量販", "09-23"),
        "super":  ("樂家康", "超市", "00-24"),
    },
    "凱絡": {
        "family": ("全台 全家便利商店\n(企業頻道節目播放)", "全省\n全家便利商店", "0700-2300"),
        "mag":    ("全台【萬家福/樂家康】\n連鎖量販店＋超市\n(企業頻道節目播放)", "萬家福(量販)", "0900-2300 "),
        "super":  ("", "樂家康(超市)", "0000-2400 "),
    },
}


def _labels(agency, key):
    """取得（媒體, 地區, 時段）列標籤；未知代理商退回 2008 式。"""
    return AGENCY_LABELS.get(agency, AGENCY_LABELS["2008傳媒"])[key]


def family_offhours(daypart):
    """
    從主時段字串解析「非主時段時數」= 24 − 主時段時數。
    支援三家格式：'07:00-23:00'、'07-23'、'0700-2300'（可含換行/空白）。
    解析失敗退回預設 8（大多數案例為 07-23 → 8 時）。
    """
    try:
        parts = str(daypart).split("-")
        if len(parts) >= 2:
            def _hh(s):
                d = re.sub(r"\D", "", s)
                return int(d[:2]) if len(d) >= 2 else int(d)
            span = _hh(parts[1]) - _hh(parts[0])
            if 0 < span < 24:
                return 24 - span
    except (ValueError, IndexError):
        pass
    return AUTO_REBATE_DEFAULT_OFFHOURS


def auto_rebate_spots(daypart, sec):
    """自帶專案回饋檔次 = round(非主時段時數 × 30 × 30 / 秒)。"""
    off = family_offhours(daypart)
    return rhu(off * AUTO_REBATE_SPOTS_PER_HOUR_30S * 30 / sec)


# =============================================================================
# 基礎數值工具
# =============================================================================
def rhu(x):
    """四捨五入取整 (round half up)。避免 Python 內建 round 的 banker's rounding。"""
    return int(Decimal(str(x)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def round_up_to_even(x):
    """四捨五入後若為奇數則進位到偶數。"""
    r = rhu(x)
    return r if r % 2 == 0 else r + 1


def minus_business_days(d, n):
    """回傳 d 之前第 n 個工作天（跳過週六、週日）。用於素材提供時間預設。"""
    cur = d
    cnt = 0
    while cnt < n:
        cur -= timedelta(days=1)
        if cur.weekday() < 5:  # 0=Mon ... 4=Fri
            cnt += 1
    return cur


# =============================================================================
# 實作價（賣給客戶的 Net）—— 三家共用，秒數採線性比例
# =============================================================================
def family_unit_net(sec):
    """全家企頻 單檔實作價：250000/480/30 × 秒數（30秒=520.83, 15秒=260.42, 10秒=173.61）。"""
    base_net, base_std, base_sec = config.AGENCY_FAMILY_BASE
    return base_net / base_std / base_sec * sec


def wjf_unit_net(sec):
    """萬家福量販 單檔實作價：250000/420/20 × 秒數（20秒=595.24, 15秒=446.43）。"""
    base_net, base_std, base_sec = config.AGENCY_WJF_BASE
    return base_net / base_std / base_sec * sec


def super_spots_from_mag(mag_spots):
    """超市(樂家康)檔次 = 量販檔次 × 720/420（計價於量販，不另收費）。"""
    a, b = config.AGENCY_SUPER_RATIO
    return rhu(mag_spots * a / b)


# =============================================================================
# 牌價（表上定價 List，每檔單價）
# =============================================================================
def get_agency_list_price(agency, platform, sec, agency_pricing=None):
    """
    查每檔牌價。查價順序：Google Sheet(agency_pricing) → config.AGENCY_LIST_PRICE_FALLBACK。
    該秒數若無資料，取最接近秒數依線性比例換算。
    agency_pricing: {(agency, platform): {sec: per_spot}}，或 None。
    """
    table = None
    if agency_pricing:
        table = agency_pricing.get((agency, platform))
    if not table:
        table = config.AGENCY_LIST_PRICE_FALLBACK.get((agency, platform))
    if not table:
        return 0
    if sec in table:
        return int(table[sec])
    # 找最接近秒數，依線性比例換算
    nearest = min(table.keys(), key=lambda s: abs(s - sec))
    return rhu(table[nearest] / nearest * sec)


# =============================================================================
# 每日檔次分配（各家習慣，由範例反推）
# =============================================================================
def dist_plain(n, d, remainder_at="front"):
    """
    平均分配，餘數放最前(front)或最後(end)。可含奇數。
    佳聖 全部用 front；2008 萬家福用 end。
    """
    if d <= 0:
        return []
    base, rem = divmod(n, d)
    result = [base] * d
    if remainder_at == "front":
        for i in range(rem):
            result[i] += 1
    else:  # end
        for i in range(rem):
            result[d - 1 - i] += 1
    return result


def dist_even_end(n, d):
    """
    2008 全家：每日偶數；base = 偶數化(向下) 的 N//d，餘數以 +2 一組放最後幾天。
    驗證：960/28 → 34×24 + 36×4（36 在最後）。
    """
    if d <= 0:
        return []
    base = n // d
    if base % 2:
        base -= 1
    result = [base] * d
    rem = n - base * d
    i = d - 1
    while rem >= 2 and i >= 0:
        result[i] += 2
        rem -= 2
        i -= 1
    if rem == 1:  # 罕見：主檔次為奇數，餘 1 補在最後一天
        result[-1] += 1
    return result


def dist_carat(n, d):
    """
    凱絡 量販/全家：前 d-2 天 = ceil(N/d) 進位到偶數，餘數平分到最後兩天。
    驗證：560/16 → 36×14 + 28 + 28。
    """
    if d <= 0:
        return []
    if d <= 2:
        return dist_plain(n, d, "front")
    per_day = round_up_to_even(math.ceil(n / d))
    head = d - 2
    used = per_day * head
    remaining = n - used
    if remaining < 0:
        # per_day 進位過頭，退回平均分配
        return dist_plain(n, d, "front")
    a = remaining // 2
    b = remaining - a
    return [per_day] * head + [a, b]


def dist_carat_super(total_super, mag_schedule):
    """
    凱絡 超市：前 d-2 天 = 當日量販 × 720/420 四捨五入到偶數，餘數平分到最後兩天。
    驗證：960/16(量販36) → 62×14 + 46 + 46。
    """
    d = len(mag_schedule)
    if d <= 0:
        return []
    if d <= 2:
        return dist_plain(total_super, d, "front")
    a_ratio, b_ratio = config.AGENCY_SUPER_RATIO
    head = d - 2
    result = [round_up_to_even(mag_schedule[i] * a_ratio / b_ratio) for i in range(head)]
    used = sum(result)
    remaining = total_super - used
    if remaining < 0:
        return dist_plain(total_super, d, "front")
    a = remaining // 2
    b = remaining - a
    return result + [a, b]


# =============================================================================
# 模型建構
# =============================================================================
def _new_row(kind, media_label, region_label, daypart, seconds, spots, schedule,
             list_per=0, list_total=0, market_per=0, uni_per=0, uni_total=0,
             net_display=0, material=""):
    return {
        "kind": kind,
        "media_label": media_label,
        "region_label": region_label,
        "daypart": daypart,
        "seconds": seconds,
        "spots": spots,
        "schedule": schedule,          # list（逐日排檔）或 None（合併顯示總數）
        "list_per": list_per,          # 每檔牌價
        "list_total": list_total,      # 定價合計 = 牌價 × 檔次
        "market_per": market_per,      # 凱絡 市場價/檔
        "uni_per": uni_per,            # 凱絡 統一價/檔
        "uni_total": uni_total,        # 凱絡 總價 = 統一價 × 檔次
        "net_display": net_display,    # 實收數字 或 "專案回饋"/"計價於量販"
        "material": material,          # 素材提供時間（顯示字串）
    }


def _build_family_sheet(agency, sec, budget, days, comp_mode, rebate_pct,
                        spots_override, material_due, ac_pct, agency_pricing, logs,
                        auto_rebate=True):
    """建立全家企頻 sheet。auto_rebate：自帶專案回饋（凌晨時數轉換），預設開。"""
    unit_net = family_unit_net(sec)
    main_spots = spots_override if spots_override else rhu(budget / unit_net)
    logs.append(f"【全家企頻】單檔實作價({sec}秒)=250000/480/30×{sec}={unit_net:.2f}")
    logs.append(f"【全家企頻】主時段檔次=round({budget}/{unit_net:.2f})={main_spots}")

    # 凌晨補償
    comp_spots = 0
    if comp_mode == COMP_MOVE50:
        comp_spots = rhu(main_spots * 0.5)
        logs.append(f"【全家企頻】凌晨補償(原50%)=round({main_spots}×0.5)={comp_spots}")
    elif comp_mode == COMP_PLAN1:
        comp_spots = rhu(0.15 * budget / unit_net)
        logs.append(f"【全家企頻】凌晨補償(方案一15%)=round(0.15×{budget}/{unit_net:.2f})={comp_spots}")
    # COMP_PLAN2 的補償走萬家福表（見 build_agency_model），此處全家不加補償列
    # COMP_NONE：無

    list_per = get_agency_list_price(agency, PLATFORM_FAMILY, sec, agency_pricing)
    fam_label, fam_region, fam_daypart = _labels(agency, "family")
    rows = []

    if agency == "2008傳媒":
        # 主列逐日排檔；補償列合併顯示總數（schedule=None）
        main_sch = dist_even_end(main_spots, days)
        net_value = rhu(unit_net * main_spots)  # 「合計」欄（G）= 實作價值
        rows.append(_new_row(
            KIND_MAIN, fam_label, fam_region, fam_daypart,
            sec, main_spots, main_sch,
            list_per=list_per, list_total=list_per * (main_spots + comp_spots),
            net_display=net_value, material=material_due,
        ))
        if comp_spots > 0:
            rows.append(_new_row(
                KIND_COMP, "", "", "", sec, comp_spots, None,
                list_per=list_per, list_total=0, net_display=0,
            ))
        budget_net = net_value
    else:
        # 佳聖 / 凱絡：補償直接併入主列（單列 = 主+補償），逐日排檔
        total_main = main_spots + comp_spots
        if agency == "凱絡":
            main_sch = dist_carat(total_main, days)
        else:  # 佳聖
            main_sch = dist_plain(total_main, days, "front")
        market_per = rhu(list_per * config.CARAT_MARKET_RATIO)
        uni_per = rhu(list_per * config.CARAT_UNI_RATIO)
        rows.append(_new_row(
            KIND_MAIN, fam_label, fam_region, fam_daypart, sec, total_main, main_sch,
            list_per=list_per, list_total=list_per * total_main,
            market_per=market_per, uni_per=uni_per, uni_total=uni_per * total_main,
            net_display=budget, material=material_due,
        ))
        budget_net = budget

    # 回饋列建構（2008 合併顯示、佳聖／凱絡逐日鋪滿），供自帶回饋與手動%回饋共用
    def _append_rebate(reb):
        if reb <= 0:
            return
        if agency == "2008傳媒":
            rows.append(_new_row(
                KIND_REBATE, "", "", "", sec, reb, None,
                list_per=list_per, list_total=list_per * reb, net_display=NET_REBATE,
            ))
        else:
            # 凱絡／佳聖 回饋皆逐日平均鋪滿（對齊範本，凱絡不再合併顯示）
            reb_sch = dist_plain(reb, days, "front")
            market_per = rhu(list_per * config.CARAT_MARKET_RATIO)
            uni_per = rhu(list_per * config.CARAT_UNI_RATIO)
            rows.append(_new_row(
                KIND_REBATE, "", "", "", sec, reb, reb_sch,
                list_per=list_per, list_total=list_per * reb,
                market_per=market_per, uni_per=uni_per, uni_total=uni_per * reb,
                net_display=NET_REBATE,
            ))

    # 自帶專案回饋（凌晨時數轉換）：非主時段時數 × 30檔 × 30/秒。三家共用。
    if auto_rebate:
        off = family_offhours(fam_daypart)
        auto_reb = auto_rebate_spots(fam_daypart, sec)
        logs.append(f"【全家企頻】自帶專案回饋(時數轉換)=(24-主時段)×30×30/{sec}"
                    f"={off}×30×{30/sec:g}={auto_reb}")
        _append_rebate(auto_reb)

    # 專案回饋（手動加贈 %）
    if rebate_pct and rebate_pct > 0:
        reb = rhu(rebate_pct / 100.0 * (main_spots + comp_spots))
        logs.append(f"【全家企頻】專案回饋={rebate_pct}%×({main_spots}+{comp_spots})={reb}")
        _append_rebate(reb)

    fees = _build_fees(agency, budget_net, ac_pct, is_rebate_wave=False)
    return {
        "platform": PLATFORM_FAMILY,
        "seconds": sec,
        "budget": budget,
        "is_rebate_wave": False,
        "rows": rows,
        "fees": fees,
    }


def _build_wjf_sheet(agency, sec, budget, days, rebate_pct, mag_override,
                     is_rebate_wave, material_due, ac_pct, agency_pricing, logs,
                     comp_mag=0, comp_super=0):
    """
    建立萬家福 sheet（量販 + 超市）。
    comp_mag / comp_super：方案二凌晨補償轉入的量販/超市檔次（另立補償列）。
    """
    unit_net = wjf_unit_net(sec)
    if is_rebate_wave:
        mag_spots = mag_override if mag_override else 0
        logs.append(f"【萬家福】整波專案回饋：量販手動={mag_spots}")
    else:
        mag_base = mag_override if mag_override else rhu(budget / unit_net)
        logs.append(f"【萬家福】量販單檔實作價({sec}秒)=250000/420/20×{sec}={unit_net:.2f}")
        logs.append(f"【萬家福】量販基礎檔次=round({budget}/{unit_net:.2f})={mag_base}")
        # 2008 專屬：固定 +10% 回饋檔直接折進每日排檔（不另立回饋列），實收不變
        if agency == "2008傳媒":
            bonus = rhu(WJF_2008_REBATE_PCT / 100.0 * mag_base)
            mag_spots = mag_base + bonus
            logs.append(f"【萬家福】2008專屬：+{int(WJF_2008_REBATE_PCT)}%回饋檔折進每日 "
                        f"= {mag_base}+round({WJF_2008_REBATE_PCT/100:.2f}×{mag_base})={mag_spots}（實收不變）")
        else:
            mag_spots = mag_base
    super_spots = super_spots_from_mag(mag_spots)
    logs.append(f"【萬家福】超市檔次={mag_spots}×720/420={super_spots}")

    list_per = get_agency_list_price(agency, PLATFORM_WJF, sec, agency_pricing)
    market_per = rhu(list_per * config.CARAT_MARKET_RATIO)
    uni_per = rhu(list_per * config.CARAT_UNI_RATIO)

    mag_label, mag_region, mag_daypart = _labels(agency, "mag")
    super_label, super_region, super_daypart = _labels(agency, "super")

    rows = []
    mag_net = NET_REBATE if is_rebate_wave else budget
    super_net = NET_REBATE if is_rebate_wave else NET_ON_MAG

    # 量販主列
    if agency == "凱絡":
        mag_sch = dist_carat(mag_spots, days)
    else:
        mag_sch = dist_plain(mag_spots, days, "front")
    rows.append(_new_row(
        KIND_MAIN, mag_label, mag_region, mag_daypart, sec, mag_spots, mag_sch,
        list_per=list_per, list_total=list_per * mag_spots,
        market_per=market_per, uni_per=uni_per, uni_total=uni_per * mag_spots,
        net_display=mag_net, material=material_due,
    ))
    # 超市列（計價於量販）
    if agency == "凱絡":
        super_sch = dist_carat_super(super_spots, mag_sch)
    else:
        super_sch = dist_plain(super_spots, days, "front")
    rows.append(_new_row(
        KIND_SUPER, super_label, super_region, super_daypart, sec, super_spots, super_sch,
        list_per=list_per, list_total=0,
        market_per=market_per, uni_per=uni_per, uni_total=uni_per * super_spots,
        net_display=super_net,
    ))

    # 方案二凌晨補償列（若有）
    if comp_mag > 0:
        rows.append(_new_row(
            KIND_MAIN, mag_label, mag_region + "(凌晨轉換)", mag_daypart, sec,
            comp_mag, None, list_per=list_per, list_total=list_per * comp_mag,
            market_per=market_per, uni_per=uni_per, uni_total=uni_per * comp_mag,
            net_display=NET_REBATE,
        ))
        rows.append(_new_row(
            KIND_SUPER, super_label, super_region + "(凌晨轉換)", super_daypart, sec,
            comp_super, None, list_per=list_per, list_total=0,
            market_per=market_per, uni_per=uni_per, uni_total=uni_per * comp_super,
            net_display=NET_ON_MAG,
        ))

    # 專案回饋（加贈）。2008 萬家福已把 10% 折進主檔每日排檔，不再另立回饋列。
    if rebate_pct and rebate_pct > 0 and not is_rebate_wave and agency != "2008傳媒":
        reb_mag = rhu(rebate_pct / 100.0 * mag_spots)
        reb_super = super_spots_from_mag(reb_mag)
        logs.append(f"【萬家福】專案回饋：量販={rebate_pct}%×{mag_spots}={reb_mag}、超市={reb_super}")
        # 凱絡／佳聖 回饋皆逐日平均鋪滿（對齊範本，凱絡不再合併顯示）
        reb_sch = dist_plain(reb_mag, days, "front")
        reb_super_sch = dist_plain(reb_super, days, "front")
        rows.append(_new_row(
            KIND_REBATE, mag_label, mag_region, mag_daypart, sec, reb_mag, reb_sch,
            list_per=list_per, list_total=list_per * reb_mag,
            market_per=market_per, uni_per=uni_per, uni_total=uni_per * reb_mag,
            net_display=NET_REBATE,
        ))
        rows.append(_new_row(
            KIND_SUPER_REBATE, super_label, super_region, super_daypart, sec, reb_super, reb_super_sch,
            list_per=list_per, list_total=0,
            market_per=market_per, uni_per=uni_per, uni_total=uni_per * reb_super,
            net_display=NET_ON_MAG,
        ))

    budget_net = 0 if is_rebate_wave else budget
    fees = _build_fees(agency, budget_net, ac_pct, is_rebate_wave=is_rebate_wave)
    return {
        "platform": PLATFORM_WJF,
        "seconds": sec,
        "budget": budget,
        "is_rebate_wave": is_rebate_wave,
        "rows": rows,
        "fees": fees,
    }


def _build_fees(agency, budget_net, ac_pct, is_rebate_wave=False):
    """依代理商建立費用區結構。金額四捨五入取整。"""
    if is_rebate_wave:
        # 整波回饋：費用全 0，TOTAL 顯示「專案回饋」
        if agency == "2008傳媒":
            return {"type": "2008", "budget_net": 0, "ac_pct": ac_pct or 0,
                    "ac": 0, "tax": 0, "total": NET_REBATE}
        if agency == "凱絡":
            return {"type": "carat", "subtotal": 0, "ac_free": True, "ac": 0,
                    "vat": 0, "grand": NET_REBATE}
        return {"type": "ddrive", "net": 0, "vat": 0, "gross": NET_REBATE}

    if agency == "2008傳媒":
        acp = ac_pct if ac_pct is not None else config.AGENCY_AC_DEFAULT.get("2008傳媒", 3.0)
        ac = rhu(budget_net * acp / 100.0)
        tax = rhu((budget_net + ac) * 0.05)
        total = budget_net + ac + tax
        return {"type": "2008", "budget_net": budget_net, "ac_pct": acp,
                "ac": ac, "tax": tax, "total": total}
    if agency == "凱絡":
        # A.C 3% 預設免收（顯示 -）；ac_pct 若為數字則收
        ac_free = ac_pct is None
        ac = 0 if ac_free else rhu(budget_net * ac_pct / 100.0)
        vat = rhu((budget_net + ac) * 0.05)
        grand = budget_net + ac + vat
        return {"type": "carat", "subtotal": budget_net, "ac_free": ac_free,
                "ac": ac, "vat": vat, "grand": grand}
    # 佳聖：無 AC
    vat = rhu(budget_net * 0.05)
    gross = budget_net + vat
    return {"type": "ddrive", "net": budget_net, "vat": vat, "gross": gross}


def build_agency_model(agency, client_name, product_name, campaign,
                       start_date, end_date, budget_total,
                       fam_cfg, wjf_cfg, comp_mode, material_due, ac_pct,
                       agency_pricing=None, payment_note="", remarks=None):
    """
    建立代理商 CUE model。

    fam_cfg / wjf_cfg：平台設定 dict，可為 None（未啟用）。
      共通鍵：enabled(bool)、seconds(int)、share(佔比%)、rebate_pct(float)、spots_override(int, 0=自動)
      fam_cfg 另有：auto_rebate(bool, 預設True=自帶專案回饋/凌晨時數轉換)
      wjf_cfg 另有：mag_override(int)、is_rebate_wave(bool)
    comp_mode：COMP_MOVE50 / COMP_PLAN1 / COMP_PLAN2 / COMP_NONE
    回傳 model：{agency, client_name, product_name, campaign, start_date, end_date,
                budget_total, sheets:[...], logs:[...], material_due, payment_note, remarks}
    """
    days = (end_date - start_date).days + 1
    logs = []
    logs.append(f"代理商={agency}；走期 {start_date}~{end_date} 共 {days} 天；總預算(未稅)={budget_total}")

    fam_on = bool(fam_cfg and fam_cfg.get("enabled"))
    wjf_on = bool(wjf_cfg and wjf_cfg.get("enabled"))

    # 依佔比拆分預算
    fam_share = fam_cfg.get("share", 100) if fam_on else 0
    wjf_share = wjf_cfg.get("share", 100) if wjf_on else 0
    if fam_on and not wjf_on:
        fam_budget = budget_total
    elif wjf_on and not fam_on:
        wjf_budget = budget_total
    if fam_on and wjf_on:
        fam_budget = rhu(budget_total * fam_share / 100.0)
        wjf_budget = budget_total - fam_budget
    if not fam_on:
        fam_budget = 0
    if not wjf_on:
        wjf_budget = 0

    sheets = []

    # 方案二：全家凌晨 20% 媒體價值轉萬家福（即使萬家福未單獨啟用也要產生補償）
    comp_mag = comp_super = 0
    comp_wjf_sec = wjf_cfg.get("seconds", 15) if wjf_cfg else 15
    if fam_on and comp_mode == COMP_PLAN2:
        comp_mag = rhu(0.20 * fam_budget / wjf_unit_net(comp_wjf_sec))
        comp_super = super_spots_from_mag(comp_mag)
        logs.append(f"【凌晨補償方案二】20%×{fam_budget}/萬家福單檔({comp_wjf_sec}秒)="
                    f"量販{comp_mag}、超市{comp_super}")

    if fam_on:
        sheets.append(_build_family_sheet(
            agency, int(fam_cfg["seconds"]), fam_budget, days, comp_mode,
            fam_cfg.get("rebate_pct", 0), int(fam_cfg.get("spots_override", 0) or 0),
            material_due, ac_pct, agency_pricing, logs,
            auto_rebate=bool(fam_cfg.get("auto_rebate", True)),
        ))

    # 方案二補償需要一張萬家福表承載；若萬家福未啟用則單獨為補償建表
    need_wjf_for_comp = (comp_mag > 0) and not wjf_on
    if wjf_on:
        sheets.append(_build_wjf_sheet(
            agency, int(wjf_cfg["seconds"]), wjf_budget, days,
            wjf_cfg.get("rebate_pct", 0), int(wjf_cfg.get("mag_override", 0) or 0),
            bool(wjf_cfg.get("is_rebate_wave", False)),
            material_due, ac_pct, agency_pricing, logs,
            comp_mag=comp_mag, comp_super=comp_super,
        ))
    elif need_wjf_for_comp:
        # 僅承載凌晨補償的萬家福表（本身無主預算）
        sheet = _build_wjf_sheet(
            agency, comp_wjf_sec, 0, days, 0, 0, False,
            material_due, ac_pct, agency_pricing, logs,
            comp_mag=comp_mag, comp_super=comp_super,
        )
        # 移除量販/超市主列（本身無預算 spots=0），只保留補償列（spots>0）
        sheet["rows"] = [r for r in sheet["rows"] if r["spots"] > 0]
        sheets.append(sheet)

    if remarks is None:
        remarks = default_remarks(agency)
    if not payment_note:
        # 以全部 sheet 的實收合計估算請款
        net_total = 0
        for s in sheets:
            f = s["fees"]
            if f["type"] == "2008":
                net_total += f["budget_net"] if isinstance(f["budget_net"], (int, float)) else 0
            elif f["type"] == "ddrive":
                net_total += f["net"] if isinstance(f["net"], (int, float)) else 0
            else:
                net_total += f["subtotal"] if isinstance(f["subtotal"], (int, float)) else 0
        payment_note = default_payment_note(start_date, end_date, net_total)

    return {
        "agency": agency,
        "client_name": client_name,
        "product_name": product_name,
        "campaign": campaign,
        "start_date": start_date,
        "end_date": end_date,
        "budget_total": budget_total,
        "material_due": material_due,
        "payment_note": payment_note,
        "remarks": remarks,
        "sheets": sheets,
        "logs": logs,
    }


# =============================================================================
# 預設備註與請款說明
# =============================================================================
def default_remarks(agency, sign_date=None, payment_note=""):
    """各代理商預設備註模板（原始範例原文）。"""
    if agency == "2008傳媒":
        return [
            "※以上各贈檔時段若已滿檔，則須更動時段",
            "※以上金額為目前之價格，往後若有調整，則以調整後價格計算",
            "※以上報價不包含錄音製作(不得指定特殊配音員或錄音室)",
            "※廣告素材至少需要播出日之7天前給予(週六.週日除外)，若無如期給予素材檔次會如期順延",
        ]
    if agency == "凱絡":
        if sign_date:
            wd = "一二三四五六日"[sign_date.weekday()]
            sd = f"{sign_date.month}/{sign_date.day}(週{wd})"
        else:
            sd = "__/__(週_)"
        pay = payment_note or "依走期各月天數比例拆分，詳如報價"
        return [
            "1.更改媒體排期表請於播出前十個工作天通知。",
            "2.以上報價僅適用於本次的專案優惠價，而非代表媒體單次購買價格；若中途停Cue，則已執行之廣告檔次，以單檔單次計費。",
            "3.客戶選擇凱絡公司提供的媒體服務，客戶並同意接受凱絡公司的條款（CARAT TERMS OF BUSINESS）的約束。 "
            "CARAT TERMS OF BUSINESS網址：http://file.carat.com.tw/legal/Carat_terms_of_business.pdf。",
            f"4.請於{sd}中午前簽回Cue表確認，並傳真至(02)2717-3011，謝謝",
            f"5.請款說明：{pay}",
            "6.■以下請PM填寫：",
            "★廣告稿件已送品牌創意檢視表電子化系統核准製作?",
            "□是，文件編號__________________費用分攤明細__________________。",
            "□否，請於上CUE前完成電子流程送簽。",
        ]
    # 佳聖：備註只有請款金額一行，由渲染器自組
    return []


def default_payment_note(start_dt, end_dt, net_total):
    """
    依走期各月天數比例拆分請款金額（千元四捨五入、尾月吃餘數）。
    回傳如：「* 請款金額：8月份Net$125,000元(未稅)。」
    """
    if not net_total or net_total <= 0:
        return "* 請款金額：Net$0元(未稅)。"
    start_d = start_dt if isinstance(start_dt, date) else start_dt.date()
    end_d = end_dt if isinstance(end_dt, date) else end_dt.date()
    total_days = (end_d - start_d).days + 1

    # 依月統計天數
    months = []  # [(month, days)]
    cur = start_d
    while cur <= end_d:
        if cur.month == 12:
            month_last = date(cur.year, 12, 31)
        else:
            month_last = date(cur.year, cur.month + 1, 1) - timedelta(days=1)
        seg_end = min(month_last, end_d)
        months.append((cur.month, (seg_end - cur).days + 1))
        cur = seg_end + timedelta(days=1)

    parts = []
    allocated = 0
    for idx, (m, dcnt) in enumerate(months):
        if idx == len(months) - 1:
            amt = net_total - allocated  # 尾月吃餘數
        else:
            amt = rhu(net_total * dcnt / total_days / 1000) * 1000  # 千元四捨五入
            allocated += amt
        parts.append(f"{m}月份Net${amt:,}元(未稅)")
    return "* 請款金額：" + "、".join(parts) + "。"


# =============================================================================
# Ragic 上傳輔助（代理商 CUE 與一般 CUE 共用同一張 Ragic 表單）
# =============================================================================
# 上傳標記：details 欄位以此字串開頭，搜尋時可據以篩出「代理商 CUE」案子。
AGENCY_RAGIC_TAG = "【代理商CUE】"
# details 內嵌完整表單輸入的 JSON 段落標記，供「載入舊案」100% 還原表單。
AGENCY_EXT_MARKER = "[AGENCY_EXT]"


def agency_grand_total(model):
    """全部 sheet 的「含稅合計」總和（給 Ragic budget_fin 用）。"""
    total = 0
    for s in model.get("sheets", []):
        f = s.get("fees", {})
        for k in ("total", "gross", "grand"):  # 2008 / 佳聖 / 凱絡
            v = f.get(k)
            if isinstance(v, (int, float)):
                total += v
                break
    return int(total)


def agency_net_total(model):
    """全部 sheet 的「未稅實收 Net」總和。"""
    total = 0
    for s in model.get("sheets", []):
        f = s.get("fees", {})
        for k in ("budget_net", "net", "subtotal"):  # 2008 / 佳聖 / 凱絡
            v = f.get(k)
            if isinstance(v, (int, float)):
                total += v
                break
    return int(total)


def agency_total_spots(model):
    """所有列的檔次加總（含補償/回饋/超市）。"""
    total = 0
    for s in model.get("sheets", []):
        for r in s.get("rows", []):
            sp = r.get("spots")
            if isinstance(sp, (int, float)):
                total += int(sp)
    return total


def agency_platform_text(model):
    """平台清單文字，如「全家企頻、萬家福」。"""
    seen = []
    for s in model.get("sheets", []):
        p = s.get("platform")
        if p and p not in seen:
            seen.append(p)
    return "、".join(seen)


def agency_seconds_union(model):
    """使用到的秒數聯集文字，如「15、20」。"""
    secs = sorted({int(s["seconds"]) for s in model.get("sheets", []) if s.get("seconds")})
    return "、".join(str(x) for x in secs)


def build_agency_ragic_details(model, form_inputs):
    """
    組出寫入 Ragic「備註(details)」欄位的字串：
      1. 人可讀摘要（代理商 / 平台 / 檔次 / 費用）
      2. [AGENCY_EXT] JSON：完整表單輸入，供「載入舊案」還原。

    form_inputs 由 UI 蒐集，須為 JSON 可序列化（日期用 isoformat 字串）。
    """
    lines = [f"{AGENCY_RAGIC_TAG}代理商：{model['agency']}"]
    if model.get("campaign"):
        lines.append(f"Campaign：{model['campaign']}")
    lines.append(f"平台：{agency_platform_text(model)}（秒數 {agency_seconds_union(model)}）")
    lines.append(f"總檔數：{agency_total_spots(model)}　實收(未稅)：${agency_net_total(model):,}"
                 f"　含稅合計：${agency_grand_total(model):,}")
    for s in model.get("sheets", []):
        f = s.get("fees", {})
        gt = next((f[k] for k in ("total", "gross", "grand") if isinstance(f.get(k), (int, float))), 0)
        lines.append(f"・{s.get('platform')}（{s.get('seconds')}秒）含稅 ${int(gt):,}")
    if model.get("payment_note"):
        lines.append(model["payment_note"])
    summary = "\n".join(lines)
    ext = json.dumps(form_inputs, ensure_ascii=False)
    return f"{summary}\n{AGENCY_EXT_MARKER}{ext}"


def parse_agency_ext(details):
    """從 details 字串取出 [AGENCY_EXT] JSON；非代理商案或損毀時回傳 None。"""
    if not details or AGENCY_EXT_MARKER not in details:
        return None
    try:
        ext_part = details.split(AGENCY_EXT_MARKER, 1)[1].strip()
        if not ext_part:
            return None
        return json.loads(ext_part)
    except (json.JSONDecodeError, ValueError, TypeError):
        return None


def is_agency_record_details(details):
    """該筆 Ragic 資料是否為代理商 CUE（供搜尋清單篩選）。"""
    if not details:
        return False
    return AGENCY_RAGIC_TAG in details or AGENCY_EXT_MARKER in details
