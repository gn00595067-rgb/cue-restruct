# -*- coding: utf-8 -*-
"""
代理商 CUE 計算模組 (Agency Cue Calculator)

負責三家配合代理商（2008傳媒 / D drive / 凱絡）專用 CUE 表的純計算：
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
import math

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

COMP_OPTIONS = [COMP_MOVE50, COMP_PLAN1, COMP_PLAN2, COMP_NONE]

# row.kind
KIND_MAIN = "main"
KIND_COMP = "comp"
KIND_SUPER = "super"
KIND_REBATE = "rebate"
KIND_SUPER_REBATE = "super_rebate"

NET_REBATE = "專案回饋"
NET_ON_MAG = "計價於量販"


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
    D drive 全部用 front；2008 萬家福用 end。
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


def _material_str_2008(material_due):
    """2008 素材提供時間顯示為 M/D 文字。"""
    if not material_due:
        return ""
    return f"{material_due.month}/{material_due.day}"


def _build_family_sheet(agency, sec, budget, days, comp_mode, rebate_pct,
                        spots_override, material_due, ac_pct, agency_pricing, logs):
    """建立全家企頻 sheet。"""
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
    rows = []

    if agency == "2008傳媒":
        # 主列逐日排檔；補償列合併顯示總數（schedule=None）
        main_sch = dist_even_end(main_spots, days)
        net_value = rhu(unit_net * main_spots)  # 「合計」欄（G）= 實作價值
        rows.append(_new_row(
            KIND_MAIN, "全省【全家便利商店】\n(企業頻道節目播放)", "全省", "07:00-23:00",
            sec, main_spots, main_sch,
            list_per=list_per, list_total=list_per * (main_spots + comp_spots),
            net_display=net_value, material=_material_str_2008(material_due),
        ))
        if comp_spots > 0:
            rows.append(_new_row(
                KIND_COMP, "", "", "", sec, comp_spots, None,
                list_per=list_per, list_total=0, net_display=0,
            ))
        budget_net = net_value
    else:
        # D drive / 凱絡：補償直接併入主列（單列 = 主+補償），逐日排檔
        total_main = main_spots + comp_spots
        if agency == "凱絡":
            main_sch = dist_carat(total_main, days)
            media_label = "全台 全家便利商店\n(企業頻道節目播放)"
            daypart = "0700-2300"
        else:  # D drive
            main_sch = dist_plain(total_main, days, "front")
            media_label = "全家便利商店店鋪廣播"
            daypart = "07:00-23:00"
        market_per = rhu(list_per * config.CARAT_MARKET_RATIO)
        uni_per = rhu(list_per * config.CARAT_UNI_RATIO)
        rows.append(_new_row(
            KIND_MAIN, media_label, "全台", daypart, sec, total_main, main_sch,
            list_per=list_per, list_total=list_per * total_main,
            market_per=market_per, uni_per=uni_per, uni_total=uni_per * total_main,
            net_display=budget, material=_material_str_2008(material_due),
        ))
        budget_net = budget

    # 專案回饋（加贈）
    if rebate_pct and rebate_pct > 0:
        reb = rhu(rebate_pct / 100.0 * (main_spots + comp_spots))
        logs.append(f"【全家企頻】專案回饋={rebate_pct}%×({main_spots}+{comp_spots})={reb}")
        if agency == "2008傳媒":
            rows.append(_new_row(
                KIND_REBATE, "", "", "", sec, reb, None,
                list_per=list_per, list_total=list_per * reb, net_display=NET_REBATE,
            ))
        else:
            if agency == "凱絡":
                reb_sch = None  # 凱絡回饋列合併顯示
            else:
                reb_sch = dist_plain(reb, days, "front")  # D drive 逐日排檔
            market_per = rhu(list_per * config.CARAT_MARKET_RATIO)
            uni_per = rhu(list_per * config.CARAT_UNI_RATIO)
            rows.append(_new_row(
                KIND_REBATE, "", "", "", sec, reb, reb_sch,
                list_per=list_per, list_total=list_per * reb,
                market_per=market_per, uni_per=uni_per, uni_total=uni_per * reb,
                net_display=NET_REBATE,
            ))

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
        mag_spots = mag_override if mag_override else rhu(budget / unit_net)
        logs.append(f"【萬家福】量販單檔實作價({sec}秒)=250000/420/20×{sec}={unit_net:.2f}")
        logs.append(f"【萬家福】量販檔次=round({budget}/{unit_net:.2f})={mag_spots}")
    super_spots = super_spots_from_mag(mag_spots)
    logs.append(f"【萬家福】超市檔次={mag_spots}×720/420={super_spots}")

    list_per = get_agency_list_price(agency, PLATFORM_WJF, sec, agency_pricing)
    market_per = rhu(list_per * config.CARAT_MARKET_RATIO)
    uni_per = rhu(list_per * config.CARAT_UNI_RATIO)

    rows = []
    mag_net = NET_REBATE if is_rebate_wave else budget
    super_net = NET_REBATE if is_rebate_wave else NET_ON_MAG

    # 量販主列
    if agency == "凱絡":
        mag_sch = dist_carat(mag_spots, days)
    else:
        mag_sch = dist_plain(mag_spots, days, "front")
    rows.append(_new_row(
        KIND_MAIN, "全省【萬家福】\n連鎖量販店", "全省", "09:00-23:00", sec, mag_spots, mag_sch,
        list_per=list_per, list_total=list_per * mag_spots,
        market_per=market_per, uni_per=uni_per, uni_total=uni_per * mag_spots,
        net_display=mag_net, material=_material_str_2008(material_due),
    ))
    # 超市列（計價於量販）
    if agency == "凱絡":
        super_sch = dist_carat_super(super_spots, mag_sch)
    else:
        super_sch = dist_plain(super_spots, days, "front")
    rows.append(_new_row(
        KIND_SUPER, "全省【樂家康】\n連鎖超市店", "全省", "00:00-24:00", sec, super_spots, super_sch,
        list_per=list_per, list_total=0,
        market_per=market_per, uni_per=uni_per, uni_total=uni_per * super_spots,
        net_display=super_net,
    ))

    # 方案二凌晨補償列（若有）
    if comp_mag > 0:
        rows.append(_new_row(
            KIND_MAIN, "全省【萬家福】\n連鎖量販店(凌晨轉換)", "全省", "09:00-23:00", sec,
            comp_mag, None, list_per=list_per, list_total=list_per * comp_mag,
            market_per=market_per, uni_per=uni_per, uni_total=uni_per * comp_mag,
            net_display=NET_REBATE,
        ))
        rows.append(_new_row(
            KIND_SUPER, "全省【樂家康】\n連鎖超市店(凌晨轉換)", "全省", "00:00-24:00", sec,
            comp_super, None, list_per=list_per, list_total=0,
            market_per=market_per, uni_per=uni_per, uni_total=uni_per * comp_super,
            net_display=NET_ON_MAG,
        ))

    # 專案回饋（加贈）
    if rebate_pct and rebate_pct > 0 and not is_rebate_wave:
        reb_mag = rhu(rebate_pct / 100.0 * mag_spots)
        reb_super = super_spots_from_mag(reb_mag)
        logs.append(f"【萬家福】專案回饋：量販={rebate_pct}%×{mag_spots}={reb_mag}、超市={reb_super}")
        reb_sch = None if agency == "凱絡" else dist_plain(reb_mag, days, "front")
        reb_super_sch = None if agency == "凱絡" else dist_plain(reb_super, days, "front")
        rows.append(_new_row(
            KIND_REBATE, "全省【萬家福】\n連鎖量販店", "全省", "09:00-23:00", sec, reb_mag, reb_sch,
            list_per=list_per, list_total=list_per * reb_mag,
            market_per=market_per, uni_per=uni_per, uni_total=uni_per * reb_mag,
            net_display=NET_REBATE,
        ))
        rows.append(_new_row(
            KIND_SUPER_REBATE, "全省【樂家康】\n連鎖超市店", "全省", "00:00-24:00", sec, reb_super, reb_super_sch,
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
    # D drive：無 AC
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
        # 移除量販/超市主列（本身無預算），只保留補償列
        sheet["rows"] = [r for r in sheet["rows"] if r["spots"] > 0 and r["kind"] not in (KIND_MAIN, KIND_SUPER)
                         or "凌晨" in r["media_label"]]
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
def default_remarks(agency, sign_date=None):
    """各代理商預設備註模板。"""
    if agency == "2008傳媒":
        return [
            "※以上各贈檔時段若已滿檔，則須更動時段",
            "※以上金額為目前之價格，往後若有調整，則以調整後價格計算",
            "※以上報價不包含錄音製作(不得指定特殊配音員或錄音室)",
            "※廣告素材至少需要播出日之7天前給予(週六.週日除外)，若無如期給予素材檔次會如期順延",
        ]
    if agency == "凱絡":
        sd = sign_date.strftime("%Y/%m/%d") if sign_date else "____/__/__"
        return [
            f"1. 請於 {sd} 前回簽及進單，方可順利上檔。",
            "2. 通路店鋪數與開機率至少七成(以上)，每日因加盟數調整、店舖改裝維護等狀況會有一定幅度增減。",
            "3. 以上節目時段如遇滿檔，上檔時間挪後或更換至同級時段。",
            "4. 廣告素材需於上檔前 5 個工作天提供。",
            "5. 請款說明：依走期各月天數比例拆分，詳如請款金額。",
            "★廣告稿件提供：",
            "□是，已提供廣告稿件，並確認以此稿件託播。",
            "□否，尚未提供廣告稿件，將於上檔前補齊。",
        ]
    # D drive
    return ["備註："]


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
