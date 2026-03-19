"""
工具函式模組 (Utility Functions)
包含字串處理、數學計算、HTML 轉義等通用功能
"""

import re
import json
from datetime import datetime, timedelta, date
from config import REGION_DISPLAY_MAP


def parse_count_to_int(x):
    """將輸入值 (可能是字串含逗號) 轉換為整數。"""
    if x is None:
        return 0
    if isinstance(x, (int, float)):
        return int(x)
    s = str(x)
    m = re.findall(r"[\d,]+", s)
    return int(m[0].replace(",", "")) if m else 0


def safe_filename(name: str) -> str:
    """移除檔名中的非法字元，避免下載時檔名錯誤。"""
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip()


def html_escape(s):
    """HTML 特殊字元轉義，防止 HTML Injection 或格式跑版。"""
    if s is None:
        return ""
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&#39;")


def region_display(region):
    """將簡寫區域轉為完整顯示名稱 (用於報表顯示)。"""
    return REGION_DISPLAY_MAP.get(region, region)


def get_sec_factor(media_type, seconds, sec_factors):
    """
    取得秒數加成係數 (Factor)。
    若查無特定秒數，嘗試以 10, 20, 15, 30 為基準進行比例換算。
    """
    factors = sec_factors.get(media_type)
    if not factors:
        if media_type == "新鮮視":
            factors = sec_factors.get("全家新鮮視")
        elif media_type == "全家廣播":
            factors = sec_factors.get("全家廣播")
    if not factors:
        return 1.0
    if seconds in factors:
        return factors[seconds]
    # 若無直接定義，嘗試推算
    for base in [10, 20, 15, 30]:
        if base in factors:
            return (seconds / base) * factors[base]
    return 1.0


def calculate_schedule(total_spots, days):
    """
    排程分配演算法：將總檔次 (total_spots) 平均分配到天數 (days)。
    若無法整除，餘數優先分配給前幾天。
    邏輯：計算單日基礎檔次，再處理餘數。
    """
    if days <= 0:
        return []
    # 確保偶數邏輯 (若需要) - 此處原始邏輯似乎是以 2 為倍數基礎進行計算
    if total_spots % 2 != 0:
        total_spots += 1
    base, rem = divmod(total_spots // 2, days)
    sch = [base + (1 if i < rem else 0) for i in range(days)]
    return [x * 2 for x in sch]


def split_period_by_months(start_dt, end_dt):
    """
    將走期依「日曆月」切為多段。若僅單月則回傳 [(start_dt, end_dt)]；
    若跨月則回傳 [(第一月start, 第一月end), (第二月start, 第二月end), ...]。
    用於 CUE 表走期超過一個月時分上下月顯示。
    """
    start_date = start_dt.date() if hasattr(start_dt, "date") else start_dt
    end_date = end_dt.date() if hasattr(end_dt, "date") else end_dt
    out = []
    cur = start_date
    while cur <= end_date:
        if cur.month == 12:
            last_in_month = date(cur.year, 12, 31)
        else:
            last_in_month = date(cur.year, cur.month + 1, 1) - timedelta(days=1)
        seg_end = min(last_in_month, end_date)
        out.append((cur, seg_end))
        if seg_end >= end_date:
            break
        cur = seg_end + timedelta(days=1)
    if hasattr(start_dt, "time") and callable(getattr(start_dt, "time", None)):
        def to_dt(d):
            if hasattr(d, "hour"):
                return d
            return datetime.combine(d, start_dt.time()) if isinstance(start_dt, datetime) else d
        return [(to_dt(a), to_dt(b)) for a, b in out]
    return out


def expand_schedule_to_calendar(schedule_active, segments, start_date, end_date):
    """
    將「僅含執行日」的排程展開為完整日曆（start_date～end_date）。
    segments: 波段 list of (start_dt, end_dt)，須在 [start_date, end_date] 內。
    未執行日填 0（報表顯示為空白）。
    """
    total_days = (end_date - start_date).days + 1
    result = [0] * total_days
    active_idx = 0
    for seg_start, seg_end in sorted(segments):
        d = seg_start
        while d <= seg_end:
            day_index = (d - start_date).days
            if 0 <= day_index < total_days and active_idx < len(schedule_active):
                result[day_index] = schedule_active[active_idx]
            active_idx += 1
            d += timedelta(days=1)
    return result


def get_remarks_text(sign_deadline, billing_month, payment_date):
    """生成標準合約備註文字 (包含日期填空)。"""
    d_str = sign_deadline.strftime("%Y/%m/%d (%a)") if sign_deadline else "____/__/__ (__)"
    p_str = payment_date.strftime("%Y/%m/%d") if payment_date else "____/__/__"
    return [
        f"1.請於 {d_str} 11:30前 回簽及進單，方可順利上檔。",
        "2.以上節目名稱如有異動，以上檔時節目名稱為主，如遇電台時段滿檔，上檔時間挪後或更換至同級時段。",
        "3.通路店鋪數與開機率開機率至少七成(以上)。每日因加盟數調整，或遇店舖年度季度改裝、設備維護升級及保修等狀況，會有一定幅度增減。",
        "4.託播方需於上檔前 5 個工作天，提供廣告帶(mp3)、影片/影像 1920x1080 (mp4)。",
        f"5.雙方同意費用請款月份 : {billing_month}，如有修正必要，將另行E-Mail告知，並視為正式合約之一部分。",
        f"6.付款兌現日期：{p_str}"
    ]


def format_campaign_details(config):
    """將複雜的媒體設定 config 轉為字串，以便存入 Ragic 的「詳細設定」欄位。"""
    details = []
    for media, settings in config.items():
        sec_str = ", ".join([f"{s}秒({p}%)" for s, p in settings.get("sec_shares", {}).items()])
        reg_str = "全省聯播" if settings.get("is_national") else "/".join(settings.get("regions", []))
        info = f"【{media}】 預算佔比: {settings.get('share')}% | 秒數分配: {sec_str} | 區域: {reg_str}"
        details.append(info)
    return "\n".join(details)


# 投放參數詳情擴充：Ragic details 欄位中 [EXT] 之後的 JSON 鍵名（供寫入/還原對照）
RAGIC_EXT_KEYS = [
    "is_barter_contract", "cue_mode", "temp_format_type", "use_date_segments", "date_segments",
    "apply_rebate", "apply_nat_rad_rebate", "rebate_nat_destination", "apply_nat_cf_rebate",
    "apply_region_rad_rebate", "rebate_region_destination",
    "bonus_rebate_pct", "bonus_cb_rad", "bonus_rad_nat", "bonus_rad_reg", "bonus_rad_share", "bonus_rad_sec",
    "bonus_cb_fv", "bonus_fv_nat", "bonus_fv_reg", "bonus_fv_share", "bonus_fv_sec",
    "bonus_cb_cf", "bonus_cf_share", "bonus_cf_sec",
    "cue_sign_deadline", "cue_billing_month", "cue_payment_date", "cue_live_update_remarks", "cue_remarks_text",
    "cue_seg_remarks", "custom_bonus",
    "rad_use_region_share", "fv_use_region_share", "_rad_region_list", "_fv_region_list",
]


def build_ragic_details_full(config, state_dict):
    """
    組裝「投放參數詳情」完整字串：媒體區塊（與舊版相容）+ [EXT] + JSON（交換/回饋/分波段/備註等），
    供上傳至 Ragic 的 details 欄位，以利 100% 還原。
    state_dict: 來自 app 的 session_state 子集，僅含可序列化鍵；日期請先轉成 str。
    """
    media_block = format_campaign_details(config)
    # 只取出可序列化的 state，且 date/datetime 轉字串
    def _to_serializable(obj):
        if obj is None:
            return None
        if isinstance(obj, (date, datetime)) and not isinstance(obj, type):
            return obj.isoformat()
        if isinstance(obj, list):
            return [_to_serializable(x) for x in obj]
        if isinstance(obj, dict):
            return {str(kk): _to_serializable(vv) for kk, vv in obj.items()}
        return obj

    out = {}
    for k, v in state_dict.items():
        if v is None:
            continue
        out[k] = _to_serializable(v)
    ext_json = json.dumps(out, ensure_ascii=False, default=str)
    return media_block + "\n[EXT]\n" + ext_json


def build_platform_detail_text(rows, config):
    """
    產出 Ragic「平台細項」欄位文字：只要有投放的平台/區域（含回饋、加贈等插入列）就記錄。
    - 家樂福：有出現就顯示「家樂福」
    - 全家廣播：全省（is_national 或 regions=['全省']）顯示「全家全頻」；非全省則依區域顯示「全家北頻/桃頻/中頻/南頻/高頻/宜頻」
    - 新鮮視：全省顯示「新鮮視全區」；非全省則依區域顯示「新鮮視北北基/桃竹苗/中彰投/雲嘉南/高高屏/宜花東」
    """
    if not rows:
        return ""

    region_to_family = {
        "北區": "北頻",
        "桃竹苗": "桃頻",
        "中區": "中頻",
        "雲嘉南": "南頻",
        "高屏": "高頻",
        "東區": "宜頻",
    }
    region_to_fv = {
        "北區": "北北基",
        "桃竹苗": "桃竹苗",
        "中區": "中彰投",
        "雲嘉南": "雲嘉南",
        "高屏": "高高屏",
        "東區": "宜花東",
    }

    # 先從 rows 收集實際出現的平台/區域（含回饋/加贈等插入列）
    seen = {"全家廣播": set(), "新鮮視": set(), "家樂福": False}
    for r in rows:
        media = r.get("media")
        if media == "家樂福":
            seen["家樂福"] = True
            continue
        region = r.get("region")
        if media in ("全家廣播", "新鮮視") and region:
            seen[media].add(str(region))

    out = []

    # 全家廣播：若 config 判定為全省，顯示「全家全頻」；否則依實際出現區域列出
    rad_cfg = config.get("全家廣播") if isinstance(config, dict) else None
    rad_is_nat = bool(rad_cfg and (rad_cfg.get("is_national") or rad_cfg.get("regions") == ["全省"]))
    if seen["全家廣播"]:
        if rad_is_nat:
            out.append("全家全頻")
        else:
            order = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]
            for reg in order:
                if reg in seen["全家廣播"] and reg in region_to_family:
                    out.append(f"全家{region_to_family[reg]}")

    # 家樂福
    if seen["家樂福"]:
        out.append("家樂福")

    # 新鮮視：若 config 判定為全省，顯示「新鮮視全區」；否則依實際出現區域列出
    fv_cfg = config.get("新鮮視") if isinstance(config, dict) else None
    fv_is_nat = bool(fv_cfg and (fv_cfg.get("is_national") or fv_cfg.get("regions") == ["全省"]))
    if seen["新鮮視"]:
        if fv_is_nat:
            out.append("新鮮視全區")
        else:
            order = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]
            for reg in order:
                if reg in seen["新鮮視"] and reg in region_to_fv:
                    out.append(f"新鮮視{region_to_fv[reg]}")

    # 去重並保序，上傳 Ragic 時以逗號隔開
    dedup = []
    for x in out:
        if x not in dedup:
            dedup.append(x)
    return ",".join(dedup)


def build_platform_text(rows):
    """
    產出 Ragic「平台」欄位文字（逗號分隔）：
    - 全家廣播相關 -> 全家企頻
    - 家樂福 -> 家樂福企頻
    - 新鮮視相關 -> 新鮮視
    """
    if not rows:
        return ""
    media_map = {
        "全家廣播": "全家企頻",
        "家樂福": "家樂福企頻",
        "新鮮視": "新鮮視",
    }
    order = ["全家企頻", "家樂福企頻", "新鮮視"]
    seen = set()
    for r in rows:
        media = r.get("media")
        label = media_map.get(media)
        if label:
            seen.add(label)
    return ",".join([x for x in order if x in seen])
