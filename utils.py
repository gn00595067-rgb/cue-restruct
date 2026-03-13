"""
工具函式模組 (Utility Functions)
包含字串處理、數學計算、HTML 轉義等通用功能
"""

import re
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
