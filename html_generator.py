"""
HTML 預覽生成模組 (HTML Preview Generator)
將運算結果轉為簡易的 HTML 表格，供使用者在網頁上直接預覽
"""

from itertools import groupby
from datetime import timedelta
from decimal import Decimal, ROUND_HALF_UP
from utils import html_escape, split_period_by_months


def _round_half_up(value):
    return int(Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def _render_one_month_table(rows, days_in_month, month_start, month_end, full_total_days, format_type, header_cls, budget_month, total_list_month, grand_total_month, prod_month):
    """產出單一月份的一張 CUE 表 HTML（表頭+日期欄+資料+總計+頁尾）。"""
    cols_def = ["Station", "Location", "Program", "Day-part", "Size", "rate<br>(Net)", "Package-cost<br>(Net)"]
    if format_type == "聲活":
        cols_def = ["頻道", "播出地區", "播出店數", "播出時間", "秒數/規格", "單價", "金額"]
    elif format_type == "鉑霖":
        cols_def = ["頻道", "播出地區", "播出店數", "播出時間", "規格", "單價", "金額"]
    th_fixed = "".join([f"<th rowspan='2' class='{header_cls}'>{c}</th>" for c in cols_def])
    th_total_right = f"<th rowspan='2' class='{header_cls}' style='min-width:50px;'>Total<br>Spots</th>"
    date_th1, date_th2 = "", ""
    curr = month_start
    weekdays = ["一", "二", "三", "四", "五", "六", "日"]
    for _ in range(days_in_month):
        wd = curr.weekday()
        bg = "bg-weekend" if wd >= 5 else ""
        date_th1 += f"<th class='{header_cls} col_day'>{curr.day}</th>"
        date_th2 += f"<th class='{bg} col_day'>{weekdays[wd]}</th>"
        curr += timedelta(days=1)
    tbody = ""
    rows_sorted = sorted(rows, key=lambda x: ({"全家廣播": 1, "新鮮視": 2, "家樂福": 3}.get(x["media"], 9), x["seconds"]))
    daily_totals = [0] * days_in_month
    for key, group in groupby(rows_sorted, lambda x: (x['media'], x['seconds'], x.get('nat_pkg_display', 0), x.get('is_rebate', False), x.get('rebate_type', ''), x.get('is_bonus_rebate', False), x.get('is_custom_bonus', False))):
        g_list = list(group)
        g_size = len(g_list)
        is_pkg = g_list[0]['is_pkg_member']
        is_rebate = g_list[0].get('is_rebate', False)
        is_custom_bonus = g_list[0].get('is_custom_bonus', False)
        pkg_label = (g_list[0].get('pkg_display', '加贈') if is_custom_bonus else g_list[0].get('pkg_display', '回饋')) if (is_rebate or is_custom_bonus) else None
        for i, r in enumerate(g_list):
            tbody += "<tr>"
            rate = f"${r['rate_display']:,}" if isinstance(r['rate_display'], (int, float)) else r['rate_display']
            pkg_val_str = ""
            if is_rebate or is_custom_bonus:
                if i == 0:
                    pkg_val_str = f"<td style='text-align:center' rowspan='{g_size}'>{pkg_label}</td>"
                else:
                    pkg_val_str = ""
            elif is_pkg:
                if i == 0:
                    val = f"${r['nat_pkg_display']:,}"
                    pkg_val_str = f"<td style='text-align:center' rowspan='{g_size}'>{val}</td>"
                else:
                    pkg_val_str = ""
            else:
                val = f"${r['pkg_display']:,}" if isinstance(r['pkg_display'], (int, float)) else r['pkg_display']
                pkg_val_str = f"<td style='text-align:center'>{val}</td>"
            if format_type == "聲活":
                sec_txt = f"{r['seconds']}秒"
                tbody += f"<td>{r['media']}</td><td>{r['region']}</td><td>{r.get('program_num','')}</td><td>{r['daypart']}</td><td>{sec_txt}</td><td>{rate}</td>{pkg_val_str}"
            elif format_type == "鉑霖":
                tbody += f"<td>{r['media']}</td><td>{r['region']}</td><td>{r.get('program_num','')}</td><td>{r['daypart']}</td><td>{r['seconds']}秒</td><td>{rate}</td>{pkg_val_str}"
            else:
                tbody += f"<td>{r['media']}</td><td>{r['region']}</td><td>{r.get('program_num','')}</td><td>{r['daypart']}</td><td>{r['seconds']}</td><td>{rate}</td>{pkg_val_str}"
            row_spots_sum = 0
            for d_idx, d in enumerate(r['schedule'][:days_in_month]):
                cell_val = "" if (d == 0 or d is None) else d
                tbody += f"<td>{cell_val}</td>"
                v = d if isinstance(d, (int, float)) else 0
                row_spots_sum += v
                if d_idx < len(daily_totals):
                    daily_totals[d_idx] += v
            tbody += f"<td style='font-weight:bold; background-color:#f0f0f0;'>{row_spots_sum}</td></tr>"
    total_row_html = "<tr><td colspan='5' style='text-align:center; font-weight:bold; background-color:#e0e0e0;'>Total</td>"
    total_row_html += f"<td style='text-align:center; font-weight:bold; background-color:#e0e0e0;'>${total_list_month:,}</td>"
    total_row_html += f"<td style='text-align:center; font-weight:bold; background-color:#e0e0e0;'>${budget_month:,}</td>"
    grand_total_spots = 0
    for day_sum in daily_totals:
        grand_total_spots += day_sum
        total_row_html += f"<td style='font-weight:bold; background-color:#e0e0e0;'>{'' if day_sum == 0 else day_sum}</td>"
    total_row_html += f"<td style='font-weight:bold; background-color:#d0d0d0; border: 2px solid #000;'>{grand_total_spots}</td></tr>"
    vat_month = _round_half_up(budget_month * 0.05)
    footer_html = f"<div style='margin-top:10px; font-weight:bold; text-align:right;'>製作費: ${prod_month:,}<br>5% VAT: ${vat_month:,}<br>Grand Total: ${grand_total_month:,}</div>"
    # Period 日期格式修正：%Y.%m.%d（原本誤用 'm' 導致畫面顯示 2026.m.01）
    return f"<div style='margin-bottom:24px;'><div style='margin-bottom:4px; font-weight:bold;'>Period: {month_start.strftime('%Y.%m.%d')} - {month_end.strftime('%Y.%m.%d')}</div><table><thead><tr>{th_fixed}{date_th1}{th_total_right}</tr><tr>{date_th2}</tr></thead><tbody>{tbody}{total_row_html}</tbody></table>{footer_html}</div>"


def generate_html_preview(rows, days_cnt, start_dt, end_dt, c_name, tax_id, p_display, format_type, remarks, total_list, grand_total, budget, prod):
    eff_days = days_cnt
    total_days = (end_dt - start_dt).days + 1
    month_ranges = split_period_by_months(start_dt, end_dt)
    # 走期 31 天內（含跨月）不拆表，一張表呈現；超過 31 天且跨月才拆成多頁
    if len(month_ranges) > 1 and total_days > 31:
        # 多月：產出多個「獨立頁」HTML，每頁都有完整標題與下方備註
        header_cls = "bg-sh-head"
        if format_type == "東吳": header_cls = "bg-dw-head"
        elif format_type == "鉑霖": header_cls = "bg-bolin-head"
        unique_media = sorted(list(set([r['media'] for r in rows])))
        order_map = {"全家廣播": 1, "新鮮視": 2, "家樂福": 3}
        unique_media.sort(key=lambda x: order_map.get(x, 99))
        medium_str = "/".join(unique_media)
        remarks_html = "<br>".join([html_escape(x) for x in remarks])
        remarks_block = f"<div style='margin-top:10px; font-size:11px;'><b>Remarks：本排程表經雙方確認後視同合約之延伸，具同等法律約束力與效力</b><br>{remarks_html}</div>"
        client_info_base = f"<b>客戶名稱：</b>{html_escape(c_name)} &nbsp; <b>統編：</b>{html_escape(tax_id)} &nbsp; <b>Product：</b>{html_escape(p_display)} &nbsp; <b>Medium：</b>{html_escape(medium_str)}"
        css = "body { font-family: sans-serif; font-size: 10px; background-color: #ffffff; color: #000000; padding: 5px; } table { border-collapse: collapse; width: 100%; } th, td { border: 0.5pt solid #000; padding: 4px; text-align: center; white-space: nowrap; } .bg-dw-head { background-color: #4472C4; color: white; } .bg-sh-head { background-color: white; color: black; font-weight: bold; border-bottom: 2px solid black; } .bg-bolin-head { background-color: #F8CBAD; color: black; } .bg-weekend { background-color: #FFFFCC; } .right { text-align: right; }"
        pages = []
        for m_start, m_end in month_ranges:
            try:
                day_offset = (m_start - start_dt).days
            except Exception:
                day_offset = 0
            day_offset = max(0, day_offset)
            days_in_month = (m_end - m_start).days + 1
            ratio = days_in_month / total_days if total_days else 0
            rows_month = []
            for r in rows:
                r2 = dict(r)
                sch = r.get("schedule", [])
                r2["schedule"] = sch[day_offset:day_offset + days_in_month]
                if isinstance(r.get("rate_display"), (int, float)):
                    r2["rate_display"] = int(round(r["rate_display"] * ratio))
                if isinstance(r.get("pkg_display"), (int, float)):
                    r2["pkg_display"] = int(round(r["pkg_display"] * ratio))
                if isinstance(r.get("nat_pkg_display"), (int, float)):
                    r2["nat_pkg_display"] = int(round(r["nat_pkg_display"] * ratio))
                rows_month.append(r2)
            total_list_month = int(round(total_list * ratio))
            budget_month = int(round(budget * ratio))
            grand_total_month = int(round((budget + _round_half_up(budget * 0.05)) * ratio))
            prod_month = int(round(prod * ratio))
            table_html = _render_one_month_table(rows_month, days_in_month, m_start, m_end, total_days, format_type, header_cls, budget_month, total_list_month, grand_total_month, prod_month)
            page_header = f"<div style='margin-bottom:10px;'>{client_info_base}<br><b>Period：</b>{m_start.strftime('%Y.%m.%d')} - {m_end.strftime('%Y.%m.%d')}</div>"
            one_page = f"<html><head><style>{css}</style></head><body>{page_header}<div style='overflow-x:auto;'>{table_html}</div>{remarks_block}</body></html>"
            pages.append(one_page)
        return pages

    # 單月：維持原本單一表格
    header_cls = "bg-sh-head"
    if format_type == "東吳":
        header_cls = "bg-dw-head"
    elif format_type == "鉑霖":
        header_cls = "bg-bolin-head"

    # 生成日期欄位 (區分週末顏色)
    date_th1, date_th2 = "", ""
    curr = start_dt
    weekdays = ["一", "二", "三", "四", "五", "六", "日"]
    for i in range(eff_days):
        wd = curr.weekday()
        bg = "bg-weekend" if wd >= 5 else ""
        date_th1 += f"<th class='{header_cls} col_day'>{curr.day}</th>"
        date_th2 += f"<th class='{bg} col_day'>{weekdays[wd]}</th>"
        curr += timedelta(days=1)

    # 定義欄位名稱 (依格式切換)
    cols_def = ["Station", "Location", "Program", "Day-part", "Size", "rate<br>(Net)", "Package-cost<br>(Net)"]
    if format_type == "聲活":
        cols_def = ["頻道", "播出地區", "播出店數", "播出時間", "秒數/規格", "單價", "金額"]
    elif format_type == "鉑霖":
        cols_def = ["頻道", "播出地區", "播出店數", "播出時間", "規格", "單價", "金額"]

    th_fixed = "".join([f"<th rowspan='2' class='{header_cls}'>{c}</th>" for c in cols_def])
    th_total_right = f"<th rowspan='2' class='{header_cls}' style='min-width:50px;'>Total<br>Spots</th>"

    # 媒體排序邏輯
    unique_media = sorted(list(set([r['media'] for r in rows])))
    order_map = {"全家廣播": 1, "新鮮視": 2, "家樂福": 3}
    unique_media.sort(key=lambda x: order_map.get(x, 99))
    medium_str = "/".join(unique_media)

    tbody = ""
    # 排序資料列，確保同一媒體與秒數在一起
    rows_sorted = sorted(rows, key=lambda x: ({"全家廣播": 1, "新鮮視": 2, "家樂福": 3}.get(x["media"], 9), x["seconds"]))
    daily_totals = [0] * eff_days

    # 分組繪製 (處理合併儲存格邏輯)
    for key, group in groupby(rows_sorted, lambda x: (x['media'], x['seconds'], x.get('nat_pkg_display', 0), x.get('is_rebate', False), x.get('rebate_type', ''), x.get('is_bonus_rebate', False), x.get('is_custom_bonus', False))):
        g_list = list(group)
        g_size = len(g_list)
        is_pkg = g_list[0]['is_pkg_member']
        is_rebate = g_list[0].get('is_rebate', False)
        is_custom_bonus = g_list[0].get('is_custom_bonus', False)
        pkg_label = (g_list[0].get('pkg_display', '加贈') if is_custom_bonus else g_list[0].get('pkg_display', '回饋')) if (is_rebate or is_custom_bonus) else None
        for i, r in enumerate(g_list):
            tbody += "<tr>"
            rate = f"${r['rate_display']:,}" if isinstance(r['rate_display'], (int, float)) else r['rate_display']
            pkg_val_str = ""
            if is_rebate or is_custom_bonus:
                if i == 0:
                    pkg_val_str = f"<td style='text-align:center' rowspan='{g_size}'>{pkg_label}</td>"
                else:
                    pkg_val_str = ""
            elif is_pkg:
                if i == 0:
                    val = f"${r['nat_pkg_display']:,}"
                    pkg_val_str = f"<td style='text-align:center' rowspan='{g_size}'>{val}</td>"
                else:
                    pkg_val_str = ""
            else:
                val = f"${r['pkg_display']:,}" if isinstance(r['pkg_display'], (int, float)) else r['pkg_display']
                pkg_val_str = f"<td style='text-align:center'>{val}</td>"

            # 填充列內容
            if format_type == "聲活":
                sec_txt = f"{r['seconds']}秒"
                tbody += f"<td>{r['media']}</td><td>{r['region']}</td><td>{r.get('program_num','')}</td><td>{r['daypart']}</td><td>{sec_txt}</td><td>{rate}</td>{pkg_val_str}"
            elif format_type == "鉑霖":
                tbody += f"<td>{r['media']}</td><td>{r['region']}</td><td>{r.get('program_num','')}</td><td>{r['daypart']}</td><td>{r['seconds']}秒</td><td>{rate}</td>{pkg_val_str}"
            else:
                tbody += f"<td>{r['media']}</td><td>{r['region']}</td><td>{r.get('program_num','')}</td><td>{r['daypart']}</td><td>{r['seconds']}</td><td>{rate}</td>{pkg_val_str}"

            # 填入每日檔次（未執行日為 0 時顯示空白）
            row_spots_sum = 0
            for d_idx, d in enumerate(r['schedule'][:eff_days]):
                cell_val = "" if (d == 0 or d is None) else d
                tbody += f"<td>{cell_val}</td>"
                v = d if isinstance(d, (int, float)) else 0
                row_spots_sum += v
                if d_idx < len(daily_totals):
                    daily_totals[d_idx] += v
            tbody += f"<td style='font-weight:bold; background-color:#f0f0f0;'>{row_spots_sum}</td></tr>"

    # 總計列（Package-cost Total 刻意用預算 budget，讓客戶看到折扣感；檔次則依實作價計算）
    total_row_html = "<tr><td colspan='5' style='text-align:center; font-weight:bold; background-color:#e0e0e0;'>Total</td>"
    total_row_html += f"<td style='text-align:center; font-weight:bold; background-color:#e0e0e0;'>${total_list:,}</td>"
    total_row_html += f"<td style='text-align:center; font-weight:bold; background-color:#e0e0e0;'>${budget:,}</td>"
    grand_total_spots = 0
    for day_sum in daily_totals:
        grand_total_spots += day_sum
        total_row_html += f"<td style='font-weight:bold; background-color:#e0e0e0;'>{'' if day_sum == 0 else day_sum}</td>"
    total_row_html += f"<td style='font-weight:bold; background-color:#d0d0d0; border: 2px solid #000;'>{grand_total_spots}</td></tr>"

    remarks_html = "<br>".join([html_escape(x) for x in remarks])
    vat = _round_half_up(budget * 0.05)
    footer_html = f"<div style='margin-top:10px; font-weight:bold; text-align:right;'>製作費: ${prod:,}<br>5% VAT: ${vat:,}<br>Grand Total: ${grand_total:,}</div>"

    # CSS 樣式
    css = """
    body { font-family: sans-serif; font-size: 10px; background-color: #ffffff; color: #000000; padding: 5px; }
    table { border-collapse: collapse; width: 100%; background-color: #ffffff; }
    th, td { border: 0.5pt solid #000; padding: 4px; text-align: center; white-space: nowrap; color: #000000; }
    .bg-dw-head { background-color: #4472C4; color: white; }
    .bg-sh-head { background-color: white; color: black; font-weight: bold; border-bottom: 2px solid black; }
    .bg-bolin-head { background-color: #F8CBAD; color: black; }
    .bg-weekend { background-color: #FFFFCC; }
    """

    client_info_html = f"<b>客戶名稱：</b>{html_escape(c_name)} &nbsp; <b>統編：</b>{html_escape(tax_id)}"

    # Period 日期格式修正：%Y.%m.%d（原本誤用 'm' 導致畫面顯示 2026.m.01）
    return f"<html><head><style>{css}</style></head><body><div style='margin-bottom:10px;'>{client_info_html} &nbsp; <b>Product：</b>{html_escape(p_display)}<br><b>Period：</b>{start_dt.strftime('%Y.%m.%d')} - {end_dt.strftime('%Y.%m.%d')} &nbsp; <b>Medium：</b>{html_escape(medium_str)}</div><div style='overflow-x:auto;'><table><thead><tr>{th_fixed}{date_th1}{th_total_right}</tr><tr>{date_th2}</tr></thead><tbody>{tbody}{total_row_html}</tbody></table></div>{footer_html}<div style='margin-top:10px; font-size:11px;'><b>Remarks：本排程表經雙方確認後視同合約之延伸，具同等法律約束力與效力</b><br>{remarks_html}</div></body></html>"
