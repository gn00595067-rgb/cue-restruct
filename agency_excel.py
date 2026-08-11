# -*- coding: utf-8 -*-
"""
代理商 CUE Excel 渲染模組 (Agency Cue Excel Renderer)

以 openpyxl 產出三種代理商版型（2008傳媒 / D drive / 凱絡），
每平台一個工作表。字型一律「微軟正黑體」、橫向 A4、fitToPage、金額會計格式。

版型布局取自實際範例檔實測（2008）與開發規格（D drive / 凱絡）。
為求 LibreOffice 轉 PDF 穩定，數字一律寫入計算後的實值（非公式）。
"""
import os
import calendar
from io import BytesIO
from datetime import date

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.drawing.image import Image as XLImage

import config
import agency_cue as ac

FONT = config.FONT_MAIN

# 代理商 Logo（隨專案打包於 assets/，確保離線與 PDF 轉檔皆可渲染）
_ASSET_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
LOGO_2008 = os.path.join(_ASSET_DIR, "logo_2008.png")
LOGO_DDRIVE = os.path.join(_ASSET_DIR, "logo_ddrive.png")


def _add_logo(ws, path, anchor_cell, width, height):
    """在指定儲存格錨點放置 Logo 圖片（單元錨定，固定顯示尺寸，單位 px）。"""
    if not os.path.exists(path):
        return
    img = XLImage(path)
    img.width = width
    img.height = height
    img.anchor = anchor_cell
    ws.add_image(img)
MONEY_FMT = '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"??_);_(@_)'
INT_FMT = "#,##0"

WEEKEND_FILL = PatternFill(fill_type="solid", fgColor="FFF2CC")  # 凱絡週末底色黃
_thin = Side(style="thin", color="000000")
_med = Side(style="medium", color="000000")
BORDER_THIN = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
BORDER_MED = Border(left=_med, right=_med, top=_med, bottom=_med)


# =============================================================================
# 共用工具
# =============================================================================
def _cn_weekday(d):
    return "一二三四五六日"[d.weekday()]


def _en_weekday(d):
    return "MTWTFSS"[d.weekday()]


def _set(ws, coord, value, size=12, bold=False, align="center", valign="center",
         wrap=False, fmt=None, fill=None):
    c = ws[coord]
    c.value = value
    c.font = Font(name=FONT, size=size, bold=bold)
    c.alignment = Alignment(horizontal=align, vertical=valign, wrap_text=wrap)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    return c


def _merge(ws, rng):
    ws.merge_cells(rng)


def _border_range(ws, min_col, min_row, max_col, max_row, border=BORDER_THIN):
    for r in range(min_row, max_row + 1):
        for cc in range(min_col, max_col + 1):
            ws.cell(row=r, column=cc).border = border


def _net_cell(ws, coord, net_display, size, bold=False, fill=None):
    """寫入實收欄：數字用會計格式，字串（專案回饋/計價於量販）用文字。"""
    if isinstance(net_display, str):
        _set(ws, coord, net_display, size=size, bold=bold, fill=fill)
    else:
        _set(ws, coord, net_display, size=size, bold=bold, fmt=MONEY_FMT, fill=fill)


def _day_columns(start_col, days):
    """回傳 [(col_index, date_offset), ...] 對應每一天的欄。"""
    return [(start_col + i, i) for i in range(days)]


def _write_day_header(ws, start_col, start_dt, days, row_month, row_date, row_wd,
                      weekday_fn, shade_weekend=False, size=12):
    """
    寫日期表頭三列：月份(合併同月)、日號、星期。
    weekday_fn: 產生星期字元的函式（中/英）。
    """
    from datetime import timedelta
    # 月份列（同月合併）
    seg_start = 0
    for i in range(days + 1):
        cur = None if i >= days else (start_dt + timedelta(days=i))
        prev = start_dt + timedelta(days=i - 1) if i > 0 else None
        boundary = (i == days) or (prev is not None and cur.month != prev.month)
        if boundary and i > 0:
            c0 = start_col + seg_start
            c1 = start_col + i - 1
            mon = (start_dt + timedelta(days=seg_start)).month
            _set(ws, f"{get_column_letter(c0)}{row_month}", calendar.month_abbr[mon],
                 size=size, bold=True)
            if c1 > c0:
                _merge(ws, f"{get_column_letter(c0)}{row_month}:{get_column_letter(c1)}{row_month}")
            seg_start = i
    # 日號 + 星期
    for i in range(days):
        d = start_dt + timedelta(days=i)
        col = get_column_letter(start_col + i)
        fill = WEEKEND_FILL if (shade_weekend and d.weekday() >= 5) else None
        _set(ws, f"{col}{row_date}", d.day, size=size, bold=True, fill=fill)
        _set(ws, f"{col}{row_wd}", weekday_fn(d), size=size, fill=fill)


def _period_str(a, b):
    return f"{a.strftime('%Y.%m.%d')}-{b.strftime('%Y.%m.%d')}"


def _group_rows(rows):
    """把 rows 分組：main(+緊接的 comp) 為一組；其餘各自一組。"""
    groups = []
    i = 0
    while i < len(rows):
        r = rows[i]
        if r["kind"] == ac.KIND_MAIN and i + 1 < len(rows) and rows[i + 1]["kind"] == ac.KIND_COMP:
            groups.append([r, rows[i + 1]])
            i += 2
        else:
            groups.append([r])
            i += 1
    return groups


# =============================================================================
# 2008 版型
# =============================================================================
def _render_2008(wb, sheet, model, made_date):
    days = (model["end_date"] - model["start_date"]).days + 1
    start_dt = model["start_date"]
    sec = sheet["seconds"]
    ws = wb.create_sheet(title=_sheet_name(model, sheet))
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9  # A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    DAY0 = 9  # 日欄自 I 起
    widths = {"A": 48.6, "B": 22.0, "C": 31.6, "D": 33.6, "E": 20.9, "F": 32.6, "G": 36.6, "H": 29.4}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    for i in range(days):
        ws.column_dimensions[get_column_letter(DAY0 + i)].width = 8.9
    last_col = DAY0 + days - 1

    # 2008 Logo（右上，比照範例錨定於日欄右側上方）
    logo_col = get_column_letter(max(DAY0, last_col - 6))
    _add_logo(ws, LOGO_2008, f"{logo_col}2", width=200, height=79)

    # 表頭 1~6（28pt bold）
    header_pairs = [
        ("Client", model["client_name"]),
        ("Media Agency", "2008傳媒行銷股份有限公司"),
        ("Advertising Agency", ""),
        ("Product", model["product_name"]),
        ("Campaign", model.get("campaign", "")),
        ("Period", _period_str(model["start_date"], model["end_date"])),
    ]
    for idx, (lab, val) in enumerate(header_pairs):
        r = idx + 1
        ws.row_dimensions[r].height = 34
        _set(ws, f"A{r}", lab, size=22, bold=True, align="center")
        _set(ws, f"C{r}", val, size=22, bold=True, align="left")
        _merge(ws, f"C{r}:H{r}")

    # 欄位表頭 8~10
    for r in (8, 9, 10):
        ws.row_dimensions[r].height = 34
    heads = [("A", "媒體型態"), ("B", "地區"), ("C", "播出時段"), ("D", "定價"),
             ("E", "單位"), ("F", "次數"), ("G", "合計"), ("H", "素材\n提供時間")]
    for col, txt in heads:
        _set(ws, f"{col}8", txt, size=18, bold=True, wrap=True)
        _merge(ws, f"{col}8:{col}10")
    _write_day_header(ws, DAY0, start_dt, days, 8, 9, 10, _cn_weekday, size=16)
    _border_range(ws, 1, 8, last_col, 10)

    day_cols = _day_columns(DAY0, days)

    # 資料列
    r = 11
    groups = _group_rows(sheet["rows"])
    data_top = r
    schedule_rows = []  # (excel_row, schedule) 供合計每日加總（僅逐日排檔列）
    total_spots = 0
    total_list = 0
    total_net = 0
    for grp in groups:
        gtop = r
        gbot = r + len(grp) - 1
        main = grp[0]
        for k, row in enumerate(grp):
            rr = r + k
            ws.row_dimensions[rr].height = 46
            _set(ws, f"F{rr}", row["spots"], size=16, fmt="0")
            total_spots += row["spots"]
            # 日欄
            if row["schedule"] is None:
                c0 = get_column_letter(DAY0)
                c1 = get_column_letter(last_col)
                _set(ws, f"{c0}{rr}", row["spots"], size=16)
                if last_col > DAY0:
                    _merge(ws, f"{c0}{rr}:{c1}{rr}")
            else:
                for (cidx, off) in day_cols:
                    _set(ws, f"{get_column_letter(cidx)}{rr}", row["schedule"][off], size=14)
                schedule_rows.append((rr, row["schedule"]))
            if isinstance(row["net_display"], (int, float)):
                total_net += row["net_display"]
        # 合併欄（整組）
        _set(ws, f"A{gtop}", main["media_label"], size=16, wrap=True)
        _set(ws, f"B{gtop}", main["region_label"], size=16)
        _set(ws, f"C{gtop}", main["daypart"], size=16)
        _set(ws, f"D{gtop}", main["list_total"], size=16, fmt=MONEY_FMT)
        _set(ws, f"E{gtop}", f"{sec}秒", size=16)
        _net_cell(ws, f"G{gtop}", main["net_display"], size=16)
        _set(ws, f"H{gtop}", main["material"], size=16, wrap=True)
        total_list += main["list_total"]
        if gbot > gtop:
            for col in ("A", "B", "C", "D", "E", "G", "H"):
                _merge(ws, f"{col}{gtop}:{col}{gbot}")
        r = gbot + 1

    data_bot = r - 1
    # 合計列
    ws.row_dimensions[r].height = 40
    _set(ws, f"B{r}", "合計", size=16, bold=True)
    _merge(ws, f"B{r}:C{r}")
    _set(ws, f"D{r}", total_list, size=16, bold=True, fmt=MONEY_FMT)
    _set(ws, f"F{r}", total_spots, size=16, bold=True, fmt=INT_FMT)
    _set(ws, f"G{r}", total_net, size=16, bold=True, fmt=MONEY_FMT)
    for (cidx, off) in day_cols:
        day_sum = sum(sch[off] for (_, sch) in schedule_rows)
        _set(ws, f"{get_column_letter(cidx)}{r}", day_sum, size=14, bold=True, fmt=INT_FMT)
    _border_range(ws, 1, data_top, last_col, r)
    total_row = r

    # 費用區（F 標籤 / G 值）
    f = sheet["fees"]
    fr = total_row + 2
    fee_lines = [
        ("Budget (net)：", f["budget_net"] if isinstance(f["budget_net"], (int, float)) else f["budget_net"]),
        (f"AC {int(f['ac_pct'])}%：", f["ac"]),
        ("5% Tax ：", f["tax"]),
        ("TOTAL ：", f["total"]),
    ]
    for i, (lab, val) in enumerate(fee_lines):
        rr = fr + i
        ws.row_dimensions[rr].height = 30
        _set(ws, f"F{rr}", lab, size=16, bold=True, align="right")
        _net_cell(ws, f"G{rr}", val, size=16, bold=True)

    # 備註（A 欄，逐行）
    rr = fr + len(fee_lines) + 1
    for line in model.get("remarks", []):
        ws.row_dimensions[rr].height = 26
        _set(ws, f"A{rr}", line, size=16, bold=True, align="left")
        _merge(ws, f"A{rr}:E{rr}")
        rr += 1
    return ws


# =============================================================================
# D drive 版型
# =============================================================================
def _render_ddrive(wb, sheet, model, made_date):
    from datetime import timedelta
    days = (model["end_date"] - model["start_date"]).days + 1
    start_dt = model["start_date"]
    sec = sheet["seconds"]
    ws = wb.create_sheet(title=_sheet_name(model, sheet))
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    DAY0 = 9  # 日欄自 I 起
    widths = {"A": 31.6, "B": 25.5, "C": 18.1, "D": 19.9, "E": 15.9, "F": 21.4, "G": 23.5, "H": 26.0}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    for i in range(days):
        ws.column_dimensions[get_column_letter(DAY0 + i)].width = 8.5
    last_col = DAY0 + days - 1

    # M Drive Logo（左上），並將上方列高撐開以容納
    for rr in (1, 2, 3, 4):
        ws.row_dimensions[rr].height = 22
    _add_logo(ws, LOGO_DDRIVE, "A1", width=100, height=89)

    # 表頭 5~7
    _set(ws, "A5", f"客戶：{model['client_name']}", size=16, bold=True, align="left")
    _merge(ws, "A5:D5")
    _set(ws, "A6", f"產品：{model['product_name']}", size=16, bold=True, align="left")
    _merge(ws, "A6:D6")
    _set(ws, "A7", f"刊期：{model['start_date'].strftime('%Y/%m/%d')} ~ {model['end_date'].strftime('%Y/%m/%d')}",
         size=16, bold=True, align="left")
    _merge(ws, "A7:D7")

    # 欄位表頭 9~11
    for r in (9, 10, 11):
        ws.row_dimensions[r].height = 26
    heads = [("A", "媒體"), ("B", "地區"), ("C", "託播秒數"), ("D", "播出時段"),
             ("E", "次數"), ("F", "素材提供時間"), ("G", "定價(Net Cost)"), ("H", "專案執行價(Net Cost)")]
    for col, txt in heads:
        _set(ws, f"{col}9", txt, size=14, bold=True, wrap=True)
        _merge(ws, f"{col}9:{col}11")
    _write_day_header(ws, DAY0, start_dt, days, 9, 10, 11,
                      lambda d: _cn_weekday(d), size=12)
    _border_range(ws, 1, 9, last_col, 11)

    day_cols = _day_columns(DAY0, days)
    r = 12
    data_top = r
    schedule_rows = []
    total_spots = 0
    total_list = 0
    total_net = 0

    # 依平台分：萬家福時 量販+超市 的 C/F/H 合併
    is_wjf = sheet["platform"] == ac.PLATFORM_WJF
    for row in sheet["rows"]:
        ws.row_dimensions[r].height = 30
        _set(ws, f"A{r}", row["media_label"], size=12, wrap=True, align="left")
        _set(ws, f"B{r}", row["region_label"], size=12)
        _set(ws, f"C{r}", f"{sec}秒", size=12)
        _set(ws, f"D{r}", row["daypart"], size=12)
        _set(ws, f"E{r}", row["spots"], size=12, fmt="0")
        _set(ws, f"F{r}", row["material"], size=12)
        # G 定價 = 牌價 × 次數
        gval = row["list_total"] if row["kind"] not in (ac.KIND_SUPER, ac.KIND_SUPER_REBATE) else ac.NET_ON_MAG
        _net_cell(ws, f"G{r}", gval, size=12)
        _net_cell(ws, f"H{r}", row["net_display"], size=12,
                  bold=(row["kind"] == ac.KIND_MAIN))
        total_spots += row["spots"]
        if isinstance(row["list_total"], (int, float)) and row["kind"] not in (ac.KIND_SUPER, ac.KIND_SUPER_REBATE):
            total_list += row["list_total"]
        if isinstance(row["net_display"], (int, float)):
            total_net += row["net_display"]
        if row["schedule"] is None:
            c0 = get_column_letter(DAY0); c1 = get_column_letter(last_col)
            _set(ws, f"{c0}{r}", row["spots"], size=12)
            if last_col > DAY0:
                _merge(ws, f"{c0}{r}:{c1}{r}")
        else:
            for (cidx, off) in day_cols:
                _set(ws, f"{get_column_letter(cidx)}{r}", row["schedule"][off], size=11)
            schedule_rows.append((r, row["schedule"]))
        r += 1

    # 小計列
    ws.row_dimensions[r].height = 28
    _set(ws, f"A{r}", "小計", size=12, bold=True)
    _merge(ws, f"A{r}:D{r}")
    _set(ws, f"E{r}", total_spots, size=12, bold=True, fmt=INT_FMT)
    _set(ws, f"G{r}", total_list, size=12, bold=True, fmt=MONEY_FMT)
    for (cidx, off) in day_cols:
        day_sum = sum(sch[off] for (_, sch) in schedule_rows)
        _set(ws, f"{get_column_letter(cidx)}{r}", day_sum, size=11, bold=True, fmt=INT_FMT)
    _border_range(ws, 1, data_top, last_col, r)
    total_row = r

    # 備註（左）＋ 請款
    rr = total_row + 2
    _set(ws, f"A{rr}", "備註：", size=12, bold=True, align="left")
    _merge(ws, f"A{rr}:F{rr}")
    rr += 1
    _set(ws, f"A{rr}", model.get("payment_note", ""), size=12, align="left")
    _merge(ws, f"A{rr}:F{rr}")

    # 費用框（右）
    f = sheet["fees"]
    fr = total_row + 2
    fee_lines = [("Total Net Cost", f["net"]), ("VAT (5%)", f["vat"]), ("Total Gross Cost", f["gross"])]
    for i, (lab, val) in enumerate(fee_lines):
        rrr = fr + i
        _set(ws, f"G{rrr}", lab, size=12, bold=True, align="right")
        _net_cell(ws, f"H{rrr}", val, size=12, bold=True)
        _border_range(ws, 7, rrr, 8, rrr)
    return ws


# =============================================================================
# 凱絡 版型
# =============================================================================
def _render_carat(wb, sheet, model, made_date):
    days = (model["end_date"] - model["start_date"]).days + 1
    start_dt = model["start_date"]
    sec = sheet["seconds"]
    ws = wb.create_sheet(title=_sheet_name(model, sheet))
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    DAY0 = 11  # 日欄自 K 起
    widths = {"A": 21.5, "B": 18.7, "C": 15.7, "D": 9.5, "E": 11.3, "F": 11.5,
              "G": 11.5, "H": 10.5, "I": 14.5, "J": 16.0}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    for i in range(days):
        ws.column_dimensions[get_column_letter(DAY0 + i)].width = 6.5
    last_col = DAY0 + days - 1

    # 標題列
    _set(ws, "A1", "凱絡媒體服務(股)公司廣播媒體排期表", size=16, bold=True, align="center")
    _merge(ws, f"A1:{get_column_letter(last_col)}1")
    made = made_date or model["start_date"]
    _set(ws, "I2", f"客 戶：{model['client_name']}", size=11, bold=True, align="left"); _merge(ws, "I2:J2")
    _set(ws, "I3", f"產 品：{model['product_name']}", size=11, bold=True, align="left"); _merge(ws, "I3:J3")
    _set(ws, "I4", f"日 期：{made.strftime('%Y/%m/%d')}", size=11, bold=True, align="left"); _merge(ws, "I4:J4")
    _set(ws, "A4", f"{start_dt.year}年{start_dt.month}月", size=12, bold=True, align="left")

    # 欄位表頭 5~7
    for r in (5, 6, 7):
        ws.row_dimensions[r].height = 22
    heads = [("A", "媒體別"), ("B", "地區"), ("C", "時段"), ("D", "素材"),
             ("E", "定價(檔/Net)"), ("F", "市場價(檔/Net)"), ("G", "統一價(檔/Net)"),
             ("H", "檔數"), ("I", "總價"), ("J", "專案價(Net)")]
    for col, txt in heads:
        _set(ws, f"{col}5", txt, size=11, bold=True, wrap=True)
        _merge(ws, f"{col}5:{col}7")
    _write_day_header(ws, DAY0, start_dt, days, 5, 6, 7, _en_weekday,
                      shade_weekend=True, size=10)
    _border_range(ws, 1, 5, last_col, 7)

    day_cols = _day_columns(DAY0, days)
    r = 8
    data_top = r
    schedule_rows = []
    media_value = 0
    actual_net = sheet["fees"]["subtotal"] if isinstance(sheet["fees"].get("subtotal"), (int, float)) else 0

    # A 欄以（量販+超市）一組合併
    groups = []
    i = 0
    rows = sheet["rows"]
    while i < len(rows):
        if rows[i]["kind"] in (ac.KIND_MAIN,) and i + 1 < len(rows) and rows[i + 1]["kind"] == ac.KIND_SUPER:
            groups.append([rows[i], rows[i + 1]]); i += 2
        elif rows[i]["kind"] == ac.KIND_REBATE and i + 1 < len(rows) and rows[i + 1]["kind"] == ac.KIND_SUPER_REBATE:
            groups.append([rows[i], rows[i + 1]]); i += 2
        else:
            groups.append([rows[i]]); i += 1

    for grp in groups:
        gtop = r
        for k, row in enumerate(grp):
            rr = r + k
            ws.row_dimensions[rr].height = 42
            _set(ws, f"B{rr}", row["region_label"], size=10, wrap=True)
            _set(ws, f"C{rr}", row["daypart"], size=10)
            _set(ws, f"D{rr}", f'{sec}"CM', size=10)
            _set(ws, f"E{rr}", row["list_per"], size=10, fmt=MONEY_FMT)
            _set(ws, f"F{rr}", row["market_per"], size=10, fmt=MONEY_FMT)
            _set(ws, f"G{rr}", row["uni_per"], size=10, fmt=MONEY_FMT)
            _set(ws, f"H{rr}", row["spots"], size=10, fmt="0")
            _set(ws, f"I{rr}", row["uni_total"], size=10, fmt=MONEY_FMT)
            _net_cell(ws, f"J{rr}", row["net_display"], size=10,
                      bold=(row["kind"] == ac.KIND_MAIN))
            media_value += row["uni_total"] if isinstance(row["uni_total"], (int, float)) else 0
            if row["schedule"] is None:
                c0 = get_column_letter(DAY0); c1 = get_column_letter(last_col)
                _set(ws, f"{c0}{rr}", row["spots"], size=10)
                if last_col > DAY0:
                    _merge(ws, f"{c0}{rr}:{c1}{rr}")
            else:
                from datetime import timedelta
                for (cidx, off) in day_cols:
                    dd = start_dt + timedelta(days=off)
                    fill = WEEKEND_FILL if dd.weekday() >= 5 else None
                    _set(ws, f"{get_column_letter(cidx)}{rr}", row["schedule"][off], size=9, fill=fill)
                schedule_rows.append((rr, row["schedule"]))
        _set(ws, f"A{gtop}", grp[0]["media_label"], size=10, wrap=True)
        if len(grp) > 1:
            _merge(ws, f"A{gtop}:A{r + len(grp) - 1}")
        r += len(grp)

    data_bot = r - 1
    _border_range(ws, 1, data_top, last_col, data_bot)

    # 媒體總價值 / 優惠總價值
    rr = data_bot + 1
    _set(ws, f"A{rr}", f"媒體總價值(NET)＝{media_value:,}", size=11, bold=True, align="left")
    _merge(ws, f"A{rr}:E{rr}")
    _set(ws, f"F{rr}", f"優惠總價值(NET)＝{media_value - actual_net:,}", size=11, bold=True, align="left")
    _merge(ws, f"F{rr}:H{rr}")

    # 費用框（右 I/J）
    f = sheet["fees"]
    ac_txt = "-" if f.get("ac_free") else f.get("ac", 0)
    fee_lines = [("Sub-Total", f["subtotal"]), ("A.C 3%", ac_txt),
                 ("VAT 5%", f["vat"]), ("Grand-Total", f["grand"])]
    for i, (lab, val) in enumerate(fee_lines):
        rrr = data_bot + 1 + i
        _set(ws, f"I{rrr}", lab, size=10, bold=True, align="right")
        _net_cell(ws, f"J{rrr}", val, size=10, bold=True)
        _border_range(ws, 9, rrr, 10, rrr)

    # 簽核列
    sr = data_bot + 2 + len(fee_lines)
    _set(ws, f"A{sr}", "部主管：______  課主管：______  媒體窗口：______  承辦PM：______",
         size=11, align="left")
    _merge(ws, f"A{sr}:H{sr}")

    # 備註
    rr = sr + 1
    _set(ws, f"A{rr}", "備 註", size=11, bold=True)
    for line in model.get("remarks", []):
        rr += 1
        _set(ws, f"B{rr}", line, size=10, align="left")
        _merge(ws, f"B{rr}:{get_column_letter(last_col)}{rr}")
    return ws


# =============================================================================
# 對外主函式
# =============================================================================
def _sheet_name(model, sheet):
    a = model["start_date"].strftime("%m%d")
    b = model["end_date"].strftime("%m%d")
    if model["agency"] == "2008傳媒":
        wan = round(sheet["budget"] / 10000) if sheet["budget"] else 0
        plat = "全家" if sheet["platform"] == ac.PLATFORM_FAMILY else "萬家福"
        name = f"{a}-{b}-{plat}-{wan}萬-{sheet['seconds']}秒"
    else:
        name = f"{sheet['platform']} {a}-{b}"
    # Excel 工作表名長度上限 31，且不可含 :\\/?*[]
    for ch in ':\\/?*[]':
        name = name.replace(ch, "")
    return name[:31]


def generate_agency_excel(model, made_date=None):
    """依 model 產出代理商 Excel（每平台一個工作表），回傳 bytes。"""
    wb = Workbook()
    wb.remove(wb.active)
    agency = model["agency"]
    for sheet in model["sheets"]:
        if agency == "2008傳媒":
            _render_2008(wb, sheet, model, made_date)
        elif agency == "D drive":
            _render_ddrive(wb, sheet, model, made_date)
        elif agency == "凱絡":
            _render_carat(wb, sheet, model, made_date)
        else:
            _render_2008(wb, sheet, model, made_date)
    if not wb.worksheets:
        wb.create_sheet(title="空")
    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()
