# -*- coding: utf-8 -*-
"""
代理商 CUE 使用者介面 (Agency Cue UI)

Streamlit 製作模式「代理商CUE」：輸入表單 → build_agency_model →
運算邏輯面板、HTML 預覽（components.html）、下載 Excel、產生/下載 PDF。
支援上傳 / 搜尋 Ragic（與一般 CUE 共用同一張表單）。
"""
import os
import base64
import mimetypes
from datetime import datetime, date, timedelta

import streamlit as st
import streamlit.components.v1 as components

import config
from config import RAGIC_MAP, RAGIC_FIELD_SERIAL
import agency_cue as ac
from agency_excel import generate_agency_excel, LOGO_2008, LOGO_DDRIVE
from data_loader import load_agency_pricing_from_cloud
from pdf_converter import xlsx_bytes_to_pdf_bytes
from ragic_api import upload_to_ragic, search_ragic_records, _ragic_number
from utils import safe_filename, html_escape

SEC_OPTIONS = [5, 10, 15, 20, 30]


# =============================================================================
# HTML 預覽
# =============================================================================
def _fmt_money(v):
    if isinstance(v, str):
        return html_escape(v)
    return f"${v:,.0f}"


def _logo_data_uri(agency):
    """回傳代理商 Logo 的 base64 data URI（供 HTML 預覽），無則 None。"""
    path = {"2008傳媒": LOGO_2008, "佳聖": LOGO_DDRIVE}.get(agency)
    if not path or not os.path.exists(path):
        return None
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")
    return f"data:image/png;base64,{b64}"


def _render_pricing_rules(agency):
    """在運算邏輯面板顯示該代理商定價規則；數值直接讀 config，永遠與算法同步。

    完整說明見 repo 根目錄「代理商定價規則.md」。
    """
    def _row(table):
        return " / ".join(str(int(table.get(s, 0))) for s in SEC_OPTIONS) if table else "—"

    fam_tbl = config.AGENCY_LIST_PRICE_FALLBACK.get((agency, ac.PLATFORM_FAMILY))
    wjf_tbl = config.AGENCY_LIST_PRICE_FALLBACK.get((agency, ac.PLATFORM_WJF))
    fn, fs, fsec = config.AGENCY_FAMILY_BASE
    wn, ws, wsec = config.AGENCY_WJF_BASE
    sr_a, sr_b = config.AGENCY_SUPER_RATIO
    ac_default = config.AGENCY_AC_DEFAULT.get(agency)

    st.markdown("---")
    st.markdown("#### 📖 定價規則（本代理商）")

    # 1. 牌價（表上定價）
    st.markdown(
        "**牌價（List，每檔定價；秒 5/10/15/20/30）**\n\n"
        f"| 平台 | 5 | 10 | 15 | 20 | 30 |\n|---|---|---|---|---|---|\n"
        f"| 全家企頻 | {_row(fam_tbl).replace(' / ', ' | ')} |\n"
        f"| 萬家福 | {_row(wjf_tbl).replace(' / ', ' | ')} |"
    )
    st.caption("查價順序：Google Sheet(AgencyPricing) → config fallback；缺該秒數取最近秒數線性換算。")

    # 2. 凱絡三層價
    if agency == "凱絡":
        st.markdown(
            f"**凱絡三層價**：市場價 = 牌價 × {config.CARAT_MARKET_RATIO}、"
            f"統一價 = 牌價 × {config.CARAT_UNI_RATIO}（表上「總價」= 統一價 × 檔次）"
        )

    # 3. 實作價（真正收錢基準）
    st.markdown(
        "**實作價（Net，秒數線性；= 基準 / 標準檔次 / 基準秒 × 秒）**\n\n"
        f"- 全家企頻：{fn:,}/{fs}/{fsec} × 秒（每秒 {fn/fs/fsec:.2f}）\n"
        f"- 萬家福量販：{wn:,}/{ws}/{wsec} × 秒（每秒 {wn/ws/wsec:.2f}）\n"
        f"- 超市（樂家康）：檔次 = 量販 × {sr_a}/{sr_b}，計價於量販、不另收費\n"
        f"- 檔次 = round(該平台預算 ÷ 單檔實作價)"
    )

    # 3b. 2008 萬家福專屬回饋規則
    if agency == "2008傳媒":
        st.warning(
            f"🎁 **2008 萬家福專屬**：量販檔次自動 +{int(ac.WJF_2008_REBATE_PCT)}% 回饋檔"
            f"（= 基礎檔次 + round({ac.WJF_2008_REBATE_PCT/100:.2f}×基礎檔次)），"
            "直接折進每日排檔、**不另立回饋列**；**實收 = 預算不變**。"
            "超市(樂家康)跟著 ×720/420 放大。佳聖 / 凱絡 不適用（回饋另立列）。"
        )

    # 4. 費用區
    if agency == "2008傳媒":
        fee = f"AC = net × {ac_default}%（預設）；稅 = (net+AC) × 5%；合計 = net+AC+稅"
    elif agency == "凱絡":
        fee = "AC = net × 3%（預設收，比照2008）；VAT = (net+AC) × 5%；合計 = net+AC+VAT；特殊案例可勾免收顯示「-」"
    else:  # 佳聖
        fee = "無 AC；VAT = net × 5%；合計 = net+VAT"
    st.markdown(f"**費用區**：{fee}")

    # 5. 例外提醒
    st.info(
        "⚠️ 秒數係數三個例外：\n"
        "1. 凱絡「全家」是純線性（每秒 100），不照全家廣播係數\n"
        "2. 2008「萬家福」套全家廣播係數（非家樂福），故比凱絡/佳聖 便宜\n"
        "3. 基準秒不同：萬家福基準 20 秒、全家/2008 基準 30 秒"
    )


def build_agency_html(model, made_date=None):
    """產出 HTML 預覽字串（供 components.html 呈現，金額含 $ 不會被誤判為 LaTeX）。"""
    css = """
    <style>
    body{font-family:'Microsoft JhengHei','微軟正黑體',sans-serif;margin:8px;color:#222;}
    h3{margin:14px 0 4px;} .meta{font-size:13px;color:#444;margin-bottom:6px;}
    table{border-collapse:collapse;margin-bottom:10px;font-size:12px;}
    th,td{border:1px solid #999;padding:3px 6px;text-align:center;white-space:nowrap;}
    th{background:#f0f0f0;} td.l{text-align:left;} .reb{color:#c0392b;}
    .fee{font-weight:bold;} .totrow{background:#fafafa;font-weight:bold;}
    .wend{background:#fff6d6;}
    </style>
    """
    a = model["agency"]
    out = [css]
    logo = _logo_data_uri(a)
    if logo:
        out.append(f"<div style='text-align:right'><img src='{logo}' style='height:56px'></div>")
    out.append(f"<div class='meta'><b>代理商：</b>{html_escape(a)}　"
               f"<b>客戶：</b>{html_escape(model['client_name'])}　"
               f"<b>產品：</b>{html_escape(model['product_name'])}　"
               f"<b>走期：</b>{model['start_date']:%Y/%m/%d}~{model['end_date']:%Y/%m/%d}</div>")
    start_dt = model["start_date"]
    days = (model["end_date"] - start_dt).days + 1
    day_dates = [start_dt + timedelta(days=i) for i in range(days)]

    for sh in model["sheets"]:
        plat = sh["platform"]
        out.append(f"<h3>{html_escape(plat)}（{sh['seconds']}秒）</h3>")
        # 表頭
        day_hdr = "".join(
            f"<th class='{'wend' if d.weekday()>=5 else ''}'>{d.month}/{d.day}</th>"
            for d in day_dates)
        if a == "凱絡":
            cols = "<th>媒體別</th><th>地區</th><th>時段</th><th>定價</th><th>市場價</th><th>統一價</th><th>檔數</th><th>總價</th><th>專案價</th>"
        else:
            cols = "<th>媒體型態</th><th>地區</th><th>時段</th><th>單位</th><th>定價</th><th>次數</th><th>實收</th>"
        out.append(f"<table><tr>{cols}{day_hdr}</tr>")
        for row in sh["rows"]:
            reb = " class='reb'" if isinstance(row["net_display"], str) and row["net_display"] == ac.NET_REBATE else ""
            label = html_escape(row["media_label"]).replace("\n", "<br>") or "&nbsp;"
            if row["schedule"] is None:
                daycells = f"<td colspan='{days}'>{row['spots']}（合併）</td>"
            else:
                daycells = "".join(
                    f"<td class='{'wend' if day_dates[i].weekday()>=5 else ''}'>{row['schedule'][i] or ''}</td>"
                    for i in range(days))
            if a == "凱絡":
                cells = (f"<td class='l'>{label}</td><td>{html_escape(row['region_label'])}</td>"
                         f"<td>{html_escape(row['daypart'])}</td><td>{_fmt_money(row['list_per'])}</td>"
                         f"<td>{_fmt_money(row['market_per'])}</td><td>{_fmt_money(row['uni_per'])}</td>"
                         f"<td>{row['spots']}</td><td>{_fmt_money(row['uni_total'])}</td>"
                         f"<td{reb}>{_fmt_money(row['net_display'])}</td>")
            else:
                cells = (f"<td class='l'>{label}</td><td>{html_escape(row['region_label'])}</td>"
                         f"<td>{html_escape(row['daypart'])}</td><td>{sh['seconds']}秒</td>"
                         f"<td>{_fmt_money(row['list_total'])}</td><td>{row['spots']}</td>"
                         f"<td{reb}>{_fmt_money(row['net_display'])}</td>")
            out.append(f"<tr>{cells}{daycells}</tr>")
        out.append("</table>")
        # 費用區
        out.append(_fee_html(sh["fees"]))
    if model.get("remarks"):
        out.append("<h3>備註</h3><div class='meta'>" +
                   "<br>".join(html_escape(x) for x in model["remarks"]) + "</div>")
    return "".join(out)


def _fee_html(f):
    def m(v):
        return html_escape(v) if isinstance(v, str) else f"${v:,.0f}"
    if f["type"] == "2008":
        rows = [("Budget (net)", m(f["budget_net"])), (f"AC {int(f['ac_pct'])}%", m(f["ac"])),
                ("5% Tax", m(f["tax"])), ("TOTAL", m(f["total"]))]
    elif f["type"] == "ddrive":
        rows = [("Total Net Cost", m(f["net"])), ("VAT (5%)", m(f["vat"])), ("Total Gross Cost", m(f["gross"]))]
    else:
        rows = [("Sub-Total", m(f["subtotal"])), ("A.C 3%", "-" if f.get("ac_free") else m(f["ac"])),
                ("VAT 5%", m(f["vat"])), ("Grand-Total", m(f["grand"]))]
    body = "".join(f"<tr><td class='l fee'>{k}</td><td class='fee'>{v}</td></tr>" for k, v in rows)
    return f"<table>{body}</table>"


# =============================================================================
# 主 UI
# =============================================================================
def render_agency_cue(sales_map=None):
    st.title("📺 代理商 CUE 表生成器")
    st.caption("三家配合代理商專用 CUE（2008傳媒／佳聖／凱絡）。平台：全家企頻、萬家福。可上傳至 Ragic 並搜尋已上傳案子。")

    # --- 搜尋 / 載入已上傳的代理商案子（與一般 CUE 共用同一張 Ragic 表單）---
    _render_agency_search()

    can_download_excel = st.session_state.get("is_supervisor", False) or st.session_state.get("allow_sales_excel_download", True)
    can_download_pdf = st.session_state.get("is_supervisor", False) or st.session_state.get("allow_sales_pdf_download", True)

    # 牌價來源
    agency_pricing = load_agency_pricing_from_cloud(config.GSHEET_SHARE_URL)
    price_src = "Google Sheet（AgencyPricing）" if agency_pricing else "程式內建 fallback"
    st.info(f"目前牌價來源：**{price_src}**")

    agency = st.radio("代理商", config.AGENCY_LIST, horizontal=True, key="ag_agency")

    c1, c2, c3 = st.columns(3)
    with c1:
        client_name = st.text_input("客戶名稱", st.session_state.get("ag_client", "統一企業"), key="ag_client")
    with c2:
        product_name = st.text_input("產品名稱", st.session_state.get("ag_product", ""), key="ag_product")
    with c3:
        campaign = st.text_input("Campaign（僅 2008 顯示）", st.session_state.get("ag_campaign", ""),
                                 key="ag_campaign") if agency == "2008傳媒" else ""

    c4, c5, c6 = st.columns(3)
    with c4:
        start_date = st.date_input("開始日", st.session_state.get("ag_start", date(2026, 8, 5)), key="ag_start")
    with c5:
        end_date = st.date_input("結束日", st.session_state.get("ag_end", date(2026, 9, 1)), key="ag_end")
    with c6:
        default_mat = ac.minus_business_days(start_date, 5)
        material_due = st.date_input("素材提供時間（預設開始日−5個工作天）", default_mat, key="ag_material")
    if end_date < start_date:
        st.error("結束日不可早於開始日")
        return
    days = (end_date - start_date).days + 1
    st.caption(f"走期共 **{days}** 天")

    total_budget = st.number_input("總預算（未稅 Net）", min_value=0, value=int(st.session_state.get("ag_budget", 250000)),
                                   step=10000, key="ag_budget")

    # 凌晨補償
    comp_mode = st.radio("凌晨補償方式（2026/03 起全家凌晨停播）", ac.COMP_OPTIONS,
                         index=ac.COMP_OPTIONS.index(ac.COMP_PLAN1), key="ag_comp",
                         help="每次製表可選；預設為方案一（15%媒體價值換全家檔次）。")

    st.markdown("---")
    st.markdown("#### 平台設定")
    colf, colw = st.columns(2)

    # --- 全家企頻 ---
    with colf:
        fam_on = st.checkbox("啟用 全家企頻", value=st.session_state.get("ag_fam_on", True), key="ag_fam_on")
        fam_cfg = None
        if fam_on:
            fam_sec = st.selectbox("全家 秒數", SEC_OPTIONS, index=SEC_OPTIONS.index(15), key="ag_fam_sec")
            fam_auto_reb = st.checkbox(
                "自帶專案回饋（凌晨時數轉換）", value=st.session_state.get("ag_fam_auto_reb", True),
                key="ag_fam_auto_reb",
                help="非主時段時數(24−主時段)×30檔×30/秒，自動另立一列「專案回饋」。"
                     "例：主時段07-23→8時、15秒=8×30×2=480檔。三家共用。")
            fam_reb = st.number_input("全家 專案回饋 %（手動另計）", 0, 100, st.session_state.get("ag_fam_reb", 0), key="ag_fam_reb")
            fam_ovr = st.number_input("全家 檔次覆寫（0=自動）", 0, value=st.session_state.get("ag_fam_ovr", 0), key="ag_fam_ovr")
            fam_cfg = {"enabled": True, "seconds": int(fam_sec), "share": 100,
                       "rebate_pct": float(fam_reb), "spots_override": int(fam_ovr),
                       "auto_rebate": bool(fam_auto_reb)}

    # --- 萬家福 ---
    with colw:
        wjf_on = st.checkbox("啟用 萬家福（量販+超市）", value=st.session_state.get("ag_wjf_on", False), key="ag_wjf_on")
        wjf_cfg = None
        if wjf_on:
            wjf_sec = st.selectbox("萬家福 秒數", SEC_OPTIONS, index=SEC_OPTIONS.index(20), key="ag_wjf_sec")
            wjf_reb = st.number_input("萬家福 專案回饋 %", 0, 100, st.session_state.get("ag_wjf_reb", 0), key="ag_wjf_reb")
            wjf_wave = st.checkbox("整波專案回饋（整張不收費、量販手動輸入）", value=st.session_state.get("ag_wjf_wave", False), key="ag_wjf_wave")
            wjf_ovr = 0
            if wjf_wave:
                wjf_ovr = st.number_input("萬家福 量販檔次（手動）", 0, value=st.session_state.get("ag_wjf_ovr", 0), key="ag_wjf_ovr")
            wjf_cfg = {"enabled": True, "seconds": int(wjf_sec), "share": 100,
                       "rebate_pct": float(wjf_reb), "mag_override": int(wjf_ovr),
                       "is_rebate_wave": bool(wjf_wave)}
        elif comp_mode == ac.COMP_PLAN2:
            # 方案二補償需萬家福秒數
            wjf_cfg = {"enabled": False, "seconds": int(st.selectbox(
                "方案二 萬家福轉換秒數", SEC_OPTIONS, index=SEC_OPTIONS.index(15), key="ag_wjf_comp_sec"))}

    if not fam_on and not wjf_on:
        st.warning("請至少啟用一個平台。")
        return

    # 預算佔比：兩平台啟用時 100% 自動連動（仿舊程式，拉一邊另一邊自動補足）
    if fam_on and wjf_on:
        fs = st.session_state.get("ag_fam_share")
        ws_ = st.session_state.get("ag_wjf_share")
        if fs is None or ws_ is None or (int(fs) + int(ws_) != 100):
            st.session_state["ag_fam_share"] = 50
            st.session_state["ag_wjf_share"] = 50

        def _link_share(changed, other):
            st.session_state[other] = max(0, 100 - int(st.session_state[changed]))

        st.markdown("**預算佔比（兩平台自動連動，合計恆為 100%）**")
        cs1, cs2 = st.columns(2)
        with cs1:
            st.slider("全家企頻 %", 0, 100, key="ag_fam_share",
                      on_change=_link_share, args=("ag_fam_share", "ag_wjf_share"))
        with cs2:
            st.slider("萬家福 %", 0, 100, key="ag_wjf_share",
                      on_change=_link_share, args=("ag_wjf_share", "ag_fam_share"))
        fam_cfg["share"] = int(st.session_state["ag_fam_share"])
        wjf_cfg["share"] = int(st.session_state["ag_wjf_share"])
        st.caption(f"目前佔比 → 全家企頻 **{fam_cfg['share']}%**　萬家福 **{wjf_cfg['share']}%**")
    # 單一平台時 share 維持 100（build 會直接用整筆預算）

    # --- 費用/AC 設定 ---
    st.markdown("---")
    ac_pct = None
    if agency == "2008傳媒":
        ac_pct = st.number_input("AC %", 0.0, 100.0, float(config.AGENCY_AC_DEFAULT.get("2008傳媒") or 3.0), step=0.5, key="ag_ac")
    elif agency == "凱絡":
        # 預設收 A.C 3%（比照2008、含進下方 VAT/Grand-Total）；特殊案例可勾免收顯示「-」
        ac_free = st.checkbox("A.C 3% 免收（顯示「-」）", value=False, key="ag_ac_free")
        ac_pct = None if ac_free else 3.0
    else:
        st.caption("佳聖 無 AC 欄位。")

    sign_date = None
    if agency == "凱絡":
        sign_date = st.date_input("凱絡 簽回期限", start_date - timedelta(days=3), key="ag_sign")

    # 備註（可編輯）；各代理商備註不同，切換公司時自動換成該家預設備註
    default_remarks = ac.default_remarks(agency, sign_date)
    default_text = "\n".join(default_remarks)
    if st.session_state.get("ag_remarks_for") != agency:
        st.session_state["ag_remarks"] = default_text
        st.session_state["ag_remarks_for"] = agency
    remarks_text = st.text_area("備註（每行一條）", key="ag_remarks", height=140)
    remarks = [x for x in remarks_text.split("\n") if x.strip()]

    payment_note = st.text_input("請款金額說明（留空自動依走期各月比例拆分）",
                                 st.session_state.get("ag_paynote", ""), key="ag_paynote")

    made_date = date.today()

    # 建模
    try:
        model = ac.build_agency_model(
            agency, client_name, product_name, campaign, start_date, end_date,
            int(total_budget), fam_cfg, wjf_cfg, comp_mode, material_due, ac_pct,
            agency_pricing=agency_pricing, payment_note=payment_note, remarks=remarks,
        )
    except Exception as e:
        st.error(f"計算失敗：{e}")
        return

    # 運算邏輯面板
    with st.expander("🧮 運算邏輯（透明化）"):
        for line in model["logs"]:
            st.text(line)
        _render_pricing_rules(agency)

    # HTML 預覽（必須用 components.html，避免 $ 被當 LaTeX）
    st.markdown("#### 預覽")
    html = build_agency_html(model, made_date)
    components.html(html, height=560, scrolling=True)

    # 檔名
    fn_base = (f"{made_date:%Y%m%d} CUE {agency}-{client_name} {product_name} "
               f"{start_date:%m%d}-{end_date:%m%d}")
    xlsx_name = safe_filename(fn_base) + ".xlsx"
    pdf_name = safe_filename(fn_base) + ".pdf"

    # Excel 下載
    xlsx_bytes = generate_agency_excel(model, made_date)
    st.markdown("---")
    d1, d2 = st.columns(2)
    with d1:
        if can_download_excel:
            st.download_button("⬇️ 下載 Excel", data=xlsx_bytes, file_name=xlsx_name,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)
        else:
            st.button("⬇️ 下載 Excel（需權限）", disabled=True, use_container_width=True)
    with d2:
        if can_download_pdf:
            if st.button("🧾 產生 PDF", use_container_width=True):
                pdf_bytes, method, err = xlsx_bytes_to_pdf_bytes(xlsx_bytes)
                if pdf_bytes:
                    st.session_state["ag_pdf_bytes"] = pdf_bytes
                    st.session_state["ag_pdf_name"] = pdf_name
                    st.success(f"PDF 已生成（{method}）")
                else:
                    st.error(f"PDF 生成失敗：{err}")
            if st.session_state.get("ag_pdf_bytes"):
                st.download_button("⬇️ 下載 PDF", data=st.session_state["ag_pdf_bytes"],
                                   file_name=st.session_state.get("ag_pdf_name", pdf_name),
                                   mime="application/pdf", use_container_width=True)
        else:
            st.button("🧾 產生 PDF（需權限）", disabled=True, use_container_width=True)

    # --- 上傳至 Ragic（與一般 CUE 共用同一張表單）---
    form_inputs = {
        "agency": agency,
        "client_name": client_name,
        "product_name": product_name,
        "campaign": campaign,
        "start_date": start_date.isoformat(),
        "end_date": end_date.isoformat(),
        "material_due": material_due.isoformat(),
        "total_budget": int(total_budget),
        "comp_mode": comp_mode,
        "fam_cfg": fam_cfg,
        "wjf_cfg": wjf_cfg,
        "ac_pct": ac_pct,
        "ac_free": bool(st.session_state.get("ag_ac_free", False)),
        "sign_date": sign_date.isoformat() if sign_date else None,
        "remarks": remarks,
        "payment_note": payment_note,
    }
    _render_agency_upload(model, xlsx_bytes, xlsx_name, form_inputs)


# =============================================================================
# Ragic 上傳
# =============================================================================
def _render_agency_upload(model, xlsx_bytes, xlsx_name, form_inputs):
    """代理商 CUE 上傳至 Ragic（含二次確認、自動附 Excel+PDF、回查流水號）。"""
    st.markdown("---")
    st.markdown("#### ☁️ 上傳至 Ragic")
    st.caption("上傳後會存進與一般 CUE 同一張 Ragic 表單，並自動附上 Excel 與 PDF；可於本頁上方「搜尋舊案」再載入。")

    if not st.session_state.get("ragic_key"):
        st.info("尚未設定 Ragic API Key（於一般 CUE 頁面的『Ragic 連線設定』設定）。")
        return

    # 上傳成功訊息（不會一閃即逝）
    if "ag_upload_success_msg" in st.session_state:
        st.success(st.session_state["ag_upload_success_msg"])
        if st.button("👌 我知道了（清除訊息）", key="ag_upload_msg_clear"):
            del st.session_state["ag_upload_success_msg"]
            st.rerun()

    client_name = model["client_name"]
    product_name = model["product_name"]

    # 自調外觀 Excel（選填）：使用者手動調整外觀後上傳，會一併附到 Ragic。
    # 用 key 保存，跨 rerun（進入確認狀態）不遺失。
    style_excel_file = st.file_uploader(
        "自調外觀excel上傳（選填）",
        type=["xlsx", "xlsm", "xls"],
        key="ag_ragic_style_excel",
        help="僅能調整外觀及呈列方式，請勿調整每日檔次及金額等重要資訊（訂檔資訊以程式設定為主）。",
    )
    st.caption("僅能調整外觀及呈列方式，請勿調整每日檔次及金額等重要資訊（訂檔資訊以程式設定為主）。")

    if not st.session_state.get("ag_ragic_confirm"):
        if st.button("🚀 上傳資料至 Ragic", type="primary", key="ag_ragic_upload_btn"):
            st.session_state["ag_ragic_confirm"] = True
            st.rerun()
        return

    st.warning(f"即將上傳【{model['agency']}｜{client_name} - {product_name}】至 Ragic，請確認？")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("❌ 取消", key="ag_ragic_cancel"):
            st.session_state["ag_ragic_confirm"] = False
            st.rerun()
    with c2:
        if st.button("✅ 確認上傳", key="ag_ragic_confirm_btn"):
            with st.spinner("正在產生 PDF 並上傳資料與檔案..."):
                pdf_bytes, _method, pdf_err = xlsx_bytes_to_pdf_bytes(xlsx_bytes)

                details = ac.build_agency_ragic_details(model, form_inputs)
                sign_deadline = form_inputs.get("sign_date") or ""
                # 只寫語意相符的欄位；代理商名/campaign 等無對應欄位者一律留在備註(details)，
                # 不硬塞進 sales（業務）或 format（報表格式）。
                data_payload = {
                    RAGIC_MAP["client"]:     client_name,
                    RAGIC_MAP["product"]:    product_name,
                    RAGIC_MAP["budget_raw"]: ac.agency_net_total(model),
                    RAGIC_MAP["budget_fin"]: ac.agency_grand_total(model),
                    RAGIC_MAP["date_start"]: model["start_date"].isoformat(),
                    RAGIC_MAP["date_end"]:   model["end_date"].isoformat(),
                    RAGIC_MAP["date_sign"]:  sign_deadline,
                    RAGIC_MAP["details"]:    details,
                }
                if RAGIC_MAP.get("platform"):
                    data_payload[RAGIC_MAP["platform"]] = ac.agency_platform_text(model)
                if RAGIC_MAP.get("total_spots"):
                    data_payload[RAGIC_MAP["total_spots"]] = ac.agency_total_spots(model)
                if RAGIC_MAP.get("seconds_union"):
                    data_payload[RAGIC_MAP["seconds_union"]] = ac.agency_seconds_union(model)

                files_payload = {
                    RAGIC_MAP["file_xls"]: (
                        xlsx_name, xlsx_bytes,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                }
                if pdf_bytes:
                    files_payload[RAGIC_MAP["file_pdf"]] = (
                        xlsx_name.rsplit(".", 1)[0] + ".pdf", pdf_bytes, "application/pdf",
                    )
                # 自調外觀 Excel（選填）→ file_style_xls 欄位
                if style_excel_file and RAGIC_MAP.get("file_style_xls"):
                    _name = safe_filename(style_excel_file.name or f"Style_{safe_filename(client_name)}.xlsx")
                    _mime = mimetypes.guess_type(_name)[0] or "application/octet-stream"
                    files_payload[RAGIC_MAP["file_style_xls"]] = (
                        _name, style_excel_file.getvalue(), _mime,
                    )

                success, msg, _rid = upload_to_ragic(
                    st.session_state["ragic_url"], st.session_state["ragic_key"],
                    data_payload, files_payload,
                )
            st.session_state["ag_ragic_confirm"] = False
            if success:
                if not pdf_bytes:
                    msg += f"（⚠️ PDF 未附上：{pdf_err}）"
                st.session_state["ag_upload_success_msg"] = msg
            else:
                st.error(f"上傳失敗: {msg}")
            st.rerun()


# =============================================================================
# Ragic 搜尋 / 載入舊案
# =============================================================================
def _render_agency_search():
    """搜尋已上傳的代理商 CUE，選一筆載入回表單（100% 還原輸入）。"""
    with st.expander("🔍 搜尋 / 載入已上傳的代理商案子"):
        if not st.session_state.get("ragic_key"):
            st.info("尚未設定 Ragic API Key（於一般 CUE 頁面的『Ragic 連線設定』設定）。")
            return

        kw = st.text_input("輸入 Cue號 或 關鍵字", key="ag_search_kw",
                           placeholder="例如：統一企業 或 1001")
        if st.button("搜尋 Ragic", key="ag_search_btn"):
            found = search_ragic_records(st.session_state["ragic_url"], st.session_state["ragic_key"], kw)
            # 只留「代理商 CUE」案子（details 含代理商標記）
            found = [r for r in found if ac.is_agency_record_details(r.get(RAGIC_MAP["details"], ""))]
            st.session_state["ag_found_records"] = found
            if not found:
                st.warning("查無代理商 CUE 資料")

        records = st.session_state.get("ag_found_records") or []
        if not records:
            return

        def _fmt(idx):
            r = records[idx]
            c = r.get(RAGIC_MAP["client"], "")
            p = r.get(RAGIC_MAP["product"], "")
            # 代理商名沒有專屬欄位，改從備註(details)的 [AGENCY_EXT] 解析
            agy = (ac.parse_agency_ext(r.get(RAGIC_MAP["details"], "")) or {}).get("agency", "")
            d = (r.get(RAGIC_MAP["date_start"], "") or "").split(" ")[0] or "無日期"
            cue = r.get(RAGIC_FIELD_SERIAL, "")
            return f"📅 {d} | 🏢 {c} - 📦 {p} [{agy}] | 🔢 {cue}"

        sel = st.selectbox("選擇一筆資料", range(len(records)), format_func=_fmt, key="ag_search_sel")
        rec = records[sel]
        st.caption(f"Cue號：{rec.get(RAGIC_FIELD_SERIAL, '未設定')}｜"
                   f"含稅合計：${_ragic_number(rec.get(RAGIC_MAP['budget_fin'])):,.0f}｜"
                   f"走期：{rec.get(RAGIC_MAP['date_start'], '')} ~ {rec.get(RAGIC_MAP['date_end'], '')}")

        if st.button("📋 載入此案設定", key="ag_search_load"):
            ok, msg = _restore_agency_state(rec)
            if ok:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)


def _restore_agency_state(record):
    """把 Ragic 記錄的 [AGENCY_EXT] JSON 還原到 st.session_state（ag_* widget keys）。"""
    ext = ac.parse_agency_ext(record.get(RAGIC_MAP["details"], ""))
    if not ext:
        return False, "❌ 此筆非代理商 CUE 或缺少可還原的設定資料。"

    def _d(s):
        try:
            return date.fromisoformat(s)
        except (ValueError, TypeError):
            return None

    ss = st.session_state
    ss["ag_agency"] = ext.get("agency", config.AGENCY_LIST[0])
    ss["ag_client"] = ext.get("client_name", "")
    ss["ag_product"] = ext.get("product_name", "")
    ss["ag_campaign"] = ext.get("campaign", "")

    for key, ek in (("ag_start", "start_date"), ("ag_end", "end_date"), ("ag_material", "material_due")):
        dv = _d(ext.get(ek))
        if dv:
            ss[key] = dv
    ss["ag_budget"] = int(ext.get("total_budget", 0) or 0)

    if ext.get("comp_mode") in ac.COMP_OPTIONS:
        ss["ag_comp"] = ext["comp_mode"]

    # 全家企頻
    fam = ext.get("fam_cfg")
    if fam and fam.get("enabled"):
        ss["ag_fam_on"] = True
        ss["ag_fam_sec"] = int(fam.get("seconds", 15))
        ss["ag_fam_reb"] = int(fam.get("rebate_pct", 0) or 0)
        ss["ag_fam_ovr"] = int(fam.get("spots_override", 0) or 0)
        ss["ag_fam_share"] = int(fam.get("share", 100) or 100)
        ss["ag_fam_auto_reb"] = bool(fam.get("auto_rebate", True))
    else:
        ss["ag_fam_on"] = False

    # 萬家福
    wjf = ext.get("wjf_cfg")
    if wjf and wjf.get("enabled"):
        ss["ag_wjf_on"] = True
        ss["ag_wjf_sec"] = int(wjf.get("seconds", 20))
        ss["ag_wjf_reb"] = int(wjf.get("rebate_pct", 0) or 0)
        ss["ag_wjf_wave"] = bool(wjf.get("is_rebate_wave", False))
        ss["ag_wjf_ovr"] = int(wjf.get("mag_override", 0) or 0)
        ss["ag_wjf_share"] = int(wjf.get("share", 100) or 100)
    else:
        ss["ag_wjf_on"] = False
        if wjf and wjf.get("seconds"):
            ss["ag_wjf_comp_sec"] = int(wjf["seconds"])

    # 費用 / AC
    if ext.get("ac_pct") is not None:
        ss["ag_ac"] = float(ext["ac_pct"])
    ss["ag_ac_free"] = bool(ext.get("ac_free", False))
    sign = _d(ext.get("sign_date"))
    if sign:
        ss["ag_sign"] = sign

    # 備註（設 ag_remarks_for 以免被預設備註覆寫）
    remarks = ext.get("remarks") or []
    ss["ag_remarks"] = "\n".join(remarks)
    ss["ag_remarks_for"] = ss["ag_agency"]
    ss["ag_paynote"] = ext.get("payment_note", "")

    return True, "✅ 已載入，請檢查下方設定。"
