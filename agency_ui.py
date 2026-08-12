# -*- coding: utf-8 -*-
"""
代理商 CUE 使用者介面 (Agency Cue UI)

Streamlit 製作模式「代理商CUE」：輸入表單 → build_agency_model →
運算邏輯面板、HTML 預覽（components.html）、下載 Excel、產生/下載 PDF。
此模式暫不支援 Ragic 上傳。
"""
import os
import base64
from datetime import datetime, date, timedelta

import streamlit as st
import streamlit.components.v1 as components

import config
import agency_cue as ac
from agency_excel import generate_agency_excel, LOGO_2008, LOGO_DDRIVE
from data_loader import load_agency_pricing_from_cloud
from pdf_converter import xlsx_bytes_to_pdf_bytes
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
    st.caption("三家配合代理商專用 CUE（2008傳媒／佳聖／凱絡）。平台：全家企頻、萬家福。此模式暫不支援 Ragic 上傳。")

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
                         index=0, key="ag_comp", help="每次製表可選；預設為原50%檔次併入主時段。")

    st.markdown("---")
    st.markdown("#### 平台設定")
    colf, colw = st.columns(2)

    # --- 全家企頻 ---
    with colf:
        fam_on = st.checkbox("啟用 全家企頻", value=st.session_state.get("ag_fam_on", True), key="ag_fam_on")
        fam_cfg = None
        if fam_on:
            fam_sec = st.selectbox("全家 秒數", SEC_OPTIONS, index=SEC_OPTIONS.index(15), key="ag_fam_sec")
            fam_reb = st.number_input("全家 專案回饋 %", 0, 100, st.session_state.get("ag_fam_reb", 0), key="ag_fam_reb")
            fam_ovr = st.number_input("全家 檔次覆寫（0=自動）", 0, value=st.session_state.get("ag_fam_ovr", 0), key="ag_fam_ovr")
            fam_cfg = {"enabled": True, "seconds": int(fam_sec), "share": 100,
                       "rebate_pct": float(fam_reb), "spots_override": int(fam_ovr)}

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

    st.caption("※ 代理商 CUE 模式暫不支援 Ragic 上傳。")
