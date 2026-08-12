"""
Cue Sheet Pro (媒體排程生成系統)
重構版本 - 模組化架構

用途: 協助業務生成媒體排程表 (Cue表)，支援 HTML 預覽、Excel/PDF 下載及 Ragic 資料庫串接。
維護注意: 此系統依賴外部 Google Sheet 作為設定檔，以及 LibreOffice 進行 PDF 轉檔。
"""

import calendar
import mimetypes
import streamlit as st
import traceback
import time
from datetime import timedelta, datetime, date
from decimal import Decimal, ROUND_HALF_UP


def _last_day_of_next_month():
    """下一個月的最後一天（用於年約季約付款兌現日預設）。"""
    today = date.today()
    if today.month == 12:
        return date(today.year + 1, 1, 31)
    _, last = calendar.monthrange(today.year, today.month + 1)
    return date(today.year, today.month + 1, last)


def _last_day_of_month_after(d):
    """給定某日，回傳「該日所在月份的下一個月」的最後一天（用於各波段付款兌現日預設）。"""
    if d is None:
        return _last_day_of_next_month()
    if d.month == 12:
        _, last = calendar.monthrange(d.year + 1, 1)
        return date(d.year + 1, 1, last)
    _, last = calendar.monthrange(d.year, d.month + 1)
    return date(d.year, d.month + 1, last)


def _slice_rows_to_segment(rows, start_date, seg_start, seg_end):
    """將 rows 的 schedule 切出波段區間，回傳新 rows（不改動原 list）。"""
    from copy import deepcopy
    seg_d0 = (seg_start - start_date).days
    seg_d1 = (seg_end - start_date).days
    seg_days = (seg_end - seg_start).days + 1
    out = []
    for r in rows:
        nr = deepcopy(r)
        if "schedule" in nr and isinstance(nr["schedule"], list):
            full = nr["schedule"]
            nr["schedule"] = full[seg_d0 : seg_d1 + 1] if seg_d0 >= 0 and seg_d1 < len(full) else [0] * seg_days
        out.append(nr)
    return out


def _round_half_up(value):
    return int(Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def _get_cue_segment_rem(i, seg_end, def_sign):
    """取得一般 CUE 第 i 波段的備註 list（從 session_state 讀取）。"""
    seg_data = st.session_state.get("cue_seg_remarks", {}).get(i, {})
    _sig = seg_data.get("sign_deadline", def_sign)
    _bill = seg_data.get("billing_month", "2026年2月")
    _pay = seg_data.get("payment_date", _last_day_of_month_after(seg_end))
    if isinstance(_sig, datetime):
        _sig = _sig.date() if hasattr(_sig, "date") else _sig
    if isinstance(_pay, datetime):
        _pay = _pay.date() if hasattr(_pay, "date") else _pay
    if seg_data.get("live_update_remarks", True):
        return get_remarks_text(_sig, _bill, _pay)
    raw = (st.session_state.get(f"cue_seg_remarks_{i}", "") or "").strip().split("\n")
    rem = [line.strip() for line in raw if line.strip()]
    return rem if rem else get_remarks_text(_sig, _bill, _pay)

# =============================================================================
# 導入所有模組
# =============================================================================
from config import (
    GSHEET_SHARE_URL,
    DEFAULT_RAGIC_URL,
    DEFAULT_RAGIC_KEY,
    RAGIC_FIELD_SERIAL,
    RAGIC_MAP,
    REGIONS_ORDER,
    DURATIONS,
    SUPERVISOR_PASSWORD
)
from utils import (
    safe_filename,
    get_remarks_text,
    format_campaign_details,
    build_ragic_details_full,
    build_platform_detail_text,
    build_platform_text,
    expand_schedule_to_calendar
)
from data_loader import load_config_from_cloud
from calculator import calculate_plan_data, render_logic_panel
from rebate import compute_rebate_rows, merge_rebate_into_rows, get_rebate_summary_text, compute_bonus_rebate_rows, compute_bonus_rebate_rows_from_allocation, get_rebate_qualified_platforms, get_rebate_qualification_detail, get_row_groups, compute_custom_bonus_rows
from html_generator import generate_html_preview
from excel_renderer import generate_excel_from_scratch
from pdf_converter import xlsx_bytes_to_pdf_bytes
from annual_quarter_cue import build_wave_rows, distribute_by_wave_days, round_to_even
from agency_ui import render_agency_cue
from ragic_api import (
    search_ragic_records,
    upload_to_ragic,
    restore_state_from_ragic,
    _ragic_number,
)

# =============================================================================
# 頁面設定 (Page Config)
# =============================================================================
st.set_page_config(
    layout="wide",
    page_title="Cue Sheet Pro"
)

# =============================================================================
# Session State 初始化 (State Initialization)
# =============================================================================
DEFAULT_STATES = {
    "is_supervisor": False,      # 主管權限開關（開啟後可修改專案優惠價／覆寫成交價）
    "allow_sales_excel_download": True,  # 主管可控制一般業務是否可下載 Excel
    "allow_sales_pdf_download": True,    # 主管可控制一般業務是否可下載 PDF
    "allow_sales_excel_download_ui": True,  # 側欄 UI 狀態（避免 widget 狀態覆蓋實際設定）
    "allow_sales_pdf_download_ui": True,    # 側欄 UI 狀態（避免 widget 狀態覆蓋實際設定）
    "rad_share": 100,            # 廣播預算佔比
    "fv_share": 0,               # 新鮮視預算佔比
    "cf_share": 0,               # 家樂福預算佔比
    "cb_rad": True,              # 啟用廣播
    "cb_fv": False,              # 啟用新鮮視
    "cb_cf": False,              # 啟用家樂福
    "ragic_url": DEFAULT_RAGIC_URL,
    "ragic_key": DEFAULT_RAGIC_KEY,
    "ragic_confirm_state": False, # 上傳確認視窗狀態
    # --- UI Widget Keys 的預設值 (避免 set_state 與 widget 初始化衝突) ---
    "rad_nat": True,
    "rad_reg": REGIONS_ORDER,
    "rad_sec": [20],
    "fv_nat": False,
    "fv_reg": ["北區"],
    "fv_sec": [10],
    "cf_sec": [20],
    # 自訂區域比例 (僅 全家廣播/新鮮視 使用；全省時 6 區加總 100%，區域時選中的區加總 100%)
    "rad_use_region_share": False,
    "fv_use_region_share": False,
    "apply_rebate": False,
    "bonus_rebate_pct": 0,
    "is_barter_contract": False,
    "cue_mode": "一般CUE",
}

def _init_session_state():
    """在 main() 內呼叫，避免 import 時 st.session_state 未就緒導致 KeyError。"""
    try:
        for key, default_val in DEFAULT_STATES.items():
            st.session_state.setdefault(key, default_val)
    except KeyError:
        pass


def _get_ragic_extra_state(row_groups):
    """
    從 session_state 彙整「投放參數詳情」擴充欄位（交換/回饋/分波段/備註/業務加贈等），
    供寫入 Ragic details 後可 100% 還原。
    """
    ss = st.session_state
    state = {}
    # 基本與分段
    for k in ("is_barter_contract", "cue_mode", "temp_format_type", "use_date_segments", "apply_rebate",
              "apply_nat_rad_rebate", "rebate_nat_destination", "apply_nat_cf_rebate",
              "apply_region_rad_rebate", "rebate_region_destination",
              "bonus_rebate_pct", "cue_live_update_remarks", "rad_use_region_share", "fv_use_region_share"):
        if k in ss:
            state[k] = ss[k]
    if ss.get("use_date_segments") and "date_segments" in ss and ss["date_segments"]:
        state["date_segments"] = [[str(s), str(e)] for s, e in ss["date_segments"]]
    # 回饋／主管加贈：bonus 相關鍵（動態鍵如 bonus_rs_20 用 session_state 掃描）
    for k, v in list(ss.items()):
        if isinstance(k, str) and k.startswith("bonus_") and k not in state:
            if isinstance(v, (bool, int, float, str)) or (isinstance(v, list) and all(isinstance(x, str) for x in v)):
                state[k] = v
    # 備註
    for k in ("cue_sign_deadline", "cue_billing_month", "cue_payment_date", "cue_remarks_text"):
        if k in ss and ss[k] is not None:
            state[k] = ss[k]
    # 各波段備註
    if "cue_seg_remarks" in ss and ss["cue_seg_remarks"]:
        state["cue_seg_remarks"] = {}
        for i, data in ss["cue_seg_remarks"].items():
            state["cue_seg_remarks"][str(i)] = {kk: (vv.isoformat() if hasattr(vv, "isoformat") and not isinstance(vv, type) else vv) for kk, vv in data.items()}
    # 區域比例
    if "_rad_region_list" in ss and ss["_rad_region_list"]:
        state["_rad_region_list"] = list(ss["_rad_region_list"])
        for r in ss["_rad_region_list"]:
            key = f"rad_region_{r}"
            if key in ss:
                state[key] = ss[key]
    if "_fv_region_list" in ss and ss["_fv_region_list"]:
        state["_fv_region_list"] = list(ss["_fv_region_list"])
        for r in ss["_fv_region_list"]:
            key = f"fv_region_{r}"
            if key in ss:
                state[key] = ss[key]
    # 業務加贈：依 row_groups 的 skey 從 session_state 讀取
    if row_groups:
        custom = []
        for g in row_groups:
            media, sec, region_key = g["key"]
            skey = f"cb_{media}_{sec}_{region_key}".replace(" ", "_")
            if ss.get(f"cb_en_{skey}") and (ss.get(f"cb_spots_{skey}") or 0) > 0:
                custom.append({
                    "key": [media, str(sec), region_key],
                    "enabled": True,
                    "spots_per_day": ss.get(f"cb_spots_{skey}", 0),
                    "date_start": ss.get(f"cb_start_{skey}"),
                    "date_end": ss.get(f"cb_end_{skey}"),
                })
        if custom:
            state["custom_bonus"] = custom
    return state


def _render_annual_quarter_cue(store_counts_num, pricing_db, sec_factors, regions_order, fmt_options, fmt_idx, sales_map):
    """年約／季約細 CUE：以已知檔次與實收分配至各波段，每波段獨立 Excel/PDF。每個廣告組合可個別設定每波檔次與實收；若不知道各波分配則填該組合總檔次／總實收後按「依天數均分」。"""
    st.caption("以已知**執行區域、平台、秒數**設定各波段檔次與實收；若不知各波分配，可填該**廣告組合**的總檔次／總實收後按「依波段天數均分」。")
    can_download_excel = st.session_state.get("is_supervisor", False) or st.session_state.get("allow_sales_excel_download", True)
    can_download_pdf = st.session_state.get("is_supervisor", False) or st.session_state.get("allow_sales_pdf_download", True)
    if "aq_combos" not in st.session_state:
        st.session_state.aq_combos = []
    if "aq_waves" not in st.session_state:
        st.session_state.aq_waves = []

    combos = st.session_state.aq_combos
    n_combos = len(combos)

    def _ensure_wave_combo_arrays():
        """確保每波都有 combo_spots / combo_net 且長度等於目前組合數。"""
        for w in st.session_state.aq_waves:
            if "combo_spots" not in w:
                w["combo_spots"] = [w.get("spots", 0)] if "spots" in w else []
                w["combo_net"] = [w.get("net", 0)] if "net" in w else []
            while len(w["combo_spots"]) < n_combos:
                w["combo_spots"].append(0)
                w["combo_net"].append(0)
            if len(w["combo_spots"]) > n_combos:
                w["combo_spots"] = w["combo_spots"][:n_combos]
                w["combo_net"] = w["combo_net"][:n_combos]

    loaded_fmt = st.session_state.get("temp_format_type", "東吳")
    try:
        fmt_idx_cur = fmt_options.index(loaded_fmt)
    except Exception:
        fmt_idx_cur = 0
    format_type = st.radio("選擇格式", fmt_options, index=fmt_idx_cur, key="aq_format", horizontal=True)
    sales_options = list(sales_map.keys()) if sales_map else []
    def_client = st.session_state.get("temp_client_name", "萬國通路")
    def_tax = st.session_state.get("temp_tax_id", "")
    def_prod = st.session_state.get("temp_product_name", "統一布丁")
    client_name = st.text_input("客戶名稱", def_client, key="aq_client")
    client_tax_id = st.text_input("統一編號", def_tax, key="aq_tax")
    product_name = st.text_input("產品名稱", def_prod, key="aq_product")
    sales_person = st.selectbox("業務名稱", options=sales_options, index=0, key="aq_sales") if sales_options else ""

    st.markdown("---")
    st.markdown("#### 廣告組合（平台、區域、秒數）")
    col_list, col_add = st.columns([1, 1])
    with col_add:
        media_aq = st.selectbox("平台", ["全家廣播", "新鮮視", "家樂福"], key="aq_add_media")
        if media_aq == "家樂福":
            region_aq = "全省"
        else:
            region_aq = st.selectbox("區域", ["全省"] + list(regions_order), key="aq_add_region")
        sec_aq = st.selectbox("秒數", DURATIONS, key="aq_add_sec")
        if st.button("➕ 加入組合"):
            st.session_state.aq_combos.append({"media": media_aq, "region": region_aq, "seconds": int(sec_aq)})
            for w in st.session_state.aq_waves:
                if "combo_spots" not in w:
                    w["combo_spots"] = [w.get("spots", 0)] if "spots" in w else []
                    w["combo_net"] = [w.get("net", 0)] if "net" in w else []
                w["combo_spots"].append(0)
                w["combo_net"].append(0)
            st.rerun()
    with col_list:
        if not st.session_state.aq_combos:
            st.info("請在右側選擇平台、區域、秒數後按「加入組合」。")
        else:
            for i, c in enumerate(st.session_state.aq_combos):
                st.text(f"【{c['media']}】{c['region']} {c['seconds']}秒")
                if st.button("刪除", key=f"aq_del_{i}"):
                    st.session_state.aq_combos.pop(i)
                    for w in st.session_state.aq_waves:
                        if "combo_spots" in w and len(w["combo_spots"]) > i:
                            w["combo_spots"].pop(i)
                        if "combo_net" in w and len(w["combo_net"]) > i:
                            w["combo_net"].pop(i)
                    st.rerun()

    _ensure_wave_combo_arrays()

    st.markdown("#### 波段設定（只設定每波起訖日）")
    if st.button("➕ 新增一波段", key="aq_add_wave"):
        st.session_state.aq_waves.append({
            "start": date(2026, 1, 1), "end": date(2026, 1, 31),
            "combo_spots": [0] * n_combos, "combo_net": [0] * n_combos
        })
        st.rerun()
    waves = st.session_state.aq_waves
    if waves:
        st.caption("可為每個波段設定「波段最終價格覆寫」：輸入後該波 CUE 表之總實收將以此金額為準（不需主管登入）。")
        for i in range(len(waves)):
            w = st.session_state.aq_waves[i]
            c1, c2, c3, c4, c5 = st.columns([1, 2, 2, 2, 1])
            with c1:
                st.caption(f"波段 {i+1}")
            with c2:
                s = st.date_input("開始日", w.get("start", date(2026, 1, 1)), key=f"aq_w_start_{i}", label_visibility="collapsed")
            with c3:
                end_default = w.get("end", date(2026, 1, 31))
                if end_default < s:
                    end_default = s
                e = st.date_input("結束日", end_default, min_value=s, key=f"aq_w_end_{i}", label_visibility="collapsed")
            with c4:
                ov = w.get("price_override")
                ov_val = int(ov) if ov is not None and ov != "" else None
                price_override = st.number_input(
                    "波段最終價格覆寫 (0=不覆寫)",
                    min_value=0,
                    value=ov_val if ov_val is not None else 0,
                    step=1000,
                    key=f"aq_w_price_override_{i}",
                    help="輸入後此波 CUE 表總實收將以此金額為準，各列依比例縮放"
                )
                st.session_state.aq_waves[i]["price_override"] = price_override if price_override and price_override > 0 else None
            with c5:
                if st.button("刪除此波", key=f"aq_w_del_{i}"):
                    st.session_state.aq_waves.pop(i)
                    st.rerun()
            st.session_state.aq_waves[i]["start"] = s
            st.session_state.aq_waves[i]["end"] = e
    else:
        st.info("請先「新增一波段」，再於下方表格填寫各波檔次與實收。")

    st.markdown("#### 檔次與實收（廣告組合 × 波段：一表輸入）")
    if n_combos > 0 and waves:
        # 總計列：用於依天數均分
        st.caption("若不知各波分配，請填下列總檔次／總實收後按「依波段天數均分」。")
        tot_cols = []
        for j in range(n_combos):
            tot_cols.extend([f"aq_ct_spots_{j}", f"aq_ct_net_{j}"])
        n_tot_cols = len(tot_cols)
        row_tot = st.columns([2] + [1] * n_tot_cols + [1])
        with row_tot[0]:
            st.markdown("**總計（均分用）**")
        for j in range(n_combos):
            c = combos[j]
            lbl = f"【{c['media']}】{c['region']} {c['seconds']}秒"
            with row_tot[1 + j*2]:
                st.number_input(lbl + " 總檔次", min_value=0, value=st.session_state.get(f"aq_ct_spots_{j}", 0), key=f"aq_ct_spots_{j}", label_visibility="collapsed")
            with row_tot[2 + j*2]:
                st.number_input(lbl + " 總實收", min_value=0, value=st.session_state.get(f"aq_ct_net_{j}", 0), key=f"aq_ct_net_{j}", label_visibility="collapsed")
        with row_tot[-1]:
            if st.button("🔄 依波段天數均分", key="aq_btn_distribute"):
                if not st.session_state.aq_waves:
                    st.warning("請先「新增一波段」再使用均分。")
                else:
                    _ensure_wave_combo_arrays()
                    wave_tuples = [(w["start"], w["end"]) for w in st.session_state.aq_waves]
                    for j in range(n_combos):
                        tot_sp = int(st.session_state.get(f"aq_ct_spots_{j}", 0) or 0)
                        tot_nt = int(st.session_state.get(f"aq_ct_net_{j}", 0) or 0)
                        distributed = distribute_by_wave_days(tot_sp, tot_nt, wave_tuples)
                        for i, (sp, nt) in enumerate(distributed):
                            if i < len(st.session_state.aq_waves) and j < len(st.session_state.aq_waves[i]["combo_spots"]):
                                st.session_state.aq_waves[i]["combo_spots"][j] = sp
                                st.session_state.aq_waves[i]["combo_net"][j] = nt
                                st.session_state[f"aq_tbl_spots_{i}_{j}"] = sp
                                st.session_state[f"aq_tbl_net_{i}_{j}"] = nt
                    st.success("已依波段天數均分至各波。")
                    st.rerun()

        # 表頭
        h_cols = st.columns([2, 1, 1] + [1, 1] * n_combos + [1])
        with h_cols[0]:
            st.markdown("**波段**")
        with h_cols[1]:
            st.markdown("**開始日**")
        with h_cols[2]:
            st.markdown("**結束日**")
        for j in range(n_combos):
            c = combos[j]
            short = f"{c['media']}{c['region']}{c['seconds']}s"
            with h_cols[3 + j*2]:
                st.markdown(f"**{short} 檔次**")
            with h_cols[4 + j*2]:
                st.markdown(f"**{short} 實收**")
        with h_cols[-1]:
            st.markdown("**操作**")

        # 資料行：每波一行，每格一個 number_input 放在對應 column
        for i in range(len(st.session_state.aq_waves)):
            w = st.session_state.aq_waves[i]
            cs = w.get("combo_spots", [0] * n_combos)
            cn = w.get("combo_net", [0] * n_combos)
            cols = st.columns([2, 1, 1] + [1, 1] * n_combos + [1])
            with cols[0]:
                st.caption(f"波段 {i+1}")
            with cols[1]:
                st.caption(str(w.get("start", "")))
            with cols[2]:
                st.caption(str(w.get("end", "")))
            for j in range(n_combos):
                key_sp = f"aq_tbl_spots_{i}_{j}"
                key_nt = f"aq_tbl_net_{i}_{j}"
                default_sp = int(cs[j]) if j < len(cs) else 0
                default_nt = int(cn[j]) if j < len(cn) else 0
                with cols[3 + j*2]:
                    if key_sp not in st.session_state:
                        sp = st.number_input("檔次", min_value=0, value=default_sp, key=key_sp, label_visibility="collapsed")
                    else:
                        sp = st.number_input("檔次", min_value=0, key=key_sp, label_visibility="collapsed")
                with cols[4 + j*2]:
                    if key_nt not in st.session_state:
                        nt = st.number_input("實收", min_value=0, value=default_nt, key=key_nt, label_visibility="collapsed")
                    else:
                        nt = st.number_input("實收", min_value=0, key=key_nt, label_visibility="collapsed")
                st.session_state.aq_waves[i]["combo_spots"][j] = sp
                st.session_state.aq_waves[i]["combo_net"][j] = nt
            with cols[-1]:
                if st.button("刪除", key=f"aq_tbl_del_{i}"):
                    st.session_state.aq_waves.pop(i)
                    st.rerun()
    elif n_combos > 0 and not waves:
        st.caption("請先在上方「波段設定」新增至少一波段。")

    combos = st.session_state.aq_combos
    waves = st.session_state.aq_waves
    if not combos:
        st.warning("請至少加入一組廣告組合。")
        return
    if not waves:
        return
    _ensure_wave_combo_arrays()

    # 各波段備註以各波自己的三個欄位為準，不再使用上方範本
    _def_sign = datetime.now() + timedelta(days=3)
    if isinstance(_def_sign, datetime):
        _def_sign = _def_sign.date()
    _def_pay = _last_day_of_next_month()

    st.markdown("---")
    st.subheader("📝 各波段備註（每個波段可分別設定）")
    st.caption("以下每個波段都有獨立的「回簽／請款月份／付款兌現日」與備註全文；付款兌現日預設為**該波段結束日的下個月最後一天**。勾選「依上列日期即時更新備註」後，修改上方欄位時下方條列會即時變動；不勾選時可手動編輯或按鈕產生。")
    for i, w in enumerate(waves):
        start_d = w["start"]
        end_d = w["end"]
        # 各波預設：付款兌現日 = 波段結束日的下個月最後一天
        _def_pay_wave = _last_day_of_month_after(end_d)
        if "payment_date" not in w or w.get("payment_date") is None:
            st.session_state.aq_waves[i]["payment_date"] = _def_pay_wave
        if "sign_deadline" not in w or w.get("sign_deadline") is None:
            st.session_state.aq_waves[i]["sign_deadline"] = _def_sign
        if "billing_month" not in w or w.get("billing_month") is None:
            st.session_state.aq_waves[i]["billing_month"] = "2026年2月"
        w = st.session_state.aq_waves[i]
        _sig = w.get("sign_deadline", _def_sign)
        _bill = w.get("billing_month", "2026年2月")
        _pay = w.get("payment_date", _def_pay_wave)
        if isinstance(_sig, datetime):
            _sig = _sig.date() if hasattr(_sig, "date") else _def_sign
        if isinstance(_pay, datetime):
            _pay = _pay.date() if hasattr(_pay, "date") else _def_pay_wave
        _ta_val = w.get("remarks_text") or "\n".join(get_remarks_text(_sig, _bill, _pay))
        with st.expander(f"**波段 {i+1}**：{start_d} ~ {end_d} — 本波備註", expanded=(i == 0)):
            st.caption("回簽及進單期限、請款月份、付款兌現日期（預設為本波結束日的下個月最後一天）；勾選「即時更新」後改欄位會即時更新下方條列。")
            r1, r2, r3 = st.columns(3)
            with r1:
                _sig = st.date_input("回簽及進單期限", _sig, key=f"aq_wave_sign_{i}")
            with r2:
                _bill = st.text_input("請款月份", _bill, key=f"aq_wave_bill_{i}", placeholder="例：2026年2月")
            with r3:
                _pay = st.date_input("付款兌現日期（預設本波結束下月最後一天）", _pay, key=f"aq_wave_pay_{i}")
            st.session_state.aq_waves[i]["sign_deadline"] = _sig
            st.session_state.aq_waves[i]["billing_month"] = _bill
            st.session_state.aq_waves[i]["payment_date"] = _pay
            live_update = st.checkbox(
                "依上列日期即時更新備註（修改欄位時下方條列會動態改變）",
                value=st.session_state.aq_waves[i].get("live_update_remarks", True),
                key=f"aq_wave_live_rem_{i}",
            )
            st.session_state.aq_waves[i]["live_update_remarks"] = live_update
            if live_update:
                _gen_rem = "\n".join(get_remarks_text(_sig, _bill, _pay))
                st.session_state[f"aq_wave_remarks_{i}"] = _gen_rem
                st.session_state.aq_waves[i]["remarks_text"] = _gen_rem
            if not live_update and st.button("依上列日期重新產生本波備註", key=f"aq_wave_rem_btn_{i}"):
                _new_rem = "\n".join(get_remarks_text(_sig, _bill, _pay))
                st.session_state.aq_waves[i]["remarks_text"] = _new_rem
                st.session_state[f"aq_wave_remarks_{i}"] = _new_rem
                st.rerun()
            if live_update:
                _display_val = _gen_rem
            else:
                _display_val = st.session_state.get(f"aq_wave_remarks_{i}", _ta_val) or _ta_val
            st.text_area(
                "本波備註內容",
                value=_display_val,
                height=200,
                key=f"aq_wave_remarks_{i}",
                label_visibility="collapsed",
                placeholder="留空則依上列三個日期欄位自動產生。可輸入多行，每行會成為一條備註。",
                disabled=live_update,
            )
            if not live_update:
                st.session_state.aq_waves[i]["remarks_text"] = st.session_state.get(f"aq_wave_remarks_{i}", _ta_val) or ""

    st.markdown("---")
    st.subheader("📥 各波段下載（每波獨立 Excel / PDF）")
    for i, w in enumerate(waves):
        start_d = w["start"]
        end_d = w["end"]
        cs = w.get("combo_spots", [])
        cn = w.get("combo_net", [])
        if not cs and not cn:
            cs = [0] * len(combos)
            cn = [0] * len(combos)
        wave_spots_list = [cs[j] if j < len(cs) else 0 for j in range(len(combos))]
        wave_net_list = [cn[j] if j < len(cn) else 0 for j in range(len(combos))]
        if sum(wave_spots_list) <= 0 and sum(wave_net_list) <= 0:
            st.caption(f"波段 {i+1}：{start_d} ~ {end_d} — 請輸入各組合檔次或實收")
            continue
        rows = build_wave_rows(combos, start_d, end_d, wave_spots_list, wave_net_list, pricing_db, sec_factors, store_counts_num, regions_order)
        if not rows:
            st.caption(f"波段 {i+1}：無法產出（請檢查定價表是否有該組合）")
            continue
        total_days_wave = (end_d - start_d).days + 1
        total_list_wave = sum(r.get("rate_display", 0) for r in rows if isinstance(r.get("rate_display"), (int, float)))
        budget_wave = sum(r.get("pkg_display", 0) for r in rows if isinstance(r.get("pkg_display"), (int, float)))
        # 波段最終價格覆寫：若有設定則以該金額為總實收，各列 pkg_display 依比例縮放
        price_override = w.get("price_override") if isinstance(w.get("price_override"), (int, float)) and w.get("price_override") > 0 else None
        if price_override is not None and budget_wave and budget_wave > 0:
            scale = price_override / budget_wave
            pkg_list = [r.get("pkg_display", 0) for r in rows if isinstance(r.get("pkg_display"), (int, float))]
            scaled = [int(round(p * scale)) for p in pkg_list]
            delta = price_override - sum(scaled)
            if scaled:
                scaled[0] = scaled[0] + delta
            for j, r in enumerate(rows):
                if isinstance(r.get("pkg_display"), (int, float)) and j < len(scaled):
                    r["pkg_display"] = scaled[j]
            budget_wave = price_override
        # 年約/季約細 CUE：標題 Product 僅顯示「秒數種類」一次，例如「5秒」而不是重複多次
        unique_secs = sorted(list({r['seconds'] for r in rows}))
        p_str = f"{'、'.join([f'{s}秒' for s in unique_secs])} {product_name}"
        _wave_default_rem = "\n".join(get_remarks_text(
            w.get("sign_deadline", _def_sign),
            w.get("billing_month", "2026年2月"),
            w.get("payment_date", _last_day_of_month_after(w["end"])),
        ))
        _wave_rem_raw = st.session_state.get(f"aq_wave_remarks_{i}") or w.get("remarks_text") or _wave_default_rem
        _wave_rem = [line.strip() for line in (_wave_rem_raw or "").strip().split("\n") if line.strip()]
        if not _wave_rem:
            _wave_rem = get_remarks_text(
                w.get("sign_deadline", _def_sign),
                w.get("billing_month", "2026年2月"),
                w.get("payment_date", _last_day_of_month_after(w["end"])),
            )
        with st.expander(f"波段 {i+1} 預覽：{start_d} ~ {end_d}", expanded=False):
            html_preview = generate_html_preview(rows, total_days_wave, start_d, end_d, client_name, client_tax_id, p_str, format_type, _wave_rem, total_list_wave, budget_wave + _round_half_up(budget_wave * 0.05), budget_wave, 0)
            if isinstance(html_preview, list):
                for idx, one_html in enumerate(html_preview):
                    st.caption(f"第 {idx+1} 頁")
                    st.components.v1.html(one_html, height=400, scrolling=True)
            else:
                st.components.v1.html(html_preview, height=400, scrolling=True)
        xlsx_bytes = generate_excel_from_scratch(format_type, start_d, end_d, client_name, client_tax_id, product_name, rows, _wave_rem, budget_wave, 0, sales_person, total_list_wave)
        pdf_bytes, _, _ = xlsx_bytes_to_pdf_bytes(xlsx_bytes)
        col_x, col_p = st.columns(2)
        with col_x:
            if can_download_excel:
                st.download_button(f"📥 波段{i+1} Excel", xlsx_bytes, f"Cue_{safe_filename(client_name)}_波段{i+1}_{start_d}_{end_d}.xlsx", key=f"aq_xlsx_{i}", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.caption("Excel 下載已由主管關閉")
        with col_p:
            if pdf_bytes and can_download_pdf:
                st.download_button(f"📥 波段{i+1} PDF", pdf_bytes, f"Cue_{safe_filename(client_name)}_波段{i+1}_{start_d}_{end_d}.pdf", key=f"aq_pdf_{i}", mime="application/pdf")
            elif not can_download_pdf:
                st.caption("PDF 下載已由主管關閉")
            else:
                st.caption("PDF 需 LibreOffice")


# =============================================================================
# 主程式邏輯 (Main Execution Block)
# =============================================================================
def main():
    _init_session_state()
    try:
        with st.spinner("正在讀取 Google 試算表設定檔..."):
            STORE_COUNTS, STORE_COUNTS_NUM, PRICING_DB, SEC_FACTORS, SALES_MAP, err_msg = load_config_from_cloud(GSHEET_SHARE_URL)
        
        if err_msg:
            st.error(f"❌ 設定檔載入失敗: {err_msg}")
            st.stop()
        
        # --- Sidebar 邏輯 (登入與設定) ---
        with st.sidebar:
            with st.expander("ℹ️ 版本", expanded=False):
                st.caption(f"Streamlit {st.__version__}")
            st.header("🕵️ 主管登入")
            if not st.session_state.is_supervisor:
                pwd = st.text_input("輸入密碼", type="password", key="pwd_input")
                if st.button("登入"):
                    if pwd == SUPERVISOR_PASSWORD:
                        st.session_state.is_supervisor = True
                        st.session_state.pop("_supervisor_final_budget", None)
                        st.session_state["supervisor_override_price"] = ""
                        st.rerun()
                    else:
                        st.error("密碼錯誤")
            else:
                st.success("✅ 目前狀態：主管模式")
                if st.button("登出"):
                    st.session_state.is_supervisor = False
                    st.rerun()
            if st.session_state.is_supervisor:
                st.markdown("#### 🔐 下載權限設定")
                excel_allowed = st.checkbox(
                    "允許一般業務下載 Excel",
                    key="allow_sales_excel_download_ui",
                )
                pdf_allowed = st.checkbox(
                    "允許一般業務下載 PDF",
                    key="allow_sales_pdf_download_ui",
                )
                # 真正生效的設定值
                st.session_state["allow_sales_excel_download"] = bool(excel_allowed)
                st.session_state["allow_sales_pdf_download"] = bool(pdf_allowed)
            else:
                # 非主管也渲染同一組 checkbox（唯讀），避免 widget state 在登出後被清掉而回到預設 True
                st.markdown("#### 🔐 下載權限設定")
                st.session_state["allow_sales_excel_download_ui"] = bool(st.session_state.get("allow_sales_excel_download", True))
                st.session_state["allow_sales_pdf_download_ui"] = bool(st.session_state.get("allow_sales_pdf_download", True))
                st.checkbox(
                    "允許一般業務下載 Excel",
                    key="allow_sales_excel_download_ui",
                    disabled=True,
                )
                st.checkbox(
                    "允許一般業務下載 PDF",
                    key="allow_sales_pdf_download_ui",
                    disabled=True,
                )
            # 覆寫欄位：預設空白，輸入數字後才啟用覆寫
            st.caption("🔒 專案優惠價覆寫")
            _ov_raw = st.text_input("最終成交價", value=st.session_state.get("supervisor_override_price", ""), key="supervisor_override_price", label_visibility="collapsed", placeholder="留空則以總預算計", disabled=not st.session_state.is_supervisor)
            if st.session_state.is_supervisor:
                _ov_stripped = (_ov_raw or "").strip().replace(",", "")
                if _ov_stripped:
                    try:
                        _ov_num = float(_ov_stripped)
                        if _ov_num >= 0:
                            st.session_state._supervisor_final_budget = _ov_num
                            st.caption(f"⚠️ 以 ${_ov_num:,.0f} 結算")
                        else:
                            st.session_state.pop("_supervisor_final_budget", None)
                            st.caption("請輸入 ≥ 0 的數字")
                    except ValueError:
                        st.session_state.pop("_supervisor_final_budget", None)
                        st.caption("請輸入有效數字")
                else:
                    st.session_state.pop("_supervisor_final_budget", None)
            
            st.markdown("---")
            # --- 新增功能: Ragic 搜尋與載入 (強化版 UI) ---
            st.subheader("🔍 搜尋舊排程")
            search_kw = st.text_input("輸入 Cue號 或 關鍵字", placeholder="例如: 1001 或 萬國通路")
            if st.button("搜尋 Ragic"):
                if not st.session_state.ragic_key:
                    st.error("請先設定 Ragic API Key")
                else:
                    found = search_ragic_records(st.session_state.ragic_url, st.session_state.ragic_key, search_kw)
                    st.session_state['found_records'] = found
                    if not found: st.warning("查無資料")
            
            if 'found_records' in st.session_state and st.session_state['found_records']:
                st.markdown("---")
                records = st.session_state['found_records']
                
                # 建立搜尋結果選單的顯示格式
                def format_search_result(idx):
                    rec = records[idx]
                    
                    # 取得各欄位資料，若無則顯示空白 (使用 RAGIC_MAP 統一管理)
                    c_name = rec.get(RAGIC_MAP['client'], '')
                    p_name = rec.get(RAGIC_MAP['product'], '')
                    s_date = rec.get(RAGIC_MAP['date_start'], '')
                    sales = rec.get(RAGIC_MAP['sales'], '')
                    
                    # 抓取真實 Cue 號 (如果 RAGIC_FIELD_SERIAL 沒填對，就只會顯示空)
                    real_cue = rec.get(RAGIC_FIELD_SERIAL, '') 
                    
                    # 日期簡化
                    date_str = s_date.split(' ')[0] if s_date else "無日期"
                    
                    return f"📅 {date_str} | 🏢 {c_name} - 📦 {p_name} ({sales}) | 🔢 {real_cue}"

                selected_idx = st.selectbox(
                    "選擇一筆資料", 
                    range(len(records)), 
                    format_func=format_search_result
                )
                
                # 顯示選中項目的詳細預覽卡片
                sel_rec = records[selected_idx]
                with st.container():
                    st.markdown(f"**詳細預覽 #{sel_rec.get('_ragicId')}**")
                    st.caption(f"Cue號: {sel_rec.get(RAGIC_FIELD_SERIAL, '未設定')}")
                    col_p1, col_p2 = st.columns(2)
                    col_p1.metric("預算", f"${_ragic_number(sel_rec.get(RAGIC_MAP['budget_raw'])):,.0f}")
                    col_p2.metric("製作費", f"${_ragic_number(sel_rec.get(RAGIC_MAP['prod_cost'])):,.0f}")
                    st.text(f"走期: {sel_rec.get(RAGIC_MAP['date_start'],'')} ~ {sel_rec.get(RAGIC_MAP['date_end'],'')}")
                
                if st.button("📋 載入此專案設定"):
                    success, msg = restore_state_from_ragic(sel_rec)
                    if success:
                        st.success(msg)
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error(msg)
                
                # 除錯區塊: 顯示所有欄位 ID
                st.markdown("---")
                with st.expander("🛠️ 除錯模式：查看欄位 ID"):
                    st.write("請對照下表找到 Cue號 對應的 Key，並修改程式碼中的 RAGIC_FIELD_SERIAL")
                    st.json(sel_rec)
            
            st.markdown("---")
            st.subheader("☁️ Ragic 連線設定")
            
            if st.session_state.is_supervisor:
                st.session_state.ragic_url = st.text_input("Ragic 表單網址", value=st.session_state.ragic_url)
                st.session_state.ragic_key = st.text_input("Ragic API Key", value=st.session_state.ragic_key, type="password")
            else:
                st.text_input("Ragic 表單網址", value=st.session_state.ragic_url, disabled=True)
            
            st.markdown("---")
            if st.button("🧹 清除快取"):
                st.cache_data.clear()
                st.rerun()

        # --- Main Content 邏輯 (輸入與報表) ---
        st.title("📺 媒體 Cue 表生成器")
        
        # 修改：處理格式的預設選擇 (東吳/聲活/鉑霖)
        loaded_fmt = st.session_state.get('temp_format_type', '東吳')
        fmt_options = ["東吳", "聲活", "鉑霖"]
        try:
            fmt_idx = fmt_options.index(loaded_fmt)
        except:
            fmt_idx = 0
        
        # 製作模式（先分流：代理商CUE 與「選擇格式／交換合約」無關，須在其之前分流）
        mode_options = ["一般CUE", "年約季約細CUE", "代理商CUE"]
        cur_mode = st.session_state.get("cue_mode", "一般CUE")
        try:
            mode_idx = mode_options.index(cur_mode)
        except ValueError:
            mode_idx = 0
        cue_mode = st.radio("製作模式", mode_options, index=mode_idx, key="cue_mode", horizontal=True,
                            help="年約季約細CUE：以已知檔次與實收分配至各波段，每波段獨立存檔。代理商CUE：2008傳媒／佳聖／凱絡專用格式。")

        if cue_mode == "代理商CUE":
            render_agency_cue(SALES_MAP)
            return

        format_type = st.radio("選擇格式", fmt_options, index=fmt_idx, horizontal=True)
        is_barter_contract = st.checkbox("是否為交換合約", value=st.session_state.get("is_barter_contract", False), key="is_barter_contract", help="交換合約：檔次依定價計算，不提供優惠回饋檔次。")

        if cue_mode == "年約季約細CUE":
            _render_annual_quarter_cue(STORE_COUNTS_NUM, PRICING_DB, SEC_FACTORS, REGIONS_ORDER, fmt_options, fmt_idx, SALES_MAP)
            return

        c1, c2, c3, c4, c5_sales = st.columns(5)
        # 載入資料後的預設值
        def_client = st.session_state.get('temp_client_name', "萬國通路")
        def_tax = st.session_state.get('temp_tax_id', "")
        def_prod = st.session_state.get('temp_product_name', "統一布丁")
        def_budget = float(st.session_state.get("_total_budget_for_sidebar", st.session_state.get('temp_budget', 1000000)))
        def_cost = float(st.session_state.get('temp_prod_cost', 0))
        def_sales_idx = 0
        
        # 嘗試還原業務選項
        loaded_sales = st.session_state.get('temp_sales')
        sales_options = list(SALES_MAP.keys()) if SALES_MAP else []
        if loaded_sales and loaded_sales in sales_options:
             def_sales_idx = sales_options.index(loaded_sales)
        elif loaded_sales:
             # 如果載入的是綽號，反查
             inv_map = {v: k for k, v in SALES_MAP.items()}
             if loaded_sales in inv_map:
                 def_sales_idx = sales_options.index(inv_map[loaded_sales])
        
        with c1: 
            client_name = st.text_input("客戶名稱", def_client)
            client_tax_id = st.text_input("統一編號", def_tax)
        with c2: product_name = st.text_input("產品名稱", def_prod)
        with c3:
            total_budget_input = st.number_input("總預算 (未稅 Net)", value=def_budget, step=10000.0)
        st.session_state["_total_budget_for_sidebar"] = float(total_budget_input)
        with c4: prod_cost_input = st.number_input("製作費 (未稅)", value=def_cost, step=1000.0)
        
        with c5_sales: 
            sales_person = st.selectbox("業務名稱", options=sales_options, index=def_sales_idx)

        final_budget_val = total_budget_input
        if st.session_state.is_supervisor:
            final_budget_val = float(st.session_state.get("_supervisor_final_budget", total_budget_input))

        c5, c6 = st.columns(2)
        def_s_date = st.session_state.get('temp_start_date', datetime(2026, 1, 1))
        def_e_date = st.session_state.get('temp_end_date', datetime(2026, 1, 31))
        with c5: start_date = st.date_input("開始日", def_s_date)
        with c6: end_date = st.date_input("結束日", def_e_date)
        days_count = (end_date - start_date).days + 1

        # 分段執行：可選擇多個波段日期，未執行日 cue 表顯示空白，檔次依執行天數計算
        if "use_date_segments" not in st.session_state:
            st.session_state.use_date_segments = False
        if "date_segments" not in st.session_state:
            st.session_state.date_segments = []

        use_date_segments = st.checkbox("分段執行（可選多段波段日期，未執行日顯示空白）", value=st.session_state.use_date_segments, key="use_date_segments")
        if use_date_segments:
            if not st.session_state.date_segments:
                st.session_state.date_segments = [(start_date, end_date)]
            # 若先前儲存的波段超出目前起訖，夾在範圍內，避免 date_input 報錯
            segs_raw = list(st.session_state.date_segments)
            segs = [(max(start_date, min(end_date, s)), max(start_date, min(end_date, e))) for s, e in segs_raw]
            if segs != segs_raw:
                st.session_state.date_segments = segs
            for i in range(len(segs)):
                col_a, col_b, col_c = st.columns([2, 2, 1])
                with col_a:
                    st.date_input("波段開始", value=segs[i][0], key=f"seg_start_{i}", min_value=start_date, max_value=end_date)
                with col_b:
                    st.date_input("波段結束", value=segs[i][1], key=f"seg_end_{i}", min_value=start_date, max_value=end_date)
                with col_c:
                    if st.button("刪除", key=f"seg_del_{i}"):
                        st.session_state.date_segments.pop(i)
                        st.rerun()
            new_segs = []
            for i in range(len(segs)):
                s = st.session_state.get(f"seg_start_{i}")
                e = st.session_state.get(f"seg_end_{i}")
                if s is not None and e is not None:
                    new_segs.append((s, e))
            if new_segs:
                st.session_state.date_segments = new_segs
            if st.button("➕ 新增一波段"):
                st.session_state.date_segments.append((start_date, end_date))
                st.rerun()
            segments = sorted(st.session_state.date_segments)
            active_days = sum((e - s).days + 1 for s, e in segments)
            st.caption(f"執行天數共 **{active_days}** 天（用於計算檔次分配）")
        else:
            segments = [(start_date, end_date)]
            active_days = days_count
            st.session_state.date_segments = []

        total_days = days_count  # 表頭與欄數仍為完整走期
        if use_date_segments:
            st.info(f"📅 走期 **{total_days}** 天（分段執行 **{active_days}** 天）")
        else:
            st.info(f"📅 走期共 **{days_count}** 天")

        _cue_def_sign = datetime.now() + timedelta(days=3)
        if isinstance(_cue_def_sign, datetime):
            _cue_def_sign = _cue_def_sign.date()
        _cue_def_pay = _last_day_of_month_after(end_date)

        # 有分波段（2 個以上）時只顯示「各波段備註」，隱藏「備註欄位設定」；取消分波段或僅一波段時才顯示「備註欄位設定」
        if not (use_date_segments and segments and len(segments) > 1):
            st.session_state.setdefault("cue_sign_deadline", _cue_def_sign)
            st.session_state.setdefault("cue_billing_month", "2026年2月")
            st.session_state.setdefault("cue_payment_date", _cue_def_pay)
            st.session_state.setdefault("cue_live_update_remarks", True)
            with st.expander("📝 備註欄位設定", expanded=False):
                st.caption("回簽／請款月份／付款兌現日（付款兌現日預設為走期結束日的下個月最後一天）；勾選「即時更新」後修改欄位時下方備註條列會動態改變。")
                rc1, rc2, rc3 = st.columns(3)
                sign_deadline = rc1.date_input("回簽及進單期限", st.session_state.get("cue_sign_deadline", _cue_def_sign), key="cue_sign_deadline")
                billing_month = rc2.text_input("請款月份", st.session_state.get("cue_billing_month", "2026年2月"), key="cue_billing_month", placeholder="例：2026年2月")
                payment_date = rc3.date_input("付款兌現日期（預設走期結束下月最後一天）", st.session_state.get("cue_payment_date", _cue_def_pay), key="cue_payment_date")
                cue_live = st.checkbox("依上列日期即時更新備註（修改欄位時下方條列會動態改變）", value=st.session_state.get("cue_live_update_remarks", True), key="cue_live_update_remarks")
                if cue_live:
                    _gen_rem = "\n".join(get_remarks_text(sign_deadline, billing_month, payment_date))
                    st.session_state.cue_remarks_text = _gen_rem
                if not cue_live and st.button("依上列日期重新產生備註", key="cue_rem_refresh"):
                    st.session_state.cue_remarks_text = "\n".join(get_remarks_text(sign_deadline, billing_month, payment_date))
                    st.rerun()
                _rem_display = "\n".join(get_remarks_text(sign_deadline, billing_month, payment_date)) if cue_live else (st.session_state.get("cue_remarks_text", "") or "\n".join(get_remarks_text(sign_deadline, billing_month, payment_date)))
                st.text_area("備註全文", value=_rem_display, height=200, key="cue_remarks_text", label_visibility="collapsed", placeholder="留空則依上列三個日期欄位自動產生。", disabled=cue_live)

        if use_date_segments and segments and len(segments) > 1:
            st.markdown("#### 📝 各波段備註（有分波段時可分別設定）")
            st.caption("每個波段可設定獨立的回簽／請款月份／付款兌現日與備註；付款兌現日預設為**該波段結束日的下個月最後一天**。")
            if "cue_seg_remarks" not in st.session_state:
                st.session_state.cue_seg_remarks = {}
            for i, (seg_start, seg_end) in enumerate(segments):
                _def_pay_seg = _last_day_of_month_after(seg_end)
                st.session_state.cue_seg_remarks.setdefault(i, {}).setdefault("sign_deadline", _cue_def_sign)
                st.session_state.cue_seg_remarks.setdefault(i, {}).setdefault("billing_month", "2026年2月")
                st.session_state.cue_seg_remarks.setdefault(i, {}).setdefault("payment_date", _def_pay_seg)
                st.session_state.cue_seg_remarks.setdefault(i, {}).setdefault("live_update_remarks", True)
                seg_data = st.session_state.cue_seg_remarks[i]
                _s = seg_data.get("sign_deadline", _cue_def_sign)
                _b = seg_data.get("billing_month", "2026年2月")
                _p = seg_data.get("payment_date", _def_pay_seg)
                if isinstance(_s, datetime):
                    _s = _s.date() if hasattr(_s, "date") else _s
                if isinstance(_p, datetime):
                    _p = _p.date() if hasattr(_p, "date") else _p
                with st.expander(f"**波段 {i+1}**：{seg_start} ~ {seg_end} — 本波備註", expanded=(i == 0)):
                    r1, r2, r3 = st.columns(3)
                    with r1:
                        _s = st.date_input("回簽及進單期限", _s, key=f"cue_seg_sign_{i}")
                    with r2:
                        _b = st.text_input("請款月份", _b, key=f"cue_seg_bill_{i}", placeholder="例：2026年2月")
                    with r3:
                        _p = st.date_input("付款兌現日期（預設本波結束下月最後一天）", _p, key=f"cue_seg_pay_{i}")
                    st.session_state.cue_seg_remarks[i]["sign_deadline"] = _s
                    st.session_state.cue_seg_remarks[i]["billing_month"] = _b
                    st.session_state.cue_seg_remarks[i]["payment_date"] = _p
                    seg_live = st.checkbox("依上列日期即時更新備註", value=seg_data.get("live_update_remarks", True), key=f"cue_seg_live_{i}")
                    st.session_state.cue_seg_remarks[i]["live_update_remarks"] = seg_live
                    if seg_live:
                        _gen_seg = "\n".join(get_remarks_text(_s, _b, _p))
                        st.session_state[f"cue_seg_remarks_{i}"] = _gen_seg
                        st.session_state.cue_seg_remarks[i]["remarks_text"] = _gen_seg
                    if not seg_live and st.button("依上列日期重新產生本波備註", key=f"cue_seg_rem_btn_{i}"):
                        _gen_seg = "\n".join(get_remarks_text(_s, _b, _p))
                        st.session_state[f"cue_seg_remarks_{i}"] = _gen_seg
                        st.session_state.cue_seg_remarks[i]["remarks_text"] = _gen_seg
                        st.rerun()
                    _seg_display = "\n".join(get_remarks_text(_s, _b, _p)) if seg_live else (st.session_state.get(f"cue_seg_remarks_{i}", "") or "\n".join(get_remarks_text(_s, _b, _p)))
                    st.text_area("本波備註", value=_seg_display, height=180, key=f"cue_seg_remarks_{i}", label_visibility="collapsed", disabled=seg_live)

        st.markdown("### 3. 媒體投放設定")
        col_cb1, col_cb2, col_cb3 = st.columns(3)
        
        # Slider 連動邏輯
        def on_media_change():
            active = []
            if st.session_state.get("cb_rad"): active.append("rad_share")
            if st.session_state.get("cb_fv"): active.append("fv_share")
            if st.session_state.get("cb_cf"): active.append("cf_share")
            if not active: return
            share = 100 // len(active)
            for key in active: st.session_state[key] = share
            rem = 100 - sum([st.session_state[k] for k in active])
            st.session_state[active[0]] += rem

        def on_slider_change(changed_key):
            active = []
            if st.session_state.get("cb_rad"): active.append("rad_share")
            if st.session_state.get("cb_fv"): active.append("fv_share")
            if st.session_state.get("cb_cf"): active.append("cf_share")
            others = [k for k in active if k != changed_key]
            if not others:
                st.session_state[changed_key] = 100
            elif len(others) == 1:
                val = st.session_state[changed_key]
                st.session_state[others[0]] = max(0, 100 - val)
            elif len(others) == 2:
                val = st.session_state[changed_key]
                rem = max(0, 100 - val)
                k1, k2 = others[0], others[1]
                sum_others = st.session_state[k1] + st.session_state[k2]
                if sum_others == 0:
                    st.session_state[k1] = rem // 2
                    st.session_state[k2] = rem - st.session_state[k1]
                else:
                    ratio = st.session_state[k1] / sum_others
                    st.session_state[k1] = int(rem * ratio)
                    st.session_state[k2] = rem - st.session_state[k1]

        def on_sec_slider_change(media_prefix, changed_sec, all_secs):
            key_changed = f"{media_prefix}{changed_sec}"
            new_val = st.session_state[key_changed]
            rem = 100 - new_val

            others = [s for s in all_secs if s != changed_sec]
            if not others:
                st.session_state[key_changed] = 100
                return

            current_sum_others = sum([st.session_state[f"{media_prefix}{s}"] for s in others])

            for i, s in enumerate(others):
                other_key = f"{media_prefix}{s}"
                if current_sum_others == 0:
                    new_other_val = rem // len(others)
                    if i == len(others) - 1:
                        new_other_val = rem - sum([st.session_state[f"{media_prefix}{x}"] for x in others if x != s])
                else:
                    ratio = st.session_state[other_key] / current_sum_others
                    new_other_val = int(rem * ratio)
                    if i == len(others) - 1:
                        allocated = new_val + sum([st.session_state[f"{media_prefix}{x}"] for x in others if x != s])
                        new_other_val = 100 - allocated

                st.session_state[other_key] = max(0, new_other_val)

        def auto_fill_region_remainder(prefix, region_list):
            """將目前未填（空或 0）的區域均分剩餘比例，使加總接近 100%。寫入 pending key，下一輪再套用到 widget key（避免 Streamlit 不允許同 run 內改 widget key）。"""
            current = {}
            for r in region_list:
                k = f"{prefix}region_{r}"
                val = st.session_state.get(k, "")
                if val == "" or val is None:
                    val = 0.0
                else:
                    try:
                        val = float(val)
                    except (TypeError, ValueError):
                        val = 0.0
                current[r] = max(0.0, min(100.0, val))
            unfilled = [r for r in region_list if current[r] == 0]
            total_filled = sum(current[r] for r in region_list) - sum(current[r] for r in unfilled)
            result = {}
            if not unfilled:
                n = len(region_list)
                base = round(100.0 / n, 1)
                for i, r in enumerate(region_list):
                    v = round(100.0 - base * (n - 1), 1) if i == n - 1 else base
                    result[r] = str(v)
            else:
                remainder = 100.0 - total_filled
                per = remainder / len(unfilled)
                for i, r in enumerate(unfilled):
                    if i == len(unfilled) - 1:
                        v = round(remainder - per * (len(unfilled) - 1), 1)
                    else:
                        v = round(per, 1)
                    result[r] = str(v)
                for r in region_list:
                    if r not in result:
                        result[r] = str(current[r]) if current[r] == int(current[r]) else str(round(current[r], 1))
            total = sum(float(result.get(r, 0) or 0) for r in region_list)
            if abs(total - 100) > 0.01:
                first_r = region_list[0]
                cur = float(result.get(first_r, 0) or 0)
                result[first_r] = str(round(cur + (100 - total), 1))
            st.session_state[f"_{prefix}region_pending"] = result

        def clear_region_values(prefix, region_list):
            """清空所有區域比例；寫入 pending key，下一輪再套用。"""
            st.session_state[f"_{prefix}region_pending"] = {r: "" for r in region_list}

        def on_region_slider_change(media_prefix, changed_region, all_regions):
            """自訂區域比例：使用者挪動某一區時，其餘區「均分」剩餘比例，使未挪動的區數值都一樣，
            這樣只有被加重的那一區會高於最低比例、觸發個別計價，其他區維持全省計價。"""
            key_changed = f"{media_prefix}region_{changed_region}"
            new_val = st.session_state[key_changed]
            rem = max(0, 100 - new_val)
            others = [r for r in all_regions if r != changed_region]
            if not others:
                st.session_state[key_changed] = 100
                return
            # 剩餘比例均分給其他區，避免未挪動的區還要另外計價
            base = rem // len(others)
            remainder = rem - base * len(others)
            for i, r in enumerate(others):
                other_key = f"{media_prefix}region_{r}"
                st.session_state[other_key] = base + (1 if i < remainder else 0)

        is_rad = col_cb1.checkbox("全家廣播", key="cb_rad", on_change=on_media_change)
        is_fv = col_cb2.checkbox("新鮮視", key="cb_fv", on_change=on_media_change)
        is_cf = col_cb3.checkbox("家樂福", key="cb_cf", on_change=on_media_change)

        m1, m2, m3 = st.columns(3)
        config = {}
        
        if is_rad:
            with m1:
                st.markdown("#### 📻 全家廣播")
                
                # 修改：移除預設值參數，完全依賴 key (已在 DEFAULT_STATES 初始化)
                is_nat = st.checkbox("全省聯播", key="rad_nat")
                regs = ["全省"] if is_nat else st.multiselect("區域", REGIONS_ORDER, key="rad_reg")
                if not is_nat and len(regs) == 6:
                    is_nat = True
                    regs = ["全省"]
                    st.info("✅ 已選滿6區，自動轉為全省聯播")
                
                # 修改：移除預設值參數
                secs = st.multiselect("秒數", DURATIONS, key="rad_sec")
                st.slider("預算 %", 0, 100, key="rad_share", on_change=on_slider_change, args=("rad_share",))
                
                sorted_secs = sorted(secs)
                if sorted_secs:
                    keys_to_check = [f"rs_{s}" for s in sorted_secs]
                    if any(k not in st.session_state for k in keys_to_check):
                        default_val = 100 // len(sorted_secs)
                        for i, s in enumerate(sorted_secs):
                            k = f"rs_{s}"
                            if i == len(sorted_secs) - 1:
                                st.session_state[k] = 100 - (default_val * (len(sorted_secs)-1))
                            else:
                                st.session_state[k] = default_val
                    
                    sec_shares = {}
                    for s in sorted_secs:
                        st.slider(
                            f"{s}秒 %", 0, 100,
                            key=f"rs_{s}",
                            on_change=on_sec_slider_change,
                            args=("rs_", s, sorted_secs)
                        )
                        sec_shares[s] = st.session_state[f"rs_{s}"]

                    # 自訂區域比例：一開始格子皆為空；可輸入加重區域，再按「填完加重區域，均分剩餘」或「清空數值」
                    region_share_regions = REGIONS_ORDER if is_nat else regs
                    with st.expander("📍 自訂區域比例", expanded=False):
                        use_region_share_rad = st.checkbox("啟用自訂區域比例", key="rad_use_region_share", help="可依所選區域調整預算分配比例；全省時最低比例用全省計價，其餘用各區計價。格子一開始為空，可輸入加重區域後按「填完加重區域，均分剩餘」。")
                        if use_region_share_rad:
                            pending_key = "_rad_region_pending"
                            if pending_key in st.session_state:
                                pending = st.session_state[pending_key]
                                for r in region_share_regions:
                                    st.session_state[f"rad_region_{r}"] = pending.get(r, "")
                                del st.session_state[pending_key]
                            keys_region = [f"rad_region_{r}" for r in region_share_regions]
                            if any(k not in st.session_state for k in keys_region):
                                for k in keys_region:
                                    st.session_state[k] = ""
                            st.session_state["_rad_region_list"] = region_share_regions
                            cols_rad = st.columns(len(region_share_regions))
                            region_shares_rad = {}
                            for idx, r in enumerate(region_share_regions):
                                with cols_rad[idx]:
                                    st.text_input(
                                        f"{r} %",
                                        key=f"rad_region_{r}",
                                        placeholder="",
                                        label_visibility="visible",
                                    )
                                    raw = st.session_state.get(f"rad_region_{r}", "") or ""
                                    try:
                                        region_shares_rad[r] = float(raw.strip()) if raw.strip() else 0.0
                                    except ValueError:
                                        region_shares_rad[r] = 0.0
                            btn_col1, btn_col2 = st.columns(2)
                            with btn_col1:
                                if st.button("填完加重區域，均分剩餘", key="rad_auto_fill_btn", help="已填的區域保留，空欄位均分剩餘比例使加總為 100%"):
                                    auto_fill_region_remainder("rad_", st.session_state.get("_rad_region_list", REGIONS_ORDER))
                                    st.rerun()
                            with btn_col2:
                                if st.button("清空數值", key="rad_clear_btn", help="清空所有區域比例"):
                                    clear_region_values("rad_", st.session_state.get("_rad_region_list", REGIONS_ORDER))
                                    st.rerun()
                        else:
                            region_shares_rad = None

                    config["全家廣播"] = {"is_national": is_nat, "regions": regs, "sec_shares": sec_shares, "share": st.session_state.rad_share, "region_shares": region_shares_rad}

        if is_fv:
            with m2:
                st.markdown("#### 📺 新鮮視")
                
                # 修改：移除預設值參數，完全依賴 key
                is_nat = st.checkbox("全省聯播", key="fv_nat")
                regs = ["全省"] if is_nat else st.multiselect("區域", REGIONS_ORDER, key="fv_reg")
                if not is_nat and len(regs) == 6:
                    is_nat = True
                    regs = ["全省"]
                    st.info("✅ 已選滿6區，自動轉為全省聯播")
                
                # 修改：移除預設值參數
                secs = st.multiselect("秒數", DURATIONS, key="fv_sec")
                st.slider("預算 %", 0, 100, key="fv_share", on_change=on_slider_change, args=("fv_share",))
                
                sorted_secs = sorted(secs)
                if sorted_secs:
                    keys_to_check = [f"fs_{s}" for s in sorted_secs]
                    if any(k not in st.session_state for k in keys_to_check):
                        default_val = 100 // len(sorted_secs)
                        for i, s in enumerate(sorted_secs):
                            k = f"fs_{s}"
                            if i == len(sorted_secs) - 1:
                                st.session_state[k] = 100 - (default_val * (len(sorted_secs)-1))
                            else:
                                st.session_state[k] = default_val
                    
                    sec_shares = {}
                    for s in sorted_secs:
                        st.slider(
                            f"{s}秒 %", 0, 100,
                            key=f"fs_{s}",
                            on_change=on_sec_slider_change,
                            args=("fs_", s, sorted_secs)
                        )
                        sec_shares[s] = st.session_state[f"fs_{s}"]

                    region_share_regions_fv = REGIONS_ORDER if is_nat else regs
                    with st.expander("📍 自訂區域比例", expanded=False):
                        use_region_share_fv = st.checkbox("啟用自訂區域比例", key="fv_use_region_share", help="可依所選區域調整預算分配比例；全省時最低比例用全省計價，其餘用各區計價。格子一開始為空，可輸入加重區域後按「填完加重區域，均分剩餘」。")
                        if use_region_share_fv:
                            pending_key_fv = "_fv_region_pending"
                            if pending_key_fv in st.session_state:
                                pending_fv = st.session_state[pending_key_fv]
                                for r in region_share_regions_fv:
                                    st.session_state[f"fv_region_{r}"] = pending_fv.get(r, "")
                                del st.session_state[pending_key_fv]
                            keys_region_fv = [f"fv_region_{r}" for r in region_share_regions_fv]
                            if any(k not in st.session_state for k in keys_region_fv):
                                for k in keys_region_fv:
                                    st.session_state[k] = ""
                            st.session_state["_fv_region_list"] = region_share_regions_fv
                            cols_fv = st.columns(len(region_share_regions_fv))
                            region_shares_fv = {}
                            for idx, r in enumerate(region_share_regions_fv):
                                with cols_fv[idx]:
                                    st.text_input(
                                        f"{r} %",
                                        key=f"fv_region_{r}",
                                        placeholder="",
                                        label_visibility="visible",
                                    )
                                    raw = st.session_state.get(f"fv_region_{r}", "") or ""
                                    try:
                                        region_shares_fv[r] = float(raw.strip()) if raw.strip() else 0.0
                                    except ValueError:
                                        region_shares_fv[r] = 0.0
                            btn_col1_fv, btn_col2_fv = st.columns(2)
                            with btn_col1_fv:
                                if st.button("填完加重區域，均分剩餘", key="fv_auto_fill_btn", help="已填的區域保留，空欄位均分剩餘比例使加總為 100%"):
                                    auto_fill_region_remainder("fv_", st.session_state.get("_fv_region_list", REGIONS_ORDER))
                                    st.rerun()
                            with btn_col2_fv:
                                if st.button("清空數值", key="fv_clear_btn", help="清空所有區域比例"):
                                    clear_region_values("fv_", st.session_state.get("_fv_region_list", REGIONS_ORDER))
                                    st.rerun()
                        else:
                            region_shares_fv = None

                    config["新鮮視"] = {"is_national": is_nat, "regions": regs, "sec_shares": sec_shares, "share": st.session_state.fv_share, "region_shares": region_shares_fv}

        if is_cf:
            with m3:
                st.markdown("#### 🛒 家樂福")
                # 修改：移除預設值參數
                secs = st.multiselect("秒數", DURATIONS, key="cf_sec")
                st.slider("預算 %", 0, 100, key="cf_share", on_change=on_slider_change, args=("cf_share",))
                
                sorted_secs = sorted(secs)
                if sorted_secs:
                    keys_to_check = [f"cs_{s}" for s in sorted_secs]
                    if any(k not in st.session_state for k in keys_to_check):
                        default_val = 100 // len(sorted_secs)
                        for i, s in enumerate(sorted_secs):
                            k = f"cs_{s}"
                            if i == len(sorted_secs) - 1:
                                st.session_state[k] = 100 - (default_val * (len(sorted_secs)-1))
                            else:
                                st.session_state[k] = default_val
                    
                    sec_shares = {}
                    for s in sorted_secs:
                        st.slider(
                            f"{s}秒 %", 0, 100, 
                            key=f"cs_{s}", 
                            on_change=on_sec_slider_change, 
                            args=("cs_", s, sorted_secs)
                        )
                        sec_shares[s] = st.session_state[f"cs_{s}"]
                
                    config["家樂福"] = {"regions": ["全省"], "sec_shares": sec_shares, "share": st.session_state.cf_share}

        # --- 運算與輸出邏輯 ---
        if config:
            # 檔次依「執行天數」計算；若有分段則產出後再展開為完整日曆（未執行日填 0）
            # 交換合約：檔次依定價計算（use_list_price_for_spots=True），且不套用回饋
            rows, total_list_accum, logs = calculate_plan_data(config, total_budget_input, active_days, PRICING_DB, SEC_FACTORS, STORE_COUNTS_NUM, REGIONS_ORDER, use_list_price_for_spots=is_barter_contract)
            if use_date_segments and segments:
                for row in rows:
                    row["schedule"] = expand_schedule_to_calendar(row["schedule"], segments, start_date, end_date)
            # 回饋贈檔：最多三種回饋並存，每種可獨立選擇要/不要（交換合約不提供回饋）
            qual = get_rebate_qualification_detail(config, total_budget_input, active_days) if not is_barter_contract else {}
            apply_nat_rad = False
            rebate_nat_destination = None
            apply_nat_cf = False
            apply_region_rad = False
            rebate_region_destination = None
            if not is_barter_contract:
                # 回饋 1：全省全家廣播達標 → 可選「全省全家預算×% 回饋全家」或「回饋家樂福」
                if qual.get("nat_rad"):
                    apply_nat_rad = st.checkbox(
                        "套用「全省全家廣播達標」回饋",
                        value=st.session_state.get("apply_nat_rad_rebate", True),
                        key="apply_nat_rad_rebate",
                        help="依全省全家廣播預算×回饋% 回饋到全家廣播或家樂福（下方擇一）。",
                    )
                    if apply_nat_rad:
                        rebate_nat_destination = st.radio(
                            "回饋到",
                            options=["全家廣播", "家樂福"],
                            index=0,
                            key="rebate_nat_destination",
                            horizontal=True,
                            help="可選擇依該預算×% 回饋在「全省全家廣播」或「全省家樂福」。",
                        )
                else:
                    for k in ("apply_nat_rad_rebate", "rebate_nat_destination"):
                        if k in st.session_state:
                            del st.session_state[k]
                # 回饋 2：家樂福達標 → 家樂福預算×% 回饋家樂福（與回饋1可併存）
                if qual.get("nat_cf"):
                    apply_nat_cf = st.checkbox(
                        "套用「家樂福達標」回饋",
                        value=st.session_state.get("apply_nat_cf_rebate", True),
                        key="apply_nat_cf_rebate",
                        help="依家樂福預算×回饋% 回饋到全省家樂福。可與「全省全家達標→家樂福」併存。",
                    )
                else:
                    if "apply_nat_cf_rebate" in st.session_state:
                        del st.session_state["apply_nat_cf_rebate"]
                # 回饋 3：單區全家達標 → 選回饋顯示區域
                if qual.get("region_rad"):
                    apply_region_rad = st.checkbox(
                        "套用「單區全家廣播達標」回饋",
                        value=st.session_state.get("apply_region_rad_rebate", True),
                        key="apply_region_rad_rebate",
                        help="單區達標僅能回饋單區全家廣播，可選擇要顯示在哪一區。",
                    )
                    if apply_region_rad:
                        region_options = ["北區", "桃竹苗", "中區", "雲嘉南"]
                        idx_opt = 0
                        if st.session_state.get("rebate_region_destination") in region_options:
                            idx_opt = region_options.index(st.session_state["rebate_region_destination"])
                        rebate_region_destination = st.selectbox(
                            "回饋顯示區域",
                            options=region_options,
                            index=idx_opt,
                            key="rebate_region_destination",
                            help="可選擇要顯示在北區/桃竹苗/中區/雲嘉南哪一區，不必與購買區域相同。",
                        )
                else:
                    for k in ("apply_region_rad_rebate", "rebate_region_destination"):
                        if k in st.session_state:
                            del st.session_state[k]
            rebate_result = compute_rebate_rows(
                config, total_budget_input, active_days, rows, PRICING_DB, SEC_FACTORS, STORE_COUNTS_NUM, REGIONS_ORDER,
                apply_nat_rad=apply_nat_rad, rebate_nat_destination=rebate_nat_destination,
                apply_nat_cf=apply_nat_cf,
                apply_region_rad=apply_region_rad, rebate_region_destination=rebate_region_destination,
            ) if not is_barter_contract else []
            if isinstance(rebate_result, tuple):
                rebate_inserts, rebate_logs = rebate_result
            else:
                rebate_inserts = rebate_result
                rebate_logs = []
            rebate_summary = get_rebate_summary_text(
                rebate_inserts, config=config, total_budget=total_budget_input, active_days=active_days,
                qual=qual, apply_nat_rad=apply_nat_rad, rebate_nat_destination=rebate_nat_destination,
                apply_nat_cf=apply_nat_cf, apply_region_rad=apply_region_rad, rebate_region_destination=rebate_region_destination,
            )
            if not is_barter_contract:
                apply_rebate = st.checkbox("套用回饋贈檔", value=st.session_state.get("apply_rebate", False), key="apply_rebate", help="勾選後，表內會顯示符合門檻的回饋贈檔列。")
            else:
                apply_rebate = False
            if rebate_summary and not is_barter_contract:
                st.caption(f"📌 **本次可回饋：** {rebate_summary}")
            # 主管加贈回饋：僅主管可填 %，回饋金額可任意分配至平台／區域／秒數（仿 3. 媒體投放設定，無自訂區域比例）
            bonus_pct_val = None
            bonus_config = {}
            if not is_barter_contract and st.session_state.get("is_supervisor"):
                bonus_input = st.number_input("加贈回饋 %（主管）", min_value=0, max_value=100, value=st.session_state.get("bonus_rebate_pct", 0) or 0, step=1, key="bonus_rebate_pct", help="例如廣告預算 100 萬、回饋 5% → 5 萬元可於下方自由分配至平台、區域、秒數及比重。")
                if bonus_input and bonus_input > 0:
                    bonus_pct_val = int(bonus_input)
                    rebate_budget = int(round(total_budget_input * bonus_pct_val / 100.0))
                    st.caption(f"💰 **主管回饋金額：${rebate_budget:,}**（可於下方分配至平台／區域／秒數）")

                    # 主管回饋分配 UI：仿「3. 媒體投放設定」，無自訂區域比例
                    def bonus_on_media_change():
                        active = []
                        if st.session_state.get("bonus_cb_rad"): active.append("bonus_rad_share")
                        if st.session_state.get("bonus_cb_fv"): active.append("bonus_fv_share")
                        if st.session_state.get("bonus_cb_cf"): active.append("bonus_cf_share")
                        if not active: return
                        share = 100 // len(active)
                        for key in active: st.session_state[key] = share
                        rem = 100 - sum([st.session_state[k] for k in active])
                        st.session_state[active[0]] += rem

                    def bonus_on_slider_change(changed_key):
                        active = []
                        if st.session_state.get("bonus_cb_rad"): active.append("bonus_rad_share")
                        if st.session_state.get("bonus_cb_fv"): active.append("bonus_fv_share")
                        if st.session_state.get("bonus_cb_cf"): active.append("bonus_cf_share")
                        others = [k for k in active if k != changed_key]
                        if not others:
                            st.session_state[changed_key] = 100
                        elif len(others) == 1:
                            val = st.session_state[changed_key]
                            st.session_state[others[0]] = max(0, 100 - val)
                        else:
                            val = st.session_state[changed_key]
                            rem = max(0, 100 - val)
                            k1, k2 = others[0], others[1]
                            sum_others = st.session_state[k1] + st.session_state[k2]
                            if sum_others == 0:
                                st.session_state[k1] = rem // 2
                                st.session_state[k2] = rem - st.session_state[k1]
                            else:
                                ratio = st.session_state[k1] / sum_others
                                st.session_state[k1] = int(rem * ratio)
                                st.session_state[k2] = rem - st.session_state[k1]

                    def bonus_on_sec_slider_change(media_prefix, changed_sec, all_secs):
                        key_changed = f"{media_prefix}{changed_sec}"
                        new_val = st.session_state[key_changed]
                        rem = 100 - new_val
                        others = [s for s in all_secs if s != changed_sec]
                        if not others:
                            st.session_state[key_changed] = 100
                            return
                        current_sum_others = sum([st.session_state.get(f"{media_prefix}{s}", 0) for s in others])
                        for i, s in enumerate(others):
                            other_key = f"{media_prefix}{s}"
                            if current_sum_others == 0:
                                new_other_val = rem // len(others)
                                if i == len(others) - 1:
                                    new_other_val = rem - sum([st.session_state.get(f"{media_prefix}{x}", 0) for x in others if x != s])
                            else:
                                ratio = st.session_state.get(other_key, 0) / current_sum_others
                                new_other_val = int(rem * ratio)
                                if i == len(others) - 1:
                                    allocated = new_val + sum([st.session_state.get(f"{media_prefix}{x}", 0) for x in others if x != s])
                                    new_other_val = 100 - allocated
                            st.session_state[other_key] = max(0, new_other_val)

                    st.markdown("#### 主管回饋分配（平台／區域／秒數／比重）")
                    b_cb1, b_cb2, b_cb3 = st.columns(3)
                    is_b_rad = b_cb1.checkbox("全家廣播", key="bonus_cb_rad", on_change=bonus_on_media_change)
                    is_b_fv = b_cb2.checkbox("新鮮視", key="bonus_cb_fv", on_change=bonus_on_media_change)
                    is_b_cf = b_cb3.checkbox("家樂福", key="bonus_cb_cf", on_change=bonus_on_media_change)

                    b_m1, b_m2, b_m3 = st.columns(3)
                    if is_b_rad:
                        with b_m1:
                            st.markdown("##### 📻 全家廣播")
                            is_b_nat_rad = st.checkbox("全省聯播", key="bonus_rad_nat")
                            b_regs_rad = ["全省"] if is_b_nat_rad else st.multiselect("區域", REGIONS_ORDER, key="bonus_rad_reg")
                            if not is_b_nat_rad and len(b_regs_rad) == 6:
                                b_regs_rad = ["全省"]
                                st.info("✅ 已選滿6區，視為全省")
                            b_secs_rad = st.multiselect("秒數", DURATIONS, key="bonus_rad_sec")
                            st.slider("預算 %", 0, 100, key="bonus_rad_share", on_change=bonus_on_slider_change, args=("bonus_rad_share",))
                            sorted_b_secs_rad = sorted(b_secs_rad)
                            if sorted_b_secs_rad:
                                for k in [f"bonus_rs_{s}" for s in sorted_b_secs_rad]:
                                    if k not in st.session_state:
                                        default_val = 100 // len(sorted_b_secs_rad)
                                        for i, s in enumerate(sorted_b_secs_rad):
                                            kk = f"bonus_rs_{s}"
                                            st.session_state[kk] = 100 - (default_val * (len(sorted_b_secs_rad) - 1)) if i == len(sorted_b_secs_rad) - 1 else default_val
                                        break
                                b_sec_shares_rad = {}
                                for s in sorted_b_secs_rad:
                                    st.slider(f"{s}秒 %", 0, 100, key=f"bonus_rs_{s}", on_change=bonus_on_sec_slider_change, args=("bonus_rs_", s, sorted_b_secs_rad))
                                    b_sec_shares_rad[s] = st.session_state.get(f"bonus_rs_{s}", 0)
                                bonus_config["全家廣播"] = {"is_national": is_b_nat_rad, "regions": b_regs_rad if b_regs_rad else ["全省"], "sec_shares": b_sec_shares_rad, "share": st.session_state.get("bonus_rad_share", 0)}
                    if is_b_fv:
                        with b_m2:
                            st.markdown("##### 📺 新鮮視")
                            is_b_nat_fv = st.checkbox("全省聯播", key="bonus_fv_nat")
                            b_regs_fv = ["全省"] if is_b_nat_fv else st.multiselect("區域", REGIONS_ORDER, key="bonus_fv_reg")
                            if not is_b_nat_fv and len(b_regs_fv) == 6:
                                b_regs_fv = ["全省"]
                                st.info("✅ 已選滿6區，視為全省")
                            b_secs_fv = st.multiselect("秒數", DURATIONS, key="bonus_fv_sec")
                            st.slider("預算 %", 0, 100, key="bonus_fv_share", on_change=bonus_on_slider_change, args=("bonus_fv_share",))
                            sorted_b_secs_fv = sorted(b_secs_fv)
                            if sorted_b_secs_fv:
                                for k in [f"bonus_fs_{s}" for s in sorted_b_secs_fv]:
                                    if k not in st.session_state:
                                        default_val = 100 // len(sorted_b_secs_fv)
                                        for i, s in enumerate(sorted_b_secs_fv):
                                            kk = f"bonus_fs_{s}"
                                            st.session_state[kk] = 100 - (default_val * (len(sorted_b_secs_fv) - 1)) if i == len(sorted_b_secs_fv) - 1 else default_val
                                        break
                                b_sec_shares_fv = {}
                                for s in sorted_b_secs_fv:
                                    st.slider(f"{s}秒 %", 0, 100, key=f"bonus_fs_{s}", on_change=bonus_on_sec_slider_change, args=("bonus_fs_", s, sorted_b_secs_fv))
                                    b_sec_shares_fv[s] = st.session_state.get(f"bonus_fs_{s}", 0)
                                bonus_config["新鮮視"] = {"is_national": is_b_nat_fv, "regions": b_regs_fv if b_regs_fv else ["全省"], "sec_shares": b_sec_shares_fv, "share": st.session_state.get("bonus_fv_share", 0)}
                    if is_b_cf:
                        with b_m3:
                            st.markdown("##### 🛒 家樂福")
                            b_secs_cf = st.multiselect("秒數", DURATIONS, key="bonus_cf_sec")
                            st.slider("預算 %", 0, 100, key="bonus_cf_share", on_change=bonus_on_slider_change, args=("bonus_cf_share",))
                            sorted_b_secs_cf = sorted(b_secs_cf)
                            if sorted_b_secs_cf:
                                for k in [f"bonus_cs_{s}" for s in sorted_b_secs_cf]:
                                    if k not in st.session_state:
                                        default_val = 100 // len(sorted_b_secs_cf)
                                        for i, s in enumerate(sorted_b_secs_cf):
                                            kk = f"bonus_cs_{s}"
                                            st.session_state[kk] = 100 - (default_val * (len(sorted_b_secs_cf) - 1)) if i == len(sorted_b_secs_cf) - 1 else default_val
                                        break
                                b_sec_shares_cf = {}
                                for s in sorted_b_secs_cf:
                                    st.slider(f"{s}秒 %", 0, 100, key=f"bonus_cs_{s}", on_change=bonus_on_sec_slider_change, args=("bonus_cs_", s, sorted_b_secs_cf))
                                    b_sec_shares_cf[s] = st.session_state.get(f"bonus_cs_{s}", 0)
                                bonus_config["家樂福"] = {"regions": ["全省"], "sec_shares": b_sec_shares_cf, "share": st.session_state.get("bonus_cf_share", 0)}

            # 業務加贈檔次：每個「同平台、同秒數、同區域」可選加贈、每日檔次、加贈日期區間
            row_groups = get_row_groups(rows, REGIONS_ORDER)
            custom_bonus_config = {}
            if st.session_state.get("is_supervisor", False):
                with st.expander("📌 業務加贈檔次", expanded=False):
                    st.caption("在相同平台、相同秒數、相同區域的廣告下方可加一列「加贈檔次」，設定每日加贈幾檔及加贈日期區間（須在走期內）。")
                    for g in row_groups:
                        media, sec, region_key = g["key"]
                        skey = f"cb_{media}_{sec}_{region_key}".replace(" ", "_")
                        label = f"【{media}】 {sec}秒 - {region_key}"
                        st.markdown(f"**{label}**")
                        c1, c2, c3, c4 = st.columns([1, 2, 2, 2])
                        with c1:
                            enabled = st.checkbox("加贈", value=st.session_state.get(f"cb_en_{skey}", False), key=f"cb_en_{skey}", label_visibility="collapsed")
                        with c2:
                            spots_per_day = st.number_input("每日加贈檔次", min_value=0, max_value=999, value=st.session_state.get(f"cb_spots_{skey}", 0) or 0, step=1, key=f"cb_spots_{skey}")
                        with c3:
                            b_start = st.date_input("加贈開始日", value=start_date, min_value=start_date, max_value=end_date, key=f"cb_start_{skey}")
                        with c4:
                            b_end = st.date_input("加贈結束日", value=end_date, min_value=start_date, max_value=end_date, key=f"cb_end_{skey}")
                        if enabled and spots_per_day and b_start and b_end:
                            custom_bonus_config[g["key"]] = {"enabled": True, "spots_per_day": spots_per_day, "date_start": b_start, "date_end": b_end}
            else:
                with st.expander("📌 業務加贈檔次", expanded=False):
                    st.caption("僅主管可設定業務加贈檔次。")

            # 合併回饋與加贈：兩者皆依「原始 rows」的 index；同一 index 先插門檻回饋再插業務加贈、主管加贈
            all_inserts = []
            bonus_rebate_logs = []
            if apply_rebate and rebate_inserts:
                all_inserts.extend(rebate_inserts)
            if custom_bonus_config:
                custom_inserts, _ = compute_custom_bonus_rows(rows, custom_bonus_config, start_date, end_date, PRICING_DB, SEC_FACTORS, STORE_COUNTS_NUM, REGIONS_ORDER)
                if custom_inserts:
                    all_inserts.extend(custom_inserts)
            if bonus_pct_val is not None and bonus_pct_val > 0 and bonus_config:
                rebate_budget = int(round(total_budget_input * bonus_pct_val / 100.0))
                bonus_inserts, bonus_rebate_logs = compute_bonus_rebate_rows_from_allocation(bonus_config, rebate_budget, active_days, rows, PRICING_DB, SEC_FACTORS, STORE_COUNTS_NUM, REGIONS_ORDER)
                if bonus_inserts:
                    all_inserts.extend(bonus_inserts)
            if all_inserts:
                # 排序：同 index 時 門檻回饋(0) -> 業務加贈(1) -> 主管加贈(2)
                all_inserts.sort(key=lambda x: (x[0], 0 if (x[1].get("is_rebate") and not x[1].get("is_bonus_rebate")) else (1 if x[1].get("is_custom_bonus") else 2)))
                rows = merge_rebate_into_rows(rows, all_inserts)
            if use_date_segments and segments:
                for row in rows:
                    if (row.get("is_rebate") or row.get("is_custom_bonus")) and "schedule" in row:
                        row["schedule"] = expand_schedule_to_calendar(row["schedule"], segments, start_date, end_date)
            prod_cost = prod_cost_input 
            vat = _round_half_up(final_budget_val * 0.05)
            grand_total = final_budget_val + vat
            
            p_str = f"{'、'.join([f'{s}秒' for s in sorted(list(set(r['seconds'] for r in rows)))])} {product_name}"
            _cue_live = st.session_state.get("cue_live_update_remarks", True)
            _sig = st.session_state.get("cue_sign_deadline", datetime.now().date())
            _bill = st.session_state.get("cue_billing_month", "2026年2月")
            _pay = st.session_state.get("cue_payment_date", _last_day_of_month_after(end_date))
            if isinstance(_sig, datetime):
                _sig = _sig.date() if hasattr(_sig, "date") else _sig
            if isinstance(_pay, datetime):
                _pay = _pay.date() if hasattr(_pay, "date") else _pay
            if _cue_live:
                rem = get_remarks_text(_sig, _bill, _pay)
            else:
                _rem_raw = (st.session_state.get("cue_remarks_text", "") or "").strip().split("\n")
                rem = [line.strip() for line in _rem_raw if line.strip()]
                if not rem:
                    rem = get_remarks_text(_sig, _bill, _pay)
            sign_deadline = _sig
            billing_month = _bill
            payment_date = _pay

            # 顯示運算邏輯（含回饋檔次計算、主管額外回饋計算）
            render_logic_panel(logs, use_list_price_for_spots=is_barter_contract, rebate_logs=rebate_logs if not is_barter_contract else [], bonus_rebate_logs=bonus_rebate_logs)
            can_download_excel = st.session_state.get("is_supervisor", False) or st.session_state.get("allow_sales_excel_download", True)
            can_download_pdf = st.session_state.get("is_supervisor", False) or st.session_state.get("allow_sales_pdf_download", True)
            
            st.markdown("---")
            st.subheader("📥 檔案下載區")

            _cue_def_s = st.session_state.get("cue_sign_deadline", datetime.now().date())
            if isinstance(_cue_def_s, datetime):
                _cue_def_s = _cue_def_s.date() if hasattr(_cue_def_s, "date") else _cue_def_s

            if use_date_segments and segments and len(segments) > 1:
                for i, (seg_start, seg_end) in enumerate(segments):
                    seg_rem = _get_cue_segment_rem(i, seg_end, _cue_def_s)
                    segment_rows = _slice_rows_to_segment(rows, start_date, seg_start, seg_end)
                    seg_days = (seg_end - seg_start).days + 1
                    html_seg = generate_html_preview(segment_rows, seg_days, seg_start, seg_end, client_name, client_tax_id, p_str, format_type, seg_rem, total_list_accum, grand_total, final_budget_val, prod_cost)
                    with st.expander(f"波段 {i+1}：{seg_start} ~ {seg_end}", expanded=(i == 0)):
                        if isinstance(html_seg, list):
                            for idx, one_html in enumerate(html_seg):
                                st.caption(f"第 {idx+1} 頁")
                                st.components.v1.html(one_html, height=500, scrolling=True)
                        else:
                            st.components.v1.html(html_seg, height=500, scrolling=True)
                        xlsx_seg = generate_excel_from_scratch(format_type, seg_start, seg_end, client_name, client_tax_id, product_name, segment_rows, seg_rem, final_budget_val, prod_cost, sales_person, total_list_accum)
                        pdf_seg, _, _ = xlsx_bytes_to_pdf_bytes(xlsx_seg)
                        c1, c2 = st.columns(2)
                        with c1:
                            if can_download_excel:
                                st.download_button(f"📥 波段{i+1} Excel", xlsx_seg, f"Cue_{safe_filename(client_name)}_波段{i+1}_{seg_start}_{seg_end}.xlsx", key=f"cue_seg_xlsx_{i}", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                            else:
                                st.caption("Excel 下載已由主管關閉")
                        with c2:
                            if pdf_seg and can_download_pdf:
                                st.download_button(f"📥 波段{i+1} PDF", pdf_seg, f"Cue_{safe_filename(client_name)}_波段{i+1}_{seg_start}_{seg_end}.pdf", key=f"cue_seg_pdf_{i}", mime="application/pdf")
                            elif not can_download_pdf:
                                st.caption("PDF 下載已由主管關閉")
                            else:
                                st.caption("PDF 需 LibreOffice")
                _ragic_col = st.columns(3)[2]
                seg0_start, seg0_end = segments[0][0], segments[0][1]
                _seg0_rem = _get_cue_segment_rem(0, seg0_end, _cue_def_s)
                _seg0_rows = _slice_rows_to_segment(rows, start_date, seg0_start, seg0_end)
                _xlsx_ragic = generate_excel_from_scratch(format_type, seg0_start, seg0_end, client_name, client_tax_id, product_name, _seg0_rows, _seg0_rem, final_budget_val, prod_cost, sales_person, total_list_accum)
                _pdf_ragic, _, _ = xlsx_bytes_to_pdf_bytes(_xlsx_ragic)
            else:
                # 單一預覽與下載（無分波段或僅一波段）
                html_preview = generate_html_preview(rows, total_days, start_date, end_date, client_name, client_tax_id, p_str, format_type, rem, total_list_accum, grand_total, final_budget_val, prod_cost)
                if isinstance(html_preview, list):
                    for idx, one_html in enumerate(html_preview):
                        st.caption(f"第 {idx+1} 頁")
                        st.components.v1.html(one_html, height=700, scrolling=True)
                else:
                    st.components.v1.html(html_preview, height=700, scrolling=True)
                xlsx_temp = generate_excel_from_scratch(format_type, start_date, end_date, client_name, client_tax_id, product_name, rows, rem, final_budget_val, prod_cost, sales_person, total_list_accum)
                col_dl1, col_dl2, col_ragic = st.columns([1, 1, 2])
                with col_dl2:
                    pdf_bytes, method, err = xlsx_bytes_to_pdf_bytes(xlsx_temp)
                    if pdf_bytes and can_download_pdf:
                        st.download_button(f"📥 下載 PDF", pdf_bytes, f"Cue_{safe_filename(client_name)}.pdf", key="pdf_dl_btn", mime="application/pdf")
                    elif not can_download_pdf:
                        st.caption("PDF 下載已由主管關閉")
                    else:
                        st.warning(f"PDF 生成失敗: {err}")
                with col_dl1:
                    if can_download_excel:
                        st.download_button("📥 下載 Excel", xlsx_temp, f"Cue_{safe_filename(client_name)}.xlsx", key="xlsx_dl_btn", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    else:
                        st.caption("Excel 下載已由主管關閉")
                _ragic_col = col_ragic
                _xlsx_ragic, _pdf_ragic = xlsx_temp, pdf_bytes

            with _ragic_col:
                st.markdown("#### ☁️ 上傳至 Ragic")
                style_excel_file = st.file_uploader(
                    "自調外觀excel上傳",
                    type=["xlsx", "xlsm", "xls"],
                    key="ragic_style_excel_upload",
                    help="僅能調整外觀及呈列方式，請勿調整每日檔次及金額等重要資訊(訂檔資訊將以程式設定為主)。",
                )
                st.caption("僅能調整外觀及呈列方式，請勿調整每日檔次及金額等重要資訊(訂檔資訊將以程式設定為主)。")
                
                # === [新增功能] 顯示上傳成功的歷史訊息 (不會一閃即逝) ===
                if 'upload_success_msg' in st.session_state:
                    st.success(st.session_state['upload_success_msg'])
                    if st.button("👌 我知道了 (清除訊息)"):
                        del st.session_state['upload_success_msg']
                        st.rerun()
                # ====================================================

                if not st.session_state.ragic_confirm_state:
                    if st.button("🚀 上傳資料至 Ragic", type="primary"):
                        st.session_state.ragic_confirm_state = True
                        st.rerun()
                else:
                    st.warning(f"即將上傳【{client_name} - {product_name}】至 Ragic，請確認？")
                    c_conf1, c_conf2 = st.columns(2)
                    
                    with c_conf1:
                        if st.button("❌ 取消"):
                            st.session_state.ragic_confirm_state = False
                            st.rerun()
                            
                    with c_conf2:
                        if st.button("✅ 確認上傳"):
                            with st.spinner("正在上傳資料與檔案..."):
                                row_groups_upload = get_row_groups(rows, REGIONS_ORDER)
                                campaign_summary = build_ragic_details_full(config, _get_ragic_extra_state(row_groups_upload))
                                platform_detail = build_platform_detail_text(rows, config)
                                platform_text = build_platform_text(rows)
                                # Ragic 追加欄位：總檔數（含一般+回饋+加贈）與秒數聯集
                                total_spots_all = int(sum(
                                    (d if isinstance(d, (int, float)) else 0)
                                    for r in rows for d in (r.get("schedule") or [])
                                ))
                                seconds_union_text = "、".join([
                                    str(s) for s in sorted({int(r.get("seconds")) for r in rows if r.get("seconds") is not None})
                                ])
                                sales_nickname = SALES_MAP.get(sales_person, sales_person)

                                data_payload = {
                                    RAGIC_MAP['client']:     client_name,
                                    RAGIC_MAP['product']:    product_name,
                                    RAGIC_MAP['budget_raw']: total_budget_input,
                                    RAGIC_MAP['budget_fin']: final_budget_val,
                                    RAGIC_MAP['prod_cost']:  prod_cost_input,
                                    RAGIC_MAP['format']:     format_type,
                                    RAGIC_MAP['sales']:      sales_nickname,
                                    RAGIC_MAP['date_start']: str(start_date),
                                    RAGIC_MAP['date_end']:   str(end_date),
                                    RAGIC_MAP['date_sign']:  str(sign_deadline),
                                    RAGIC_MAP['bill_month']: billing_month,
                                    RAGIC_MAP['date_pay']:   str(payment_date),
                                    RAGIC_MAP['details']:    campaign_summary,
                                    RAGIC_MAP['tax_id']:     client_tax_id
                                }
                                if RAGIC_MAP.get('platform_detail'):
                                    data_payload[RAGIC_MAP['platform_detail']] = platform_detail
                                if RAGIC_MAP.get('platform'):
                                    data_payload[RAGIC_MAP['platform']] = platform_text
                                if RAGIC_MAP.get('total_spots'):
                                    data_payload[RAGIC_MAP['total_spots']] = total_spots_all
                                if RAGIC_MAP.get('seconds_union'):
                                    data_payload[RAGIC_MAP['seconds_union']] = seconds_union_text

                                files_payload = {}
                                files_payload[RAGIC_MAP['file_xls']] = (
                                    f"Cue_{safe_filename(client_name)}.xlsx", 
                                    _xlsx_ragic,
                                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                                )
                                if _pdf_ragic:
                                    files_payload[RAGIC_MAP['file_pdf']] = (
                                        f"Cue_{safe_filename(client_name)}.pdf", 
                                        _pdf_ragic, 
                                        'application/pdf'
                                    )
                                if style_excel_file and RAGIC_MAP.get('file_style_xls'):
                                    _name_raw = style_excel_file.name or f"Style_{safe_filename(client_name)}.xlsx"
                                    _name = safe_filename(_name_raw)
                                    _mime = mimetypes.guess_type(_name)[0] or "application/octet-stream"
                                    files_payload[RAGIC_MAP['file_style_xls']] = (
                                        _name,
                                        style_excel_file.getvalue(),
                                        _mime,
                                    )

                                success, msg, rid = upload_to_ragic(
                                    st.session_state.ragic_url,
                                    st.session_state.ragic_key,
                                    data_payload,
                                    files_payload
                                )
                                
                                if success:
                                    # 成功：將訊息存入 Session，然後重新整理頁面
                                    st.session_state['upload_success_msg'] = msg
                                    st.session_state.ragic_confirm_state = False
                                    st.rerun()
                                else:
                                    # 失敗：直接顯示紅字錯誤
                                    st.error(f"上傳失敗: {msg}")
                            
                            st.session_state.ragic_confirm_state = False
                            time.sleep(1)
                            st.rerun()

    except Exception as e:
        st.error("程式執行發生錯誤，請聯絡開發者。")
        st.error(traceback.format_exc())

if __name__ == "__main__":
    main()
