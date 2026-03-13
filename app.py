"""
Cue Sheet Pro (媒體排程生成系統)
重構版本 - 模組化架構

用途: 協助業務生成媒體排程表 (Cue表)，支援 HTML 預覽、Excel/PDF 下載及 Ragic 資料庫串接。
維護注意: 此系統依賴外部 Google Sheet 作為設定檔，以及 LibreOffice 進行 PDF 轉檔。
"""

import streamlit as st
import traceback
import time
from datetime import timedelta, datetime, date

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
    expand_schedule_to_calendar
)
from data_loader import load_config_from_cloud
from calculator import calculate_plan_data, render_logic_panel
from rebate import compute_rebate_rows, merge_rebate_into_rows, get_rebate_summary_text, compute_bonus_rebate_rows, compute_bonus_rebate_rows_from_allocation, get_rebate_qualified_platforms, get_rebate_qualification_detail
from html_generator import generate_html_preview
from excel_renderer import generate_excel_from_scratch
from pdf_converter import xlsx_bytes_to_pdf_bytes
from annual_quarter_cue import build_wave_rows, distribute_by_wave_days, round_to_even
from ragic_api import (
    search_ragic_records,
    upload_to_ragic,
    restore_state_from_ragic
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

for key, default_val in DEFAULT_STATES.items():
    if key not in st.session_state:
        st.session_state[key] = default_val


def _render_annual_quarter_cue(store_counts_num, pricing_db, sec_factors, regions_order, fmt_options, fmt_idx, sales_map):
    """年約／季約細 CUE：以已知檔次與實收分配至各波段，每波段獨立 Excel/PDF。每個廣告組合可個別設定每波檔次與實收；若不知道各波分配則填該組合總檔次／總實收後按「依天數均分」。"""
    st.caption("以已知**執行區域、平台、秒數**設定各波段檔次與實收；若不知各波分配，可填該**廣告組合**的總檔次／總實收後按「依波段天數均分」。")
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
        for i in range(len(waves)):
            w = st.session_state.aq_waves[i]
            c1, c2, c3, c4 = st.columns([1, 2, 2, 1])
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
    rem = get_remarks_text(datetime.now() + timedelta(days=3), "2026年2月", datetime(2026, 3, 31))
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
        p_str = f"{'、'.join([str(r['seconds']) + '秒' for r in rows])} {product_name}"
        html_preview = generate_html_preview(rows, total_days_wave, start_d, end_d, client_name, client_tax_id, p_str, format_type, rem, total_list_wave, budget_wave + int(budget_wave * 0.05), budget_wave, 0)
        with st.expander(f"波段 {i+1} 預覽：{start_d} ~ {end_d}", expanded=False):
            if isinstance(html_preview, list):
                for idx, one_html in enumerate(html_preview):
                    st.caption(f"第 {idx+1} 頁")
                    st.components.v1.html(one_html, height=400, scrolling=True)
            else:
                st.components.v1.html(html_preview, height=400, scrolling=True)
        xlsx_bytes = generate_excel_from_scratch(format_type, start_d, end_d, client_name, client_tax_id, product_name, rows, rem, budget_wave, 0, sales_person, total_list_wave)
        pdf_bytes, _, _ = xlsx_bytes_to_pdf_bytes(xlsx_bytes)
        col_x, col_p = st.columns(2)
        with col_x:
            st.download_button(f"📥 波段{i+1} Excel", xlsx_bytes, f"Cue_{safe_filename(client_name)}_波段{i+1}_{start_d}_{end_d}.xlsx", key=f"aq_xlsx_{i}", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col_p:
            if pdf_bytes:
                st.download_button(f"📥 波段{i+1} PDF", pdf_bytes, f"Cue_{safe_filename(client_name)}_波段{i+1}_{start_d}_{end_d}.pdf", key=f"aq_pdf_{i}", mime="application/pdf")
            else:
                st.caption("PDF 需 LibreOffice")


# =============================================================================
# 主程式邏輯 (Main Execution Block)
# =============================================================================
def main():
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
                        for k in ("supervisor_last_total_budget", "supervisor_override_price"):
                            st.session_state.pop(k, None)
                        st.rerun()
                    else:
                        st.error("密碼錯誤")
            else:
                st.success("✅ 目前狀態：主管模式")
                if st.button("登出"):
                    st.session_state.is_supervisor = False
                    st.rerun()
            # 覆寫欄位「始終」渲染（僅主管時啟用），避免登入時 widget 樹變動觸發 Cached ForwardMsg MISS
            _total = float(st.session_state.get("_total_budget_for_sidebar", st.session_state.get("temp_budget", 1000000)))
            if "supervisor_last_total_budget" not in st.session_state:
                st.session_state.supervisor_last_total_budget = _total
            if _total != st.session_state.supervisor_last_total_budget:
                st.session_state.supervisor_last_total_budget = _total
                _override_display = _total
            else:
                _override_display = float(st.session_state.get("supervisor_override_price", _total))
            st.caption("🔒 專案優惠價覆寫")
            _ov = st.number_input("最終成交價", value=_override_display, step=10000.0, key="supervisor_override_price", label_visibility="collapsed", disabled=not st.session_state.is_supervisor)
            if st.session_state.is_supervisor:
                st.session_state._supervisor_final_budget = _ov
                if _ov != _total:
                    st.caption(f"⚠️ 以 ${_ov:,.0f} 結算")
            
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
                    col_p1.metric("預算", f"${float(sel_rec.get(RAGIC_MAP['budget_raw'], 0)):,.0f}")
                    col_p2.metric("製作費", f"${float(sel_rec.get(RAGIC_MAP['prod_cost'], 0)):,.0f}")
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
        
        format_type = st.radio("選擇格式", fmt_options, index=fmt_idx, horizontal=True)
        is_barter_contract = st.checkbox("是否為交換合約", value=st.session_state.get("is_barter_contract", False), key="is_barter_contract", help="交換合約：檔次依定價計算，不提供優惠回饋檔次。")
        cue_mode = st.radio("製作模式", ["一般CUE", "年約季約細CUE"], index=0 if st.session_state.get("cue_mode", "一般CUE") == "一般CUE" else 1, key="cue_mode", horizontal=True, help="年約季約細CUE：以已知檔次與實收分配至各波段，每波段獨立存檔。")

        if cue_mode == "年約季約細CUE":
            _render_annual_quarter_cue(STORE_COUNTS_NUM, PRICING_DB, SEC_FACTORS, REGIONS_ORDER, fmt_options, fmt_idx, SALES_MAP)
            return

        c1, c2, c3, c4, c5_sales = st.columns(5)
        # 載入資料後的預設值
        def_client = st.session_state.get('temp_client_name', "萬國通路")
        def_tax = st.session_state.get('temp_tax_id', "")
        def_prod = st.session_state.get('temp_product_name', "統一布丁")
        def_budget = float(st.session_state.get('temp_budget', 1000000))
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
        with c3: total_budget_input = st.number_input("總預算 (未稅 Net)", value=def_budget, step=10000.0)
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

        with st.expander("📝 備註欄位設定", expanded=False):
            rc1, rc2, rc3 = st.columns(3)
            # 修改：處理備註欄位的預設值
            def_sign = st.session_state.get('temp_sign_date', datetime.now() + timedelta(days=3))
            def_bill = st.session_state.get('temp_bill_month', "2026年2月")
            def_pay = st.session_state.get('temp_pay_date', datetime(2026, 3, 31))

            sign_deadline = rc1.date_input("回簽截止日", def_sign)
            billing_month = rc2.text_input("請款月份", def_bill)
            payment_date = rc3.date_input("付款兌現日", def_pay)

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

            # 合併回饋與加贈：兩者皆依「原始 rows」的 index；同一 index 先插門檻回饋再插加贈回饋
            all_inserts = []
            bonus_rebate_logs = []
            if apply_rebate and rebate_inserts:
                all_inserts.extend(rebate_inserts)
            if bonus_pct_val is not None and bonus_pct_val > 0 and bonus_config:
                rebate_budget = int(round(total_budget_input * bonus_pct_val / 100.0))
                bonus_inserts, bonus_rebate_logs = compute_bonus_rebate_rows_from_allocation(bonus_config, rebate_budget, active_days, rows, PRICING_DB, SEC_FACTORS, STORE_COUNTS_NUM, REGIONS_ORDER)
                if bonus_inserts:
                    all_inserts.extend(bonus_inserts)
            if all_inserts:
                all_inserts.sort(key=lambda x: (x[0], 1 if x[1].get("is_bonus_rebate") else 0))  # 同 index 時門檻回饋在前
                rows = merge_rebate_into_rows(rows, all_inserts)
            if use_date_segments and segments:
                for row in rows:
                    if row.get("is_rebate") and "schedule" in row:
                        row["schedule"] = expand_schedule_to_calendar(row["schedule"], segments, start_date, end_date)
            prod_cost = prod_cost_input 
            vat = int(round(final_budget_val * 0.05))
            grand_total = final_budget_val + vat
            
            p_str = f"{'、'.join([f'{s}秒' for s in sorted(list(set(r['seconds'] for r in rows)))])} {product_name}"
            rem = get_remarks_text(sign_deadline, billing_month, payment_date)
            
            # 生成 HTML 預覽（表頭與欄數用完整走期 total_days）
            html_preview = generate_html_preview(rows, total_days, start_date, end_date, client_name, client_tax_id, p_str, format_type, rem, total_list_accum, grand_total, final_budget_val, prod_cost)
            if isinstance(html_preview, list):
                for idx, one_html in enumerate(html_preview):
                    st.caption(f"第 {idx+1} 頁")
                    st.components.v1.html(one_html, height=700, scrolling=True)
            else:
                st.components.v1.html(html_preview, height=700, scrolling=True)
            
            # 顯示運算邏輯（含回饋檔次計算、主管額外回饋計算）
            render_logic_panel(logs, use_list_price_for_spots=is_barter_contract, rebate_logs=rebate_logs if not is_barter_contract else [], bonus_rebate_logs=bonus_rebate_logs)
            
            st.markdown("---")
            st.subheader("📥 檔案下載區")
            
            # 生成 Excel (In-Memory)
            xlsx_temp = generate_excel_from_scratch(format_type, start_date, end_date, client_name, client_tax_id, product_name, rows, rem, final_budget_val, prod_cost, sales_person, total_list_accum)
            
            col_dl1, col_dl2, col_ragic = st.columns([1, 1, 2])
            
            # PDF 下載 (透過 LibreOffice)
            with col_dl2:
                pdf_bytes, method, err = xlsx_bytes_to_pdf_bytes(xlsx_temp)
                if pdf_bytes:
                    st.download_button(
                        f"📥 下載 PDF", 
                        pdf_bytes, 
                        f"Cue_{safe_filename(client_name)}.pdf", 
                        key="pdf_dl_btn",
                        mime="application/pdf"
                    )
                else:
                    st.warning(f"PDF 生成失敗: {err}")

            # Excel 下載（所有人皆可下載）
            with col_dl1:
                st.download_button(
                    "📥 下載 Excel",
                    xlsx_temp,
                    f"Cue_{safe_filename(client_name)}.xlsx",
                    key="xlsx_dl_btn",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # Ragic 上傳
            with col_ragic:
                st.markdown("#### ☁️ 上傳至 Ragic")
                
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
                                
                                campaign_summary = format_campaign_details(config)
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

                                files_payload = {}
                                files_payload[RAGIC_MAP['file_xls']] = (
                                    f"Cue_{safe_filename(client_name)}.xlsx", 
                                    xlsx_temp,
                                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                                )
                                
                                if pdf_bytes:
                                    files_payload[RAGIC_MAP['file_pdf']] = (
                                        f"Cue_{safe_filename(client_name)}.pdf", 
                                        pdf_bytes, 
                                        'application/pdf'
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
