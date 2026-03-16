"""
Ragic API 整合模組 (Ragic API Integration)
負責與 Ragic 資料庫進行 CRUD 操作 (搜尋、單筆讀取、新增)
"""

import time
import re
import json
import streamlit as st
import requests
from datetime import datetime, date
from config import RAGIC_FIELD_SERIAL, RAGIC_MAP


def get_ragic_record(api_url, api_key, rid):
    """
    透過 Record ID 讀取單筆 Ragic 資料。
    參數 naming='EID' 非常重要，確保回傳的是 Field ID 而非欄位名稱，
    這樣程式碼才能穩定讀取 (欄位名稱可能會被使用者修改)。
    """
    if not api_url or not api_key:
        return None
    base_url = api_url.split("?")[0].rstrip("/")
    headers = {"Authorization": f"Basic {api_key}"}
    params = {"api": "", "naming": "EID"}
    try:
        resp = requests.get(f"{base_url}/{rid}", headers=headers, params=params, timeout=30)
        if resp.status_code == 200:
            return resp.json()
    except:
        pass
    return None


def upload_to_ragic(api_url, api_key, data_dict, files_dict=None):
    """
    上傳資料至 Ragic。
    包含一個特殊的「回查機制」：上傳成功後，系統會嘗試多次讀取該筆資料，
    以獲取 Ragic 系統自動生成的「流水號 (Cue表編號)」。
    """
    if not api_url or not api_key:
        return False, "API URL 或 API Key 未設定", None
    base_url = api_url.split("?")[0]
    headers = {"Authorization": f"Basic {api_key}"}
    payload = dict(data_dict)
    payload["api"] = ""
    payload["v"] = "3"
    try:
        # 1. 執行 POST 上傳動作
        resp = requests.post(
            base_url, headers=headers, data=payload, files=files_dict, timeout=120
        )
        try:
            j = resp.json()
        except:
            j = None
        if resp.status_code != 200:
            return False, f"HTTP {resp.status_code}: {resp.text[:200]}", None
        if not j:
            return False, f"Ragic 回傳非 JSON 格式: {resp.text[:200]}", None

        # 2. 上傳成功，開始回查流水號
        if j.get("status") == "SUCCESS":
            ragic_id = str(j.get('ragicId'))  # 強制轉字串
            serial_no = ""

            # --- 智能重試與解析機制 ---
            # 因為 Ragic 生成流水號需要一點時間，故重試 3 次
            for i in range(3):
                time.sleep(1.5)  # 等待伺服器運算

                full_rec = get_ragic_record(api_url, api_key, ragic_id)

                if full_rec and isinstance(full_rec, dict):
                    # === 資料結構解析 ===
                    target_data = {}

                    # 情況 A: 資料包在 ID 裡面 (例如 {"28": {"1015306": "..."}})
                    if ragic_id in full_rec:
                        target_data = full_rec[ragic_id]
                    # 情況 B: 資料沒有包 (API 版本差異可能導致)
                    else:
                        target_data = full_rec

                    # 嘗試抓取 Cue 號 (RAGIC_FIELD_SERIAL)
                    val = target_data.get(RAGIC_FIELD_SERIAL)
                    if val:
                        serial_no = val
                        break

            if not serial_no:
                # 若抓不到，為了避免使用者以為失敗，顯示原始 ID
                serial_no = f"(生成中...ID:{ragic_id})"

            success_msg = f"✅ 上傳成功! Ragic流水號: {serial_no}"
            return True, success_msg, ragic_id

        return False, f"❌ Ragic 錯誤 (Code: {j.get('code')}): {j.get('msg')}", None
    except Exception as e:
        return False, f"❌ 連線異常: {str(e)}", None


def search_ragic_records(api_url, api_key, keyword=None, limit=20):
    """
    搜尋 Ragic 資料 (全文檢索)。
    使用 naming='EID' 以利前端解析。
    """
    if not api_url or not api_key:
        return []
    base_url = api_url.split("?")[0]
    headers = {"Authorization": f"Basic {api_key}"}

    params = {"limit": limit, "api": "", "naming": "EID"}
    if keyword:
        params["fts"] = keyword

    try:
        resp = requests.get(base_url, headers=headers, params=params, timeout=30)
        if resp.status_code == 200:
            data = resp.json()
            records = []
            if not data:
                return []

            for rid, row in data.items():
                if not isinstance(row, dict):
                    continue
                row['_ragicId'] = rid
                records.append(row)

            # 排序：ID 大的在前面 (新的資料)
            records.sort(key=lambda x: int(x['_ragicId']) if x['_ragicId'].isdigit() else 0, reverse=True)
            return records
    except Exception as e:
        st.error(f"搜尋失敗: {e}")
    return []


def _ragic_number(val, default=0):
    """Ragic 可能回傳數字或字串（含逗號），轉成 float。"""
    if val is None or val == "":
        return float(default)
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(",", "")
    try:
        return float(s)
    except ValueError:
        return float(default)


def restore_state_from_ragic(record):
    """
    從 Ragic Record 解析資料並還原到 st.session_state。
    這是「載入舊專案設定」的核心功能。
    需解析儲存在 Ragic 備註欄位中的字串 (details) 來還原 checkbox 與 slider 狀態。
    """
    try:
        # 1. 基礎欄位還原 (使用 RAGIC_MAP 統一管理欄位 ID)
        st.session_state['temp_client_name'] = record.get(RAGIC_MAP['client'], '') or ''
        st.session_state['temp_product_name'] = record.get(RAGIC_MAP['product'], '') or ''

        st.session_state['temp_budget'] = _ragic_number(record.get(RAGIC_MAP['budget_raw']), 0)
        st.session_state['temp_prod_cost'] = _ragic_number(record.get(RAGIC_MAP['prod_cost']), 0)
        # 讓主畫面「總預算」顯示載入值：同步到 _total_budget_for_sidebar，並清除主管覆寫
        st.session_state['_total_budget_for_sidebar'] = st.session_state['temp_budget']
        st.session_state.pop('_supervisor_final_budget', None)

        st.session_state['temp_sales'] = record.get(RAGIC_MAP['sales'], '')
        st.session_state['temp_tax_id'] = record.get(RAGIC_MAP['tax_id'], '')

        # --- 格式與備註欄位還原 ---
        st.session_state['temp_format_type'] = record.get(RAGIC_MAP['format'], '東吳')  # 還原報表格式
        st.session_state['temp_bill_month'] = record.get(RAGIC_MAP['bill_month'], '')    # 還原請款月份

        # 日期解析 helper
        def parse_date(d_str):
            if not d_str:
                return datetime.now()
            for fmt in ['%Y/%m/%d', '%Y-%m-%d']:
                try:
                    return datetime.strptime(d_str.split(' ')[0], fmt)
                except:
                    pass
            return datetime.now()

        st.session_state['temp_start_date'] = parse_date(record.get(RAGIC_MAP['date_start'], ''))
        st.session_state['temp_end_date'] = parse_date(record.get(RAGIC_MAP['date_end'], ''))

        st.session_state['temp_sign_date'] = parse_date(record.get(RAGIC_MAP['date_sign'], ''))
        st.session_state['temp_pay_date'] = parse_date(record.get(RAGIC_MAP['date_pay'], ''))

        # 2. 複雜設定還原 (解析 Details 字串)
        # 需解析如: "【全家廣播】 預算佔比: 50% | 秒數分配: 20秒(100%) | 區域: 全省聯播"
        details = record.get(RAGIC_MAP['details'], '')

        # 先重置開關
        st.session_state['cb_rad'] = False
        st.session_state['cb_fv'] = False
        st.session_state['cb_cf'] = False

        lines = details.split('\n')
        for line in lines:
            line = line.strip()
            if not line:
                continue

            # 使用 Regex 提取資料
            match = re.search(r"【(.*?)】 預算佔比: (\d+)% \| 秒數分配: (.*?) \| 區域: (.*)", line)
            if match:
                media, share, sec_str, reg_str = match.groups()
                share = int(share)

                if media == "全家廣播":
                    st.session_state['cb_rad'] = True
                    st.session_state['rad_share'] = share
                    st.session_state['rad_nat'] = (reg_str == "全省聯播")
                    if not st.session_state['rad_nat']:
                        st.session_state['rad_reg'] = reg_str.split('/')

                    secs_matches = re.findall(r"(\d+)秒\((\d+)%\)", sec_str)
                    selected_secs = []
                    for s_val, s_pct in secs_matches:
                        s_int = int(s_val)
                        selected_secs.append(s_int)
                        st.session_state[f"rs_{s_int}"] = int(s_pct)
                    st.session_state['rad_sec'] = selected_secs

                elif media == "新鮮視":
                    st.session_state['cb_fv'] = True
                    st.session_state['fv_share'] = share
                    st.session_state['fv_nat'] = (reg_str == "全省聯播")
                    if not st.session_state['fv_nat']:
                        st.session_state['fv_reg'] = reg_str.split('/')

                    secs_matches = re.findall(r"(\d+)秒\((\d+)%\)", sec_str)
                    selected_secs = []
                    for s_val, s_pct in secs_matches:
                        s_int = int(s_val)
                        selected_secs.append(s_int)
                        st.session_state[f"fs_{s_int}"] = int(s_pct)
                    st.session_state['fv_sec'] = selected_secs

                elif media == "家樂福":
                    st.session_state['cb_cf'] = True
                    st.session_state['cf_share'] = share

                    secs_matches = re.findall(r"(\d+)秒\((\d+)%\)", sec_str)
                    selected_secs = []
                    for s_val, s_pct in secs_matches:
                        s_int = int(s_val)
                        selected_secs.append(s_int)
                        st.session_state[f"cs_{s_int}"] = int(s_pct)
                    st.session_state['cf_sec'] = selected_secs

        # 3. 投放參數詳情擴充 [EXT] JSON：交換／回饋／分波段／備註／業務加贈等
        if "[EXT]" in details:
            try:
                ext_part = details.split("[EXT]", 1)[-1].strip()
                if ext_part:
                    ext_data = json.loads(ext_part)
                    _apply_ext_state(ext_data)
            except (json.JSONDecodeError, TypeError) as e:
                pass  # 舊資料無 [EXT] 或 JSON 損壞時略過，不影響媒體區塊已還原

        return True, "✅ 資料載入成功！請檢查下方設定。"
    except Exception as e:
        return False, f"❌ 解析失敗: {e}"


def _parse_iso_date(s):
    """字串轉 date/datetime，供還原用。"""
    if s is None or s == "":
        return None
    if isinstance(s, (date, datetime)):
        return s.date() if hasattr(s, "date") else s
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            d = datetime.strptime(str(s).split("T")[0].split(" ")[0], fmt)
            return d.date() if hasattr(d, "date") else d
        except ValueError:
            continue
    return None


def _apply_ext_state(ext_data):
    """將 [EXT] JSON 的鍵值寫回 st.session_state，支援 100% 還原。"""
    if not ext_data or not isinstance(ext_data, dict):
        return
    ss = st.session_state
    date_keys = ("cue_sign_deadline", "cue_payment_date")
    date_seg_keys = ("sign_deadline", "payment_date")  # cue_seg_remarks[i] 內

    for k, v in ext_data.items():
        if v is None:
            continue
        if k == "date_segments" and isinstance(v, list):
            segs = []
            for pair in v:
                if isinstance(pair, (list, tuple)) and len(pair) >= 2:
                    s, e = _parse_iso_date(pair[0]), _parse_iso_date(pair[1])
                    if s and e:
                        segs.append((s, e))
            if segs:
                ss["date_segments"] = segs
                ss["use_date_segments"] = True
            continue
        if k == "cue_seg_remarks" and isinstance(v, dict):
            seg_rem = {}
            for ki, data in v.items():
                if not isinstance(data, dict):
                    continue
                idx = int(ki) if str(ki).isdigit() else ki
                seg_rem[idx] = {}
                for kk, vv in data.items():
                    if kk in date_seg_keys and vv:
                        seg_rem[idx][kk] = _parse_iso_date(vv)
                    else:
                        seg_rem[idx][kk] = vv
            ss["cue_seg_remarks"] = seg_rem
            continue
        if k == "custom_bonus" and isinstance(v, list):
            for item in v:
                if not isinstance(item, dict) or not item.get("enabled"):
                    continue
                key_list = item.get("key")
                if isinstance(key_list, list) and len(key_list) >= 3:
                    media, sec, region_key = key_list[0], key_list[1], key_list[2]
                    skey = f"cb_{media}_{sec}_{region_key}".replace(" ", "_")
                    ss[f"cb_en_{skey}"] = True
                    ss[f"cb_spots_{skey}"] = int(item.get("spots_per_day", 0) or 0)
                    start_d = _parse_iso_date(item.get("date_start"))
                    end_d = _parse_iso_date(item.get("date_end"))
                    if start_d:
                        ss[f"cb_start_{skey}"] = start_d
                    if end_d:
                        ss[f"cb_end_{skey}"] = end_d
            continue
        if k in date_keys and v:
            parsed = _parse_iso_date(v)
            if parsed:
                ss[k] = parsed
            continue
        # 其餘鍵直接寫入（bool, int, str, list）
        try:
            ss[k] = v
        except Exception:
            pass
