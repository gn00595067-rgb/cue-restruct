"""
排程運算核心模組 (Calculation Engine)
負責預算分配與檔次計算邏輯
"""

import math
import streamlit as st
from config import REFERENCE_STD_SPOTS
from utils import get_sec_factor, calculate_schedule


def render_logic_panel(logs, use_list_price_for_spots=False, rebate_logs=None, bonus_rebate_logs=None):
    """
    在 UI 上繪製詳細運算邏輯面板 (透明化運算過程)。
    顯示每個項目的：預算、單價、係數、最終檔次。
    use_list_price_for_spots: 若 True（交換合約），面板顯示為「定價」計算；否則為「實作價」。
    rebate_logs: 回饋檔次計算明細列表（由 compute_rebate_rows 產出），若有則顯示「回饋檔次計算」區塊。
    bonus_rebate_logs: 主管額外回饋計算明細（由 compute_bonus_rebate_rows_from_allocation 產出），若有則顯示「主管額外回饋計算」區塊。
    """
    if not logs and not (rebate_logs or []) and not (bonus_rebate_logs or []):
        return

    st.markdown("### 🧮 運算邏輯詳細面板 (透明化運算)")
    if use_list_price_for_spots:
        st.caption("📌 **交換合約**：檔次依**定價 (List Price)** 計算，不提供優惠回饋。")

    price_label = "定價 (List Price)" if use_list_price_for_spots else "實作價 (Net Price)"
    price_latex = r"\text{List Price}" if use_list_price_for_spots else r"\text{Net Price}"

    for idx, item in enumerate(logs or []):
        title = f"#{idx+1} 【{item['media']}】 {item['seconds']}秒 - {item['region']}"
        with st.expander(title, expanded=False):
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("分配預算 (Budget)", f"${int(item['budget']):,}")
            c2.metric("單檔成本 (Unit Cost)", f"${item['unit_cost_actual']:.2f}")
            c3.metric("秒數係數 (Factor)", f"{item['factor']}")
            c4.metric("最終檔次 (Spots)", item['spots'])

            st.markdown("---")
            st.markdown("#### 1. 基礎參數")
            st.text(f"• 媒體與區域: {item['media']} ({item['region']})")
            st.text(f"• {price_label}: ${item['base_net_price']:,} (依據 Pricing 表)")
            st.text(f"• 計算用標準檔次 (Std Spots): {item['std_spots']} 檔")
            if item.get("std_spots_note"):
                st.caption(f"📌 {item['std_spots_note']}")
            st.text(f"• 秒數: {item['seconds']}秒 (Factor: {item['factor']})")

            st.markdown("#### 2. 單檔成本計算")
            st.latex(r"\text{Unit Cost} = \frac{" + price_latex + r"}{\text{Std Spots}} \times \text{Factor}")
            st.code(f"{item['base_net_price']} / {item['std_spots']} * {item['factor']} = {item['unit_cost_actual']:.4f}")

            st.markdown("#### 3. 檔次計算")
            st.text(f"• 初估檔次 = 預算 / 單檔成本 = {item['budget']:.0f} / {item['unit_cost_actual']:.2f} = {item['spots_init_raw']:.2f}")
            st.latex(r"\text{Final Spots} = \text{Ceil}\left(\frac{\text{Budget}}{\text{Unit Cost}}\right)")
            st.code(f"{item['budget']:.0f} / {item['unit_cost_actual']:.2f} = {item.get('spots_final_raw_penalty', item['spots_init_raw']):.2f} -> 無條件進位 -> {item['spots']}")

            if item.get('list_price') is not None and item.get('pkg_display') is not None:
                st.markdown("#### 4. rate (Net) / Package-cost (Net) 計算")
                st.text("報表上 rate (Net) = int(定價 ÷ 標準檔次 × 係數) × 檔次；Package-cost (Net) 為該列或合併顯示之金額。")
                list_p = item['list_price']
                std_s = item['std_spots']
                fac = item['factor']
                sp = item['spots']
                unit_rate = int((list_p / std_s) * fac)
                total_rate = unit_rate * sp
                st.latex(r"\text{rate (Net)} = \left\lfloor \frac{\text{定價}}{\text{標準檔次}} \times \text{係數} \right\rfloor \times \text{檔次}")
                st.code(f"int({list_p:,} / {std_s} × {fac}) × {sp} = {unit_rate:,} × {sp} = ${total_rate:,}")
                st.text(f"• 定價 (List)：${list_p:,}")
                st.text(f"• Package-cost (Net) 顯示：${item['pkg_display']:,}")
                if item.get('rate_pkg_note'):
                    st.caption(item['rate_pkg_note'])

            if item.get("region_detail"):
                st.markdown("#### 5. 各區價錢計算（分區 Std_Spots 影響各區 rate (Net)）")
                st.caption("全省聯播時，各列 rate (Net) 依該區定價與該區 Std_Spots 計算；分區 Std_Spots 與全省可能不同。")
                for rd in item["region_detail"]:
                    st.text(f"• **{rd['region']}**：定價 ${rd['list_price']:,}、Std_Spots = {rd['std_spots']} → int({rd['list_price']:,} / {rd['std_spots']} × {item['factor']}) × {item['spots']} = {rd['unit_rate']:,} × {item['spots']} = **${rd['rate_display']:,}**")

            if item.get('note'):
                st.info(f"備註: {item['note']}")

    # 回饋檔次計算區塊
    if rebate_logs:
        st.markdown("---")
        st.markdown("#### 📌 回饋檔次計算")
        for idx, item in enumerate(rebate_logs):
            title = f"回饋 #{idx+1} 【{item.get('media','')}】 {item.get('region','')} - {item.get('rebate_type','')}"
            with st.expander(title, expanded=False):
                c1, c2, c3 = st.columns(3)
                c1.metric("回饋%", f"{item.get('rebate_pct', 0)}%")
                c2.metric("回饋檔次", f"{item.get('rebate_spots', 0)} 檔")
                if item.get("base_spots") is not None:
                    c3.metric("基準檔次", f"{item['base_spots']} 檔")
                elif item.get("rebate_budget") is not None:
                    c3.metric("回饋預算", f"${int(item['rebate_budget']):,}")
                else:
                    c3.metric("", "")
                st.markdown("**計算式**")
                st.text(item.get("formula", ""))

    # 主管額外回饋計算區塊
    if bonus_rebate_logs:
        st.markdown("---")
        st.markdown("#### 📌 主管額外回饋計算")
        for idx, item in enumerate(bonus_rebate_logs):
            title = f"額外回饋 #{idx+1} 【{item.get('media','')}】 {item.get('seconds','')}秒 - {item.get('region','')}"
            with st.expander(title, expanded=False):
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("主管回饋金額", f"${int(item.get('rebate_budget_total', 0)):,}")
                c2.metric("平台分配預算", f"${int(item.get('platform_budget', 0)):,}")
                c3.metric("該項預算", f"${int(item.get('region_budget', item.get('sec_budget', 0))):,}")
                c4.metric("檔次", f"{item.get('spots', 0)} 檔")
                st.markdown("**明細**")
                st.text(f"• 平台比重：{item.get('platform_share_pct', 0)}% → ${int(item.get('platform_budget', 0)):,}")
                st.text(f"• 秒數比重：{item.get('sec_share_pct', 0)}% → ${int(item.get('sec_budget', 0)):,}")
                if item.get("region_budget") is not None:
                    st.text(f"• 該區分配（{item.get('region', '')}）：${int(item['region_budget']):,}")
                st.text(f"• 單檔成本(實作價)：${item.get('unit_cost_net', 0):,.2f}")
                st.markdown("**計算式**")
                st.text(item.get("formula", ""))


def _get_std_spots(db, region_key, is_national, media):
    """依 Pricing 表取得該區 Std_Spots；缺值時以 config.REFERENCE_STD_SPOTS 為 fallback（以 Google Sheet 為主）。"""
    region_std = db.get("_Region_Std_Spots") or {}
    ref = REFERENCE_STD_SPOTS.get(media, {})
    if is_national or region_key == "全省":
        return region_std.get("全省") or db.get("Std_Spots") or ref.get("全省") or 480
    v = region_std.get(region_key)
    if v is not None:
        return v
    # 以參考表為 fallback；僅當 media 不在參考表時用 480 避免除錯（仍應以 Google Sheet 為主）
    return db.get("Std_Spots") or ref.get(region_key) or ref.get("全省") or 480


def calculate_plan_data(config, total_budget, days_count, pricing_db, sec_factors, store_counts_num, regions_order, use_list_price_for_spots=False):
    """
    排程運算核心函式。

    檔次計算預設依「實作價（Net）」；若 use_list_price_for_spots=True（交換合約），則改依「定價（List）」計算檔次。
    單檔成本 = 價格/標準檔次 × 秒數係數，檔次 = 分配預算 ÷ 單檔成本。
    新鮮視：全省與分區依 Pricing 表各區 Std_Spots 計算，兩者可能不同。
    參數:
        config: 使用者在 Sidebar 選取的設定 (媒體、佔比、秒數)
        total_budget: 總預算
        use_list_price_for_spots: 若 True（交換合約），檔次依定價計算；否則依實作價。
    回傳:
        rows: 計算完的每一列資料 (含 schedule 陣列)
        total_list_accum: 定價總額 (報表分項加總用)
        logs: 運算過程紀錄 (供 render_logic_panel 使用)
    """
    rows, total_list_accum = [], 0
    logs = []
    # 交換合約時用定價 [0] / List，否則用實作價 [1] / Net
    def price_for_spots(db_entry):
        if use_list_price_for_spots:
            if isinstance(db_entry, (list, tuple)):
                return db_entry[0]
            return db_entry.get("List", db_entry.get("Net"))
        if isinstance(db_entry, (list, tuple)):
            return db_entry[1]
        return db_entry.get("Net", db_entry.get("List"))

    for m, cfg in config.items():
        # 1. 根據媒體佔比計算該媒體預算
        m_budget_total = total_budget * (cfg["share"] / 100.0)

        for sec, sec_pct in cfg["sec_shares"].items():
            # 2. 根據秒數佔比計算該秒數預算
            s_budget = m_budget_total * (sec_pct / 100.0)
            if s_budget <= 0:
                continue

            factor = get_sec_factor(m, sec, sec_factors)

            if m in ["全家廣播", "新鮮視"]:
                db = pricing_db[m]
                calc_regs = ["全省"] if cfg["is_national"] else cfg["regions"]
                display_regs = regions_order if cfg["is_national"] else cfg["regions"]
                std_spots_ref = db["Std_Spots"]
                region_shares = cfg.get("region_shares")  # 自訂區域比例 { "北區": 40, "桃竹苗": 30, ... }

                # ---------- 有啟用「自訂區域比例」時的計價邏輯 ----------
                if region_shares and len(region_shares) > 0:
                    # 全省：每一區都有 min_ratio 用全省計價，故全省總預算 = 6 × min_ratio × s_budget（總和才會 100%）
                    if cfg["is_national"]:
                        min_ratio = min(region_shares.values())
                        n_regions = len(region_shares)  # 6
                        budget_nat = (n_regions * min_ratio / 100.0) * s_budget
                        calc_std_spots_nat = _get_std_spots(db, "全省", True, m)
                        price_nat = price_for_spots(db["全省"])
                        unit_net_nat = (price_nat / calc_std_spots_nat) * factor
                        if budget_nat > 0 and unit_net_nat > 0:
                            spots_nat_raw = budget_nat / unit_net_nat
                            spots_nat_init = math.ceil(spots_nat_raw)
                            spots_nat = spots_nat_init
                            if spots_nat % 2 != 0:
                                spots_nat += 1
                            if spots_nat == 0:
                                spots_nat = 2
                            nat_list = db["全省"][0]
                            nat_unit_price = int((nat_list / calc_std_spots_nat) * factor)
                            nat_pkg_display = nat_unit_price * spots_nat
                            total_list_accum += nat_pkg_display
                            sch_nat = calculate_schedule(spots_nat, days_count)
                            _log = {
                                "media": m,
                                "region": "全省聯播(自訂比例-最低比例)",
                                "seconds": sec,
                                "budget": budget_nat,
                                "base_net_price": price_nat,
                                "std_spots": calc_std_spots_nat,
                                "factor": factor,
                                "unit_cost_actual": unit_net_nat,
                                "spots_init_raw": spots_nat_raw,
                                "is_under_target": False,
                                "spots_final_raw_penalty": spots_nat_raw,
                                "spots": spots_nat,
                                "note": "自訂區域比例：全省部分為 6 區皆有的最低比例，總預算 = 6×min_ratio×s_budget",
                                "list_price": nat_list,
                                "pkg_display": nat_pkg_display,
                            }
                            if m == "新鮮視":
                                _log["std_spots_note"] = f"新鮮視：全省 Std_Spots = {calc_std_spots_nat}（依據 Pricing 表）"
                            _region_detail = []
                            for r in display_regs:
                                list_price_region = db[r][0]
                                std_r = _get_std_spots(db, r, False, m)
                                unit_rate_region = int((list_price_region / std_r) * factor)
                                total_rate_display_region = unit_rate_region * spots_nat
                                _region_detail.append({"region": r, "list_price": list_price_region, "std_spots": std_r, "unit_rate": unit_rate_region, "rate_display": total_rate_display_region})
                                rows.append({
                                    "media": m,
                                    "region": r,
                                    "program_num": store_counts_num.get(f"新鮮視_{r}" if m == "新鮮視" else r, 0),
                                    "daypart": db["Day_Part"],
                                    "seconds": sec,
                                    "spots": spots_nat,
                                    "schedule": sch_nat,
                                    "rate_display": total_rate_display_region,
                                    "pkg_display": nat_pkg_display,
                                    "is_pkg_member": True,
                                    "nat_pkg_display": nat_pkg_display
                                })
                            _log["region_detail"] = _region_detail
                            _log["rate_pkg_note"] = "rate (Net) 各列依該區定價與該區 Std_Spots 計算；Package-cost (Net) 為全省合併顯示。"
                            logs.append(_log)

                        # 各區「超出最低比例」部分：依個別縣市價錢計價（新鮮視各區 Std_Spots 依 Pricing 表）
                        for r, ratio in region_shares.items():
                            if ratio <= min_ratio or r not in db or r == "全省":
                                continue
                            extra_budget = ((ratio - min_ratio) / 100.0) * s_budget
                            if extra_budget <= 0:
                                continue
                            calc_std_spots_region = _get_std_spots(db, r, False, m)
                            price_r = price_for_spots(db[r])
                            unit_net_r = (price_r / calc_std_spots_region) * factor
                            if unit_net_r <= 0:
                                continue
                            spots_r_raw = extra_budget / unit_net_r
                            spots_r_init = math.ceil(spots_r_raw)
                            spots_r = spots_r_init
                            if spots_r % 2 != 0:
                                spots_r += 1
                            if spots_r == 0:
                                spots_r = 2
                            list_r = db[r][0]
                            unit_rate_r = int((list_r / calc_std_spots_region) * factor)
                            row_pkg_r = unit_rate_r * spots_r
                            total_list_accum += row_pkg_r
                            sch_r = calculate_schedule(spots_r, days_count)
                            _log = {
                                "media": m,
                                "region": f"{r}(自訂比例加重)",
                                "seconds": sec,
                                "budget": extra_budget,
                                "base_net_price": price_r,
                                "std_spots": calc_std_spots_region,
                                "factor": factor,
                                "unit_cost_actual": unit_net_r,
                                "spots_init_raw": spots_r_raw,
                                "is_under_target": False,
                                "spots_final_raw_penalty": spots_r_raw,
                                "spots": spots_r,
                                "note": f"自訂區域比例：{r} 超出最低比例部分依區計價",
                                "list_price": list_r,
                                "pkg_display": row_pkg_r,
                            }
                            if m == "新鮮視":
                                _log["std_spots_note"] = f"新鮮視：分區({r}) Std_Spots = {calc_std_spots_region}（依據 Pricing 表）"
                            logs.append(_log)
                            rows.append({
                                "media": m,
                                "region": r,
                                "program_num": store_counts_num.get(f"新鮮視_{r}" if m == "新鮮視" else r, 0),
                                "daypart": db["Day_Part"],
                                "seconds": sec,
                                "spots": spots_r,
                                "schedule": sch_r,
                                "rate_display": row_pkg_r,
                                "pkg_display": row_pkg_r,
                                "is_pkg_member": False,
                                "nat_pkg_display": 0
                            })

                    else:
                        # 區域：依自訂比例分配預算，各區個別計價（新鮮視各區 Std_Spots 依 Pricing 表）
                        total_ratio = sum(region_shares.get(r, 0) for r in calc_regs)
                        if total_ratio <= 0:
                            total_ratio = 100.0
                        for r in calc_regs:
                            pct = region_shares.get(r, 0) / total_ratio
                            budget_r = s_budget * pct
                            if budget_r <= 0:
                                continue
                            calc_std_spots = _get_std_spots(db, r, False, m)
                            price_r = price_for_spots(db[r])
                            unit_net_r = (price_r / calc_std_spots) * factor
                            if unit_net_r <= 0:
                                continue
                            spots_r_raw = budget_r / unit_net_r
                            spots_r_init = math.ceil(spots_r_raw)
                            spots_r = spots_r_init
                            if spots_r % 2 != 0:
                                spots_r += 1
                            if spots_r == 0:
                                spots_r = 2
                            list_r = db[r][0]
                            unit_rate_r = int((list_r / calc_std_spots) * factor)
                            row_pkg_r = unit_rate_r * spots_r
                            total_list_accum += row_pkg_r
                            sch_r = calculate_schedule(spots_r, days_count)
                            _log = {
                                "media": m,
                                "region": r,
                                "seconds": sec,
                                "budget": budget_r,
                                "base_net_price": price_r,
                                "std_spots": calc_std_spots,
                                "factor": factor,
                                "unit_cost_actual": unit_net_r,
                                "spots_init_raw": spots_r_raw,
                                "is_under_target": False,
                                "spots_final_raw_penalty": spots_r_raw,
                                "spots": spots_r,
                                "note": "自訂區域比例：依比例分配預算，各區個別計價",
                                "list_price": list_r,
                                "pkg_display": row_pkg_r,
                            }
                            if m == "新鮮視":
                                _log["std_spots_note"] = f"新鮮視：分區({r}) Std_Spots = {calc_std_spots}（依據 Pricing 表）"
                            logs.append(_log)
                            rows.append({
                                "media": m,
                                "region": r,
                                "program_num": store_counts_num.get(f"新鮮視_{r}" if m == "新鮮視" else r, 0),
                                "daypart": db["Day_Part"],
                                "seconds": sec,
                                "spots": spots_r,
                                "schedule": sch_r,
                                "rate_display": row_pkg_r,
                                "pkg_display": row_pkg_r,
                                "is_pkg_member": False,
                                "nat_pkg_display": 0
                            })
                    continue  # 已處理此 (m, sec)，跳過下方原有邏輯
                # ---------- 以下為原本未啟用自訂區域比例的邏輯 ----------

                base_net_price_sum = 0
                is_nat = cfg["is_national"]
                calc_std_spots = _get_std_spots(db, "全省", is_nat, m)
                note_text = "若選全省聯播，實作價為全省定價；若選區域，則為各區實作價加總。"
                std_spots_note = None
                if m == "新鮮視":
                    if is_nat:
                        std_spots_note = f"新鮮視：全省 Std_Spots = {calc_std_spots}（依據 Pricing 表）"
                    else:
                        note_text = "新鮮視分區：各區 Std_Spots 依 Pricing 表，與全省可能不同。"
                        std_spots_note = "新鮮視：分區各區 Std_Spots 依 Pricing 表"

                # 計算該組合的總單檔成本 (Unit Cost)（交換合約時用定價；分區時各區用該區 Std_Spots）
                unit_net_sum = 0
                for r in calc_regs:
                    price_r = price_for_spots(db[r])
                    base_net_price_sum += price_r
                    std_r = _get_std_spots(db, r, is_nat, m)
                    unit_net_sum += (price_r / std_r) * factor

                if unit_net_sum == 0:
                    continue

                # 初算檔次
                spots_init_raw = s_budget / unit_net_sum
                spots_init = math.ceil(spots_init_raw)

                # 最終檔次計算（無懲罰）
                spots_final_raw = s_budget / unit_net_sum
                spots_final = math.ceil(spots_final_raw)

                # 檔次需為偶數 (業務邏輯需求)
                if spots_final % 2 != 0:
                    spots_final += 1
                if spots_final == 0:
                    spots_final = 2

                # 計算顯示用價格 (Rate & Package)，供 log 與後續寫入
                nat_pkg_display = 0
                nat_list = None
                if cfg["is_national"]:
                    nat_list = db["全省"][0]
                    nat_unit_price = int((nat_list / calc_std_spots) * factor)
                    nat_pkg_display = nat_unit_price * spots_final
                    total_list_accum += nat_pkg_display
                first_region_list = db[display_regs[0]][0] if display_regs else None

                log_entry = {
                    "media": m,
                    "region": "全省聯播" if cfg["is_national"] else "/".join(cfg["regions"]),
                    "seconds": sec,
                    "budget": s_budget,
                    "base_net_price": base_net_price_sum,
                    "std_spots": calc_std_spots,
                    "factor": factor,
                    "unit_cost_actual": unit_net_sum,
                    "spots_init_raw": spots_init_raw,
                    "is_under_target": False,
                    "spots_final_raw_penalty": spots_final_raw,
                    "spots": spots_final,
                    "note": note_text
                }
                if std_spots_note:
                    log_entry["std_spots_note"] = std_spots_note
                if cfg["is_national"] and nat_list is not None:
                    log_entry["list_price"] = nat_list
                    log_entry["pkg_display"] = nat_pkg_display
                    log_entry["rate_pkg_note"] = "rate (Net) 各列依該區定價與該區 Std_Spots 計算；Package-cost (Net) 為全省合併顯示。"
                else:
                    std_first = _get_std_spots(db, display_regs[0], False, m) if display_regs else calc_std_spots
                    unit_rate_ex = int((first_region_list / std_first) * factor) if first_region_list else 0
                    log_entry["list_price"] = first_region_list
                    log_entry["pkg_display"] = unit_rate_ex * spots_final if first_region_list else 0
                logs.append(log_entry)

                # 將總檔次分配到每一天
                sch = calculate_schedule(spots_final, days_count)

                # 計算顯示用價格 (Rate & Package) — 全省時 nat_pkg_display 已算過
                if not cfg["is_national"]:
                    nat_pkg_display = 0

                # 全省聯播時仍依「各區 Std_Spots」計算各區 rate (Net)，並寫入 log 供邏輯面板顯示
                region_detail = [] if cfg["is_national"] and display_regs else None
                for i, r in enumerate(display_regs):
                    list_price_region = db[r][0]
                    # 各區 rate (Net) 依該區 Std_Spots 計算（全省聯播也適用，分區 Std_Spots 影響價錢）
                    std_r = _get_std_spots(db, r, False, m)
                    unit_rate_display = int((list_price_region / std_r) * factor)
                    total_rate_display = unit_rate_display * spots_final
                    row_pkg_display = total_rate_display
                    if region_detail is not None:
                        region_detail.append({
                            "region": r,
                            "list_price": list_price_region,
                            "std_spots": std_r,
                            "unit_rate": unit_rate_display,
                            "rate_display": total_rate_display,
                        })
                    if not cfg["is_national"]:
                        total_list_accum += row_pkg_display

                    rows.append({
                        "media": m,
                        "region": r,
                        "program_num": store_counts_num.get(f"新鮮視_{r}" if m == "新鮮視" else r, 0),
                        "daypart": db["Day_Part"],
                        "seconds": sec,
                        "spots": spots_final,
                        "schedule": sch,
                        "rate_display": total_rate_display,
                        "pkg_display": row_pkg_display,
                        "is_pkg_member": cfg["is_national"],
                        "nat_pkg_display": nat_pkg_display
                    })
                if region_detail:
                    logs[-1]["region_detail"] = region_detail

            elif m == "家樂福":
                # 家樂福特殊邏輯: 分為量販與超市，但預算計算主要基於量販
                db = pricing_db["家樂福"]
                base_std = db["量販_全省"]["Std_Spots"]
                price_for_calc = price_for_spots(db["量販_全省"])
                unit_net = (price_for_calc / base_std) * factor

                spots_init_raw = s_budget / unit_net
                spots_init = math.ceil(spots_init_raw)

                spots_final_raw = s_budget / unit_net
                spots_final = math.ceil(spots_final_raw)

                if spots_final % 2 != 0:
                    spots_final += 1

                base_list = db["量販_全省"]["List"]
                unit_rate_h = int((base_list / base_std) * factor)
                total_rate_h = unit_rate_h * spots_final

                logs.append({
                    "media": m,
                    "region": "全省量販+超市",
                    "seconds": sec,
                    "budget": s_budget,
                    "base_net_price": price_for_calc,
                    "std_spots": base_std,
                    "factor": factor,
                    "unit_cost_actual": unit_net,
                    "spots_init_raw": spots_init_raw,
                    "is_under_target": False,
                    "spots_final_raw_penalty": spots_final_raw,
                    "spots": spots_final,
                    "note": f"超市檔次會依照比例自動計算 (量販:{spots_final})",
                    "list_price": base_list,
                    "pkg_display": total_rate_h,
                })

                # 量販列
                sch_h = calculate_schedule(spots_final, days_count)
                total_list_accum += total_rate_h

                rows.append({
                    "media": m,
                    "region": "全省量販",
                    "program_num": store_counts_num["家樂福_量販"],
                    "daypart": db["量販_全省"]["Day_Part"],
                    "seconds": sec,
                    "spots": spots_final,
                    "schedule": sch_h,
                    "rate_display": total_rate_h,
                    "pkg_display": total_rate_h,
                    "is_pkg_member": False
                })

                # 超市列 (依量販檔次比例換算)
                spots_s = int(spots_final * (db["超市_全省"]["Std_Spots"] / base_std))
                sch_s = calculate_schedule(spots_s, days_count)
                rows.append({
                    "media": m,
                    "region": "全省超市",
                    "program_num": store_counts_num["家樂福_超市"],
                    "daypart": db["超市_全省"]["Day_Part"],
                    "seconds": sec,
                    "spots": spots_s,
                    "schedule": sch_s,
                    "rate_display": "計量販",
                    "pkg_display": "計量販",
                    "is_pkg_member": False
                })

    return rows, total_list_accum, logs
