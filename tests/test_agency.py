# -*- coding: utf-8 -*-
"""
代理商 CUE 驗收測試（不依賴 streamlit runtime）。
數字全部出自實際範例檔與轉換說明書，須與開發規格 §4 完全一致。

執行：  py -m pytest tests/test_agency.py -q
或：    py tests/test_agency.py
"""
import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import agency_cue as ac


def _sheet(model, platform):
    for s in model["sheets"]:
        if s["platform"] == platform:
            return s
    raise AssertionError(f"找不到平台 {platform}")


def _row(sheet, kind):
    for r in sheet["rows"]:
        if r["kind"] == kind:
            return r
    raise AssertionError(f"找不到 kind={kind}")


def _fam(enabled=True, seconds=15, share=100, rebate_pct=0, spots_override=0,
         auto_rebate=False):
    # auto_rebate 預設 False：既有數值測試僅驗手動回饋/補償，自帶回饋另有專測。
    return {"enabled": enabled, "seconds": seconds, "share": share,
            "rebate_pct": rebate_pct, "spots_override": spots_override,
            "auto_rebate": auto_rebate}


def _wjf(enabled=True, seconds=20, share=100, rebate_pct=0, mag_override=0, is_rebate_wave=False):
    return {"enabled": enabled, "seconds": seconds, "share": share,
            "rebate_pct": rebate_pct, "mag_override": mag_override,
            "is_rebate_wave": is_rebate_wave}


# ---------------------------------------------------------------------------
# 單元：實作價 / 牌價 / 分配
# ---------------------------------------------------------------------------
def test_unit_prices():
    assert round(ac.family_unit_net(30), 2) == 520.83
    assert round(ac.family_unit_net(15), 2) == 260.42
    assert round(ac.family_unit_net(10), 2) == 173.61
    assert round(ac.wjf_unit_net(20), 2) == 595.24
    assert round(ac.wjf_unit_net(15), 2) == 446.43
    assert ac.super_spots_from_mag(210) == 360
    assert ac.super_spots_from_mag(560) == 960
    assert ac.super_spots_from_mag(413) == 708


def test_list_prices():
    assert ac.get_agency_list_price("2008傳媒", "全家企頻", 15) == 1950
    assert ac.get_agency_list_price("佳聖", "全家企頻", 15) == 1300
    assert ac.get_agency_list_price("佳聖", "萬家福", 20) == 4000
    assert ac.get_agency_list_price("佳聖", "萬家福", 15) == 3400
    assert ac.get_agency_list_price("凱絡", "全家企頻", 15) == 1500
    assert ac.get_agency_list_price("凱絡", "萬家福", 15) == 3400


def test_distributions():
    # 2008 全家 960/28 → 34×24 + 36×4（36 在最後）
    s = ac.dist_even_end(960, 28)
    assert sum(s) == 960
    assert s.count(34) == 24 and s[-4:] == [36, 36, 36, 36]
    # 2008 萬家福 413/21 → 19×7 + 20×14（餘數放最後）
    s = ac.dist_plain(413, 21, "end")
    assert sum(s) == 413 and s.count(20) == 14 and s.count(19) == 7 and s[0] == 19
    # 佳聖 1440/14 → 103×12 + 102×2（餘數放最前）
    s = ac.dist_plain(1440, 14, "front")
    assert sum(s) == 1440 and s[:12] == [103] * 12 and s[12:] == [102, 102]
    # 佳聖 576/14 → 42,42,41×12
    s = ac.dist_plain(576, 14, "front")
    assert sum(s) == 576 and s[:2] == [42, 42] and s[2:] == [41] * 12
    # 佳聖 210/20 → 11×10 + 10×10
    s = ac.dist_plain(210, 20, "front")
    assert sum(s) == 210 and s[:10] == [11] * 10 and s[10:] == [10] * 10
    # 凱絡 量販 560/16 → 36×14 + 28 + 28
    s = ac.dist_carat(560, 16)
    assert sum(s) == 560 and s[:14] == [36] * 14 and s[14:] == [28, 28]
    # 凱絡 超市 960/16（量販36）→ 62×14 + 46 + 46
    mag = ac.dist_carat(560, 16)
    ss = ac.dist_carat_super(960, mag)
    assert sum(ss) == 960 and ss[:14] == [62] * 14 and ss[14:] == [46, 46]


# ---------------------------------------------------------------------------
# A. 2008／全家
# ---------------------------------------------------------------------------
def test_A_2008_family():
    m = ac.build_agency_model(
        "2008傳媒", "統一企業", "統一木瓜牛乳", "統一木瓜牛乳",
        date(2026, 8, 5), date(2026, 9, 1), 250000,
        _fam(seconds=15), None, ac.COMP_MOVE50, date(2026, 7, 24), 3.0,
    )
    sh = _sheet(m, "全家企頻")
    main = _row(sh, ac.KIND_MAIN)
    comp = _row(sh, ac.KIND_COMP)
    assert main["spots"] == 960
    assert comp["spots"] == 480 and comp["schedule"] is None
    assert sum(main["schedule"]) == 960
    assert main["schedule"][:24] == [34] * 24 and main["schedule"][24:] == [36] * 4
    # 定價 = 1950 × (960+480) = 2,808,000
    assert main["list_total"] == 2_808_000
    f = sh["fees"]
    assert f["budget_net"] == 250000
    assert f["ac"] == 7500 and f["tax"] == 12875 and f["total"] == 270375


# ---------------------------------------------------------------------------
# B. 佳聖／全家（回饋 40%、補償併主列、無 AC）
# ---------------------------------------------------------------------------
def test_B_ddrive_family():
    m = ac.build_agency_model(
        "佳聖", "客戶", "產品", "",
        date(2024, 6, 5), date(2024, 6, 18), 250000,
        _fam(seconds=15, rebate_pct=40), None, ac.COMP_MOVE50, date(2024, 5, 29), None,
    )
    sh = _sheet(m, "全家企頻")
    main = _row(sh, ac.KIND_MAIN)
    reb = _row(sh, ac.KIND_REBATE)
    assert main["spots"] == 1440  # 960 + 480 併一列
    assert main["schedule"][:12] == [103] * 12 and main["schedule"][12:] == [102, 102]
    assert reb["spots"] == 576
    assert reb["schedule"][:2] == [42, 42] and reb["schedule"][2:] == [41] * 12
    assert main["list_total"] == 1_872_000
    assert reb["list_total"] == 748_800
    f = sh["fees"]
    assert f["net"] == 250000 and f["vat"] == 12500 and f["gross"] == 262500


# ---------------------------------------------------------------------------
# C. 佳聖／萬家福
# ---------------------------------------------------------------------------
def test_C_ddrive_wjf():
    m = ac.build_agency_model(
        "佳聖", "客戶", "產品", "",
        date(2026, 8, 8), date(2026, 8, 27), 125000,
        None, _wjf(seconds=20), ac.COMP_NONE, date(2026, 8, 1), None,
    )
    sh = _sheet(m, "萬家福")
    mag = _row(sh, ac.KIND_MAIN)
    sup = _row(sh, ac.KIND_SUPER)
    assert mag["spots"] == 210
    assert mag["schedule"][:10] == [11] * 10 and mag["schedule"][10:] == [10] * 10
    assert sup["spots"] == 360 and all(x == 18 for x in sup["schedule"])
    assert mag["list_total"] == 840_000
    f = sh["fees"]
    assert f["net"] == 125000 and f["vat"] == 6250 and f["gross"] == 131250


# ---------------------------------------------------------------------------
# D. 凱絡／萬家福（回饋 10%、A.C 免收）
# ---------------------------------------------------------------------------
def test_D_carat_wjf():
    m = ac.build_agency_model(
        "凱絡", "客戶", "產品", "",
        date(2026, 8, 17), date(2026, 9, 1), 250000,
        None, _wjf(seconds=15, rebate_pct=10), ac.COMP_NONE, date(2026, 8, 10), None,
    )
    sh = _sheet(m, "萬家福")
    mag = _row(sh, ac.KIND_MAIN)
    sup = _row(sh, ac.KIND_SUPER)
    reb = _row(sh, ac.KIND_REBATE)
    reb_s = _row(sh, ac.KIND_SUPER_REBATE)
    assert mag["spots"] == 560 and mag["schedule"][:14] == [36] * 14 and mag["schedule"][14:] == [28, 28]
    assert sup["spots"] == 960 and sup["schedule"][:14] == [62] * 14 and sup["schedule"][14:] == [46, 46]
    # 凱絡回饋改逐日平均鋪滿（不再合併顯示）
    assert reb["spots"] == 56 and reb["schedule"] == [4] * 8 + [3] * 8
    assert reb_s["spots"] == 96 and reb_s["schedule"] == [6] * 16
    assert mag["market_per"] == 2720 and mag["uni_per"] == 2550
    assert mag["uni_total"] == 1_428_000
    assert sup["uni_total"] == 2_448_000
    assert reb["uni_total"] == 142_800
    assert reb_s["uni_total"] == 244_800
    media_value = mag["uni_total"] + sup["uni_total"] + reb["uni_total"] + reb_s["uni_total"]
    assert media_value == 4_263_600
    f = sh["fees"]
    assert f["subtotal"] == 250000 and f["ac_free"] is True
    assert f["vat"] == 12500 and f["grand"] == 262500
    assert media_value - f["subtotal"] == 4_013_600


# ---------------------------------------------------------------------------
# E. 轉換說明書方案（預算 250,000、全家 10秒）
# ---------------------------------------------------------------------------
def test_E1_plan1():
    # 方案一：主 1440 + 補償 216；40%回饋 = round((1440+216)×0.4)=662
    m = ac.build_agency_model(
        "佳聖", "客戶", "產品", "",
        date(2026, 3, 1), date(2026, 3, 14), 250000,
        _fam(seconds=10, rebate_pct=40), None, ac.COMP_PLAN1, date(2026, 2, 22), None,
    )
    sh = _sheet(m, "全家企頻")
    main = _row(sh, ac.KIND_MAIN)
    reb = _row(sh, ac.KIND_REBATE)
    assert main["spots"] == 1440 + 216  # 主+補償併主列
    assert reb["spots"] == 662


def test_E2_plan2():
    # 方案二：全家主 1440 + 回饋 576；萬家福表(15秒) 量販 112、超市 192
    m = ac.build_agency_model(
        "佳聖", "客戶", "產品", "",
        date(2026, 3, 1), date(2026, 3, 14), 250000,
        _fam(seconds=10, rebate_pct=40), _wjf(enabled=False, seconds=15),
        ac.COMP_PLAN2, date(2026, 2, 22), None,
    )
    fam = _sheet(m, "全家企頻")
    main = _row(fam, ac.KIND_MAIN)
    reb = _row(fam, ac.KIND_REBATE)
    assert main["spots"] == 1440  # 方案二全家不加補償列
    assert reb["spots"] == 576    # round(1440×0.4)
    wjf = _sheet(m, "萬家福")
    def _has_comp(r):
        return "凌晨" in (r["media_label"] + r["region_label"])
    comp_mag = [r for r in wjf["rows"] if _has_comp(r) and r["kind"] == ac.KIND_MAIN][0]
    comp_sup = [r for r in wjf["rows"] if _has_comp(r) and r["kind"] == ac.KIND_SUPER][0]
    assert comp_mag["spots"] == 112 and comp_mag["net_display"] == ac.NET_REBATE
    assert comp_sup["spots"] == 192


# ---------------------------------------------------------------------------
# F. 2008／萬家福整波回饋（量販手動 413）
# ---------------------------------------------------------------------------
def test_F_2008_wjf_rebate_wave():
    m = ac.build_agency_model(
        "2008傳媒", "統一企業", "產品", "",
        date(2026, 8, 8), date(2026, 8, 28), 0,
        None, _wjf(seconds=15, mag_override=413, is_rebate_wave=True),
        ac.COMP_NONE, date(2026, 8, 1), 3.0,
    )
    sh = _sheet(m, "萬家福")
    mag = _row(sh, ac.KIND_MAIN)
    sup = _row(sh, ac.KIND_SUPER)
    assert mag["spots"] == 413
    assert sup["spots"] == 708
    assert mag["list_total"] == 805_350  # 1950 × 413
    assert sh["is_rebate_wave"] is True
    assert mag["net_display"] == ac.NET_REBATE
    assert sh["fees"]["total"] == ac.NET_REBATE


# ---------------------------------------------------------------------------
# 自帶專案回饋（凌晨時數轉換）
# ---------------------------------------------------------------------------
def test_offhours_parse_all_formats():
    # 三家時段格式都要解析出 07→23 = 16 時 → off 8
    assert ac.family_offhours("07:00-23:00") == 8   # 2008
    assert ac.family_offhours("07-23") == 8          # 佳聖
    assert ac.family_offhours("0700-2300") == 8      # 凱絡
    assert ac.family_offhours("0700-2300 ") == 8     # 含尾空白
    assert ac.family_offhours("亂七八糟") == 8        # 解析失敗退回預設 8


def test_auto_rebate_spots_formula():
    # 檔 = off(8) × 30 × 30 / 秒
    assert ac.auto_rebate_spots("07:00-23:00", 30) == 240   # 8×30×1
    assert ac.auto_rebate_spots("07:00-23:00", 15) == 480   # 8×30×2  ← 合作夥伴案例
    assert ac.auto_rebate_spots("07:00-23:00", 20) == 360   # 8×30×1.5
    assert ac.auto_rebate_spots("07:00-23:00", 10) == 720   # 8×30×3


def _rows(sheet, kind):
    return [r for r in sheet["rows"] if r["kind"] == kind]


def test_auto_rebate_row_2008():
    # 統一案例：全家 15秒、預算 250000、自帶回饋開 → 480 檔專案回饋列
    m = ac.build_agency_model(
        "2008傳媒", "統一企業", "全家企頻", "",
        date(2026, 8, 5), date(2026, 9, 1), 250000,
        _fam(seconds=15, auto_rebate=True), None,
        ac.COMP_NONE, date(2026, 7, 29), None,
    )
    fam = _sheet(m, "全家企頻")
    main = _row(fam, ac.KIND_MAIN)
    rebs = _rows(fam, ac.KIND_REBATE)
    assert main["spots"] == 960
    assert len(rebs) == 1
    assert rebs[0]["spots"] == 480
    assert rebs[0]["net_display"] == ac.NET_REBATE
    # 自帶回饋不計實收：費用 net 仍等於實作價值（不含回饋列）
    assert fam["fees"]["budget_net"] == ac.rhu(ac.family_unit_net(15) * 960)


def test_auto_rebate_stacks_with_manual_and_comp():
    # 自帶(480) + 手動%(另一列) 兩列並存，三家共用算法（凱絡逐日鋪滿）
    m = ac.build_agency_model(
        "凱絡", "統一企業", "全家企頻", "",
        date(2026, 8, 5), date(2026, 9, 1), 250000,
        _fam(seconds=15, rebate_pct=20, auto_rebate=True), None,
        ac.COMP_PLAN1, date(2026, 7, 29), None,
    )
    fam = _sheet(m, "全家企頻")
    rebs = _rows(fam, ac.KIND_REBATE)
    assert len(rebs) == 2
    assert rebs[0]["spots"] == 480                       # 自帶（先列）
    assert sum(rebs[0]["schedule"]) == 480               # 凱絡逐日鋪滿
    main = _row(fam, ac.KIND_MAIN)
    # 手動 20% × (主+補償)
    assert rebs[1]["spots"] == ac.rhu(0.20 * main["spots"])


def test_auto_rebate_off_by_default_when_disabled():
    m = ac.build_agency_model(
        "2008傳媒", "統一企業", "全家企頻", "",
        date(2026, 8, 5), date(2026, 9, 1), 250000,
        _fam(seconds=15, auto_rebate=False), None,
        ac.COMP_NONE, date(2026, 7, 29), None,
    )
    fam = _sheet(m, "全家企頻")
    assert _rows(fam, ac.KIND_REBATE) == []


if __name__ == "__main__":
    import traceback
    fns = [v for k, v in sorted(globals().items()) if k.startswith("test_") and callable(v)]
    passed = failed = 0
    for fn in fns:
        try:
            fn()
            print(f"PASS  {fn.__name__}")
            passed += 1
        except Exception:
            print(f"FAIL  {fn.__name__}")
            traceback.print_exc()
            failed += 1
    print(f"\n{passed} passed, {failed} failed")
    sys.exit(1 if failed else 0)
