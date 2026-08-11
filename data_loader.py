"""
資料讀取模組 (Data Loader)
負責從 Google Sheet 讀取設定檔資料
"""

import re
import pandas as pd
import streamlit as st
from config import REFERENCE_STD_SPOTS


@st.cache_data(ttl=300)
def load_config_from_cloud(share_url):
    """
    從 Google Spreadsheet 讀取所有基礎設定 (Store Count, Pricing, Factors, Sales)。
    使用 gviz API 將 Sheet 轉為 CSV 讀取。
    """
    try:
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", share_url)
        if not match:
            return None, None, None, None, None, "連結格式錯誤"
        file_id = match.group(1)

        def read_sheet(sheet_name):
            url = f"https://docs.google.com/spreadsheets/d/{file_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
            return pd.read_csv(url)

        df_store = read_sheet("Stores")
        df_store.columns = [c.strip() for c in df_store.columns]
        store_counts = dict(zip(df_store['Key'], df_store['Display_Name']))
        store_counts_num = dict(zip(df_store['Key'], df_store['Count']))

        df_fact = read_sheet("Factors")
        df_fact.columns = [c.strip() for c in df_fact.columns]
        sec_factors = {}
        for _, row in df_fact.iterrows():
            if row['Media'] not in sec_factors:
                sec_factors[row['Media']] = {}
            sec_factors[row['Media']][int(row['Seconds'])] = float(row['Factor'])

        name_map = {"全家新鮮視": "新鮮視", "全家廣播": "全家廣播", "家樂福": "家樂福"}
        for k, v in name_map.items():
            if k in sec_factors and v not in sec_factors:
                sec_factors[v] = sec_factors[k]

        df_price = read_sheet("Pricing")
        df_price.columns = [c.strip() for c in df_price.columns]
        pricing_db = {}
        for _, row in df_price.iterrows():
            m, r = row['Media'], row['Region']
            if m == "家樂福":
                if m not in pricing_db:
                    pricing_db[m] = {}
                pricing_db[m][r] = {
                    "List": int(row['List_Price']),
                    "Net": int(row['Net_Price']),
                    "Std_Spots": int(row['Std_Spots']),
                    "Day_Part": row['Day_Part']
                }
            else:
                if m not in pricing_db:
                    pricing_db[m] = {"Day_Part": row['Day_Part'], "_Region_Std_Spots": {}}
                pricing_db[m]["_Region_Std_Spots"][r] = int(row['Std_Spots'])
                pricing_db[m][r] = [int(row['List_Price']), int(row['Net_Price'])]
        # 全家廣播/新鮮視：Std_Spots 預設為全省值，供未改寫處相容；缺值時以 REFERENCE_STD_SPOTS 為 fallback
        for m in list(pricing_db.keys()):
            if m != "家樂福" and "_Region_Std_Spots" in pricing_db[m]:
                rs = pricing_db[m]["_Region_Std_Spots"]
                fallback = REFERENCE_STD_SPOTS.get(m, {}).get("全省", 480)
                pricing_db[m]["Std_Spots"] = rs.get("全省", next(iter(rs.values()), fallback))

        df_sales = read_sheet("Sales")
        df_sales.columns = [c.strip() for c in df_sales.columns]
        if 'Name' in df_sales.columns and 'Nickname' in df_sales.columns:
            sales_map = dict(zip(df_sales['Name'], df_sales['Nickname']))
        else:
            sales_map = {name: name for name in df_sales.iloc[:, 0].tolist()}

        return store_counts, store_counts_num, pricing_db, sec_factors, sales_map, None
    except Exception as e:
        return None, None, None, None, None, f"讀取失敗: {str(e)}"


@st.cache_data(ttl=300)
def load_agency_pricing_from_cloud(share_url):
    """
    從 Google Sheet 的 `AgencyPricing` 分頁讀取代理商牌價（定價）。
    欄位：Agency | Platform | Seconds | List_Per_Spot。
    回傳 {(agency, platform): {seconds: per_spot}}；讀不到或欄位不齊回傳 None（改用 fallback）。
    """
    try:
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", share_url)
        if not match:
            return None
        file_id = match.group(1)
        url = f"https://docs.google.com/spreadsheets/d/{file_id}/gviz/tq?tqx=out:csv&sheet=AgencyPricing"
        df = pd.read_csv(url)
        df.columns = [c.strip() for c in df.columns]
        need = {"Agency", "Platform", "Seconds", "List_Per_Spot"}
        if not need.issubset(set(df.columns)):
            return None
        out = {}
        for _, row in df.iterrows():
            try:
                agency = str(row["Agency"]).strip()
                platform = str(row["Platform"]).strip()
                sec = int(row["Seconds"])
                per = int(float(str(row["List_Per_Spot"]).replace(",", "")))
            except (ValueError, TypeError):
                continue
            out.setdefault((agency, platform), {})[sec] = per
        return out or None
    except Exception:
        return None
