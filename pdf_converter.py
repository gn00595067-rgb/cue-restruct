"""
PDF 轉檔模組 (PDF Converter)
依賴伺服器端的 LibreOffice 進行 Excel -> PDF 轉檔
"""

import os
import shutil
import tempfile
import subprocess
import gc
import requests
import streamlit as st
from config import BOLIN_LOGO_URL


def find_soffice_path():
    """尋找系統中的 LibreOffice 執行檔路徑 (支援 Linux 與 Windows)。"""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        return soffice
    if os.name == "nt":  # Windows fallback
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
        ]
        for p in candidates:
            if os.path.exists(p):
                return p
    return None


@st.cache_data(show_spinner="正在下載 Logo...", ttl=3600)
def get_cloud_logo_bytes():
    """下載鉑霖 Logo 圖片檔 (快取 1 小時)。"""
    try:
        response = requests.get(BOLIN_LOGO_URL, timeout=10)
        return response.content if response.status_code == 200 else None
    except:
        return None


@st.cache_data(show_spinner="正在生成 PDF (LibreOffice)...", ttl=3600)
def xlsx_bytes_to_pdf_bytes(xlsx_bytes: bytes):
    """
    使用 LibreOffice CLI 將 Excel bytes 轉為 PDF bytes。
    過程: 寫入暫存檔 -> 呼叫 soffice 轉檔 -> 讀取 PDF -> 清除暫存。
    """
    soffice = find_soffice_path()
    if not soffice:
        return None, "Fail", "伺服器未安裝 LibreOffice"
    try:
        with tempfile.TemporaryDirectory() as tmp:
            xlsx_path = os.path.join(tmp, "cue.xlsx")
            with open(xlsx_path, "wb") as f:
                f.write(xlsx_bytes)
            # 呼叫 LibreOffice 轉檔指令
            subprocess.run(
                [soffice, "--headless", "--nologo", "--convert-to", "pdf:calc_pdf_Export", "--outdir", tmp, xlsx_path],
                capture_output=True,
                timeout=60
            )
            pdf_path = os.path.join(tmp, "cue.pdf")
            if not os.path.exists(pdf_path):
                # 嘗試尋找任何產出的 pdf
                for fn in os.listdir(tmp):
                    if fn.endswith(".pdf"):
                        pdf_path = os.path.join(tmp, fn)
                        break
            if os.path.exists(pdf_path):
                with open(pdf_path, "rb") as f:
                    return f.read(), "LibreOffice", ""
            return None, "Fail", "LibreOffice 未產出檔案"
    except Exception as e:
        return None, "Fail", str(e)
    finally:
        gc.collect()
