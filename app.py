import streamlit as st
import pandas as pd
import base64
import io
import re
import os
import numpy as np
import matplotlib.pyplot as plt
import urllib.request
import matplotlib.font_manager as fm

st.set_page_config(
    page_title="ระบบ Report ข้อมูลจิตเวช",
    page_icon="🌿",
    layout="wide",
)
plt.switch_backend("Agg")

PROTO = "https" + "://"
SHEET_ID = "1dtpMxycg0en1_zdtsQreeLohqOVEqNyZyxCI26O5Zlc"
SHEET_URL = f"{PROTO}docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"
EDIT_URL = f"{PROTO}docs.google.com/spreadsheets/d/{SHEET_ID}/edit"
FONT_URL = f"{PROTO}github.com/googlefonts/sarabun/raw/main/fonts/ttf/Sarabun-Regular.ttf"

ID_RED = "15P_z1gObqnm29vn-afAJ4JeMRQ4Y-IZw"
ID_ORG = "1Vkl3jyY4W9h3Mv_l17xlWmbNw1A4p4-P"
ID_YEL = "1YlAPW2PBMUbkuRt0unWjJTolQ9aAp48Y"
ID_GRN = "1Wu3vMN2idLhA5fWlY4ZsGZ64Uf_c-f-B"
ID_LOGO = "1KYrHcRg6dvs2h0nfDf7ZxpzWpLnCqnjY"

@st.cache_resource
def download_thai_font():
    font_path = "Sarabun-Regular.ttf"
    if not os.path.exists(font_path):
        try:
            urllib.request.urlretrieve(FONT_URL, font_path)
        except Exception:
            pass
    return font_path

THAI_FONT_PATH = download_thai_font()
try:
    fm.fontManager.addfont(THAI_FONT_PATH)
    THAI_FONT_NAME = fm.FontProperties(fname=THAI_FONT_PATH).get_name()
    plt.rcParams["font.family"] = THAI_FONT_NAME
except Exception:
    THAI_FONT_NAME = "sans-serif"


def extract_gdrive_id(url_str):
    s = str(url_str).strip()
    if not s or s.lower() == "nan":
        return None
    m1 = re.search(r"/d/([a-zA-Z0-9_-]+)", s)
    m2 = re.search(r"id=([a-zA-Z0-9_-]+)", s)
    if m1:
        return m1.group(1)
    if m2:
        return m2.group(1)
    return None


def apply_pdpa_mask(df_in):
    df_out = df_in.copy()
    if "ชื่อ-สกุล" in df_out.columns:
        df_out["ชื่อ-สกุล"] = [
            f"ผู้ป่วยรายที่ {i + 1}" for i in range(len(df_out))
        ]
    sensitive_cols = [
        "คำนำหน้า",
        "ชื่อเล่น-ฉายา",
        "เลขที่บัตรประชาชน",
        "วัน/เดือน/ปี เกิด",
        "ที่อยู่หลังพันโทษ",
        "ญาติที่สามารถติดต่อได้",
        "เบอร์โทรศัพท์ญาติ",
        "รูป",
    ]
    for c in sensitive_cols:
        if c in df_out.columns:
            df_out[c] = "*** (ซ่อนข้อมูล) ***"
    return df_out


@st.cache_data(ttl=300)
def generate_excel_with_images(export_df, export_cols):
    import openpyxl
    from openpyxl.utils import get_column_letter
    from PIL import Image as PILImage

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="ข้อมูลผู้ป่วย")
        if "หน้า" in export_cols:
            ws = writer.sheets["ข้อมูลผู้ป่วย"]
            col_idx = export_cols.index("หน้า") + 1
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = 15

            for row_idx, face_url in enumerate(export_df["หน้า"], start=2):
                ws.row_dimensions[row_idx].height = 65
                ws.cell(row=row_idx, column=col_idx).value
