import streamlit as st
import pandas as pd
import base64
from weasyprint import HTML
import io
import re
import os
import numpy as np
import matplotlib.pyplot as plt
import urllib.request
import matplotlib.font_manager as fm

# ==========================================
# ⚙️ โหลดฟอนต์ Sarabun
# ==========================================
@st.cache_resource
def download_thai_font():
    font_path = "Sarabun-Regular.ttf"
    if not os.path.exists(font_path):
        try:
            urllib.request.urlretrieve("https://github.com/googlefonts/sarabun/raw/main/fonts/ttf/Sarabun-Regular.ttf", font_path)
        except Exception:
            pass
    return font_path

THAI_FONT_PATH = download_thai_font()
try:
    fm.fontManager.addfont(THAI_FONT_PATH)
    THAI_FONT_NAME = fm.FontProperties(fname=THAI_FONT_PATH).get_name()
    plt.rcParams['font.family'] = THAI_FONT_NAME
except Exception:
    THAI_FONT_NAME = 'sans-serif'

plt.switch_backend('Agg')

# ==========================================
# 🚀 ฟังก์ชันจัดเตรียมไฟล์ Excel พร้อมรูปภาพ
# ==========================================
@st.cache_data(ttl=300)
def generate_excel_with_images(export_df, export_cols):
    import openpyxl
    from openpyxl.utils import get_column_letter
    from PIL import Image as PILImage
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        export_df.to_excel(writer, index=False, sheet_name='ข้อมูลผู้ป่วย')
        
        if 'หน้า' in export_cols:
            worksheet = writer.sheets['ข้อมูลผู้ป่วย']
            col_idx = export_cols.index('หน้า') + 1
            col_letter = get_column_letter(col_idx)
            
            worksheet.column_dimensions[col_letter].width = 15
            
            for row_idx, face_url in enumerate(export_df['หน้า'], start=2):
                worksheet.row_dimensions[row_idx].height = 65 
                worksheet.cell(row=row_idx, column=col_idx).value = "" 
                
                face_url = str(face_url).strip()
                match1 = re.search(r'/d/([a-zA-Z0-9_-]+)', face_url)
                match2 = re.search(r'id=([a-zA-Z0-9_-]+)', face_url)
                
                gdrive_id = None
                if match1: gdrive_id = match1.group(1)
                elif match2: gdrive_id = match2.group(1)
                
                if gdrive_id:
                    try:
                        direct_img_url = f"https://drive.google.com/uc?id={gdrive_id}"
                        req = urllib.request.Request(direct_img_url, headers={'User-Agent': 'Mozilla/5.0'})
                        with urllib.request.urlopen(req) as response:
                            img_data = response.read()
                            img = PILImage.open(io.BytesIO(img_data))
                            img.thumbnail((80, 80)) 
                            
                            img_byte_arr = io.BytesIO()
                            img.save(img_byte_arr, format='PNG')
                            img_byte_arr.seek(0)
                            
                            xl_img = openpyxl.drawing.image.Image(img_byte_arr)
                            xl_img.anchor = f"{col_letter}{row_idx}"
                            worksheet.add_image(xl_img)
                    except Exception:
                        worksheet.cell(row=row_idx, column=col_idx).value = "โหลดรูปไม่ได้"
                        
    return buffer.getvalue()

# ==========================================
# 🚀 เริ่มต้นโปรแกรม Streamlit
# ==========================================
st.set_page_config(page_title="ระบบ Report ข้อมูลจิตเวช", page_icon="🌿", layout="wide")

st.markdown("""
<style>
    .stApp { background-color: #F8FAFC; }
    h1 {
        background: -webkit-linear-gradient(45deg, #0284c7, #0d9488);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800 !important;
        font-family: 'Sarabun', sans-serif;
    }
    h2, h3 { color: #1e293b; font-family: 'Sarabun', sans-serif; font-weight: 700; }
    [data-testid="stMetric"] {
        background-color: #ffffff; border-radius: 16px; padding: 20px;
        box-shadow: 0 4px 15px -3px rgba(0, 0, 0, 0.05);
        border: 1px solid #e2e8f0; border-left: 6px solid #0ea5e9; 
    }
    .stButton>button {
        background: linear-gradient(135deg, #0ea5e9 0%, #
