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
# ⚙️ โหลดฟอนต์ Sarabun (แก้ปัญหาสี่เหลี่ยมแบบเด็ดขาด)
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

# 🔥 แทรก CSS ตกแต่งหน้าเว็บ สไตล์ Modern UI / SaaS Dashboard
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
        box-shadow: 0 4px 15px -3px rgba(0, 0, 0, 0.05), 0 4px 6px -4px rgba(0, 0, 0, 0.025);
        border: 1px solid #e2e8f0; border-left: 6px solid #0ea5e9; 
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    [data-testid="stMetric"]:hover {
        transform: translateY(-3px); box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1);
    }
    [data-testid="stMetricLabel"] {
        font-size: 13px !important; font-weight: 600 !important;
        color: #64748b !important; text-transform: uppercase; letter-spacing: 0.5px;
    }
    [data-testid="stMetricValue"] {
        font-size: 28px !important; color: #0f172a !important; font-weight: 800 !important; margin-top: 5px;
    }
    
    .stMultiSelect div[data-baseweb="select"] {
        border-radius: 10px; background-color: #ffffff; border: 1px solid #cbd5e1;
    }
    
    .stTabs [data-baseweb="tab-list"] { gap: 12px; border-bottom: 2px solid #e2e8f0; }
    .stTabs [data-baseweb="tab"] {
        background-color: transparent; border-radius: 8px 8px 0 0;
        padding: 12px 24px; color: #64748b; font-weight: 600; border: none; transition: all 0.2s;
    }
    .stTabs [aria-selected="true"] {
        color: #0ea5e9; background-color: #f0f9ff; border-bottom: 3px solid #0ea5e9; 
    }
    
    .stButton>button {
        background: linear-gradient(135deg, #0ea5e9 0%, #0284c7 100%); color: white;
        border-radius: 10px; border: none; padding: 12px 24px;
        box-shadow: 0 4px 6px -1px rgba(14, 165, 233, 0.3); transition: all 0.3s ease; font-weight: 600;
    }
    .stButton>button:hover {
        transform: translateY(-2px); box-shadow: 0 6px 10px -1px rgba(14, 165, 233, 0.4); color: white;
    }
    
    [data-testid="stDataFrame"] {
        border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); border: 1px solid #e2e8f0;
    }
    .stTextArea textarea { border-radius: 10px; border: 1px solid #cbd5e1; }
</style>
""", unsafe_allow_html=True)

# URL ของ Google Sheet
SHEET_URL = "https://docs.google.com/spreadsheets/d/1dtpMxycg0en1_zdtsQreeLohqOVEqNyZyxCI26O5Zlc/export?format=csv"

col_title, col_btn = st.columns([3, 1])
with col_title:
    st.title("🌿 ระบบ Report ข้อมูลจิตเวช")
    st.markdown("<p style='color:#64748b; font-size:16px; margin-top:-10px; font-weight:500;'>ระบบจัดการข้อมูลและส่งออกรายงานอัตโนมัติ (PDF / Excel) เพื่อสุขภาพจิตที่ดี</p>", unsafe_allow_html=True)
with col_btn:
    st.write("") 
    st.link_button("📝 เปิดแก้ไขข้อมูลถาวรบน Google Sheet", SHEET_URL.replace("/export?format=csv", "/edit"), type="secondary", use_container_width=True)

# 🔥 เพิ่ม Spinner ตอนโหลดข้อมูลครั้งแรก
@st.cache_data(ttl=60, show_spinner="⏳ กำลังโหลดและซิงค์ข้อมูลจาก Google Sheet กรุณารอสักครู่...")
def load_data(url):
    try:
        df = pd.read_csv(url, on_bad_lines='skip')
        return df
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการดึงข้อมูล: {e}")
        return None

df = load_data(SHEET_URL)

if df is not None:
    # 🔥 1. ล้างอักขระแปลกปลอมและช่องว่างหน้า-หลัง ในชื่อหัวตารางทั้งหมด
    df.columns = df.columns.astype(str).str.strip().str.replace('\u200b', '')
    
    # 🔥 2. สร้างเกราะป้องกัน KeyError: ตรวจสอบคอลัมน์ที่จำเป็น ถ้าไม่มีให้สร้างคอลัมน์ว่างอัตโนมัติ
    required_cols = ['ชื่อ-สกุล', 'เพศ', 'สถานะ', 'Dx', 'อาการปัจจุบัน', 'หน้า', 'ห้อง', 'แพทย์', 'นัด']
    for col in required_cols:
        if col not in df.columns:
            # ค้นหาแบบ Fuzzy เผื่อพิมพ์ผิด เช่น 'แพทย์ประจำตัว'
            match = next((c for c in df.columns if col.lower() in c.lower()), None)
            if match:
                df.rename(columns={match: col}, inplace=True)
            else:
                df[col] = "" # เติมคอลัมน์ว่างให้ระบบทำงานต่อไปได้โดยไม่จอแดง
                
    # 🔍 ระบบค้นหาคอลัมน์วันที่อัตโนมัติ
    if "คอลัมน์ 1" in df.columns:
        target_col = "คอลัมน์ 1"
    elif "วันที่" in df.columns:
        target_col = "วันที่"
    elif "Date" in df.columns:
        target_col = "Date"
    else:
        target_col = df.columns[0]
    
    if target_col in df.columns:
        df[target_col] = df[target_col].astype(str).str.strip()
        
        unique_dates = ["ทั้งหมด"] + sorted(list(set([d for d in df[target_col].unique() if d and str(d).lower() != 'nan'])))
        unique_doctors = ["ทั้งหมด"] + sorted(list(set([str(d).strip() for d in df['แพทย์'].unique() if pd.notna(d) and str(d).lower() != 'nan'])))
        unique_statuses = ["ทั้งหมด"] + sorted(list(set([str(d).strip() for d in df['สถานะ'].unique() if pd.notna(d) and str(d).lower() != 'nan'])))
        
        # 🔥 แก้ไขแบ่งเป็น 4 คอลัมน์เพื่อใส่ตัวกรองสี
        col_filter1, col_filter2, col_filter3, col_filter4 = st.columns(4)
        with col_filter1:
            selected_dates = st.multiselect(f"📅 เลือกวันที่ ({target_col}):", options=unique_dates, default=["ทั้งหมด"])
        with col_filter2:
            selected_doctors = st.multiselect("🩺 เลือกแพทย์:", options=unique_doctors, default=["ทั้งหมด"])
        with
