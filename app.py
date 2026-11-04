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
        background: linear-gradient(135deg, #0ea5e9 0%, #0284c7 100%); color: white;
        border-radius: 10px; border: none; padding: 12px 24px; font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# 🔥 ล็อก URL ไปที่หน้าชีตที่ถูกต้อง (gid=1105708359)
SHEET_URL = "https://docs.google.com/spreadsheets/d/1dtpMxycg0en1_zdtsQreeLohqOVEqNyZyxCI26O5Zlc/export?format=csv&gid=1105708359"

col_title, col_btn = st.columns([3, 1])
with col_title:
    st.title("🌿 ระบบ Report ข้อมูลจิตเวช")
    st.markdown("<p style='color:#64748b; font-size:16px; margin-top:-10px; font-weight:500;'>ระบบจัดการข้อมูลและส่งออกรายงานอัตโนมัติ (PDF / Excel)</p>", unsafe_allow_html=True)
with col_btn:
    st.write("") 
    st.link_button("📝 เปิดแก้ไขข้อมูลบน Google Sheet", "https://docs.google.com/spreadsheets/d/1dtpMxycg0en1_zdtsQreeLohqOVEqNyZyxCI26O5Zlc/edit#gid=1105708359", type="secondary", use_container_width=True)

@st.cache_data(ttl=60, show_spinner="⏳ กำลังโหลดข้อมูล กรุณารอสักครู่...")
def load_data(url):
    try:
        return pd.read_csv(url, on_bad_lines='skip')
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการดึงข้อมูล: {e}")
        return None

df = load_data(SHEET_URL)

if df is not None:
    df.columns = df.columns.astype(str).str.strip().str.replace('\u200b', '')
    
    # 🔥 รองรับการเปลี่ยนชื่อคอลัมน์ นัดแพทย์ -> นัด
    if 'นัดแพทย์' in df.columns and 'นัด' not in df.columns:
        df.rename(columns={'นัดแพทย์': 'นัด'}, inplace=True)
        
    required_cols = ['ชื่อ-สกุล', 'เพศ', 'สถานะ', 'Dx', 'อาการปัจจุบัน', 'หน้า', 'ห้อง', 'แพทย์', 'นัด']
    for col in required_cols:
        if col not in df.columns:
            df[col] = "" 
                
    if "คอลัมน์ 1" in df.columns:
        target_col = "คอลัมน์ 1"
    else:
        target_col = df.columns[0]
    
    if target_col in df.columns:
        df[target_col] = df[target_col].astype(str).str.strip()
        
        unique_dates = ["ทั้งหมด"] + sorted(list(set([d for d in df[target_col].dropna().unique() if str(d).strip() != '' and str(d).lower() != 'nan'])))
        unique_doctors = ["ทั้งหมด"] + sorted(list(set([str(d).strip() for d in df['แพทย์'].dropna().unique() if str(d).strip() != '' and str(d).lower() != 'nan'])))
        unique_statuses = ["ทั้งหมด"] + sorted(list(set([str(d).strip() for d in df['สถานะ'].dropna().unique() if str(d).strip() != '' and str(d).lower() != 'nan'])))
        
        col_filter1, col_filter2, col_filter3, col_filter4 = st.columns(4)
        with col_filter1:
            selected_dates = st.multiselect(f"📅 เลือกวันที่:", options=unique_dates, default=["ทั้งหมด"])
        with col_filter2:
            selected_doctors = st.multiselect("🩺 เลือกแพทย์:", options=unique_doctors, default=["ทั้งหมด"])
        with col_filter3:
            selected_statuses = st.multiselect("📌 เลือกสถานะ:", options=unique_statuses, default=["ทั้งหมด"])
        with col_filter4:
            selected_colors = st.multiselect("🎨 เลือกระดับสี:", options=["ทั้งหมด", "🔴 แดง", "🟠 ส้ม", "🟡 เหลือง", "🟢 เขียว", "⚪ ไม่มีข้อมูล"], default=["ทั้งหมด"])
            
        if not selected_dates or not selected_doctors or not selected_statuses or not selected_colors:
            st.warning("⚠️ กรุณาเลือกตัวกรองอย่างน้อย 1 รายการในแต่ละช่อง")
        else:
            filtered_df = df.copy()
            if "ทั้งหมด" not in selected_dates:
                filtered_df = filtered_df[filtered_df[target_col].isin(selected_dates)]
            if "ทั้งหมด" not in selected_doctors:
                filtered_df = filtered_df[filtered_df['แพทย์'].astype(str).str.strip().isin(selected_doctors)]
            if "ทั้งหมด" not in selected_statuses:
                filtered_df = filtered_df[filtered_df['สถานะ'].astype(str).str.strip().isin(selected_statuses)]
                
            if "ทั้งหมด" not in selected_colors:
                mask = pd.Series(False, index=filtered_df.index)
                if "🔴 แดง" in selected_colors: mask |= filtered_df['หน้า'].fillna("").astype(str).str.contains("15P_z1gObqnm29vn-afAJ4JeMRQ4Y-IZw")
                if "🟠 ส้ม" in selected_colors: mask |= filtered_df['หน้า'].fillna("").astype(str).str.contains("1Vkl3jyY4W9h3Mv_l17xlWmbNw1A4p4-P")
                if "🟡 เหลือง" in selected_colors: mask |= filtered_df['หน้า'].fillna("").astype(str).str.contains("1YlAPW2PBMUbkuRt0unWjJTolQ9aAp48Y")
                if "🟢 เขียว" in selected_colors: mask |= filtered_df['หน้า'].fillna("").astype(str).str.contains("1Wu3vMN2idLhA5fWlY4ZsGZ64Uf_c-f-B")
                if "⚪ ไม่มีข้อมูล" in selected_colors:
                    has_color = filtered_df['หน้า'].fillna("").astype(str).str.contains("15P_z1gObqnm29vn-afAJ4JeMRQ4Y-IZw|1Vkl3jyY4W9h3Mv_l17xlWmbNw1A4p4-P|1YlAPW2PBMUbkuRt0unWjJTolQ9aAp48Y|1Wu3vMN2idLhA5fWlY4ZsGZ64Uf_c-f-B")
                    mask |= (~has_color)
                filtered_df = filtered_df[mask]

            if len(filtered_df) > 0:
                filtered_df['นัด'] = filtered_df['นัด'].fillna('')
                filtered_df = filtered_df.sort_values(by='นัด').reset_index(drop=True)

                st.markdown("---")
                
                total_patients = len(filtered_df)
                new_patients = len(filtered_df[filtered_df['สถานะ'].astype(str).str.strip() == 'รายใหม่'])
                red_cases = sum(filtered_df['หน้า'].fillna("").astype(str).str.contains("15P_z1gObqnm29vn-afAJ4JeMRQ4Y-IZw", na=False))
                
                filtered_df['เพศ'] = filtered_df['เพศ'].astype(str).str.strip()
                male_count = len(filtered_df[filtered_df['เพศ'] == 'ชาย'])
                female_count = len(filtered_df[filtered_df['เพศ'] == 'หญิง'])
                
                dash_col1, dash_col2, dash_col3, dash_col4 = st.columns([1, 1.2, 1.2, 1.5])
                dash_col1.metric("👥 ผู้ป่วยทั้งหมด", f"{total_patients} ราย")
                dash_col2.metric("🆕 ผู้ป่วยรายใหม่", f"{new_patients} ราย")
                dash_col3.metric("🚨 เคสเฝ้าระวัง (แดง)", f"{red_cases} ราย")
                
                with dash_col4:
                    st.markdown("<p style='font-size:13px; font-weight:700; color:#64748b; margin-bottom:-10px;'>🚻 สัดส่วนเพศ (Gender)</p>", unsafe_allow_html=True)
                    if male_count + female_count > 0:
                        fig_dash, ax_dash = plt.subplots(figsize=(2.5, 1.5))
                        fig_dash.patch.set_alpha(0.0) 
                        sizes, labels, colors = [], [], []
                        if male_count > 0: sizes.append(male_count); labels.append('Male'); colors.append('#38bdf8')
                        if female_count > 0: sizes.append(female_count); labels.append('Female'); colors.append('#f472b6')
                        wedges, texts, autotexts = ax_dash.pie(sizes, labels=labels, autopct='%1.1f%%', colors=colors, startangle=90, textprops={'fontsize': 8, 'fontfamily': THAI_FONT_NAME, 'fontweight': 'bold'})
                        for text in texts: text.set_color('#475569') 
                        for autotext in autotexts: autotext.set_color('#ffffff') 
                        ax_dash.axis('equal')
                        st.pyplot(fig_dash)
                        plt.close(fig_dash)
                    else:
                        st.info("ไม่มีข้อมูลเพศ")
                        
                st.markdown("---")

                url_mapping = {
                    "แดง": "https://drive.google.com/uc?id=15P_z1gObqnm29vn-afAJ4JeMRQ4Y-IZw",
                    "ส้ม": "https://drive.google.com/uc?id=1Vkl3jyY4W9h3Mv_l17xlWmbNw1A4p4-P",
                    "เหลือง": "https://drive.google.com/uc?id=1YlAPW2PBMUbkuRt0unWjJTolQ9aAp48Y",
                    "เขียว": "https://drive.google.com/uc?id=1Wu3vMN2idLhA5fWlY4ZsGZ64Uf_c-f-B",
                    "ไม่มีข้อมูล": ""
                }
                
                def url_to_label(url):
                    url_str = str(url).strip()
                    if "15P_z1gObqnm29vn-afAJ4JeMRQ4Y-IZw" in url_str: return "แดง"
                    if "1Vkl3jyY4W9h3Mv_l17xlWmbNw1A4p4-P" in url_str: return "ส้ม"
                    if "1YlAPW2PBMUbkuRt0unWjJTolQ9aAp48Y" in url_str: return "เหลือง"
                    if "1Wu3vMN2idLhA5fWlY4ZsGZ64Uf_c-f-B" in url_str: return "เขียว"
                    return "ไม่มีข้อมูล"

                def highlight_and_group(subset_df):
                    styles = pd.DataFrame('', index=subset_df.index, columns=subset_df.columns)
                    if 'นัด' in subset_df.columns:
                        valid_dates = [d for d in subset_df['นัด'].unique() if str(d).strip() != '' and str(d).lower() != 'nan']
                        pastel_colors = ['#E0F2F1', '#FFF9C4', '#F3E5F5', '#E3F2FD', '#FBE9E7']
                        color_map = {date: pastel_colors[i % len(pastel_colors)] for i, date in enumerate(valid_dates)}
                        for idx, row in subset_df.iterrows():
                            appt_date = row.get('นัด', '')
                            if appt_date in color_map:
                                styles.loc[idx, 'นัด'] = f'background-color: {color_map[appt_date]}; font-weight: 600;'
                    if 'หน้า' in subset_df.columns:
                        red_mask = subset_df['หน้า'].fillna("").astype(str).str.contains("แดง|15P_z1gObqnm29vn-afAJ4JeMRQ4Y-IZw")
                        for col in styles.columns: styles.loc[red_mask, col] = 'background-color: #FCE4EC;' 
                    return styles

                tab_pdf, tab_excel = st.tabs(["📄 Telepsychiatry Report", "📊 รายงานทั่วไป (Excel)"])
                
                # ------------------------------------------
                # TAB 1: ระบบรายงาน PDF
                # ------------------------------------------
                with tab_pdf:
                    problem_text = st.text_area("✍ บันทึกปัญหา / อุปสรรค (ถ้ามี):", placeholder="พิมพ์ปัญหาที่พบ...")
                    suggestion_text = st.text_area("💡 ข้อเสนอแนะ (ถ้ามี):", placeholder="พิมพ์ข้อเสนอแนะ...")
                    
                    st.markdown("**📋 Preview ข้อมูล (เรียงตามวันนัดแล้ว / ดับเบิลคลิกแก้ไขได้เลยครับ):**")
                    preview_cols = ['ชื่อ-สกุล', 'สถานะ', 'Dx', 'อาการปัจจุบัน', 'หน้า', 'นัด', 'ห้อง', 'แพทย์']
                    avail_cols = [c for c in preview_cols if c in filtered_df.columns]
                    
                    pdf_display_df = filtered_df[avail_cols].copy().fillna("")
                    if 'หน้า' in pdf_display_df.columns:
                        pdf_display_df['หน้า'] = pdf_display_df['หน้า'].apply(url_to_label).astype(str)

                    edited_pdf_display = st.data_editor(
                        pdf_display_df.style.apply(highlight_and_group, axis=None),
                        use_container_width=True, hide_index=True, column_order=avail_cols, key="pdf_editor_with_dropdown",
                        column_config={
                            "หน้า": st.column_config.SelectboxColumn("ระดับสี", options=["แดง", "ส้ม", "เหลือง", "เขียว", "ไม่มีข้อมูล"]),
                            "นัด": st.column_config.TextColumn("นัดครั้งถัดไป")
                        }
                    )
                    
                    edited_df = filtered_df.copy()
                    for col in avail_cols:
                        if col in edited_pdf_display.columns:
                            edited_df[col] = edited_pdf_display[col]
                            
                    if 'หน้า' in edited_pdf_display.columns:
                        edited_df['หน้า'] = edited_pdf_display['หน้า'].apply(lambda x: url_mapping.get(x, ""))
                    
                    if st.button("🚀 สร้างรายงาน PDF", type="primary"):
                        with st.status("⏳ กำลังสร้างรายงาน PDF กรุณารอสักครู่...", expanded=True) as status:
                            html_rows = ""
                            for row_num, (idx, row) in enumerate(edited_df.iterrows(), start=1):
                                name = str(row.get('ชื่อ-สกุล', '')).replace('nan', '')
                                status_txt = str(row.get('สถานะ', '')).replace('nan', '')
                                dx = str(row.get('Dx', '')).replace('nan', '')
                                symptom = str(row.get('อาการปัจจุบัน', '')).replace('nan', '')
                                room = str(row.get('ห้อง', '')).replace('nan', '') 
                                appt = str(row.get('นัด', '')).replace('nan', '')
                                doc = str(row.get('แพทย์', '')).replace('nan', '')
                                face_url = str(row.get('หน้า', '')).strip()
                                match1 = re.search(r'/d/([a-zA-Z0-9_-]+)', face_url)
                                match2 = re.search(r'id=([a-zA-Z0-9_-]+)', face_url)
                                gdrive_id = match1.group(1) if match1 else (match2.group(1) if match2 else None)
                                img_tag = f'<img src="https://drive.google.com/uc?id={gdrive_id}" style="width:26px;height:26px;object-fit:cover;border-radius:4px;"/>' if gdrive_id else "-"
                                
                                html_rows += f"""<tr>
                                    <td style="text-align:center;">{row_num}</td>
                                    <td style="font-weight:bold; color:#1e293
