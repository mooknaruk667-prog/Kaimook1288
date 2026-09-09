import streamlit as st
import pandas as pd
import base64
from weasyprint import HTML
import io
import re
import os
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import urllib.request
import matplotlib.font_manager as fm

# ==========================================
# ⚙️ โหลดฟอนต์ Sarabun สำหรับวาดกราฟ
# ==========================================
@st.cache_resource
def setup_thai_font():
    font_path = "Sarabun-Regular.ttf"
    if not os.path.exists(font_path):
        try:
            # ใช้ฟอนต์ Sarabun แบบธรรมดา (ไม่หนา)
            urllib.request.urlretrieve("https://github.com/googlefonts/sarabun/raw/main/fonts/ttf/Sarabun-Regular.ttf", font_path)
        except Exception:
            pass
    try:
        fm.fontManager.addfont(font_path)
        plt.rcParams['font.family'] = 'Sarabun'
    except Exception:
        pass

setup_thai_font()
plt.switch_backend('Agg')

# ==========================================
# 🚀 เริ่มต้นโปรแกรม Streamlit
# ==========================================
st.set_page_config(page_title="ระบบสร้างรายงาน PDF", page_icon="📄", layout="wide")

st.title("📄 ระบบสร้างรายงาน Telepsychiatry (PDF)")
st.markdown("ดึงข้อมูลจาก Google Sheet และสรุปเป็น PDF")

# URL ของ Google Sheet
SHEET_URL = "https://docs.google.com/spreadsheets/d/1dtpMxycg0en1_zdtsQreeLohqOVEqNyZyxCI26O5Zlc/export?format=csv"

@st.cache_data(ttl=60)
def load_data(url):
    try:
        df = pd.read_csv(url, on_bad_lines='skip')
        return df
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการดึงข้อมูล: {e}")
        return None

df = load_data(SHEET_URL)

if df is not None:
    st.success(f"โหลดข้อมูลสำเร็จ! จำนวนทั้งหมด {len(df)} รายการ")
    
    target_col = "คอลัมน์ 1"
    
    if target_col in df.columns:
        df[target_col] = df[target_col].astype(str).str.strip()
        unique_dates = ["ทั้งหมด"] + sorted(list(set([d for d in df[target_col].unique() if d and str(d).lower() != 'nan'])))
        
        selected_date = st.selectbox("📅 เลือกวันที่ (จาก คอลัมน์ 1):", options=unique_dates)
        
        problem_text = st.text_area("✍️ บันทึกปัญหา / อุปสรรค (ถ้ามี):", placeholder="พิมพ์ปัญหาหรืออุปสรรคที่พบในวันนี้ที่นี่...")
        suggestion_text = st.text_area("💡 ข้อเสนอแนะ (ถ้ามี):", placeholder="พิมพ์ข้อเสนอแนะเพิ่มเติมที่นี่...")
        
        if st.button("🚀 สร้าง PDF", type="primary"):
            with st.spinner('กำลังประมวลผลข้อมูลและสร้างไฟล์ PDF...'):
                
                filtered_df = df.copy()
                if selected_date != "ทั้งหมด":
                    filtered_df = filtered_df[filtered_df[target_col] == selected_date]
                
                if len(filtered_df) == 0:
                    st.warning("ไม่พบข้อมูลในวันที่เลือก")
                else:
                    html_rows = ""
                    for row_num, (idx, row) in enumerate(filtered_df.iterrows(), start=1):
                        name = str(row.get('ชื่อ-สกุล', '')).replace('nan', '')
                        status = str(row.get('สถานะ', '')).replace('nan', '')
                        dx = str(row.get('Dx', '')).replace('nan', '')
                        symptom = str(row.get('อาการปัจจุบัน', '')).replace('nan', '')
                        
                        face_url = str(row.get('หน้า', '')).strip()
                        
                        match1 = re.search(r'/d/([a-zA-Z0-9_-]+)', face_url)
                        match2 = re.search(r'id=([a-zA-Z0-9_-]+)', face_url)
                        
                        gdrive_id = None
                        if match1: gdrive_id = match1.group(1)
                        elif match2: gdrive_id = match2.group(1)
                            
                        if gdrive_id:
                            direct_img_url = f"https://drive.google.com/uc?id={gdrive_id}"
                            img_tag = f'<img src="{direct_img_url}" style="width:28px;height:28px;object-fit:cover;border-radius:4px; box-shadow: 0 1px 3px rgba(0,0,0,0.2);"/>'
                        else:
                            img_tag = "-"
                        
                        appt = str(row.get('นัด', '')).replace('nan', '')
                        doc = str(row.get('แพทย์', '')).replace('nan', '')
                        
                        html_rows += f"""<tr>
                            <td style="text-align:center;">{row_num}</td>
                            <td style="font-weight:bold; color:#1e293b;">{name}</td>
                            <td>{status}</td>
                            <td>{dx}</td>
                            <td>{symptom}</td>
                            <td style="text-align:center;">{img_tag}</td>
                            <td style="color:#0369a1;">{appt}</td>
                            <td>{doc}</td>
                        </tr>"""

                    filtered_df['สถานะ'] = filtered_df['สถานะ'].astype(str).str.strip()
                    filtered_df['เพศ'] = filtered_df['เพศ'].astype(str).str.strip()

                    old_cases = filtered_df[filtered_df['สถานะ'] == 'รายเก่า']
                    new_cases = filtered_df[filtered_df['สถานะ'] == 'รายใหม่']

                    old_m = len(old_cases[old_cases['เพศ'] == 'ชาย'])
                    old_f = len(old_cases[old_cases['เพศ'] == 'หญิง'])
                    new_m = len(new_cases[new_cases['เพศ'] == 'ชาย'])
                    new_f = len(new_cases[new_cases['เพศ'] == 'หญิง'])

                    # ==========================================
                    # 📊 1. กราฟวงกลม Dx
                    # ==========================================
                    dx_counts = filtered_df['Dx'].value_counts()
                    chart_img_tag = ""
                    
                    valid_dx = {k: v for k, v in dx_counts.items() if str(k).lower() != 'nan'}
                    if valid_dx:
                        fig1, ax1 = plt.subplots(figsize=(2.8, 2.8))
                        total_dx = sum(valid_dx.values())
                        
                        labels1 = [f"{k}\n{v} ({v/total_dx*100:.1f}%)" for k, v in valid_dx.items()]
                        wedges1, texts1 = ax1.pie(
                            valid_dx.values(), 
                            labels=labels1, 
                            labeldistance=0.5, 
                            startangle=90,
                            # ตั้งค่า font weight เป็น normal และใช้สีเข้ม
                            textprops={'fontsize': 10, 'color': '#111827', 'weight': 'normal', 'ha': 'center'},
                            colors=plt.cm.tab20.colors,
                            radius=1 
                        )
                            
                        ax1.axis('equal') 
                        img_buf1 = io.BytesIO()
                        plt.savefig(img_buf1, format='png', bbox_inches='tight', transparent=True, dpi=120)
                        img_buf1.seek(0)
                        chart_base64_1 = base64.b64encode(img_buf1.read()).decode('utf-8')
                        plt.close(fig1)
                        chart_img_tag = f'<img src="data:image/png;base64,{chart_base64_1}" style="width:100%; max-width:200px; display:block; margin:auto;"/>'
                    else:
                        chart_img_tag = "<p style='text-align:center; color:#94a3b8;'>ไม่มีข้อมูล Dx</p>"

                    # ==========================================
                    # 📊 2. กราฟวงกลม สรุปเคสสี
                    # ==========================================
                    color_map = {'แดง': '#F44336', 'ส้ม': '#FF9800', 'เหลือง': '#FACC15', 'เขียว': '#4CAF50', 'เทา': '#9E9E9E'}
                    level_counts = {'แดง': 0, 'ส้ม': 0, 'เหลือง': 0, 'เขียว': 0, 'เทา': 0}
                    
                    for face_url in filtered_df['หน้า'].astype(str):
                        if "15P_z1gObqnm29vn-afAJ4JeMRQ4Y-IZw" in face_url: level_counts['แดง'] += 1
                        elif "1Vkl3jyY4W9h3Mv_l17xlWmbNw1A4p4-P" in face_url: level_counts['ส้ม'] += 1
                        elif "1YlAPW2PBMUbkuRt0unWjJTolQ9aAp48Y" in face_url: level_counts['เหลือง'] += 1
                        elif "1Wu3vMN2idLhA5fWlY4ZsGZ64Uf_c-f-B" in face_url: level_counts['เขียว'] += 1
                        else: level_counts['เทา'] += 1
                        
                    active_levels = {k: v for k, v in level_counts.items() if v > 0}
                    level_chart_img_tag = ""
                    
                    if active_levels:
                        fig2, ax2 = plt.subplots(figsize=(2.8, 2.8))
                        colors2 = [color_map[k] for k in active_levels.keys()]
                        total_levels = sum(active_levels.values())
                        
                        wedges2, texts2, autotexts2 = ax2.pie(
                            active_levels.values(),
                            autopct=lambda p: f"{int(round(p * total_levels / 100))}\n({p:.1f}%)",
                            startangle=90,
                            labeldistance=0.5,
                            # ตั้งค่า font weight เป็น normal และใช้สีเข้ม
                            textprops={'fontsize': 10, 'color': '#111827', 'weight': 'normal', 'ha': 'center'},
                            colors=colors2,
                            radius=1 
                        )
                                
                        ax2.axis('equal')
                        img_buf2 = io.BytesIO()
                        plt.savefig(img_buf2, format='png', bbox_inches='tight', transparent=True, dpi=120)
                        img_buf2.seek(0)
                        chart_base64_2 = base64.b64encode(img_buf2.read()).decode('utf-8')
                        plt.close(fig2)
                        level_chart_img_tag = f'<img src="data:image/png;base64,{chart_base64_2}" style="width:100%; max-width:200px; display:block; margin:auto;"/>'
                    else:
                        level_chart_img_tag = "<p style='text-align:center; color:#94a3b8;'>ไม่มีข้อมูล</p>"

                    # ==========================================
                    # 🖼️ ดึงไฟล์ภาพโลโก้จากลิงก์ Google Drive
                    # ==========================================
                    logo_src = "https://drive.google.com/uc?id=1KYrHcRg6dvs2h0nfDf7ZxpzWpLnCqnjY"

                    title_text = f"รายงาน Telepsychiatry วันที่ {selected_date}" if selected_date != "ทั้งหมด" else "รายงาน Telepsychiatry (ทั้งหมด)"

                    bottom_sections = ""
                    if problem_text.strip():
                        bottom_sections += f"""
                        <div style="background-color: #fef2f2; border-left: 5px solid #ef4444; padding: 15px; margin-top: 20px; border-radius: 4px;">
                            <h3 style="color: #b91c1c; margin-top: 0; font-size: 13pt;">ปัญหา / อุปสรรค</h3>
                            <p style="margin: 0; line-height: 1.6; white-space: pre-line;">{problem_text}</p>
                        </div>
                        """
                    if suggestion_text.strip():
                        bottom_sections += f"""
                        <div style="background-color: #f0fdf4; border-left: 5px solid #22c55e; padding: 15px; margin-top: 15px; border-radius: 4px;">
                            <h3 style="color: #15803d; margin-top: 0; font-size: 13pt;">ข้อเสนอแนะ</h3>
                            <p style="margin: 0; line-height: 1.6; white-space: pre-line;">{suggestion_text}</p>
                        </div>
                        """

                    html_content = f"""
                    <!DOCTYPE html>
                    <html lang="th">
                    <head>
                    <meta charset="UTF-8">
                    <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap" rel="stylesheet">
                    <style>
                        @page {{ 
                            size: A4 landscape; 
                            margin: 15mm; 
                        }}
                        body {{ 
                            font-family: 'TH Sarabun PSK', 'Sarabun', sans-serif; 
                            font-size: 12pt; 
                            color: #334155; 
                            line-height: 1.5;
                        }}
                        
                        /* โลโก้ ฝั่งซ้าย และห่างจากขอบกระดาษซ้าย 3 ซม. ห่างจากขอบบน 1 ซม. */
                        /* margin-top: -5mm (หน้ากระดาษ 15mm - 5mm = 10mm คือ 1 ซม.) */
                        /* margin-left: 15mm (หน้ากระดาษ 15mm + 15mm = 30mm คือ 3 ซม.) */
                        .header-logo {{
                            position: absolute;
                            top: -5mm; 
                            left: 15mm; 
                            text-align: center;
                            width: 180px;
                        }}
                        .header-logo img {{
                            width: 65px; 
                            height: auto;
                        }}
                        .header-logo p {{
                            font-size: 11pt;
                            font-weight: bold;
                            color: #1e293b;
                            margin-top: 5px;
                            line-height: 1.1;
                        }}

                        h1 {{ 
                            text-align: center; 
                            color: #0f172a; 
                            border-bottom: 2px solid #cbd5e1; 
                            padding-bottom: 10px; 
                            margin-top: 20px;
                            margin-bottom: 25px;
                            font-size: 22pt;
                            padding-left: 180px; 
                            padding-right: 180px; 
                        }}
                        
                        .summary-container {{
                            width: 100%;
                            border-collapse: separate;
                            border-spacing: 15px 0; 
                            margin-bottom: 25px;
                        }}
                        .summary-box {{
                            background-color: #f8fafc;
                            border: 1px solid #e2e8f0;
                            border-radius: 8px;
                            padding: 15px 20px;
                            vertical-align: middle;
                        }}
                        .summary-box h3 {{
                            margin-top: 0; 
                            color: #0369a1; 
                            font-size: 14pt;
                            border-bottom: 1px solid #cbd5e1;
                            padding-bottom: 8px;
                            margin-bottom: 12px;
                        }}

                        .data-table {{ 
                            width: 100%; 
                            border-collapse: collapse; 
                            margin-top: 10px;
                        }}
                        .data-table th, .data-table td {{ 
                            padding: 10px 8px; 
                            vertical-align: middle; 
                        }}
                        .data-table th {{ 
                            background-color: #1e293b; 
                            color: #ffffff; 
                            text-align: center;
                            font-weight: 600;
                            font-size: 12pt;
                        }}
                        .data-table td {{
                            border-bottom: 1px solid #e2e8f0;
                        }}
                        .data-table tr:nth-child(even) {{ 
                            background-color: #f8fafc; 
                        }}

                        .signature-section {{ 
                            margin-top: 40px; 
                            text-align: left; 
                            page-break-inside: avoid; 
                        }}
                    </style>
                    </head>
                    <body>
                        <div class="header-logo">
                            <img src="{logo_src}" alt="Logo">
                            <p>สถานพยาบาลเรือนจำ<br>จังหวัดบุรีรัมย์</p>
                        </div>

                        <h1>{title_text}</h1>
                        
                        <table class="summary-container" style="margin-left: -15px; margin-right: -15px; width: calc(100% + 30px);">
                            <tr>
                                <td class="summary-box" style="width: 30%; vertical-align: top;">
                                    <h3>สรุปสถานะผู้ป่วย</h3>
                                    <p style="margin: 10px 0 0 0; line-height: 1.8;">
                                        <strong>ผู้ป่วยรายเก่า:</strong> {len(old_cases)} ราย (ชาย {old_m}, หญิง {old_f})<br>
                                        <strong>ผู้ป่วยรายใหม่:</strong> {len(new_cases)} ราย (ชาย {new_m}, หญิง {new_f})
                                    </p>
                                </td>
                                <td class="summary-box" style="width: 35%; text-align: center;">
                                    <h3 style="text-align: left;">สรุปการวินิจฉัยโรค (Dx)</h3>
                                    {chart_img_tag}
                                </td>
                                <td class="summary-box" style="width: 35%; text-align: center;">
                                    <h3 style="text-align: left;">สรุปเคสสีตามระดับ</h3>
                                    {level_chart_img_tag}
                                </td>
                            </tr>
                        </table>

                        <table class="data-table">
                            <thead>
                                <tr>
                                    <th style="width: 5%;">ที่</th>
                                    <th style="width: 16%;">ชื่อ-สกุล</th>
                                    <th style="width: 8%;">สถานะ</th>
                                    <th style="width: 10%;">Dx</th>
                                    <th style="width: 32%;">อาการปัจจุบัน</th>
                                    <th style="width: 6%;">ระดับ</th>
                                    <th style="width: 10%;">นัดครั้งถัดไป</th>
                                    <th style="width: 13%;">แพทย์</th>
                                </tr>
                            </thead>
                            <tbody>{html_rows}</tbody>
                        </table>

                        {bottom_sections}

                        <div class="signature-section">
                            <p style="margin-bottom: 20px; line-height: 1.6;">เรียน ผู้บัญชาการเรือนจำฯ<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- เพื่อโปรดทราบ</p>
                            <br><br><br>
                            <div style="display: inline-block; text-align: center;">
                                <p style="margin: 0; font-weight: bold; font-size: 13pt;">(นางสาวเดือนนภา เบี้ยชาติไทย)</p>
                                <p style="margin: 5px 0 0 0; color: #475569;">นักจิตวิทยาปฏิบัติการ</p>
                            </div>
                        </div>
                    </body>
                    </html>
                    """
                    
                    pdf_bytes = HTML(string=html_content).write_pdf()
                    
                    st.success(f"สร้าง PDF สำเร็จ! (ข้อมูล {len(filtered_df)} รายการ)")
                    st.download_button(
                        label="📥 คลิกที่นี่เพื่อดาวน์โหลดไฟล์ PDF",
                        data=pdf_bytes,
                        file_name=f"Telepsychiatry_Report_{selected_date.replace('/', '-') if selected_date != 'ทั้งหมด' else 'All'}.pdf",
                        mime="application/pdf"
                    )
    else:
         st.error(f"❌ เกิดข้อผิดพลาด: ไม่พบคอลัมน์ชื่อ '{target_col}' ในไฟล์ Sheet ของคุณ")
