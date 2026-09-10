import streamlit as st
import pandas as pd
import base64
from weasyprint import HTML
import io
import re
import os
import numpy as np
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
st.markdown("ระบบสรุปข้อมูลและออกรายงาน PDF อัตโนมัติ")

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
    target_col = "คอลัมน์ 1"
    
    if target_col in df.columns:
        df[target_col] = df[target_col].astype(str).str.strip()
        
        # ดึงรายชื่อวันที่, แพทย์, และสถานะ
        unique_dates = ["ทั้งหมด"] + sorted(list(set([d for d in df[target_col].unique() if d and str(d).lower() != 'nan'])))
        unique_doctors = ["ทั้งหมด"] + sorted(list(set([str(d).strip() for d in df['แพทย์'].unique() if pd.notna(d) and str(d).lower() != 'nan'])))
        unique_statuses = ["ทั้งหมด"] + sorted(list(set([str(d).strip() for d in df['สถานะ'].unique() if pd.notna(d) and str(d).lower() != 'nan'])))
        
        # ส่วน UI ตัวกรอง (แสดงแบบ 3 คอลัมน์คู่กัน)
        col_filter1, col_filter2, col_filter3 = st.columns(3)
        with col_filter1:
            selected_date = st.selectbox("📅 เลือกวันที่:", options=unique_dates)
        with col_filter2:
            selected_doctor = st.selectbox("🩺 เลือกแพทย์ผู้ตรวจ:", options=unique_doctors)
        with col_filter3:
            selected_status = st.selectbox("📌 เลือกสถานะผู้ป่วย:", options=unique_statuses)
        
        problem_text = st.text_area("✍️ บันทึกปัญหา / อุปสรรค (ถ้ามี):", placeholder="พิมพ์ปัญหาหรืออุปสรรคที่พบในวันนี้ที่นี่...")
        suggestion_text = st.text_area("💡 ข้อเสนอแนะ (ถ้ามี):", placeholder="พิมพ์ข้อเสนอแนะเพิ่มเติมที่นี่...")
        
        # ประมวลผลการกรองข้อมูลตามเงื่อนไขที่เลือกทั้งหมด
        filtered_df = df.copy()
        if selected_date != "ทั้งหมด":
            filtered_df = filtered_df[filtered_df[target_col] == selected_date]
        if selected_doctor != "ทั้งหมด":
            filtered_df = filtered_df[filtered_df['แพทย์'].astype(str).str.strip() == selected_doctor]
        if selected_status != "ทั้งหมด":
            filtered_df = filtered_df[filtered_df['สถานะ'].astype(str).str.strip() == selected_status]
            
        # ==========================================
        # 📊 1. ส่วน Dashboard บนหน้าเว็บ
        # ==========================================
        if len(filtered_df) > 0:
            st.markdown("---")
            st.subheader("📊 สรุปข้อมูลเบื้องต้นก่อนพิมพ์")
            
            # คำนวณตัวเลข
            total_patients = len(filtered_df)
            new_patients = len(filtered_df[filtered_df['สถานะ'].astype(str).str.strip() == 'รายใหม่'])
            red_cases = sum(filtered_df['หน้า'].astype(str).str.contains("15P_z1gObqnm29vn-afAJ4JeMRQ4Y-IZw", na=False))
            
            dash_col1, dash_col2, dash_col3 = st.columns(3)
            dash_col1.metric("👥 ผู้ป่วยทั้งหมด (ตามตัวกรอง)", f"{total_patients} ราย")
            dash_col2.metric("🆕 ผู้ป่วยรายใหม่", f"{new_patients} ราย")
            dash_col3.metric("🚨 เคสเฝ้าระวัง (สีแดง)", f"{red_cases} ราย")
            
            # ==========================================
            # 📋 2. ตาราง Preview ข้อมูลบนเว็บ
            # ==========================================
            st.markdown("**📋 Preview รายชื่อผู้ป่วย:**")
            preview_cols = ['ชื่อ-สกุล', 'สถานะ', 'Dx', 'อาการปัจจุบัน', 'แพทย์']
            avail_cols = [c for c in preview_cols if c in filtered_df.columns]
            st.dataframe(filtered_df[avail_cols], use_container_width=True, hide_index=True)
            st.markdown("---")
            
            if st.button("🚀 สร้าง PDF", type="primary"):
                with st.spinner('กำลังประมวลผลข้อมูลและสร้างไฟล์ PDF...'):
                    
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
                            
                        # ไฮไลต์สีพื้นหลังแถวสำหรับเคสแดงและส้ม
                        row_bg_color = ""
                        if "15P_z1gObqnm29vn-afAJ4JeMRQ4Y-IZw" in face_url: # สีแดง
                            row_bg_color = "background-color: #fee2e2;" # ชมพูอ่อน
                        elif "1Vkl3jyY4W9h3Mv_l17xlWmbNw1A4p4-P" in face_url: # สีส้ม
                            row_bg_color = "background-color: #ffedd5;" # ส้มอ่อน
                            
                        if gdrive_id:
                            direct_img_url = f"https://drive.google.com/uc?id={gdrive_id}"
                            img_tag = f'<img src="{direct_img_url}" style="width:26px;height:26px;object-fit:cover;border-radius:4px; box-shadow: 0 1px 3px rgba(0,0,0,0.2);"/>'
                        else:
                            img_tag = "-"
                        
                        appt = str(row.get('นัด', '')).replace('nan', '')
                        doc = str(row.get('แพทย์', '')).replace('nan', '')
                        
                        # ใส่ style row_bg_color ลงไปใน tr
                        html_rows += f"""<tr style="{row_bg_color}">
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
                    # 📊 กราฟวงกลม Dx
                    # ==========================================
                    dx_counts = filtered_df['Dx'].value_counts()
                    chart_img_tag = ""
                    
                    valid_dx = {k: v for k, v in dx_counts.items() if str(k).lower() != 'nan'}
                    if valid_dx:
                        fig1, ax1 = plt.subplots(figsize=(2.8, 2.8))
                        total_dx = sum(valid_dx.values())
                        labels1 = [f"{k}\n{v} ({v/total_dx*100:.1f}%)" for k, v in valid_dx.items()]
                        wedges1, texts1 = ax1.pie(valid_dx.values(), startangle=90, colors=plt.cm.tab20.colors, radius=0.55)
                        
                        kw = dict(arrowprops=dict(arrowstyle="-", color="#64748b", lw=1.0), zorder=0, va="center")
                        for i, p in enumerate(wedges1):
                            ang = (p.theta2 - p.theta1)/2. + p.theta1
                            y = np.sin(np.deg2rad(ang))
                            x = np.cos(np.deg2rad(ang))
                            x_sign = -1 if x < 0 else 1
                            horizontalalignment = "right" if x_sign == -1 else "left"
                            kw["arrowprops"].update({"connectionstyle": f"angle,angleA=0,angleB={ang}"})
                            ax1.annotate(labels1[i], xy=(x*0.55, y*0.55), xytext=(0.75*x_sign, 0.8*y),
                                         horizontalalignment=horizontalalignment, fontsize=8, color='#111827', **kw) 
                        ax1.axis('equal') 
                        img_buf1 = io.BytesIO()
                        plt.savefig(img_buf1, format='png', bbox_inches='tight', transparent=True, dpi=120)
                        img_buf1.seek(0)
                        chart_img_tag = f'<img src="data:image/png;base64,{base64.b64encode(img_buf1.read()).decode("utf-8")}" style="width:100%; max-width:180px; display:block; margin:auto;"/>'
                        plt.close(fig1)
                    else:
                        chart_img_tag = "<p style='text-align:center; color:#94a3b8; font-size: 9pt;'>ไม่มีข้อมูล Dx</p>"

                    # ==========================================
                    # 📊 กราฟวงกลม สรุปเคสตามระดับสี
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
                        fig2, ax2 = plt.subplots(figsize=(2.2, 2.2))
                        colors2 = [color_map[k] for k in active_levels.keys()]
                        total_levels = sum(active_levels.values())
                        wedges2, texts2, autotexts2 = ax2.pie(
                            active_levels.values(), autopct=lambda p: f"{int(round(p * total_levels / 100))}\n({p:.1f}%)",
                            startangle=90, labeldistance=0.5, textprops={'fontsize': 8, 'color': '#111827', 'weight': 'normal', 'ha': 'center'}, 
                            colors=colors2, radius=0.85 
                        )
                        ax2.axis('equal')
                        img_buf2 = io.BytesIO()
                        plt.savefig(img_buf2, format='png', bbox_inches='tight', transparent=True, dpi=120)
                        img_buf2.seek(0)
                        level_chart_img_tag = f'<img src="data:image/png;base64,{base64.b64encode(img_buf2.read()).decode("utf-8")}" style="width:100%; max-width:130px; display:block; margin:auto;"/>'
                        plt.close(fig2)
                    else:
                        level_chart_img_tag = "<p style='text-align:center; color:#94a3b8; font-size: 9pt;'>ไม่มีข้อมูล</p>"

                    # ==========================================
                    # 📊 กราฟวงกลม สรุปเพศ
                    # ==========================================
                    gender_counts = {'ชาย': old_m + new_m, 'หญิง': old_f + new_f}
                    active_genders = {k: v for k, v in gender_counts.items() if v > 0}
                    gender_chart_img_tag = ""
                    if active_genders:
                        fig3, ax3 = plt.subplots(figsize=(2.2, 2.2))
                        colors3 = ['#3b82f6' if k == 'ชาย' else '#ec4899' for k in active_genders.keys()]
                        total_genders = sum(active_genders.values())
                        wedges3, texts3, autotexts3 = ax3.pie(
                            active_genders.values(), labels=active_genders.keys(), 
                            autopct=lambda p: f"{int(round(p * total_genders / 100))}\n({p:.1f}%)",
                            startangle=90, labeldistance=1.1, textprops={'fontsize': 8, 'color': '#111827', 'weight': 'normal', 'ha': 'center'}, 
                            colors=colors3, radius=0.75 
                        )
                        for autotext in autotexts3:
                            autotext.set_color('white')
                        ax3.axis('equal')
                        img_buf3 = io.BytesIO()
                        plt.savefig(img_buf3, format='png', bbox_inches='tight', transparent=True, dpi=120)
                        img_buf3.seek(0)
                        gender_chart_img_tag = f'<img src="data:image/png;base64,{base64.b64encode(img_buf3.read()).decode("utf-8")}" style="width:100%; max-width:110px; display:block; margin:auto; margin-top: 10px;"/>'
                        plt.close(fig3)
                    else:
                        gender_chart_img_tag = "<p style='text-align:center; color:#94a3b8; font-size: 9pt;'>ไม่มีข้อมูล</p>"

                    # ==========================================
                    # 🖼️ สร้าง HTML และ PDF
                    # ==========================================
                    logo_src = "https://drive.google.com/uc?id=1KYrHcRg6dvs2h0nfDf7ZxpzWpLnCqnjY"
                    title_text = f"รายงาน Telepsychiatry วันที่ {selected_date}" if selected_date != "ทั้งหมด" else "รายงาน Telepsychiatry (ทั้งหมด)"
                    
                    # อัปเดตหัวข้อตามตัวกรองแพทย์และสถานะ (ให้มีรายละเอียดครบ)
                    if selected_doctor != "ทั้งหมด":
                        title_text += f" (แพทย์: {selected_doctor})"
                    if selected_status != "ทั้งหมด":
                        title_text += f" (เฉพาะผู้ป่วย{selected_status})"

                    bottom_sections = ""
                    if problem_text.strip():
                        bottom_sections += f"""
                        <div style="background-color: #fef2f2; border-left: 5px solid #ef4444; padding: 12px; margin-top: 15px; border-radius: 4px; page-break-inside: avoid;">
                            <h3 style="color: #b91c1c; margin-top: 0; font-size: 10pt;">ปัญหา / อุปสรรค</h3>
                            <p style="margin: 0; line-height: 1.5; white-space: pre-line; font-size: 9pt;">{problem_text}</p>
                        </div>
                        """
                    if suggestion_text.strip():
                        bottom_sections += f"""
                        <div style="background-color: #f0fdf4; border-left: 5px solid #22c55e; padding: 12px; margin-top: 15px; border-radius: 4px; page-break-inside: avoid;">
                            <h3 style="color: #15803d; margin-top: 0; font-size: 10pt;">ข้อเสนอแนะ</h3>
                            <p style="margin: 0; line-height: 1.5; white-space: pre-line; font-size: 9pt;">{suggestion_text}</p>
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
                            size: A4 portrait; 
                            margin: 5mm 10mm 15mm 10mm; 
                            @bottom-right {{
                                content: "หน้า " counter(page) " / " counter(pages);
                                font-family: 'TH Sarabun PSK', 'Sarabun', sans-serif;
                                font-size: 9pt;
                                color: #64748b;
                            }}
                        }}
                        body {{ 
                            font-family: 'TH Sarabun PSK', 'Sarabun', sans-serif; 
                            font-size: 11pt; 
                            color: #334155; 
                            line-height: 1.5;
                        }}
                        .header-table {{ width: 100%; border-collapse: collapse; margin-bottom: 20px; border-bottom: 2px solid #cbd5e1; }}
                        .header-table td {{ border: none; padding-bottom: 10px; vertical-align: bottom; }}
                        .header-logo {{ width: 130px; text-align: center; }}
                        .header-logo img {{ width: 55px; height: auto; }}
                        .header-logo p {{ font-size: 9pt; font-weight: bold; color: #1e293b; margin-top: 5px; margin-bottom: 0; line-height: 1.1; }}
                        h1 {{ text-align: center; color: #0f172a; margin: 0; padding: 0; font-size: 12pt; }}
                        .summary-container {{ width: 100%; border-collapse: separate; border-spacing: 8px 0; margin-bottom: 20px; table-layout: fixed; }}
                        .summary-box {{ background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px; vertical-align: top; width: 33.33%; }}
                        .summary-box h3 {{ margin-top: 0; color: #0369a1; font-size: 10pt; border-bottom: 1px solid #cbd5e1; padding-bottom: 6px; margin-bottom: 8px; }}
                        
                        .data-table {{ width: 100%; border-collapse: collapse; margin-top: 10px; table-layout: auto; }}
                        .data-table th, .data-table td {{ padding: 6px 4px; vertical-align: middle; font-size: 9pt; }}
                        .data-table th {{ background-color: #1e293b; color: #ffffff; text-align: center; font-weight: 600; }}
                        .data-table td {{ border-bottom: 1px solid #e2e8f0; }}
                        
                        .data-table tr:nth-child(even) {{ background-color: #f8fafc; }}

                        .signature-section {{ margin-top: 30px; text-align: left; page-break-inside: avoid; }}
                    </style>
                    </head>
                    <body>
                        
                        <table class="header-table">
                            <tr>
                                <td class="header-logo"><img src="{logo_src}"><p>สถานพยาบาลเรือนจำ<br>จังหวัดบุรีรัมย์</p></td>
                                <td style="vertical-align: middle;"><h1>{title_text}</h1></td>
                                <td style="width: 130px;"></td> 
                            </tr>
                        </table>
                        
                        <table class="summary-container">
                            <tr>
                                <td class="summary-box" style="text-align: left;">
                                    <h3>สรุปสถานะผู้ป่วย</h3>
                                    <p style="margin: 0 0 5px 0; line-height: 1.5; font-size: 9pt;">
                                        <strong>ผู้ป่วยรายเก่า:</strong> {len(old_cases)} ราย (ช {old_m}, ญ {old_f})<br>
                                        <strong>ผู้ป่วยรายใหม่:</strong> {len(new_cases)} ราย (ช {new_m}, ญ {new_f})
                                    </p>
                                    {gender_chart_img_tag}
                                </td>
                                <td class="summary-box" style="text-align: center;"><h3>สรุปการวินิจฉัยโรค</h3>{chart_img_tag}</td>
                                <td class="summary-box" style="text-align: center;"><h3>สรุปเคสตามระดับสี</h3>{level_chart_img_tag}</td>
                            </tr>
                        </table>

                        <table class="data-table">
                            <thead>
                                <tr>
                                    <th style="width: 5%;">ที่</th>
                                    <th style="width: 16%;">ชื่อ-สกุล</th>
                                    <th style="width: 7%;">สถานะ</th>
                                    <th style="width: 9%;">Dx</th>
                                    <th style="width: 30%;">อาการปัจจุบัน</th>
                                    <th style="width: 6%;">ระดับ</th>
                                    <th style="width: 13%;">นัดครั้งถัดไป</th>
                                    <th style="width: 14%;">แพทย์</th>
                                </tr>
                            </thead>
                            <tbody>{html_rows}</tbody>
                        </table>

                        {bottom_sections}

                        <div class="signature-section">
                            <p style="margin-bottom: 20px; line-height: 1.6; font-size: 11pt;">เรียน ผู้บัญชาการเรือนจำฯ<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- เพื่อโปรดทราบ</p>
                            <br><br><br>
                            <div style="display: inline-block; text-align: center;">
                                <p style="margin: 0; font-size: 9pt;">(นางสาวเดือนนภา เบี้ยชาติไทย)</p>
                                <p style="margin: 5px 0 0 0; color: #475569; font-size: 9pt;">นักจิตวิทยาปฏิบัติการ</p>
                            </div>
                        </div>
                    </body>
                    </html>
                    """
                    
                    pdf_bytes = HTML(string=html_content).write_pdf()
                    
                    st.success(f"🎉 สร้าง PDF สำเร็จ! (ข้อมูล {len(filtered_df)} รายการ)")
                    st.download_button(
                        label="📥 คลิกที่นี่เพื่อดาวน์โหลดไฟล์ PDF",
                        data=pdf_bytes,
                        file_name=f"Telepsychiatry_Report.pdf",
                        mime="application/pdf"
                    )
        else:
             st.warning("ไม่พบข้อมูลผู้ป่วยในเงื่อนไขที่เลือก")
    else:
         st.error(f"❌ เกิดข้อผิดพลาด: ไม่พบคอลัมน์ชื่อ '{target_col}' ในไฟล์ Sheet ของคุณ")
