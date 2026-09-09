import streamlit as st
import pandas as pd
import base64
from weasyprint import HTML
import io
import re

st.set_page_config(page_title="ระบบสร้างรายงาน PDF", page_icon="📄", layout="wide")

st.title("📄 ระบบสร้างรายงาน Telepsychiatry (PDF)")
st.markdown("ดึงข้อมูลจาก Google Sheet และสรุปเป็น PDF")

# URL ของ Google Sheet (รูปแบบส่งออกเป็น CSV)
SHEET_URL = "https://docs.google.com/spreadsheets/d/1dtpMxycg0en1_zdtsQreeLohqOVEqNyZyxCI26O5Zlc/export?format=csv"

@st.cache_data(ttl=60) # เก็บ Cache 1 นาที
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
    
    # 📌 บังคับให้ใช้ "คอลัมน์ 1" สำหรับกรองวันที่
    target_col = "คอลัมน์ 1"
    
    if target_col in df.columns:
        df[target_col] = df[target_col].astype(str).str.strip()
        unique_dates = ["ทั้งหมด"] + sorted(list(set([d for d in df[target_col].unique() if d and str(d).lower() != 'nan'])))
        
        selected_date = st.selectbox("📅 เลือกวันที่ (จาก คอลัมน์ 1):", options=unique_dates)
        
        # 📝 ช่องสำหรับพิมพ์ปัญหา/อุปสรรค และข้อเสนอแนะ
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
                        if match1:
                            gdrive_id = match1.group(1)
                        elif match2:
                            gdrive_id = match2.group(1)
                            
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

                    dx_counts = filtered_df['Dx'].value_counts()
                    dx_html = "".join([f"<li style='margin-bottom:4px;'><strong>{k}</strong>: {v} ราย</li>" for k, v in dx_counts.items() if str(k).lower() != 'nan'])

                    title_text = f"รายงาน Telepsychiatry วันที่ {selected_date}" if selected_date != "ทั้งหมด" else "รายงาน Telepsychiatry (ทั้งหมด)"

                    # ส่วนของปัญหา/อุปสรรค และข้อเสนอแนะ (ดีไซน์ใหม่)
                    bottom_sections = ""
                    if problem_text.strip():
                        bottom_sections += f"""
                        <div style="background-color: #fef2f2; border-left: 5px solid #ef4444; padding: 15px; margin-top: 20px; border-radius: 4px;">
                            <h3 style="color: #b91c1c; margin-top: 0; font-size: 13pt;">⚠️ ปัญหา / อุปสรรค</h3>
                            <p style="margin: 0; line-height: 1.6; white-space: pre-line;">{problem_text}</p>
                        </div>
                        """
                    if suggestion_text.strip():
                        bottom_sections += f"""
                        <div style="background-color: #f0fdf4; border-left: 5px solid #22c55e; padding: 15px; margin-top: 15px; border-radius: 4px;">
                            <h3 style="color: #15803d; margin-top: 0; font-size: 13pt;">💡 ข้อเสนอแนะ</h3>
                            <p style="margin: 0; line-height: 1.6; white-space: pre-line;">{suggestion_text}</p>
                        </div>
                        """

                    # HTML หลัก พร้อม CSS ที่ออกแบบใหม่
                    html_content = f"""
                    <!DOCTYPE html>
                    <html lang="th">
                    <head>
                    <meta charset="UTF-8">
                    <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap" rel="stylesheet">
                    <style>
                        @page {{ size: A4 landscape; margin: 15mm; }}
                        body {{ 
                            font-family: 'Sarabun', sans-serif; 
                            font-size: 11pt; 
                            color: #334155; 
                            line-height: 1.5;
                        }}
                        h1 {{ 
                            text-align: center; 
                            color: #0f172a; 
                            border-bottom: 2px solid #cbd5e1; 
                            padding-bottom: 10px; 
                            margin-bottom: 25px;
                            font-size: 20pt;
                        }}
                        
                        /* Layout สำหรับกล่องสรุป */
                        .summary-container {{
                            width: 100%;
                            border-collapse: separate;
                            border-spacing: 15px 0; /* ระยะห่างระหว่างกล่องซ้ายขวา */
                            margin-bottom: 25px;
                        }}
                        .summary-box {{
                            background-color: #f8fafc;
                            border: 1px solid #e2e8f0;
                            border-radius: 8px;
                            padding: 15px 20px;
                            vertical-align: top;
                            width: 50%;
                        }}
                        .summary-box h3 {{
                            margin-top: 0; 
                            color: #0369a1; 
                            font-size: 13pt;
                            border-bottom: 1px solid #cbd5e1;
                            padding-bottom: 8px;
                            margin-bottom: 12px;
                        }}
                        .summary-box ul {{
                            margin: 0;
                            padding-left: 20px;
                        }}

                        /* ดีไซน์ตาราง */
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
                            font-size: 11pt;
                        }}
                        .data-table td {{
                            border-bottom: 1px solid #e2e8f0;
                        }}
                        .data-table tr:nth-child(even) {{ 
                            background-color: #f8fafc; 
                        }}

                        /* ส่วนลงนาม */
                        .signature-section {{ 
                            margin-top: 40px; 
                            text-align: left; 
                            page-break-inside: avoid; 
                        }}
                    </style>
                    </head>
                    <body>
                        <h1>{title_text}</h1>
                        
                        <table class="summary-container" style="margin-left: -15px; margin-right: -15px; width: calc(100% + 30px);">
                            <tr>
                                <td class="summary-box">
                                    <h3>📊 สรุปสถานะผู้ป่วย</h3>
                                    <ul>
                                        <li style="margin-bottom:4px;"><strong>ผู้ป่วยรายเก่า:</strong> {len(old_cases)} ราย (ชาย {old_m}, หญิง {old_f})</li>
                                        <li><strong>ผู้ป่วยรายใหม่:</strong> {len(new_cases)} ราย (ชาย {new_m}, หญิง {new_f})</li>
                                    </ul>
                                </td>
                                <td class="summary-box">
                                    <h3>🩺 สรุปการวินิจฉัยโรค (Dx)</h3>
                                    <ul style="column-count: 2; column-gap: 20px;">
                                        {dx_html}
                                    </ul>
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
                                <p style="margin: 0; font-weight: bold; font-size: 12pt;">(นางสาวเดือนนภา เบี้ยชาติไทย)</p>
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
