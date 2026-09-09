import streamlit as st
import pandas as pd
import base64
from weasyprint import HTML
import io

st.set_page_config(page_title="ระบบสร้างรายงาน PDF", page_icon="📄", layout="wide")

st.title("📄 ระบบสร้างรายงานสุขภาพจิต (PDF)")
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
    
    # ตรวจสอบว่ามีคอลัมน์ 1 อยู่ในข้อมูลจริงๆ หรือไม่
    if target_col in df.columns:
        
        # จัดการข้อมูลวันที่ให้สะอาดและดึงวันที่ที่ไม่ซ้ำกันมาแสดงใน Dropdown
        df[target_col] = df[target_col].astype(str).str.strip()
        unique_dates = ["ทั้งหมด"] + sorted(list(set([d for d in df[target_col].unique() if d and str(d).lower() != 'nan'])))
        
        # สร้าง Dropdown ให้เลือกวันที่
        selected_date = st.selectbox("📅 เลือกวันที่ (จาก คอลัมน์ 1):", options=unique_dates)
        
        if st.button("🚀 สร้าง PDF", type="primary"):
            with st.spinner('กำลังประมวลผลข้อมูลและสร้างไฟล์ PDF...'):
                
                # กรองข้อมูล
                filtered_df = df.copy()
                if selected_date != "ทั้งหมด":
                    filtered_df = filtered_df[filtered_df[target_col] == selected_date]
                
                if len(filtered_df) == 0:
                    st.warning("ไม่พบข้อมูลในวันที่เลือก")
                else:
                    html_rows = ""
                    for idx, row in filtered_df.iterrows():
                        no = idx + 1
                        name = str(row.get('ชื่อ-สกุล', '')).replace('nan', '')
                        status = str(row.get('สถานะ', '')).replace('nan', '')
                        dx = str(row.get('Dx', '')).replace('nan', '')
                        symptom = str(row.get('อาการปัจจุบัน', '')).replace('nan', '')
                        
                        face_url = str(row.get('หน้า', ''))
                        if "1YlAPW2PBMUbkuRt0unWjJTolQ9aAp48Y" in face_url:
                            color = "#4CAF50" # ปกติ (เขียว)
                        elif "1Vkl3jyY4W9h3Mv_l17xlWmbNw1A4p4-P" in face_url:
                            color = "#FFEB3B" # เฝ้าระวัง (เหลือง)
                        elif "15P_z1gObqnm29vn-afAJ4JeMRQ4Y-IZw" in face_url:
                            color = "#F44336" # รุนแรง (แดง)
                        elif "1Wu3vMN2idLhA5fWlY4ZsGZ64Uf_c-f-B" in face_url:
                            color = "#2196F3" # ฟ้า
                        else:
                            color = "#9E9E9E" # เทา
                            
                        svg = f'''<svg width="24" height="24" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><circle cx="12" cy="12" r="10" fill="{color}" stroke="#333" stroke-width="1"/></svg>'''
                        b64_svg = base64.b64encode(svg.encode('utf-8')).decode('utf-8')
                        img_tag = f'<img src="data:image/svg+xml;base64,{b64_svg}" style="width:20px;height:20px;"/>'
                        
                        appt = str(row.get('นัด', '')).replace('nan', '')
                        doc = str(row.get('แพทย์', '')).replace('nan', '')
                        
                        html_rows += f"""<tr>
                            <td style="text-align:center;">{no}</td><td>{name}</td><td>{status}</td>
                            <td>{dx}</td><td>{symptom}</td><td style="text-align:center;">{img_tag}</td>
                            <td>{appt}</td><td>{doc}</td>
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
                    dx_html = "".join([f"<li><strong>{k}</strong>: {v} ราย</li>" for k, v in dx_counts.items() if str(k).lower() != 'nan'])

                    title_text = f"รายงานฐานข้อมูลสุขภาพจิต (วันที่ {selected_date})" if selected_date != "ทั้งหมด" else "รายงานฐานข้อมูลสุขภาพจิต (ทั้งหมด)"

                    html_content = f"""
                    <!DOCTYPE html>
                    <html lang="th">
                    <head>
                    <meta charset="UTF-8">
                    <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700&display=swap" rel="stylesheet">
                    <style>
                        @page {{ size: A4 landscape; margin: 15mm; }}
                        body {{ font-family: 'Sarabun', sans-serif; font-size: 11pt; }}
                        table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
                        th, td {{ border: 1px solid #bdc3c7; padding: 6px; }}
                        th {{ background-color: #34495e; color: white; text-align: center; }}
                        .summary-box {{ border: 1px solid #ddd; padding: 15px; border-radius: 8px; }}
                        .summary-box h3 {{ margin-top: 0; color: #2980b9; }}
                    </style>
                    </head>
                    <body>
                        <h1 style="text-align:center;">{title_text}</h1>
                        <div class="summary-box">
                            <table style="border: none; margin-top:0;">
                                <tr style="border: none;">
                                    <td style="border: none; vertical-align: top; width: 50%;">
                                        <h3>สรุปสถานะผู้ป่วย</h3>
                                        <ul>
                                            <li><strong>รายเก่า:</strong> {len(old_cases)} ราย (ชาย {old_m}, หญิง {old_f})</li>
                                            <li><strong>รายใหม่:</strong> {len(new_cases)} ราย (ชาย {new_m}, หญิง {new_f})</li>
                                        </ul>
                                    </td>
                                    <td style="border: none; vertical-align: top; width: 50%;">
                                        <h3>สรุปการวินิจฉัยโรค (Dx)</h3>
                                        <ul style="column-count: 2;">{dx_html}</ul>
                                    </td>
                                </tr>
                            </table>
                        </div>
                        <table>
                            <thead>
                                <tr>
                                    <th style="width: 5%;">ลำดับ</th><th style="width: 15%;">ชื่อ-สกุล</th>
                                    <th style="width: 8%;">สถานะ</th><th style="width: 10%;">Dx</th>
                                    <th style="width: 32%;">อาการปัจจุบัน</th><th style="width: 5%;">ระดับ</th>
                                    <th style="width: 10%;">นัด</th><th style="width: 15%;">แพทย์</th>
                                </tr>
                            </thead>
                            <tbody>{html_rows}</tbody>
                        </table>
                    </body>
                    </html>
                    """
                    
                    pdf_bytes = HTML(string=html_content).write_pdf()
                    
                    st.success(f"สร้าง PDF สำเร็จ! (ข้อมูล {len(filtered_df)} รายการ)")
                    st.download_button(
                        label="📥 คลิกที่นี่เพื่อดาวน์โหลดไฟล์ PDF",
                        data=pdf_bytes,
                        file_name=f"Mental_Health_Report_{selected_date.replace('/', '-') if selected_date != 'ทั้งหมด' else 'All'}.pdf",
                        mime="application/pdf"
                    )
    else:
         st.error(f"❌ เกิดข้อผิดพลาด: ไม่พบคอลัมน์ชื่อ '{target_col}' ในไฟล์ Sheet ของคุณ")
