import streamlit as st
import pandas as pd
import io
import os
import urllib.request
from datetime import datetime
from fpdf import FPDF
from streamlit_gsheets import GSheetsConnection
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

# 1. ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="ระบบบันทึกข้อมูลคลินิกคลายเครียด", layout="wide")

FONT_URL = "https://github.com/google/fonts/raw/main/ofl/sarabun/Sarabun-Regular.ttf"
FONT_PATH = "Sarabun-Regular.ttf"

@st.cache_resource
def download_font():
    if not os.path.exists(FONT_PATH):
        urllib.request.urlretrieve(FONT_URL, FONT_PATH)

download_font() 

# ================== เชื่อมต่อ Google Sheets ==================
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1wMxwcdcF3zXliTTifINkhinh76VYBh-xSj_7GLLqDxY/edit?usp=sharing"
INMATE_DB_URL = "https://docs.google.com/spreadsheets/d/1dtpMxycg0en1_zdtsQreeLohqOVEqNyZyxCI26O5Zlc/edit?usp=drivesdk"

conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    try:
        df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet=0, ttl=0)
        df = df.fillna("")
        
        if df.empty or len(df.columns) == 0:
            df = pd.DataFrame(columns=[
                "วันที่", "ชื่อ-สกุล", "เพศ", "อายุ", "ห้อง", "คดี", "ครั้งที่", 
                "ประเภทบริการ", "ผลประเมิน", "บันทึกอาการ", "ปัญหาและอุปสรรค", "สถานะติดตาม", "วันที่นัดติดตาม"
            ])
            conn.update(spreadsheet=SPREADSHEET_URL, worksheet=0, data=df)
        else:
            needs_update = False
            
            if 'ชื่อ' in df.columns and 'นามสกุล' in df.columns:
                df['ชื่อ-สกุล'] = df['ชื่อ'].astype(str) + " " + df['นามสกุล'].astype(str)
                df['ชื่อ-สกุล'] = df['ชื่อ-สกุล'].str.strip()
                df = df.drop(columns=['ชื่อ', 'นามสกุล'])
                needs_update = True
                
            if 'แดน/ห้อง' in df.columns:
                df = df.rename(columns={'แดน/ห้อง': 'ห้อง'})
                needs_update = True
                
            if 'บันทึกติดตาม' in df.columns:
                df = df.rename(columns={'บันทึกติดตาม': 'บันทึกอาการ'})
                needs_update = True
                
            cols = ["วันที่", "ชื่อ-สกุล", "เพศ", "อายุ", "ห้อง", "คดี", "ครั้งที่", "ประเภทบริการ", "ผลประเมิน", "บันทึกอาการ", "ปัญหาและอุปสรรค", "สถานะติดตาม", "วันที่นัดติดตาม"]
            for c in cols:
                if c not in df.columns:
                    df[c] = ""
                    needs_update = True

            if needs_update:
                df = df[cols]
                conn.update(spreadsheet=SPREADSHEET_URL, worksheet=0, data=df)
                
        return df
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการเชื่อมต่อ Google Sheets หลัก: {e}")
        return pd.DataFrame(columns=[
            "วันที่", "ชื่อ-สกุล", "เพศ", "อายุ", "ห้อง", "คดี", "ครั้งที่", 
            "ประเภทบริการ", "ผลประเมิน", "บันทึกอาการ", "ปัญหาและอุปสรรค", "สถานะติดตาม", "วันที่นัดติดตาม"
        ])

if 'patient_data' not in st.session_state:
    st.session_state['patient_data'] = load_data()

st.title("🏥 ระบบบันทึกข้อมูลคลินิกคลายเครียด (สจ.21)")
st.markdown("ระบบออนไลน์ เชื่อมต่อฐานข้อมูล Cloud (ข้อมูลปลอดภัย 100%)")

tab1, tab2 = st.tabs(["📝 บันทึกข้อมูลรายบุคคล", "📊 สรุปและออกรายงาน สจ.21"])

SYMPTOM_OPTIONS = ["เครียด", "เหนื่อยล้า", "ซึมเศร้า", "อยากตาย", "นอนไม่หลับ", "วิตกกังวล", "ปวดศีรษะ", "อื่นๆ (ให้ระบุ)"]
STATUS_OPTIONS = ["รอดำเนินการ", "ติดตามแล้ว", "ติดตามต่อ", "ใกล้พ้นโทษ", "ปิดเคส"]

# ================= TAB 1: บันทึกข้อมูล =================
with tab1:
    
    # ---- ส่วนดึงข้อมูล Auto-fill ----
    st.subheader("📥 ดึงข้อมูลผู้ต้องขัง (Auto-fill)")
    
    def_name, def_gender, def_age, def_case = "", "ชาย", 30, ""
    
    try:
        inmate_df = conn.read(spreadsheet=INMATE_DB_URL, worksheet=0, ttl=10)
        inmate_df = inmate_df.fillna("")
        
        if not inmate_df.empty and len(inmate_df.columns) > 0:
            # พยายามหาคอลัมน์ แดน หรือ ห้อง เป็นค่าเริ่มต้น
            default_filter_col = inmate_df.columns[0]
            for col in inmate_df.columns:
                if "แดน" in str(col) or "ห้อง" in str(col):
                    default_filter_col = col
                    break
            
            # ให้ผู้ใช้เลือกคอลัมน์สำหรับกรองได้เอง
            filter_col = st.selectbox("1️⃣ เลือกคอลัมน์สำหรับกรองข้อมูล:", inmate_df.columns.tolist(), index=inmate_df.columns.tolist().index(default_filter_col))
            
            unique_vals = [v for v in inmate_df[filter_col].unique() if str(v).strip() != ""]
            options_1 = ["-- กรุณาเลือกข้อมูล --"] + unique_vals
            
            selected_val = st.selectbox(f"👉 เลือกข้อมูลจาก [{filter_col}]:", options_1)
            
            if selected_val != "-- กรุณาเลือกข้อมูล --":
                filtered_inmates = inmate_df[inmate_df[filter_col].astype(str) == str(selected_val)]
                
                # พยายามหาคอลัมน์ชื่อ
                name_col = filtered_inmates.columns[0]
                for col in filtered_inmates.columns:
                    if "ชื่อ" in str(col):
                        name_col = col
                        break
                
                options_2 = ["-- กรุณาพิมพ์หรือเลือกรายชื่อ --"] + filtered_inmates[name_col].astype(str).tolist()
                
                selected_inmate = st.selectbox("2️⃣ ค้นหาและเลือกรายชื่อผู้ต้องขัง:", options_2)
                
                if selected_inmate != "-- กรุณาพิมพ์หรือเลือกรายชื่อ --":
                    row = filtered_inmates[filtered_inmates[name_col] == selected_inmate].iloc[0]
                    def_name = str(row[name_col])
                    
                    if "เพศ" in row.index and pd.notna(row["เพศ"]) and str(row["เพศ"]) != "":
                        def_gender = "หญิง" if "หญิง" in str(row["เพศ"]) else "ชาย"
                        
                    if "อายุ" in row.index and pd.notna(row["อายุ"]) and str(row["อายุ"]) != "":
                        try: def_age = int(float(row["อายุ"]))
                        except ValueError: def_age = 30
                            
                    if "คดี" in row.index and pd.notna(row["คดี"]):
                        def_case = str(row["คดี"])
        else:
            st.info("💡 ไม่พบคอลัมน์ข้อมูลในชีตทะเบียนผู้ต้องขัง")
            
    except Exception as e:
        st.error(f"⚠️ ไม่สามารถดึงข้อมูลจากชีตทะเบียนได้ (ตรวจสอบการแชร์ไฟล์ให้ Email Bot หรือลิงก์) Error: {e}")

    # ---- ส่วนฟอร์มกรอกข้อมูล ----
    with st.form("patient_form", clear_on_submit=True):
        st.subheader("📝 บันทึกข้อมูลเข้ารับบริการ")
        col1, col2 = st.columns(2)
        
        with col1:
            full_name = st.text_input("ชื่อ-สกุล", value=def_name, placeholder="เช่น สมชาย มั่นคง")
            gender_index = 0 if def_gender == "ชาย" else 1
            gender = st.radio("เพศ", ["ชาย", "หญิง"], index=gender_index, horizontal=True)
            age = st.number_input("อายุ (ปี)", min_value=15, max_value=100, step=1, value=def_age)
            room = st.text_input("ห้อง")
            
            symptoms = st.multiselect("บันทึกอาการ (เลือกได้หลายข้อ)", SYMPTOM_OPTIONS)
            other_symptom = ""
            if "อื่นๆ (ให้ระบุ)" in symptoms:
                other_symptom = st.text_input("ระบุอาการอื่นๆ")
                
            obs_problem = st.text_input("ปัญหาและอุปสรรค (ถ้ามี)")
            
        with col2:
            case_type = st.text_input("ฐานความผิด / คดี", value=def_case)
            visit_count = st.number_input("รับบริการครั้งที่", min_value=1, step=1)
            
            service_type = st.multiselect("ประเภทการเข้ารับบริการ (ตาม สจ.21)", [
                "คัดกรองผู้ต้องขังเข้าใหม่", "ผู้ต้องขังรายเก่าที่ได้รับการคัดกรองซ้ำ",
                "ให้บริการตรวจรักษาในเรือนจำ", "ตรวจผ่านระบบ Telepsychiatry",
                "การบริการคลินิกคลายเครียดให้การปรึกษา", "จิตบำบัด", "เฝ้าระวังผู้ต้องขังมีพฤติกรรมเสี่ยงฆ่าตัวตาย",
                "ฆ่าตัวตายไม่สำเร็จ", "ฆ่าตัวตายสำเร็จ", "อบรมผู้ต้องขังช่วยเหลืองานด้านสุขภาพจิต"
            ])
            
            result = st.multiselect("ผลการประเมิน / การดำเนินการ", [
                "ปกติ", "พบความผิดปกติ", "ผู้ที่พบปัญหาสุขภาพจิตและได้รับการดูแลรักษา",
                "รับการประเมินเพื่อวินิจฉัยโรคทางจิตเวช", "ส่งต่อไปรับการรักษานอกเรือนจำ",
                "โรงพยาบาลรับเป็นผู้ป่วยใน (admit)", "ปิดเคส"
            ])
            
            st.info("💡 **Auto-fill:** \n- หากเลือก **'พบความผิดปกติ'** ระบบจะเพิ่ม Telepsychiatry และการวินิจฉัยให้อัตโนมัติ \n- หากเลือก **'จิตบำบัด'** ระบบจะเพิ่ม คลินิกคลายเครียด ให้อัตโนมัติ")
        
        if st.form_submit_button("💾 บันทึกข้อมูลใหม่"):
            if full_name.strip():
                
                if "พบความผิดปกติ" in result:
                    if "ผู้ที่พบปัญหาสุขภาพจิตและได้รับการดูแลรักษา" not in result:
                        result.append("ผู้ที่พบปัญหาสุขภาพจิตและได้รับการดูแลรักษา")
                    if "รับการประเมินเพื่อวินิจฉัยโรคทางจิตเวช" not in result:
                        result.append("รับการประเมินเพื่อวินิจฉัยโรคทางจิตเวช")
                    
                    if "ตรวจผ่านระบบ Telepsychiatry" not in service_type:
                        service_type.append("ตรวจผ่านระบบ Telepsychiatry")
                        
                initial_status = "รอดำเนินการ"
                if "ปิดเคส" in result:
                    initial_status = "ปิดเคส"
                
                if "จิตบำบัด" in service_type:
                    if "การบริการคลินิกคลายเครียดให้การปรึกษา" not in service_type:
                        service_type.append("การบริการคลินิกคลายเครียดให้การปรึกษา")

                service_type_str = ", ".join(service_type) if service_type else "ไม่ได้ระบุ"
                result_str = ", ".join(result) if result else "ไม่ได้ระบุ"
                
                final_symptoms = [s for s in symptoms if s != "อื่นๆ (ให้ระบุ)"]
                if other_symptom.strip():
                    final_symptoms.append(other_symptom.strip())
                symptom_str = ", ".join(final_symptoms)
                
                formatted_date = datetime.now().strftime("%d/%m/%Y")
                
                new_data = {
                    "วันที่": formatted_date,
                    "ชื่อ-สกุล": full_name.strip(), 
                    "เพศ": gender, "อายุ": age,
                    "ห้อง": room, "คดี": case_type, "ครั้งที่": visit_count,
                    "ประเภทบริการ": service_type_str, "ผลประเมิน": result_str,
                    "บันทึกอาการ": symptom_str, "ปัญหาและอุปสรรค": obs_problem.strip(),
                    "สถานะติดตาม": initial_status, "วันที่นัดติดตาม": ""
                }
                
                new_df = pd.DataFrame([new_data])
                st.session_state['patient_data'] = pd.concat([st.session_state['patient_data'], new_df], ignore_index=True)
                
                conn.update(spreadsheet=SPREADSHEET_URL, worksheet=0, data=st.session_state['patient_data'])
                st.success(f"✅ บันทึกข้อมูลของ {full_name} ขึ้นฐานข้อมูลสำเร็จ!")
            else:
                st.error("⚠️ กรุณากรอกชื่อ-สกุล")

    # ================= ส่วนตารางที่แก้ไขได้ =================
    st.markdown("---")
    st.subheader("📋 ตารางข้อมูลปัจจุบัน (สามารถแก้ไขข้อมูลในตารางได้โดยตรง)")
    st.info("💡 **วิธีใช้งาน:** ดับเบิลคลิกที่ช่องเพื่อพิมพ์แก้ไขหรือเลือก Drop-down และสามารถคลิกเลือกแถวเพื่อลบข้อมูลได้ เมื่อแก้เสร็จแล้วให้กดปุ่ม **'ยืนยันการแก้ไข'** ด้านล่าง")

    df_current = st.session_state['patient_data']
    
    if not df_current.empty:
        edited_df = st.data_editor(
            df_current, 
            num_rows="dynamic", 
            use_container_width=True,
            column_config={
                "เพศ": st.column_config.SelectboxColumn("เพศ", options=["ชาย", "หญิง"]),
                "ประเภทบริการ": st.column_config.SelectboxColumn("ประเภทบริการ", options=[
                    "คัดกรองผู้ต้องขังเข้าใหม่", "ผู้ต้องขังรายเก่าที่ได้รับการคัดกรองซ้ำ",
                    "ให้บริการตรวจรักษาในเรือนจำ", "ตรวจผ่านระบบ Telepsychiatry",
                    "การบริการคลินิกคลายเครียดให้การปรึกษา", "จิตบำบัด", "เฝ้าระวังผู้ต้องขังมีพฤติกรรมเสี่ยงฆ่าตัวตาย",
                    "ฆ่าตัวตายไม่สำเร็จ", "ฆ่าตัวตายสำเร็จ", "อบรมผู้ต้องขังช่วยเหลืองานด้านสุขภาพจิต"
                ]),
                "ผลประเมิน": st.column_config.SelectboxColumn("ผลประเมิน", options=[
                    "ปกติ", "พบความผิดปกติ", "ผู้ที่พบปัญหาสุขภาพจิตและได้รับการดูแลรักษา",
                    "รับการประเมินเพื่อวินิจฉัยโรคทางจิตเวช", "ส่งต่อไปรับการรักษานอกเรือนจำ", "โรงพยาบาลรับเป็นผู้ป่วยใน (admit)", "ปิดเคส"
                ]),
                "สถานะติดตาม": st.column_config.SelectboxColumn("สถานะติดตาม", options=STATUS_OPTIONS)
            }
        )
        
        if st.button("💾 ยืนยันการแก้ไขและบันทึกลงฐานข้อมูล"):
            edited_df = edited_df.reset_index(drop=True) 
            st.session_state['patient_data'] = edited_df
            conn.update(spreadsheet=SPREADSHEET_URL, worksheet=0, data=edited_df)
            st.success("✅ บันทึกการแก้ไขทั้งหมดลง Google Sheets เรียบร้อยแล้ว!")
            st.rerun()
    else:
        st.dataframe(df_current, use_container_width=True)

    # ================= ส่วนแจ้งเตือนติดตามผู้ป่วย =================
    st.markdown("---")
    st.subheader("🔔 แจ้งเตือนเคสที่ต้องติดตามต่อ / ใกล้พ้นโทษ")
    main_db = st.session_state['patient_data']
    if not main_db.empty:
        to_follow_up = main_db[main_db['สถานะติดตาม'].isin(['ติดตามต่อ', 'ใกล้พ้นโทษ'])]
        if not to_follow_up.empty:
            for idx, row in to_follow_up.iterrows():
                date_str = row['วันที่นัดติดตาม'] if pd.notna(row['วันที่นัดติดตาม']) and str(row['วันที่นัดติดตาม']).strip() != "" else "ไม่ได้ระบุวัน"
                if row['สถานะติดตาม'] == 'ใกล้พ้นโทษ':
                    st.info(f"⏳ **เคสใกล้พ้นโทษ:** {row.get('ชื่อ-สกุล', '')} (ห้อง: {row.get('ห้อง', '')}) — สถานะ: **ใกล้พ้นโทษ**")
                else:
                    st.warning(f"📅 **นัดติดตามอาการ:** {row.get('ชื่อ-สกุล', '')} (ห้อง: {row.get('ห้อง', '')}) — นัดหมายวันที่: **{date_str}**")
        else:
            st.info("🎉 ปัจจุบันไม่มีเคสที่ค้างการติดตามหรือใกล้พ้นโทษ")

    # ================= ส่วนระบบติดตามผู้ป่วย (Follow-up) =================
    st.markdown("---")
    st.subheader("🚨 ระบบอัปเดตสถานะผู้ป่วย")
    
    if not main_db.empty:
        abnormal_patients = main_db[(main_db['ผลประเมิน'].astype(str).str.contains("พบความผิดปกติ", na=False)) & (main_db['สถานะติดตาม'] != "บันทึกครั้งใหม่แล้ว")]
        if not abnormal_patients.empty:
            for idx, row in abnormal_patients.iterrows():
                if row['สถานะติดตาม'] == "ปิดเคส":
                    status_icon = "🟢"
                elif row['สถานะติดตาม'] == "ใกล้พ้นโทษ":
                    status_icon = "⏳"
                elif row['สถานะติดตาม'] == "ติดตามต่อ":
                    status_icon = "🟡"
                else:
                    status_icon = "🔴"
                
                with st.expander(f"{status_icon} อัปเดตอาการ: {row.get('ชื่อ-สกุล', '')} [สถานะ: {row['สถานะติดตาม']}]"):
                    st.write(f"**วันที่รับบริการล่าสุด:** {row['วันที่']} | **เข้ารับบริการครั้งที่:** {row['ครั้งที่']}")
                    
                    valid_srv_options = ["คัดกรองผู้ต้องขังเข้าใหม่", "ผู้ต้องขังรายเก่าที่ได้รับการคัดกรองซ้ำ", "ให้บริการตรวจรักษาในเรือนจำ", "ตรวจผ่านระบบ Telepsychiatry", "การบริการคลินิกคลายเครียดให้การปรึกษา", "จิตบำบัด", "เฝ้าระวังผู้ต้องขังมีพฤติกรรมเสี่ยงฆ่าตัวตาย", "ฆ่าตัวตายไม่สำเร็จ", "ฆ่าตัวตายสำเร็จ", "อบรมผู้ต้องขังช่วยเหลืองานด้านสุขภาพจิต"]
                    old_srv = [s.strip() for s in str(row.get('ประเภทบริการ', '')).split(",")]
                    new_service_list = st.multiselect("ประเภทการเข้ารับบริการ (ครั้งนี้)", valid_srv_options, default=[s for s in old_srv if s in valid_srv_options], key=f"srv_{idx}")
                    
                    valid_res_options = ["ปกติ", "พบความผิดปกติ", "ผู้ที่พบปัญหาสุขภาพจิตและได้รับการดูแลรักษา", "รับการประเมินเพื่อวินิจฉัยโรคทางจิตเวช", "ส่งต่อไปรับการรักษานอกเรือนจำ", "โรงพยาบาลรับเป็นผู้ป่วยใน (admit)", "ปิดเคส"]
                    old_res = [r.strip() for r in str(row.get('ผลประเมิน', '')).split(",")]
                    new_result_list = st.multiselect("ผลประเมิน (อัปเดตล่าสุด)", valid_res_options, default=[r for r in old_res if r in valid_res_options], key=f"res_{idx}")
                    
                    st.caption("💡 ระบบจะเพิ่ม คลินิกคลายเครียด และ Telepsychiatry ให้อัตโนมัติตามเงื่อนไขเมื่อกดบันทึก")
                    
                    current_status = row['สถานะติดตาม'] if pd.notna(row['สถานะติดตาม']) and row['สถานะติดตาม'] != "" else "รอดำเนินการ"
                    new_status = st.selectbox("สถานะการติดตาม", STATUS_OPTIONS, index=STATUS_OPTIONS.index(current_status) if current_status in STATUS_OPTIONS else 0, key=f"status_{idx}")
                    
                    new_date_str = row['วันที่นัดติดตาม']
                    if new_status == "ติดตามต่อ":
                        parsed_date = datetime.strptime(str(row['วันที่นัดติดตาม']), "%Y-%m-%d").date() if pd.notna(row['วันที่นัดติดตาม']) and str(row['วันที่นัดติดตาม']).strip() != "" else datetime.now().date()
                        new_date_str = st.date_input("ระบุวันที่นัดติดตามครั้งต่อไป", value=parsed_date, key=f"date_{idx}").strftime("%Y-%m-%d")
                    else:
                        new_date_str = "" 
                    
                    existing_notes_raw = str(row.get('บันทึกอาการ', ''))
                    existing_notes_list = [s.strip() for s in existing_notes_raw.split(",")] if existing_notes_raw else []
                    
                    default_symp = []
                    other_existing_list = []
                    for s in existing_notes_list:
                        if s in SYMPTOM_OPTIONS:
                            default_symp.append(s)
                        elif s:
                            other_existing_list.append(s)
                            
                    if other_existing_list and "อื่นๆ (ให้ระบุ)" not in default_symp:
                        default_symp.append("อื่นๆ (ให้ระบุ)")
                        
                    other_existing_str = ", ".join(other_existing_list)
                    
                    selected_symptoms = st.multiselect("บันทึกอาการ (อัปเดตล่าสุด):", SYMPTOM_OPTIONS, default=default_symp, key=f"symp_{idx}")
                    
                    other_note = ""
                    if "อื่นๆ (ให้ระบุ)" in selected_symptoms:
                        other_note = st.text_input("ระบุอาการอื่นๆ / รายละเอียดเพิ่มเติม:", value=other_existing_str, key=f"other_note_{idx}")
                        
                    new_prob = st.text_input("ปัญหาและอุปสรรคที่พบ:", value=row.get('ปัญหาและอุปสรรค', ''), key=f"prob_{idx}")
                    
                    next_visit_num = int(float(row['ครั้งที่'])) + 1 if pd.notna(row['ครั้งที่']) and str(row['ครั้งที่']).strip() != "" else 2
                    create_new_visit = st.checkbox(f"✅ บันทึกเป็นประวัติการเข้ารับบริการครั้งใหม่ (ปรับเป็นครั้งที่ {next_visit_num})", value=True, key=f"new_visit_{idx}")
                    
                    if create_new_visit:
                        record_date_update = st.date_input("📅 วันที่รับบริการ (สำหรับการบันทึกประวัติครั้งใหม่)", value=datetime.now(), key=f"new_date_{idx}")
                    
                    if st.button("💾 บันทึกอัปเดต", key=f"save_note_{idx}"):
                        
                        if "พบความผิดปกติ" in new_result_list:
                            if "ผู้ที่พบปัญหาสุขภาพจิตและได้รับการดูแลรักษา" not in new_result_list:
                                new_result_list.append("ผู้ที่พบปัญหาสุขภาพจิตและได้รับการดูแลรักษา")
                            if "รับการประเมินเพื่อวินิจฉัยโรคทางจิตเวช" not in new_result_list:
                                new_result_list.append("รับการประเมินเพื่อวินิจฉัยโรคทางจิตเวช")
                            
                            if "ตรวจผ่านระบบ Telepsychiatry" not in new_service_list:
                                new_service_list.append("ตรวจผ่านระบบ Telepsychiatry")
                                
                        if "จิตบำบัด" in new_service_list:
                            if "การบริการคลินิกคลายเครียดให้การปรึกษา" not in new_service_list:
                                new_service_list.append("การบริการคลินิกคลายเครียดให้การปรึกษา")

                        final_note_parts = [s for s in selected_symptoms if s != "อื่นๆ (ให้ระบุ)"]
                        if other_note.strip():
                            final_note_parts.append(other_note.strip())
                        new_note = ", ".join(final_note_parts)

                        if create_new_visit:
                            st.session_state['patient_data'].at[idx, 'สถานะติดตาม'] = "บันทึกครั้งใหม่แล้ว"
                            new_row = row.to_dict()
                            
                            update_formatted_date = record_date_update.strftime("%d/%m/%Y")
                            new_row.update({
                                "วันที่": update_formatted_date, 
                                "ครั้งที่": next_visit_num, 
                                "ประเภทบริการ": ", ".join(new_service_list), 
                                "ผลประเมิน": ", ".join(new_result_list), 
                                "บันทึกอาการ": new_note, 
                                "ปัญหาและอุปสรรค": new_prob.strip(),
                                "สถานะติดตาม": new_status, 
                                "วันที่นัดติดตาม": new_date_str
                            })
                            st.session_state['patient_data'] = pd.concat([st.session_state['patient_data'], pd.DataFrame([new_row])], ignore_index=True)
                        else:
                            st.session_state['patient_data'].loc[idx, ['ประเภทบริการ', 'ผลประเมิน', 'บันทึกอาการ', 'ปัญหาและอุปสรรค', 'สถานะติดตาม', 'วันที่นัดติดตาม']] = [", ".join(new_service_list), ", ".join(new_result_list), new_note, new_prob.strip(), new_status, new_date_str]
                        
                        conn.update(spreadsheet=SPREADSHEET_URL, worksheet=0, data=st.session_state['patient_data'])
                        st.success("บันทึกการติดตามเรียบร้อยแล้ว!")
                        st.rerun() 
        else:
            st.success("ไม่มีผู้ป่วยที่พบความผิดปกติที่ต้องติดตาม")

# ================= TAB 2: สรุปรายงาน สจ.21 =================
with tab2:
    st.subheader("⚙️ ตั้งค่ารายงานประจำเดือน (ดึงข้อมูลเฉพาะเดือนที่เลือก)")
    
    col_m, col_y, col_d = st.columns(3)
    with col_m:
        months_th = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        report_month = st.selectbox("ประจำเดือน", months_th, index=datetime.now().month-1)
    with col_y:
        current_year_th = datetime.now().year + 543
        report_year = st.number_input("ปี (พ.ศ.)", value=current_year_th, step=1)
    with col_d:
        report_date = st.date_input("ข้อมูล ณ วันที่พิมพ์รายงาน", datetime.now())
        
    col_act1, col_act2 = st.columns([3, 1])
    with col_act1:
        other_act_name = st.text_input("กิจกรรมส่งเสริมสุขภาพจิตอื่นๆ (ระบุชื่อกิจกรรม)", placeholder="เช่น จัดบอร์ดความรู้, เสียงตามสาย...")
    with col_act2:
        other_act_count = st.number_input("จำนวน (ครั้ง)", min_value=0, step=1)
        
    problems = st.text_area("4. ปัญหาและอุปสรรคที่พบ (ถ้ามี)", placeholder="พิมพ์ปัญหาหรืออุปสรรคที่นี่...")

    full_df = st.session_state['patient_data'].copy()
    
    target_month_num = months_th.index(report_month) + 1
    target_year_gregorian = report_year - 543
    
    temp_date = pd.to_datetime(full_df['วันที่'], dayfirst=True, errors='coerce')
    mask = (temp_date.dt.month == target_month_num) & (temp_date.dt.year == target_year_gregorian)
    
    df_report = full_df[mask]

    def count_data(service_kw="", result_kw="", gender=""):
        mask = pd.Series(True, index=df_report.index)
        if gender: mask = mask & (df_report.get('เพศ', '') == gender)
        if service_kw: mask = mask & df_report.get('ประเภทบริการ', pd.Series(dtype=str)).astype(str).str.contains(service_kw, na=False)
        if result_kw: mask = mask & df_report.get('ผลประเมิน', pd.Series(dtype=str)).astype(str).str.contains(result_kw, na=False)
        return mask.sum()

    def count_unique_person(service_kw, gender):
        if 'เพศ' not in df_report.columns or 'ประเภทบริการ' not in df_report.columns or 'ชื่อ-สกุล' not in df_report.columns: return 0
        filtered_df = df_report[(df_report['เพศ'] == gender) & (df_report['ประเภทบริการ'].astype(str).str.contains(service_kw, na=False))]
        return filtered_df['ชื่อ-สกุล'].nunique()

    def count_status(status_val, gender=""):
        if 'สถานะติดตาม' not in df_report.columns or 'เพศ' not in df_report.columns: return 0
        mask = df_report['สถานะติดตาม'].astype(str) == status_val
        if gender:
            mask = mask & (df_report['เพศ'] == gender)
        return mask.sum()

    st.markdown("---")
    st.subheader(f"📊 สรุปตัวเลขรายงาน สจ.21 ประจำเดือน {report_month} {report_year}")
    
    if not df_report.empty:
        summary_data = {
            "รายการ (ตาม สจ.21)": [
                "1. คัดกรองผู้ต้องขังเข้าใหม่", " - พบความผิดปกติ (เข้าใหม่)", " - ได้รับการดูแลรักษา (เข้าใหม่)",
                "ผู้ต้องขังรายเก่าที่ได้รับการคัดกรองซ้ำ", " - พบความผิดปกติ (รายเก่า)", " - ได้รับการดูแลรักษา (รายเก่า)",
                "2. รับการประเมินเพื่อวินิจฉัยโรคทางจิตเวช", "ให้บริการตรวจรักษาในเรือนจำ", "ตรวจผ่านระบบ Telepsychiatry",
                "ส่งต่อไปรับการรักษานอกเรือนจำ", "โรงพยาบาลรับเป็นผู้ป่วยใน (admit)",
                "3. การบริการคลินิกคลายเครียดให้การปรึกษา (จำนวนราย)", "การบริการคลินิกคลายเครียดให้การปรึกษา (จำนวนครั้ง)",
                "เฝ้าระวังผู้ต้องขังมีพฤติกรรมเสี่ยงฆ่าตัวตาย", "ฆ่าตัวตายไม่สำเร็จ", "ฆ่าตัวตายสำเร็จ",
                "อบรมผู้ต้องขังช่วยเหลืองานด้านสุขภาพจิต",
                "4. ผู้ต้องขังใกล้พ้นโทษที่ต้องดูแลต่อเนื่อง"
            ],
            "ชาย": [
                count_data("เข้าใหม่", "", "ชาย"), count_data("เข้าใหม่", "พบความผิดปกติ", "ชาย"), count_data("เข้าใหม่", "ได้รับการดูแลรักษา", "ชาย"),
                count_data("รายเก่า", "", "ชาย"), count_data("รายเก่า", "พบความผิดปกติ", "ชาย"), count_data("รายเก่า", "ได้รับการดูแลรักษา", "ชาย"),
                count_data("", "วินิจฉัย", "ชาย"), count_data("รักษาในเรือนจำ", "", "ชาย"), count_data("Telepsychiatry", "", "ชาย"),
                count_data("", "ส่งต่อไปรับ", "ชาย"), count_data("", "admit", "ชาย"),
                count_unique_person("ให้การปรึกษา", "ชาย"), count_data("ให้การปรึกษา", "", "ชาย"),
                count_data("เฝ้าระวัง", "", "ชาย"), count_data("ฆ่าตัวตายไม่สำเร็จ", "", "ชาย"), count_data("ฆ่าตัวตายสำเร็จ", "", "ชาย"), count_data("อบรม", "", "ชาย"),
                count_status("ใกล้พ้นโทษ", "ชาย")
            ],
            "หญิง": [
                count_data("เข้าใหม่", "", "หญิง"), count_data("เข้าใหม่", "พบความผิดปกติ", "หญิง"), count_data("เข้าใหม่", "ได้รับการดูแลรักษา", "หญิง"),
                count_data("รายเก่า", "", "หญิง"), count_data("รายเก่า", "พบความผิดปกติ", "หญิง"), count_data("รายเก่า", "ได้รับการดูแลรักษา", "หญิง"),
                count_data("", "วินิจฉัย", "หญิง"), count_data("รักษาในเรือนจำ", "", "หญิง"), count_data("Telepsychiatry", "", "หญิง"),
                count_data("", "ส่งต่อไปรับ", "หญิง"), count_data("", "admit", "หญิง"),
                count_unique_person("ให้การปรึกษา", "หญิง"), count_data("ให้การปรึกษา", "", "หญิง"),
                count_data("เฝ้าระวัง", "", "หญิง"), count_data("ฆ่าตัวตายไม่สำเร็จ", "", "หญิง"), count_data("ฆ่าตัวตายสำเร็จ", "", "หญิง"), count_data("อบรม", "", "หญิง"),
                count_status("ใกล้พ้นโทษ", "หญิง")
            ]
        }
        
        summary_df = pd.DataFrame(summary_data)
        st.table(summary_df)
        
        # 1. ฟังก์ชันสร้างรายงานสรุป สจ.21
        def generate_pdf():
            pdf = FPDF()
            pdf.add_page()
            
            if os.path.exists(FONT_PATH):
                pdf.add_font("Sarabun", style="", fname=FONT_PATH)
                pdf.set_font("Sarabun", size=18)
            else:
                pdf.set_font("Arial", size=16)

            date_str = report_date.strftime("%d/%m/%Y")
            
            pdf.cell(0, 10, f"รายงานคลินิกคลายเครียด ประจำเดือน{report_month} พ.ศ. {report_year}", ln=True, align="C")
            pdf.set_font("Sarabun", size=16)
            pdf.cell(0, 10, f"เรือนจำจังหวัดบุรีรัมย์ (ข้อมูล ณ วันที่ {date_str})", ln=True, align="C")
                
            pdf.ln(5)

            pdf.set_font("Sarabun", size=13)
            for idx, row in summary_df.iterrows():
                item = row["รายการ (ตาม สจ.21)"]
                m = row["ชาย"]
                f = row["หญิง"]
                pdf.cell(130, 7, txt=item, border=0)
                pdf.cell(30, 7, txt=f"ชาย: {m} ราย", border=0)
                pdf.cell(30, 7, txt=f"หญิง: {f} ราย", border=0, ln=True)

            pdf.ln(5)
            pdf.set_font("Sarabun", size=14)
            act_text = other_act_name if other_act_name.strip() else "- ไม่ได้ระบุ -"
            pdf.cell(0, 8, txt=f"กิจกรรมส่งเสริมสุขภาพจิตอื่นๆ (ระบุ): {act_text}    จำนวน {other_act_count} ครั้ง", ln=True)
            pdf.cell(0, 8, txt="ปัญหาและอุปสรรค:", ln=True)
            
            pdf.set_font("Sarabun", size=12)
            pdf.multi_cell(0, 6, txt=problems if problems.strip() else "- ไม่มี -")
            
            pdf.ln(10)
            pdf.set_font("Sarabun", size=13)
            pdf.cell(90) 
            pdf.cell(100, 7, txt="เรียน ผู้บังคับบัญชาการเรือนจำจังหวัดบุรีรัมย์", ln=True, align="C")
            pdf.cell(90) 
            pdf.cell(100, 7, txt="      เพื่อโปรดทราบ", ln=True, align="L")
            pdf.ln(21) 
            pdf.cell(90)
            pdf.cell(100, 7, txt="นางสาวเดือนนภา เบี้ยชาติไทย", ln=True, align="C")
            pdf.cell(90)
            pdf.cell(100, 7, txt="(นักจิตวิทยาปฏิบัติการ)", ln=True, align="C")
            
            return bytes(pdf.output())

        # 2. ฟังก์ชันสร้างรายงานตารางข้อมูลดิบ
        def generate_raw_data_pdf():
            pdf = FPDF(orientation="L", unit="mm", format="A4") 
            pdf.add_page()
            
            if os.path.exists(FONT_PATH):
                pdf.add_font("Sarabun", style="", fname=FONT_PATH)
                pdf.set_font("Sarabun", size=18)
            else:
                pdf.set_font("Arial", size=16)

            pdf.cell(0, 10, f"รายงานผู้รับบริการคลินิกคลายเครียด ประจำเดือน{report_month} พ.ศ. {report_year}", ln=True, align="C")
            pdf.ln(2)

            df_clinic = df_report[df_report['ประเภทบริการ'].astype(str).str.contains('คลินิกคลายเครียด', na=False)]

            if not df_clinic.empty:
                try:
                    plt.figure(figsize=(15, 4))
                    thai_font = fm.FontProperties(fname=FONT_PATH)
                    
                    # 1) กราฟวงกลม สัดส่วนเพศ
                    plt.subplot(1, 3, 1)
                    gender_counts = df_clinic['เพศ'].value_counts()
                    if not gender_counts.empty:
                        plt.pie(gender_counts, labels=gender_counts.index, autopct='%1.1f%%', 
                                textprops={'fontproperties': thai_font, 'fontsize': 12}, startangle=90, colors=['#4c72b0', '#dd8452'])
                        plt.title('สัดส่วนเพศผู้รับบริการ', fontproperties=thai_font, fontsize=14)
                    else:
                        plt.text(0.5, 0.5, 'ไม่มีข้อมูลเพศ', ha='center', va='center', fontproperties=thai_font)
                        plt.axis('off')
                        
                    # 2) กราฟแท่ง ช่วงอายุ
                    plt.subplot(1, 3, 2)
                    ages = pd.to_numeric(df_clinic['อายุ'], errors='coerce').dropna()
                    bins = [0, 19, 29, 39, 49, 59, 150]
                    labels = ['<20', '20-29', '30-39', '40-49', '50-59', '60+']
                    if not ages.empty:
                        age_groups = pd.cut(ages, bins=bins, labels=labels).value_counts().reindex(labels, fill_value=0)
                        bars = plt.bar(age_groups.index.astype(str), age_groups.values, color='#55a868')
                        plt.title('ช่วงอายุผู้รับบริการ', fontproperties=thai_font, fontsize=14)
                        plt.xticks(fontproperties=thai_font, fontsize=10)
                        plt.yticks(fontproperties=thai_font, fontsize=10)
                        
                        for bar in bars:
                            yval = bar.get_height()
                            if yval > 0:
                                plt.text(bar.get_x() + bar.get_width()/2, yval + 0.1, int(yval), ha='center', va='bottom', fontproperties=thai_font)
                    else:
                        plt.text(0.5, 0.5, 'ไม่มีข้อมูลอายุ', ha='center', va='center', fontproperties=thai_font)
                        plt.axis('off')
                        
                    # 3) กราฟวงกลม สัดส่วนอาการ
                    plt.subplot(1, 3, 3)
                    all_symptoms = []
                    for s_str in df_clinic['บันทึกอาการ'].dropna():
                        if str(s_str).strip():
                            all_symptoms.extend([s.strip() for s in str(s_str).split(',') if s.strip()])
                    
                    symptom_series = pd.Series(all_symptoms)
                    symptom_counts = symptom_series.value_counts()
                    
                    if not symptom_counts.empty:
                        plt.pie(symptom_counts, labels=symptom_counts.index, autopct='%1.1f%%', 
                                textprops={'fontproperties': thai_font, 'fontsize': 10}, startangle=90)
                        plt.title('สัดส่วนอาการ', fontproperties=thai_font, fontsize=14)
                    else:
                        plt.text(0.5, 0.5, 'ไม่มีข้อมูลอาการ', ha='center', va='center', fontproperties=thai_font)
                        plt.axis('off')

                    plt.tight_layout()
                    
                    chart_path = "temp_chart.png"
                    plt.savefig(chart_path, format='png', bbox_inches='tight')
                    plt.close()
                    
                    y_pos = pdf.get_y()
                    pdf.image(chart_path, x=(297-210)/2, y=y_pos, w=210)
                    pdf.set_y(y_pos + 65)  
                    
                    if os.path.exists(chart_path):
                        os.remove(chart_path)
                except Exception as e:
                    pdf.set_font("Sarabun", size=12)
                    pdf.cell(0, 10, f"(เกิดข้อผิดพลาดในการวาดกราฟ: {e})", ln=True, align="C")
            else:
                pdf.ln(5)

            pdf.set_font("Sarabun", size=11)
            
            w_no, w_date, w_name, w_gender, w_age = 10, 22, 45, 12, 12
            w_room, w_visit, w_note, w_next = 18, 15, 115, 30
            
            pdf.cell(w_no, 10, "ลำดับ", border=1, align="C")
            pdf.cell(w_date, 10, "วันที่", border=1, align="C")
            pdf.cell(w_name, 10, "ชื่อ-สกุล", border=1, align="C")
            pdf.cell(w_gender, 10, "เพศ", border=1, align="C")
            pdf.cell(w_age, 10, "อายุ", border=1, align="C")
            pdf.cell(w_room, 10, "ห้อง", border=1, align="C")
            pdf.cell(w_visit, 10, "ครั้งที่", border=1, align="C")
            pdf.cell(w_note, 10, "บันทึกอาการ", border=1, align="C")
            pdf.cell(w_next, 10, "วันที่นัด", border=1, align="C")
            pdf.ln()

            pdf.set_font("Sarabun", size=10)
            
            if df_clinic.empty:
                total_width = w_no + w_date + w_name + w_gender + w_age + w_room + w_visit + w_note + w_next
                pdf.cell(total_width, 10, "ไม่มีข้อมูลผู้รับบริการคลินิกคลายเครียดในเดือนนี้", border=1, align="C")
                pdf.ln()
            else:
                for i, (index, row) in enumerate(df_clinic.iterrows(), start=1):
                    def trunc(t, l):
                        s = str(t).replace('\n', ' ').strip()
                        return s[:l] + '..' if len(s) > l else s
                    
                    date_only = str(row.get("วันที่", "")).split(" ")[0]
                    next_date_only = str(row.get("วันที่นัดติดตาม", "")).split(" ")[0]
                    age_str = str(row.get("อายุ", "")).replace(".0", "")
                    visit_str = str(row.get("ครั้งที่", "")).replace(".0", "")
                    
                    pdf.cell(w_no, 8, str(i), border=1, align="C")
                    pdf.cell(w_date, 8, trunc(date_only, 15), border=1, align="C")
                    pdf.cell(w_name, 8, trunc(row.get("ชื่อ-สกุล", ""), 25), border=1)
                    pdf.cell(w_gender, 8, trunc(row.get("เพศ", ""), 10), border=1, align="C")
                    pdf.cell(w_age, 8, trunc(age_str, 10), border=1, align="C")
                    pdf.cell(w_room, 8, trunc(row.get("ห้อง", ""), 12), border=1, align="C")
                    pdf.cell(w_visit, 8, trunc(visit_str, 10), border=1, align="C")
                    pdf.cell(w_note, 8, trunc(row.get("บันทึกอาการ", ""), 65), border=1)
                    pdf.cell(w_next, 8, trunc(next_date_only, 15), border=1, align="C")
                    pdf.ln()
                    
            pdf.ln(10)
            pdf.set_font("Sarabun", size=14)
            pdf.cell(0, 8, "ปัญหาและอุปสรรค:", ln=True)
            pdf.set_font("Sarabun", size=12)
            pdf.multi_cell(0, 6, txt=problems if problems.strip() else "- ไม่มี -")
            
            # --- รูปแบบลายเซ็น เสนอลงนาม (A4 แนวนอน) ---
            pdf.ln(15)
            pdf.set_font("Sarabun", size=14)
            pdf.cell(177) 
            pdf.cell(100, 8, txt="เรียน ผู้บังคับบัญชาการเรือนจำจังหวัดบุรีรัมย์", ln=True, align="C")
            pdf.cell(177) 
            pdf.cell(100, 8, txt="      เพื่อโปรดทราบ", ln=True, align="L")
            pdf.ln(24) 
            pdf.cell(177)
            pdf.cell(100, 8, txt="นางสาวเดือนนภา เบี้ยชาติไทย", ln=True, align="C")
            pdf.cell(177)
            pdf.cell(100, 8, txt="(นักจิตวิทยาปฏิบัติการ)", ln=True, align="C")
                
            return bytes(pdf.output())

        st.markdown("### 📥 ดาวน์โหลดรายงาน")
        col_btn1, col_btn2, col_btn3 = st.columns(3)
        
        with col_btn1:
            output_excel = io.BytesIO()
            with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                summary_df.to_excel(writer, index=False, sheet_name='สรุปรายงาน_สจ21')
                df_report.to_excel(writer, index=False, sheet_name='ข้อมูลดิบ')
            st.download_button("📥 โหลดไฟล์ Excel", data=output_excel.getvalue(), file_name=f"Report_Sj21_{report_month}.xlsx")
            
        with col_btn2:
            try:
                pdf_bytes = generate_pdf()
                st.download_button("📄 โหลด PDF (สรุป สจ.21)", data=pdf_bytes, file_name=f"Report_Clinic_{report_month}_{report_year}.pdf", mime="application/pdf")
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดในการสร้าง PDF สรุป: {e}")
                
        with col_btn3:
            try:
                pdf_raw_bytes = generate_raw_data_pdf()
                st.download_button("📄 โหลด PDF (รายงานคลินิกคลายเครียด)", data=pdf_raw_bytes, file_name=f"Report_RawData_Clinic_{report_month}_{report_year}.pdf", mime="application/pdf")
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดในการสร้าง PDF ข้อมูลดิบ: {e}")
                
    else:
        st.info(f"ไม่มีข้อมูลการรับบริการในเดือน **{report_month} {report_year}** ครับ")
