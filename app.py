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
                ws.cell(row=row_idx, column=col_idx).value = ""
                gid = extract_gdrive_id(face_url)
                if gid:
                    try:
                        img_url = f"{PROTO}drive.google.com/uc?id={gid}"
                        headers = {"User-Agent": "Mozilla/5.0"}
                        req = urllib.request.Request(img_url, headers=headers)
                        with urllib.request.urlopen(req) as resp:
                            img = PILImage.open(io.BytesIO(resp.read()))
                            img.thumbnail((80, 80))
                            img_bytes = io.BytesIO()
                            img.save(img_bytes, format="PNG")
                            img_bytes.seek(0)
                            xl_img = openpyxl.drawing.image.Image(img_bytes)
                            xl_img.anchor = f"{col_letter}{row_idx}"
                            ws.add_image(xl_img)
                    except Exception:
                        ws.cell(row=row_idx, column=col_idx).value = "โหลดรูปไม่ได้"
    return buffer.getvalue()


css_lines = [
    "<style>",
    ".stApp { background-color: #F8FAFC; }",
    "h1 { color: #0284c7; font-weight: 800 !important; font-family: 'Sarabun', sans-serif; }",
    "h2, h3 { color: #1e293b; font-family: 'Sarabun', sans-serif; font-weight: 700; }",
    "[data-testid='stMetric'] { background-color: #ffffff; border-radius: 16px; padding: 20px; border: 1px solid #e2e8f0; border-left: 6px solid #0ea5e9; }",
    ".stButton>button { background-color: #0284c7; color: white; border-radius: 10px; border: none; padding: 12px 24px; font-weight: 600; }",
    "</style>",
]
st.markdown("".join(css_lines), unsafe_allow_html=True)

col_title, col_btn = st.columns([3, 1])
with col_title:
    st.title("🌿 ระบบ Report ข้อมูลจิตเวช")
    st.caption("ระบบจัดการข้อมูลและส่งออกรายงานอัตโนมัติ (PDF / Excel)")
with col_btn:
    st.write("")
    st.link_button(
        "📝 เปิดแก้ไขข้อมูลบน Google Sheet",
        EDIT_URL,
        type="secondary",
        use_container_width=True,
    )


@st.cache_data(ttl=60, show_spinner="⏳ กำลังซิงค์ข้อมูลจาก Google Sheet...")
def load_data(url):
    try:
        raw_df = pd.read_csv(url, on_bad_lines="skip")
        raw_df.columns = raw_df.columns.astype(str).str.strip().str.replace("\u200b", "")
        if "ชื่อ-สกุล" not in raw_df.columns and "สถานะ" not in raw_df.columns:
            for i in range(min(10, len(raw_df))):
                row_vals = [str(v).strip() for v in raw_df.iloc[i].values]
                if "ชื่อ-สกุล" in row_vals or "สถานะ" in row_vals:
                    raw_df.columns = row_vals
                    raw_df = raw_df.iloc[i + 1 :].reset_index(drop=True)
                    break
        return raw_df
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการดึงข้อมูล: {e}")
        return None


df = load_data(SHEET_URL)

if df is not None:
    df.columns = df.columns.astype(str).str.strip().str.replace("\u200b", "")

    if "นัดแพทย์" in df.columns and "นัด" not in df.columns:
        df.rename(columns={"นัดแพทย์": "นัด"}, inplace=True)

    req_cols = ["ชื่อ-สกุล", "เพศ", "สถานะ", "Dx", "อาการปัจจุบัน", "หน้า", "ห้อง", "แพทย์", "นัด"]
    for req_col in req_cols:
        if req_col not in df.columns:
            df[req_col] = ""

    if "คอลัมน์ 1" in df.columns:
        target_col = "คอลัมน์ 1"
    elif "วันที่" in df.columns:
        target_col = "วันที่"
    else:
        target_col = df.columns[0]

    df[target_col] = df[target_col].astype(str).str.strip()

    dates_list = [d for d in df[target_col].dropna().unique() if str(d).strip() != "" and str(d).lower() != "nan"]
    docs_list = [str(d).strip() for d in df["แพทย์"].dropna().unique() if str(d).strip() != "" and str(d).lower() != "nan"]
    status_list = [str(d).strip() for d in df["สถานะ"].dropna().unique() if str(d).strip() != "" and str(d).lower() != "nan"]

    unique_dates = ["ทั้งหมด"] + sorted(list(set(dates_list)))
    unique_doctors = ["ทั้งหมด"] + sorted(list(set(docs_list)))
    unique_statuses = ["ทั้งหมด"] + sorted(list(set(status_list)))
    color_options = ["ทั้งหมด", "🔴 แดง", "🟠 ส้ม", "🟡 เหลือง", "🟢 เขียว", "⚪ ไม่มีข้อมูล"]

    f_col1, f_col2, f_col3, f_col4 = st.columns(4)
    with f_col1:
        selected_dates = st.multiselect("📅 เลือกวันที่:", options=unique_dates, default=["ทั้งหมด"])
    with f_col2:
        selected_doctors = st.multiselect("🩺 เลือกแพทย์:", options=unique_doctors, default=["ทั้งหมด"])
    with f_col3:
        selected_statuses = st.multiselect("📌 เลือกสถานะผู้ป่วย:", options=unique_statuses, default=["ทั้งหมด"])
    with f_col4:
        selected_colors = st.multiselect("🎨 เลือกระดับสี:", options=color_options, default=["ทั้งหมด"])

    has_filters = selected_dates and selected_doctors and selected_statuses and selected_colors
    if not has_filters:
        st.warning("⚠️ กรุณาเลือกตัวกรองอย่างน้อย 1 รายการในแต่ละช่อง")
    else:
        filtered_df = df.copy()
        if "ทั้งหมด" not in selected_dates:
            filtered_df = filtered_df[filtered_df[target_col].isin(selected_dates)]
        if "ทั้งหมด" not in selected_doctors:
            filtered_df = filtered_df[filtered_df["แพทย์"].astype(str).str.strip().isin(selected_doctors)]
        if "ทั้งหมด" not in selected_statuses:
            filtered_df = filtered_df[filtered_df["สถานะ"].astype(str).str.strip().isin(selected_statuses)]

        if "ทั้งหมด" not in selected_colors:
            f_col = filtered_df["หน้า"].fillna("").astype(str)
            mask = pd.Series(False, index=filtered_df.index)
            if "🔴 แดง" in selected_colors:
                mask |= f_col.str.contains(ID_RED)
            if "🟠 ส้ม" in selected_colors:
                mask |= f_col.str.contains(ID_ORG)
            if "🟡 เหลือง" in selected_colors:
                mask |= f_col.str.contains(ID_YEL)
            if "🟢 เขียว" in selected_colors:
                mask |= f_col.str.contains(ID_GRN)
            if "⚪ ไม่มีข้อมูล" in selected_colors:
                any_c = (f_col.str.contains(ID_RED) | f_col.str.contains(ID_ORG) | f_col.str.contains(ID_YEL) | f_col.str.contains(ID_GRN))
                mask |= ~any_c
            filtered_df = filtered_df[mask]

        if len(filtered_df) > 0:
            filtered_df["นัด"] = filtered_df["นัด"].fillna("")
            filtered_df = filtered_df.sort_values(by="นัด").reset_index(drop=True)
            st.markdown("---")

            total_p = len(filtered_df)
            new_p = len(filtered_df[filtered_df["สถานะ"].astype(str).str.strip() == "รายใหม่"])
            red_p = sum(filtered_df["หน้า"].fillna("").astype(str).str.contains(ID_RED, na=False))

            filtered_df["เพศ"] = filtered_df["เพศ"].astype(str).str.strip()
            male_c = len(filtered_df[filtered_df["เพศ"] == "ชาย"])
            female_c = len(filtered_df[filtered_df["เพศ"] == "หญิง"])

            new_pct = (new_p / total_p * 100) if total_p > 0 else 0
            red_pct = (red_p / total_p * 100) if total_p > 0 else 0

            d_col1, d_col2, d_col3, d_col4 = st.columns([1, 1.2, 1.2, 1.5])
            d_col1.metric("👥 ผู้ป่วยทั้งหมด", f"{total_p} ราย")
            d_col2.metric("🆕 ผู้ป่วยรายใหม่", f"{new_p} ราย ({new_pct:.1f}%)")
            d_col3.metric("🚨 เคสเฝ้าระวัง (แดง)", f"{red_p} ราย ({red_pct:.1f}%)")

            with d_col4:
                st.caption("🚻 สัดส่วนเพศ (Gender)")
                if male_c + female_c > 0:
                    fig_d, ax_d = plt.subplots(figsize=(2.5, 1.5))
                    fig_d.patch.set_alpha(0.0)
                    sizes, labels, colors = [], [], []
                    if male_c > 0:
                        sizes.append(male_c)
                        labels.append("Male")
                        colors.append("#38bdf8")
                    if female_c > 0:
                        sizes.append(female_c)
                        labels.append("Female")
                        colors.append("#f472b6")
                    
                    tot_g = sum(sizes)
                    
                    wedges, texts, autotexts = ax_d.pie(
                        sizes, 
                        labels=labels, 
                        autopct=lambda p: f"{int(round(p * tot_g / 100))} ({p:.1f}%)", 
                        colors=colors, 
                        startangle=90, 
                        textprops={"fontsize": 8, "fontweight": "bold"}
                    )
                    
                    for t in texts:
                        t.set_color("#475569")
                    for at in autotexts:
                        at.set_color("black")
                    ax_d.axis("equal")
                    st.pyplot(fig_d)
                    plt.close(fig_d)
                else:
                    st.info("ไม่มีข้อมูลเพศ")

            st.markdown("---")
            p_col1, p_col2 = st.columns(2)
            with p_col1:
                show_table = st.toggle("👁️ แสดงตารางข้อมูลผู้ป่วยบนหน้าจอ", value=True, help="ปิดสวิตช์นี้เพื่อซ่อนตารางทั้งหมดขณะนำเสนอในที่ประชุม")
            with p_col2:
                mask_pii = st.toggle("🔒 ปิดบังชื่อและข้อมูลส่วนบุคคล (PDPA Mode)", value=False, help="เปลี่ยนชื่อ-สกุลเป็นรหัสและซ่อนข้อมูลส่วนตัวทั้งบนจอและในไฟล์")

            working_df = apply_pdpa_mask(filtered_df) if mask_pii else filtered_df.copy()

            def get_direct_url_preview(url):
                gid = extract_gdrive_id(url)
                if gid:
                    return f"{PROTO}drive.google.com/thumbnail?id={gid}&sz=w100"
                return ""

            def highlight_and_group(subset_df):
                styles = pd.DataFrame("", index=subset_df.index, columns=subset_df.columns)
                if "นัด" in subset_df.columns:
                    v_dates = [d for d in subset_df["นัด"].unique() if str(d).strip() != "" and str(d).lower() != "nan"]
                    p_colors = ["#E0F2F1", "#FFF9C4", "#F3E5F5", "#E3F2FD", "#FBE9E7"]
                    c_map = {d: p_colors[i % len(p_colors)] for i, d in enumerate(v_dates)}
                    for idx, r in subset_df.iterrows():
                        ad = r.get("นัด", "")
                        if ad in c_map:
                            styles.loc[idx, "นัด"] = f"background-color: {c_map[ad]}; font-weight: 600;"
                if "หน้า" in subset_df.columns:
                    r_mask = subset_df["หน้า"].fillna("").astype(str).str.contains(ID_RED, na=False)
                    for col in styles.columns:
                        styles.loc[r_mask, col] = "background-color: #FCE4EC;"
                return styles

            tab_pdf, tab_excel = st.tabs(["📄 Telepsychiatry Report", "📊 รายงานทั่วไป (Excel)"])

            with tab_pdf:
                problem_text = st.text_area("✍️ บันทึกปัญหา / อุปสรรค (ถ้ามี):", placeholder="พิมพ์ปัญหาหรืออุปสรรคที่พบในวันนี้ที่นี่...")
                suggestion_text = st.text_area("💡 ข้อเสนอแนะ (ถ้ามี):", placeholder="พิมพ์ข้อเสนอแนะเพิ่มเติมที่นี่...")
                preview_cols = ["ชื่อ-สกุล", "สถานะ", "Dx", "อาการปัจจุบัน", "หน้า", "นัด", "ห้อง", "แพทย์"]
                avail_cols = [c for c in preview_cols if c in working_df.columns]
                edited_df = working_df.copy()

                if show_table:
                    st.markdown("**📋 Preview ข้อมูล (ดับเบิลคลิกแก้ไขในตารางได้):**")
                    pdf_display_df = working_df[avail_cols].copy().fillna("")
                    if "หน้า" in pdf_display_df.columns:
                        pdf_display_df["หน้า"] = pdf_display_df["หน้า"].apply(get_direct_url_preview)
                    edited_pdf_display = st.data_editor(
                        pdf_display_df.style.apply(highlight_and_group, axis=None),
                        use_container_width=True, hide_index=True, column_order=avail_cols, key="pdf_editor_table",
                        column_config={
                            "หน้า": st.column_config.ImageColumn("ระดับ (รูปภาพ)"),
                            "นัด": st.column_config.TextColumn("นัดครั้งถัดไป"),
                            "ห้อง": st.column_config.TextColumn("ห้อง"),
                        },
                    )
                    for col in avail_cols:
                        if col in edited_pdf_display.columns:
                            edited_df[col] = edited_pdf_display[col]
                else:
                    st.info("🔒 ซ่อนตารางรายชื่อผู้ป่วยอยู่ (โหมดนำเสนอที่ประชุม) — คุณยังสามารถกดปุ่มสร้างรายงาน PDF ด้านล่างได้ตามปกติ")

                if st.button("🚀 สร้างรายงาน PDF", type="primary", key="btn_pdf"):
                    from weasyprint import HTML
                    
                    with st.status("⏳ กำลังสร้างรายงาน PDF...", expanded=True) as status:
                        st.markdown("1️⃣ **กำลังจัดรูปแบบตาราง...**")
                        rows_list = []
                        for r_num, (_, row) in enumerate(edited_df.iterrows(), start=1):
                            name = str(row.get("ชื่อ-สกุล", "")).replace("nan", "")
                            st_txt = str(row.get("สถานะ", "")).replace("nan", "")
                            dx = str(row.get("Dx", "")).replace("nan", "")
                            symp = str(row.get("อาการปัจจุบัน", "")).replace("nan", "")
                            room = str(row.get("ห้อง", "")).replace("nan", "")
                            appt = str(row.get("นัด", "")).replace("nan", "")
                            doc = str(row.get("แพทย์", "")).replace("nan", "")
                            gid = extract_gdrive_id(row.get("หน้า", ""))
                            if gid:
                                u = f"{PROTO}drive.google.com/uc?id={gid}"
                                img_t = f"<img src='{u}' style='width:26px;height:26px;object-fit:cover;border-radius:4px;'/>"
                            else:
                                img_t = "-"
                            row_html = "".join([
                                "<tr>", f"<td style='text-align:center;'>{r_num}</td>",
                                f"<td style='font-weight:bold;color:#1e293b;'>{name}</td>",
                                f"<td>{st_txt}</td>", f"<td>{dx}</td>", f"<td>{symp}</td>",
                                f"<td style='text-align:center;'>{img_t}</td>",
                                f"<td style='color:#0369a1;'>{appt}</td>",
                                f"<td style='text-align:center;'>{room}</td>", f"<td>{doc}</td>", "</tr>",
                            ])
                            rows_list.append(row_html)
                        html_rows = "".join(rows_list)

                        st.markdown("2️⃣ **กำลังประมวลผลสถิติและวาดกราฟ...**")
                        st_s = edited_df["สถานะ"].astype(str).str.strip()
                        gd_s = edited_df["เพศ"].astype(str).str.strip()
                        old_c = edited_df[st_s == "รายเก่า"]
                        new_c = edited_df[st_s == "รายใหม่"]
                        nr_c = edited_df[st_s == "ใกล้พ้นโทษ"]
                        old_m = len(old_c[gd_s.loc[old_c.index] == "ชาย"])
                        old_f = len(old_c[gd_s.loc[old_c.index] == "หญิง"])
                        new_m = len(new_c[gd_s.loc[new_c.index] == "ชาย"])
                        new_f = len(new_c[gd_s.loc[new_c.index] == "หญิง"])
                        nr_m = len(nr_c[gd_s.loc[nr_c.index] == "ชาย"])
                        nr_f = len(nr_c[gd_s.loc[nr_c.index] == "หญิง"])
                        
                        dx_counts = edited_df["Dx"].value_counts()
                        valid_dx = {k: v for k, v in dx_counts.items() if str(k).strip() != "" and str(k).lower() != "nan"}
                        if valid_dx:
                            fig1, ax1 = plt.subplots(figsize=(2.8, 2.8))
                            tot_dx = sum(valid_dx.values())
                            labels1 = [f"{k}\n{v} ({v / tot_dx * 100:.1f}%)" for k, v in valid_dx.items()]
                            wedges1, _ = ax1.pie(valid_dx.values(), startangle=90, colors=plt.cm.tab20.colors, radius=0.55)
                            kw = dict(arrowprops=dict(arrowstyle="-", color="#64748b", lw=1.0), zorder=0, va="center")
                            for i, p in enumerate(wedges1):
                                ang = (p.theta2 - p.theta1) / 2.0 + p.theta1
                                safe_ang = ang + 1.0 if abs(ang % 180) < 1 else ang
                                y = np.sin(np.deg2rad(ang))
                                x = np.cos(np.deg2rad(ang))
                                x_sign = -1 if x < 0 else 1
                                ha = "right" if x_sign == -1 else "left"
                                conn = f"angle,angleA=0,angleB={safe_ang}"
                                kw["arrowprops"].update({"connectionstyle": conn})
                                ax1.annotate(
                                    labels1[i], xy=(x * 0.55, y * 0.55), xytext=(0.75 * x_sign, 0.8 * y),
                                    horizontalalignment=ha, fontsize=8, color="#111827", fontfamily=THAI_FONT_NAME, **kw
                                )
                            ax1.axis("equal")
                            buf1 = io.BytesIO()
                            plt.savefig(buf1, format="png", bbox_inches="tight", transparent=True, dpi=120)
                            b64_1 = base64.b64encode(buf1.getvalue()).decode("utf-8")
                            chart_img_tag = f"<img src='data:image/png;base64,{b64_1}' style='width:100%;max-width:180px;display:block;margin:auto;'/>"
                            plt.close(fig1)
                        else:
                            chart_img_tag = "<p style='font-size:9pt;'>ไม่มีข้อมูล Dx</p>"

                        c_map = {"แดง": "#F44336", "ส้ม": "#FF9800", "เหลือง": "#FACC15", "เขียว": "#4CAF50", "เทา": "#9E9E9E"}
                        lvl_cnt = {"แดง": 0, "ส้ม": 0, "เหลือง": 0, "เขียว": 0, "เทา": 0}
                        for f_url in edited_df["หน้า"]:
                            fs = str(f_url)
                            if ID_RED in fs: lvl_cnt["แดง"] += 1
                            elif ID_ORG in fs: lvl_cnt["ส้ม"] += 1
                            elif ID_YEL in fs: lvl_cnt["เหลือง"] += 1
                            elif ID_GRN in fs: lvl_cnt["เขียว"] += 1
                            else: lvl_cnt["เทา"] += 1
                        act_lvl = {k: v for k, v in lvl_cnt.items() if v > 0}
                        if act_lvl:
                            fig2, ax2 = plt.subplots(figsize=(2.2, 2.2))
                            cols2 = [c_map[k] for k in act_lvl.keys()]
                            tot_lvl = sum(act_lvl.values())
                            ax2.pie(
                                act_lvl.values(), autopct=lambda p: f"{int(round(p * tot_lvl / 100))} ({p:.1f}%)",
                                startangle=90, colors=cols2, textprops={"fontsize": 8, "fontfamily": THAI_FONT_NAME}
                            )
                            ax2.axis("equal")
                            buf2 = io.BytesIO()
                            plt.savefig(buf2, format="png", bbox_inches="tight", transparent=True, dpi=120)
                            b64_2 = base64.b64encode(buf2.getvalue()).decode("utf-8")
                            level_chart_img_tag = f"<img src='data:image/png;base64,{b64_2}' style='width:100%;max-width:130px;display:block;margin:auto;'/>"
                            plt.close(fig2)
                        else:
                            level_chart_img_tag = "<p style='font-size:9pt;'>ไม่มีข้อมูล</p>"

                        logo_src = f"{PROTO}drive.google.com/uc?id={ID_LOGO}"
                        if "ทั้งหมด" in selected_dates:
                            title_text = "รายงาน Tele psychiatry"
                        else:
                            d_str = ", ".join(selected_dates)
                            title_text = f"รายงาน Tele psychiatry วันที่ {d_str}"

                        extra_boxes = []
                        if problem_text.strip():
                            extra_boxes.extend([
                                "<div style='background-color:#fef2f2;border-left:5px solid #ef4444;padding:12px;margin-top:15px;'>",
                                "<h3 style='color:#b91c1c;margin-top:0;font-size:10pt;'>ปัญหา / อุปสรรค</h3>",
                                f"<p style='margin:0;font-size:9pt;white-space:pre-line;'>{problem_text}</p></div>",
                            ])
                        if suggestion_text.strip():
                            extra_boxes.extend([
                                "<div style='background-color:#f0fdf4;border-left:5px solid #22c55e;padding:12px;margin-top:15px;'>",
                                "<h3 style='color:#15803d;margin-top:0;font-size:10pt;'>ข้อเสนอแนะ</h3>",
                                f"<p style='margin:0;font-size:9pt;white-space:pre-line;'>{suggestion_text}</p></div>",
                            ])
                        bottom_sections = "".join(extra_boxes)

                        st.markdown("3️⃣ **กำลังสร้างไฟล์ PDF...**")
                        font_css = f"{PROTO}fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap"
                        html_parts = [
                            "<!DOCTYPE html><html lang='th'><head><meta charset='UTF-8'>",
                            f"<link href='{font_css}' rel='stylesheet'><style>",
                            "@page { size: A4 portrait; margin: 10mm 10mm 15mm 10mm; }",
                            "body { font-family: 'Sarabun', sans-serif; font-size: 11pt; color: #334155; line-height: 1.5; }",
                            ".header-table { width: 100%; border-collapse: collapse; margin-bottom: 20px; border-bottom: 2px solid #cbd5e1; }",
                            ".header-table td { border: none; padding-bottom: 10px; vertical-align: bottom; }",
                            ".header-logo { width: 130px; text-align: center; }",
                            ".header-logo img { width: 55px; height: auto; }",
                            ".header-logo p { font-size: 9pt; font-weight: bold; margin-top: 5px; margin-bottom: 0; }",
                            "h1 { text-align: center; margin: 0; font-size: 12pt; }",
                            ".summary-container { width: 100%; border-collapse: separate; border-spacing: 8px 0; margin-bottom: 20px; table-layout: fixed; }",
                            ".summary-box { background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px; vertical-align: top; width: 33.33%; }",
                            ".summary-box h3 { margin-top: 0; color: #0369a1; font-size: 10pt; border-bottom: 1px solid #cbd5e1; padding-bottom: 6px; }",
                            ".data-table { width: 100%; border-collapse: collapse; margin-top: 10px; table-layout: auto; }",
                            ".data-table th, .data-table td { padding: 6px 4px; vertical-align: middle; font-size: 9pt; border-bottom: 1px solid #e2e8f0; }",
                            ".data-table th { background-color: #1e293b; color: #ffffff; text-align: center; font-weight: 600; }",
                            ".data-table tr:nth-child(even) { background-color: #f8fafc; }",
                            ".signature-section { margin-top: 30px; text-align: left; page-break-inside: avoid; }</style></head><body>",
                            f"<table class='header-table'><tr><td class='header-logo'><img src='{logo_src}'><p>สถานพยาบาลเรือนจำ<br>จังหวัดบุรีรัมย์</p></td>",
                            f"<td style='vertical-align:middle;'><h1>{title_text}</h1></td><td style='width:130px;'></td></tr></table>",
                            "<table class='summary-container'><tr>",
                            "<td class='summary-box' style='text-align:left;'><h3>สรุปสถานะผู้ป่วย</h3><p style='margin:0 0 5px 0;font-size:9pt;'>",
                            f"<strong>ผู้ป่วยรายเก่า:</strong> {len(old_c)} ราย (ช {old_m}, ญ {old_f})<br>",
                            f"<strong>ผู้ป่วยรายใหม่:</strong> {len(new_c)} ราย (ช {new_m}, ญ {new_f})<br>",
                            f"<strong>ใกล้พ้นโทษ:</strong> {len(nr_c)} ราย (ช {nr_m}, ญ {nr_f})</p></td>",
                            f"<td class='summary-box' style='text-align:center;'><h3>สรุปการวินิจฉัยโรค</h3>{chart_img_tag}</td>",
                            f"<td class='summary-box' style='text-align:center;'><h3>สรุปเคสตามระดับสี</h3>{level_chart_img_tag}</td></tr></table>",
                            "<table class='data-table'><thead><tr><th style='width:4%;'>ที่</th><th style='width:15%;'>ชื่อ-สกุล</th>",
                            "<th style='width:7%;'>สถานะ</th><th style='width:8%;'>Dx</th><th style='width:27%;'>อาการปัจจุบัน</th>",
                            "<th style='width:6%;'>ระดับ</th><th style='width:11%;'>นัดครั้งถัดไป</th><th style='width:10%;'>ห้อง</th>",
                            f"<th style='width:12%;'>แพทย์</th></tr></thead><tbody>{html_rows}</tbody></table>",
                            bottom_sections,
                            "<div class='signature-section'><p style='margin-bottom:20px;font-size:11pt;'>เรียน ผู้บัญชาการเรือนจำฯ<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- เพื่อโปรดทราบ</p>",
                            "<br><br><div style='display:inline-block;text-align:center;'><p style='margin:0;font-size:9pt;'>(นางสาวเดือนนภา เบี้ยชาติไทย)</p>",
                            "<p style='margin:5px 0 0 0;font-size:9pt;'>นักจิตวิทยาปฏิบัติการ</p></div></div></body></html>",
                        ]
                        pdf_bytes = HTML(string="".join(html_parts)).write_pdf()
                        status.update(label="✅ สร้างไฟล์ PDF เสร็จสมบูรณ์!", state="complete", expanded=False)

                    if "ทั้งหมด" in selected_dates:
                        f_date = "All_Dates"
                    else:
                        f_date = "_".join([d.replace("/", "-") for d in selected_dates])
                    st.success(f"🎉 พร้อมดาวน์โหลดแล้ว! ({len(edited_df)} รายการ)")
                    st.download_button(
                        label="📥 ดาวน์โหลดไฟล์ PDF", data=pdf_bytes, file_name=f"Telepsychiatry_report_{f_date}.pdf", mime="application/pdf"
                    )

            with tab_excel:
                st.markdown("### 📊 ส่งออกข้อมูลรูปแบบตาราง (Excel)")
                all_cols = working_df.columns.tolist()
                pref_cols = ["ชื่อ-สกุล", "เพศ", "สถานะ", "Dx", "อาการปัจจุบัน", "หน้า", "ห้อง", "แพทย์", "นัด"]
                def_cols = [c for c in pref_cols if c in all_cols]
                selected_export_cols = st.multiselect("📌 เลือกคอลัมน์ที่จะส่งออก:", options=all_cols, default=def_cols if def_cols else all_cols)
                
                if selected_export_cols:
                    edited_ex_df = working_df[selected_export_cols].copy().fillna("")
                    if show_table:
                        preview_ex_df = edited_ex_df.copy()
                        if "หน้า" in preview_ex_df.columns:
                            preview_ex_df["หน้า"] = preview_ex_df["หน้า"].apply(get_direct_url_preview)
                            edited_ex_df = st.data_editor(
                                preview_ex_df.style.apply(highlight_and_group, axis=None),
                                use_container_width=True, hide_index=True, key="excel_editor_img",
                                column_config={"หน้า": st.column_config.ImageColumn("ระดับ (รูปภาพ)")},
                            )
                        else:
                            edited_ex_df = st.data_editor(
                                preview_ex_df.style.apply(highlight_and_group, axis=None),
                                use_container_width=True, hide_index=True, key="excel_editor_no_img"
                            )
                    else:
                        st.info("🔒 ซ่อนตารางรายชื่อผู้ป่วยอยู่ (โหมดนำเสนอที่ประชุม) — คุณยังสามารถกดปุ่มสร้างรายงาน Excel ได้ตามปกติ")

                    if st.button("🚀 สร้างรายงาน Excel", type="primary", key="btn_ex"):
                        with st.status("⏳ กำลังสร้างไฟล์ Excel...", expanded=True) as status:
                            ex_final = edited_ex_df.copy()
                            if "ลำดับ" in ex_final.columns:
                                ex_final = ex_final.drop(columns=["ลำดับ"])
                            ex_final.insert(0, "ลำดับ", range(1, len(ex_final) + 1))
                            cols_final = ["ลำดับ"] + [c for c in selected_export_cols if c != "ลำดับ"]
                            excel_bytes = generate_excel_with_images(ex_final, cols_final)
                            status.update(label="✅ เตรียมไฟล์ Excel สำเร็จ!", state="complete", expanded=False)

                        if "ทั้งหมด" in selected_dates:
                            f_date = "All_Dates"
                        else:
                            f_date = "_".join([d.replace("/", "-") for d in selected_dates])
                        xlsx_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        st.download_button(
                            label="📥 ดาวน์โหลดไฟล์ Excel", data=excel_bytes,
                            file_name=f"Psychiatry_report_{f_date}.xlsx", mime=xlsx_mime, type="primary"
                        )
        else:
            st.warning("ไม่พบข้อมูลผู้ป่วยในเงื่อนไขที่เลือก")
