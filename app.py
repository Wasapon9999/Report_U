import streamlit as st
import pandas as pd
import os
import time
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle,
    Image as RLImage, Paragraph, Spacer, PageBreak, KeepTogether
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import Image, ExifTags

# Google Drive API
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from google.oauth2 import service_account

# --- 1. การตั้งค่าหน้าเว็บและ Config ---
st.set_page_config(page_title="USO1-Report Manager", layout="wide")

GOOGLE_DRIVE_FOLDER_ID = '1-4OwgP-ODbelbtwSg5-m-rm4cyOTcW7O'

@st.cache_resource
def get_drive_service():
    try:
        if "gcp_service_account" in st.secrets:
            info = st.secrets["gcp_service_account"]
            creds = service_account.Credentials.from_service_account_info(
                info, scopes=['https://www.googleapis.com/auth/drive']
            )
        elif os.path.exists("service_account.json"):
            creds = service_account.Credentials.from_service_account_file(
                'service_account.json', scopes=['https://www.googleapis.com/auth/drive']
            )
        else: return None
        return build('drive', 'v3', credentials=creds)
    except Exception as e:
        st.error(f"⚠️ Drive Connection Error: {e}"); return None

def init_fonts():
    try:
        pdfmetrics.registerFont(TTFont('THSarabun', 'THSarabunNew.ttf'))
        pdfmetrics.registerFont(TTFont('THSarabun-Bold', 'THSarabunNew Bold.ttf'))
        return 'THSarabun', 'THSarabun-Bold'
    except: return 'Helvetica', 'Helvetica-Bold'

F_REG, F_BOLD = init_fonts()

# --- 2. Google Drive Helpers ---

@st.cache_data(ttl=600, show_spinner=False)
def download_image_direct(file_name):
    service = get_drive_service()
    if not service or not file_name or file_name in ["0", "nan", ""]: return None
    try:
        query = f"'{GOOGLE_DRIVE_FOLDER_ID}' in parents and name = '{file_name}' and trashed = false"
        results = service.files().list(q=query, fields="files(id)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
        items = results.get('files', [])
        if not items: return None
        file_id = items[0]['id']
        request = service.files().get_media(fileId=file_id)
        fh = BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)
        return fh.getvalue()
    except: return None

def upload_and_overwrite(target_filename, content_bytes):
    service = get_drive_service()
    if not service: return
    try:
        query = f"'{GOOGLE_DRIVE_FOLDER_ID}' in parents and name = '{target_filename}' and trashed = false"
        results = service.files().list(q=query, fields="files(id)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
        for f in results.get('files', []):
            try: service.files().delete(fileId=f['id'], supportsAllDrives=True).execute()
            except: pass
        file_metadata = {'name': target_filename, 'parents': [GOOGLE_DRIVE_FOLDER_ID]}
        media = MediaIoBaseUpload(BytesIO(content_bytes), mimetype='image/jpeg', resumable=True)
        service.files().create(body=file_metadata, media_body=media, supportsAllDrives=True).execute()
        download_image_direct.clear(target_filename)
    except Exception as e:
        st.error(f"❌ อัปโหลดล้มเหลว: {str(e)}")

# --- 3. Utility & PDF Generator ---
def apply_exif_orientation(img):
    try:
        exif = img._getexif()
        if exif:
            for tag, value in exif.items():
                if ExifTags.TAGS.get(tag) == 'Orientation':
                    if value == 3: img = img.transpose(Image.ROTATE_180)
                    elif value == 6: img = img.transpose(Image.ROTATE_270)
                    elif value == 8: img = img.transpose(Image.ROTATE_90)
                    break
    except: pass
    return img

def fmt_time(t):
    if not t or pd.isna(t) or str(t).strip() == "": return ""
    t = str(t).strip().replace(".", ":")
    try:
        parts = t.split(":")
        return f"{int(parts[0]):02d}:{int(parts[1]):02d}"
    except: return t

def parse_thai_date_simple(s):
    month_thai_name = {1: "มกราคม", 2: "กุมภาพันธ์", 3: "มีนาคม", 4: "เมษายน", 5: "พฤษภาคม", 6: "มิถุนายน",
                       7: "กรกฎาคม", 8: "สิงหาคม", 9: "กันยายน", 10: "ตุลาคม", 11: "พฤศจิกายน", 12: "ธันวาคม"}
    if not s or pd.isna(s): return pd.NaT, ""
    try:
        parts = str(s).strip().split()
        if len(parts) == 3:
            day, month_name, year = parts
            m_map = {"มกราคม":1,"กุมภาพันธ์":2,"มีนาคม":3,"เมษายน":4,"พฤษภาคม":5,"มิถุนายน":6,"กรกฎาคม":7,"สิงหาคม":8,"กันยายน":9,"ตุลาคม":10,"พฤศจิกายน":11,"ธันวาคม":12}
            m_int = m_map.get(month_name, 1)
            y_int = int(year); y_int = y_int - 543 if y_int > 2500 else y_int
            dt = pd.to_datetime(f"{y_int}-{m_int:02d}-{int(day):02d}")
            return dt, f"{day} {month_thai_name[m_int]} {int(year)}"
    except: pass
    return pd.NaT, str(s)

def generate_pdf_original_style(df, center_name):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=25, bottomMargin=15)
    thai_styles = {
        "Normal": ParagraphStyle("ThaiNormal", fontName=F_REG, fontSize=14, leading=18, alignment=1),
        "Title": ParagraphStyle("ThaiTitle", fontName=F_BOLD, fontSize=18, leading=24, alignment=1),
        "Heading2": ParagraphStyle("ThaiHeading2", fontName=F_BOLD, fontSize=14, leading=20, alignment=1),
        "Signature": ParagraphStyle("ThaiSignature", fontName=F_REG, fontSize=14, leading=18, alignment=1),
        "HeaderStyle": ParagraphStyle("H", fontName=F_BOLD, fontSize=10, leading=11, alignment=1),
        "CellStyle": ParagraphStyle("C", fontName=F_REG, fontSize=10, leading=11, alignment=1),
    }
    story = []
    story.append(Paragraph("รายงานเวลาปฏิบัติงาน USO1-Renew", thai_styles["Title"]))
    story.append(Paragraph(f"ศูนย์ : {center_name}", thai_styles["Title"]))
    dt_first, date_str = parse_thai_date_simple(df.iloc[0]['date'])
    if pd.notna(dt_first): story.append(Paragraph(f"เดือน : {date_str.split(' ', 1)[1]}", thai_styles["Heading2"]))
    valid_names = df["name"].loc[df["name"].str.strip() != ""]
    emp_name = valid_names.iloc[0] if not valid_names.empty else ""
    story.append(Paragraph(f"เจ้าหน้าที่ดูแลประจำศูนย์ : {emp_name}", thai_styles["Heading2"]))
    story.append(Spacer(1, 2))
    table_data = [[Paragraph(h, thai_styles["HeaderStyle"]) for h in ["ลำดับ", "วันที่", "ชื่อ - นามสกุล", "เวลาเข้า", "เวลาออก", "ตำแหน่ง", "หมายเหตุ"]]]
    for i, row in df.iterrows():
        _, d_thai = parse_thai_date_simple(row['date'])
        table_data.append([Paragraph(str(i+1), thai_styles["CellStyle"]), Paragraph(d_thai, thai_styles["CellStyle"]), Paragraph(row['name'], thai_styles["CellStyle"]), Paragraph(fmt_time(row['time_in']), thai_styles["CellStyle"]), Paragraph(fmt_time(row['time_out']), thai_styles["CellStyle"]), Paragraph(row['status'], thai_styles["CellStyle"]), Paragraph("", thai_styles["CellStyle"])])
    tbl = Table(table_data, colWidths=[35, 100, 130, 60, 60, 80, 70], repeatRows=1)
    tbl.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black), ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("ALIGN", (0, 0), (-1, -1), "CENTER")]))
    story.append(tbl)
    story.append(Spacer(1, 30))
    sig_style = thai_styles["Signature"]
    sig_left = [Paragraph("....................................", sig_style), Spacer(1, 6), Paragraph(f"( {emp_name} )", sig_style), Paragraph("ผดล.ประจำศูนย์", sig_style)]
    sig_right = [Paragraph("....................................", sig_style), Spacer(1, 6), Paragraph("( ...................................... )", sig_style), Paragraph("ตำแหน่ง_______________________", sig_style)]
    story.append(KeepTogether(Table([[sig_left, sig_right]], colWidths=[260, 260])))
    for _, r in df.iterrows():
        story.append(PageBreak())
        _, d_thai = parse_thai_date_simple(r['date'])
        story.append(Paragraph(f"วันที่ : {d_thai}", thai_styles["Heading2"]))
        story.append(Spacer(1, 12))
        story.append(Paragraph(f"ชื่อ : <b>{r['name']}</b> &nbsp; ตำแหน่ง : <b>{r['status']}</b>", thai_styles["Normal"]))
        for label, col_img, col_time in [("เข้า (เช้า)", "img_in1", "time_in"), ("ออก (เย็น)", "img_out1", "time_out")]:
            img_bytes = download_image_direct(r[col_img])
            if img_bytes:
                try:
                    with Image.open(BytesIO(img_bytes)) as PIL_img:
                        PIL_img = apply_exif_orientation(PIL_img)
                        temp_io = BytesIO()
                        PIL_img.convert('RGB').save(temp_io, format="JPEG", quality=85)
                        temp_io.seek(0)
                        im = RLImage(temp_io)
                        im._restrictSize(310, 260)
                        img_tbl = Table([[im]], colWidths=[450])
                        img_tbl.setStyle(TableStyle([('ALIGN', (0, 0), (-1, -1), 'CENTER')]))
                        story.append(img_tbl)
                except: pass
            story.append(Paragraph(f"เวลา{label} : <b>{fmt_time(r[col_time])}</b>", thai_styles["Normal"]))
            story.append(Spacer(1, 18))
    doc.build(story)
    return buffer.getvalue()

# --- 4. Main UI ---

if 'main_df' not in st.session_state:
    try:
        st.session_state.main_df = pd.read_csv("03-2026.csv").fillna("")
    except:
        st.error("❌ ไม่พบไฟล์ CSV"); st.stop()

if 'img_refresh_keys' not in st.session_state:
    st.session_state.img_refresh_keys = {}

# --- ฟังก์ชันจัดการอัปโหลดรายจุด (Fragment) ---
@st.fragment
def image_editor_fragment(idx, col, target_filename):
    refresh_key = st.session_state.img_refresh_keys.get(f"{idx}_{col}", 0)
    img_bytes = download_image_direct(target_filename)
    if img_bytes:
        st.image(img_bytes, caption=f"Drive: {target_filename}", use_container_width=True)
    else:
        st.warning(f"❌ ไม่พบรูป: {target_filename}")
    new_f = st.file_uploader(f"เปลี่ยนรูป {col}", type=['jpg','png','jpeg'], key=f"fu_{idx}_{col}_{refresh_key}")
    if new_f:
        with st.spinner("กำลังอัปโหลด..."):
            upload_and_overwrite(target_filename, new_f.getbuffer())
            st.session_state.img_refresh_keys[f"{idx}_{col}"] = refresh_key + 1
            st.toast("อัปโหลดสำเร็จ!"); time.sleep(0.5); st.rerun()

# --- ฟังก์ชันแสดงหน้าจอหลัก ---
def render_main_ui(center):
    with main_container.container():
        st.title(f"🚀 ศูนย์: {center}")
        df_idx = st.session_state.main_df[st.session_state.main_df['file_name'] == center].index
        # st.info("💡 ข้อมูลถูกกางออกเรียบร้อยแล้ว ตรวจสอบได้ทันที")
        
        for idx in df_idx:
            row = st.session_state.main_df.loc[idx]
            # ✅ expanded=True กางข้อมูลออกทั้งหมด
            with st.expander(f"📅 {row['date']} - {row['name']}", expanded=True):
                c = st.columns([2, 2, 1, 1])
                st.session_state.main_df.at[idx, 'name'] = c[0].text_input("ชื่อ", row['name'], key=f"n_{idx}")
                st.session_state.main_df.at[idx, 'status'] = c[1].text_input("ตำแหน่ง", row['status'], key=f"s_{idx}")
                st.session_state.main_df.at[idx, 'time_in'] = c[2].text_input("เข้า", row['time_in'], key=f"i_{idx}")
                st.session_state.main_df.at[idx, 'time_out'] = c[3].text_input("ออก", row['time_out'], key=f"o_{idx}")
                c_img = st.columns(2)
                with c_img[0]: image_editor_fragment(idx, "img_in1", str(row["img_in1"]))
                with c_img[1]: image_editor_fragment(idx, "img_out1", str(row["img_out1"]))

        st.divider()
        if st.button("🖨️ ออกรายงาน PDF", use_container_width=True, type="primary"):
            with st.spinner("กำลังสร้าง PDF..."):
                pdf = generate_pdf_original_style(st.session_state.main_df.loc[df_idx], center)
                st.download_button("📥 ดาวน์โหลด PDF", pdf, f"{center}.pdf", "application/pdf", use_container_width=True)

# --- ส่วนควบคุมหลัก ---
st.sidebar.title("เมนู")

# ✅ ปรับการเรียงลำดับศูนย์ (Sorting)
# ดึงค่า unique มาก่อน แล้วใช้ sorted() เพื่อเรียง 1, 2, 3...
# 1. ดึงชื่อศูนย์ทั้งหมดแบบ unique
unique_centers = st.session_state.main_df['file_name'].unique()

# 2. ใช้ฟังก์ชันเรียงลำดับแบบฉลาด (ดึงตัวเลขหน้าชื่อมาเรียง)
def natural_sort_key(s):
    try:
        # แยกข้อความตรง " - " แล้วเอาส่วนแรก (ตัวเลข) มาแปลงเป็น int
        return int(str(s).split('-')[0].strip())
    except:
        # ถ้าไม่มีตัวเลข หรือแยกไม่ได้ ให้เรียงตามข้อความปกติ
        return s

# 3. สั่งเรียงลำดับโดยใช้ key ที่เราสร้างไว้
centers = sorted(unique_centers, key=natural_sort_key)

# 4. นำไปใส่ใน selectbox เหมือนเดิม
sel_center = st.sidebar.selectbox("เลือกศูนย์", centers)

main_container = st.empty()

if sel_center:
    if "current_center" not in st.session_state or st.session_state.current_center != sel_center:
        st.session_state.current_center = sel_center
        progress_container = st.empty()
        with progress_container.container():
            st.markdown(f"### ✨ กำลังเตรียมข้อมูล: {sel_center}")
            prog_bar = st.progress(0)
            status_text = st.empty()
            target_df = st.session_state.main_df[st.session_state.main_df['file_name'] == sel_center]
            total_rows = len(target_df)
            for i, (idx, r) in enumerate(target_df.iterrows()):
                percent_complete = int(((i + 1) / total_rows) * 100)
                status_text.text(f"โหลดข้อมูลวันที่ {r['date']}... ({percent_complete}%)")
                download_image_direct(str(r['img_in1']))
                download_image_direct(str(r['img_out1']))
                prog_bar.progress(percent_complete)
            status_text.text("🚀 พร้อมตรวจสอบ!")
            time.sleep(0.3)
        progress_container.empty()
    render_main_ui(sel_center)

if st.sidebar.button("💾 บันทึก ", use_container_width=True):
    st.session_state.main_df.to_csv("03-2026.csv", index=False)
    st.sidebar.success("บันทึกสำเร็จ!")
