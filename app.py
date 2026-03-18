import streamlit as st
import pandas as pd
import os, gc, threading, time
from io import BytesIO
from PIL import Image, ExifTags

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from google.oauth2 import service_account

# ================= CONFIG =================
st.set_page_config(page_title="USO Report", layout="wide")

GOOGLE_DRIVE_FOLDER_ID = 'YOUR_FOLDER_ID'
pdf_lock = threading.Lock()

# ================= GOOGLE DRIVE =================
@st.cache_resource
def get_drive():
    try:
        if "gcp_service_account" in st.secrets:
            creds = service_account.Credentials.from_service_account_info(
                st.secrets["gcp_service_account"],
                scopes=['https://www.googleapis.com/auth/drive']
            )
        else:
            creds = service_account.Credentials.from_service_account_file(
                "service_account.json",
                scopes=['https://www.googleapis.com/auth/drive']
            )
        return build('drive', 'v3', credentials=creds)
    except:
        return None

# ================= DATA =================
@st.cache_data(ttl=300)
def load_data():
    return pd.read_csv("03-2026.csv").fillna("")

def get_center_df(center):
    df = load_data()
    return df[df['file_name'] == center].copy()

# ================= IMAGE =================
@st.cache_data(ttl=180, max_entries=30)
def load_image(file_name):
    service = get_drive()
    if not service or not file_name:
        return None
    try:
        q = f"'{GOOGLE_DRIVE_FOLDER_ID}' in parents and name='{file_name}'"
        res = service.files().list(q=q, fields="files(id)").execute()
        if not res['files']:
            return None

        file_id = res['files'][0]['id']
        req = service.files().get_media(fileId=file_id)

        fh = BytesIO()
        downloader = MediaIoBaseDownload(fh, req)

        done = False
        while not done:
            _, done = downloader.next_chunk()

        img = Image.open(BytesIO(fh.getvalue()))
        img.thumbnail((350, 350))

        out = BytesIO()
        img.convert("RGB").save(out, format="JPEG", quality=55)

        return out.getvalue()

    except:
        return None

def upload_image(name, file):
    service = get_drive()
    if not service:
        return

    media = MediaIoBaseUpload(BytesIO(file), mimetype='image/jpeg')
    service.files().create(
        body={"name": name, "parents": [GOOGLE_DRIVE_FOLDER_ID]},
        media_body=media
    ).execute()

# ================= PDF =================
def generate_pdf(df, center):
    buffer = BytesIO()

    doc = SimpleDocTemplate(buffer, pagesize=A4)

    styles = {
        "title": ParagraphStyle("t", fontSize=16, alignment=1),
        "body": ParagraphStyle("b", fontSize=10)
    }

    story = []

    story.append(Paragraph(f"ศูนย์: {center}", styles["title"]))
    story.append(Spacer(1, 10))

    table_data = [["วันที่","ชื่อ","เข้า","ออก"]]

    for _, r in df.iterrows():
        table_data.append([
            r["date"], r["name"], r["time_in"], r["time_out"]
        ])

    table = Table(table_data)
    table.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.5,colors.black)
    ]))

    story.append(table)

    doc.build(story)

    del story
    gc.collect()

    return buffer.getvalue()

def safe_pdf(df, center):
    with pdf_lock:
        return generate_pdf(df, center)

# ================= UI =================
st.sidebar.title("เมนู")

df_all = load_data()
centers = sorted(df_all['file_name'].unique())

center = st.sidebar.selectbox("เลือกศูนย์", centers)

if center:

    df = get_center_df(center)

    st.title(f"ศูนย์: {center}")

    # -------- Pagination --------
    page_size = 5
    page = st.number_input("หน้า", min_value=1, value=1)

    start = (page-1)*page_size
    end = start + page_size

    for idx, row in df.iloc[start:end].iterrows():

        with st.expander(f"{row['date']} - {row['name']}"):

            c = st.columns(4)

            row['name'] = c[0].text_input("ชื่อ", row['name'], key=f"n{idx}")
            row['status'] = c[1].text_input("ตำแหน่ง", row['status'], key=f"s{idx}")
            row['time_in'] = c[2].text_input("เข้า", row['time_in'], key=f"i{idx}")
            row['time_out'] = c[3].text_input("ออก", row['time_out'], key=f"o{idx}")

            col1, col2 = st.columns(2)

            with col1:
                if st.button("โหลดรูปเข้า", key=f"btn_in{idx}"):
                    img = load_image(row["img_in1"])
                    if img:
                        st.image(img)

                up = st.file_uploader("เปลี่ยนรูปเข้า", key=f"up_in{idx}")
                if up:
                    upload_image(row["img_in1"], up.getbuffer())
                    st.success("อัปโหลดแล้ว")

            with col2:
                if st.button("โหลดรูปออก", key=f"btn_out{idx}"):
                    img = load_image(row["img_out1"])
                    if img:
                        st.image(img)

                up = st.file_uploader("เปลี่ยนรูปออก", key=f"up_out{idx}")
                if up:
                    upload_image(row["img_out1"], up.getbuffer())
                    st.success("อัปโหลดแล้ว")

    # -------- PDF --------
    if st.button("สร้าง PDF"):
        with st.spinner("กำลังสร้าง..."):
            pdf = safe_pdf(df, center)

        st.download_button("ดาวน์โหลด", pdf, f"{center}.pdf")

# ================= SAVE =================
if st.sidebar.button("💾 บันทึก"):
    df_all.to_csv("03-2026.csv", index=False)
    st.sidebar.success("บันทึกแล้ว")
