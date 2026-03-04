import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from streamlit_drawable_canvas import st_canvas
import datetime
import io
from PIL import Image

st.set_page_config(page_title="衛生組生理用品發放系統", layout="centered")
st.title("🌸 校園生理用品領取系統")

SHEET_URL = "https://docs.google.com/spreadsheets/d/13bMCf_cgdfByYH_DgUZynKHKZAHd2qmXJKyeCfCQOg8/edit"
DRIVE_FOLDER_ID = "19vaojtAD2eSmS7XSla431ryIUP2LG2LY"

@st.cache_resource
def get_gcp_credentials():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    return Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)

def get_gsheet_client():
    return gspread.authorize(get_gcp_credentials())

def get_gdrive_service():
    return build("drive", "v3", credentials=get_gcp_credentials())

@st.cache_data(ttl=300)
def fetch_data(url):
    gc = get_gsheet_client()
    doc = gc.open_by_url(url)
    sheet = doc.worksheet("工作表1")
    return pd.DataFrame(sheet.get_all_records())

def canvas_to_png_bytes(canvas_image_data) -> bytes:
    """canvas RGBA ndarray -> PNG bytes"""
    img = Image.fromarray(canvas_image_data.astype("uint8"), mode="RGBA")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()

def upload_png_to_drive(png_bytes: bytes, student_id: str) -> tuple[str, str]:
    """回傳 (file_id, web_view_link)"""
    drive = get_gdrive_service()
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"signature_{student_id}_{ts}.png"

    media = MediaIoBaseUpload(io.BytesIO(png_bytes), mimetype="image/png", resumable=False)
    metadata = {"name": filename, "parents": [DRIVE_FOLDER_ID]}

    created = drive.files().create(
        body=metadata,
        media_body=media,
        fields="id,webViewLink",
        supportsAllDrives=True,
    ).execute()

    return created["id"], created["webViewLink"]

# 讀取資料
try:
    data = fetch_data(SHEET_URL)
except Exception as e:
    st.error(f"❌ 資料讀取失敗：{e}")
    st.stop()

# --- UI 美化（手機/平板友善） ---
st.markdown("""
<style>
.toolbar {border:1px solid rgba(0,0,0,.08); border-radius:16px; padding:12px; background:white;}
.canvas-wrap {border-radius:16px; overflow:hidden; border:1px solid rgba(0,0,0,.10);}
.small-label {font-size:12px; color:rgba(0,0,0,.6); margin-bottom:4px;}
</style>
""", unsafe_allow_html=True)

# --- session_state 初始化 ---
if "verified_student_id" not in st.session_state:
    st.session_state.verified_student_id = None
if "eligible" not in st.session_state:
    st.session_state.eligible = False
if "canvas_nonce" not in st.session_state:
    st.session_state.canvas_nonce = 0

st.write("---")
st.subheader("🔍 領取登記")

student_id = st.text_input("👉 請輸入或掃描學生證學號：").strip()

# --- 驗證學號 ---
if student_id:
    if "學號" not in data.columns:
        st.error("❌ 試算表缺少「學號」欄位")
        st.stop()

    student_info = data[data["學號"].astype(str) == str(student_id)]
    if student_info.empty:
        st.session_state.verified_student_id = None
        st.session_state.eligible = False
        st.error(f"❌ 查無學號 {student_id}")
    else:
        status = str(student_info.iloc[0].get("本次領取狀態", "")).strip()
        if status == "已領取":
            st.session_state.verified_student_id = None
            st.session_state.eligible = False
            st.warning(f"⚠️ 學號 {student_id} 已經領取過囉！")
        else:
            st.session_state.verified_student_id = str(student_id)
            st.session_state.eligible = True
            st.success("✅ 符合資格！請簽名。")

# --- 只有符合資格才顯示簽名板 ---
if st.session_state.eligible and st.session_state.verified_student_id == str(student_id) and student_id:
    st.markdown('<div class="toolbar">', unsafe_allow_html=True)

    c1, c2, c3, c4, c5 = st.columns([1.3, 1, 1, 1, 1])
    with c1:
        st.markdown('<div class="small-label">工具</div>', unsafe_allow_html=True)
        tool = st.selectbox(
            "tool",
            ["筆刷", "橡皮擦", "直線", "矩形", "圓形"],
            label_visibility="collapsed",
        )
    with c2:
        st.markdown('<div class="small-label">筆觸顏色</div>', unsafe_allow_html=True)
        stroke_color = st.color_picker("stroke", "#111111", label_visibility="collapsed")
    with c3:
        st.markdown('<div class="small-label">背景色</div>', unsafe_allow_html=True)
        bg_color = st.color_picker("bg", "#F1F5F9", label_visibility="collapsed")
    with c4:
        st.markdown('<div class="small-label">粗細</div>', unsafe_allow_html=True)
        stroke_width = st.slider("w", 1, 20, 4, label_visibility="collapsed")
    with c5:
        st.markdown('<div class="small-label">清除</div>', unsafe_allow_html=True)
        if st.button("🧹", use_container_width=True):
            st.session_state.canvas_nonce += 1
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

    # 工具映射
    if tool == "筆刷":
        drawing_mode = "freedraw"
        effective_stroke = stroke_color
    elif tool == "橡皮擦":
        drawing_mode = "freedraw"
        effective_stroke = bg_color
    elif tool == "直線":
        drawing_mode = "line"
        effective_stroke = stroke_color
    elif tool == "矩形":
        drawing_mode = "rect"
        effective_stroke = stroke_color
    else:
        drawing_mode = "circle"
        effective_stroke = stroke_color

    st.markdown('<div class="canvas-wrap">', unsafe_allow_html=True)
    canvas_result = st_canvas(
        fill_color="rgba(0,0,0,0)",
        stroke_width=stroke_width,
        stroke_color=effective_stroke,
        background_color=bg_color,
        height=240,
        drawing_mode=drawing_mode,
        update_streamlit=True,
        key=f"sig_canvas_{st.session_state.canvas_nonce}",
    )
    st.markdown("</div>", unsafe_allow_html=True)

    # --- 確認送出（Drive 存 PNG + Sheet 存連結）---
    if st.button("🚀 確認領取並存檔", use_container_width=True):
        if canvas_result.image_data is None:
            st.warning("請先簽名再點擊送出。")
            st.stop()

        try:
            with st.spinner("正在存檔中..."):
                # 1) 轉 PNG bytes
                png_bytes = canvas_to_png_bytes(canvas_result.image_data)

                # 2) 上傳 Drive
                file_id, web_link = upload_png_to_drive(png_bytes, student_id)

                # 3) 寫入 Sheet（先抓最新避免快取造成重複領取）
                latest = fetch_data(SHEET_URL)
                student_info2 = latest[latest["學號"].astype(str) == str(student_id)]
                if student_info2.empty:
                    st.error("❌ 資料已更新，找不到該學號，請重新輸入。")
                    st.stop()

                status2 = str(student_info2.iloc[0].get("本次領取狀態", "")).strip()
                if status2 == "已領取":
                    st.warning("⚠️ 剛剛已被登記領取（可能重複操作），請勿重複送出。")
                    st.stop()

                gc = get_gsheet_client()
                doc = gc.open_by_url(SHEET_URL)
                sheet_main = doc.worksheet("工作表1")
                sheet_log = doc.worksheet("領取日誌")

                row_idx = int(student_info2.index[0]) + 2
                sheet_main.update_cell(row_idx, 2, "已領取")

                now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                # ✅ 日誌只存 file_id + link（不存 base64）
                sheet_log.append_row([str(student_id), now, file_id, web_link])

            st.cache_data.clear()
            st.session_state.verified_student_id = None
            st.session_state.eligible = False
            st.session_state.canvas_nonce += 1

            st.success(f"🎉 登記成功！學號 {student_id} 已完成領取。")
            st.info("✅ 簽名已上傳 Google Drive，日誌已記錄檔案連結。")

        except Exception as e:
            st.error(f"💔 存檔失敗：{e}")