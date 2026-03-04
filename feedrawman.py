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

# --- 1. 系統設定 ---
st.set_page_config(page_title="衛生組生理用品發放系統", layout="centered")
st.title("🌸 校園生理用品領取系統")

# --- 2. 參數設定 ---
SHEET_URL = "https://docs.google.com/spreadsheets/d/13bMCf_cgdfByYH_DgUZynKHKZAHd2qmXJKyeCfCQOg8/edit"
DRIVE_FOLDER_ID = "19vaojtAD2eSmS7XSla431ryIUP2LG2LY" 
ADMIN_PASSWORD = "admin" 

# --- 3. Google API 連線設定 ---
@st.cache_resource
def get_gcp_credentials():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    try:
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
        return creds
    except Exception as e:
        st.error(f"無法讀取 Secrets 金鑰。錯誤: {e}")
        st.stop()

def get_gsheet_client():
    return gspread.authorize(get_gcp_credentials())

def get_gdrive_service():
    return build('drive', 'v3', credentials=get_gcp_credentials())

# 讀取 Google Sheet 資料
try:
    gc = get_gsheet_client()
    doc = gc.open_by_url(SHEET_URL)
    sheet_main = doc.worksheet("工作表1")
    sheet_log = doc.worksheet("領取日誌")
    data = pd.DataFrame(sheet_main.get_all_records())
except Exception as e:
    st.error(f"❌ Google Sheets 連線失敗：{e}")
    st.stop()

# --- 4. 管理員功能 ---
with st.sidebar:
    st.header("⚙️ 管理員選單")
    admin_input = st.text_input("管理密碼", type="password")
    if admin_input == ADMIN_PASSWORD:
        if st.button("📢 重置所有學生領取狀態"):
            num_rows = len(data) + 1 
            cell_list = sheet_main.range(f'B2:B{num_rows}')
            for cell in cell_list: cell.value = ''
            sheet_main.update_cells(cell_list)
            st.success("已重置狀態！")
            st.rerun()

# --- 5. 領取操作介面 ---
st.write("---")
st.subheader("🔍 領取登記登記")

# 步驟 1：輸入學號
student_id = st.text_input("👉 請輸入或掃描學生證學號：", key="id_input")

if student_id:
    # 檢查學號
    student_info = data[data['學號'].astype(str) == str(student_id)]
    
    if student_info.empty:
        st.error(f"❌ 查無學號 {student_id}")
    elif str(student_info.iloc[0].get('本次領取狀態', '')) == "已領取":
        st.warning(f"⚠️ 學號 {student_id} 已經領取過了。")
    else:
        st.success(f"✅ 符合資格！請在下方簽名。")
        
        # 步驟 2：簽名板 (移除 st.form，確保穩定)
        st.write("✍️ **學生簽名確認：**")
        canvas_result = st_canvas(
            fill_color="rgba(255, 255, 255, 0)",
            stroke_width=3,
            stroke_color="#000000",
            background_color="#eeeeee",
            height=150,
            drawing_mode="freedraw", # 明確設定繪圖模式
            key="canvas_main",
        )

        # 步驟 3：送出按鈕
        if st.button("🚀 確認領取並存檔"):
            if canvas_result.image_data is not None:
                try:
                    with st.spinner('正在存檔中...'):
                        # 處理圖片
                        img_data = canvas_result.image_data
                        img = Image.fromarray((img_data).astype('uint8'), mode='RGBA')
                        img_byte_arr = io.BytesIO()
                        img.save(img_byte_arr, format='PNG')
                        img_byte_arr.seek(0)

                        # 上傳 Drive
                        now = datetime.datetime.now()
                        file_name = f"簽名_{student_id}_{now.strftime('%Y%m%d_%H%M%S')}.png"
                        drive_service = get_gdrive_service()
                        file_metadata = {'name': file_name, 'parents': [DRIVE_FOLDER_ID]}
                        media = MediaIoBaseUpload(img_byte_arr, mimetype='image/png', resumable=True)
                        file = drive_service.files().create(
                            body=file_metadata, media_body=media, fields='id, webViewLink', supportsAllDrives=True 
                        ).execute()
                        img_url = file.get('webViewLink')

                        # 更新 Sheet
                        row_idx = int(student_info.index[0]) + 2
                        sheet_main.update_cell(row_idx, 2, "已領取")
                        sheet_log.append_row([str(student_id), now.strftime("%Y-%m-%d %H:%M:%S"), img_url])

                    st.balloons()
                    st.success(f"🎉 登記成功！請發放用品。")
                    st.info(f"簽名檔：[點此查看]({img_url})")
                except Exception as e:
                    st.error(f"💔 發生錯誤：{e}")
            else:
                st.warning("請先簽名再點擊送出。")
