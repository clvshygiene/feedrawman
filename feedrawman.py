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
ADMIN_PASSWORD = "admin" # 您可以在此修改管理員密碼

# --- 3. Google API 連線設定 ---
@st.cache_resource
def get_gcp_credentials():
    # 設定權限範圍
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    # 從 Streamlit Cloud 的 Secrets 讀取金鑰
    try:
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
        return creds
    except Exception as e:
        st.error(f"無法讀取 Secrets 金鑰，請檢查 Streamlit 後台設定。錯誤: {e}")
        st.stop()

def get_gsheet_client():
    creds = get_gcp_credentials()
    return gspread.authorize(creds)

def get_gdrive_service():
    creds = get_gcp_credentials()
    return build('drive', 'v3', credentials=creds)

# 讀取 Google Sheet 資料
try:
    gc = get_gsheet_client()
    doc = gc.open_by_url(SHEET_URL)
    sheet_main = doc.worksheet("工作表1")
    sheet_log = doc.worksheet("領取日誌")
    
    # 讀取主要名單
    data = pd.DataFrame(sheet_main.get_all_records())
except Exception as e:
    st.error(f"❌ Google Sheets 連線失敗：{e}")
    st.info("請確認：1. Sheet 已共用給 Service Account 信箱 2. 工作表名稱為「工作表1」與「領取日誌」")
    st.stop()

# --- 4. 管理員功能：開啟新的一波領取 (側邊欄) ---
with st.sidebar:
    st.header("⚙️ 管理員選單")
    admin_input = st.text_input("管理密碼", type="password")
    if admin_input == ADMIN_PASSWORD:
        if st.button("📢 開啟全新一波領取 (重置狀態)"):
            num_rows = len(data) + 1 
            if num_rows > 1:
                cell_list = sheet_main.range(f'B2:B{num_rows}')
                for cell in cell_list:
                    cell.value = ''
                sheet_main.update_cells(cell_list)
                st.success("已重置所有學生領取狀態！")
                st.rerun()

# --- 5. 領取操作介面 (使用 Form 防止輸入時跳掉) ---
st.write("---")
st.subheader("🔍 領取登記登記")

# 使用表單包覆輸入與簽名，確保按下「送出」前畫面不會隨意跳轉
with st.form("collection_form", clear_on_submit=True):
    student_id = st.text_input("👉 請輸入或掃描學生證學號：")
    
    st.write("✍️ **學生簽名確認：**")
    # 簽名板優化：加入 update_freq 提高靈敏度，防止中斷
    canvas_result = st_canvas(
        fill_color="rgba(255, 255, 255, 0)",
        stroke_width=3,
        stroke_color="#000000",
        background_color="#eeeeee",
        height=150,
        update_freq=100, # 提高更新頻率，防止簽一半消失
        key="canvas",
    )
    
    submit_btn = st.form_submit_button("✅ 確認領取並存檔")

# --- 6. 表單送出後的邏輯處理 ---
if submit_btn:
    if not student_id:
        st.warning("請先輸入學號。")
    elif canvas_result.image_data is None:
        st.warning("請完成簽名再送出。")
    else:
        # 檢查學號是否存在於清單中
        student_info = data[data['學號'].astype(str) == str(student_id)]
        
        if student_info.empty:
            st.error(f"❌ 查無學號 {student_id}，請確認是否在申請名單內。")
        elif str(student_info.iloc[0].get('本次領取狀態', '')) == "已領取":
            st.error(f"⚠️ 學號 {student_id} 本波已經領取過，不可重複領取。")
        else:
            try:
                with st.spinner('正在更新紀錄，請稍候...'):
                    # A. 處理簽名圖片
                    img_data = canvas_result.image_data
                    img = Image.fromarray((img_data).astype('uint8'), mode='RGBA')
                    img_byte_arr = io.BytesIO()
                    img.save(img_byte_arr, format='PNG')
                    img_byte_arr.seek(0)

                    # B. 上傳至 Google Drive (修正 403 空間權限問題)
                    now = datetime.datetime.now()
                    now_str = now.strftime("%Y-%m-%d_%H-%M-%S")
                    file_name = f"簽名_{student_id}_{now_str}.png"
                    
                    drive_service = get_gdrive_service()
                    file_metadata = {
                        'name': file_name,
                        'parents': [DRIVE_FOLDER_ID]
                    }
                    media = MediaIoBaseUpload(img_byte_arr, mimetype='image/png', resumable=True)
                    
                    # 加上 supportsAllDrives=True 確保檔案能繼承資料夾空間權限
                    file = drive_service.files().create(
                        body=file_metadata, 
                        media_body=media, 
                        fields='id, webViewLink',
                        supportsAllDrives=True 
                    ).execute()
                    img_url = file.get('webViewLink')

                    # C. 更新 Google Sheet
                    # 找到該生在試算表中的列數
                    row_idx = int(student_info.index[0]) + 2
                    sheet_main.update_cell(row_idx, 2, "已領取")

                    # D. 寫入領取日誌
                    log_time = now.strftime("%Y-%m-%d %H:%M:%S")
                    sheet_log.append_row([str(student_id), log_time, img_url])

                st.balloons()
                st.success(f"🎉 學號 {student_id} 登記成功！請發放用品。")
                st.info(f"簽名檔已存至雲端：[點此查看]({img_url})")
                
            except Exception as e:
                st.error(f"💔 存檔過程發生錯誤：{e}")
                st.write("請檢查 Google Drive 資料夾權限是否已設為「編輯者」。")
