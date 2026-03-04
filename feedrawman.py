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

# --- 2. 參數設定 (請替換為您的實際資料) ---
SHEET_URL = "https://docs.google.com/spreadsheets/d/13bMCf_cgdfByYH_DgUZynKHKZAHd2qmXJKyeCfCQOg8/"
# 👇 已經為您填入專屬的資料夾ID
DRIVE_FOLDER_ID = "19vaojtAD2eSmS7XSla431ryIUP2LG2LY" 
# 👇 這裡可以改成您自己好記的管理員密碼（用於一鍵清空發放狀態）
ADMIN_PASSWORD = "lls"

# --- 3. Google API 連線設定 ---
@st.cache_resource
def get_gcp_credentials():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    # 從 Streamlit Cloud 的 Secrets 讀取金鑰
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    return creds

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
    st.error("Google Sheets 連線失敗，請檢查 API 權限或網址。")
    st.stop()

# --- 4. 管理員功能：開啟新的一波領取 ---
with st.sidebar:
    st.header("⚙️ 衛生組管理員選單")
    st.write("當您有足夠庫存，想開啟新一波領取時，請輸入密碼重置狀態。")
    admin_input = st.text_input("輸入管理密碼", type="password")
    if admin_input == ADMIN_PASSWORD:
        st.success("身分確認：管理模式已開啟。")
        if st.button("📢 開啟全新一波領取 (清空所有狀態)"):
            num_rows = len(data) + 1 
            if num_rows > 1:
                # 清空 B 欄 (也就是「本次領取狀態」)
                cell_list = sheet_main.range(f'B2:B{num_rows}')
                for cell in cell_list:
                    cell.value = ''
                sheet_main.update_cells(cell_list)
                st.success("✅ 已成功重置所有學生的領取狀態！現在大家可以領取新的一波了。")
                st.rerun() # 重新整理頁面刷新資料

# --- 5. 護理師/學生 操作介面 ---
st.write("---")
st.subheader("🔍 學生身分核對區")
student_id = st.text_input("👉 請輸入或掃描學生證學號：")

if student_id:
    # 確保學號為字串格式進行比對
    student_info = data[data['學號'].astype(str) == str(student_id)]
    
    if not student_info.empty:
        # 取得該學生在 DataFrame 中的 index 以計算 Google Sheet 列數
        df_index = student_info.index[0]
        row_idx = int(df_index) + 2  # +2 是因為 df index 從 0 開始，且 Excel 有標題列
        
        status = str(student_info.iloc[0].get('本次領取狀態', ''))
        
        if status == "已領取":
            st.error(f"⚠️ 注意：學號 {student_id} 在這一波活動中 **已經領取過了**。")
        else:
            st.success(f"✅ 確認符合資格！學號 {student_id} 尚未領取。")
            
            st.write("---")
            st.subheader("✍️ 學生簽名確認")
            st.info("請同學在下方灰色區域簽名：")
            canvas_result = st_canvas(
                fill_color="rgba(255, 255, 255, 0)",
                stroke_width=3,
                stroke_color="#000000",
                background_color="#eeeeee",
                height=150,
                key="canvas",
            )

            if st.button("確認領取並存檔"):
                if canvas_result.image_data is not None:
                    try:
                        with st.spinner('正在上傳簽名與更新紀錄，請稍候...'):
                            # 1. 處理簽名圖片
                            img_data = canvas_result.image_data
                            img = Image.fromarray((img_data).astype('uint8'), mode='RGBA')
                            img_byte_arr = io.BytesIO()
                            img.save(img_byte_arr, format='PNG')
                            img_byte_arr.seek(0)

                            # 2. 上傳到 Google Drive
                            now_str = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                            file_name = f"簽名_{student_id}_{now_str}.png"
                            drive_service = get_gdrive_service()
                            file_metadata = {
                                'name': file_name,
                                'parents': [DRIVE_FOLDER_ID]
                            }
                            # 設定圖片上傳
                            media = MediaIoBaseUpload(img_byte_arr, mimetype='image/png', resumable=True)
                            file = drive_service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
                            img_url = file.get('webViewLink')

                            # 3. 更新「工作表1」的狀態為 "已領取"
                            sheet_main.update_cell(row_idx, 2, "已領取")

                            # 4. 寫入「領取日誌」
                            log_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            sheet_log.append_row([str(student_id), log_time, img_url])

                        st.balloons()
                        st.success("🎉 領取紀錄與簽名已成功存檔！護理師可以發放生理用品了。")
                        
                    except Exception as e:
                        st.error(f"存檔過程中發生錯誤：{e}")
                else:
                    st.warning("請在上方簽名框簽名後再點擊送出。")
    else:
        st.error("❌ 查無此學號，請確認學生是否在申請名單內。")
