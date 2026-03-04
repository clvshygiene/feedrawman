import streamlit as st
from streamlit_drawable_canvas import st_canvas
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

# 1. 頁面設定
st.set_page_config(page_title="校園生理用品領取系統", layout="centered")
st.title("🌸 生理用品領取確認系統")

# 2. 連結 Google Sheets (使用 Secrets)
def get_gsheet_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    # 這裡從 Streamlit Secrets 讀取憑證，確保個資安全
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
    client = gspread.authorize(creds)
    return client

try:
    client = get_gsheet_client()
    # 替換成你的 Google Sheet 名稱或 ID
    sheet = client.open("生理用品發放名單").sheet1
    data = pd.DataFrame(sheet.get_all_records())
except Exception as e:
    st.error("系統連線失敗，請檢查權限設定。")
    st.stop()

# 3. 查詢介面
st.subheader("護理師核對區域")
student_id = st.text_input("請輸入學生學號進行查詢：")

if student_id:
    # 搜尋學號
    student_info = data[data['學號'].astype(str) == student_id]
    
    if not student_info.empty:
        st.success(f"確認符合領取資格！")
        # 僅顯示必要資訊，保護個資
        st.info(f"備註說明：{student_info.iloc[0]['原因']}")
        
        if student_info.iloc[0]['是否領取'] == "已領取":
            st.warning("提醒：該同學已於先前領取過。")
        
        # 4. 簽名板功能
        st.write("---")
        st.subheader("學生簽名確認")
        canvas_result = st_canvas(
            fill_color="rgba(255, 255, 255, 0)",
            stroke_width=3,
            stroke_color="#000000",
            background_color="#eeeeee",
            height=150,
            key="canvas",
        )

        # 5. 送出領取紀錄
        if st.button("確認領取並存檔"):
            if canvas_result.image_data is not None:
                # 這裡可以加入更新 Google Sheet 的邏輯
                # 找到對應的儲存格列號並更新
                row_idx = student_info.index[0] + 2 # +2 因為 header 且 index 從 0 開始
                sheet.update_cell(row_idx, 3, "已領取") # 假設第三欄是'是否領取'
                
                st.balloons()
                st.success("領取紀錄已成功更新！")
            else:
                st.error("請在上方簽名框簽名。")
    else:
        st.error("查無此學號，請再次確認。")