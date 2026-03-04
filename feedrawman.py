import streamlit as st
from streamlit_drawable_canvas import st_canvas
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime

# 設定發放週期（例如：90天領一次）
RECYCLE_DAYS = 90

st.title("🌸 校園生理用品定期發放系統")

# --- Google Sheets 連線與讀取 ---
# (此處省略 credentials 設定，同前一次回覆)
# sheet = client.open("生理用品名冊").sheet1 

# 假設欄位是：學號, 最後領取日期 (格式為 YYYY-MM-DD)
# --------------------------------

student_id = st.text_input("請輸入學生學號：")

if student_id:
    # 搜尋學生資料
    student_info = data[data['學號'].astype(str) == student_id]
    
    if not student_info.empty:
        last_date_str = str(student_info.iloc[0]['最後領取日期'])
        
        # 判斷是否可以領取
        can_receive = False
        if not last_date_str or last_date_str == "None" or last_date_str == "":
            can_receive = True # 從未領過
            st.success("該同學符合資格（首次領取）")
        else:
            # 計算距離上次領取幾天
            last_date = datetime.strptime(last_date_str, "%Y-%m-%d")
            days_passed = (datetime.now() - last_date).days
            
            if days_passed >= RECYCLE_DAYS:
                can_receive = True
                st.success(f"符合資格（距離上次領取已過 {days_passed} 天）")
            else:
                st.error(f"不符合資格。上次領取日期：{last_date_str}，需再等 {RECYCLE_DAYS - days_passed} 天。")

        if can_receive:
            st.write("---")
            st.subheader("學生簽名")
            canvas_result = st_canvas(stroke_width=3, height=150, key="canvas")

            if st.button("確認領取"):
                if canvas_result.image_data is not None:
                    # 1. 更新 Google Sheet 的 '最後領取日期' 為今天
                    row_idx = student_info.index[0] + 2
                    today_str = datetime.now().strftime("%Y-%m-%d")
                    sheet.update_cell(row_idx, 2, today_str) 
                    
                    st.success(f"紀錄已更新！今日日期：{today_str}")
                    st.balloons()
                else:
                    st.warning("請先完成簽名")
    else:
        st.error("此學號不在名單內，請聯繫系統管理員。")
