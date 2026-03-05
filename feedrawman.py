import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from streamlit_drawable_canvas import st_canvas
import datetime
import io
import time
from zoneinfo import ZoneInfo
from PIL import Image

# Email
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# PDF
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.utils import ImageReader

# ----------------------------
# 基本設定
# ----------------------------
TZ_TAIPEI = ZoneInfo("Asia/Taipei")

st.set_page_config(page_title="衛生組生理用品發放系統", layout="centered")
st.title("🌸 校園生理用品領取系統")

SHEET_URL = "https://docs.google.com/spreadsheets/d/13bMCf_cgdfByYH_DgUZynKHKZAHd2qmXJKyeCfCQOg8/edit"
DRIVE_FOLDER_ID = "1WPfK0coynKAwb15EfVp0st5VEGjNmUAl"

STUDENT_EMAIL_DOMAIN = "@g.clvs.tyc.edu.tw"
PICKED_TEXT = "已領取"

# ----------------------------
# Google API
# ----------------------------
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

@st.cache_data(ttl=120)
def fetch_sheet(url, worksheet_name):
    gc = get_gsheet_client()
    doc = gc.open_by_url(url)
    ws = doc.worksheet(worksheet_name)
    return pd.DataFrame(ws.get_all_records())

def get_worksheet(url, worksheet_name):
    gc = get_gsheet_client()
    doc = gc.open_by_url(url)
    return doc.worksheet(worksheet_name)

# ----------------------------
# 小工具：偵測狀態欄位
# ----------------------------
def detect_status_column(df: pd.DataFrame) -> str:
    """
    使用者沒有提供狀態欄位時：
    1) 優先找欄名同時包含「領取」與「狀態」
    2) 其次找包含「領取」或「狀態」
    3) 都找不到：用第 2 欄（通常是狀態）
    """
    cols = list(df.columns)
    for c in cols:
        if ("領取" in c) and ("狀態" in c):
            return c
    for c in cols:
        if ("領取" in c) or ("狀態" in c):
            return c
    if len(cols) >= 2:
        return cols[1]
    return cols[0]

# ----------------------------
# 簽名 / 上傳 / PDF
# ----------------------------
def canvas_to_png_bytes(canvas_image_data) -> bytes:
    img = Image.fromarray(canvas_image_data.astype("uint8"), mode="RGBA")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()

def upload_bytes_to_drive(file_bytes: bytes, filename: str, mimetype: str) -> tuple[str, str]:
    drive = get_gdrive_service()
    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mimetype, resumable=False)
    metadata = {"name": filename, "parents": [DRIVE_FOLDER_ID]}
    created = drive.files().create(
        body=metadata,
        media_body=media,
        fields="id,webViewLink",
        supportsAllDrives=True,
    ).execute()
    return created["id"], created["webViewLink"]

@st.cache_resource
def register_pdf_fonts():
    """
    若你 repo 有 fonts/NotoSansTC-Black.ttf 就會用中文黑體
    沒有也不會壞掉：會 fallback Helvetica
    """
    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        pdfmetrics.registerFont(TTFont("NotoSansTC-Black", "fonts/NotoSansTC-Black.ttf"))
        return True
    except Exception:
        return False

def make_receipt_pdf(student_id: str, ts_str: str, signature_png_bytes: bytes) -> bytes:
    ok_font = register_pdf_fonts()
    title_font = "NotoSansTC-Black" if ok_font else "Helvetica-Bold"
    text_font = "NotoSansTC-Black" if ok_font else "Helvetica"

    buf = io.BytesIO()
    c = pdf_canvas.Canvas(buf, pagesize=A4)
    width, height = A4

    c.setFont(title_font, 18)
    c.drawString(40, height - 60, "生理用品領取簽收單")

    c.setFont(text_font, 12)
    c.drawString(40, height - 98, f"學號：{student_id}")
    c.drawString(40, height - 120, f"時間：{ts_str}")

    c.setFont(text_font, 12)
    c.drawString(40, height - 160, "簽名：")

    sig_img = ImageReader(io.BytesIO(signature_png_bytes))
    img_w = 420
    img_h = 170
    x = 40
    y = height - 160 - img_h - 10

    c.rect(x, y, img_w, img_h, stroke=1, fill=0)
    c.drawImage(sig_img, x, y, width=img_w, height=img_h, mask="auto")

    c.setFont(text_font, 10)
    c.drawString(40, 40, "本簽收單由系統自動產生")

    c.showPage()
    c.save()
    return buf.getvalue()

# ----------------------------
# SMTP 寄信
# ----------------------------
def send_email(to_addr: str, subject: str, body: str) -> None:
    """
    需要 secrets.toml:
    [smtp]
    host="smtp.gmail.com"
    port=587
    username="你的寄件帳號"
    password="16碼金鑰"
    from_addr="同 username"
    from_name="衛生組"
    use_tls=true
    """
    cfg = st.secrets["smtp"]
    host = cfg["host"]
    port = int(cfg.get("port", 587))
    username = cfg.get("username", "")
    password = cfg.get("password", "")
    from_addr = cfg.get("from_addr", username)
    from_name = cfg.get("from_name", "衛生組")
    use_tls = bool(cfg.get("use_tls", True))

    msg = MIMEMultipart()
    msg["From"] = f"{from_name} <{from_addr}>"
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))

    with smtplib.SMTP(host, port, timeout=25) as server:
        server.ehlo()
        if use_tls:
            server.starttls()
            server.ehlo()
        if username and password:
            server.login(username, password)
        server.send_message(msg)

# ----------------------------
# 可選：庫存（若沒有「庫存」工作表就跳過）
# ----------------------------
def try_load_inventory():
    try:
        inv = fetch_sheet(SHEET_URL, "庫存")
        if inv.empty:
            return None, None
        # 預期欄位：品項 / 庫存 / 安全庫存 (可自行改)
        return inv, "庫存"
    except Exception:
        return None, None

def try_update_inventory(delta: int = -1) -> tuple[bool, str]:
    """
    預設：每成功領取扣 1
    工作表「庫存」第 1 列第一筆品項當作總庫存
    欄位：庫存、(可選)安全庫存
    """
    try:
        ws = get_worksheet(SHEET_URL, "庫存")
        df = pd.DataFrame(ws.get_all_records())
        if df.empty:
            return False, "庫存表是空的"

        # 偵測庫存欄位
        stock_col = None
        for c in df.columns:
            if "庫存" in c:
                stock_col = c
                break
        if not stock_col:
            return False, "庫存表找不到「庫存」欄位"

        cur = df.iloc[0].get(stock_col, 0)
        try:
            cur = int(cur)
        except Exception:
            return False, f"庫存欄位不是數字：{cur}"

        new_val = cur + delta
        if new_val < 0:
            return False, "庫存不足，無法扣帳"

        # 更新到試算表：第一筆資料的庫存欄位
        # ws.get_all_records() 會從第 2 列開始，所以 df.iloc[0] 對應到 sheet row=2
        row = 2
        col = list(df.columns).index(stock_col) + 1
        ws.update_cell(row, col, new_val)
        return True, f"庫存 {cur} → {new_val}"
    except Exception as e:
        return False, f"庫存更新失敗：{e}"

# ----------------------------
# 讀取主表
# ----------------------------
try:
    main_df = fetch_sheet(SHEET_URL, "工作表1")
except Exception as e:
    st.error(f"❌ 資料讀取失敗：{e}")
    st.stop()

if main_df.empty:
    st.error("❌ 工作表1 沒有資料")
    st.stop()

if "學號" not in main_df.columns:
    st.error("❌ 工作表1 缺少「學號」欄位")
    st.stop()

STATUS_COL = detect_status_column(main_df)

# ----------------------------
# Session state
# ----------------------------
if "verified_student_id" not in st.session_state:
    st.session_state.verified_student_id = None
if "eligible" not in st.session_state:
    st.session_state.eligible = False
if "canvas_nonce" not in st.session_state:
    st.session_state.canvas_nonce = 0
if "student_id_input" not in st.session_state:
    st.session_state.student_id_input = ""
if "sig_landscape" not in st.session_state:
    st.session_state.sig_landscape = True  # 預設大簽名
if "sig_height" not in st.session_state:
    st.session_state.sig_height = 420

# ----------------------------
# Sidebar：功能選單
# ----------------------------
with st.sidebar:
    st.header("⚙️ 管理功能")

    # 快速看狀態欄位偵測結果
    st.caption(f"狀態欄位：{STATUS_COL}")

    # 庫存區（可選）
    inv_df, _ = try_load_inventory()
    if inv_df is None:
        st.info("（可選）要啟用庫存：請在試算表新增工作表「庫存」\n建議欄位：品項、庫存、安全庫存")
    else:
        # 顯示第一筆庫存
        stock_col = None
        safety_col = None
        for c in inv_df.columns:
            if "安全" in c:
                safety_col = c
            if c == "庫存" or ("庫存" in c):
                stock_col = stock_col or c

        if stock_col:
            cur_stock = inv_df.iloc[0].get(stock_col, "")
            st.metric("目前庫存", cur_stock)
            if safety_col:
                safety = inv_df.iloc[0].get(safety_col, "")
                st.caption(f"安全庫存：{safety}")
        else:
            st.warning("庫存表找不到「庫存」欄位")

    st.divider()

    # 寄信區
    st.header("✉️ 未領取自動寄信")

    # 管理密碼（可選）
    admin_ok = True
    if "admin" in st.secrets and "password" in st.secrets["admin"]:
        pwd = st.text_input("管理密碼", type="password")
        admin_ok = (pwd == st.secrets["admin"]["password"])

    if not admin_ok:
        st.warning("需要管理密碼才能寄信")
    else:
        mode = st.selectbox("寄送模式", ["寄給所有未領取", "寄給指定學號"])

        subject = st.text_input("主旨", value="生理用品領取提醒")
        body = st.text_area(
            "內容（可用 {student_id}）",
            value="同學您好：\n\n提醒您可至衛生組領取生理用品。\n學號：{student_id}\n\n謝謝。",
            height=170,
        )
        throttle = st.slider("每封間隔(秒)（避免被鎖）", 0.3, 2.0, 0.8, 0.1)
        max_send = st.number_input("本次最多寄送封數", min_value=1, max_value=500, value=80)

        ids_to_send = []
        if mode == "寄給指定學號":
            ids_text = st.text_area("學號（每行一個）", height=120)
            ids_to_send = [x.strip() for x in ids_text.splitlines() if x.strip()]
        else:
            # 所有未領取
            tmp = main_df.copy()
            tmp_status = tmp[STATUS_COL].astype(str).str.strip()
            not_picked = tmp[tmp_status != PICKED_TEXT]
            ids_to_send = not_picked["學號"].astype(str).tolist()

        if st.button("📨 開始寄送", use_container_width=True):
            if "smtp" not in st.secrets:
                st.error("secrets.toml 缺少 [smtp] 設定")
            elif not ids_to_send:
                st.error("沒有找到收件人")
            else:
                ids_to_send = list(dict.fromkeys(ids_to_send))  # 去重
                ids_to_send = ids_to_send[: int(max_send)]

                ok = 0
                failed = []

                prog = st.progress(0)
                for i, sid in enumerate(ids_to_send, start=1):
                    to_addr = f"{sid}{STUDENT_EMAIL_DOMAIN}"
                    try:
                        send_email(
                            to_addr=to_addr,
                            subject=subject.format(student_id=sid),
                            body=body.format(student_id=sid),
                        )
                        ok += 1
                    except Exception as e:
                        failed.append((sid, str(e)))

                    prog.progress(i / len(ids_to_send))
                    time.sleep(float(throttle))

                st.success(f"✅ 已寄出 {ok} 封")
                if failed:
                    st.error(f"❌ 失敗 {len(failed)} 封（顯示前 10 筆）")
                    for sid, err in failed[:10]:
                        st.write(f"- {sid}: {err}")

    st.divider()

    # 儀表板切換
    show_dashboard = st.toggle("📊 顯示統計儀表板", value=True)

# ----------------------------
# 儀表板
# ----------------------------
if "show_dashboard" not in locals():
    show_dashboard = True

if show_dashboard:
    st.subheader("📊 領取統計儀表板")

    # 已領取/未領取
    status_series = main_df[STATUS_COL].astype(str).str.strip()
    picked = (status_series == PICKED_TEXT).sum()
    total = len(main_df)
    not_picked = total - picked
    rate = (picked / total * 100) if total else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("總人數", total)
    c2.metric("已領取", picked)
    c3.metric("未領取", not_picked)
    c4.metric("領取率", f"{rate:.1f}%")

    # 從「領取日誌」拉趨勢（若沒有就跳過）
    try:
        log_df = fetch_sheet(SHEET_URL, "領取日誌")
    except Exception:
        log_df = pd.DataFrame()

    if not log_df.empty and len(log_df.columns) >= 2:
        # 預期第 2 欄是時間字串
        time_col = log_df.columns[1]
        # parse 成台灣時間日期
        def to_date(s):
            try:
                dt = datetime.datetime.strptime(str(s).strip(), "%Y-%m-%d %H:%M:%S").replace(tzinfo=TZ_TAIPEI)
                return dt.date()
            except Exception:
                return None

        log_df["_date"] = log_df[time_col].apply(to_date)
        daily = log_df.dropna(subset=["_date"]).groupby("_date").size().reset_index(name="count")
        daily = daily.sort_values("_date")

        # 近 7 日
        today = datetime.datetime.now(TZ_TAIPEI).date()
        start = today - datetime.timedelta(days=6)
        daily_7 = daily[daily["_date"] >= start]

        st.line_chart(daily_7.set_index("_date")["count"])

        # 今日人次
        today_count = int((log_df["_date"] == today).sum())
        st.caption(f"今日領取人次：{today_count}")
    else:
        st.info("尚未建立「領取日誌」或格式不足，暫不顯示趨勢圖。")

st.write("---")
st.subheader("✍️ 領取登記（手動輸入學號）")

# ----------------------------
# 登記區：輸入學號 -> 顯示簽名板
# ----------------------------
student_id = st.text_input("👉 請輸入學生證學號：", value=st.session_state.student_id_input).strip()
st.session_state.student_id_input = student_id

if student_id:
    info = main_df[main_df["學號"].astype(str) == str(student_id)]

    if info.empty:
        st.session_state.verified_student_id = None
        st.session_state.eligible = False
        st.error(f"❌ 查無學號 {student_id}")
    else:
        status = str(info.iloc[0].get(STATUS_COL, "")).strip()
        if status == PICKED_TEXT:
            st.session_state.verified_student_id = None
            st.session_state.eligible = False
            st.warning(f"⚠️ 學號 {student_id} 已經領取過囉！")
        else:
            # 庫存不足就阻擋（如果有庫存表）
            inv_df, _ = try_load_inventory()
            if inv_df is not None:
                stock_col = None
                for c in inv_df.columns:
                    if c == "庫存" or ("庫存" in c):
                        stock_col = c
                        break
                if stock_col:
                    try:
                        cur_stock = int(inv_df.iloc[0].get(stock_col, 0))
                    except Exception:
                        cur_stock = 0
                    if cur_stock <= 0:
                        st.error("❌ 目前庫存不足，請先補貨後再登記。")
                        st.stop()

            st.session_state.verified_student_id = str(student_id)
            st.session_state.eligible = True
            st.success("✅ 符合資格！請簽名。")

# ----------------------------
# 簽名板（可切橫向/加大）
# ----------------------------
if st.session_state.eligible and st.session_state.verified_student_id == str(student_id) and student_id:
    st.markdown("""
    <style>
    .toolbar {border:1px solid rgba(0,0,0,.08); border-radius:16px; padding:12px; background:white;}
    .canvas-wrap {border-radius:16px; overflow:hidden; border:1px solid rgba(0,0,0,.10);}
    .small-label {font-size:12px; color:rgba(0,0,0,.6); margin-bottom:4px;}
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="toolbar">', unsafe_allow_html=True)

    # 工具列
    c1, c2, c3, c4, c5, c6 = st.columns([1.2, 1, 1, 1, 1, 1.2])
    with c1:
        st.markdown('<div class="small-label">工具</div>', unsafe_allow_html=True)
        tool = st.selectbox("tool", ["筆刷", "橡皮擦", "直線", "矩形", "圓形"], label_visibility="collapsed")
    with c2:
        st.markdown('<div class="small-label">筆色</div>', unsafe_allow_html=True)
        stroke_color = st.color_picker("stroke", "#111111", label_visibility="collapsed")
    with c3:
        st.markdown('<div class="small-label">背景</div>', unsafe_allow_html=True)
        bg_color = st.color_picker("bg", "#FFFFFF", label_visibility="collapsed")
    with c4:
        st.markdown('<div class="small-label">粗細</div>', unsafe_allow_html=True)
        stroke_width = st.slider("w", 1, 20, 6, label_visibility="collapsed")
    with c5:
        st.markdown('<div class="small-label">簽名大小</div>', unsafe_allow_html=True)
        size_mode = st.selectbox("size", ["大(推薦)", "超大"], label_visibility="collapsed")
    with c6:
        st.markdown('<div class="small-label">清除</div>', unsafe_allow_html=True)
        if st.button("🧹 清空簽名", use_container_width=True):
            st.session_state.canvas_nonce += 1
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

    # 大小設定：直式也給你「橫向手感」的大區域
    if size_mode == "大(推薦)":
        canvas_h = 420
    else:
        canvas_h = 520

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
        height=canvas_h,
        drawing_mode=drawing_mode,
        update_streamlit=True,
        key=f"sig_canvas_{st.session_state.canvas_nonce}",
    )
    st.markdown("</div>", unsafe_allow_html=True)

    # ----------------------------
    # 存檔
    # ----------------------------
    if st.button("🚀 確認領取並存檔", use_container_width=True):
        if canvas_result.image_data is None:
            st.warning("請先簽名再點擊送出。")
            st.stop()

        try:
            with st.spinner("正在存檔中..."):
                # 抓最新狀態避免快取
                latest = fetch_sheet(SHEET_URL, "工作表1")
                status_col_latest = detect_status_column(latest)

                info2 = latest[latest["學號"].astype(str) == str(student_id)]
                if info2.empty:
                    st.error("❌ 資料已更新，找不到該學號，請重新輸入。")
                    st.stop()

                status2 = str(info2.iloc[0].get(status_col_latest, "")).strip()
                if status2 == PICKED_TEXT:
                    st.warning("⚠️ 剛剛已被登記領取（可能重複操作），請勿重複送出。")
                    st.stop()

                # 1) PNG
                png_bytes = canvas_to_png_bytes(canvas_result.image_data)

                ts = datetime.datetime.now(TZ_TAIPEI)
                ts_str = ts.strftime("%Y-%m-%d %H:%M:%S")
                ts_file = ts.strftime("%Y%m%d_%H%M%S")

                png_name = f"signature_{student_id}_{ts_file}.png"
                png_id, png_link = upload_bytes_to_drive(png_bytes, png_name, "image/png")

                # 2) PDF
                pdf_bytes = make_receipt_pdf(student_id, ts_str, png_bytes)
                pdf_name = f"receipt_{student_id}_{ts_file}.pdf"
                pdf_id, pdf_link = upload_bytes_to_drive(pdf_bytes, pdf_name, "application/pdf")

                # 3) 寫回 Sheet
                ws_main = get_worksheet(SHEET_URL, "工作表1")
                ws_log = get_worksheet(SHEET_URL, "領取日誌")

                row_idx = int(info2.index[0]) + 2  # df index -> sheet row
                # 狀態欄位位置（用最新 df 的欄位）
                status_col_index = list(latest.columns).index(status_col_latest) + 1
                ws_main.update_cell(row_idx, status_col_index, PICKED_TEXT)

                ws_log.append_row([str(student_id), ts_str, png_id, png_link, pdf_id, pdf_link])

                # 4) 扣庫存（若有庫存表）
                inv_df, _ = try_load_inventory()
                if inv_df is not None:
                    ok, msg = try_update_inventory(delta=-1)
                    if not ok:
                        st.warning(f"⚠️ 已完成登記，但庫存扣帳失敗：{msg}")

            st.cache_data.clear()
            st.session_state.verified_student_id = None
            st.session_state.eligible = False
            st.session_state.canvas_nonce += 1
            st.session_state.student_id_input = ""

            st.success(f"🎉 登記成功！學號 {student_id} 已完成領取。")
            st.info("✅ 已上傳簽名 PNG 與簽收 PDF 到 Google Drive，並寫入日誌連結。")

        except Exception as e:
            st.error(f"💔 存檔失敗：{e}")
