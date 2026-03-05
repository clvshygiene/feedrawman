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
import os
from zoneinfo import ZoneInfo
from PIL import Image, ImageDraw, ImageFont

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
WATERMARK_TEXT = "衛生組專用"

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
# 狀態欄位偵測
# ----------------------------
def detect_status_column(df: pd.DataFrame) -> str:
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
# 浮水印字體
# ----------------------------
@st.cache_resource
def get_watermark_font(img_width: int):
    """
    優先使用 repo 內 fonts/NotoSansTC-Black.ttf
    找不到則 fallback（中文可能不漂亮，但不會炸）
    """
    size = max(18, int(img_width * 0.06))
    for fp in [
        os.path.join("fonts", "NotoSansTC-Black.ttf"),
        os.path.join("fonts", "NotoSansTC-Black.otf"),
    ]:
        if os.path.exists(fp):
            return ImageFont.truetype(fp, size)
    return ImageFont.load_default()

# ----------------------------
# 簽名：強制旋轉 + 浮水印
# ----------------------------
def canvas_to_png_bytes(canvas_image_data) -> bytes:
    """
    1) canvas -> PNG
    2) 強制旋轉 90 度（學生橫拿手機簽名 -> 存成直的）
    3) 加浮水印：衛生組專用（上方半透明）
    """
    img = Image.fromarray(canvas_image_data.astype("uint8"), mode="RGBA")

    # ✅ 強制旋轉（每次都轉）
    img = img.rotate(90, expand=True)
    w, h = img.size

    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)

    font = get_watermark_font(w)
    text = WATERMARK_TEXT

    # 置中上方
    bbox = draw.textbbox((0, 0), text, font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    x = (w - tw) // 2
    y = max(8, int(h * 0.03))

    draw.text((x, y), text, font=font, fill=(0, 0, 0, 80))
    img = Image.alpha_composite(img, overlay)

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
    若 repo 有 fonts/NotoSansTC-Black.ttf：PDF 文字用中文黑體
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
    c.drawString(40, height - 160, "簽名（含浮水印）：")

    sig_img = ImageReader(io.BytesIO(signature_png_bytes))
    img_w = 420
    img_h = 190
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
    secrets.toml:
    [smtp]
    host="smtp.gmail.com"
    port=587
    username="你的寄件帳號"
    password="16碼金鑰(去空格)"
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
            return None
        return inv
    except Exception:
        return None

def try_update_inventory(delta: int = -1) -> tuple[bool, str]:
    """
    預設：每成功領取扣 1
    「庫存」工作表第一筆資料的「庫存」欄位扣帳
    """
    try:
        ws = get_worksheet(SHEET_URL, "庫存")
        df = pd.DataFrame(ws.get_all_records())
        if df.empty:
            return False, "庫存表是空的"

        stock_col = None
        for c in df.columns:
            if c == "庫存" or ("庫存" in c):
                stock_col = c
                break
        if not stock_col:
            return False, "庫存表找不到「庫存」欄位"

        cur = df.iloc[0].get(stock_col, 0)
        cur = int(cur)
        new_val = cur + delta
        if new_val < 0:
            return False, "庫存不足，無法扣帳"

        # df.iloc[0] 對應 sheet row=2
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

# ----------------------------
# Sidebar：管理功能（寄信/儀表板/庫存狀態）
# ----------------------------
with st.sidebar:
    st.header("⚙️ 管理功能")
    st.caption(f"狀態欄位（自動偵測）：{STATUS_COL}")

    inv_df = try_load_inventory()
    if inv_df is None:
        st.info("（可選）要啟用庫存：請新增工作表「庫存」\n建議欄位：品項、庫存、安全庫存")
    else:
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
                st.caption(f"安全庫存：{inv_df.iloc[0].get(safety_col, '')}")
        else:
            st.warning("庫存表找不到「庫存」欄位")

    st.divider()

    st.header("✉️ 未領取自動寄信")

    admin_ok = True
    if "admin" in st.secrets and "password" in st.secrets["admin"]:
        pwd = st.text_input("管理密碼", type="password")
        admin_ok = (pwd == st.secrets["admin"]["password"])

    if not admin_ok:
        st.warning("需要管理密碼才能寄信")
        show_dashboard = st.toggle("📊 顯示統計儀表板", value=True)
    else:
        show_dashboard = st.toggle("📊 顯示統計儀表板", value=True)

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

                prog = st.progress(0.0)
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

# ----------------------------
# 儀表板
# ----------------------------
if show_dashboard:
    st.subheader("📊 領取統計儀表板")

    status_series = main_df[STATUS_COL].astype(str).str.strip()
    picked = int((status_series == PICKED_TEXT).sum())
    total = int(len(main_df))
    not_picked = total - picked
    rate = (picked / total * 100) if total else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("總人數", total)
    c2.metric("已領取", picked)
    c3.metric("未領取", not_picked)
    c4.metric("領取率", f"{rate:.1f}%")

    # 趨勢圖（可選）
    try:
        log_df = fetch_sheet(SHEET_URL, "領取日誌")
    except Exception:
        log_df = pd.DataFrame()

    if not log_df.empty and len(log_df.columns) >= 2:
        time_col = log_df.columns[1]

        def to_date(s):
            try:
                dt = datetime.datetime.strptime(str(s).strip(), "%Y-%m-%d %H:%M:%S").replace(tzinfo=TZ_TAIPEI)
                return dt.date()
            except Exception:
                return None

        log_df["_date"] = log_df[time_col].apply(to_date)
        daily = log_df.dropna(subset=["_date"]).groupby("_date").size().reset_index(name="count").sort_values("_date")

        today = datetime.datetime.now(TZ_TAIPEI).date()
        start = today - datetime.timedelta(days=6)
        daily_7 = daily[daily["_date"] >= start]

        st.line_chart(daily_7.set_index("_date")["count"])
        today_count = int((log_df["_date"] == today).sum())
        st.caption(f"今日領取人次：{today_count}")
    else:
        st.info("尚未建立「領取日誌」或格式不足，暫不顯示趨勢圖。")

st.write("---")
st.subheader("✍️ 領取登記（手動輸入學號）")
st.info("📱 建議把手機橫過來簽名（系統會自動轉正並加浮水印）")

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
            inv_df2 = try_load_inventory()
            if inv_df2 is not None:
                stock_col = None
                for c in inv_df2.columns:
                    if c == "庫存" or ("庫存" in c):
                        stock_col = c
                        break
                if stock_col:
                    try:
                        cur_stock = int(inv_df2.iloc[0].get(stock_col, 0))
                    except Exception:
                        cur_stock = 0
                    if cur_stock <= 0:
                        st.error("❌ 目前庫存不足，請先補貨後再登記。")
                        st.stop()

            st.session_state.verified_student_id = str(student_id)
            st.session_state.eligible = True
            st.success("✅ 符合資格！請簽名。")

# ----------------------------
# 簽名板：只留筆刷 + 高畫布
# ----------------------------
if st.session_state.eligible and st.session_state.verified_student_id == str(student_id) and student_id:
    st.markdown("""
    <style>
    .canvas-wrap {border-radius:16px; overflow:hidden; border:1px solid rgba(0,0,0,.10);}
    </style>
    """, unsafe_allow_html=True)

    # ✅ 只留筆刷：固定黑筆、白底、粗細固定
    stroke_color = "#111111"
    bg_color = "#FFFFFF"
    stroke_width = 8

    # ✅ 畫布加高（你可調 650~780）
    canvas_h = 700

    st.markdown('<div class="canvas-wrap">', unsafe_allow_html=True)
    canvas_result = st_canvas(
        fill_color="rgba(0,0,0,0)",
        stroke_width=stroke_width,
        stroke_color=stroke_color,
        background_color=bg_color,
        height=canvas_h,
        drawing_mode="freedraw",
        update_streamlit=True,
        key=f"sig_canvas_{st.session_state.canvas_nonce}",
    )
    st.markdown("</div>", unsafe_allow_html=True)

    cbtn1, cbtn2 = st.columns([1, 1])
    with cbtn1:
        if st.button("🧹 清空簽名", use_container_width=True):
            st.session_state.canvas_nonce += 1
            st.rerun()
    with cbtn2:
        submit = st.button("🚀 確認領取並存檔", use_container_width=True)

    if submit:
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

                # 1) PNG（已強制轉正 + 浮水印）
                png_bytes = canvas_to_png_bytes(canvas_result.image_data)

                ts = datetime.datetime.now(TZ_TAIPEI)
                ts_str = ts.strftime("%Y-%m-%d %H:%M:%S")
                ts_file = ts.strftime("%Y%m%d_%H%M%S")

                png_name = f"signature_{student_id}_{ts_file}.png"
                png_id, png_link = upload_bytes_to_drive(png_bytes, png_name, "image/png")

                # 2) PDF（用已處理過的 png_bytes）
                pdf_bytes = make_receipt_pdf(student_id, ts_str, png_bytes)
                pdf_name = f"receipt_{student_id}_{ts_file}.pdf"
                pdf_id, pdf_link = upload_bytes_to_drive(pdf_bytes, pdf_name, "application/pdf")

                # 3) 寫回 Sheet
                ws_main = get_worksheet(SHEET_URL, "工作表1")
                ws_log = get_worksheet(SHEET_URL, "領取日誌")

                row_idx = int(info2.index[0]) + 2
                status_col_index = list(latest.columns).index(status_col_latest) + 1
                ws_main.update_cell(row_idx, status_col_index, PICKED_TEXT)

                ws_log.append_row([str(student_id), ts_str, png_id, png_link, pdf_id, pdf_link])

                # 4) 扣庫存（若有庫存表）
                inv_df3 = try_load_inventory()
                if inv_df3 is not None:
                    ok, msg = try_update_inventory(delta=-1)
                    if not ok:
                        st.warning(f"⚠️ 已完成登記，但庫存扣帳失敗：{msg}")

            # ✅ 清空狀態，像簽收機一樣回到輸入畫面
            st.cache_data.clear()
            st.session_state.verified_student_id = None
            st.session_state.eligible = False
            st.session_state.canvas_nonce += 1
            st.session_state.student_id_input = ""

            st.success(f"🎉 登記成功！學號 {student_id} 已完成領取。")
            st.info("✅ 簽名 PNG（已轉正＋浮水印）與簽收 PDF 已上傳，並寫入日誌。")

            # ✅ 自動回到輸入學號畫面
            time.sleep(0.6)
            st.rerun()

        except Exception as e:
            st.error(f"💔 存檔失敗：{e}")
