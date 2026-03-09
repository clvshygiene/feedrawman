import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from streamlit_drawable_canvas import st_canvas

import datetime
import io
import os
import time
import uuid
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

st.set_page_config(page_title="衛生組生理用品領取系統", layout="centered")

SHEET_URL = "https://docs.google.com/spreadsheets/d/13bMCf_cgdfByYH_DgUZynKHKZAHd2qmXJKyeCfCQOg8/edit"
DRIVE_FOLDER_ID = "1WPfK0coynKAwb15EfVp0st5VEGjNmUAl"

STUDENT_EMAIL_DOMAIN = "@g.clvs.tyc.edu.tw"
PICKED_TEXT = "已領取"
WATERMARK_TEXT = "衛生組專用"


# ----------------------------
# 安全：雙密碼 + 錯誤鎖定
# ----------------------------
def check_app_password():
    if "app_unlocked" not in st.session_state:
        st.session_state.app_unlocked = False
    if "login_fail_count" not in st.session_state:
        st.session_state.login_fail_count = 0
    if "login_locked_until" not in st.session_state:
        st.session_state.login_locked_until = None

    if st.session_state.app_unlocked:
        return

    passwords = st.secrets["app"]["passwords"]
    max_attempts = int(st.secrets["app"].get("max_attempts", 5))
    lock_seconds = int(st.secrets["app"].get("lock_seconds", 30))

    st.title("🔒 衛生組生理用品領取系統")
    st.warning("請先輸入網站密碼")

    now = datetime.datetime.now(TZ_TAIPEI)
    locked_until = st.session_state.login_locked_until

    if locked_until and now < locked_until:
        remaining = int((locked_until - now).total_seconds())
        st.error(f"登入失敗次數過多，請 {remaining} 秒後再試")
        st.stop()

    pwd = st.text_input("網站密碼", type="password")

    if st.button("登入", use_container_width=True):
        if pwd in passwords:
            st.session_state.app_unlocked = True
            st.session_state.login_fail_count = 0
            st.session_state.login_locked_until = None
            st.rerun()
        else:
            st.session_state.login_fail_count += 1
            remain_attempts = max_attempts - st.session_state.login_fail_count

            if st.session_state.login_fail_count >= max_attempts:
                st.session_state.login_locked_until = now + datetime.timedelta(seconds=lock_seconds)
                st.error(f"密碼錯誤次數過多，已鎖定 {lock_seconds} 秒")
            else:
                st.error(f"密碼錯誤，剩餘 {remain_attempts} 次嘗試")

    st.stop()


check_app_password()
st.title("🌸 校園生理用品領取系統")


# ----------------------------
# Google API
# ----------------------------
@st.cache_resource
def get_gcp_credentials():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    return Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scopes
    )

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
# 字體
# ----------------------------
def get_font_for_image(img_width: int):
    size = max(18, int(img_width * 0.06))
    for fp in [
        os.path.join("fonts", "NotoSansTC-Black.ttf"),
        os.path.join("fonts", "NotoSansTC-Black.otf"),
    ]:
        if os.path.exists(fp):
            return ImageFont.truetype(fp, size)
    return ImageFont.load_default()


# ----------------------------
# 簽名圖：強制旋轉 + 浮水印
# ----------------------------
def canvas_to_png_bytes(canvas_image_data) -> bytes:
    img = Image.fromarray(canvas_image_data.astype("uint8"), mode="RGBA")

    # 學生固定橫著簽，強制轉正
    img = img.rotate(90, expand=True)

    w, h = img.size
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)

    font = get_font_for_image(w)
    bbox = draw.textbbox((0, 0), WATERMARK_TEXT, font=font)
    tw = bbox[2] - bbox[0]

    x = (w - tw) // 2
    y = max(8, int(h * 0.03))

    draw.text((x, y), WATERMARK_TEXT, font=font, fill=(0, 0, 0, 80))
    img = Image.alpha_composite(img, overlay)

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


# ----------------------------
# Drive：上傳 + 去公開
# ----------------------------
def remove_public_permissions(file_id: str):
    """
    移除 anyone 公開權限。
    注意：若 Shared Drive/父資料夾本身有更寬的繼承權限，這裡無法完全推翻繼承。
    """
    drive = get_gdrive_service()

    try:
        perms = drive.permissions().list(
            fileId=file_id,
            supportsAllDrives=True,
            fields="permissions(id,type,role)"
        ).execute()

        permissions = perms.get("permissions", [])
        for p in permissions:
            if p.get("type") == "anyone":
                try:
                    drive.permissions().delete(
                        fileId=file_id,
                        permissionId=p["id"],
                        supportsAllDrives=True
                    ).execute()
                except Exception:
                    pass
    except Exception:
        # 權限清理失敗不致命，但不宣稱一定完全成功
        pass


def upload_bytes_to_drive(file_bytes: bytes, filename: str, mimetype: str) -> tuple[str, str]:
    drive = get_gdrive_service()
    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mimetype, resumable=False)
    metadata = {"name": filename, "parents": [DRIVE_FOLDER_ID]}

    created = drive.files().create(
        body=metadata,
        media_body=media,
        fields="id",
        supportsAllDrives=True
    ).execute()

    file_id = created["id"]

    # 去掉任何 public 權限
    remove_public_permissions(file_id)

    # ✅ 不回傳 webViewLink，避免連結外流
    return file_id, ""


# ----------------------------
# PDF
# ----------------------------
@st.cache_resource
def register_pdf_fonts():
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
# SMTP
# ----------------------------
def send_email(to_addr: str, subject: str, body: str) -> None:
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
# 可選：庫存
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

        cur = int(df.iloc[0].get(stock_col, 0))
        new_val = cur + delta
        if new_val < 0:
            return False, "庫存不足，無法扣帳"

        row = 2
        col = list(df.columns).index(stock_col) + 1
        ws.update_cell(row, col, new_val)
        return True, f"庫存 {cur} → {new_val}"
    except Exception as e:
        return False, f"庫存更新失敗：{e}"


# ----------------------------
# 讀主表
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
# Session State
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
# Sidebar
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

    show_dashboard = st.toggle("📊 顯示統計儀表板", value=True)

    st.divider()

    st.header("✉️ 未領取自動寄信")
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
            ids_to_send = list(dict.fromkeys(ids_to_send))
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
        daily = (
            log_df.dropna(subset=["_date"])
            .groupby("_date")
            .size()
            .reset_index(name="count")
            .sort_values("_date")
        )

        today = datetime.datetime.now(TZ_TAIPEI).date()
        start = today - datetime.timedelta(days=6)
        daily_7 = daily[daily["_date"] >= start]

        if not daily_7.empty:
            st.line_chart(daily_7.set_index("_date")["count"])
        else:
            st.info("近 7 日尚無領取資料")

        today_count = int((log_df["_date"] == today).sum())
        st.caption(f"今日領取人次：{today_count}")
    else:
        st.info("尚未建立「領取日誌」或格式不足，暫不顯示趨勢圖。")


# ----------------------------
# 主畫面：登記
# ----------------------------
st.write("---")
st.subheader("✍️ 領取登記（手動輸入學號）")
st.info("📱 請將手機橫過來簽名，系統會自動轉正並加上浮水印")

student_id = st.text_input("👉 請輸入學生證學號：", value=st.session_state.student_id_input).strip()
st.session_state.student_id_input = student_id

if student_id:
    # 基本輸入限制（不是 IP 防爆破，但可減少亂打）
    if not student_id.isdigit() or not (4 <= len(student_id) <= 12):
        st.error("❌ 學號格式不正確")
        st.stop()

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
# 簽名板
# ----------------------------
if st.session_state.eligible and st.session_state.verified_student_id == str(student_id) and student_id:
    st.markdown("""
    <style>
    .canvas-wrap {border-radius:16px; overflow:hidden; border:1px solid rgba(0,0,0,.10);}
    </style>
    """, unsafe_allow_html=True)

    stroke_color = "#111111"
    bg_color = "#FFFFFF"
    stroke_width = 8
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

                # ✅ 簽名圖：轉正 + 浮水印
                png_bytes = canvas_to_png_bytes(canvas_result.image_data)

                ts = datetime.datetime.now(TZ_TAIPEI)
                ts_str = ts.strftime("%Y-%m-%d %H:%M:%S")
                ts_file = ts.strftime("%Y%m%d_%H%M%S")

                # ✅ UUID 檔名，不暴露學號
                png_uuid = str(uuid.uuid4())
                pdf_uuid = str(uuid.uuid4())

                png_name = f"signature_{png_uuid}_{ts_file}.png"
                png_id, _ = upload_bytes_to_drive(png_bytes, png_name, "image/png")

                pdf_bytes = make_receipt_pdf(student_id, ts_str, png_bytes)
                pdf_name = f"receipt_{pdf_uuid}_{ts_file}.pdf"
                pdf_id, _ = upload_bytes_to_drive(pdf_bytes, pdf_name, "application/pdf")

                ws_main = get_worksheet(SHEET_URL, "工作表1")
                ws_log = get_worksheet(SHEET_URL, "領取日誌")

                row_idx = int(info2.index[0]) + 2
                status_col_index = list(latest.columns).index(status_col_latest) + 1
                ws_main.update_cell(row_idx, status_col_index, PICKED_TEXT)

                # ✅ 只存 file_id，不存 webViewLink
                ws_log.append_row([str(student_id), ts_str, png_id, pdf_id])

                inv_df3 = try_load_inventory()
                if inv_df3 is not None:
                    ok, msg = try_update_inventory(delta=-1)
                    if not ok:
                        st.warning(f"⚠️ 已完成登記，但庫存扣帳失敗：{msg}")

            st.cache_data.clear()
            st.session_state.verified_student_id = None
            st.session_state.eligible = False
            st.session_state.canvas_nonce += 1
            st.session_state.student_id_input = ""

            st.success(f"🎉 登記成功！學號 {student_id} 已完成領取。")
            st.info("✅ 簽名圖已加浮水印、檔名已 UUID 化、Drive 已去除公開權限。")

            time.sleep(0.6)
            st.rerun()

        except Exception as e:
            st.error(f"💔 存檔失敗：{e}")
