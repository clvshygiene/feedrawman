import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from streamlit_drawable_canvas import st_canvas
import datetime
import io
import re
from PIL import Image

# 台灣時間
from zoneinfo import ZoneInfo

# 寄信
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# PDF
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.utils import ImageReader


# ----------------------------
# 0) 常數
# ----------------------------
TZ_TAIPEI = ZoneInfo("Asia/Taipei")

st.set_page_config(page_title="衛生組生理用品發放系統", layout="centered")
st.title("🌸 校園生理用品領取系統")

SHEET_URL = "https://docs.google.com/spreadsheets/d/13bMCf_cgdfByYH_DgUZynKHKZAHd2qmXJKyeCfCQOg8/edit"
DRIVE_FOLDER_ID = "1WPfK0coynKAwb15EfVp0st5VEGjNmUAl"

# 你校內學號長度（請依你學校調整）
STUDENT_ID_MIN_LEN = 6
STUDENT_ID_MAX_LEN = 10


# ----------------------------
# 1) Google API
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

@st.cache_data(ttl=300)
def fetch_data(url):
    gc = get_gsheet_client()
    doc = gc.open_by_url(url)
    sheet = doc.worksheet("工作表1")
    return pd.DataFrame(sheet.get_all_records())


# ----------------------------
# 2) 工具：簽名、上傳、PDF
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
    # 你已經成功的版本：若你已放字體檔，這行就會生效
    # 沒放字體檔也沒關係：下面 make_receipt_pdf 會 fallback
    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        pdfmetrics.registerFont(TTFont("NotoSansTC-Black", "fonts/NotoSansTC-Black.ttf"))
        return True
    except Exception:
        return False

def make_receipt_pdf(student_id: str, ts_str: str, signature_png_bytes: bytes) -> bytes:
    """
    中文簽收單（優先使用 NotoSansTC-Black；失敗則 fallback Helvetica）
    """
    ok_font = register_pdf_fonts()
    font_name = "NotoSansTC-Black" if ok_font else "Helvetica-Bold"

    buf = io.BytesIO()
    c = pdf_canvas.Canvas(buf, pagesize=A4)
    width, height = A4

    # 標題
    c.setFont(font_name, 18)
    c.drawString(40, height - 60, "生理用品領取簽收單")

    # 基本資訊
    c.setFont(font_name, 12)
    c.drawString(40, height - 98, f"學號：{student_id}")
    c.drawString(40, height - 120, f"時間：{ts_str}")

    # 簽名區
    c.setFont(font_name, 12)
    c.drawString(40, height - 160, "簽名：")

    sig_img = ImageReader(io.BytesIO(signature_png_bytes))
    img_w = 420
    img_h = 160
    x = 40
    y = height - 160 - img_h - 10

    c.rect(x, y, img_w, img_h, stroke=1, fill=0)
    c.drawImage(sig_img, x, y, width=img_w, height=img_h, mask="auto")

    c.setFont(font_name, 10)
    c.drawString(40, 40, "本簽收單由系統自動產生")

    c.showPage()
    c.save()
    return buf.getvalue()


# ----------------------------
# 3) 掃碼：QuaggaJS（掃到後寫入 URL ?sid=xxxx）
# ----------------------------
def barcode_scanner_quagga(height: int = 420):
    html = """
    <div class="w-full h-full">
      <script src="https://cdn.tailwindcss.com"></script>
      <script src="https://cdnjs.cloudflare.com/ajax/libs/quagga/0.12.1/quagga.min.js"></script>

      <div class="p-3 space-y-3">
        <div class="flex items-center justify-between">
          <div class="text-sm text-slate-600">對準條碼（亮一點、距離 10–18cm）。掃到會自動回填。</div>
          <div class="text-xs text-slate-500" id="status">未啟動</div>
        </div>

        <div class="flex gap-2">
          <button id="startBtn"
            class="px-3 py-2 rounded-xl bg-slate-900 text-white text-sm active:scale-[0.99]">
            開始掃描
          </button>
          <button id="stopBtn"
            class="px-3 py-2 rounded-xl bg-slate-200 text-slate-900 text-sm active:scale-[0.99]">
            停止
          </button>
        </div>

        <div class="rounded-2xl overflow-hidden border border-slate-200 bg-black">
          <div id="scanner" style="width: 100%; height: 280px;"></div>
        </div>

        <div class="rounded-2xl border border-slate-200 bg-white p-3">
          <div class="text-xs text-slate-500">掃描結果</div>
          <div class="text-lg font-semibold" id="result">—</div>
        </div>
      </div>
    </div>

    <script>
      const statusEl = document.getElementById("status");
      const resultEl = document.getElementById("result");
      const startBtn = document.getElementById("startBtn");
      const stopBtn = document.getElementById("stopBtn");

      let lastCode = "";
      let scanning = false;

      function setStatus(t) { statusEl.textContent = t; }

      function sendToStreamlit(value) {
        // 將掃描結果寫入父頁 URL query param：?sid=xxxx
        const url = new URL(window.parent.location.href);
        url.searchParams.set("sid", value);
        // 加一個 cache-bust，避免某些瀏覽器不重載
        url.searchParams.set("_t", Date.now().toString());
        window.parent.location.href = url.toString();
      }

      function start() {
        if (scanning) return;
        setStatus("啟動中…");

        Quagga.init({
          inputStream: {
            name: "Live",
            type: "LiveStream",
            target: document.querySelector('#scanner'),
            constraints: {
              facingMode: "environment",
              width: { ideal: 1280 },
              height: { ideal: 720 },
            },
          },
          decoder: {
            readers: [
              "code_128_reader",
              "code_39_reader",
              "ean_reader",
              "ean_8_reader",
              "upc_reader"
            ]
          },
          locate: true,
          numOfWorkers: 0
        }, function(err) {
          if (err) {
            console.log(err);
            setStatus("啟動失敗");
            alert("相機啟動失敗：請確認瀏覽器相機權限已允許");
            return;
          }
          Quagga.start();
          scanning = true;
          setStatus("掃描中");
        });

        Quagga.onDetected(function(data) {
          const code = (data && data.codeResult && data.codeResult.code) ? data.codeResult.code : "";
          if (!code) return;
          if (code === lastCode) return;
          lastCode = code;

          resultEl.textContent = code;
          stop();
          sendToStreamlit(code);
        });
      }

      function stop() {
        if (!scanning) return;
        try { Quagga.stop(); } catch (e) {}
        scanning = false;
        setStatus("已停止");
      }

      startBtn.addEventListener("click", start);
      stopBtn.addEventListener("click", stop);
    </script>
    """
    return components.html(html, height=height)


# ----------------------------
# 4) 把掃描字串轉成「學號」
# ----------------------------
def parse_student_id(raw: str) -> tuple[str | None, str]:
    """
    回傳 (student_id_or_none, debug_message)
    規則：抓最像學號的一段連續數字（長度 6~10 可調）
    """
    if raw is None:
        return None, "raw=None"

    s = str(raw).strip()
    # 有些情況 query_params 可能給 list
    if isinstance(raw, list) and raw:
        s = str(raw[0]).strip()

    # 抓出所有連續數字片段
    nums = re.findall(r"\d+", s)
    if not nums:
        return None, f"raw='{s}' 找不到任何數字片段"

    # 優先選「長度符合學號範圍」的片段，且取最長/最像的
    candidates = [x for x in nums if STUDENT_ID_MIN_LEN <= len(x) <= STUDENT_ID_MAX_LEN]
    if candidates:
        # 取最長的（通常學號固定長度）
        candidates.sort(key=len, reverse=True)
        return candidates[0], f"raw='{s}' nums={nums} -> pick={candidates[0]}"

    # 若沒有符合長度的，就退而求其次：取最長一段
    nums.sort(key=len, reverse=True)
    return nums[0], f"raw='{s}' nums={nums} -> fallback_pick={nums[0]}"


# ----------------------------
# 5) SMTP 寄信
# ----------------------------
def send_email(to_addr: str, subject: str, body: str) -> None:
    """
    從 st.secrets 讀 SMTP 設定：
    st.secrets["smtp"]["host"], ["port"], ["username"], ["password"], ["from_name"], ["from_addr"], ["use_tls"]
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

    with smtplib.SMTP(host, port, timeout=20) as server:
        server.ehlo()
        if use_tls:
            server.starttls()
            server.ehlo()
        if username and password:
            server.login(username, password)
        server.send_message(msg)


# ----------------------------
# 6) 讀取名單
# ----------------------------
try:
    data = fetch_data(SHEET_URL)
except Exception as e:
    st.error(f"❌ 資料讀取失敗：{e}")
    st.stop()


# ----------------------------
# 7) UI 美化 + session_state
# ----------------------------
st.markdown("""
<style>
.toolbar {border:1px solid rgba(0,0,0,.08); border-radius:16px; padding:12px; background:white;}
.canvas-wrap {border-radius:16px; overflow:hidden; border:1px solid rgba(0,0,0,.10);}
.small-label {font-size:12px; color:rgba(0,0,0,.6); margin-bottom:4px;}
.hint {font-size:13px; color:rgba(0,0,0,.65); margin-bottom:8px;}
</style>
""", unsafe_allow_html=True)

if "verified_student_id" not in st.session_state:
    st.session_state.verified_student_id = None
if "eligible" not in st.session_state:
    st.session_state.eligible = False
if "canvas_nonce" not in st.session_state:
    st.session_state.canvas_nonce = 0
if "student_id_input" not in st.session_state:
    st.session_state.student_id_input = ""
if "scan_debug" not in st.session_state:
    st.session_state.scan_debug = ""


# ----------------------------
# 8) 管理功能：寄提醒信（在 sidebar）
# ----------------------------
with st.sidebar:
    st.header("✉️ 提醒信寄送")

    # 你可以自行加管理密碼（避免學生亂寄）
    admin_ok = True
    if "admin" in st.secrets:
        pwd = st.text_input("管理密碼", type="password")
        admin_ok = (pwd == st.secrets["admin"].get("password", ""))

    if not admin_ok:
        st.warning("需要管理密碼才能寄信")
    else:
        st.caption("收件人規則：學號 + @g.clvs.tyc.edu.tw")

        # 你可以貼一個或多個學號（每行一個）
        ids_text = st.text_area(
            "要寄信的學號（可多筆，每行一個）",
            placeholder="例如：\n311117\n311118",
            height=120,
        )

        subject = st.text_input("主旨", value="生理用品領取提醒")
        body = st.text_area(
            "內容（可用 {student_id} 代入學號）",
            value="同學您好：\n\n提醒您可至衛生組領取生理用品。\n學號：{student_id}\n\n謝謝。",
            height=180,
        )

        if st.button("📨 寄出提醒信", use_container_width=True):
            ids = [x.strip() for x in ids_text.splitlines() if x.strip()]
            if not ids:
                st.error("請至少輸入 1 個學號")
            else:
                failed = []
                sent = 0
                for sid in ids:
                    email = f"{sid}@g.clvs.tyc.edu.tw"
                    try:
                        send_email(
                            to_addr=email,
                            subject=subject.format(student_id=sid),
                            body=body.format(student_id=sid),
                        )
                        sent += 1
                    except Exception as e:
                        failed.append((sid, str(e)))

                if sent:
                    st.success(f"已寄出 {sent} 封")
                if failed:
                    st.error("部分寄送失敗：")
                    for sid, err in failed[:10]:
                        st.write(f"- {sid}: {err}")


# ----------------------------
# 9) 領取登記 + 掃描
# ----------------------------
st.write("---")
st.subheader("🔍 領取登記")

# 讀 URL query params：sid（掃描結果）
try:
    raw_sid = st.query_params.get("sid", None)
except Exception:
    raw_sid = None

if raw_sid:
    student_id_parsed, dbg = parse_student_id(raw_sid)
    st.session_state.scan_debug = dbg
    if student_id_parsed:
        st.session_state.student_id_input = student_id_parsed
    # 清 query params（避免一直重覆套用）
    try:
        st.query_params.clear()
    except Exception:
        pass

st.markdown('<div class="hint">若掃描回填失敗，下面會顯示掃描 debug（只給你看）。</div>', unsafe_allow_html=True)
enable_scan = st.toggle("📷 用手機相機掃學生證條碼", value=False)
if enable_scan:
    barcode_scanner_quagga()

# 顯示 debug（你要更乾淨可關掉）
if st.session_state.scan_debug:
    with st.expander("🛠️ 掃描 debug（若回填不對請看這裡）"):
        st.code(st.session_state.scan_debug)

student_id = st.text_input(
    "👉 請輸入或掃描學生證學號：",
    value=st.session_state.student_id_input
).strip()
st.session_state.student_id_input = student_id


# ----------------------------
# 10) 驗證學號
# ----------------------------
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


# ----------------------------
# 11) 簽名板 + 存檔（PNG + PDF）
# ----------------------------
if st.session_state.eligible and st.session_state.verified_student_id == str(student_id) and student_id:
    st.markdown('<div class="toolbar">', unsafe_allow_html=True)

    c1, c2, c3, c4, c5 = st.columns([1.3, 1, 1, 1, 1])
    with c1:
        st.markdown('<div class="small-label">工具</div>', unsafe_allow_html=True)
        tool = st.selectbox("tool", ["筆刷", "橡皮擦", "直線", "矩形", "圓形"], label_visibility="collapsed")
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

    if st.button("🚀 確認領取並存檔", use_container_width=True):
        if canvas_result.image_data is None:
            st.warning("請先簽名再點擊送出。")
            st.stop()

        try:
            with st.spinner("正在存檔中..."):
                # 0) 抓最新狀態，避免快取造成重複領取
                latest = fetch_data(SHEET_URL)
                student_info2 = latest[latest["學號"].astype(str) == str(student_id)]
                if student_info2.empty:
                    st.error("❌ 資料已更新，找不到該學號，請重新輸入。")
                    st.stop()

                status2 = str(student_info2.iloc[0].get("本次領取狀態", "")).strip()
                if status2 == "已領取":
                    st.warning("⚠️ 剛剛已被登記領取（可能重複操作），請勿重複送出。")
                    st.stop()

                # 1) canvas -> PNG
                png_bytes = canvas_to_png_bytes(canvas_result.image_data)

                # 台灣時間
                ts = datetime.datetime.now(TZ_TAIPEI)
                ts_str = ts.strftime("%Y-%m-%d %H:%M:%S")
                ts_file = ts.strftime("%Y%m%d_%H%M%S")

                png_name = f"signature_{student_id}_{ts_file}.png"
                png_id, png_link = upload_bytes_to_drive(png_bytes, png_name, "image/png")

                # 2) 產 PDF + 上傳
                pdf_bytes = make_receipt_pdf(student_id, ts_str, png_bytes)
                pdf_name = f"receipt_{student_id}_{ts_file}.pdf"
                pdf_id, pdf_link = upload_bytes_to_drive(pdf_bytes, pdf_name, "application/pdf")

                # 3) 寫入 Sheet
                gc = get_gsheet_client()
                doc = gc.open_by_url(SHEET_URL)
                sheet_main = doc.worksheet("工作表1")
                sheet_log = doc.worksheet("領取日誌")

                row_idx = int(student_info2.index[0]) + 2
                sheet_main.update_cell(row_idx, 2, "已領取")

                sheet_log.append_row([str(student_id), ts_str, png_id, png_link, pdf_id, pdf_link])

            st.cache_data.clear()
            st.session_state.verified_student_id = None
            st.session_state.eligible = False
            st.session_state.canvas_nonce += 1
            st.session_state.student_id_input = ""
            st.session_state.scan_debug = ""

            st.success(f"🎉 登記成功！學號 {student_id} 已完成領取。")
            st.info("✅ 已上傳簽名 PNG 與簽收 PDF 到 Google Drive，並寫入日誌連結。")

        except Exception as e:
            st.error(f"💔 存檔失敗：{e}")
