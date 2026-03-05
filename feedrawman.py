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
from PIL import Image

# PDF
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ----------------------------
# 1) 基本設定
# ----------------------------
st.set_page_config(page_title="衛生組生理用品發放系統", layout="centered")
st.title("🌸 校園生理用品領取系統")

SHEET_URL = "https://docs.google.com/spreadsheets/d/13bMCf_cgdfByYH_DgUZynKHKZAHd2qmXJKyeCfCQOg8/edit"
DRIVE_FOLDER_ID = "1WPfK0coynKAwb15EfVp0st5VEGjNmUAl"  # ✅ 你新的資料夾

# ----------------------------
# 2) Google API
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
# 3) 工具：簽名、上傳、PDF
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
    # 你 repo 裡的字體檔路徑
    pdfmetrics.registerFont(TTFont("NotoSansTC-Black", "fonts/NotoSansTC-Black.ttf"))

def make_receipt_pdf(student_id: str, ts_str: str, signature_png_bytes: bytes) -> bytes:
    """
    產出中文 PDF 簽收單（使用 NotoSansTC Black）
    """
    register_pdf_fonts()

    buf = io.BytesIO()
    c = pdf_canvas.Canvas(buf, pagesize=A4)
    width, height = A4

    # 標題
    c.setFont("NotoSansTC-Black", 18)
    c.drawString(40, height - 60, "生理用品領取簽收單")

    # 基本資訊
    c.setFont("NotoSansTC-Black", 12)
    c.drawString(40, height - 98, f"學號：{student_id}")
    c.drawString(40, height - 120, f"時間：{ts_str}")

    # 簽名區
    c.setFont("NotoSansTC-Black", 12)
    c.drawString(40, height - 160, "簽名：")

    sig_img = ImageReader(io.BytesIO(signature_png_bytes))

    # 簽名框位置/大小（你可調整）
    img_w = 420
    img_h = 160
    x = 40
    y = height - 160 - img_h - 10

    # 畫簽名框
    c.rect(x, y, img_w, img_h, stroke=1, fill=0)

    # 放簽名圖（PNG 透明背景也 OK）
    c.drawImage(sig_img, x, y, width=img_w, height=img_h, mask="auto")

    # 底部註記
    c.setFont("NotoSansTC-Black", 10)
    c.drawString(40, 40, "本簽收單由系統自動產生")

    c.showPage()
    c.save()
    return buf.getvalue()

# ----------------------------
# 4) 手機相機掃一維條碼（QuaggaJS）
#    掃到後：把結果寫進 URL ?sid=xxxx，觸發整頁 rerun
# ----------------------------
def barcode_scanner_quagga(height: int = 420):
    html = """
    <div class="w-full h-full">
      <script src="https://cdn.tailwindcss.com"></script>
      <script src="https://cdnjs.cloudflare.com/ajax/libs/quagga/0.12.1/quagga.min.js"></script>

      <div class="p-3 space-y-3">
        <div class="flex items-center justify-between">
          <div class="text-sm text-slate-600">對準條碼，亮一點、距離 10–18cm。掃到會自動停止並回填學號。</div>
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
          <button id="torchBtn"
            class="px-3 py-2 rounded-xl bg-slate-200 text-slate-900 text-sm active:scale-[0.99]">
            手電筒
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
      const torchBtn = document.getElementById("torchBtn");

      let lastCode = "";
      let scanning = false;
      let track = null;
      let torchOn = false;

      function setStatus(t) { statusEl.textContent = t; }

      function sendToStreamlit(value) {
        // ✅ 將掃描結果寫入父頁 URL query param：?sid=xxxx
        const url = new URL(window.parent.location.href);
        url.searchParams.set("sid", value);
        window.parent.location.href = url.toString();
      }

      async function tryEnableTorch() {
        try {
          if (!track) return;
          const caps = track.getCapabilities ? track.getCapabilities() : null;
          if (!caps || !caps.torch) {
            alert("此裝置/瀏覽器不支援手電筒");
            return;
          }
          torchOn = !torchOn;
          await track.applyConstraints({ advanced: [{ torch: torchOn }] });
        } catch (e) {
          console.log(e);
          alert("手電筒啟用失敗（可能不支援）");
        }
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

          setTimeout(() => {
            const video = document.querySelector("#scanner video");
            if (video && video.srcObject) {
              const tracks = video.srcObject.getVideoTracks();
              if (tracks && tracks.length) track = tracks[0];
            }
          }, 800);
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
      torchBtn.addEventListener("click", tryEnableTorch);
    </script>
    """
    return components.html(html, height=height)

# ----------------------------
# 5) 讀取名單
# ----------------------------
try:
    data = fetch_data(SHEET_URL)
except Exception as e:
    st.error(f"❌ 資料讀取失敗：{e}")
    st.stop()

# ----------------------------
# 6) UI 美化 + 狀態
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

st.write("---")
st.subheader("🔍 領取登記")

# ----------------------------
# 7) 從 URL 讀 sid 自動回填（掃描成功會帶 ?sid=xxxx）
# ----------------------------
try:
    sid = st.query_params.get("sid", None)
except Exception:
    sid = None

if sid:
    # 有些版本會是 list
    if isinstance(sid, list) and sid:
        sid = sid[0]
    sid = str(sid).strip()
    sid_digits = "".join(ch for ch in sid if ch.isdigit())
    if sid_digits:
        st.session_state.student_id_input = sid_digits
    # 清掉 query params，避免重整又重覆套用
    try:
        st.query_params.clear()
    except Exception:
        pass

# ----------------------------
# 8) 掃描 UI
# ----------------------------
st.markdown('<div class="hint">學生證是「一維條碼」。開啟掃描後，允許相機權限，對準條碼即可（掃到會自動回填學號）。</div>', unsafe_allow_html=True)
enable_scan = st.toggle("📷 用手機相機掃學生證條碼", value=False)

if enable_scan:
    barcode_scanner_quagga()

student_id = st.text_input(
    "👉 請輸入或掃描學生證學號：",
    value=st.session_state.student_id_input
).strip()
st.session_state.student_id_input = student_id

# ----------------------------
# 9) 驗證學號
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
# 10) 簽名板 + 存檔（PNG + PDF）
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
                ts = datetime.datetime.now()
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

                # 日誌：學號 / 時間 / png_id / png_link / pdf_id / pdf_link
                sheet_log.append_row([str(student_id), ts_str, png_id, png_link, pdf_id, pdf_link])

            st.cache_data.clear()
            st.session_state.verified_student_id = None
            st.session_state.eligible = False
            st.session_state.canvas_nonce += 1
            st.session_state.student_id_input = ""

            st.success(f"🎉 登記成功！學號 {student_id} 已完成領取。")
            st.info("✅ 已上傳簽名 PNG 與簽收 PDF 到 Google Drive，並寫入日誌連結。")

        except Exception as e:
            st.error(f"💔 存檔失敗：{e}")
