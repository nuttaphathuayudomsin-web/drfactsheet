import streamlit as st
import io
import copy
from pptx import Presentation
from pptx.util import Pt
import qrcode
from PIL import Image
from datetime import datetime

# ── PAGE CONFIG ───────────────────────────────────────────────
st.set_page_config(
    page_title="DR Factsheet Generator",
    page_icon="📊",
    layout="wide",
)

# ── PASSWORD GATE ─────────────────────────────────────────────
APP_PASSWORD = "Password1234"  # ← change this

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("""
    <div style="max-width:380px;margin:80px auto;">
        <div style="background:#141720;border:1px solid #252A3A;border-radius:14px;padding:36px 32px;text-align:center;">
            <div style="font-size:36px;margin-bottom:8px;">🔒</div>
            <div style="font-size:18px;font-weight:600;color:#E8ECF4;margin-bottom:4px;">DR Factsheet Generator</div>
            <div style="font-size:12px;color:#5A637A;margin-bottom:24px;">กรุณาใส่รหัสผ่านเพื่อเข้าใช้งาน</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    with st.form("login_form"):
        # center the form with columns
        _, mid, _ = st.columns([1, 2, 1])
        with mid:
            password = st.text_input("รหัสผ่าน", type="password", placeholder="Enter password")
            login = st.form_submit_button("เข้าสู่ระบบ", use_container_width=True, type="primary")

    if login:
        if password == APP_PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            _, mid, _ = st.columns([1, 2, 1])
            with mid:
                st.error("❌ รหัสผ่านไม่ถูกต้อง")

    st.stop()  # block everything below until authenticated

# ── FIXED VALUES (never change) ───────────────────────────────
FIXED = {
    "exchange"     : "ตลาดหลักทรัพย์แห่งประเทศไทย (SET)",
    "depositary"   : "ธนาคารกรุงไทย จำกัด (มหาชน)",
    "offering_type": "Direct Listing",
    "price_info"   : "เป็นไปตามกลไกราคาของตลาดในเวลาที่เสนอขาย โดยอ้างอิงจากราคาหลักทรัพย์อ้างอิงต่างประเทศ อัตราแลกเปลี่ยน อัตราส่วนอ้างอิง และค่าธรรมเนียมอื่น ๆ",
    "ktb_contact"  : "02-208-3748, 02-208-4669",
}

# ── STYLES ────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@300;400;500;600&display=swap');

    html, body, [class*="css"] { font-family: 'IBM Plex Sans Thai', sans-serif; }

    .main { background: #0D0F14; }

    .block-container { padding: 2rem 2.5rem; max-width: 1200px; }

    /* Header */
    .app-header {
        background: linear-gradient(135deg, #141720 0%, #1C2030 100%);
        border: 1px solid #252A3A;
        border-radius: 12px;
        padding: 20px 28px;
        margin-bottom: 24px;
        display: flex;
        align-items: center;
        gap: 16px;
    }

    .header-badge {
        background: #3B6FFF;
        color: white;
        font-weight: 700;
        font-size: 18px;
        width: 44px; height: 44px;
        border-radius: 10px;
        display: flex; align-items: center; justify-content: center;
    }

    /* Section labels */
    .section-label {
        font-size: 11px;
        font-weight: 600;
        letter-spacing: 0.1em;
        text-transform: uppercase;
        color: #5A637A;
        margin: 20px 0 10px;
    }

    /* Fixed values display */
    .fixed-grid {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 10px;
        margin-bottom: 20px;
    }

    .fixed-card {
        background: #141720;
        border: 1px solid #252A3A;
        border-radius: 8px;
        padding: 10px 14px;
    }

    .fixed-label { font-size: 10px; color: #5A637A; text-transform: uppercase; letter-spacing: 0.06em; }
    .fixed-value { font-size: 12px; color: #8892A4; margin-top: 3px; }

    /* Generated fields */
    .gen-card {
        background: #0D1A12;
        border: 1px solid #1A3A25;
        border-radius: 8px;
        padding: 12px 14px;
        margin-bottom: 8px;
    }

    .gen-label { font-size: 10px; color: #00D4AA; text-transform: uppercase; letter-spacing: 0.06em; }
    .gen-value { font-size: 12px; color: #B0C4B8; margin-top: 4px; line-height: 1.5; }

    /* Output box */
    .output-box {
        background: #141720;
        border: 1px solid #252A3A;
        border-radius: 10px;
        padding: 20px;
    }

    /* Streamlit button overrides */
    .stButton > button {
        font-family: 'IBM Plex Sans Thai', sans-serif !important;
        border-radius: 8px !important;
        font-weight: 500 !important;
    }

    div[data-testid="stForm"] {
        background: #141720;
        border: 1px solid #252A3A;
        border-radius: 12px;
        padding: 20px;
    }

    .stTextInput input, .stNumberInput input, .stSelectbox select {
        background: #1C2030 !important;
        border: 1px solid #252A3A !important;
        color: #E8ECF4 !important;
        border-radius: 6px !important;
    }

    hr { border-color: #252A3A !important; }
</style>
""", unsafe_allow_html=True)

# ── HEADER ────────────────────────────────────────────────────
header_col, logout_col = st.columns([6, 1])
with header_col:
    st.markdown("""
    <div class="app-header">
        <div class="header-badge">DR</div>
        <div>
            <div style="font-size:18px;font-weight:600;color:#E8ECF4;">DR Factsheet Generator</div>
            <div style="font-size:12px;color:#5A637A;margin-top:2px;">กรอกข้อมูลเพื่อสร้าง Factsheet อัตโนมัติ</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
with logout_col:
    st.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)
    if st.button("🔓 ออกจากระบบ", use_container_width=True):
        st.session_state.authenticated = False
        st.session_state.form_data = None
        st.rerun()

# ── SESSION STATE ─────────────────────────────────────────────
if "generated_files" not in st.session_state:
    st.session_state.generated_files = []

if "form_data" not in st.session_state:
    st.session_state.form_data = None

# ── HELPER: build concatenated fields ─────────────────────────
def build_data(ticker, company_name, stock_code, foreign_exchange_full,
               total_units, ratio, first_trading_date, filing_url, foreign_exchange_short):
    fmt_units = f"{int(total_units):,} หน่วย"
    fmt_ratio = f"1 : {int(ratio):,}"
    full_name_thai = f"ตราสารแสดงสิทธิในหลักทรัพย์ต่างประเทศของบริษัท {company_name} ออกโดยธนาคารกรุงไทย จำกัด (มหาชน)"
    full_name_eng  = f"Depositary receipt on {company_name} issued by Krungthai Bank Public Company Limited"
    underlying     = f"{company_name} ({stock_code})"

    return {
        "ticker"             : ticker,
        "full_name_thai"     : full_name_thai,
        "full_name_eng"      : full_name_eng,
        "exchange"           : FIXED["exchange"],
        "underlying_stock"   : underlying,
        "underlying_exchange": foreign_exchange_full,
        "depositary"         : FIXED["depositary"],
        "offering_type"      : FIXED["offering_type"],
        "total_units"        : fmt_units,
        "ratio"              : fmt_ratio,
        "first_trading_date" : first_trading_date,
        "price_info"         : FIXED["price_info"],
        "ktb_contact"        : FIXED["ktb_contact"],
        "filing_url"         : filing_url,
        "foreign_exchange"   : foreign_exchange_short,
    }

# ── HELPER: replace text in runs (handles split runs) ─────────
def replace_in_paragraph(paragraph, placeholder, value):
    """Reconstruct full paragraph text, replace, then rewrite into first run."""
    full_text = "".join(run.text for run in paragraph.runs)
    if placeholder not in full_text:
        return
    new_text = full_text.replace(placeholder, value)
    # Put all text in first run, clear the rest
    if paragraph.runs:
        paragraph.runs[0].text = new_text
        for run in paragraph.runs[1:]:
            run.text = ""

def replace_in_textframe(tf, placeholder_map):
    for para in tf.paragraphs:
        for ph, val in placeholder_map.items():
            replace_in_paragraph(para, ph, str(val) if val else "")

def replace_all_placeholders(presentation, data):
    placeholder_map = {
        "{{ticker}}"             : data["ticker"],
        "{{full_name_thai}}"     : data["full_name_thai"],
        "{{full_name_eng}}"      : data["full_name_eng"],
        "{{exchange}}"           : data["exchange"],
        "{{underlying_stock}}"   : data["underlying_stock"],
        "{{underlying_exchange}}": data["underlying_exchange"],
        "{{depositary}}"         : data["depositary"],
        "{{offering_type}}"      : data["offering_type"],
        "{{total_units}}"        : data["total_units"],
        "{{ratio}}"              : data["ratio"],
        "{{first_trading_date}}" : data["first_trading_date"],
        "{{price_info}}"         : data["price_info"],
        "{{ktb_contact}}"        : data["ktb_contact"],
        "{{filing_url}}"         : data["filing_url"],
        "{{foreign_exchange}}"   : data["foreign_exchange"],
    }

    for slide in presentation.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                replace_in_textframe(shape.text_frame, placeholder_map)
            if shape.shape_type == 19:  # Table
                for row in shape.table.rows:
                    for cell in row.cells:
                        replace_in_textframe(cell.text_frame, placeholder_map)

# ── HELPER: insert QR code ─────────────────────────────────────
def insert_qr_code(presentation, url):
    if not url:
        return
    # Generate QR
    qr = qrcode.QRCode(box_size=10, border=2)
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img_bytes = io.BytesIO()
    img.save(img_bytes, format="PNG")
    img_bytes.seek(0)

    for slide in presentation.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                txt = shape.text_frame.text.strip()
                if txt == "{{qr_code}}":
                    left, top, width, height = shape.left, shape.top, shape.width, shape.height
                    sp = shape._element
                    sp.getparent().remove(sp)
                    slide.shapes.add_picture(img_bytes, left, top, width, height)
                    img_bytes.seek(0)  # reset for next slide if needed
                    break

# ── HELPER: generate PPTX bytes ───────────────────────────────
def generate_pptx(template_bytes, data):
    prs = Presentation(io.BytesIO(template_bytes))
    replace_all_placeholders(prs, data)
    if data.get("filing_url"):
        insert_qr_code(prs, data["filing_url"])
    output = io.BytesIO()
    prs.save(output)
    return output.getvalue()

# ══════════════════════════════════════════════════════════════
#  LAYOUT — two columns
# ══════════════════════════════════════════════════════════════
left_col, right_col = st.columns([3, 2], gap="large")

# ── LEFT: FORM ────────────────────────────────────────────────
with left_col:

    # Template upload (persistent in session)
    if "template_bytes" not in st.session_state:
        st.session_state.template_bytes = None

    with st.expander("📎 อัปโหลด Template (.pptx)", expanded=st.session_state.template_bytes is None):
        uploaded = st.file_uploader("เลือกไฟล์ template", type=["pptx"], label_visibility="collapsed")
        if uploaded:
            st.session_state.template_bytes = uploaded.read()
            st.success(f"✅ โหลด template แล้ว: {uploaded.name}")

    st.markdown("---")

    # ── FIXED VALUES DISPLAY ──────────────────────────────────
    st.markdown('<div class="section-label">🔒 ข้อมูลคงที่ (ไม่ต้องกรอก)</div>', unsafe_allow_html=True)
    st.markdown(f"""
    <div class="fixed-grid">
        <div class="fixed-card">
            <div class="fixed-label">ตลาดรอง</div>
            <div class="fixed-value">{FIXED['exchange']}</div>
        </div>
        <div class="fixed-card">
            <div class="fixed-label">ผู้ออกตราสาร</div>
            <div class="fixed-value">{FIXED['depositary']}</div>
        </div>
        <div class="fixed-card">
            <div class="fixed-label">รูปแบบการเสนอขาย</div>
            <div class="fixed-value">{FIXED['offering_type']}</div>
        </div>
        <div class="fixed-card" style="grid-column:span 2;">
            <div class="fixed-label">ราคาตราสาร</div>
            <div class="fixed-value">{FIXED['price_info']}</div>
        </div>
        <div class="fixed-card">
            <div class="fixed-label">เบอร์ติดต่อ KTB</div>
            <div class="fixed-value">{FIXED['ktb_contact']}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    # ── INPUT FORM ────────────────────────────────────────────
    st.markdown('<div class="section-label">✏️ กรอกข้อมูล DR</div>', unsafe_allow_html=True)

    with st.form("dr_form", clear_on_submit=False):

        col1, col2 = st.columns(2)
        with col1:
            ticker = st.text_input("ชื่อย่อ (Ticker) *", placeholder="เช่น SUNNY80")
        with col2:
            stock_code = st.text_input("รหัสหลักทรัพย์อ้างอิง *", placeholder="เช่น 2383 HK")

        company_name = st.text_input("ชื่อบริษัทอ้างอิง (ภาษาอังกฤษ) *",
                                      placeholder="เช่น Sunny Optical Technology (Group) Co., Ltd.")

        col3, col4 = st.columns(2)
        with col3:
            foreign_exchange_full = st.text_input("ตลาดจดทะเบียน (ชื่อเต็ม) *",
                                                   placeholder="เช่น Hong Kong Stock Exchange")
        with col4:
            foreign_exchange_short = st.text_input("ชื่อย่อตลาด *", placeholder="เช่น HKEX, NDX, TSE")

        col5, col6 = st.columns(2)
        with col5:
            total_units = st.number_input("จำนวนหน่วยที่อนุมัติ *",
                                           min_value=1,
                                           value=10_000_000_000,
                                           step=1_000_000_000,
                                           format="%d")
        with col6:
            ratio = st.number_input("อัตราอ้างอิง (1 : X) *",
                                     min_value=1,
                                     value=100,
                                     step=1)

        first_trading_date = st.text_input("ประมาณการวันซื้อขายวันแรก *",
                                            placeholder="เช่น 11 มี.ค. 69")

        filing_url = st.text_input("ลิงก์ข้อมูล Filing (สำหรับ QR Code)",
                                    placeholder="https://capital.sec.or.th/...")

        st.markdown("")
        submitted = st.form_submit_button("⚡ Generate Factsheet", use_container_width=True, type="primary")

    if submitted:
        # Validate
        errors = []
        if not ticker.strip():           errors.append("กรุณากรอก Ticker")
        if not company_name.strip():     errors.append("กรุณากรอกชื่อบริษัท")
        if not stock_code.strip():       errors.append("กรุณากรอกรหัสหลักทรัพย์")
        if not foreign_exchange_full.strip(): errors.append("กรุณากรอกตลาดจดทะเบียน")
        if not foreign_exchange_short.strip(): errors.append("กรุณากรอกชื่อย่อตลาด")
        if not first_trading_date.strip(): errors.append("กรุณากรอกวันซื้อขายวันแรก")

        if errors:
            for e in errors:
                st.error(e)
        else:
            data = build_data(
                ticker.strip(),
                company_name.strip(),
                stock_code.strip(),
                foreign_exchange_full.strip(),
                total_units,
                ratio,
                first_trading_date.strip(),
                filing_url.strip(),
                foreign_exchange_short.strip(),
            )
            st.session_state.form_data = data
            st.rerun()

# ── RIGHT: PREVIEW + DOWNLOAD ─────────────────────────────────
with right_col:
    data = st.session_state.form_data

    if data is None:
        st.markdown("""
        <div style="background:#141720;border:1px dashed #252A3A;border-radius:12px;
                    padding:60px 20px;text-align:center;color:#5A637A;">
            <div style="font-size:40px;margin-bottom:12px;">📋</div>
            <div style="font-size:14px;font-weight:500;">กรอกข้อมูลด้านซ้ายแล้วกด Generate</div>
            <div style="font-size:12px;margin-top:6px;">ระบบจะสร้าง Factsheet ให้อัตโนมัติ</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown('<div class="section-label">✅ ข้อมูลที่จะใส่ใน Factsheet</div>', unsafe_allow_html=True)

        # Show all generated fields
        fields = [
            ("ชื่อย่อ",              data["ticker"]),
            ("ชื่อเต็ม (ไทย)",       data["full_name_thai"]),
            ("ชื่อเต็ม (อังกฤษ)",    data["full_name_eng"]),
            ("หลักทรัพย์อ้างอิง",    data["underlying_stock"]),
            ("ตลาดจดทะเบียน",        data["underlying_exchange"]),
            ("จำนวนหน่วย",           data["total_units"]),
            ("อัตราอ้างอิง",         data["ratio"]),
            ("วันซื้อขายวันแรก",     data["first_trading_date"]),
            ("ชื่อย่อตลาดต่างประเทศ",data["foreign_exchange"]),
            ("ลิงก์ Filing",         data["filing_url"] or "—"),
        ]

        for label, value in fields:
            short_val = value if len(value) < 80 else value[:77] + "..."
            st.markdown(f"""
            <div class="gen-card">
                <div class="gen-label">{label}</div>
                <div class="gen-value">{short_val}</div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("---")

        # Download section
        st.markdown('<div class="section-label">⬇️ ดาวน์โหลด</div>', unsafe_allow_html=True)

        if st.session_state.template_bytes is None:
            st.warning("⚠️ กรุณาอัปโหลด template .pptx ก่อนดาวน์โหลด")
        else:
            with st.spinner("กำลังสร้างไฟล์..."):
                try:
                    pptx_bytes = generate_pptx(st.session_state.template_bytes, data)
                    filename = f"{data['ticker']}.pptx"
                    st.download_button(
                        label=f"⬇️  ดาวน์โหลด {filename}",
                        data=pptx_bytes,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True,
                        type="primary",
                    )
                    st.success(f"✅ พร้อมดาวน์โหลด: {filename}")
                except Exception as e:
                    st.error(f"❌ เกิดข้อผิดพลาด: {str(e)}")

        # Clear button
        if st.button("🗑️ ล้างข้อมูล / สร้างใหม่", use_container_width=True):
            st.session_state.form_data = None
            st.rerun()

# ── FOOTER ────────────────────────────────────────────────────
st.markdown("---")
st.markdown("""
<div style="text-align:center;font-size:11px;color:#5A637A;padding:8px 0;">
    DR Factsheet Generator · KTB Securities · ข้อมูลทั้งหมดเป็นความลับ
</div>
""", unsafe_allow_html=True)
