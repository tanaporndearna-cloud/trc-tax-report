import streamlit as st
import gspread
import pandas as pd
import csv
import re
import io
import json
from datetime import date, datetime
from google.oauth2.service_account import Credentials

# ============================================================
#  CONFIG — แก้ Sheet ID ตรงนี้ถ้าเปลี่ยน Sheet
# ============================================================
FORM_SHEET_ID     = "1O8x5pN7exw44wUEYOd_VcdIJ4xkc-PE8kHJmJMPSdsg"  # Form Responses
TAX_REPORT_SHEET_ID = "1bApVK4OjhlkaAIRVeslPIryyjNymeEz1VVVJA99dKvk"  # รายงานภาษีขาย

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

# ============================================================
#  ERP Column mapping
# ============================================================
def col_idx(col):
    col = col.upper()
    r = 0
    for c in col:
        r = r * 26 + (ord(c) - 64)
    return r - 1

ERP_COLS = {
    "RefDocNo":          col_idx("J"),
    "TransDate":         col_idx("B"),
    "UserRealSurName":   204,
    "OrderName":         col_idx("JY"),
    "OrderAddress":      col_idx("JZ"),
    "OrderTaxId":        col_idx("KE"),
    "IcProductDescription": 164,
    "RevenueQuantity":   col_idx("QI"),
    "PriceEach":         col_idx("CJ"),
    "PropAvailable":     col_idx("GE"),
}

# ============================================================
#  Helpers
# ============================================================
def clean_name(n):
    return re.sub(r"^#\d+\s+", "", n).strip()

def clean_item(text, maxlen=30):
    text = text.split("\n")[0].strip()
    text = re.sub(r"\s*[\(\[][^\)\]]{0,40}[\)\]]\s*$", "", text).strip()
    m = re.search(r"\s+ทรง\b", text)
    if m:
        text = text[: m.start()].strip()
    if len(text) > maxlen:
        t = text[:maxlen]
        ls = t.rfind(" ")
        if ls > maxlen // 2:
            t = t[:ls]
        text = t.rstrip()
    return text

def parse_erp_date(s):
    try:
        dt = datetime.strptime(s.strip(), "%m/%d/%Y %I:%M:%S %p")
        return date(dt.year, dt.month, dt.day)
    except:
        return None

def be_date_str(d):
    return f"{d.day:02d}/{d.month:02d}/{d.year + 543}"

def parse_erp_csv(file_bytes):
    try:
        content = file_bytes.decode("tis-620")
    except Exception:
        content = file_bytes.decode("utf-8", errors="replace")
    reader = csv.reader(io.StringIO(content))
    rows = list(reader)
    return rows[1:]  # skip header row

# ============================================================
#  Google Sheets helpers
# ============================================================
@st.cache_resource(show_spinner=False)
def connect_gspread(creds_str):
    creds_dict = json.loads(creds_str)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

def read_form_responses(gc):
    sh = gc.open_by_key(FORM_SHEET_ID)
    ws = sh.get_worksheet(0)
    rows = ws.get_all_values()
    result = {}
    for row in rows[1:]:
        if len(row) > 1 and row[1].strip():
            doc_no = row[1].strip()
            result[doc_no] = {
                "channel":    row[2].strip() if len(row) > 2 else "",
                "email_addr": row[3].strip() if len(row) > 3 else "",
            }
    return result

def read_tax_sheet(gc):
    """อ่านรายงานภาษีขาย → คืน (worksheet, inv_map)
    inv_map: {inv_no: {row_idx, has_data, date_str}}
    """
    sh = gc.open_by_key(TAX_REPORT_SHEET_ID)
    ws = sh.get_worksheet(0)
    rows = ws.get_all_values()
    inv_map = {}
    for i, row in enumerate(rows):
        if i < 2:
            continue  # skip title + header
        if len(row) > 1 and row[1].strip():
            inv_no = row[1].strip()
            has_data = bool(len(row) > 3 and row[3].strip())  # col D = เลขที่ ABB
            inv_map[inv_no] = {
                "row_idx":  i + 1,   # gspread 1-based
                "has_data": has_data,
                "date_str": row[0].strip(),
            }
    return ws, inv_map

# ============================================================
#  Core processing
# ============================================================
def process(erp_bytes, form_data, ws, inv_map):
    erp_rows = parse_erp_csv(erp_bytes)

    # Build ERP map: doc_no → list of rows
    erp_map = {}
    for row in erp_rows:
        if len(row) <= max(ERP_COLS.values()):
            continue
        doc = row[ERP_COLS["RefDocNo"]].strip()
        if doc in form_data:
            erp_map.setdefault(doc, []).append(row)

    if not erp_map:
        return [], []

    # date_str → sorted list of available (empty) invoice numbers
    date_slots = {}
    for inv_no, info in inv_map.items():
        if not info["has_data"]:
            date_slots.setdefault(info["date_str"], []).append(inv_no)
    for k in date_slots:
        date_slots[k].sort()

    # Check which docs are already written (col D has ABB number)
    written_docs = set()
    for info in inv_map.values():
        if info["has_data"]:
            # We can't easily get col D value here without re-reading,
            # but inv_map's has_data flag is enough to skip those slots.
            pass

    # Sort docs by (transaction_date, doc_no) for consistent assignment
    doc_info = {}
    for doc_no, rows in erp_map.items():
        txn_date = parse_erp_date(rows[0][ERP_COLS["TransDate"]])
        if txn_date:
            doc_info[doc_no] = {"date": txn_date, "date_str": be_date_str(txn_date)}

    sorted_docs = sorted(doc_info, key=lambda d: (doc_info[d]["date"], d))

    date_cursor = {}  # date_str → next available slot index
    preview_rows = []
    write_ops = []    # list of (row_idx, [C..O values])

    for doc_no in sorted_docs:
        d_str = doc_info[doc_no]["date_str"]
        rows = erp_map[doc_no]
        fd = form_data[doc_no]
        first = rows[0]

        # Extract product lines
        lines = []
        for r in rows:
            prop = r[ERP_COLS["PropAvailable"]].strip()
            desc = r[ERP_COLS["IcProductDescription"]].strip()
            try:
                price = float(r[ERP_COLS["PriceEach"]])
            except:
                price = 0.0
            try:
                qty = float(r[ERP_COLS["RevenueQuantity"]])
            except:
                qty = 0.0
            if prop and price > 0:
                lines.append({
                    "name":  clean_item(desc if desc else prop),
                    "qty":   int(qty),
                    "price": price,
                })
        if not lines:
            continue

        avail = date_slots.get(d_str, [])
        cursor = date_cursor.get(d_str, 0)

        if cursor + len(lines) > len(avail):
            st.warning(f"⚠️ {doc_no}: slot ไม่พอสำหรับวันที่ {d_str} (ต้องการ {len(lines)} slot, เหลือ {len(avail)-cursor})")
            continue

        for i, line in enumerate(lines):
            inv_no = avail[cursor + i]
            row_idx = inv_map[inv_no]["row_idx"]
            sale  = round(line["qty"] * line["price"], 2)
            vat   = round(sale * 0.07, 2)
            total = round(sale + vat, 2)

            values = [
                first[ERP_COLS["UserRealSurName"]].strip(),          # C: สาขา
                doc_no,                                                # D: เลขที่ ABB
                clean_name(first[ERP_COLS["OrderName"]].strip()),     # E: ชื่อลูกค้า
                first[ERP_COLS["OrderAddress"]].strip(),              # F: ที่อยู่
                first[ERP_COLS["OrderTaxId"]].strip(),               # G: เลขผู้เสียภาษี
                line["name"],                                          # H: รายการ
                line["qty"],                                           # I: จำนวน
                line["price"],                                         # J: หน่วยละ
                sale,                                                  # K: ยอดขาย
                vat,                                                   # L: VAT 7%
                total,                                                 # M: รวม
                fd.get("channel", ""),                                 # N: ช่องทาง
                fd.get("email_addr", ""),                              # O: อีเมล/ที่อยู่จัดส่ง
            ]
            write_ops.append((row_idx, inv_no, values))
            preview_rows.append({
                "เลขที่":    inv_no,
                "วันที่":    d_str,
                "เลขที่ ABB": doc_no,
                "ลูกค้า":   clean_name(first[ERP_COLS["OrderName"]].strip())[:25],
                "รายการ":   line["name"],
                "จำนวน":    line["qty"],
                "หน่วยละ":  line["price"],
                "ยอดขาย":   sale,
                "VAT":       vat,
                "รวม":       total,
            })

        date_cursor[d_str] = cursor + len(lines)

    return preview_rows, write_ops

# ============================================================
#  Streamlit UI
# ============================================================
st.set_page_config(
    page_title="TRC Tax Report",
    ="wide",
)

st.title("TRC Motorsport — Tax Report Update")
st.caption("Match ERP CSV + Form Responses and write to Google Sheet")

# ── Sidebar: Credentials ──────────────────────────────────────
with st.sidebar:
    st.header("🔑 Google Credentials")
    creds_file = st.file_uploader("อัปโหลด Service Account JSON", type=["json"])

    if creds_file:
        creds_str = creds_file.read().decode("utf-8")
        st.success("✅ โหลด credentials แล้ว")
    else:
        creds_str = None
        st.info("กรุณาอัปโหลด Service Account JSON")
        with st.expander("📖 วิธีสร้าง credentials (คลิกดู)"):
            st.markdown("""
**1.** ไปที่ [Google Cloud Console](https://console.cloud.google.com)

**2.** เปิด APIs:
- Google Sheets API
- Google Drive API

**3.** สร้าง Service Account:
- IAM & Admin → Service Accounts → Create
- ดาวน์โหลด JSON key

**4.** Share Google Sheet ให้กับ email ของ Service Account
(email รูปแบบ `xxx@project.iam.gserviceaccount.com`)
- ให้สิทธิ์ **Editor** สำหรับ Sheet รายงานภาษีขาย
- ให้สิทธิ์ **Viewer** สำหรับ Sheet Form Responses
            """)

    st.divider()
    st.caption(f"Form Sheet ID:\n`{FORM_SHEET_ID}`")
    st.caption(f"Tax Report Sheet ID:\n`{TAX_REPORT_SHEET_ID}`")

if not creds_str:
    st.warning("⬅️ กรุณาอัปโหลด Service Account JSON ในแถบซ้ายก่อนค่ะ")
    st.stop()

# ── Step 1: Connect & Load Form Responses ────────────────────
st.subheader("1️⃣  โหลด Form Responses จาก Google Sheet")

col1, col2 = st.columns([1, 4])
with col1:
    load_btn = st.button("🔄 โหลดข้อมูล", use_container_width=True)

if load_btn or "form_data" in st.session_state:
    if load_btn:
        with st.spinner("กำลังเชื่อมต่อ Google Sheets..."):
            try:
                gc = connect_gspread(creds_str)
                form_data = read_form_responses(gc)
                st.session_state["form_data"] = form_data
                st.session_state["gc"] = gc
            except Exception as e:
                st.error(f"❌ เชื่อมต่อไม่ได้: {e}")
                st.stop()

    form_data = st.session_state.get("form_data", {})
    if form_data:
        st.success(f"✅ พบ {len(form_data)} เอกสารใน Form Responses")
        with st.expander("ดูรายการเอกสาร"):
            st.dataframe(
                pd.DataFrame([
                    {"เลขที่เอกสาร": k, "ช่องทาง": v["channel"], "อีเมล/ที่อยู่": v["email_addr"]}
                    for k, v in form_data.items()
                ]),
                use_container_width=True,
                hide_index=True,
            )
    else:
        st.info("ไม่พบข้อมูลใน Form Responses ค่ะ")

# ── Step 2: Upload ERP CSV ────────────────────────────────────
st.subheader("2️⃣  อัปโหลดไฟล์ ERP CSV")
erp_file = st.file_uploader("เลือกไฟล์ ERPTax.csv (TIS-620)", type=["csv"])

# ── Step 3: Process & Preview ────────────────────────────────
if erp_file and "form_data" in st.session_state and "gc" in st.session_state:
    st.subheader("3️⃣  ตรวจสอบข้อมูลก่อนบันทึก")

    with st.spinner("กำลังประมวลผล..."):
        try:
            gc = st.session_state["gc"]
            form_data = st.session_state["form_data"]
            ws, inv_map = read_tax_sheet(gc)
            erp_bytes = erp_file.read()
            preview_rows, write_ops = process(erp_bytes, form_data, ws, inv_map)
            st.session_state["write_ops"] = write_ops
            st.session_state["ws"] = ws
        except Exception as e:
            st.error(f"❌ Error: {e}")
            st.stop()

    if preview_rows:
        st.success(f"✅ พบข้อมูลที่ match ได้ **{len(preview_rows)} แถว**")
        st.dataframe(
            pd.DataFrame(preview_rows),
            use_container_width=True,
            hide_index=True,
        )

        # ── Step 4: Write ─────────────────────────────────────
        st.subheader("4️⃣  บันทึกลง Google Sheet")
        st.info("ระบบจะเขียนเฉพาะแถวที่ว่างอยู่ ไม่แตะข้อมูลที่มีอยู่แล้วค่ะ")

        if st.button("✍️ บันทึกข้อมูลลง Google Sheet", type="primary", use_container_width=False):
            progress_bar = st.progress(0, text="กำลังบันทึก...")
            errors = []
            for i, (row_idx, inv_no, values) in enumerate(st.session_state["write_ops"]):
                try:
                    st.session_state["ws"].update(
                        f"C{row_idx}:O{row_idx}", [values]
                    )
                except Exception as e:
                    errors.append(f"แถว {inv_no}: {e}")
                progress_bar.progress(
                    (i + 1) / len(st.session_state["write_ops"]),
                    text=f"บันทึก {inv_no}... ({i+1}/{len(st.session_state['write_ops'])})",
                )

            if errors:
                st.error("เกิดข้อผิดพลาดบางส่วน:\n" + "\n".join(errors))
            else:
                st.success(f"🎉 บันทึกเรียบร้อย {len(write_ops)} แถวค่ะ!")
                st.balloons()
    else:
        st.warning("ไม่พบข้อมูลที่ match หรือทุก slot มีข้อมูลแล้วค่ะ")
