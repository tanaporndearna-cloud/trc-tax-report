import streamlit as st
import gspread
import pandas as pd
import csv
import re
import io
import json
from datetime import date, datetime
from google.oauth2.service_account import Credentials

FORM_SHEET_ID       = "1O8x5pN7exw44wUEYOd_VcdIJ4xkc-PE8kHJmJMPSdsg"
TAX_REPORT_SHEET_ID = "1bApVK4OjhlkaAIRVeslPIryyjNymeEz1VVVJA99dKvk"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

def col_idx(col):
    col = col.upper()
    r = 0
    for c in col:
        r = r * 26 + (ord(c) - 64)
    return r - 1

ERP_COLS = {
    "RefDocNo":             col_idx("J"),
    "TransDate":            col_idx("B"),
    "UserRealSurName":      204,
    "OrderName":            col_idx("JY"),
    "OrderAddress":         col_idx("JZ"),
    "OrderTaxId":           col_idx("KE"),
    "IcProductDescription": 164,
    "RevenueQuantity":      col_idx("QI"),
    "PriceEach":            col_idx("CJ"),
    "PropAvailable":        col_idx("GE"),
}

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
    except Exception:
        return None

def be_date_str(d):
    return f"{d.day:02d}/{d.month:02d}/{d.year + 543}"

def parse_erp_csv(file_bytes):
    for enc in ("cp874", "tis-620", "utf-8-sig", "utf-8"):
        try:
            content = file_bytes.decode(enc)
            if "?" not in content[:500] and "�" not in content[:500]:
                break
        except Exception:
            continue
    else:
        content = file_bytes.decode("cp874", errors="replace")
    reader = csv.reader(io.StringIO(content))
    rows = list(reader)
    return rows[1:]

@st.cache_resource(show_spinner=False)
def connect_gspread(creds_str):
    creds_dict = json.loads(creds_str) if isinstance(creds_str, str) else creds_str
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

def get_creds_from_secrets():
    try:
        s = st.secrets["gcp_service_account"]
        return dict(s)
    except Exception:
        return None

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
    sh = gc.open_by_key(TAX_REPORT_SHEET_ID)
    ws = sh.get_worksheet(0)
    rows = ws.get_all_values()
    inv_map = {}
    written_docs = set()
    for i, row in enumerate(rows):
        if i < 2:
            continue
        if len(row) > 1 and row[1].strip():
            inv_no = row[1].strip()
            has_data = bool(len(row) > 3 and row[3].strip())
            if has_data and len(row) > 3:
                written_docs.add(row[3].strip())
            if inv_no not in inv_map:
                inv_map[inv_no] = {
                    "row_idx":  i + 1,
                    "has_data": has_data,
                    "date_str": row[0].strip(),
                }
    return ws, inv_map, written_docs

def process(erp_bytes, form_data, ws, inv_map, written_docs):
    erp_rows = parse_erp_csv(erp_bytes)
    erp_map = {}
    for row in erp_rows:
        if len(row) <= max(ERP_COLS.values()):
            continue
        doc = row[ERP_COLS["RefDocNo"]].strip()
        if doc in form_data:
            erp_map.setdefault(doc, []).append(row)
    if not erp_map:
        return [], []

    date_slots = {}
    for inv_no, info in inv_map.items():
        if not info["has_data"]:
            date_slots.setdefault(info["date_str"], []).append(inv_no)
    for k in date_slots:
        date_slots[k].sort()

    doc_info = {}
    for doc_no, rows in erp_map.items():
        if doc_no in written_docs:
            continue
        txn_date = parse_erp_date(rows[0][ERP_COLS["TransDate"]])
        if txn_date:
            doc_info[doc_no] = {"date": txn_date, "date_str": be_date_str(txn_date)}

    sorted_docs = sorted(doc_info, key=lambda d: (doc_info[d]["date"], d))
    date_cursor = {}
    preview_rows = []
    write_ops = []

    for doc_no in sorted_docs:
        d_str = doc_info[doc_no]["date_str"]
        rows = erp_map[doc_no]
        fd = form_data[doc_no]
        first = rows[0]

        lines = []
        for r in rows:
            prop = r[ERP_COLS["PropAvailable"]].strip()
            desc = r[ERP_COLS["IcProductDescription"]].strip()
            try:
                price = float(r[ERP_COLS["PriceEach"]])
            except Exception:
                price = 0.0
            try:
                qty = float(r[ERP_COLS["RevenueQuantity"]])
            except Exception:
                qty = 0.0
            if prop and price > 0:
                lines.append({"name": clean_item(desc if desc else prop), "qty": int(qty), "price": price})
        if not lines:
            continue

        avail = date_slots.get(d_str, [])
        cursor = date_cursor.get(d_str, 0)
        if cursor + 1 > len(avail):
            st.warning(f"⚠️ {doc_no}: ไม่มีสล็อตว่างสำหรับวันที่ {d_str}")
            continue

        inv_no = avail[cursor]
        base_row_idx = inv_map[inv_no]["row_idx"]

        for i, line in enumerate(lines):
            row_idx = base_row_idx + i
            sale  = round(line["qty"] * line["price"], 2)
            vat   = round(sale * 0.07, 2)
            total = round(sale + vat, 2)
            values = [
                inv_no,
                first[ERP_COLS["UserRealSurName"]].strip(),
                doc_no,
                clean_name(first[ERP_COLS["OrderName"]].strip()),
                first[ERP_COLS["OrderAddress"]].strip(),
                first[ERP_COLS["OrderTaxId"]].strip(),
                line["name"],
                line["qty"],
                line["price"],
                sale,
                vat,
                total,
                fd.get("channel", ""),
                fd.get("email_addr", ""),
            ]
            write_ops.append({
                "row_idx":   row_idx,
                "inv_no":    inv_no,
                "values":    values,
                "is_insert": i > 0,
                "date_str":  d_str,
            })
            note = "แทรกแถว" if i > 0 else ""
            preview_rows.append({
                "เลขที่":     inv_no,
                "วันที่":     d_str,
                "เลขที่ ABB": doc_no,
                "ลูกค้า":     clean_name(first[ERP_COLS["OrderName"]].strip())[:25],
                "รายการ":     line["name"],
                "จำนวน":      line["qty"],
                "หน่วยละ":    line["price"],
                "ยอดขาย":     sale,
                "VAT":        vat,
                "รวม":        total,
                "หมายเหตุ":   note,
            })
        date_cursor[d_str] = cursor + 1

    return preview_rows, write_ops

# ---- UI ----
st.set_page_config(page_title="TRC Tax Report", layout="wide")
st.title("TRC Motorsport — Tax Report Update")
st.caption("Match ERP CSV + Form Responses and write to Google Sheet")

with st.sidebar:
    st.header("\U0001f511 Google Credentials")
    secret_creds = get_creds_from_secrets()
    if secret_creds:
        creds_str = secret_creds
        st.success("✅ โหลด credentials จาก Secrets แล้ว")
    else:
        creds_file = st.file_uploader("อัปโหลด Service Account JSON", type=["json"])
        if creds_file:
            creds_str = creds_file.read().decode("utf-8")
            st.success("✅ โหลด credentials แล้ว")
        else:
            creds_str = None
            st.info("กรุณาอัปโหลด Service Account JSON")
    st.divider()
    st.caption(f"Form Sheet ID:\n`{FORM_SHEET_ID}`")
    st.caption(f"Tax Report Sheet ID:\n`{TAX_REPORT_SHEET_ID}`")

if not creds_str:
    st.warning("❪️ กรุณาอัปโหลด Service Account JSON ในแถบซ้ายก่อนค่ะ")
    st.stop()

st.subheader("1️⃣  โหลด Form Responses จาก Google Sheet")
col1, col2 = st.columns([1, 4])
with col1:
    load_btn = st.button("\U0001f504 โหลดข้อมูล", use_container_width=True)

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
                pd.DataFrame([{"เลขที่เอกสาร": k, "ช่องทาง": v["channel"], "อีเมล/ที่อยู่": v["email_addr"]} for k, v in form_data.items()]),
                use_container_width=True, hide_index=True,
            )
    else:
        st.info("ไม่พบข้อมูลใน Form Responses ค่ะ")

st.subheader("2️⃣  อัปโหลดไฟล์ ERP CSV")
erp_file = st.file_uploader("เลือกไฟล์ ERPTax.csv (TIS-620/CP874)", type=["csv"])

if erp_file and "form_data" in st.session_state and "gc" in st.session_state:
    st.subheader("3️⃣  ตรวจสอบข้อมูลก่อนบันทึก")
    with st.spinner("กำลังประมวลผล..."):
        try:
            gc = st.session_state["gc"]
            form_data = st.session_state["form_data"]
            ws, inv_map, written_docs = read_tax_sheet(gc)
            erp_bytes = erp_file.read()
            preview_rows, write_ops = process(erp_bytes, form_data, ws, inv_map, written_docs)
            st.session_state["write_ops"] = write_ops
            st.session_state["ws"] = ws
        except Exception as e:
            st.error(f"❌ Error: {e}")
            st.stop()

    if preview_rows:
        insert_count = sum(1 for op in write_ops if op["is_insert"])
        note = f" (แทรกแถวใหม่ {insert_count} แถว)" if insert_count else ""
        st.success(f"✅ พบข้อมูลที่ match ได้ **{len(preview_rows)} แถว**{note}")
        st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)

        st.subheader("4️⃣  บันทึกลง Google Sheet")
        st.info("แถวที่มีหลายรายการจะ **แทรกแถวใหม่** เลขที่ใบกำกับภาษีเดิมจะไม่ถูกลบค่ะ")

        if st.button("✍️ บันทึกข้อมูลลง Google Sheet", type="primary", use_container_width=False):
            progress_bar = st.progress(0, text="กำลังบันทึก...")
            errors = []
            total = len(st.session_state["write_ops"])
            row_offset = 0

            for i, op in enumerate(st.session_state["write_ops"]):
                actual_row = op["row_idx"] + row_offset
                inv_no = op["inv_no"]
                try:
                    if op["is_insert"]:
                        row_data = [op["date_str"]] + op["values"]
                        st.session_state["ws"].insert_rows([row_data], actual_row, inherit_from_before=True)
                        row_offset += 1
                    else:
                        st.session_state["ws"].update(f"B{actual_row}:O{actual_row}", [op["values"]])
                except Exception as e:
                    errors.append(f"แถว {inv_no}: {e}")
                action = "แทรก" if op["is_insert"] else "บันทึก"
                progress_bar.progress((i + 1) / total, text=f"{action} {inv_no}... ({i+1}/{total})")

            if errors:
                st.error("เกิดข้อผิดพลาดบางส่วน:\n" + "\n".join(errors))
            else:
                st.success(f"\U0001f389 บันทึกเรียบร้อย {total} แถวค่ะ!")
                st.balloons()
    else:
        if written_docs:
            st.info(f"ℹ️ ทุก doc ใน ERP ถูกบันทึกลง Sheet แล้ว ({len(written_docs)} รายการ) ไม่มีข้อมูลใหม่ค่ะ")
        else:
            st.warning("ไม่พบข้อมูลที่ match หรือทุก slot มีข้อมูลแล้วค่ะ")
