import streamlit as st
import gspread
import pandas as pd
import csv
import re
import io
import json
import time
import calendar as _cal
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
    header = rows[0] if rows else []
    return header, rows[1:]

def parse_month_from_doc(doc_no):
    """ABBTCB26070001 -> (7, 2569) or None"""
    try:
        if len(doc_no) < 10:
            return None
        ce_yy = int(doc_no[6:8])
        month = int(doc_no[8:10])
        be_year = 2000 + ce_yy + 543
        if 1 <= month <= 12:
            return month, be_year
        return None
    except Exception:
        return None

def sheet_name_for_month(month, be_year):
    return f"{month:02d}/{be_year}"

def sheets_call(fn, max_retries=6):
    """Retry on quota 429 with exponential backoff"""
    for attempt in range(max_retries):
        try:
            return fn()
        except gspread.exceptions.APIError as e:
            if "429" in str(e) and attempt < max_retries - 1:
                wait = 2 ** attempt  # 1,2,4,8,16,32 sec
                time.sleep(wait)
            else:
                raise

def populate_new_worksheet(ws, month, be_year):
    """Fill 20 slots per day for every day in the month"""
    ce_year = be_year - 543
    days_in_month = _cal.monthrange(ce_year, month)[1]
    rows = []
    seq = 1
    for day in range(1, days_in_month + 1):
        date_str = f"{day:02d}/{month:02d}/{be_year}"
        for _ in range(20):
            inv_no = f"{be_year}{month:02d}{seq:03d}"
            rows.append([date_str, inv_no])
            seq += 1
    # rows 1-2 = header (blank), rows 3+ = slots
    all_rows = [["", ""], ["", ""]] + rows
    sheets_call(lambda: ws.update(f"A1:B{len(all_rows)}", all_rows))

def get_or_create_worksheet(sh, month, be_year):
    name = sheet_name_for_month(month, be_year)
    try:
        ws = sh.worksheet(name)
        return ws, False
    except gspread.exceptions.WorksheetNotFound:
        ce_year = be_year - 543
        days_in_month = _cal.monthrange(ce_year, month)[1]
        ws = sh.add_worksheet(title=name, rows=days_in_month * 20 + 10, cols=20)
        populate_new_worksheet(ws, month, be_year)
        return ws, True

def read_tax_sheet_ws(ws):
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
    return inv_map, written_docs

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

def process(erp_bytes, form_data, sh):
    header, erp_rows = parse_erp_csv(erp_bytes)
    try:
        branch_col = header.index("ShipToBranchNumber")
    except ValueError:
        branch_col = ERP_COLS["UserRealSurName"]

    erp_map = {}
    for row in erp_rows:
        if len(row) <= max(ERP_COLS.values()):
            continue
        doc = row[ERP_COLS["RefDocNo"]].strip()
        if doc in form_data:
            erp_map.setdefault(doc, []).append(row)
    if not erp_map:
        return [], [], []

    doc_by_month = {}
    for doc_no in erp_map:
        mth = parse_month_from_doc(doc_no)
        if mth:
            doc_by_month.setdefault(mth, []).append(doc_no)

    preview_rows = []
    write_ops = []
    new_sheets = []

    for (month, be_year), month_docs in sorted(doc_by_month.items()):
        ws, is_new = get_or_create_worksheet(sh, month, be_year)
        if is_new:
            new_sheets.append(sheet_name_for_month(month, be_year))

        inv_map, written_docs = read_tax_sheet_ws(ws)

        date_slots = {}
        for inv_no, info in inv_map.items():
            if not info["has_data"]:
                date_slots.setdefault(info["date_str"], []).append(inv_no)
        for k in date_slots:
            date_slots[k].sort()

        doc_info = {}
        for doc_no in month_docs:
            if doc_no in written_docs:
                continue
            rows_for_doc = erp_map[doc_no]
            txn_date = parse_erp_date(rows_for_doc[0][ERP_COLS["TransDate"]])
            if txn_date:
                doc_info[doc_no] = {"date": txn_date, "date_str": be_date_str(txn_date)}

        sorted_docs = sorted(doc_info, key=lambda d: (doc_info[d]["date"], d))
        date_cursor = {}
        sname = sheet_name_for_month(month, be_year)

        for doc_no in sorted_docs:
            d_str = doc_info[doc_no]["date_str"]
            rows_for_doc = erp_map[doc_no]
            fd = form_data[doc_no]
            first = rows_for_doc[0]

            lines = []
            for r in rows_for_doc:
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
                st.warning(f"no slot: {doc_no} {d_str} (sheet {sname})")
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
                    first[branch_col].strip() if branch_col < len(first) else "",
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
                    "ws":         ws,
                    "sheet_name": sname,
                    "row_idx":    row_idx,
                    "inv_no":     inv_no,
                    "values":     values,
                    "is_insert":  i > 0,
                    "date_str":   d_str,
                })
                preview_rows.append({
                    "Sheet":      sname,
                    "No":         inv_no,
                    "Date":       d_str,
                    "ABB No":     doc_no,
                    "Customer":   clean_name(first[ERP_COLS["OrderName"]].strip())[:25],
                    "Item":       line["name"],
                    "Qty":        line["qty"],
                    "Price":      line["price"],
                    "Sale":       sale,
                    "VAT":        vat,
                    "Total":      total,
                    "Note":       "insert" if i > 0 else "",
                })
            date_cursor[d_str] = cursor + 1

    return preview_rows, write_ops, new_sheets

# ---- UI ----
st.set_page_config(page_title="TRC Tax Report", layout="wide")
st.title("TRC Motorsport — Tax Report Update")
st.caption("Match ERP CSV + Form Responses and write to Google Sheet")

with st.sidebar:
    st.header("\U0001f511 Google Credentials")
    secret_creds = get_creds_from_secrets()
    if secret_creds:
        creds_str = secret_creds
        st.success("credentials loaded from Secrets")
    else:
        creds_file = st.file_uploader("Upload Service Account JSON", type=["json"])
        if creds_file:
            creds_str = creds_file.read().decode("utf-8")
            st.success("credentials loaded")
        else:
            creds_str = None
            st.info("please upload Service Account JSON")
    st.divider()
    st.caption(f"Form Sheet ID:\n`{FORM_SHEET_ID}`")
    st.caption(f"Tax Report Sheet ID:\n`{TAX_REPORT_SHEET_ID}`")

if not creds_str:
    st.warning("please upload Service Account JSON in the sidebar")
    st.stop()

st.subheader("1.  Load Form Responses from Google Sheet")
col1, col2 = st.columns([1, 4])
with col1:
    load_btn = st.button("\U0001f504 Load data", use_container_width=True)

if load_btn or "form_data" in st.session_state:
    if load_btn:
        with st.spinner("Connecting to Google Sheets..."):
            try:
                gc = connect_gspread(creds_str)
                form_data = read_form_responses(gc)
                st.session_state["form_data"] = form_data
                st.session_state["gc"] = gc
            except Exception as e:
                st.error(f"Connection error: {e}")
                st.stop()
    form_data = st.session_state.get("form_data", {})
    if form_data:
        st.success(f"Found {len(form_data)} documents in Form Responses")
        with st.expander("View documents"):
            st.dataframe(
                pd.DataFrame([{"Doc No": k, "Channel": v["channel"], "Email/Addr": v["email_addr"]} for k, v in form_data.items()]),
                use_container_width=True, hide_index=True,
            )
    else:
        st.info("No data in Form Responses")

st.subheader("2.  Upload ERP CSV")
erp_file = st.file_uploader("Select ERPTax.csv (TIS-620/CP874)", type=["csv"])

if erp_file and "form_data" in st.session_state and "gc" in st.session_state:
    st.subheader("3.  Preview before saving")
    with st.spinner("Processing..."):
        try:
            gc = st.session_state["gc"]
            form_data = st.session_state["form_data"]
            sh = gc.open_by_key(TAX_REPORT_SHEET_ID)
            erp_bytes = erp_file.read()
            preview_rows, write_ops, new_sheets = process(erp_bytes, form_data, sh)
            st.session_state["write_ops"] = write_ops
        except Exception as e:
            st.error(f"Error: {e}")
            st.stop()

    if new_sheets:
        st.success(f"Created new sheets with 20 slots/day: {', '.join(new_sheets)}")

    if preview_rows:
        insert_count = sum(1 for op in write_ops if op["is_insert"])
        note = f" ({insert_count} inserted rows)" if insert_count else ""
        st.success(f"Found **{len(preview_rows)} rows** to write{note}")
        st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)

        st.subheader("4.  Save to Google Sheet")
        st.info("Multi-line docs will INSERT new rows — invoice numbers will not be deleted")

        if st.button("Save to Google Sheet", type="primary", use_container_width=False):
            progress_bar = st.progress(0, text="Saving...")
            errors = []
            total = len(st.session_state["write_ops"])
            row_offsets = {}

            for i, op in enumerate(st.session_state["write_ops"]):
                ws = op["ws"]
                sname = op["sheet_name"]
                offset = row_offsets.get(sname, 0)
                actual_row = op["row_idx"] + offset
                inv_no = op["inv_no"]
                try:
                    if op["is_insert"]:
                        row_data = [op["date_str"]] + list(op["values"])
                        row_data[11] = f"=K{actual_row}*0.07"
                        row_data[12] = f"=K{actual_row}+L{actual_row}"
                        sheets_call(lambda rd=row_data, ar=actual_row: ws.insert_rows([rd], ar, inherit_from_before=True, value_input_option="USER_ENTERED"))
                        row_offsets[sname] = offset + 1
                    else:
                        row_values = list(op["values"])
                        row_values[10] = f"=K{actual_row}*0.07"
                        row_values[11] = f"=K{actual_row}+L{actual_row}"
                        sheets_call(lambda rv=row_values, ar=actual_row: ws.update(f"B{ar}:O{ar}", [rv], value_input_option="USER_ENTERED"))
                    time.sleep(0.5)
                except Exception as e:
                    errors.append(f"{inv_no} ({sname}): {e}")
                action = "insert" if op["is_insert"] else "save"
                progress_bar.progress((i + 1) / total, text=f"{action} {inv_no} ({sname})... ({i+1}/{total})")

            if errors:
                st.error("Some errors:\n" + "\n".join(errors))
            else:
                st.success(f"Done! {total} rows saved")
                st.balloons()
    else:
        st.info("No new data to write (all docs already processed or no matching slots)")
