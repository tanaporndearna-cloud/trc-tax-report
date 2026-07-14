import streamlit as st
import gspread
import pandas as pd
import csv, re, io, json
from datetime import date, datetime
from google.oauth2.service_account import Credentials

FORM_SHEET_ID = '1O8x5pN7exw44wUEYOd_VcdIJ4xkc-PE8kHJmJMPSdsg'
TAX_REPORT_SHEET_ID = '1bApVK4OjhlkaAIRVeslPIryyjNymeEz1VVVJA99dKvk'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive.readonly']

def col_idx(col):
    col = col.upper(); r = 0
    for c in col: r = r * 26 + (ord(c) - 64)
    return r - 1

ERP_COLS = {
    'RefDocNo': col_idx('J'), 'TransDate': col_idx('B'),
    'UserRealSurName': 204, 'OrderName': col_idx('JY'),
    'OrderAddress': col_idx('JZ'), 'OrderTaxId': col_idx('KE'),
    'IcProductDescription': 164, 'RevenueQuantity': col_idx('QI'),
    'PriceEach': col_idx('CJ'), 'PropAvailable': col_idx('GE'),
}

def clean_name(n): return re.sub(r'^#\d+\s+', '', n).strip()

def clean_item(text, maxlen=30):
    text = text.split('\n')[0].strip()
    text = re.sub(r'\s*[\(\[][^\)\]]{0,40}[\)\]]\s*$', '', text).strip()
    m = re.search(r'\s+\u0e17\u0e23\u0e07\b', text)
    if m: text = text[:m.start()].strip()
    if len(text) > maxlen:
        t = text[:maxlen]; ls = t.rfind(' ')
        if ls > maxlen // 2: t = t[:ls]
        text = t.rstrip()
    return text

def parse_erp_date(s):
    try:
        dt = datetime.strptime(s.strip(), '%m/%d/%Y %I:%M:%S %p')
        return date(dt.year, dt.month, dt.day)
    except: return None

def be_date_str(d): return f'{d.day:02d}/{d.month:02d}/{d.year+543}'

def parse_erp_csv(b):
    try: content = b.decode('tis-620')
    except: content = b.decode('utf-8', errors='replace')
    return list(csv.reader(io.StringIO(content)))[1:]

@st.cache_resource(show_spinner=False)
def connect_gspread(s):
    creds = Credentials.from_service_account_info(json.loads(s), scopes=SCOPES)
    return gspread.authorize(creds)

def read_form_responses(gc):
    rows = gc.open_by_key(FORM_SHEET_ID).get_worksheet(0).get_all_values()
    result = {}
    for row in rows[1:]:
        if len(row) > 1 and row[1].strip():
            result[row[1].strip()] = {
                'channel': row[2].strip() if len(row) > 2 else '',
                'email_addr': row[3].strip() if len(row) > 3 else '',
            }
    return result

def read_tax_sheet(gc):
    ws = gc.open_by_key(TAX_REPORT_SHEET_ID).get_worksheet(0)
    rows = ws.get_all_values(); inv_map = {}
    for i, row in enumerate(rows):
        if i < 2: continue
        if len(row) > 1 and row[1].strip():
            inv_map[row[1].strip()] = {
                'row_idx': i + 1,
                'has_data': bool(len(row) > 3 and row[3].strip()),
                'date_str': row[0].strip(),
            }
    return ws, inv_map

def process(erp_bytes, form_data, ws, inv_map):
    erp_rows = parse_erp_csv(erp_bytes); erp_map = {}
    for row in erp_rows:
        if len(row) <= max(ERP_COLS.values()): continue
        doc = row[ERP_COLS['RefDocNo']].strip()
        if doc in form_data: erp_map.setdefault(doc, []).append(row)
    if not erp_map: return [], []
    date_slots = {}
    for inv_no, info in inv_map.items():
        if not info['has_data']:
            date_slots.setdefault(info['date_str'], []).append(inv_no)
    for k in date_slots: date_slots[k].sort()
    doc_info = {}
    for doc_no, rows in erp_map.items():
        txn = parse_erp_date(rows[0][ERP_COLS['TransDate']])
        if txn: doc_info[doc_no] = {'date': txn, 'date_str': be_date_str(txn)}
    sorted_docs = sorted(doc_info, key=lambda d: (doc_info[d]['date'], d))
    dc = {}; preview = []; ops = []
    for doc_no in sorted_docs:
        d_str = doc_info[doc_no]['date_str']
        rows = erp_map[doc_no]; fd = form_data[doc_no]; first = rows[0]; lines = []
        for r in rows:
            prop = r[ERP_COLS['PropAvailable']].strip()
            desc = r[ERP_COLS['IcProductDescription']].strip()
            try: price = float(r[ERP_COLS['PriceEach']])
            except: price = 0.0
            try: qty = float(r[ERP_COLS['RevenueQuantity']])
            except: qty = 0.0
            if prop and price > 0:
                lines.append({'name': clean_item(desc if desc else prop), 'qty': int(qty), 'price': price})
        if not lines: continue
        avail = date_slots.get(d_str, []); cur = dc.get(d_str, 0)
        if cur + len(lines) > len(avail):
            st.warning(f'slot \u0e44\u0e21\u0e48\u0e1e\u0e2d: {doc_no}'); continue
        for i, line in enumerate(lines):
            inv_no = avail[cur + i]; ri = inv_map[inv_no]['row_idx']
            sale = round(line['qty'] * line['price'], 2)
            vat = round(sale * 0.07, 2); total = round(sale + vat, 2)
            v = [first[ERP_COLS['UserRealSurName']].strip(), doc_no,
                 clean_name(first[ERP_COLS['OrderName']].strip()),
                 first[ERP_COLS['OrderAddress']].strip(),
                 first[ERP_COLS['OrderTaxId']].strip(),
                 line['name'], line['qty'], line['price'], sale, vat, total,
                 fd.get('channel',''), fd.get('email_addr','')]
            ops.append((ri, inv_no, v))
            preview.append({'\u0e40\u0e25\u0e02\u0e17\u0e35\u0e48': inv_no, 'ABB': doc_no,
                '\u0e23\u0e32\u0e22\u0e01\u0e32\u0e23': line['name'],
                '\u0e08\u0e33\u0e19\u0e27\u0e19': line['qty'], '\u0e23\u0e27\u0e21': total})
        dc[d_str] = cur + len(lines)
    return preview, ops

# ── UI ──
st.set_page_config(page_title='TRC \u0e20\u0e32\u0e29\u0e35\u0e02\u0e32\u0e22', page_icon='\ud83d\udccb', layout='wide')
st.title('\ud83d\udccb TRC Motorsport \u2014 \u0e23\u0e32\u0e22\u0e07\u0e32\u0e19\u0e20\u0e32\u0e29\u0e35\u0e02\u0e32\u0e22')
with st.sidebar:
    st.header('\ud83d\udd11 Google Credentials')
    cf = st.file_uploader('\u0e2d\u0e31\u0e1b\u0e42\u0e2b\u0e25\u0e14 Service Account JSON', type=['json'])
    if cf: creds_str = cf.read().decode('utf-8'); st.success('\u2705 \u0e42\u0e2b\u0e25\u0e14\u0e41\u0e25\u0e49\u0e27')
    else: creds_str = None; st.info('\u0e01\u0e23\u0e38\u0e13\u0e32\u0e2d\u0e31\u0e1b\u0e42\u0e2b\u0e25\u0e14 JSON')
    st.divider()
    st.caption(f'Form: {FORM_SHEET_ID}'); st.caption(f'Tax: {TAX_REPORT_SHEET_ID}')
if not creds_str: st.warning('\u25c4 \u0e2d\u0e31\u0e1b\u0e42\u0e2b\u0e25\u0e14 credentials \u0e01\u0e48\u0e2d\u0e19\u0e04\u0e48\u0e30'); st.stop()
if st.button('\ud83d\udd04 \u0e42\u0e2b\u0e25\u0e14 Form Responses') or 'form_data' in st.session_state:
    if 'form_data' not in st.session_state:
        try:
            gc = connect_gspread(creds_str)
            st.session_state['form_data'] = read_form_responses(gc)
            st.session_state['gc'] = gc
        except Exception as e: st.error(f'\u274c {e}'); st.stop()
    fd = st.session_state['form_data']
    st.success(f'\u2705 \u0e1e\u0e1a {len(fd)} \u0e40\u0e2d\u0e01\u0e2a\u0e32\u0e23') if fd else st.info('\u0e44\u0e21\u0e48\u0e1e\u0e1a')
erp_file = st.file_uploader('\u2b06\ufe0f \u0e2d\u0e31\u0e1b\u0e42\u0e2b\u0e25\u0e14 ERPTax.csv', type=['csv'])
if erp_file and 'form_data' in st.session_state and 'gc' in st.session_state:
    with st.spinner('\u0e1b\u0e23\u0e30\u0e21\u0e27\u0e25\u0e1c\u0e25...'):
        try:
            ws, inv_map = read_tax_sheet(st.session_state['gc'])
            prev, ops = process(erp_file.read(), st.session_state['form_data'], ws, inv_map)
            st.session_state['ops'] = ops; st.session_state['ws'] = ws
        except Exception as e: st.error(f'\u274c {e}'); st.stop()
    if prev:
        st.success(f'\u2705 match {len(prev)} \u0e41\u0e16\u0e27')
        st.dataframe(pd.DataFrame(prev), use_container_width=True, hide_index=True)
        if st.button('\u270d\ufe0f \u0e1a\u0e31\u0e19\u0e17\u0e36\u0e01\u0e25\u0e07 Google Sheet', type='primary'):
            bar = st.progress(0)
            for i, (ri, inv_no, v) in enumerate(st.session_state['ops']):
                try: st.session_state['ws'].update(f'C{ri}:O{ri}', [v])
                except: pass
                bar.progress((i+1)/len(st.session_state['ops']))
            st.success('\ud83c\udf89 \u0e40\u0e23\u0e35\u0e22\u0e1a\u0e23\u0e49\u0e2d\u0e22\u0e04\u0e48\u0e30!'); st.balloons()
    else: st.warning('\u0e44\u0e21\u0e48\u0e1e\u0e1a\u0e02\u0e49\u0e2d\u0e21\u0e39\u0e25 \u0e2b\u0e23\u0e37\u0e2d\u0e17\u0e38\u0e01 slot \u0e21\u0e35\u0e02\u0e49\u0e2d\u0e21\u0e39\u0e25\u0e41\u0e25\u0e49\u0e27')
