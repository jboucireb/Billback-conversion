#!/usr/bin/env python3
"""Monin Billback Processor v2 — exact Tellus Sheet1 format"""

import os, re, io, json, threading, webbrowser, email, uuid, shutil, tempfile, traceback
from http.server import HTTPServer, BaseHTTPRequestHandler
from datetime import datetime
import pandas as pd
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font
import warnings
warnings.filterwarnings('ignore')

TODAY = datetime.now().strftime('%Y%m%d')

# ─── EXACT TELLUS SHEET1 COLUMNS (20 cols, no blank first col) ───────────────
S1_COLS = [
    'Source', 'Vendor #', 'Program #', 'Customer Reference',
    'Payee Type', 'Distributor ID', 'Operator ID', 'Closing Method',
    'Billback Date', 'BB Start Date', 'BB End Date',
    'Item Number', 'Item Volume Qty', 'Item Dollar Amount',
    'Associated Distributor ID', 'UOM', 'Trade Indicator',
    'Lump Sum Indicator', 'File Name', 'Component'
]

# ─── DEFAULT SUPPLIER CONFIG ─────────────────────────────────────────────────
DEFAULT_SUPPLIER_CONFIG = {
    'KAST':       {'program_num': '1004089', 'dist_id': '134810000', 'trade': 'D'},
    'SOFO':       {'program_num': '',        'dist_id': '',           'trade': 'D'},
    'PFS':        {'program_num': '',        'dist_id': '',           'trade': 'D'},
    'LABATT':     {'program_num': '',        'dist_id': '',           'trade': 'D'},
    'Y_HATA':     {'program_num': '',        'dist_id': '',           'trade': 'D'},
    'BEK':        {'program_num': '',        'dist_id': '',           'trade': 'D'},
    'NICH_CO':    {'program_num': '',        'dist_id': '',           'trade': 'D'},
    'SHAMROCK':   {'program_num': '',        'dist_id': '',           'trade': 'D'},
    'DOT_CBBB':   {'program_num': '',        'dist_id': '',           'trade': 'D'},
    'MCLANE':     {'program_num': '',        'dist_id': '',           'trade': 'D'},
    'S_AND_W':    {'program_num': '',        'dist_id': '',           'trade': 'D'},
    'BLAIR_CANDY':{'program_num': '',        'dist_id': '',           'trade': 'D'},
    'HARBOR':      {'program_num': '',        'dist_id': '',           'trade': 'D'},
    'MARTIN_BROS':  {'program_num': '',        'dist_id': '',           'trade': 'D'},
    'DOT_FOODS_BB': {'program_num': '',        'dist_id': '',           'trade': 'D'},
    'DRISCOLL':     {'program_num': '',        'dist_id': '',           'trade': 'D'},
    'UNKNOWN':      {'program_num': '',        'dist_id': '',           'trade': 'D'},
}

# ─── DATE HELPERS ─────────────────────────────────────────────────────────────
def to_yyyymmdd(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ''
    if isinstance(val, str):
        val = val.strip().replace('/', '-')
        for fmt in ('%m-%d-%Y','%Y-%m-%d','%m-%d-%y','%Y%m%d','%d-%b-%Y'):
            try:
                return datetime.strptime(val, fmt).strftime('%Y%m%d')
            except: pass
        digits = re.sub(r'\D','',val)
        if len(digits) == 8: return digits
        return val
    if isinstance(val, datetime):
        return val.strftime('%Y%m%d')
    if hasattr(val, 'strftime'):
        return val.strftime('%Y%m%d')
    return str(val)

def parse_date_range(text):
    m = re.search(r'(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\s*[-–to]+\s*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})', str(text))
    if m:
        return to_yyyymmdd(m.group(1).replace('-','/')), to_yyyymmdd(m.group(2).replace('-','/'))
    return '', ''

def month_name_to_range(period_str):
    """Convert 'December 2025' → ('20251201', '20251231')"""
    import calendar
    months = {'january':1,'february':2,'march':3,'april':4,'may':5,'june':6,
              'july':7,'august':8,'september':9,'october':10,'november':11,'december':12}
    s = str(period_str).strip().lower()
    for name, num in months.items():
        if name in s:
            m = re.search(r'(\d{4})', s)
            if m:
                year = int(m.group(1))
                last = calendar.monthrange(year, num)[1]
                start = f'{year}{num:02d}01'
                end   = f'{year}{num:02d}{last:02d}'
                return start, end
    return '', ''

def clean_amount(val):
    if val is None: return 0.0
    s = str(val).replace('$','').replace(',','').strip()
    neg = s.startswith('(') and s.endswith(')')
    s = s.strip('()')
    try:
        v = float(s)
        return -v if neg else v
    except: return 0.0

# ─── TELLUS ROW BUILDER (returns dict keyed by S1_COLS) ──────────────────────
def make_row(source='', program_num='', customer_ref='', dist_id='',
             bill_date='', start_date='', end_date='',
             item='', qty=0, amount=0.0, trade='D'):
    try:
        dist_id_int = int(str(dist_id).strip()) if str(dist_id).strip() else None
    except:
        dist_id_int = None
    try:
        qty_int = int(float(str(qty))) if qty not in ('', None) else 0
    except:
        qty_int = 0
    try:
        amt_float = round(float(str(amount)), 4) if amount not in ('', None) else 0.0
    except:
        amt_float = 0.0
    return {
        'Source':                  str(source),
        'Vendor #':                ' ',
        'Program #':               str(program_num) if program_num else ' ',
        'Customer Reference':      str(customer_ref),
        'Payee Type':              'D',
        'Distributor ID':          dist_id_int,
        'Operator ID':             ' ',
        'Closing Method':          'D',
        'Billback Date':           str(bill_date) if bill_date else TODAY,
        'BB Start Date':           str(start_date),
        'BB End Date':             str(end_date),
        'Item Number':             str(item),
        'Item Volume Qty':         qty_int,
        'Item Dollar Amount':      amt_float,
        'Associated Distributor ID': dist_id_int,
        'UOM':                     'CS',
        'Trade Indicator':         str(trade),
        'Lump Sum Indicator':      None,
        'File Name':               None,
        'Component':               ' ',
    }

# ─── FORMAT DETECTOR ─────────────────────────────────────────────────────────
def detect_supplier(filename):
    fn = os.path.basename(filename).upper()
    if 'KAST' in fn: return 'KAST'
    if 'SOFO' in fn: return 'SOFO'
    if 'PFS' in fn: return 'PFS'
    if 'BLAIR' in fn: return 'BLAIR_CANDY'
    if 'LABATT' in fn: return 'LABATT'
    if 'SHAMROCK' in fn: return 'SHAMROCK'
    if re.search(r'S[\s_]*AND[\s_]*W|S\s*&\s*W', fn): return 'S_AND_W'
    if 'BEK' in fn: return 'BEK'
    if 'NICH' in fn: return 'NICH_CO'
    if 'CBBB' in fn: return 'DOT_CBBB'
    if re.search(r'Y[\s.]?HATA|Y_HATA|TM\s+\d{6}', fn): return 'Y_HATA'
    if 'DRISCOLL' in fn: return 'DRISCOLL'
    if 'HARBOR' in fn: return 'HARBOR'
    if 'SUPPLIER BILLBACK' in fn or 'SUPPLIER_BILLBACK' in fn: return 'HARBOR'
    m = re.search(r'CUST\s+(\d+)', fn.replace('  ',' '))
    if m:
        cust = m.group(1)
        if cust in ('15101', '141699') and fn.endswith('.XLSX'): return 'MCLANE_OR_DOT'
        if fn.endswith('.PDF'): return 'PFS_STYLE'
    return 'UNKNOWN'

# ─── PARSERS ─────────────────────────────────────────────────────────────────

def parse_bek_or_nich(filepath, source_name, cfg, customer_ref):
    rows = []
    try:
        df = pd.read_csv(filepath, sep='\t', dtype=str, low_memory=False)
        df.columns = df.columns.str.strip()
        for _, r in df.iterrows():
            item = str(r.get('MID','')).strip()
            if not item or item == 'nan' or not item.startswith('M-'): continue
            inv = str(r.get('InvoiceNumber','')).strip()
            bill = to_yyyymmdd(str(r.get('InvoiceDate','')).strip())
            start, end = parse_date_range(str(r.get('ProgramPeriod','')))
            qty = r.get('Quantity','0')
            amt = clean_amount(r.get('ExtendedAmount','0'))
            rows.append(make_row(
                source=source_name,
                program_num=cfg['program_num'],
                customer_ref=customer_ref or inv,
                dist_id=cfg['dist_id'],
                bill_date=bill,
                start_date=start,
                end_date=end,
                item=item,
                qty=qty,
                amount=amt,
                trade=cfg['trade']
            ))
    except Exception as e:
        rows.append({'_error': str(e)})
    return rows

def parse_shamrock(filepath, cfg, customer_ref):
    rows = []
    try:
        wb = __import__('openpyxl').load_workbook(filepath, data_only=True)
        ws = wb.active
        header_row = None
        col_map = {}
        for i, row in enumerate(ws.iter_rows(values_only=True), 1):
            row_vals = [str(v).strip() if v else '' for v in row]
            if 'MFG PROD #' in row_vals or 'MFG PROD#' in row_vals or any('MFG' in c and 'PROD' in c for c in row_vals):
                header_row = i
                col_map = {v.upper(): j for j, v in enumerate(row_vals)}
                break
        if header_row is None:
            return [{'_error': 'SHAMROCK: header not found'}]
        def gc(key_options):
            for k in key_options:
                for ck, ci in col_map.items():
                    if k in ck: return ci
            return None
        ci_item   = gc(['MFG PROD'])
        ci_qty    = gc(['QUANTITY','QTY'])
        ci_amt    = gc(['NET ALLOW','ALLOW','AMOUNT'])
        ci_bill   = gc(['INVOICE DATE','BILL DATE'])
        ci_start  = gc(['START','FROM'])
        ci_end    = gc(['END','TO','THRU'])
        for row in ws.iter_rows(min_row=header_row+1, values_only=True):
            row_vals = [str(v).strip() if v is not None else '' for v in row]
            if not any(row_vals): continue
            item = row_vals[ci_item] if ci_item is not None else ''
            if not item or not item.upper().startswith('M-'): continue
            qty  = row_vals[ci_qty]  if ci_qty  is not None else 0
            amt  = clean_amount(row_vals[ci_amt]  if ci_amt  is not None else 0)
            bill = to_yyyymmdd(row_vals[ci_bill]  if ci_bill is not None else '')
            start= to_yyyymmdd(row_vals[ci_start] if ci_start is not None else '')
            end  = to_yyyymmdd(row_vals[ci_end]   if ci_end  is not None else '')
            rows.append(make_row(
                source='Shamrock',
                program_num=cfg['program_num'],
                customer_ref=customer_ref,
                dist_id=cfg['dist_id'],
                bill_date=bill,
                start_date=start,
                end_date=end,
                item=item,
                qty=qty,
                amount=amt,
                trade=cfg['trade']
            ))
    except Exception as e:
        rows.append({'_error': str(e)})
    return rows

def parse_dot_cbbb(filepath, cfg, customer_ref):
    rows = []
    try:
        wb = __import__('openpyxl').load_workbook(filepath, data_only=True)
        ws = wb.active
        header_row = None
        col_map = {}
        for i, row in enumerate(ws.iter_rows(values_only=True), 1):
            row_vals = [str(v).strip().upper() if v else '' for v in row]
            if 'ITEM #' in row_vals or 'ITEM#' in row_vals or any('ITEM' in c and '#' in c for c in row_vals):
                header_row = i
                col_map = {v: j for j, v in enumerate(row_vals)}
                break
        if header_row is None:
            return [{'_error': 'DOT_CBBB: header not found'}]
        def gc(keys):
            for k in keys:
                for ck, ci in col_map.items():
                    if k in ck: return ci
            return None
        ci_item  = gc(['ITEM #','ITEM#'])
        ci_qty   = gc(['QTY','QUANTITY','CASES'])
        ci_amt   = gc(['ALLOW AMT','ALLOWANCE','AMOUNT'])
        ci_bill  = gc(['INV DATE','INVOICE DATE','DATE'])
        ci_start = gc(['START'])
        ci_end   = gc(['END','THRU'])
        for row in ws.iter_rows(min_row=header_row+1, values_only=True):
            row_vals = [str(v).strip() if v is not None else '' for v in row]
            if not any(row_vals): continue
            item = row_vals[ci_item] if ci_item is not None else ''
            if not item or not item.upper().startswith('M-'): continue
            qty  = row_vals[ci_qty]  if ci_qty  is not None else 0
            amt  = clean_amount(row_vals[ci_amt]  if ci_amt  is not None else 0)
            bill = to_yyyymmdd(row_vals[ci_bill]  if ci_bill is not None else '')
            start= to_yyyymmdd(row_vals[ci_start] if ci_start is not None else '')
            end  = to_yyyymmdd(row_vals[ci_end]   if ci_end  is not None else '')
            rows.append(make_row(
                source='DOT_CBBB',
                program_num=cfg['program_num'],
                customer_ref=customer_ref,
                dist_id=cfg['dist_id'],
                bill_date=bill,
                start_date=start,
                end_date=end,
                item=item,
                qty=qty,
                amount=amt,
                trade=cfg['trade']
            ))
    except Exception as e:
        rows.append({'_error': str(e)})
    return rows

def parse_mclane(filepath, cfg, customer_ref):
    rows = []
    try:
        wb = __import__('openpyxl').load_workbook(filepath, data_only=True)
        ws = wb.active
        header_row = None
        col_map = {}
        for i, row in enumerate(ws.iter_rows(values_only=True), 1):
            row_vals = [str(v).strip().upper() if v else '' for v in row]
            if any('M-CODE' in c or 'MCODE' in c or ('ITEM' in c and 'CODE' in c) for c in row_vals):
                header_row = i
                col_map = {v: j for j, v in enumerate(row_vals)}
                break
        if header_row is None:
            return [{'_error': 'MCLANE: header not found'}]
        def gc(keys):
            for k in keys:
                for ck, ci in col_map.items():
                    if k in ck: return ci
            return None
        ci_item  = gc(['M-CODE','MCODE','ITEM CODE','ITEM #'])
        ci_qty   = gc(['QTY','QUANTITY','CASES'])
        ci_amt   = gc(['ALLOW','AMOUNT','NET AMT'])
        ci_bill  = gc(['INV DATE','INVOICE DATE','DATE'])
        ci_start = gc(['START'])
        ci_end   = gc(['END','THRU'])
        for row in ws.iter_rows(min_row=header_row+1, values_only=True):
            row_vals = [str(v).strip() if v is not None else '' for v in row]
            if not any(row_vals): continue
            item = row_vals[ci_item] if ci_item is not None else ''
            if not item or not item.upper().startswith('M-'): continue
            qty  = row_vals[ci_qty]  if ci_qty  is not None else 0
            amt  = clean_amount(row_vals[ci_amt]  if ci_amt  is not None else 0)
            bill = to_yyyymmdd(row_vals[ci_bill]  if ci_bill is not None else '')
            start= to_yyyymmdd(row_vals[ci_start] if ci_start is not None else '')
            end  = to_yyyymmdd(row_vals[ci_end]   if ci_end  is not None else '')
            rows.append(make_row(
                source='McLane',
                program_num=cfg['program_num'],
                customer_ref=customer_ref,
                dist_id=cfg['dist_id'],
                bill_date=bill,
                start_date=start,
                end_date=end,
                item=item,
                qty=qty,
                amount=amt,
                trade=cfg['trade']
            ))
    except Exception as e:
        rows.append({'_error': str(e)})
    return rows

def parse_sw(filepath, cfg, customer_ref):
    """S&W Wholesale Foods vendor invoice Excel format.
    Header row contains: Mfq. Product Code | Total case qty. | Incentive $$ ...
    Date range is extracted from the header area (Invoice Activity Date row).
    Amount per row = Total case qty × Incentive $$.
    Items with numeric Mfq. codes (no M-prefix) are flagged as warnings.
    """
    rows = []
    try:
        wb = __import__('openpyxl').load_workbook(filepath, data_only=True)
        ws = wb.active

        # Extract date range from header area (before data section)
        start_date = end_date = bill_date = TODAY
        inv_num = customer_ref
        for row in ws.iter_rows(max_row=30, values_only=True):
            for cell in row:
                if not cell or not isinstance(cell, str): continue
                m = re.search(r'(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})', cell)
                if m:
                    start_date = to_yyyymmdd(m.group(1))
                    end_date   = to_yyyymmdd(m.group(2))
                if 'Invoice Number' in cell or 'Invoice #' in cell:
                    pass  # handled by adjacent cell below
            # Check for invoice number in adjacent cells
            row_list = list(row)
            for j, cell in enumerate(row_list):
                if cell and isinstance(cell, str) and 'Invoice Number' in cell:
                    if j+1 < len(row_list) and row_list[j+1]:
                        inv_num = inv_num or str(row_list[j+1]).strip()

        cref = customer_ref or inv_num

        # Find column header row
        header_row = None
        col_map = {}
        for i, row in enumerate(ws.iter_rows(values_only=True), 1):
            row_vals = [str(v).strip().upper() if v else '' for v in row]
            if any('MFQ' in c and 'PRODUCT' in c for c in row_vals):
                header_row = i
                col_map = {v: j for j, v in enumerate(row_vals)}
                break
        if header_row is None:
            return [{'_error': 'S&W: column header not found'}]

        def gc(keys):
            for k in keys:
                for ck, ci in col_map.items():
                    if k in ck: return ci
            return None

        ci_item = gc(['MFQ. PRODUCT CODE', 'MFQ PRODUCT', 'MFG. PRODUCT'])
        ci_qty  = gc(['TOTAL CASE QTY', 'CASE QTY', 'QTY', 'CASES', 'QUANTITY'])
        ci_inc  = gc(['INCENTIVE'])          # Incentive $$ per case
        ci_desc = gc(['PRODUCT # : NAME', 'PRODUCT NAME', 'DESCRIPTION'])

        item_map = cfg.get('item_map', {})

        for row in ws.iter_rows(min_row=header_row+1, values_only=True):
            if not any(row): continue
            raw_item = str(row[ci_item]).strip() if ci_item is not None and row[ci_item] else ''
            if not raw_item or raw_item.upper() in ('NONE', ''): continue

            qty  = float(row[ci_qty] or 0) if ci_qty  is not None else 0
            inc  = clean_amount(str(row[ci_inc] or 0)) if ci_inc is not None else 0.0
            amt  = round(qty * inc, 4)

            desc = str(row[ci_desc]).strip() if ci_desc is not None and row[ci_desc] else raw_item

            if re.match(r'M-[A-Z0-9]+', raw_item, re.I):
                rows.append(make_row(
                    source='S&W',
                    program_num=cfg['program_num'],
                    customer_ref=cref,
                    dist_id=cfg['dist_id'],
                    bill_date=bill_date,
                    start_date=start_date,
                    end_date=end_date,
                    item=raw_item.upper(),
                    qty=qty,
                    amount=amt,
                    trade=cfg['trade']
                ))
            else:
                mapped = item_map.get(raw_item.strip())
                if mapped and re.match(r'M-[A-Z0-9]+', str(mapped), re.I):
                    rows.append(make_row(
                        source='S&W',
                        program_num=cfg['program_num'],
                        customer_ref=cref,
                        dist_id=cfg['dist_id'],
                        bill_date=bill_date,
                        start_date=start_date,
                        end_date=end_date,
                        item=mapped.upper(),
                        qty=qty,
                        amount=amt,
                        trade=cfg['trade']
                    ))
                else:
                    rows.append({'_warning': True, 'code': raw_item, 'desc': desc[:60], 'amount': amt})
    except Exception as e:
        rows.append({'_error': f'S&W: {e}'})
    return rows

def parse_kast_po_report(filepath, cfg, customer_ref):
    """KAST PO Receipts Report format.
    Each line: DATE  PO#  REQ#  PROD#  Description M-CODE  QTY  CS  WEIGHT  UNITCOST  CS  EXTCOST
    M-code is embedded in the description; amount = ExtCost * 4%.
    """
    rows = []
    try:
        with pdfplumber.open(filepath) as pdf:
            all_text = '\n'.join(page.extract_text() or '' for page in pdf.pages)

        # Bill date — first date in the header (report run date)
        bill_date = ''
        m = re.search(r'(\d{1,2}/\d{1,2}/\d{4})', all_text)
        if m: bill_date = to_yyyymmdd(m.group(1))

        # Date range from "Receipt Date (MM-DD-YY - MM-DD-YY)"
        start_date = end_date = ''
        dr = re.search(r'Receipt\s+Date\s*\((\d{2}-\d{2}-\d{2})\s*-\s*(\d{2}-\d{2}-\d{2})\)', all_text, re.I)
        if dr:
            def yy_to_yyyy(s):
                # MM-DD-YY → YYYYMMDD
                parts = s.split('-')
                if len(parts) == 3:
                    mm, dd, yy = parts
                    yyyy = '20' + yy if int(yy) < 50 else '19' + yy
                    return f'{yyyy}{mm}{dd}'
                return ''
            start_date = yy_to_yyyy(dr.group(1))
            end_date   = yy_to_yyyy(dr.group(2))

        # Parse each line for M-code + qty + extended cost
        # Pattern: M-CODE  QTY  CS  WEIGHT  UNITCOST  CS  EXTCOST
        # After M-code there are always: int  CS  int  decimal  CS  decimal
        line_pat = re.compile(
            r'(M-[A-Z0-9]+)\s+(\d+)\s+CS\s+[\d,]+\s+([\d,.]+)\s+CS\s+([\d,.]+)',
            re.I
        )
        seen = set()
        for line in all_text.splitlines():
            # Skip totals / headers
            if re.search(r'Total For|Receipt Date|Product Description|Page \d', line, re.I):
                continue
            m = line_pat.search(line)
            if not m:
                continue
            item     = m.group(1).upper()
            qty      = int(m.group(2))
            ext_cost = clean_amount(m.group(4))
            amount   = round(ext_cost * 0.04, 4)
            rows.append(make_row(
                source='Kast',
                program_num=cfg['program_num'],
                customer_ref=customer_ref,
                dist_id=cfg['dist_id'],
                bill_date=bill_date,
                start_date=start_date,
                end_date=end_date,
                item=item,
                qty=qty,
                amount=amount,
                trade=cfg['trade']
            ))
    except Exception as e:
        rows.append({'_error': f'KAST PO Report: {e}'})
    return rows


def parse_kast(filepath, cfg, customer_ref):
    # Detect format: PO Receipts Report vs invoice
    try:
        with pdfplumber.open(filepath) as pdf:
            first_page = pdf.pages[0].extract_text() or ''
        if 'PO Receipts Report' in first_page or 'Receipt Date' in first_page:
            return parse_kast_po_report(filepath, cfg, customer_ref)
    except Exception:
        pass

    rows = []
    try:
        with pdfplumber.open(filepath) as pdf:
            bill_date = start_date = end_date = ''
            for page in pdf.pages:
                text = page.extract_text() or ''
                if not bill_date:
                    m = re.search(r'(?:Invoice\s*Date|Date)[:\s]+(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})', text, re.I)
                    if m: bill_date = to_yyyymmdd(m.group(1))
                if not start_date:
                    s, e = parse_date_range(text)
                    if s: start_date, end_date = s, e

                # Table rows: M-code  desc  qty  ext_cost
                item_pat = re.compile(
                    r'(M-[A-Z0-9]+)\s+.{0,50?}\s+(\d+)\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)',
                    re.I
                )
                for m in item_pat.finditer(text):
                    item     = m.group(1).upper()
                    qty      = m.group(2)
                    ext_cost = clean_amount(m.group(4))
                    amt      = round(ext_cost * 0.04, 4)
                    rows.append(make_row(
                        source='Kast',
                        program_num=cfg['program_num'],
                        customer_ref=customer_ref,
                        dist_id=cfg['dist_id'],
                        bill_date=bill_date,
                        start_date=start_date,
                        end_date=end_date,
                        item=item,
                        qty=qty,
                        amount=amt,
                        trade=cfg['trade']
                    ))

                # Table extraction fallback
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        if not row: continue
                        cells = [str(c).strip() if c else '' for c in row]
                        item_cell = next((c for c in cells if re.match(r'M-[A-Z0-9]+', c, re.I)), None)
                        if not item_cell: continue
                        nums = [clean_amount(c) for c in cells if re.match(r'[\d,]+\.?\d*$', c.replace(',',''))]
                        if len(nums) >= 3:
                            qty = int(nums[0]) if nums[0] > 0 else 1
                            ext_cost = nums[-1]
                            amt = round(ext_cost * 0.04, 4)
                            # avoid duplicates from text extraction
                            already = any(r.get('Item Number') == item_cell.upper() for r in rows[-10:])
                            if not already:
                                rows.append(make_row(
                                    source='Kast',
                                    program_num=cfg['program_num'],
                                    customer_ref=customer_ref,
                                    dist_id=cfg['dist_id'],
                                    bill_date=bill_date,
                                    start_date=start_date,
                                    end_date=end_date,
                                    item=item_cell.upper(),
                                    qty=qty,
                                    amount=amt,
                                    trade=cfg['trade']
                                ))
    except Exception as e:
        rows.append({'_error': str(e)})
    return rows

def parse_sofo(filepath, cfg, customer_ref):
    """SOFO Foods — may arrive as DOT Foods BB format or Trackmax format; detect by content."""
    if filepath.lower().endswith('.pdf'):
        try:
            with pdfplumber.open(filepath) as pdf:
                first_page = pdf.pages[0].extract_text() or ''
        except Exception as e:
            return [{'_error': f'SOFO: could not read PDF: {e}'}]
        # SOFO files that come through DOT Foods use the BB Dept/BB Vendor format
        if 'BB Dept' in first_page and 'BB Vendor' in first_page:
            return parse_dot_foods_bb(filepath, cfg, customer_ref)
        # Trackmax-hosted SOFO files
        if 'Powered byTrackmax' in first_page or ('Product ID' in first_page and 'DID' in first_page):
            return parse_trackmax(filepath, cfg, customer_ref, source_name='SOFO')
    # Native SOFO format (XLS/XLSX/PDF with direct M-code columns)
    rows = []
    try:
        with pdfplumber.open(filepath) as pdf:
            bill_date = start_date = end_date = ''
            for page in pdf.pages:
                text = page.extract_text() or ''
                if not bill_date:
                    m = re.search(r'(?:Date)[:\s]+(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})', text, re.I)
                    if m: bill_date = to_yyyymmdd(m.group(1))
                if not start_date:
                    s, e = parse_date_range(text)
                    if s: start_date, end_date = s, e
                # Conservative pattern: M-code then explicit qty (small integer) then amounts
                item_pat = re.compile(
                    r'(M-[A-Z0-9]+)\s+\S+\s+(\d{1,4})\s+\$?([\d,.]+)\s+\$?([\d,.]+)',
                    re.I
                )
                for m in item_pat.finditer(text):
                    item = m.group(1).upper()
                    qty  = m.group(2)
                    amt  = clean_amount(m.group(4))
                    rows.append(make_row(
                        source='SOFO',
                        program_num=cfg['program_num'],
                        customer_ref=customer_ref,
                        dist_id=cfg['dist_id'],
                        bill_date=bill_date,
                        start_date=start_date,
                        end_date=end_date,
                        item=item,
                        qty=qty,
                        amount=amt,
                        trade=cfg['trade']
                    ))
    except Exception as e:
        rows.append({'_error': str(e)})
    return rows

def parse_pfs(filepath, cfg, customer_ref):
    """PFS Roma / PFG — uses Trackmax format (same engine as Driscoll Foods)."""
    if filepath.lower().endswith('.pdf'):
        # PFS Roma uses Powered-by-Trackmax format; route through the generic Trackmax parser.
        return parse_trackmax(filepath, cfg, customer_ref, source_name='PFS')
    # Non-PDF fallback (CSV/XLSX) — generic M-code search
    rows = []
    try:
        import pandas as pd
        df = pd.read_csv(filepath, dtype=str) if filepath.lower().endswith('.csv') else pd.read_excel(filepath, dtype=str)
        for _, r in df.iterrows():
            row_str = ' '.join(str(v) for v in r.values if v)
            m = re.search(r'(M-[A-Z0-9]+)', row_str, re.I)
            if not m: continue
            qty_m = re.search(r'\b(\d+)\b', row_str)
            amt_m = re.search(r'\$?([\d,.]+)\s*$', row_str)
            rows.append(make_row(source='PFS', program_num=cfg['program_num'],
                customer_ref=customer_ref, dist_id=cfg['dist_id'],
                item=m.group(1).upper(),
                qty=qty_m.group(1) if qty_m else 0,
                amount=clean_amount(amt_m.group(1)) if amt_m else 0,
                trade=cfg['trade']))
    except Exception as e:
        rows.append({'_error': str(e)})
    return rows

def parse_yhata(filepath, cfg, customer_ref):
    rows = []
    try:
        with pdfplumber.open(filepath) as pdf:
            bill_date = start_date = end_date = ''
            for page in pdf.pages:
                text = page.extract_text() or ''
                if not bill_date:
                    m = re.search(r'(?:Date|Dated)[:\s]+(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})', text, re.I)
                    if m: bill_date = to_yyyymmdd(m.group(1))
                if not start_date:
                    s, e = parse_date_range(text)
                    if s: start_date, end_date = s, e

                # Pattern: M-code  desc  UPC(8-20 digits)  qty  price  ext  allowance
                detail_pat = re.compile(
                    r'(M-[A-Z0-9]+)\s+.{0,60?}\s+\d{8,20}\s+(\d+)\s+\$?([\d,.]+)\s+\$?([\d,.]+)\s+\(?\$?([\d,.]+)\)?',
                    re.I
                )
                detail_pat2 = re.compile(
                    r'(M-[A-Z0-9]+)\s+.{0,60?}\s+\d+\s+(\d+)\s+\$?([\d,.]+)\s+\$?([\d,.]+)\s+\(?\$?([\d,.]+)\)?',
                    re.I
                )
                matched = set()
                for pat in [detail_pat, detail_pat2]:
                    for m in pat.finditer(text):
                        item = m.group(1).upper()
                        if item in matched: continue
                        matched.add(item)
                        qty  = m.group(2)
                        amt  = clean_amount(m.group(5))
                        rows.append(make_row(
                            source='Y.Hata',
                            program_num=cfg['program_num'],
                            customer_ref=customer_ref,
                            dist_id=cfg['dist_id'],
                            bill_date=bill_date,
                            start_date=start_date,
                            end_date=end_date,
                            item=item,
                            qty=qty,
                            amount=amt,
                            trade=cfg['trade']
                        ))
    except Exception as e:
        rows.append({'_error': str(e)})
    return rows

def parse_labatt(filepath, cfg, customer_ref):
    rows = []
    try:
        with pdfplumber.open(filepath) as pdf:
            bill_date = start_date = end_date = ''
            for page in pdf.pages:
                text = page.extract_text() or ''
                if not bill_date:
                    m = re.search(r'(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})', text)
                    if m: bill_date = to_yyyymmdd(m.group(1))
                if not start_date:
                    s, e = parse_date_range(text)
                    if s: start_date, end_date = s, e

                item_pat = re.compile(
                    r'(M-[A-Z0-9]+)\s+\S+\s+(\d+)\s+\$?([\d,.]+)\s+\$?([\d,.]+)',
                    re.I
                )
                for m in item_pat.finditer(text):
                    item = m.group(1).upper()
                    qty  = m.group(2)
                    amt  = clean_amount(m.group(4))
                    rows.append(make_row(
                        source='Labatt',
                        program_num=cfg['program_num'],
                        customer_ref=customer_ref,
                        dist_id=cfg['dist_id'],
                        bill_date=bill_date,
                        start_date=start_date,
                        end_date=end_date,
                        item=item,
                        qty=qty,
                        amount=amt,
                        trade=cfg['trade']
                    ))
    except Exception as e:
        rows.append({'_error': str(e)})
    return rows

def parse_trackmax(filepath, cfg, customer_ref, source_name=''):
    """Generic Trackmax 'Supplier Billback' PDF parser (Driscoll Foods, PFS Roma, etc.).
    Columns: Inv.Number Inv.Date CustomerID CustomerName Brand PackSize Description
             ProductID(M-code) DID UPC Quantity TotalWeight TotalCharges ProgramAmount AmountDue
    Credit/return lines have parenthesized quantities and amounts, e.g. (1.00) and ($9.04).
    These must be captured as NEGATIVE amounts to match the Grand Total.
    source_name: the row Source label; auto-detected from PDF header if empty.
    """
    rows = []
    try:
        with pdfplumber.open(filepath) as pdf:
            all_text = '\n'.join(page.extract_text() or '' for page in pdf.pages)

        # Auto-detect distributor name from header ("IMPORTANT!\n<Company Name>")
        if not source_name:
            m = re.search(r'IMPORTANT!\s*\n([^\n]+)', all_text)
            source_name = m.group(1).strip() if m else 'Trackmax Dist'

        # Statement date and program date range
        bill_date = start_date = end_date = ''
        m = re.search(r'(?:generated|posted)\s+on\s+(\d{1,2}/\d{1,2}/\d{4})', all_text, re.I)
        if m: bill_date = to_yyyymmdd(m.group(1))
        m = re.search(r'between\s+(\d{1,2}/\d{1,2}/\d{4})\s+and\s+(\d{1,2}/\d{1,2}/\d{4})', all_text, re.I)
        if m:
            start_date = to_yyyymmdd(m.group(1))
            end_date   = to_yyyymmdd(m.group(2))
        # Also check "Start Date: / Stop Date:" header format
        if not start_date:
            m2 = re.search(r'Start\s+Date[:\s]+(\d{1,2}/\d{1,2}/\d{4}).*?Stop\s+Date[:\s]+(\d{1,2}/\d{1,2}/\d{4})', all_text, re.I | re.S)
            if m2:
                start_date = to_yyyymmdd(m2.group(1))
                end_date   = to_yyyymmdd(m2.group(2))

        # Invoice number for customer ref
        inv_num = ''
        m2 = re.search(r'Our\s+Invoice\s+Number[:\s]+(\w+)', all_text, re.I)
        if m2: inv_num = m2.group(1).strip()
        cref = customer_ref or inv_num

        # Data line pattern — matches both regular and credit/return lines:
        #   Regular:  019926 02/13/2026 ... M-FR035F 350377 10738337884136 1.00  11.53 $31.44 BB To 22.400/unit $9.04
        #   Credit:   X97707 02/20/2026 ... M-FR035F 350377 10738337884136 (1.00) (11.53) ($31.44) BB To 22.400/unit ($9.04)
        # Also handles truncated UPC (10-11 chars) when pdfplumber wraps the line.
        line_pat = re.compile(
            r'(M-[A-Z0-9]+)\s+\d{4,9}\s+\d{8,14}\s+'   # M-code, DID (4-9d), UPC (8-14d)
            r'(\([\d.]+\)|[\d.]+)\s+'                    # qty: positive or (negative)
            r'[\d.()]+\s+'                               # weight (ignore)
            r'(?:\(\$[\d,.]+\)|\$[-\d,.]+)\s+'          # Total Charges (ignore, may be negative $-X.XX)
            r'\S.*?'                                     # Program Amount text
            r'(\(\$[\d,.]+\)|\$[-\d,.]+)',               # Amount Due (last $ value)
            re.I
        )

        seen_lines = set()
        for line in all_text.splitlines():
            if re.search(r'^Totals for|^Grand Total|Powered by|Invoice is Due', line, re.I):
                continue
            m = line_pat.search(line)
            if not m:
                continue
            if line in seen_lines:
                continue
            seen_lines.add(line)

            item = m.group(1).upper()
            raw_qty = m.group(2)
            raw_amt = m.group(3)

            # Parenthesized qty means a return (negative)
            if raw_qty.startswith('('):
                qty = -float(raw_qty.strip('()'))
            else:
                qty = float(raw_qty)

            amount = clean_amount(raw_amt)  # handles ($9.04) and $-2.80 → negative

            rows.append(make_row(
                source=source_name,
                program_num=cfg['program_num'],
                customer_ref=cref,
                dist_id=cfg['dist_id'],
                bill_date=bill_date,
                start_date=start_date,
                end_date=end_date,
                item=item,
                qty=qty,
                amount=amount,
                trade=cfg['trade']
            ))

        if not rows:
            rows.append({'_error': f'{source_name} PDF: no M-code rows found'})
    except Exception as e:
        rows.append({'_error': f'Trackmax parser: {e}'})
    return rows


def parse_driscoll(filepath, cfg, customer_ref):
    """Driscoll Foods wrapper — uses generic Trackmax parser."""
    return parse_trackmax(filepath, cfg, customer_ref, source_name='Driscoll Foods')


def parse_martin_bros(filepath, cfg, customer_ref):
    """Martin Bros. Dist. Co. 'Supplier Billback' PDF format.
    Columns: InvNum PO# InvDate RcvdDate Supplier Brand PackSize Desc M-CODE DID UPC Qty FOB $DEL $DEL %program $Amount
    M-code is Product ID; billback amount is last $X.XX on the line.
    """
    rows = []
    try:
        with pdfplumber.open(filepath) as pdf:
            all_text = '\n'.join(page.extract_text() or '' for page in pdf.pages)

        # Statement date and program date range
        bill_date = start_date = end_date = ''
        m = re.search(r'generated on\s+(\d{1,2}/\d{1,2}/\d{4})', all_text, re.I)
        if m: bill_date = to_yyyymmdd(m.group(1))
        m = re.search(r'between\s+(\d{1,2}/\d{1,2}/\d{4})\s+and\s+(\d{1,2}/\d{1,2}/\d{4})', all_text, re.I)
        if m:
            start_date = to_yyyymmdd(m.group(1))
            end_date   = to_yyyymmdd(m.group(2))

        # Invoice number for customer ref
        inv_num = ''
        m2 = re.search(r'Invoice\s+Number[:\s]+(\d+)', all_text, re.I)
        if not m2:
            m2 = re.search(r'Our\s+Invoice\s+Number[:\s\n]+(\d+)', all_text, re.I)
        if m2: inv_num = m2.group(1).strip()
        cref = customer_ref or inv_num

        # Each data line: ... M-CODE  DID(8-9d)  UPC(12-14d)  QTY  FOB  $DEL  $DEL  X%...  $AMOUNT
        # After M-code: short DID, long UPC, then qty as float
        line_pat = re.compile(
            r'(M-[A-Z0-9]+)\s+\d{5,9}\s+\d{10,14}\s+([\d.]+)\s+[\d.]+\s+\$([\d,.]+)\s+\$([\d,.]+)\s+[\d.]+%.*?\$([\d,.]+)',
            re.I
        )
        seen_lines = set()   # deduplicate page-boundary carryover lines
        for line in all_text.splitlines():
            if re.search(r'^Totals for|^Invoice|Inv\.\s*Number|program activity', line, re.I):
                continue
            m = line_pat.search(line)
            if not m: continue
            item   = m.group(1).upper()
            qty    = m.group(2)
            amount = clean_amount(m.group(5))   # last dollar column = billback amount
            # Lines at page breaks appear on both the bottom of one page and the top of
            # the next; deduplicate by exact raw line content
            if line in seen_lines:
                continue
            seen_lines.add(line)
            rows.append(make_row(
                source='Martin Bros',
                program_num=cfg['program_num'],
                customer_ref=cref,
                dist_id=cfg['dist_id'],
                bill_date=bill_date,
                start_date=start_date,
                end_date=end_date,
                item=item,
                qty=qty,
                amount=amount,
                trade=cfg['trade']
            ))
        if not rows:
            rows.append({'_error': 'Martin Bros PDF: no rows extracted'})
    except Exception as e:
        rows.append({'_error': f'Martin Bros: {e}'})
    return rows


def parse_dot_foods_bb(filepath, cfg, customer_ref):
    """DOT Foods 'Supplier Billback' multi-vendor report.
    Columns: BB Dept | BB Vendor # | BB Vendor Name | Chain | Item | Item Description | Vendor Item# | ...
    Aggregates by item using subtotal lines. Vendor Item# may contain M-code or a numeric code.
    Items without M-codes (including non-Monin products) are flagged as warnings.
    """
    rows = []
    try:
        with pdfplumber.open(filepath) as pdf:
            all_text = '\n'.join(p.extract_text() or '' for p in pdf.pages)

        # Date range from invoice dates in file
        bill_date = start_date = end_date = TODAY
        dates_found = re.findall(r'\b(\d{1,2}/\d{1,2}/\d{4})\b', all_text)
        if dates_found:
            from datetime import datetime
            dt_list = []
            for d in dates_found:
                try: dt_list.append(datetime.strptime(d, '%m/%d/%Y'))
                except: pass
            if dt_list:
                start_date = to_yyyymmdd(min(dt_list).strftime('%m/%d/%Y'))
                end_date   = to_yyyymmdd(max(dt_list).strftime('%m/%d/%Y'))

        # Customer ref: use passed value or DocID from filename
        cref = customer_ref
        if not cref:
            m = re.search(r'DocID\s+([\w]+)', os.path.basename(filepath), re.I)
            if m: cref = m.group(1)

        item_map = cfg.get('item_map', {})   # user-supplied DOT item# → M-code

        # Pass 1: build item# → {mcode, desc} from detail lines
        item_mcode = {}
        item_desc  = {}
        for line in all_text.splitlines():
            if re.search(r'\bTotal\b|BB Dept|BB Vendor', line, re.I):
                continue
            # Item# followed by description followed by M-code in Vendor Item# slot
            m = re.search(r'(\d{5,6})([A-Z][A-Z0-9\s/]+?)\s+(M-[A-Z0-9]+)\s+\d{6,9}', line, re.I)
            if m:
                item_mcode[m.group(1)] = m.group(3).upper()
                item_desc[m.group(1)]  = m.group(2).strip()
                continue
            # Item# followed by description with numeric Vendor Item# — capture desc only
            m2 = re.search(r'(\d{5,6})([A-Z][A-Z0-9\s/.-]+?)\s+\d{4,8}\s+\d{7,9}[A-Z]', line)
            if m2 and m2.group(1) not in item_desc:
                item_desc[m2.group(1)] = m2.group(2).strip()

        # Pass 2: process item-level Total lines (one row per item)
        # Formats: "280930 Total 7 - $ 34.65"  /  "140078 Total 28 280.00 $ 140.00"  /  "280978 Total 1 3 - $ 36.55"
        # Handle variable number of numeric/dash tokens before the final $ amount
        total_pat = re.compile(r'^(\d{5,6})\s+Total\s+(\d+)(?:\s+[-\d,.]+)*\s+\$\s*([\d,.]+)', re.M)
        for m in total_pat.finditer(all_text):
            item_num = m.group(1)
            qty      = m.group(2)
            amount   = clean_amount(m.group(3))
            mcode    = item_mcode.get(item_num) or item_map.get(item_num)
            desc     = item_desc.get(item_num, item_num)
            if mcode and re.match(r'M-[A-Z0-9]+', str(mcode), re.I):
                rows.append(make_row(
                    source='DOT Foods BB',
                    program_num=cfg['program_num'],
                    customer_ref=cref,
                    dist_id=cfg['dist_id'],
                    bill_date=bill_date,
                    start_date=start_date,
                    end_date=end_date,
                    item=str(mcode).upper(),
                    qty=qty,
                    amount=amount,
                    trade=cfg['trade']
                ))
            else:
                rows.append({'_warning': True, 'code': item_num, 'desc': desc, 'amount': amount})

        if not rows:
            rows.append({'_error': 'DOT Foods BB: no items found'})
    except Exception as e:
        rows.append({'_error': f'DOT Foods BB: {e}'})
    return rows


def parse_supplier_billback_pdf(filepath, cfg, customer_ref):
    """Content-based dispatcher for 'Supplier Billback' PDFs — multiple distributors use this filename."""
    try:
        with pdfplumber.open(filepath) as pdf:
            first_page = pdf.pages[0].extract_text() or ''
    except Exception as e:
        return [{'_error': f'Could not read PDF: {e}'}]

    if 'Martin Bros' in first_page:
        return parse_martin_bros(filepath, cfg, customer_ref)
    if 'Driscoll Foods' in first_page:
        return parse_trackmax(filepath, cfg, customer_ref, source_name='Driscoll Foods')
    # Any Trackmax-format distributor: auto-detect company name from header
    if 'Powered byTrackmax' in first_page or ('Product ID' in first_page and 'DID' in first_page and 'UPC' in first_page):
        return parse_trackmax(filepath, cfg, customer_ref)
    if 'BB Dept' in first_page and 'BB Vendor' in first_page:
        return parse_dot_foods_bb(filepath, cfg, customer_ref)
    if 'Harbor Food' in first_page:
        # Harbor sends XLS usually; PDF variant — generic fallback
        return [{'_error': 'Harbor Foodservice PDF format not yet supported — please send the XLSX version'}]

    # Unknown distributor — try generic M-code extraction as best effort
    rows = []
    try:
        with pdfplumber.open(filepath) as pdf:
            all_text = '\n'.join(page.extract_text() or '' for page in pdf.pages)
        bill_date = start_date = end_date = ''
        m = re.search(r'(\d{1,2}/\d{1,2}/\d{4})', all_text)
        if m: bill_date = to_yyyymmdd(m.group(1))
        s, e = parse_date_range(all_text)
        if s: start_date, end_date = s, e
        for line in all_text.splitlines():
            if re.search(r'total|header|description', line, re.I): continue
            m = re.search(r'(M-[A-Z0-9]+).*?([\d.]+)\s+.*?\$([\d,.]+)\s*$', line, re.I)
            if not m: continue
            rows.append(make_row(
                source='Dist PDF',
                program_num=cfg['program_num'],
                customer_ref=customer_ref,
                dist_id=cfg['dist_id'],
                bill_date=bill_date,
                start_date=start_date,
                end_date=end_date,
                item=m.group(1).upper(),
                qty=m.group(2),
                amount=clean_amount(m.group(3)),
                trade=cfg['trade']
            ))
        if not rows:
            rows.append({'_error': f'Unknown distributor PDF — could not identify company from content. Please add support for this format.'})
    except Exception as e:
        rows.append({'_error': f'Supplier Billback PDF: {e}'})
    return rows


def parse_harbor(filepath, cfg, customer_ref):
    """Harbor Foodservice 'Supplier Billback' — XLSX only; PDFs routed by content."""
    if filepath.lower().endswith('.pdf'):
        return parse_supplier_billback_pdf(filepath, cfg, customer_ref)
    rows = []
    try:
        wb = __import__('openpyxl').load_workbook(filepath, data_only=True)
        ws = wb.active

        # Find invoice date (row with 'Invoice Date:' label, typically row 6)
        bill_date = ''
        period_str = ''
        for row in ws.iter_rows(max_row=10, values_only=True):
            for i, cell in enumerate(row):
                if isinstance(cell, str) and 'invoice date' in cell.lower():
                    # Date value is in the next cell
                    if i + 1 < len(row) and row[i+1] is not None:
                        bill_date = to_yyyymmdd(row[i+1])
                    break

        # Find header row (has 'Vendor Item Num' or 'Billback Amt')
        header_row_idx = None
        col_map = {}
        for i, row in enumerate(ws.iter_rows(values_only=True), 1):
            row_vals = [str(v).strip() if v is not None else '' for v in row]
            if any('vendor item' in c.lower() for c in row_vals):
                header_row_idx = i
                col_map = {v.strip().lower(): j for j, v in enumerate(row_vals)}
                break

        if header_row_idx is None:
            return [{'_error': 'Harbor: could not find header row'}]

        # Column indices
        def gc(*keys):
            for k in keys:
                for ck, ci in col_map.items():
                    if k.lower() in ck: return ci
            return None

        ci_mcode   = gc('vendor item num', 'vendor item')
        ci_period  = gc('invoice period', 'period')
        ci_recqty  = gc('rec qty', 'received qty')
        ci_ordqty  = gc('order qty')
        ci_bbamt   = gc('billback amt', 'billback amount')
        ci_invoice = gc('invoice #', 'invoice#')

        for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
            if not any(v is not None for v in row): continue
            cells = [str(v).strip() if v is not None else '' for v in row]

            # Skip total row
            if any('total' in c.lower() for c in cells if c): continue

            # Get M-code — extract first M-XXXXX pattern from cell
            raw_mcode = cells[ci_mcode] if ci_mcode is not None else ''
            m = re.search(r'M-[A-Z0-9]+', raw_mcode, re.I)
            if not m:
                # Check user-defined item code mapping
                item_map = cfg.get('item_map', {})
                mapped = item_map.get(raw_mcode.strip(), '')
                if mapped and re.match(r'M-[A-Z0-9]+', mapped, re.I):
                    item = mapped.upper()
                else:
                    # Still unknown — flag it as a structured warning
                    if raw_mcode and raw_mcode not in ('', 'None'):
                        desc    = cells[10] if len(cells) > 10 else ''
                        amt_val = clean_amount(cells[ci_bbamt]) if ci_bbamt is not None else 0.0
                        rows.append({'_warning': True, 'code': raw_mcode, 'desc': desc, 'amount': amt_val})
                    continue
            else:
                item = m.group(0).upper()

            # Period for date range
            if ci_period is not None and not period_str:
                period_str = cells[ci_period]
            start_date, end_date = month_name_to_range(cells[ci_period] if ci_period is not None else period_str)

            # Qty: prefer rec qty, fall back to order qty
            qty_val = cells[ci_recqty] if ci_recqty is not None else ''
            if not qty_val and ci_ordqty is not None:
                qty_val = cells[ci_ordqty]

            # Amount
            amt_val = clean_amount(cells[ci_bbamt]) if ci_bbamt is not None else 0.0

            # Invoice # as customer ref if not provided
            inv_num = cells[ci_invoice] if ci_invoice is not None else ''
            cref = customer_ref or inv_num

            rows.append(make_row(
                source='Harbor',
                program_num=cfg['program_num'],
                customer_ref=cref,
                dist_id=cfg['dist_id'],
                bill_date=bill_date,
                start_date=start_date,
                end_date=end_date,
                item=item,
                qty=qty_val,
                amount=amt_val,
                trade=cfg['trade']
            ))
    except Exception as e:
        rows.append({'_error': f'Harbor: {e}'})
    return rows


# ─── DISPATCHER ──────────────────────────────────────────────────────────────
def detect_and_parse(filepath, user_config=None, customer_ref='', file_override=None, original_filename=None):
    # Honour user-selected supplier from the UI dropdown; fall back to auto-detection
    forced = (file_override or {}).get('supplier', '')
    if forced and forced != 'UNKNOWN':
        supplier = forced
    else:
        # Use original filename for detection — safe_name (spaces→underscores) breaks some patterns
        supplier = detect_supplier(original_filename or filepath)
    cfg = dict(DEFAULT_SUPPLIER_CONFIG.get(supplier, DEFAULT_SUPPLIER_CONFIG['UNKNOWN']))
    if user_config and supplier in user_config:
        uc = user_config[supplier]
        if uc.get('program_num'): cfg['program_num'] = uc['program_num']
        if uc.get('dist_id'):     cfg['dist_id']     = uc['dist_id']
        if uc.get('trade'):       cfg['trade']       = uc['trade']
        if uc.get('item_map'):    cfg['item_map']    = uc['item_map']  # {harbor_code: m_code}
    # Per-file overrides take highest priority (entered by user in the file row)
    if file_override:
        if file_override.get('program_num'): cfg['program_num'] = file_override['program_num']
        if file_override.get('dist_id'):     cfg['dist_id']     = file_override['dist_id']
        if file_override.get('customer_ref') and not customer_ref:
            customer_ref = file_override['customer_ref']

    fn = os.path.basename(filepath).upper()
    ext = os.path.splitext(filepath)[1].lower()

    if supplier == 'BEK':
        return supplier, parse_bek_or_nich(filepath, 'BEK', cfg, customer_ref)
    elif supplier == 'NICH_CO':
        return supplier, parse_bek_or_nich(filepath, 'Nich&Co', cfg, customer_ref)
    elif supplier == 'SHAMROCK':
        return supplier, parse_shamrock(filepath, cfg, customer_ref)
    elif supplier == 'DOT_CBBB':
        return supplier, parse_dot_cbbb(filepath, cfg, customer_ref)
    elif supplier in ('MCLANE','MCLANE_OR_DOT'):
        return supplier, parse_mclane(filepath, cfg, customer_ref)
    elif supplier == 'S_AND_W':
        return supplier, parse_sw(filepath, cfg, customer_ref)
    elif supplier == 'KAST':
        return supplier, parse_kast(filepath, cfg, customer_ref)
    elif supplier == 'SOFO':
        return supplier, parse_sofo(filepath, cfg, customer_ref)
    elif supplier == 'PFS' or supplier == 'PFS_STYLE':
        return supplier, parse_pfs(filepath, cfg, customer_ref)
    elif supplier == 'Y_HATA':
        return supplier, parse_yhata(filepath, cfg, customer_ref)
    elif supplier == 'LABATT':
        return supplier, parse_labatt(filepath, cfg, customer_ref)
    elif supplier == 'HARBOR':
        result = parse_harbor(filepath, cfg, customer_ref)
        # Content-based routing may have identified a different distributor — label accordingly
        if result and isinstance(result[0], dict):
            src = result[0].get('Source', '')
            if src == 'Martin Bros':
                return 'MARTIN_BROS', result
            if src == 'Driscoll Foods':
                return 'DRISCOLL', result
        return supplier, result
    elif supplier == 'MARTIN_BROS':
        return supplier, parse_martin_bros(filepath, cfg, customer_ref)
    elif supplier == 'DRISCOLL':
        return supplier, parse_driscoll(filepath, cfg, customer_ref)
    elif supplier == 'BLAIR_CANDY':
        return supplier, [{'_error': 'Blair Candy uses scanned/image PDFs — text extraction not supported. Please enter manually.'}]
    else:
        # Try generic PDF
        if ext == '.pdf':
            return supplier, parse_pfs(filepath, cfg, customer_ref)
        return supplier, [{'_error': f'Unknown supplier format: {os.path.basename(filepath)}'}]

# ─── OUTPUT BUILDER ──────────────────────────────────────────────────────────
def build_output(all_rows):
    """Write all_rows (list of dicts) to a new workbook with exact Sheet1 format."""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'

    # Header row — bold
    for col_idx, col_name in enumerate(S1_COLS, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = Font(bold=True)

    # Data rows
    for row_idx, row_dict in enumerate(all_rows, 2):
        if '_error' in row_dict:
            ws.cell(row=row_idx, column=1, value=f"ERROR: {row_dict['_error']}")
            continue
        for col_idx, col_name in enumerate(S1_COLS, 1):
            val = row_dict.get(col_name)
            ws.cell(row=row_idx, column=col_idx, value=val)

    # Auto-width
    for col_idx, col_name in enumerate(S1_COLS, 1):
        max_len = max(len(str(col_name)), 10)
        from openpyxl.utils import get_column_letter
        ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 2

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

# ─── HTML UI ─────────────────────────────────────────────────────────────────
HTML_PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Monin Billback Processor</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Segoe UI',Arial,sans-serif;background:#f0f2f5;color:#222}
  .header{background:linear-gradient(135deg,#1a3a5c,#2e6da4);color:#fff;padding:20px 32px;display:flex;align-items:center;gap:16px}
  .header h1{font-size:1.4rem;font-weight:600}
  .header .sub{font-size:.85rem;opacity:.8}
  .container{max-width:1100px;margin:24px auto;padding:0 16px}
  .card{background:#fff;border-radius:10px;box-shadow:0 2px 8px rgba(0,0,0,.08);padding:24px;margin-bottom:20px}
  .card h2{font-size:1rem;font-weight:600;color:#1a3a5c;margin-bottom:16px;border-bottom:2px solid #e8eef4;padding-bottom:8px}
  .drop-zone{border:2px dashed #2e6da4;border-radius:8px;padding:36px;text-align:center;cursor:pointer;transition:.2s;background:#f8fafc}
  .drop-zone:hover,.drop-zone.drag-over{background:#e8f0fb;border-color:#1a3a5c}
  .drop-zone svg{color:#2e6da4;margin-bottom:8px}
  .drop-zone p{color:#666;font-size:.9rem}
  .drop-zone strong{color:#2e6da4}
  #file-input{display:none}
  .file-list{margin-top:12px}
  .file-item{display:flex;flex-direction:column;gap:6px;padding:10px 14px;background:#f8fafc;border-radius:6px;margin-bottom:8px;border:1px solid #e2e8f0}
  .file-item-top{display:flex;align-items:center;gap:10px;width:100%}
  .file-item-fields{display:flex;gap:10px;flex-wrap:wrap;align-items:center;padding:6px 8px;background:#f1f5f9;border-radius:5px;border:1px dashed #cbd5e1}
  .file-item-fields label{font-size:.72rem;color:#64748b;white-space:nowrap;display:flex;align-items:center;gap:4px}
  .file-badge{font-size:.7rem;font-weight:700;padding:2px 7px;border-radius:10px;background:#dbeafe;color:#1d4ed8;white-space:nowrap}
  .file-badge.unknown{background:#fef3c7;color:#92400e}
  .file-badge.error{background:#fee2e2;color:#991b1b}
  .file-name{flex:1;font-size:.85rem;font-weight:500;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .cref-input,.override-input{font-size:.8rem;border:1px solid #cbd5e1;border-radius:4px;padding:4px 8px;background:#fff}
  .cref-input{width:200px}
  .override-input{width:120px}
  /* Searchable distributor combobox */
  .combo-wrap{position:relative;display:inline-block;min-width:180px;max-width:240px}
  .combo-input{font-size:.78rem;font-weight:700;border:1px solid #93c5fd;border-radius:10px;padding:3px 26px 3px 10px;background:#dbeafe;color:#1d4ed8;cursor:text;width:100%;box-sizing:border-box;outline:none}
  .combo-input::placeholder{color:#93c5fd;font-weight:400}
  .combo-input.unknown{background:#fef9c3;border-color:#f59e0b;color:#92400e}
  .combo-arrow{position:absolute;right:7px;top:50%;transform:translateY(-50%);pointer-events:none;font-size:.65rem;color:#93c5fd}
  .combo-list{position:absolute;left:0;top:calc(100% + 3px);width:240px;max-height:220px;overflow-y:auto;background:#fff;border:1px solid #93c5fd;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,.12);z-index:9999;display:none}
  .combo-item{padding:5px 12px;font-size:.8rem;cursor:pointer;color:#1e293b}
  .combo-item:hover,.combo-item.focused{background:#dbeafe;color:#1d4ed8}
  .combo-item.no-match{color:#94a3b8;font-style:italic;cursor:default}
  .remove-btn{background:none;border:none;cursor:pointer;color:#94a3b8;font-size:1rem;padding:2px 6px}
  .remove-btn:hover{color:#dc2626}
  .config-toggle{font-size:.8rem;color:#2e6da4;cursor:pointer;text-decoration:underline;margin-bottom:10px;display:inline-block}
  .config-section{display:none}
  .config-section.open{display:block}
  .config-table{width:100%;border-collapse:collapse;font-size:.82rem}
  .config-table th{background:#f1f5f9;color:#475569;font-weight:600;padding:7px 10px;text-align:left;border-bottom:2px solid #e2e8f0}
  .config-table td{padding:6px 10px;border-bottom:1px solid #f1f5f9}
  .config-table input,.config-table select{font-size:.8rem;border:1px solid #cbd5e1;border-radius:4px;padding:3px 7px;width:100%}
  .btn{display:inline-flex;align-items:center;gap:8px;padding:11px 24px;border-radius:6px;font-weight:600;font-size:.9rem;cursor:pointer;border:none;transition:.2s}
  .btn-primary{background:#2e6da4;color:#fff}
  .btn-primary:hover{background:#1a3a5c}
  .btn-primary:disabled{background:#94a3b8;cursor:not-allowed}
  .btn-secondary{background:#f1f5f9;color:#334155;border:1px solid #e2e8f0}
  .btn-secondary:hover{background:#e2e8f0}
  .actions{display:flex;gap:12px;margin-top:8px}
  .progress{display:none;margin-top:16px}
  .progress-bar{height:6px;background:#e2e8f0;border-radius:3px;overflow:hidden;margin-bottom:8px}
  .progress-fill{height:100%;background:linear-gradient(90deg,#2e6da4,#60a5fa);width:0%;transition:width .3s;border-radius:3px}
  .progress-text{font-size:.82rem;color:#64748b}
  .results{display:none;margin-top:16px}
  .result-row{display:flex;align-items:center;justify-content:space-between;padding:8px 12px;border-radius:5px;margin-bottom:5px;font-size:.84rem}
  .result-ok{background:#f0fdf4;border-left:3px solid #22c55e}
  .result-err{background:#fff1f2;border-left:3px solid #f43f5e}
  .result-skip{background:#fffbeb;border-left:3px solid #f59e0b}
  .count-badge{font-weight:700;font-size:.8rem;padding:1px 8px;border-radius:8px;background:#dbeafe;color:#1d4ed8}
  .dl-box{background:linear-gradient(135deg,#f0fdf4,#dcfce7);border:1px solid #86efac;border-radius:8px;padding:16px 20px;text-align:center;margin-top:16px}
  .dl-box p{font-size:.9rem;color:#166534;margin-bottom:10px}
  .total-bar{background:#1a3a5c;color:#fff;border-radius:6px;padding:8px 16px;display:inline-block;font-weight:700;margin-bottom:12px;font-size:.95rem}
</style>
</head>
<body>
<div class="header">
  <img src="data:image/png;base64,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LMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMAiEzn2uAfq614AAAAASUVORK5CYII=" alt="Monin" style="height:48px;object-fit:contain;flex-shrink:0;filter:brightness(0) invert(1)">
  <div class="sub" style="color:rgba(255,255,255,.75);font-size:.88rem;margin-left:14px">Supplier Billback Processor</div>
</div>

<div class="container">

  <!-- CONFIG PANEL -->
  <div class="card">
    <h2>⚙️ Supplier Configuration</h2>
    <div style="display:flex;align-items:center;gap:16px;margin-bottom:10px">
      <span class="config-toggle" onclick="toggleConfig()" style="margin-bottom:0">▶ Show / Edit Supplier Settings</span>
      <button class="btn btn-secondary" onclick="saveConfig()" style="padding:5px 14px;font-size:.8rem">💾 Save Settings</button>
    </div>
    <div style="font-size:.78rem;color:#94a3b8;margin-bottom:8px">Settings are saved in your browser and restored automatically next time.</div>
    <div class="config-section" id="config-section">
      <table class="config-table">
        <thead><tr><th>Supplier</th><th>Program #</th><th>Distributor ID</th><th>Trade</th></tr></thead>
        <tbody id="config-tbody"></tbody>
      </table>
    </div>
  </div>

  <!-- FILE DROP -->
  <div class="card">
    <h2>📂 Upload Billback Files</h2>
    <div class="drop-zone" id="drop-zone" onclick="document.getElementById('file-input').click()">
      <svg width="40" height="40" fill="none" stroke="currentColor" stroke-width="1.5" viewBox="0 0 24 24"><path d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"/></svg>
      <p><strong>Click to browse</strong> or drag &amp; drop files here</p>
      <p style="margin-top:4px;font-size:.8rem">Supports: .xlsx, .xls, .txt, .tab, .pdf</p>
    </div>
    <input type="file" id="file-input" multiple accept=".xlsx,.xls,.txt,.tab,.pdf,.csv">
    <div class="file-list" id="file-list"></div>
    <div class="actions">
      <button class="btn btn-primary" id="process-btn" onclick="processFiles()" disabled>
        <span>▶ Process Files</span>
      </button>
      <button class="btn btn-secondary" onclick="clearAll()">✕ Clear All</button>
    </div>
    <div class="progress" id="progress">
      <div class="progress-bar"><div class="progress-fill" id="progress-fill"></div></div>
      <div class="progress-text" id="progress-text">Processing...</div>
    </div>
    <div class="results" id="results"></div>
  </div>

</div>

<script>
const SUPPLIERS = ['KAST','SOFO','PFS','LABATT','Y_HATA','BEK','NICH_CO','SHAMROCK','DOT_CBBB','MCLANE','S_AND_W','HARBOR','MARTIN_BROS','DOT_FOODS_BB','DRISCOLL'];
const DEFAULT_CFG = {
  KAST:    {program_num:'1004089', dist_id:'134810000', trade:'D'},
  SOFO:    {program_num:'', dist_id:'', trade:'D'},
  PFS:     {program_num:'', dist_id:'', trade:'D'},
  LABATT:  {program_num:'', dist_id:'', trade:'D'},
  Y_HATA:  {program_num:'', dist_id:'', trade:'D'},
  BEK:     {program_num:'', dist_id:'', trade:'D'},
  NICH_CO: {program_num:'', dist_id:'', trade:'D'},
  SHAMROCK:{program_num:'', dist_id:'', trade:'D'},
  DOT_CBBB:{program_num:'', dist_id:'', trade:'D'},
  MCLANE:  {program_num:'', dist_id:'', trade:'D'},
  S_AND_W: {program_num:'', dist_id:'', trade:'D'},
  HARBOR:      {program_num:'', dist_id:'', trade:'D'},
  MARTIN_BROS:  {program_num:'', dist_id:'', trade:'D'},
  DOT_FOODS_BB: {program_num:'', dist_id:'', trade:'D'},
  DRISCOLL:     {program_num:'', dist_id:'', trade:'D'},
};

// ── Config persistence (localStorage) ────────────────────────────────────────
const SAVE_KEY = 'monin_billback_cfg_v2';

function saveConfig() {
  try { localStorage.setItem(SAVE_KEY, JSON.stringify(getConfigRaw())); } catch(e){}
  showToast('Settings saved ✓');
}

function loadSavedConfig() {
  try {
    const saved = JSON.parse(localStorage.getItem(SAVE_KEY) || 'null');
    if (!saved) return;
    SUPPLIERS.forEach(s => {
      if (!saved[s]) return;
      const d = saved[s];
      const pe = document.getElementById(`cfg_${s}_prog`);
      const de = document.getElementById(`cfg_${s}_dist`);
      const te = document.getElementById(`cfg_${s}_trade`);
      if (pe && d.program_num !== undefined) pe.value = d.program_num;
      if (de && d.dist_id !== undefined)     de.value = d.dist_id;
      if (te && d.trade !== undefined)       te.value = d.trade;
    });
    // Restore Harbor item mappings
    if (saved.HARBOR && saved.HARBOR.item_map) {
      Object.entries(saved.HARBOR.item_map).forEach(([code, mcode]) => addHarborMapping(code, mcode));
    }
  } catch(e) {}
}

function showToast(msg) {
  let t = document.getElementById('toast');
  if (!t) {
    t = document.createElement('div');
    t.id = 'toast';
    t.style.cssText = 'position:fixed;bottom:24px;right:24px;background:#1a3a5c;color:#fff;padding:10px 18px;border-radius:6px;font-size:.85rem;z-index:999;opacity:0;transition:.3s';
    document.body.appendChild(t);
  }
  t.textContent = msg;
  t.style.opacity = '1';
  setTimeout(() => { t.style.opacity = '0'; }, 2000);
}

// ── Harbor item mapping ───────────────────────────────────────────────────────
let harborMappings = []; // [{code, mcode}]

function addHarborMapping(code='', mcode='') {
  harborMappings.push({code, mcode});
  renderHarborMappings();
}

function removeHarborMapping(idx) {
  harborMappings.splice(idx, 1);
  renderHarborMappings();
}

function renderHarborMappings() {
  const container = document.getElementById('harbor-mappings');
  if (!container) return;
  container.innerHTML = harborMappings.map((m, i) => `
    <div style="display:flex;gap:6px;margin-bottom:5px;align-items:center">
      <input value="${m.code}" onchange="harborMappings[${i}].code=this.value"
        placeholder="Harbor code (e.g. 505018)" style="flex:1;font-size:.8rem;border:1px solid #cbd5e1;border-radius:4px;padding:3px 7px">
      <span style="color:#94a3b8;font-size:.9rem">→</span>
      <input value="${m.mcode}" onchange="harborMappings[${i}].mcode=this.value"
        placeholder="M-code (e.g. M-FR109F)" style="flex:1;font-size:.8rem;border:1px solid #cbd5e1;border-radius:4px;padding:3px 7px">
      <button onclick="removeHarborMapping(${i})" style="background:none;border:none;cursor:pointer;color:#94a3b8;padding:2px 5px">✕</button>
    </div>`).join('');
}

// ── Config table ─────────────────────────────────────────────────────────────
const tbody = document.getElementById('config-tbody');
SUPPLIERS.forEach(s => {
  const d = DEFAULT_CFG[s] || {program_num:'',dist_id:'',trade:'D'};
  let extra = '';
  if (s === 'HARBOR') {
    extra = `<tr id="harbor-map-row" style="display:none"><td colspan="4" style="padding:8px 10px 12px">
      <div style="font-size:.8rem;font-weight:600;color:#475569;margin-bottom:6px">
        🔗 Item Code Mapping <span style="font-weight:400;color:#94a3b8">(map Harbor codes with no M-code to Monin M-codes)</span>
      </div>
      <div id="harbor-mappings"></div>
      <button onclick="addHarborMapping()" style="font-size:.78rem;color:#2e6da4;background:none;border:1px dashed #93c5fd;border-radius:4px;padding:3px 10px;cursor:pointer;margin-top:4px">+ Add mapping</button>
    </td></tr>`;
  }
  tbody.innerHTML += `<tr id="row_${s}">
    <td><strong>${s.replace(/_/g,' ')}</strong>${s==='HARBOR'?' <span style="font-size:.7rem;color:#2e6da4;cursor:pointer;text-decoration:underline" onclick="toggleHarborMap()">item codes</span>':''}</td>
    <td><input id="cfg_${s}_prog" value="${d.program_num}" placeholder="e.g. 1004089"></td>
    <td><input id="cfg_${s}_dist" value="${d.dist_id}" placeholder="e.g. 134810000"></td>
    <td><select id="cfg_${s}_trade"><option value="D" ${d.trade==='D'?'selected':''}>D</option><option value="O" ${d.trade==='O'?'selected':''}>O</option></select></td>
  </tr>${extra}`;
});

function toggleHarborMap() {
  const row = document.getElementById('harbor-map-row');
  row.style.display = row.style.display === 'none' ? '' : 'none';
  renderHarborMappings();
}

function toggleConfig() {
  const sec = document.getElementById('config-section');
  const tog = document.querySelector('.config-toggle');
  sec.classList.toggle('open');
  tog.textContent = sec.classList.contains('open') ? '▼ Hide Supplier Settings' : '▶ Show / Edit Supplier Settings';
}

function getConfigRaw() {
  const cfg = {};
  SUPPLIERS.forEach(s => {
    cfg[s] = {
      program_num: document.getElementById(`cfg_${s}_prog`).value.trim(),
      dist_id:     document.getElementById(`cfg_${s}_dist`).value.trim(),
      trade:       document.getElementById(`cfg_${s}_trade`).value,
    };
  });
  // Add Harbor item mappings
  if (harborMappings.length) {
    const map = {};
    harborMappings.forEach(m => { if (m.code && m.mcode) map[m.code.trim()] = m.mcode.trim(); });
    cfg['HARBOR'].item_map = map;
  }
  return cfg;
}

function getConfig() { return getConfigRaw(); }

// Load saved config on startup
window.addEventListener('DOMContentLoaded', loadSavedConfig);

let selectedFiles = [];

function detectSupplier(filename) {
  const fn = filename.toUpperCase();
  if (fn.includes('KAST'))    return 'KAST';
  if (fn.includes('SOFO'))    return 'SOFO';
  if (fn.includes('PFS'))     return 'PFS';
  if (fn.includes('BLAIR'))   return 'BLAIR_CANDY';
  if (fn.includes('LABATT'))  return 'LABATT';
  if (fn.includes('SHAMROCK')) return 'SHAMROCK';
  if (/S\s*AND\s*W|S\s*&\s*W/.test(fn)) return 'S_AND_W';
  if (fn.includes('BEK'))     return 'BEK';
  if (fn.includes('NICH'))    return 'NICH_CO';
  if (fn.includes('CBBB'))    return 'DOT_CBBB';
  if (/Y[\s.]?HATA|Y_HATA/.test(fn)) return 'Y_HATA';
  if (fn.includes('DRISCOLL')) return 'DRISCOLL';
  if (fn.includes('HARBOR') || fn.includes('SUPPLIER BILLBACK') || fn.includes('SUPPLIER_BILLBACK')) return 'HARBOR';
  return 'UNKNOWN';
}

// Full distributor list from Tellus (169 unique names)
const DIST_LIST = ["6 Degrees Grp","Accardi","Affiliated Foods","Alamode","Aloha","Atlas","Avalon FS","BEK","Bakemark","Balford","BiRite","Blair Candy","Brown Foodservice","CA Curtze","CASH-WA","CBCSpecialty","CCappuccino Connection","CORA","Cadillac","Cappuccino Connection","Carisam","ChefsWhse","Cheney","Cheney Bros","Cheung","Christ Panos","City Line","Clark","Clark Assoc","Country Club","Custom Food Svc","DDDI","DOT","Darden","Delco Foods","Dennis","DiCarlo","Dillanos","Dora's Natural","Dough Works","Driscoll Foods","EK Beverage","EVCO","EdwardDon","Espresso","FSA","Feesers","Fellers Foodservice","Ferraro","Fischer Foods","Fontana","Food Pro","Fred Hutch","GFS","Greco & Sons","HFM","HPC","HT Hackney","Harbor","Hardies","Harold Levinson","Henry's Foods","Hillcrest","Hillcrest Foodservice","Houstons","I Supply","IDF FD Serv Dist","IFD","Imperial Bag","Imperial Dade","Incredible","Indianhead FS","Individual FS","Iowa Des Moines","Iowa Des Moines Supply","Ital","JMC","JPoelp","Jacmar","Jakes","John Gross","John Gross & Co","Julius Silvert","KINEXO","Kaleel","Kast","Katsiroubas Bros.","Kohl Wholesale","Kuna","Labatt FS","Lineage","Lucky Goat","Lund","M&K","MAINES","MBM","Martin Bros","Martin Brower","Maximum Quality","McLane FS","McLane Grocery","Merchants","Merlino Foods","Nicholas&Co","Northern Lights","Odeko","PFD","PFG","PFS","Palmer","Paulines","Pippin","Pocono","Pon Food","Reinhart","Reliant","S&W","SF Supply","SHAHEEN BROS","SHAMROCK","SOFO","SOTF","SOTO","SPECS","SSA","SW Traders","SYSCO","SYSOCSYGMA","Saladino","Saval","Scavuzzos","Schenck","Schiffs","Sheetz","Shoreline","Six Degrees Grp","Smart&Final","Snowcap","Southeastern Food","Southern Star","Springfield Grocer","Sunbelt","Sunrise","Sutherland","Sweetwaters","Sygma","Sysco","Tankersley","Tapia","Testa","The Chefs Warehouse","Tidewater","Two Valleys","Two Valleys Dist","USCI","USF Culinary","USFS","Upper Lake","Van Eerden","Velmar","Vitco","WB Mason","Wasserstrom","West Coast Ship","White Castle","Willow Run","Win Depot","Y Hata","Youngs"];

// Map display name → internal parser key
function distNameToKey(name) {
  if (!name) return 'UNKNOWN';
  const u = name.toUpperCase().replace(/[^A-Z0-9]/g,'');
  if (u === 'KAST' || u.includes('KAST')) return 'KAST';
  if (u === 'SOFO') return 'SOFO';
  if (u === 'PFS' || u === 'PFSROMA') return 'PFS';
  if (u.includes('BLAIR')) return 'BLAIR_CANDY';
  if (u.includes('LABATT')) return 'LABATT';
  if (u.includes('SHAMROCK')) return 'SHAMROCK';
  if (u === 'SW' || u === 'SANDW' || u.includes('SANW')) return 'S_AND_W';
  if (u === 'BEK') return 'BEK';
  if (u.includes('NICHOL') || u === 'NICHCO') return 'NICH_CO';
  if (u.includes('CBBB')) return 'DOT_CBBB';
  if (u === 'YHATA' || u.includes('HATA')) return 'Y_HATA';
  if (u.includes('MCLANE')) return 'MCLANE';
  if (u.includes('MARTINBROS')) return 'MARTIN_BROS';
  if (u.includes('DRISCOLL')) return 'DRISCOLL';
  if (u.includes('HARBOR')) return 'HARBOR';
  // Everything else: route through content-based PDF dispatcher
  return 'HARBOR';
}

// Map parser key → display name for auto-detect
const KEY_TO_DISPLAY = {
  KAST:'Kast', SOFO:'SOFO', PFS:'PFS', LABATT:'Labatt FS',
  Y_HATA:'Y Hata', BEK:'BEK', NICH_CO:'Nicholas&Co', SHAMROCK:'SHAMROCK',
  DOT_CBBB:'DOT', MCLANE:'McLane FS', S_AND_W:'S&W',
  HARBOR:'Harbor', MARTIN_BROS:'Martin Bros',
  DOT_FOODS_BB:'DOT', DRISCOLL:'Driscoll Foods',
  BLAIR_CANDY:'Blair Candy', UNKNOWN:''
};

// Track currently open combobox index
let openComboIdx = null;

function supplierCombo(idx, detectedKey) {
  const displayName = KEY_TO_DISPLAY[detectedKey] || '';
  const isUnknown = !displayName || detectedKey === 'UNKNOWN';
  return `<div class="combo-wrap" id="combo_${idx}">
    <input type="text" class="combo-input${isUnknown?' unknown':''}" id="supplier_text_${idx}"
      value="${displayName.replace(/"/g,'&quot;')}"
      placeholder="Search distributor…"
      autocomplete="off"
      oninput="filterDist(${idx})"
      onkeydown="comboKey(event,${idx})"
      onfocus="showDist(${idx})">
    <span class="combo-arrow">▼</span>
    <input type="hidden" id="supplier_${idx}" value="${detectedKey}">
    <div class="combo-list" id="dist_list_${idx}"></div>
  </div>`;
}

function buildDistItems(idx, filter) {
  const q = (filter||'').toLowerCase();
  const matches = q ? DIST_LIST.filter(n => n.toLowerCase().includes(q)) : DIST_LIST;
  if (!matches.length) return `<div class="combo-item no-match">No match — will use generic PDF parser</div>`;
  return matches.map(n =>
    `<div class="combo-item" onmousedown="selectDist(${idx},'${n.replace(/\\/g,'\\\\').replace(/'/g,"\\'")}')">${n}</div>`
  ).join('');
}

function showDist(idx) {
  if (openComboIdx !== null && openComboIdx !== idx) hideDist(openComboIdx);
  openComboIdx = idx;
  const list = document.getElementById('dist_list_'+idx);
  const inp  = document.getElementById('supplier_text_'+idx);
  list.innerHTML = buildDistItems(idx, inp.value);
  list.style.display = 'block';
}

function hideDist(idx) {
  const list = document.getElementById('dist_list_'+idx);
  if (list) list.style.display = 'none';
  if (openComboIdx === idx) openComboIdx = null;
}

function filterDist(idx) {
  const inp  = document.getElementById('supplier_text_'+idx);
  const list = document.getElementById('dist_list_'+idx);
  list.innerHTML = buildDistItems(idx, inp.value);
  list.style.display = 'block';
  openComboIdx = idx;
  // Update hidden key as user types (in case they clear and type something new)
  if (!inp.value.trim()) {
    document.getElementById('supplier_'+idx).value = 'UNKNOWN';
    inp.className = 'combo-input unknown';
  }
}

function selectDist(idx, name) {
  const key  = distNameToKey(name);
  document.getElementById('supplier_text_'+idx).value = name;
  document.getElementById('supplier_'+idx).value = key;
  const inp  = document.getElementById('supplier_text_'+idx);
  inp.className = 'combo-input';
  hideDist(idx);
  // Update placeholder hints from config
  const cfg = getConfig()[key] || {};
  const pe = document.getElementById('prog_'+idx);
  const de = document.getElementById('dist_'+idx);
  if (pe) pe.placeholder = cfg.program_num ? `default: ${cfg.program_num}` : 'Program # (optional)';
  if (de) de.placeholder = cfg.dist_id     ? `default: ${cfg.dist_id}`     : 'Distributor ID (optional)';
}

function comboKey(e, idx) {
  const list = document.getElementById('dist_list_'+idx);
  const items = list.querySelectorAll('.combo-item:not(.no-match)');
  const focused = list.querySelector('.combo-item.focused');
  if (e.key === 'Escape') { hideDist(idx); return; }
  if (e.key === 'ArrowDown') {
    e.preventDefault();
    const next = focused ? focused.nextElementSibling : items[0];
    if (focused) focused.classList.remove('focused');
    if (next && !next.classList.contains('no-match')) next.classList.add('focused');
    if (next) next.scrollIntoView({block:'nearest'});
    return;
  }
  if (e.key === 'ArrowUp') {
    e.preventDefault();
    const prev = focused ? focused.previousElementSibling : items[items.length-1];
    if (focused) focused.classList.remove('focused');
    if (prev && !prev.classList.contains('no-match')) prev.classList.add('focused');
    if (prev) prev.scrollIntoView({block:'nearest'});
    return;
  }
  if (e.key === 'Enter' && focused) {
    e.preventDefault();
    focused.onmousedown();
  }
}

// Close combo when clicking elsewhere
document.addEventListener('click', function(e) {
  if (openComboIdx !== null) {
    const wrap = document.getElementById('combo_'+openComboIdx);
    if (wrap && !wrap.contains(e.target)) hideDist(openComboIdx);
  }
});

function onSupplierChange(idx) {
  // kept for backward compat — actual updates now in selectDist
}

function renderFiles() {
  const list = document.getElementById('file-list');
  if (!selectedFiles.length) { list.innerHTML=''; return; }
  list.innerHTML = selectedFiles.map((f,i) => {
    const sup = detectSupplier(f.name);
    const cfg = getConfig()[sup] || {};
    const progPlaceholder = cfg.program_num ? `default: ${cfg.program_num}` : 'Program # (optional)';
    const distPlaceholder = cfg.dist_id     ? `default: ${cfg.dist_id}`     : 'Distributor ID (optional)';
    const unknownStyle = sup === 'UNKNOWN' ? 'background:#fef9c3;' : '';
    return `<div class="file-item">
      <div class="file-item-top">
        ${supplierCombo(i, sup)}
        <span class="file-name" title="${f.name}">${f.name}</span>
        <button class="remove-btn" onclick="removeFile(${i})" title="Remove">✕</button>
      </div>
      <div class="file-item-fields">
        <label title="Override Program # for this file only">
          Prog #
          <input class="override-input" id="prog_${i}" placeholder="${progPlaceholder}">
        </label>
        <label title="Override Distributor ID for this file only">
          Dist ID
          <input class="override-input" id="dist_${i}" placeholder="${distPlaceholder}">
        </label>
        <label title="Customer Reference — leave blank to auto-detect from file">
          Customer Ref
          <input class="cref-input" id="cref_${i}" placeholder="auto-detect from file">
        </label>
      </div>
    </div>`;
  }).join('');
  document.getElementById('process-btn').disabled = false;
}

function removeFile(idx) {
  selectedFiles.splice(idx, 1);
  renderFiles();
  if (!selectedFiles.length) document.getElementById('process-btn').disabled = true;
}

function clearAll() {
  selectedFiles = [];
  renderFiles();
  document.getElementById('results').style.display = 'none';
  document.getElementById('progress').style.display = 'none';
}

document.getElementById('file-input').onchange = function(e) {
  Array.from(e.target.files).forEach(f => selectedFiles.push(f));
  renderFiles();
  this.value = '';
};

const dz = document.getElementById('drop-zone');
dz.ondragover = e => { e.preventDefault(); dz.classList.add('drag-over'); };
dz.ondragleave = () => dz.classList.remove('drag-over');
dz.ondrop = e => {
  e.preventDefault(); dz.classList.remove('drag-over');
  Array.from(e.dataTransfer.files).forEach(f => selectedFiles.push(f));
  renderFiles();
};

async function processFiles() {
  if (!selectedFiles.length) return;
  const btn = document.getElementById('process-btn');
  btn.disabled = true;
  const prog = document.getElementById('progress');
  const fill = document.getElementById('progress-fill');
  const ptxt = document.getElementById('progress-text');
  const results = document.getElementById('results');
  prog.style.display = 'block';
  results.style.display = 'none';
  results.innerHTML = '';

  const config = getConfig();
  const fileOverrides = {};
  selectedFiles.forEach((f,i) => {
    fileOverrides[f.name] = {
      supplier:    document.getElementById(`supplier_${i}`)?.value || '',
      program_num: document.getElementById(`prog_${i}`)?.value.trim() || '',
      dist_id:     document.getElementById(`dist_${i}`)?.value.trim() || '',
      customer_ref:document.getElementById(`cref_${i}`)?.value.trim() || '',
    };
  });

  const form = new FormData();
  selectedFiles.forEach(f => form.append('files', f));
  form.append('config', JSON.stringify(config));
  form.append('file_overrides', JSON.stringify(fileOverrides));

  fill.style.width = '30%';
  ptxt.textContent = 'Uploading files...';

  try {
    const resp = await fetch('/process', {method:'POST', body: form});
    fill.style.width = '70%';
    ptxt.textContent = 'Generating Tellus file...';
    const data = await resp.json();
    fill.style.width = '100%';

    let html = '';
    let totalRows = 0;
    let totalAmount = 0;
    let totalQty = 0;
    let hasErrors = false;

    data.results.forEach(r => {
      if (r.error) {
        hasErrors = true;
        html += `<div class="result-row result-err">
          <span>❌ <strong>${r.file}</strong> — ${r.error}</span>
        </div>`;
      } else if (r.rows === 0) {
        html += `<div class="result-row result-skip">
          <span>⚠️ <strong>${r.file}</strong> <em>(${r.supplier})</em> — No Monin items found</span>
          <span class="count-badge" style="background:#fef3c7;color:#92400e">0 rows</span>
        </div>`;
      } else {
        totalRows += r.rows;
        totalAmount += r.total_amount || 0;
        totalQty += r.total_qty || 0;
        const hasWarns = r.warnings && r.warnings.length > 0;
        const warnHtml = hasWarns
          ? `<div style="margin-top:8px;padding:8px 10px;background:#fffbeb;border-radius:5px;border-left:3px solid #f59e0b">
               <div style="font-size:.78rem;font-weight:600;color:#92400e;margin-bottom:4px">
                 ⚠️ ${r.warnings.length} item${r.warnings.length>1?'s':''} missing M-code — $${r.warn_total.toFixed(2)} not included in this file
               </div>
               <div style="font-size:.75rem;color:#b45309;line-height:1.6">
                 ${r.warnings.map(w => `• ${w}`).join('<br>')}
               </div>
               <div style="font-size:.75rem;color:#92400e;margin-top:5px;font-style:italic">
                 These items need to be entered manually in Tellus. Ask your supplier to include the Monin M-code in future files, or add the mapping in Supplier Settings above.
               </div>
             </div>` : '';
        const amtStr = (r.total_amount || 0).toLocaleString('en-US',{style:'currency',currency:'USD'});
        const qtyStr = (r.total_qty || 0).toLocaleString();
        html += `<div class="result-row result-ok" style="flex-direction:column;align-items:flex-start">
          <div style="display:flex;justify-content:space-between;width:100%;flex-wrap:wrap;gap:4px">
            <span>✅ <strong>${r.file}</strong> <em>(${r.supplier})</em></span>
            <span style="display:flex;gap:6px;align-items:center;flex-wrap:wrap">
              <span class="count-badge">${r.rows} rows</span>
              <span class="count-badge" style="background:#dcfce7;color:#166534">Qty: ${qtyStr}</span>
              <span class="count-badge" style="background:#eff6ff;color:#1e40af">${amtStr}</span>
              ${hasWarns?`<span class="count-badge" style="background:#fef3c7;color:#92400e">⚠️ ${r.warnings.length} manual</span>`:''}
            </span>
          </div>${warnHtml}
        </div>`;
      }
    });

    const grandAmt = totalAmount.toLocaleString('en-US',{style:'currency',currency:'USD'});
    html = `<div class="total-bar">📋 ${totalRows} rows · ${totalQty.toLocaleString()} cases · ${grandAmt} total</div>` + html;

    if (data.download_id) {
      html += `<div class="dl-box">
        <p>✅ Tellus upload file is ready!</p>
        <a href="/download/${data.download_id}" class="btn btn-primary" style="text-decoration:none;display:inline-flex">
          ⬇️ Download Tellus_Upload_${new Date().toLocaleDateString('en-CA').replace(/-/g,'')}.xlsx
        </a>
      </div>`;
    }

    results.innerHTML = html;
    results.style.display = 'block';
  } catch(err) {
    results.innerHTML = `<div class="result-row result-err">❌ Server error: ${err.message}</div>`;
    results.style.display = 'block';
  }
  prog.style.display = 'none';
  btn.disabled = false;
}
</script>
</body>
</html>
"""

# ─── HTTP SERVER ──────────────────────────────────────────────────────────────
_downloads = {}   # id → bytes


# ── Simple session-based auth (only active when APP_PASSWORD env var is set) ──
import hashlib, secrets as _secrets
_active_sessions = set()  # set of valid session tokens

APP_PASSWORD = os.environ.get('APP_PASSWORD', '')  # set this on Render

LOGIN_PAGE = """<!DOCTYPE html>
<html><head><meta charset="utf-8">
<title>Monin Billback — Login</title>
<style>
  body{margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;
       background:#f0f4ff;font-family:'Segoe UI',sans-serif}
  .card{background:#fff;border-radius:16px;padding:48px 40px;box-shadow:0 4px 24px rgba(0,0,0,.1);
        text-align:center;width:320px}
  h2{margin:0 0 8px;color:#1e3a5f;font-size:1.4rem}
  p{color:#64748b;margin:0 0 28px;font-size:.9rem}
  input{width:100%;box-sizing:border-box;padding:10px 14px;border:1px solid #cbd5e1;
        border-radius:8px;font-size:1rem;margin-bottom:14px;outline:none}
  input:focus{border-color:#3b82f6}
  button{width:100%;padding:11px;background:#1d4ed8;color:#fff;border:none;border-radius:8px;
         font-size:1rem;font-weight:600;cursor:pointer}
  button:hover{background:#1e40af}
  .err{color:#dc2626;font-size:.85rem;margin-bottom:12px}
</style></head>
<body><div class="card">
  <img src="data:image/png;base64,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LMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMBCPAEAAACwEE8AAAAALMQTAAAAAAvxBAAAAMAiEzn2uAfq614AAAAASUVORK5CYII=" alt="Monin" style="height:64px;object-fit:contain;margin-bottom:12px">
  <h2>Billback Processor</h2>
  <p>Enter the team password to continue.</p>
  {error}
  <form method="post" action="/login">
    <input type="password" name="pw" placeholder="Password" autofocus>
    <button type="submit">Sign In</button>
  </form>
</div></body></html>"""

def _get_session(headers):
    """Return session token from Cookie header, or None."""
    cookie = headers.get('Cookie', '')
    for part in cookie.split(';'):
        k, _, v = part.strip().partition('=')
        if k.strip() == 'bb_session':
            return v.strip()
    return None

def _is_authed(headers):
    """Return True if auth is disabled or valid session cookie present."""
    if not APP_PASSWORD:
        return True  # running locally without password
    tok = _get_session(headers)
    return tok in _active_sessions


class BillbackHandler(BaseHTTPRequestHandler):
    """Clean multipart handler using email library for proper filename extraction."""
    def log_message(self, format, *args): pass

    def do_GET(self):
        # ── Auth check ──
        if not _is_authed(self.headers):
            self._serve_login()
            return

        if self.path == '/':
            self.send_response(200)
            self.send_header('Content-Type','text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(HTML_PAGE.encode())
            return
        if self.path.startswith('/download/'):
            did = self.path.split('/')[-1]
            data = _downloads.get(did)
            if data:
                self.send_response(200)
                self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                fname = f'Tellus_Upload_{TODAY}.xlsx'
                self.send_header('Content-Disposition',f'attachment; filename="{fname}"')
                self.send_header('Content-Length', str(len(data)))
                self.end_headers()
                self.wfile.write(data)
                return
        self.send_response(404); self.end_headers()

    def _serve_login(self, error=''):
        err_html = f'<p class="err">{error}</p>' if error else ''
        page = LOGIN_PAGE.replace('{error}', err_html).encode()
        self.send_response(200)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.end_headers()
        self.wfile.write(page)

    def do_POST(self):
        # ── Login form submission ──
        if self.path == '/login':
            length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(length).decode()
            from urllib.parse import parse_qs
            params = parse_qs(body)
            pw = params.get('pw', [''])[0]
            if APP_PASSWORD and pw == APP_PASSWORD:
                tok = _secrets.token_hex(24)
                _active_sessions.add(tok)
                self.send_response(302)
                self.send_header('Location', '/')
                self.send_header('Set-Cookie', f'bb_session={tok}; Path=/; HttpOnly; SameSite=Lax')
                self.end_headers()
            else:
                self._serve_login(error='Incorrect password — please try again.')
            return

        if not _is_authed(self.headers):
            self.send_response(403); self.end_headers(); return

        if self.path != '/process':
            self.send_response(404); self.end_headers(); return
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            content_type = self.headers.get('Content-Type', '')

            # Parse multipart using email package
            msg_bytes = f'Content-Type: {content_type}\r\n\r\n'.encode() + body
            msg = email.message_from_bytes(msg_bytes)

            user_config = {}
            file_overrides = {}   # {filename: {program_num, dist_id, customer_ref}}
            file_parts = []

            for part in msg.walk():
                cd = part.get('Content-Disposition', '')
                if not cd: continue
                name_m = re.search(r'name="([^"]+)"', cd)
                fname_m = re.search(r'filename="([^"]+)"', cd)
                field_name = name_m.group(1) if name_m else ''
                filename   = fname_m.group(1) if fname_m else ''
                payload = part.get_payload(decode=True)
                if payload is None: continue

                if field_name == 'config':
                    try: user_config = json.loads(payload.decode())
                    except: pass
                elif field_name in ('file_overrides', 'customer_refs'):
                    # Support both new file_overrides and legacy customer_refs field
                    try:
                        parsed = json.loads(payload.decode())
                        if field_name == 'file_overrides':
                            file_overrides = parsed
                        else:
                            # Legacy: customer_refs was {filename: ref_str}
                            for fn, ref in parsed.items():
                                file_overrides.setdefault(fn, {})['customer_ref'] = ref
                    except: pass
                elif field_name == 'files' and filename:
                    file_parts.append((filename, payload))

            tmpdir = tempfile.mkdtemp()
            results = []
            all_rows = []

            for filename, content in file_parts:
                # Save with original filename for supplier detection
                safe_name = re.sub(r'[^\w.\-]', '_', filename)
                tmp_path = os.path.join(tmpdir, safe_name)
                with open(tmp_path, 'wb') as f:
                    f.write(content)

                fo   = file_overrides.get(filename, {})
                cref = fo.get('customer_ref', '')
                try:
                    supplier, rows = detect_and_parse(tmp_path, user_config, cref, file_override=fo, original_filename=filename)
                    fatal_errs = [r['_error'] for r in rows if '_error' in r]
                    warnings   = [r for r in rows if '_warning' in r]
                    ok         = [r for r in rows if '_error' not in r and '_warning' not in r]
                    all_rows.extend(ok)
                    warn_total = round(sum(w['amount'] for w in warnings), 4)
                    warn_list  = [f"{w['code']} — {w['desc']} (${w['amount']})" for w in warnings]
                    if fatal_errs and not ok:
                        results.append({'file': filename, 'supplier': supplier,
                                        'error': fatal_errs[0], 'rows': 0,
                                        'warnings': [], 'warn_total': 0,
                                        'total_amount': 0.0, 'total_qty': 0})
                    else:
                        total_amount = round(sum(
                            float(r.get('Item Dollar Amount') or 0) for r in ok), 2)
                        total_qty = int(round(sum(
                            float(r.get('Item Volume Qty') or 0) for r in ok)))
                        results.append({'file': filename, 'supplier': supplier,
                                        'rows': len(ok), 'error': None,
                                        'warnings': warn_list, 'warn_total': warn_total,
                                        'total_amount': total_amount, 'total_qty': total_qty})
                except Exception as ex:
                    results.append({'file': filename, 'supplier': 'ERROR',
                                    'error': str(ex), 'rows': 0,
                                    'warnings': [], 'warn_total': 0})

            shutil.rmtree(tmpdir, ignore_errors=True)

            download_id = None
            if all_rows:
                xlsx_bytes = build_output(all_rows)
                download_id = str(uuid.uuid4())[:8]
                _downloads[download_id] = xlsx_bytes

            resp = {'results': results, 'download_id': download_id,
                    'total_rows': len(all_rows)}
            self._respond_json(resp)

        except Exception as e:
            self._respond_json({'error': str(e), 'trace': traceback.format_exc(),
                                'results': [], 'download_id': None})

    def _respond_json(self, data):
        body = json.dumps(data).encode()
        self.send_response(200)
        self.send_header('Content-Type','application/json')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

def run():
    # When deployed online (Render, Railway, etc.) use the PORT env var.
    # When running locally as .exe, fall back to 8765-8775.
    env_port = os.environ.get('PORT')
    is_web = bool(env_port)

    if is_web:
        port = int(env_port)
        host = '0.0.0.0'
        server = HTTPServer((host, port), BillbackHandler)
        print(f'Monin Billback running on port {port}')
        try:
            server.serve_forever()
        except KeyboardInterrupt:
            pass
    else:
        port = 8765
        host = 'localhost'
        for p in range(8765, 8776):
            try:
                server = HTTPServer((host, p), BillbackHandler)
                port = p
                break
            except OSError:
                continue
        else:
            print('ERROR: Could not find a free port (8765-8775). Close other apps and try again.')
            input('Press Enter to exit.')
            return
        url = f'http://{host}:{port}'
        threading.Timer(1.5, lambda: webbrowser.open(url)).start()
        try:
            server.serve_forever()
        except KeyboardInterrupt:
            pass

if __name__ == '__main__':
    run()
