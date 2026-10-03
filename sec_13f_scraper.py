"""
SEC 13F Filing Scraper
Fetches all 13F-HR filings for a fund from EDGAR and outputs an Excel workbook.

Usage:
    python sec_13f_scraper.py --cik 0001649339 --output aqr_13f.xlsx
    python sec_13f_scraper.py --name "Pershing Square" --output pershing_13f.xlsx

Columns: Quarter | Ticker | CUSIP | Security Name | Shares | Market Value ($) | % of Fund | QoQ Delta MV ($)
"""
import argparse
import time
import re
import sys
from io import StringIO
import xml.etree.ElementTree as ET
import requests
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.styles.numbers import FORMAT_NUMBER_COMMA_SEPARATED1

EDGAR_BASE = 'https://data.sec.gov'
EFTS_BASE = 'https://efts.sec.gov'
EDGAR_BROWSE = 'https://www.sec.gov'

HEADERS = {
    'User-Agent': '13F-Scraper research@example.com',
    'Accept-Encoding': 'gzip, deflate',
}

NS_PATTERN = re.compile(r'\s+xmlns[^"]*"[^"]*"')

# The SEC required the <value> field in the 13F information table to be reported
# "in thousands of dollars" up through year-end 2022. Starting with filings filed
# on or after 2023-01-03 (SEC Release No. 34-95148 / amendments to Form 13F),
# filers report the actual whole-dollar value instead. Filings before that date
# still need the old *1000 conversion; filings on/after it must NOT be multiplied.
VALUE_IN_THOUSANDS_CUTOFF = '2023-01-03'


def _value_already_in_dollars(filing_date):
    """True if this filing's <value> field is already whole dollars (post 2023-01-03)."""
    if not filing_date:
        # Unknown filing date: assume modern format (whole dollars) since nearly
        # all filings encountered today are post-cutoff.
        return True
    return filing_date >= VALUE_IN_THOUSANDS_CUTOFF


def search_cik_by_name(name):
    """Return (cik, entity_name) for the best match on EDGAR full-text search."""
    url = f'{EDGAR_BASE}/submissions/search.json?company={requests.utils.quote(name)}&type=13F-HR&dateb=&owner=include&count=10&search_text='

    url = f'{EDGAR_BROWSE}/cgi-bin/browse-edgar?company={requests.utils.quote(name)}&CIK=&type=13F-HR&dateb=&owner=include&count=10&search_text=&action=getcompany&output=atom'

    r = requests.get(url, headers=HEADERS, timeout=30)
    r.raise_for_status()

    root = ET.fromstring(r.text)
    ns = {'atom': 'http://www.w3.org/2005/Atom'}

    entries = root.findall('atom:entry', ns)
    if not entries:
        sys.exit(f"No 13F filers found matching '{name}'. Try --cik directly.")

    cik_el = entries[0].find('atom:id', ns)
    name_el = entries[0].find('atom:company-name', ns)

    cik_match = re.search('CIK=(\\d+)', cik_el.text if cik_el is not None else '')
    if not cik_match:
        sys.exit('Could not parse CIK from search results.')

    cik = cik_match.group(1).lstrip('0')
    entity = name_el.text if name_el is not None else name
    return cik, entity


def get_submissions(cik):
    """Fetch the submissions JSON for a CIK (pads to 10 digits)."""
    padded = cik.zfill(10)
    url = f'{EDGAR_BASE}/submissions/CIK{padded}.json'
    r = requests.get(url, headers=HEADERS, timeout=30)
    r.raise_for_status()
    return r.json()


def get_all_13f_filings(cik, dedupe=True):
    """Return list of dicts with keys: accessionNumber, filingDate, primaryDocument."""
    data = get_submissions(cik)
    entity_name = data.get('name', cik)
    print(f'Entity: {entity_name}  |  CIK: {cik}')

    filings = []
    recent = data.get('filings', {}).get('recent', {})
    forms = recent.get('form', [])
    accessions = recent.get('accessionNumber', [])
    dates = recent.get('filingDate', [])
    docs = recent.get('primaryDocument', [])

    for form, acc, date, doc in zip(forms, accessions, dates, docs):
        if form not in ('13F-HR', '13F-HR/A'):
            continue
        filings.append({'accessionNumber': acc, 'filingDate': date, 'primaryDocument': doc, 'form': form})

    for extra_file in data.get('filings', {}).get('files', []):
        extra_url = f"{EDGAR_BASE}/submissions/{extra_file['name']}"
        r = requests.get(extra_url, headers=HEADERS, timeout=30)
        r.raise_for_status()
        extra = r.json()
        ex_forms = extra.get('form', [])
        ex_acc = extra.get('accessionNumber', [])
        ex_dates = extra.get('filingDate', [])
        ex_docs = extra.get('primaryDocument', [])
        for form, acc, date, doc in zip(ex_forms, ex_acc, ex_dates, ex_docs):
            if form not in ('13F-HR', '13F-HR/A'):
                continue
            filings.append({'accessionNumber': acc, 'filingDate': date, 'primaryDocument': doc, 'form': form})

    filings.sort(key=lambda x: x['filingDate'])

    if dedupe:
        filings = dedupe_by_quarter(filings)

    print(f'Found {len(filings)} 13F-HR filings.')

    return filings, entity_name


def dedupe_by_quarter(filings):
    """Keep the latest amendment per calendar quarter."""
    by_quarter = {}
    for f in filings:
        q = date_to_quarter(f['filingDate'])
        by_quarter[q] = f
    return [by_quarter[q] for q in sorted(by_quarter)]


def date_to_quarter(date_str):
    """'2023-05-15' -> '2023-Q1' (based on calendar quarter of filing month)."""
    dt = pd.to_datetime(date_str)
    return f'{dt.year}-Q{dt.quarter}'


def quarter_label(date_str):
    """Return a display label like 'Q1 2023' for the period covered by a filing date.

    13F filings are due 45 days after quarter end, so a May filing covers Q1.
    """
    dt = pd.to_datetime(date_str)

    month = dt.month
    year = dt.year

    if month <= 2:
        return f'Q4 {year - 1}'
    if month <= 5:
        return f'Q1 {year}'
    if month <= 8:
        return f'Q2 {year}'
    if month <= 11:
        return f'Q3 {year}'

    return f'Q4 {year}'


def _get_with_retry(url, retries=5):
    """GET with exponential backoff on 429 / 5xx."""
    delay = 2.0
    for attempt in range(retries):
        r = requests.get(url, headers=HEADERS, timeout=30)
        if r.status_code == 429 or r.status_code >= 500:
            wait = delay * 2 ** attempt
            print(f'    Rate-limited ({r.status_code}), retrying in {wait:.0f}s...')
            time.sleep(wait)
            continue
        r.raise_for_status()
        return r

    r.raise_for_status()
    return r


def find_info_table_url(cik, accession):
    """Find the XML information table document URL within a 13F filing."""
    padded = cik.zfill(10)
    acc_nodash = accession.replace('-', '')
    base = f'https://www.sec.gov/Archives/edgar/data/{padded}/{acc_nodash}'

    index_url = f'{base}/{accession}-index.htm'

    try:
        r = _get_with_retry(index_url)

        xml_links = re.findall('href="([^"]+\\.xml)"', r.text, re.IGNORECASE)
        xml_links = [l for l in xml_links if 'xsl' not in l.lower()]

        for link in xml_links:
            name = link.lower().split('/')[-1]
            if 'infotable' in name:
                full = link if link.startswith('http') else f'https://www.sec.gov{link}'
                return full

        for link in xml_links:
            name = link.lower().split('/')[-1]
            if 'form13f' in name:
                continue
            if 'primary' in name:
                continue
            full = link if link.startswith('http') else f'https://www.sec.gov{link}'
            return full

        if xml_links:
            link = xml_links[0]
            return link if link.startswith('http') else f'https://www.sec.gov{link}'
    except Exception as e:
        print(f'    Index fetch error: {e}')

    return None


def parse_info_table(xml_url, filing_date=None):
    """Download and parse a 13F information table XML, return list of holdings.

    filing_date (YYYY-MM-DD, from the filing index) determines whether the
    reported <value> is in thousands of dollars (pre-2023 filings) or whole
    dollars (2023-01-03 onward) — see VALUE_IN_THOUSANDS_CUTOFF above.
    """
    r = _get_with_retry(xml_url)
    r.raise_for_status()
    try:
        root = ET.fromstring(r.content)
    except ET.ParseError as e:
        print(f'    XML parse error: {e}')
        print(f'    URL was: {xml_url}')
        print(f'    Content start: {r.text[:300]!r}')
        return []

    ns_match = re.match(r'\{([^}]+)\}', root.tag)
    ns = ns_match.group(1) if ns_match else ''

    def tag(t):
        if ns:
            return f'{{{ns}}}{t}'
        return t

    entries = list(root.iter(tag('infoTable')))
    if not entries:
        entries = list(root.iter('infoTable'))

    holdings = []
    for entry in entries:
        name = _text(entry, tag('nameOfIssuer')) or _text(entry, 'nameOfIssuer')
        cusip = _text(entry, tag('cusip')) or _text(entry, 'cusip')
        value = _int(entry, tag('value')) or _int(entry, 'value')
        ssh = entry.find(tag('shrsOrPrnAmt'))
        if ssh is None:
            ssh = entry.find('shrsOrPrnAmt')
        shares = 0
        if ssh is not None:
            amt = ssh.find(tag('sshPrnamt'))
            if amt is None:
                amt = ssh.find('sshPrnamt')
            if amt is not None and amt.text:
                shares = int(amt.text.replace(',', '').strip())
        ticker = _text(entry, tag('ticker')) or _text(entry, 'ticker')
        market_value = value if _value_already_in_dollars(filing_date) else value * 1000
        holdings.append({
            'security_name': name,
            'cusip': cusip,
            'ticker': ticker,
            'shares': shares,
            'market_value': market_value,
        })
    return holdings


def _text(el, tag):
    child = el.find(tag)
    if child is not None and child.text:
        return child.text.strip()
    return ''


def _int(el, tag):
    child = el.find(tag)
    if child is None or not child.text:
        return 0
    try:
        return int(child.text.replace(',', '').strip())
    except ValueError:
        return 0


_CUSIP_TO_TICKER = {}


def load_sec_ticker_map():
    """Load the SEC's company_tickers_exchange.json for CUSIP->ticker mapping."""
    return None


def build_dataframe(all_quarters):
    """Combine all quarterly holdings into a single DataFrame and compute derived columns."""
    rows = []
    for quarter, holdings in all_quarters:
        total_mv = sum(h['market_value'] for h in holdings)
        for h in holdings:
            pct = h['market_value'] / total_mv * 100 if total_mv else 0
            rows.append({
                'Quarter': quarter,
                'Ticker': h['ticker'],
                'CUSIP': h['cusip'],
                'Security Name': h['security_name'],
                'Shares': h['shares'],
                'Market Value ($)': h['market_value'],
                '% of Fund': round(pct, 4),
            })

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    def quarter_sort_key(q):
        num, year = q.split(' ')
        return (int(year), int(num[1]))

    df['_qkey'] = df['Quarter'].map(quarter_sort_key)

    df = df.sort_values(['CUSIP', '_qkey'])

    df['QoQ Delta MV ($)'] = df.groupby('CUSIP')['Market Value ($)'].diff()

    df = df.sort_values(['_qkey', 'Market Value ($)'], ascending=[True, False])

    df = df.drop(columns=['_qkey'])

    return df.reset_index(drop=True)


def build_common_holdings(funds):
    """
    For each (year, CUSIP) held by any fund, produce one row with:
      Year | CUSIP | Security Name | <FundA>_held | <FundA>_weight | ...
    Weight = average % of fund across quarters in that year (0 if not held).
    """
    fund_year = {}

    for entity_name, df in funds:
        d = df.copy()
        d['Year'] = d['Quarter'].str.split(' ').str[1]

        by_year = {}
        for year, ydf in d.groupby('Year'):
            cusip_data = {}
            for cusip, cdf in ydf.groupby('CUSIP'):
                cusip_data[cusip] = {
                    'name': cdf['Security Name'].iloc[0],
                    'avg_weight': cdf['% of Fund'].mean(),
                }
            by_year[year] = cusip_data
        fund_year[entity_name] = by_year

    all_years = sorted({y for fh in fund_year.values() for y in fh})

    fund_names = [name for name, _ in funds]

    rows = []
    for year in all_years:
        all_cusips = {}
        for fname in fund_names:
            for cusip, info in fund_year.get(fname, {}).get(year, {}).items():
                if cusip not in all_cusips:
                    all_cusips[cusip] = info['name']

        for cusip, sec_name in all_cusips.items():
            row = {'Year': year, 'CUSIP': cusip, 'Security Name': sec_name}
            for fname in fund_names:
                info = fund_year.get(fname, {}).get(year, {}).get(cusip)
                row[f'{fname}_held'] = 1 if info else 0
                row[f'{fname}_weight'] = round(info['avg_weight'], 4) if info else 0.0
            rows.append(row)

    return pd.DataFrame(rows)


def _add_common_holdings_sheet(wb, funds):
    from openpyxl.worksheet.datavalidation import DataValidation

    df = build_common_holdings(funds)
    if df.empty:
        return None

    ws = wb.create_sheet('Common Holdings')

    fund_names = [name for name, _ in funds]
    n_funds = len(fund_names)

    title_font = Font(name='Calibri', bold=True, size=14, color='1F3864')
    hdr_font = Font(name='Calibri', bold=True, size=10, color='FFFFFF')
    hdr_fill = PatternFill('solid', fgColor='1F3864')
    yes_fill = PatternFill('solid', fgColor='C6EFCE')
    no_fill = PatternFill('solid', fgColor='FFCCCC')
    label_font = Font(name='Calibri', bold=True, size=10)
    body_font = Font(name='Calibri', size=10)
    section_font = Font(name='Calibri', bold=True, size=11, color='1F3864')

    ws['A1'] = 'Common Holdings Analysis'
    ws['A1'].font = title_font

    ws['A3'] = 'Fund'
    ws['B3'] = 'Include in Comparison?'
    for cell in (ws['A3'], ws['B3']):
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal='center')

    selector_start = 4
    selector_end = selector_start + n_funds - 1

    dv = DataValidation(type='list', formula1='"YES,NO"', allow_blank=False)
    ws.add_data_validation(dv)

    for i, fname in enumerate(fund_names):
        row = selector_start + i
        ws.cell(row=row, column=1, value=fname).font = label_font
        cell = ws.cell(row=row, column=2, value='YES')
        cell.font = label_font
        cell.fill = yes_fill
        cell.alignment = Alignment(horizontal='center')
        dv.add(cell)

    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 24

    tbl_hdr_row = selector_end + 2
    tbl_data_row = tbl_hdr_row + 1

    fixed_cols = ['Year', 'CUSIP', 'Security Name', '# Selected Funds']

    fund_col_pairs = [(f'{fn} Held', f'{fn} Weight %') for fn in fund_names]

    all_col_headers = fixed_cols + [c for pair in fund_col_pairs for c in pair]

    for col_idx, col_name in enumerate(all_col_headers, 1):
        cell = ws.cell(row=tbl_hdr_row, column=col_idx, value=col_name)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal='center')

    col_widths = {'Year': 8, 'CUSIP': 12, 'Security Name': 36, '# Selected Funds': 18}

    for col_idx, col_name in enumerate(all_col_headers, 1):
        letter = get_column_letter(col_idx)
        w = col_widths.get(col_name)
        if w:
            ws.column_dimensions[letter].width = w
            continue
        if 'Held' in col_name:
            ws.column_dimensions[letter].width = 8
            continue
        if 'Weight' not in col_name:
            continue
        ws.column_dimensions[letter].width = max(14, min(len(col_name) + 2, 28))

    flag_col_start = len(fixed_cols) + 1

    held_col_letters = [get_column_letter(flag_col_start + i * 2) for i in range(n_funds)]

    df['_total'] = df[[f'{fn}_held' for fn in fund_names]].sum(axis=1)

    df = df.sort_values(['Year', '_total'], ascending=[True, False]).drop(columns=['_total'])

    for row_offset, (_, data_row) in enumerate(df.iterrows()):
        r = tbl_data_row + row_offset
        fill = ALT_FILL if row_offset % 2 == 0 else None

        ws.cell(row=r, column=1, value=data_row['Year']).font = body_font
        ws.cell(row=r, column=2, value=data_row['CUSIP']).font = body_font
        ws.cell(row=r, column=3, value=data_row['Security Name']).font = body_font

        formula = '=' + '+'.join(
            f'($B${selector_start + i}="YES")*{l}{r}'
            for i, l in enumerate(held_col_letters)
        )

        count_cell = ws.cell(row=r, column=4, value=formula)
        count_cell.font = body_font
        count_cell.alignment = Alignment(horizontal='center')

        col_cursor = flag_col_start
        for fname in fund_names:
            held = int(data_row[f'{fname}_held'])
            weight = data_row[f'{fname}_weight']

            held_cell = ws.cell(row=r, column=col_cursor, value=held)
            held_cell.font = body_font
            held_cell.alignment = Alignment(horizontal='center')
            if held:
                held_cell.fill = PatternFill('solid', fgColor='C6EFCE')

            w_cell = ws.cell(row=r, column=col_cursor + 1, value=weight / 100 if held else None)
            w_cell.font = body_font
            w_cell.number_format = '0.00%'
            w_cell.alignment = Alignment(horizontal='right')
            if held:
                w_cell.fill = PatternFill('solid', fgColor='C6EFCE')

            col_cursor += 2

        if fill:
            for col_idx in range(1, 4):
                ws.cell(row=r, column=col_idx).fill = fill

    ws.auto_filter.ref = f'A{tbl_hdr_row}:{get_column_letter(len(all_col_headers))}{tbl_data_row + len(df) - 1}'

    ws.freeze_panes = f'A{tbl_data_row}'

    ws['A2'] = 'Toggle YES/NO in column B to include/exclude funds from the # Selected Funds count.'
    ws['A2'].font = Font(name='Calibri', italic=True, size=9, color='666666')


HEADER_FILL = PatternFill('solid', fgColor='1F3864')
ALT_FILL = PatternFill('solid', fgColor='EFF3FB')
HEADER_FONT = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
BODY_FONT = Font(name='Calibri', size=10)
BORDER_SIDE = Side(style='thin', color='C0C0C0')
THIN_BORDER = Border(bottom=BORDER_SIDE)


def safe_sheet_name(name):
    """Truncate and strip characters illegal in Excel sheet names."""
    illegal = '\\/:*?[]'
    for ch in illegal:
        name = name.replace(ch, '')
    return name[:31]


def generate_html_report(funds, output_path):
    """Generate a standalone interactive HTML visualization of common holdings with weights."""
    import json

    df = build_common_holdings(funds)

    fund_names = [name for name, _ in funds]

    COLORS = ['#378ADD', '#1D9E75', '#D85A30', '#7F77DD', '#BA7517', '#D4537E', '#639922', '#E24B4A', '#888780', '#185FA5']

    years = sorted(df['Year'].unique())

    held_cols = [f'{fn}_held' for fn in fund_names]

    # Keep every security held by 2+ funds in any year (the report only displays
    # overlaps). No top-N cut: truncating let selection order decide which funds
    # appeared at all.
    funds_per_cusip = df.groupby('CUSIP')[held_cols].max().sum(axis=1)
    overlapping = set(funds_per_cusip[funds_per_cusip >= 2].index)

    data_by_year = {}
    for year in years:
        ydf = df[(df['Year'] == year) & df['CUSIP'].isin(overlapping)].copy()

        ydf['_total'] = ydf[held_cols].sum(axis=1)

        ydf = ydf[ydf['_total'] >= 1].sort_values('_total', ascending=False, kind='stable')

        data_by_year[year] = [
            {
                'cusip': r['CUSIP'],
                'name': r['Security Name'],
                'held': [int(r[f'{fn}_held']) for fn in fund_names],
                'wt': [round(float(r[f'{fn}_weight']) / 100, 4) for fn in fund_names],
            }
            for _, r in ydf.iterrows()
        ]

    payload = json.dumps({
        'funds': fund_names,
        'colors': [COLORS[i % len(COLORS)] for i in range(len(fund_names))],
        'years': years,
        'data': data_by_year,
    })

    html_head = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>13F Portfolio Overlap</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f4f6f9;color:#1a1a1a;padding:2rem;max-width:1200px;margin:0 auto}
h1{font-size:22px;font-weight:500;color:#1F3864;margin-bottom:.2rem}
.sub{color:#666;font-size:13px;margin-bottom:1.5rem}
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(148px,1fr));gap:10px;margin-bottom:1.25rem}
.stat{background:#fff;border-radius:8px;border:.5px solid #e0e0e0;padding:.9rem 1.1rem}
.stat-lbl{font-size:11px;color:#888;margin-bottom:3px;text-transform:uppercase;letter-spacing:.04em}
.stat-val{font-size:22px;font-weight:500;color:#1F3864}
.stat-val.warn{color:#c0392b}
.controls{display:flex;gap:1.5rem;align-items:flex-start;flex-wrap:wrap;background:#fff;border-radius:10px;border:.5px solid #e0e0e0;padding:1rem 1.25rem;margin-bottom:1.5rem}
.clbl{font-size:11px;font-weight:500;color:#888;text-transform:uppercase;letter-spacing:.04em;margin-bottom:6px}
.fbtn{padding:4px 13px;border-radius:20px;border:1.5px solid;cursor:pointer;font-size:12px;font-weight:500;transition:all .15s}
.fbtn.active{color:#fff}
.fbtn.off{opacity:.3}
.fund-toggles{display:flex;gap:6px;flex-wrap:wrap}
select{padding:5px 9px;border-radius:6px;border:1px solid #ddd;font-size:12px;background:#fff;cursor:pointer}
.yr-range{display:flex;align-items:center;gap:8px;font-size:12px;color:#555}
.trow{display:flex;align-items:center;gap:10px;font-size:12px;color:#555}
.trow input{width:120px;cursor:pointer}
.card{background:#fff;border-radius:10px;border:.5px solid #e0e0e0;padding:1.5rem;margin-bottom:1.25rem}
.card-hdr{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:.75rem}
.card-title{font-size:14px;font-weight:500;color:#1a1a1a}
.card-hint{font-size:11px;color:#888}
.leg{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:10px;font-size:12px;color:#555}
.ldot{width:10px;height:10px;border-radius:2px;display:inline-block;margin-right:4px;vertical-align:middle}
.hm{overflow-x:auto}
table{border-collapse:collapse;width:100%;font-size:11px}
th{padding:5px 8px;text-align:left;font-weight:500;color:#888;border-bottom:1px solid #eee;white-space:nowrap}
td{padding:4px 8px;border-bottom:.5px solid #f0f0f0;white-space:nowrap}
.hmc{display:inline-block;min-width:56px;text-align:center;padding:3px 6px;border-radius:4px;font-weight:500;font-size:11px;line-height:1.5}
.hmc small{display:block;font-weight:400;font-size:10px;opacity:.75}
.alert-row td:first-child{border-left:3px solid #c0392b}
.flag{display:inline-block;background:#c0392b;color:#fff;font-size:10px;padding:1px 5px;border-radius:3px;margin-left:4px;vertical-align:middle}
.empty{text-align:center;color:#aaa;padding:2.5rem;font-size:13px}
</style>
</head>
<body>
<h1>13F Portfolio Overlap Analysis</h1>
<p class="sub">Identify unintentional duplicate exposure across funds — weights averaged across selected year range</p>
<div class="stats">
  <div class="stat"><div class="stat-lbl">Funds selected</div><div class="stat-val" id="s-funds">-</div></div>
  <div class="stat"><div class="stat-lbl">Overlapping holdings</div><div class="stat-val" id="s-hold">-</div></div>
  <div class="stat"><div class="stat-lbl">Above threshold</div><div class="stat-val warn" id="s-alerts">-</div></div>
  <div class="stat"><div class="stat-lbl">Highest combined</div><div class="stat-val" id="s-max">-</div></div>
  <div class="stat"><div class="stat-lbl">Years in range</div><div class="stat-val" id="s-yrs">-</div></div>
</div>
<div class="controls">
  <div>
    <div class="clbl">Year range</div>
    <div class="yr-range">
      <select id="yr-start"></select>
      <span>to</span>
      <select id="yr-end"></select>
    </div>
  </div>
  <div>
    <div class="clbl">Alert threshold</div>
    <div class="trow"><input type="range" id="thresh" min="2" max="30" value="10" step="1"><span id="tlbl" style="font-weight:500;min-width:36px">10%</span></div>
  </div>
  <div><div class="clbl">Funds</div><div class="fund-toggles" id="toggles"></div></div>
</div>
<div class="card">
  <div class="card-hdr"><div class="card-title">Combined exposure risk</div><div class="card-hint">Avg allocation per fund across selected years &middot; red line = alert threshold</div></div>
  <div class="leg" id="lg1"></div>
  <div id="w1" style="position:relative"></div>
</div>
<div class="card">
  <div class="card-hdr"><div class="card-title">Fund-by-fund weight breakdown</div><div class="card-hint">Average allocation per fund across years held &middot; hover for exact weight</div></div>
  <div class="leg" id="lg2"></div>
  <div id="w2" style="position:relative"></div>
</div>
<div class="card">
  <div class="card-hdr"><div class="card-title">Overlap heatmap</div><div class="card-hint">Avg weight / years held in range &middot; deeper blue = higher average weight &middot; flagged rows exceed threshold</div></div>
  <div class="hm" id="hm"></div>
</div>
<script>
const D="""

    html_tail = """;
const names=D.funds,colors=D.colors,years=D.years,byYear=D.data;
let startY=years[0],endY=years[years.length-1],active=new Set(names),thresh=0.10,c1=null,c2=null;

// Populate year dropdowns
const yrS=document.getElementById('yr-start'),yrE=document.getElementById('yr-end');
years.forEach(y=>{
  const a=document.createElement('option');a.value=y;a.textContent=y;yrS.appendChild(a);
  const b=document.createElement('option');b.value=y;b.textContent=y;yrE.appendChild(b);
});
yrS.value=startY;yrE.value=endY;
yrS.addEventListener('change',()=>{
  startY=yrS.value;
  if(endY<startY){yrE.value=startY;endY=startY;}
  render();
});
yrE.addEventListener('change',()=>{
  endY=yrE.value;
  if(startY>endY){yrS.value=endY;startY=endY;}
  render();
});

const ti=document.getElementById('thresh');
ti.addEventListener('input',()=>{thresh=parseInt(ti.value)/100;document.getElementById('tlbl').textContent=ti.value+'%';render();});

const tg=document.getElementById('toggles');
names.forEach((n,i)=>{
  const b=document.createElement('button');
  b.className='fbtn active';b.style.borderColor=colors[i];b.style.background=colors[i];b.textContent=n;
  b.addEventListener('click',()=>{
    if(active.has(n)&&active.size>1){active.delete(n);b.classList.add('off');b.classList.remove('active');b.style.background='#fff';b.style.color=colors[i];}
    else if(!active.has(n)){active.add(n);b.classList.remove('off');b.classList.add('active');b.style.background=colors[i];b.style.color='#fff';}
    render();
  });
  tg.appendChild(b);
});

function getDataForRange(){
  const inRange=years.filter(y=>y>=startY&&y<=endY);
  const cusipMap={};
  inRange.forEach(y=>{
    (byYear[y]||[]).forEach(r=>{
      if(!cusipMap[r.cusip])cusipMap[r.cusip]={name:r.name,fundWts:names.map(()=>[]),fundYrs:names.map(()=>0)};
      r.wt.forEach((w,fi)=>{if(w>0){cusipMap[r.cusip].fundWts[fi].push(w);cusipMap[r.cusip].fundYrs[fi]++;} });
    });
  });
  return Object.values(cusipMap).map(d=>{
    const wt=d.fundWts.map(ws=>ws.length?ws.reduce((s,w)=>s+w,0)/ws.length:0);
    const held=wt.map(w=>w>0?1:0);
    return{name:d.name,held,wt,yrsHeld:d.fundYrs,totalYrs:inRange.length};
  });
}

function wtColor(w,mx){
  if(w<=0)return'#f4f6f9';
  const t=Math.min(w/Math.max(mx,0.0001),1);
  return`rgb(${Math.round(230-t*150)},${Math.round(243-t*120)},${Math.round(255-t*120)})`;
}

function render(){
  const inRange=years.filter(y=>y>=startY&&y<=endY);
  const ai=names.map((n,i)=>active.has(n)?i:-1).filter(i=>i>=0);
  const raw=getDataForRange();
  const data=raw.map(r=>{
    const count=ai.reduce((s,i)=>s+r.held[i],0);
    const totalWt=ai.reduce((s,i)=>s+r.wt[i],0);
    return{...r,count,totalWt};
  }).filter(r=>r.count>=2).sort((a,b)=>b.totalWt-a.totalWt).slice(0,20);

  const alerts=data.filter(r=>r.totalWt>=thresh).length;
  const maxWt=data.length?data[0].totalWt:0;
  document.getElementById('s-funds').textContent=active.size;
  document.getElementById('s-hold').textContent=data.length;
  document.getElementById('s-alerts').textContent=alerts;
  document.getElementById('s-max').textContent=maxWt?(maxWt*100).toFixed(1)+'%':'—';
  document.getElementById('s-yrs').textContent=inRange.length;

  const legH=ai.map(fi=>`<span><span class="ldot" style="background:${colors[fi]}"></span>${names[fi]}</span>`).join('');
  document.getElementById('lg1').innerHTML=legH;
  document.getElementById('lg2').innerHTML=legH;

  if(!data.length){
    ['w1','w2','hm'].forEach(id=>document.getElementById(id).innerHTML='<div class="empty">No overlapping holdings for selected funds in this range.</div>');
    if(c1){c1.destroy();c1=null;}if(c2){c2.destroy();c2=null;}return;
  }

  const labels=data.map(r=>r.name.length>32?r.name.slice(0,30)+'…':r.name);
  const thrPct=thresh*100;

  (()=>{
    const h=Math.max(260,data.length*36+80);
    const el=document.getElementById('w1');el.style.height=h+'px';el.style.position='relative';
    if(!el.querySelector('canvas'))el.innerHTML='<canvas id="c1" role="img" aria-label="Combined exposure chart"></canvas>';
    const ds=ai.map(fi=>{return{label:names[fi],data:data.map(r=>+(r.wt[fi]*100).toFixed(2)),
      backgroundColor:data.map(r=>r.totalWt>=thresh?colors[fi]+'cc':colors[fi]+'55'),
      borderColor:data.map(r=>r.totalWt>=thresh?colors[fi]:'transparent'),
      borderWidth:1,barThickness:18,borderRadius:2};});
    const cfg={type:'bar',data:{labels,datasets:ds},options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
      scales:{
        x:{stacked:true,ticks:{font:{size:10},callback:v=>v+'%'},grid:{color:'rgba(0,0,0,0.05)'},
          title:{display:true,text:'Average combined allocation (across years held)',font:{size:10}}},
        y:{stacked:true,ticks:{font:{size:11},color:(ctx)=>data[ctx.index]&&data[ctx.index].totalWt>=thresh?'#c0392b':'#333'},grid:{display:false}},
      },
      plugins:{legend:{display:false},
        tooltip:{callbacks:{
          title:i=>data[i[0].dataIndex].name,
          label:i=>i.raw>0?` ${i.dataset.label}: ${i.raw.toFixed(2)}% avg`:null,
          afterBody:i=>{const r=data[i[0].dataIndex];const lines=[`Combined avg: ${(r.totalWt*100).toFixed(2)}%`];if(r.totalWt>=thresh)lines.push('⚠ Exceeds alert threshold');return lines;},
        },filter:i=>i.raw>0},
      },
    }};
    if(c1){c1.data=cfg.data;c1.options=cfg.options;c1.update('none');}else{c1=new Chart(document.getElementById('c1'),cfg);}
  })();

  (()=>{
    const h=Math.max(260,data.length*50+80);
    const el=document.getElementById('w2');el.style.height=h+'px';el.style.position='relative';
    if(!el.querySelector('canvas'))el.innerHTML='<canvas id="c2" role="img" aria-label="Per-fund weight breakdown"></canvas>';
    const ds=ai.map(fi=>{return{label:names[fi],data:data.map(r=>+(r.wt[fi]*100).toFixed(2)),
      backgroundColor:colors[fi]+'aa',borderColor:colors[fi],borderWidth:1,barThickness:10,borderRadius:2};});
    const cfg={type:'bar',data:{labels,datasets:ds},options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
      scales:{
        x:{stacked:false,ticks:{font:{size:10},callback:v=>v+'%'},grid:{color:'rgba(0,0,0,0.05)'},
          title:{display:true,text:'Average allocation % (years held only)',font:{size:10}}},
        y:{stacked:false,ticks:{font:{size:11}},grid:{display:false}},
      },
      plugins:{legend:{display:false},
        tooltip:{callbacks:{title:i=>data[i[0].dataIndex].name,
          label:i=>{if(!i.raw)return null;const r=data[i[0].dataIndex];const fi=ai[i.datasetIndex];const yh=r.yrsHeld[fi];const ty=r.totalYrs;return` ${i.dataset.label}: ${i.raw.toFixed(2)}% avg (${yh}/${ty} yrs)`;}},filter:i=>i.raw>0},
      },
    }};
    if(c2){c2.data=cfg.data;c2.options=cfg.options;c2.update('none');}else{c2=new Chart(document.getElementById('c2'),cfg);}
  })();

  (()=>{
    const mx=Math.max(...data.flatMap(r=>ai.map(i=>r.wt[i])));
    let h='<table><thead><tr><th>Security</th>';
    ai.forEach(fi=>h+=`<th style="color:${colors[fi]}">${names[fi]}</th>`);
    h+='<th>Combined avg</th></tr></thead><tbody>';
    data.forEach(r=>{
      const isA=r.totalWt>=thresh;
      h+=`<tr class="${isA?'alert-row':''}"><td style="font-weight:500">${r.name}${isA?'<span class="flag">!</span>':''}</td>`;
      ai.forEach(fi=>{
        const w=r.wt[fi];const yh=r.yrsHeld[fi];const ty=r.totalYrs;
        const bg=wtColor(w,mx);const tc=w>mx*0.5?'#1a3a6b':'#555';
        h+=`<td><span class="hmc" style="background:${bg};color:${tc}">${w>0?(w*100).toFixed(1)+'%':'—'}${w>0?`<small>${yh}/${ty} yrs</small>`:'' }</span></td>`;
      });
      const cb=isA?'#fde8e8':wtColor(r.totalWt,mx*ai.length);const cc=isA?'#c0392b':'#1a3a6b';
      h+=`<td><span class="hmc" style="background:${cb};color:${cc};font-weight:600">${(r.totalWt*100).toFixed(1)}%</span></td></tr>`;
    });
    h+='</tbody></table>';
    document.getElementById('hm').innerHTML=h;
  })();
}
render();
</script>
</body>
</html>"""

    html = html_head + payload + html_tail

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f'Saved HTML:  {output_path}')


def write_excel(funds, output_path):
    """Write one sheet per fund plus a summary sheet into a single workbook."""
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        wb = writer.book

        for entity_name, df in funds:
            sheet_name = safe_sheet_name(entity_name)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]
            _format_sheet(ws, df)

        _add_summary_sheet(wb, funds)
        if len(funds) > 1:
            _add_common_holdings_sheet(wb, funds)

    print(f'\nSaved: {output_path}')


def _format_sheet(ws, df):
    col_widths = {
        'Quarter': 12,
        'Ticker': 10,
        'CUSIP': 12,
        'Security Name': 36,
        'Shares': 18,
        'Market Value ($)': 20,
        '% of Fund': 12,
        'QoQ Delta MV ($)': 22,
    }

    for col_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.column_dimensions[get_column_letter(col_idx)].width = col_widths.get(col_name, 14)

    ws.freeze_panes = 'A2'

    num_cols = {
        'Shares': '#,##0',
        'Market Value ($)': '"$"#,##0',
        '% of Fund': '0.00%',
        'QoQ Delta MV ($)': '"$"#,##0;[Red]-"$"#,##0',
    }

    col_map = {name: idx + 1 for idx, name in enumerate(df.columns)}

    for row_idx in range(2, len(df) + 2):
        fill = ALT_FILL if row_idx % 2 == 0 else None
        for col_name, col_idx in col_map.items():
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = BODY_FONT
            cell.border = THIN_BORDER
            if fill:
                cell.fill = fill
            if col_name not in num_cols:
                continue
            cell.number_format = num_cols[col_name]
            if col_name != '% of Fund':
                continue
            if cell.value is None:
                continue
            cell.value = cell.value / 100

    ws.auto_filter.ref = ws.dimensions


def _add_summary_sheet(wb, funds):
    ws = wb.create_sheet('Summary', 0)

    ws.column_dimensions['A'].width = 36
    ws.column_dimensions['B'].width = 16
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 20
    ws.column_dimensions['F'].width = 20

    title_font = Font(name='Calibri', bold=True, size=14, color='1F3864')
    header_font = Font(name='Calibri', bold=True, size=10, color='FFFFFF')
    label_font = Font(name='Calibri', size=10)
    header_fill = PatternFill('solid', fgColor='1F3864')

    ws['A1'] = '13F Holdings — Summary'
    ws['A1'].font = title_font

    headers = ['Fund', 'Sheet', 'First Quarter', 'Last Quarter', 'Quarters', 'Unique CUSIPs']

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')

    for row, (entity_name, df) in enumerate(funds, start=4):
        sheet_name = safe_sheet_name(entity_name)

        def quarter_sort_key(q):
            num, year = q.split(' ')
            return (int(year), int(num[1]))

        quarters = sorted(df['Quarter'].unique(), key=quarter_sort_key)

        ws.cell(row=row, column=1, value=entity_name).font = label_font
        ws.cell(row=row, column=2, value=sheet_name).font = label_font
        ws.cell(row=row, column=3, value=quarters[0] if quarters else '').font = label_font
        ws.cell(row=row, column=4, value=quarters[-1] if quarters else '').font = label_font
        ws.cell(row=row, column=5, value=len(quarters)).font = label_font
        ws.cell(row=row, column=6, value=df['CUSIP'].nunique()).font = label_font


def run_scraper(ciks, output_path, limit=None, progress_cb=None, skipped_out=None):
    """Scrape 13F holdings for one or more CIKs and write Excel (+ HTML overlap
    report if 2+ funds produced data). Returns (funds, html_path) where funds is
    a list of (entity_name, DataFrame) tuples and html_path is str|None.

    limit is the number of most recent quarters *with usable holdings* per fund;
    a newer filing with no holdings table (e.g. a cover-page-only amendment) is
    passed over in favor of the next one. progress_cb, if given, is called with
    each log line instead of printing it. skipped_out, if given, is a list that
    receives (fund name, reason) for every fund that produced no data.
    """
    def log(msg):
        if progress_cb:
            progress_cb(msg)
        else:
            print(msg)

    def skip(name, reason):
        log(f'WARNING: {name}: {reason} -- skipping.')
        if skipped_out is not None:
            skipped_out.append((name, reason))

    funds = []
    for raw_cik in ciks:
        cik = raw_cik.strip().lstrip('0') or '0'
        log('=' * 56)
        try:
            filings, entity_name = get_all_13f_filings(cik, dedupe=False)
        except Exception as e:
            skip(f'CIK {cik}', f'could not load filing list from EDGAR ({e})')
            continue

        by_quarter = {}
        for f in sorted(filings, key=lambda x: x['filingDate'], reverse=True):
            by_quarter.setdefault(quarter_label(f['filingDate']), []).append(f)
        log(f'Processing {entity_name} ({len(by_quarter)} quarters on file)...')

        all_quarters = []
        problems = []
        for quarter, candidates in by_quarter.items():
            if limit and len(all_quarters) >= limit:
                break
            for filing in candidates:
                fdate = filing['filingDate']
                log(f'  {quarter}  ({fdate})  {filing["form"]}')
                xml_url = find_info_table_url(cik, filing['accessionNumber'])
                if not xml_url:
                    problems.append(f'{fdate}: no information table found')
                    log('    -> No info table found.')
                    time.sleep(0.5)
                    continue
                try:
                    holdings = parse_info_table(xml_url, filing_date=fdate)
                    log(f'    -> {len(holdings)} positions')
                except Exception as e:
                    problems.append(f'{fdate}: error reading information table ({e})')
                    log(f'    -> ERROR: {e}')
                    holdings = []
                time.sleep(0.5)
                if holdings:
                    all_quarters.append((quarter, holdings))
                    break
                if not problems or not problems[-1].startswith(fdate):
                    problems.append(f'{fdate}: information table had 0 positions')

        if not all_quarters:
            skip(entity_name, '; '.join(problems) or 'no 13F-HR filings found')
            continue

        log(f'Building dataset from {len(all_quarters)} quarters...')
        df = build_dataframe(all_quarters)
        log(f'Total rows: {len(df):,}  |  Unique CUSIPs: {df["CUSIP"].nunique():,}')
        funds.append((entity_name, df))

    if not funds:
        raise ValueError('No data retrieved for any fund. Check the CIK values and try again.')

    log(f'Writing Excel: {output_path}')
    write_excel(funds, output_path)
    log('Excel saved.')

    html_path = None
    if len(funds) > 1:
        html_path = output_path.replace('.xlsx', '.html')
        log('Generating HTML report...')
        try:
            generate_html_report(funds, html_path)
            log(f'HTML saved: {html_path}')
        except Exception as e:
            log(f'WARNING: HTML generation failed: {e}')
            html_path = None
    else:
        log('Note: HTML overlap report requires 2 or more funds.')

    return funds, html_path


def main():
    parser = argparse.ArgumentParser(description='Scrape SEC 13F filings into Excel.')
    parser.add_argument('--cik', nargs='+', help='One or more EDGAR CIK numbers (e.g. 1336528 1350694)')
    parser.add_argument('--output', default='13f_holdings.xlsx', help='Output Excel file path')
    parser.add_argument('--limit', type=int, default=None, help='Limit to N most recent filings per fund (for testing)')
    args = parser.parse_args()

    if not args.cik:
        sys.exit('Provide at least one CIK with --cik.')

    funds = []

    for raw_cik in args.cik:
        cik = raw_cik.lstrip('0') or '0'

        print(f"\n{'=' * 60}")

        filings, entity_name = get_all_13f_filings(cik)

        if args.limit:
            filings = filings[-args.limit:]

        print(f'\nProcessing {len(filings)} filings for {entity_name}...\n')

        all_quarters = []

        for i, filing in enumerate(filings, 1):
            quarter = quarter_label(filing['filingDate'])
            acc = filing['accessionNumber']

            print(f'  [{i:>3}/{len(filings)}] {quarter}  ({filing["filingDate"]})  {acc}', end='  ')

            xml_url = find_info_table_url(cik, acc)
            if not xml_url:
                print('-> No info table found, skipping.')
                continue

            try:
                holdings = parse_info_table(xml_url, filing_date=filing['filingDate'])
                print(f'-> {len(holdings)} positions')
            except Exception as e:
                print(f'-> ERROR: {e}')
                holdings = []

            if holdings:
                all_quarters.append((quarter, holdings))

            time.sleep(0.5)

        if not all_quarters:
            print(f'WARNING: No holdings data for {entity_name}, skipping.')
            continue

        print(f'\nBuilding dataset from {len(all_quarters)} quarters...')

        df = build_dataframe(all_quarters)
        print(f'Total rows: {len(df):,}  |  Unique CUSIPs: {df["CUSIP"].nunique():,}')

        funds.append((entity_name, df))

    if not funds:
        sys.exit('No data retrieved for any fund.')

    print(f"\n{'=' * 60}")

    print(f'Writing Excel: {args.output}')
    write_excel(funds, args.output)

    if len(funds) > 1:
        html_path = args.output.replace('.xlsx', '.html')
        generate_html_report(funds, html_path)
        return None

    return None


if __name__ == '__main__':
    main()


# ---------------------------------------------------------------------------
# NOTE: This module was reconstructed from decompiled/disassembled bytecode
# (no original source was available) as part of a legitimate reverse-engineering
# effort on the requesting user's own commercial tool. While control flow,
# string literals, and logic were carefully cross-checked against the raw
# CPython 3.14 bytecode, this file should be spot-checked against real SEC
# EDGAR responses (and diffed/tested) before being relied upon in production,
# particularly around: the XML namespace/tag fallback logic in
# parse_info_table/_text/_int, the exact HTML/JS report template in
# generate_html_report, and the openpyxl formatting helpers.
# ---------------------------------------------------------------------------
