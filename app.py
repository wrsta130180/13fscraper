"""
13F Fund Analyzer — web front end.

A Streamlit port of the original desktop (Tkinter) 13F Analyzer. Scrapes SEC
EDGAR 13F-HR filings for one or more institutional managers and produces:
  - An Excel workbook (one sheet per fund + a Summary sheet, + a Common
    Holdings sheet when 2+ funds are run together)
  - An interactive HTML "portfolio overlap" report (when 2+ funds are run)

All scraping/analysis logic lives in sec_13f_scraper.py (unchanged core module,
recovered from the original tool and fixed for the SEC's 2023 reporting-format
change — see VALUE_IN_THOUSANDS_CUTOFF in that file).
"""
import io
import json
import os
import tempfile

import streamlit as st

import sec_13f_scraper as scraper

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
CUSTOM_FUNDS_FILE = os.path.join(DATA_DIR, 'custom_funds.json')

COMMON_FUNDS = [
    ('Praetorian PR LLC', '1949877'),
    ('Carronade Capital Management', '1866872'),
    ('TCI Fund Management', '1647251'),
    ('CastleKnight Management', '1835751'),
    ('Naya Capital Management', '1665012'),
    ('LFL Advisers', '1694127'),
    ('Dorsal Capital Management', '1547007'),
    ('Merewether Investment Mgmt', '1736852'),
]


def load_custom_funds():
    try:
        with open(CUSTOM_FUNDS_FILE, 'r') as f:
            return json.load(f)
    except Exception:
        return []


def save_custom_funds(funds):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(CUSTOM_FUNDS_FILE, 'w') as f:
        json.dump(funds, f, indent=2)


st.set_page_config(page_title='13F Fund Analyzer', layout='wide')

if 'custom_funds' not in st.session_state:
    st.session_state.custom_funds = load_custom_funds()
if 'log_lines' not in st.session_state:
    st.session_state.log_lines = []
if 'result' not in st.session_state:
    st.session_state.result = None  # dict: xlsx_bytes, html_text, funds (list of names)

st.title('13F Fund Analyzer')
st.caption(
    'Scrapes SEC EDGAR 13F-HR filings and builds an Excel workbook plus an '
    'interactive portfolio-overlap report. Runs entirely in your browser session — '
    'nothing is installed locally.'
)

with st.sidebar:
    st.header('Tracked funds')
    st.caption('Shared list — anyone using this app can add to it.')

    all_tracked = COMMON_FUNDS + [tuple(f) for f in st.session_state.custom_funds]

    with st.form('add_fund_form', clear_on_submit=True):
        new_name = st.text_input('Manager name')
        new_cik = st.text_input('CIK')
        submitted = st.form_submit_button('Add fund')
        if submitted:
            name = new_name.strip()
            cik = new_cik.strip().lstrip('0') or '0'
            if not (name and cik):
                st.error('Enter both a manager name and CIK.')
            elif any(c == cik for _, c in all_tracked):
                st.warning(f'CIK {cik} is already in the list.')
            else:
                st.session_state.custom_funds.append([name, cik])
                save_custom_funds(st.session_state.custom_funds)
                st.rerun()

    st.divider()
    st.caption('Built-in funds')
    for name, cik in COMMON_FUNDS:
        st.text(f'{name}  ({cik})')

    if st.session_state.custom_funds:
        st.caption('Custom tracked funds')
        for i, (name, cik) in enumerate(st.session_state.custom_funds):
            col1, col2 = st.columns([5, 1])
            col1.text(f'{name}  ({cik})')
            if col2.button('✕', key=f'remove_{i}', help='Remove'):
                st.session_state.custom_funds.pop(i)
                save_custom_funds(st.session_state.custom_funds)
                st.rerun()

st.subheader('Run analysis')

all_tracked = COMMON_FUNDS + [tuple(f) for f in st.session_state.custom_funds]
labels = [f'{name} ({cik})' for name, cik in all_tracked]
label_to_cik = {f'{name} ({cik})': cik for name, cik in all_tracked}

selected_labels = st.multiselect('Select tracked funds to include', labels)
extra_ciks_text = st.text_area(
    'Additional CIKs (optional — comma, space, or newline separated)',
    placeholder='e.g. 0001350694, 1336528',
)

col1, col2 = st.columns(2)
with col1:
    limit = st.number_input(
        'Number of quarters', min_value=1, value=1, step=1,
        help='Fetches this many of each fund\'s most recent quarterly 13F filings.',
    )
with col2:
    output_name = st.text_input('Output file name', value='fund_overlap')

run_clicked = st.button('▶  Run analysis', type='primary')

if run_clicked:
    ciks = [label_to_cik[l] for l in selected_labels]
    extra = extra_ciks_text.replace(',', ' ').split()
    ciks += [c.strip() for c in extra if c.strip()]
    ciks = list(dict.fromkeys(ciks))  # de-dupe, preserve order

    if not ciks:
        st.error('Select at least one tracked fund or enter a CIK.')
    else:
        st.session_state.result = None
        log_box = st.empty()
        lines = []

        def progress_cb(msg):
            lines.append(msg)
            log_box.code('\n'.join(lines[-25:]), language=None)

        fname = output_name.strip() or 'fund_overlap'
        if not fname.endswith('.xlsx'):
            fname += '.xlsx'

        with tempfile.TemporaryDirectory() as tmp:
            xlsx_path = os.path.join(tmp, fname)
            try:
                with st.spinner('Scraping SEC EDGAR...'):
                    funds, html_path = scraper.run_scraper(
                        ciks,
                        xlsx_path,
                        limit=(limit or None),
                        progress_cb=progress_cb,
                    )
                with open(xlsx_path, 'rb') as f:
                    xlsx_bytes = f.read()
                html_text = None
                if html_path and os.path.exists(html_path):
                    with open(html_path, 'r', encoding='utf-8') as f:
                        html_text = f.read()
                st.session_state.result = {
                    'xlsx_bytes': xlsx_bytes,
                    'xlsx_name': fname,
                    'html_text': html_text,
                    'html_name': fname.replace('.xlsx', '.html'),
                    'fund_names': [name for name, _ in funds],
                }
                st.success(f'Done — {len(funds)} fund(s) processed.')
            except Exception as e:
                st.error(f'Run failed: {e}')

result = st.session_state.result
if result:
    st.subheader('Results')
    st.write(', '.join(result['fund_names']))

    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            'Download Excel workbook',
            data=result['xlsx_bytes'],
            file_name=result['xlsx_name'],
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
    with dl2:
        if result['html_text']:
            st.download_button(
                'Download HTML overlap report',
                data=result['html_text'],
                file_name=result['html_name'],
                mime='text/html',
            )
        else:
            st.caption('HTML overlap report requires 2+ funds.')

    if result['html_text']:
        st.subheader('Portfolio overlap report (preview)')
        st.components.v1.html(result['html_text'], height=1200, scrolling=True)
