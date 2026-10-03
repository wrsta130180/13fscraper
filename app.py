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
FUNDS_FILE = os.path.join(DATA_DIR, 'funds.json')
LEGACY_CUSTOM_FILE = os.path.join(DATA_DIR, 'custom_funds.json')


def _read_json_list(path):
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return [[str(n), str(c)] for n, c in data]
    except Exception:
        return []


def load_funds():
    """The single shared fund list: a JSON array of [name, cik] pairs in data/funds.json."""
    funds = _read_json_list(FUNDS_FILE)
    # One-time migration of funds added under the old "custom funds" scheme.
    legacy = _read_json_list(LEGACY_CUSTOM_FILE)
    if legacy:
        known = {c for _, c in funds}
        funds += [f for f in legacy if f[1] not in known]
        save_funds(funds)
        os.replace(LEGACY_CUSTOM_FILE, LEGACY_CUSTOM_FILE + '.migrated')
    return funds


def save_funds(funds):
    os.makedirs(DATA_DIR, exist_ok=True)
    tmp = FUNDS_FILE + '.tmp'
    with open(tmp, 'w', encoding='utf-8', newline='\n') as f:
        json.dump(funds, f, indent=2)
        f.write('\n')
    os.replace(tmp, FUNDS_FILE)

st.set_page_config(page_title='13F Fund Analyzer', layout='wide')

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
    funds_list = load_funds()
    st.header('Built-in funds')
    st.caption(
        'Funds you add here are saved to the built-in list for everyone using this app, '
        'until you remove them.'
    )

    with st.form('add_fund_form', clear_on_submit=True):
        new_name = st.text_input('Manager name')
        new_cik = st.text_input('CIK')
        submitted = st.form_submit_button('Add fund')
        if submitted:
            name = new_name.strip()
            cik = new_cik.strip().lstrip('0') or '0'
            if not (name and new_cik.strip()):
                st.error('Enter both a manager name and CIK.')
            elif not cik.isdigit():
                st.error('CIK must be a number.')
            elif any(c == cik for _, c in funds_list):
                st.warning(f'CIK {cik} is already in the list.')
            else:
                save_funds(funds_list + [[name, cik]])
                st.rerun()

    st.divider()
    with st.expander(f'Remove funds ({len(funds_list)} in list)'):
        remove_labels = st.multiselect(
            'Select funds to remove',
            [f'{n} ({c})' for n, c in funds_list],
            key='remove_select',
        )
        if st.button('Remove selected', disabled=not remove_labels):
            gone = set(remove_labels)
            save_funds([[n, c] for n, c in funds_list if f'{n} ({c})' not in gone])
            st.session_state.pop('remove_select', None)
            st.rerun()

    for name, cik in funds_list:
        st.text(f'{name}  ({cik})')

st.subheader('Run analysis')

all_tracked = [tuple(f) for f in funds_list]
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

        skipped = []
        with tempfile.TemporaryDirectory() as tmp:
            xlsx_path = os.path.join(tmp, fname)
            try:
                with st.spinner('Scraping SEC EDGAR...'):
                    funds, html_path = scraper.run_scraper(
                        ciks,
                        xlsx_path,
                        limit=(limit or None),
                        progress_cb=progress_cb,
                        skipped_out=skipped,
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
                    'requested': len(ciks),
                    'skipped': skipped,
                    'log': '\n'.join(lines),
                }
                st.success(f'Done — {len(funds)} of {len(ciks)} fund(s) returned data.')
            except Exception as e:
                st.error(f'Run failed: {e}')

result = st.session_state.result
if result:
    st.subheader('Results')
    st.write(', '.join(result['fund_names']))

    if result['skipped']:
        st.warning(
            f"{len(result['skipped'])} fund(s) returned no data and are NOT in the results:\n\n"
            + '\n'.join(f'- **{name}** — {reason}' for name, reason in result['skipped'])
        )
    with st.expander('Full run log'):
        st.code(result['log'], language=None)

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
