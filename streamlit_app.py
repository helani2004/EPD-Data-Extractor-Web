"""Browser interface for the existing EPD extraction engine."""
import copy
import io
import os
import tempfile
from datetime import datetime
from pathlib import Path

import streamlit as st
from app import process, save_excel, HEADERS, GROUPS, IMPACT_FIELDS
from ai_review import review_pdf

st.set_page_config(page_title='EPD Data Extractor', page_icon='📄', layout='wide')
st.title('📄 EPD Data Extractor')
st.caption('Upload an EPD PDF → extract data with Python and Gemini → review/edit → download Excel.')

if 'records' not in st.session_state:
    st.session_state.records = []
if 'current' not in st.session_state:
    st.session_state.current = None
if 'current_name' not in st.session_state:
    st.session_state.current_name = ''

try:
    api_key = st.secrets.get('GEMINI_API_KEY', '')
except Exception:
    api_key = ''
api_key = api_key or os.environ.get('GEMINI_API_KEY', '')
with st.sidebar:
    st.header('Settings')
    st.write('Gemini API key: ' + ('Configured securely ✓' if api_key else 'Not configured'))
    model = st.text_input('Gemini model', value='gemini-flash-latest')
    use_ai = st.checkbox('Run Gemini AI extraction', value=bool(api_key), disabled=not bool(api_key))
    st.caption('Only the website owner configures the key in Streamlit Cloud Secrets. Uploaded PDFs are sent to Google when AI is enabled.')
    st.info('Records are kept only in this browser session. Download your Excel before closing it.')

uploads = st.file_uploader('Choose EPD PDF files', type=['pdf'], accept_multiple_files=True)

if st.button('Extract uploaded PDFs', type='primary', disabled=not uploads):
    for idx, uploaded in enumerate(uploads):
        st.write(f'Processing {idx+1}/{len(uploads)}: {uploaded.name}')
        with st.spinner('Reading PDF tables and fields...'):
            with tempfile.TemporaryDirectory(prefix='epd_web_') as folder:
                path = Path(folder) / Path(uploaded.name).name
                path.write_bytes(uploaded.getvalue())
                try:
                    result = process(path)
                except Exception as exc:
                    st.error(f'Python extraction failed for {uploaded.name}: {exc}')
                    continue
                if use_ai:
                    with st.spinner('Gemini is reviewing the PDF (this may take several minutes)...'):
                        try:
                            result = review_pdf(path, result, HEADERS, GROUPS, IMPACT_FIELDS, api_key, model, Path(folder)/'ai_cache')
                        except Exception as exc:
                            result['warnings'].insert(0, 'AI PASS FAILED: ' + str(exc))
                            st.warning(f'Gemini failed for {uploaded.name}; Python results retained.')
                st.session_state.records.append({
                    'id': len(st.session_state.records)+1,
                    'created': datetime.now().isoformat(timespec='seconds'),
                    'result': result,
                })
                st.session_state.current = len(st.session_state.records)-1
                st.session_state.current_name = uploaded.name
    st.success('Processing finished. Review your fields below.')

records = st.session_state.records
if records:
    choices = [f"{i+1}. {rec['result'].get('file', 'EPD')}" for i,rec in enumerate(records)]
    selected = st.selectbox('Choose a record to review', range(len(records)), index=st.session_state.current or 0, format_func=lambda i: choices[i])
    rec = records[selected]
    result = rec['result']
    st.subheader('Review / Edit extracted values')
    st.caption('Edit a value directly in the table. Leave unavailable fields blank. Changes are included in the Excel download.')
    import pandas as pd
    keys = list(result['data'].keys())
    frame = pd.DataFrame({'Field': keys, 'Value': [str(result['data'].get(k,'') or '') for k in keys]})
    edited = st.data_editor(frame, hide_index=True, use_container_width=True, disabled=['Field'], key=f'edit_{selected}', num_rows='fixed')
    if st.button('Save edits', key=f'save_{selected}'):
        for _, row in edited.iterrows():
            result['data'][row['Field']] = row['Value']
        st.success('Edits saved for this session.')
    with st.expander('Extraction warnings and source evidence'):
        for warning in result.get('warnings',[]):
            st.write('• '+str(warning))
        if result.get('ai_review'):
            st.dataframe(pd.DataFrame(result['ai_review']), use_container_width=True)
        if result.get('evidence'):
            st.dataframe(pd.DataFrame(result['evidence']), use_container_width=True)

    if st.button('Remove selected record'):
        records.pop(selected)
        st.session_state.current = 0 if records else None
        st.rerun()

    with tempfile.TemporaryDirectory(prefix='epd_export_') as folder:
        output = Path(folder)/'EPD_Extracted_Results.xlsx'
        try:
            save_excel(records, output)
            st.download_button('⬇ Download Excel (all records)', data=output.read_bytes(), file_name='EPD_Extracted_Results.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', type='primary')
        except Exception as exc:
            st.error(f'Excel export failed: {exc}')
else:
    st.info('Upload at least one PDF to begin.')
