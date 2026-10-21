"""Render-ready Flask frontend for the existing EPD extraction engine."""
import os
import io
import math
import time
import uuid
import tempfile
import threading
from datetime import datetime
from pathlib import Path
from collections import OrderedDict

from flask import Flask, request, redirect, url_for, session, render_template_string, send_file, flash, abort
from werkzeug.exceptions import RequestEntityTooLarge
from extractor_core import process, save_excel, HEADERS, GROUPS, IMPACT_FIELDS
from ai_review import review_pdf
from category_export import save_category_excel
from material_categories import CATEGORIES
from database import (
    list_records,
    get_record,
    insert_record,
    update_record,
    delete_record,
    DatabaseError,
)

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET_KEY') or os.urandom(32)
app.config['MAX_CONTENT_LENGTH'] = 35 * 1024 * 1024
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['SESSION_COOKIE_SECURE'] = os.environ.get('RENDER', '').lower() == 'true'
API_KEY = os.environ.get('GEMINI_API_KEY', '').strip()
MODEL = os.environ.get('EPD_GEMINI_MODEL', 'gemini-flash-latest').strip()


PAGE = '''<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>EPD Data Extractor</title><style>
:root{font-family:system-ui,-apple-system,Segoe UI,sans-serif;color:#192b3d;background:#f3f6fa}*{box-sizing:border-box}body{margin:0}
header{background:#152a43;color:white;padding:23px max(20px,calc((100vw - 1150px)/2))}header h1{margin:0 0 5px;font-size:26px}header p{margin:0;color:#c8d6e6}
main{max-width:1150px;margin:24px auto;padding:0 16px}.panel{background:white;border:1px solid #e0e7ef;border-radius:12px;padding:22px;margin-bottom:20px;box-shadow:0 3px 15px #192b3d0a}
h2{font-size:19px;margin:0 0 15px}label{font-weight:600}input[type=file],input[type=text],select{padding:10px;border:1px solid #c7d3df;border-radius:7px;max-width:100%;font:inherit}input[type=file]{display:block;width:100%;margin:12px 0}.actions{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin-top:15px}button,.btn{background:#1469a3;border:0;border-radius:7px;color:white;padding:11px 17px;text-decoration:none;cursor:pointer;font-weight:600;font-size:14px;display:inline-block}.btn.secondary,button.secondary{background:#e7eef5;color:#173a58}.btn.danger,button.danger{background:#9b3434}button:hover,.btn:hover{filter:brightness(.93)}small,.muted{color:#5a6979}.notice{padding:12px 15px;border-radius:7px;margin-bottom:12px;background:#e8f4ff}.notice.error{background:#ffebeb}.status{font-size:14px;color:#526579}.scroll{max-height:550px;overflow:auto;border:1px solid #e4e9ef;border-radius:8px}table{border-collapse:collapse;width:100%}td,th{padding:9px 12px;border-bottom:1px solid #e7edf2;text-align:left;font-size:13px;vertical-align:middle}th{background:#f0f5f9;position:sticky;top:0}td:first-child{width:48%;word-break:break-word}td input{width:100%;min-width:180px}.topline{display:flex;justify-content:space-between;align-items:center;gap:15px;flex-wrap:wrap}.fieldhead{font-weight:700;margin:16px 0 8px}.badge{font-size:12px;background:#e5f3ea;color:#22673b;padding:4px 9px;border-radius:12px}.warning{background:#fff6e3;padding:10px;border-radius:7px;margin:7px 0;word-break:break-word}details{margin-top:14px}details summary{cursor:pointer;font-weight:600}code{font-size:12px} .grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}@media(max-width:650px){.grid{grid-template-columns:1fr}header{padding:20px}}
</style></head><body><header><h1>📄 EPD Data Extractor</h1><p>Upload PDF → extract → review and edit → download 111-column Excel</p></header><main>
{% for kind,message in get_flashed_messages(with_categories=true) %}<div class="notice {{ 'error' if kind=='error' else '' }}">{{ message }}</div>{% endfor %}
<section class="panel"><div class="topline"><h2>1. Upload EPD PDFs</h2><span class="badge">Gemini: {{ 'Configured' if ai_ready else 'Not configured' }}</span></div>
<form method="post" action="{{ url_for('upload') }}" enctype="multipart/form-data">
<label for="category">Material Category (required)</label>
<select id="category" name="material_category" required>
    <option value="">Select a material category</option>
    {% for category in categories %}
    <option value="{{ category }}">{{ category }}</option>
    {% endfor %}
</select>

<p class="status">
    The selected category applies to every PDF in this upload.
    Upload different categories separately.
</p>

<label for="pdfs">Choose one or more PDF files</label>
<input id="pdfs" type="file" name="pdfs"
       accept=".pdf,application/pdf" multiple required>
{% if ai_ready %}<label><input type="checkbox" name="use_ai" value="1" checked> Run Gemini AI extraction after Python extraction</label>{% else %}<p class="status">Python extraction works without Gemini. To enable AI, the website owner must add GEMINI_API_KEY in Render environment settings.</p>{% endif %}
<div class="actions"><button type="submit">Extract uploaded PDFs</button><small>Processing can take several minutes. Please keep this tab open.</small></div></form></section>
{% if records %}<section class="panel"><div class="topline"><h2>2. Review extracted results ({{ records|length }} records)</h2><a class="btn" href="{{ url_for('download') }}">⬇ Download Excel (all records)</a></div>
<form method="get" action="{{ url_for('index') }}" class="actions"><label for="record">Select record</label><select name="record" id="record" onchange="this.form.submit()">{% for r in records %}<option value="{{ loop.index0 }}" {{ 'selected' if loop.index0 == selected else '' }}>{{ loop.index }}. {{ r['result'].get('file','EPD') }}</option>{% endfor %}</select><noscript><button type="submit">Open</button></noscript></form>
{% if record %}<p class="status">Edit any field and click <strong>Save edits</strong> before downloading. Blank fields mean unavailable or not recovered.</p>
<form method="post" action="{{ url_for('edit',record_id=record['id']) }}"><div class="scroll"><table><thead><tr><th>Excel field</th><th>Extracted value (editable)</th></tr></thead><tbody>
{% for key,value in record['result']['data'].items() %}
<tr>
    <td>{{ key.replace('__',' | ') }}</td>
    <td>
        {% if key == 'Material Category' %}
        <select name="f_{{ loop.index0 }}" required>
            <option value="">Select a material category</option>
            {% for category in categories %}
            <option value="{{ category }}"
                    {% if category == value %}selected{% endif %}>
                {{ category }}
            </option>
            {% endfor %}
        </select>
        {% else %}
        <input type="text"
               name="f_{{ loop.index0 }}"
               value="{{ value if value is not none else '' }}"
               autocomplete="off">
        {% endif %}
    </td>
</tr>
{% endfor %}
</tbody></table></div><div class="actions"><button type="submit">Save edits</button><a class="btn secondary" href="{{ url_for('download') }}">Download Excel</a></div></form>
<form method="post" action="{{ url_for('remove',record_id=record['id']) }}" class="actions"><button class="danger" type="submit" onclick="return confirm('Remove this record?')">Remove selected record</button></form>
<details><summary>Extraction warnings ({{ record['result'].get('warnings',[])|length }})</summary>{% for w in record['result'].get('warnings',[]) %}<div class="warning">{{ w }}</div>{% endfor %}</details>
<details><summary>Source evidence ({{ record['result'].get('evidence',[])|length }})</summary><div class="scroll"><table><tr><th>Field</th><th>Value</th><th>Page</th><th>Source text</th></tr>{% for e in record['result'].get('evidence',[]) %}<tr><td>{{ e.get('field','') }}</td><td>{{ e.get('value','') }}</td><td>{{ e.get('page','') }}</td><td>{{ e.get('raw','') }}</td></tr>{% endfor %}</table></div></details>
<details><summary>Gemini AI review ({{ record['result'].get('ai_review',[])|length }})</summary><div class="scroll"><table><tr><th>Field</th><th>AI proposal</th><th>Decision</th><th>Page</th></tr>{% for e in record['result'].get('ai_review',[]) %}<tr><td>{{ e.get('field','') }}</td><td>{{ e.get('proposed','') }}</td><td>{{ e.get('decision','') }}</td><td>{{ e.get('page','') }}</td></tr>{% endfor %}</table></div></details>
{% endif %}</section>{% else %}<section class="panel"><p class="muted">Your extracted records will appear here.</p></section>{% endif %}
<section class="panel"><small>Records are stored in the shared project database. All website users can view, edit, export and remove saved records.</small></section>
<p>
    <a class="btn secondary" href="{{ url_for('download') }}">
        Download database Excel — Combined + 63 categories
    </a>
</p>
</main></body></html>'''


def user_records():
    """Load saved records from the shared Turso database."""
    return list_records()


@app.get('/')
def index():
    records = user_records()
    try:
        selected = int(request.args.get('record', len(records)-1))
    except ValueError:
        selected = 0
    selected = max(0, min(selected, len(records)-1)) if records else 0
    return render_template_string(PAGE, records=records, record=records[selected] if records else None,
                                  selected=selected,
                                  ai_ready=bool(API_KEY),
                                  categories=CATEGORIES)


@app.post('/upload')
def upload():
    category = request.form.get('material_category', '')

    if category not in CATEGORIES:
        flash('Please select a valid material category.', 'error')
        return redirect(url_for('index'))

    uploads = [f for f in request.files.getlist('pdfs') if f.filename]
    if not uploads:
        flash('Please select at least one PDF.', 'error')
        return redirect(url_for('index'))
    if len(uploads) > 5:
        flash('Please upload no more than 5 PDFs at a time.', 'error')
        return redirect(url_for('index'))
    use_ai = bool(API_KEY and request.form.get('use_ai') == '1')
    for upload_file in uploads:
        name = Path(upload_file.filename).name
        if not name.lower().endswith('.pdf'):
            flash(f'Skipped {name}: PDF files only.', 'error')
            continue
        with tempfile.TemporaryDirectory(prefix='epd_render_') as directory:
            path = Path(directory) / 'input.pdf'
            content = upload_file.read()
            if not content.startswith(b'%PDF-'):
                flash(f'Skipped {name}: file is not a valid PDF.', 'error')
                continue
            if len(content) > 15 * 1024 * 1024:
                flash(f'Skipped {name}: maximum file size is 15 MB.', 'error')
                continue
            path.write_bytes(content)
            try:
                result = process(path)
                result['file'] = name
                for item in result.get('values', []):
                    item['file'] = name
                if use_ai:
                    try:
                        result = review_pdf(path, result, HEADERS, GROUPS, IMPACT_FIELDS,
                                            API_KEY, MODEL, Path(directory) / 'ai_cache')
                    except Exception as exc:
                        result.setdefault('warnings', []).insert(0, 'AI PASS FAILED: ' + str(exc))
                        result['ai_failed'] = True
                                    result['data']['Material Category'] = category
                                    insert_record({
                                      'id': uuid.uuid4().hex,
                                      'created': datetime.now().isoformat(
                                         timespec='microseconds'
                                       ),
                                       'result': result,
                                     })
                flash(f'Processed {name}.' + (' Gemini failed; Python results retained.' if result.get('ai_failed') else ''))
            except Exception as exc:
                flash(f'Extraction failed for {name}: {exc}', 'error')
    return redirect(url_for('index'))


@app.post('/edit/<record_id>')
def edit(record_id):
    record = get_record(record_id)
    if record is None:
        abort(404)

    result = record['result']
       category_index = list(result['data']).index('Material Category')
    selected_category = request.form.get(
        f'f_{category_index}', ''
    )

    if selected_category not in CATEGORIES:
        flash(
            'Please select a valid material category. '
            'No changes were saved.',
            'error',
        )
        return redirect(url_for('index'))
    for i, key in enumerate(result['data']):
        val = request.form.get(f'f_{i}', '').strip()
        if val and (key.startswith('__') or key in {HEADERS[j] for j in (15,17,19,21,40,41,42,43,110)}):
            try:
                num = float(val)
                if math.isfinite(num):
                    val = num
            except ValueError:
                pass  # Preserve user-entered text for review.
        if val != result['data'][key]:
            result.setdefault('evidence', []).append({'field': key, 'value': val, 'page': None,
                'raw': str(val), 'method': 'manual web review edit', 'status': 'user supplied'})
            result['data'][key] = val
    result['reviewed'] = True
    update_record(record_id, result)
    flash('Record saved to the database.')
    return redirect(url_for('index'))


@app.post('/remove/<record_id>')
def remove(record_id):
    record = get_record(record_id)
    if record is None:
        abort(404)

    delete_record(record_id)
    flash('Record removed from the database.')
    return redirect(url_for('index'))


@app.get('/download')
def download():
    records = user_records()

    with tempfile.TemporaryDirectory(
        prefix='epd_excel_'
    ) as directory:
        target = Path(directory) / 'EPD_Extracted_Results.xlsx'

        try:
            save_category_excel(records, target)
        except ValueError as exc:
            flash(str(exc), 'error')
            return redirect(url_for('index'))

        data = target.read_bytes()

    return send_file(
        io.BytesIO(data),
        mimetype=(
            'application/vnd.openxmlformats-officedocument.'
            'spreadsheetml.sheet'
        ),
        as_attachment=True,
        download_name='EPD_Extracted_Results.xlsx',
    )


@app.errorhandler(RequestEntityTooLarge)
def too_large(exc):
    flash('Upload too large. Select up to 5 PDFs, each under 15 MB.', 'error')
    return redirect(url_for('index'))

@app.errorhandler(DatabaseError)
def database_error(exc):
    return render_template_string(
        """
        <h1>Database connection needs attention</h1>
        <p>{{ message }}</p>
        <p>Check the Turso configuration, then reload the website.</p>
        """,
        message=str(exc),
    ), 503
@app.get('/health')
def health():
    return 'ok', 200


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', '10000')), debug=False)
