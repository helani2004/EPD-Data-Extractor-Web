"""Render-ready Flask frontend for the existing EPD extraction engine."""
import os
import io
import math
import time
import uuid
import tempfile
import threading
from datetime import datetime, timezone
from itsdangerous import URLSafeSerializer, BadSignature
from pathlib import Path
from collections import OrderedDict

from flask import Flask, request, redirect, url_for, session, render_template_string, send_file, flash, abort
from werkzeug.exceptions import RequestEntityTooLarge
from extractor_core import process, save_excel, HEADERS, GROUPS, IMPACT_FIELDS
from ai_review import review_pdf
from category_export import save_category_excel
from material_categories import CATEGORIES
from record_fields import excel_fields, filled_count
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
#processing-indicator[hidden] {
    display: none;
}

#processing-indicator {
    position: fixed;
    inset: 0;
    z-index: 9999;
    display: flex;
    align-items: center;
    justify-content: center;
    background: rgba(21, 42, 67, 0.65);
    padding: 20px;
}

.processing-card {
    background: white;
    border-radius: 12px;
    padding: 30px;
    width: 100%;
    max-width: 420px;
    text-align: center;
    box-shadow: 0 10px 40px rgba(0, 0, 0, 0.2);
}

.processing-spinner {
    width: 48px;
    height: 48px;
    margin: 0 auto 18px;
    border: 5px solid #e0e7ef;
    border-top-color: #1469a3;
    border-radius: 50%;
    animation: processing-spin 0.9s linear infinite;
}

@keyframes processing-spin {
    to {
        transform: rotate(360deg);
    }
}

@media (prefers-reduced-motion: reduce) {
    .processing-spinner {
        animation: none;
    }
}</style></head><body><header><h1>📄 EPD Data Extractor</h1><p>Upload PDF → extract → review and edit → download 111-column Excel</p></header><main>
{% for kind,message in get_flashed_messages(with_categories=true) %}<div class="notice {{ 'error' if kind=='error' else '' }}">{{ message }}</div>{% endfor %}
<section class="panel"><div class="topline"><h2>1. Upload EPD PDFs</h2><span class="badge">Gemini: {{ 'Configured' if ai_ready else 'Not configured' }}</span></div>
<form id="upload-form" method="post"
      action="{{ url_for('upload') }}"
      enctype="multipart/form-data">
<input type="hidden" name="visit_token" id="visit-token">
<p id="session-status" class="status" aria-live="polite">
    Preparing your session…
</p>
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
<div class="actions"><button id="extract-button" type="submit" disabled>
    Extract uploaded PDFs
</button>
<a class="btn secondary" href="{{ url_for('index') }}">
    Clear / New entry
</a>

<button id="new-session-button" type="button" class="secondary">
    Start new session
</button>
<small>Processing can take several minutes. Please keep this tab open.</small></div></form></section>
{% if records %}<section class="panel"><div class="topline"><h2>2. Review extracted results ({{ records|length }} records)</h2></div>
<form method="get" action="{{ url_for('index') }}" class="actions">
    <label for="record">View a saved record</label>
    <select name="record" id="record" onchange="this.form.submit()">
        <option value="">Select a saved record</option>
        {% for r in records %}
        <option value="{{ r['id'] }}"
                {% if r['id'] == selected_id %}selected{% endif %}>
            {{ loop.index }}. {{ r['result'].get('file','EPD') }}
        </option>
        {% endfor %}
    </select>
    <noscript><button type="submit">Open</button></noscript>
</form>
{% if record %}
<p class="notice">
    <strong>{{ filled_columns }}/{{ total_columns }} columns filled</strong>
</p>
<p class="status">
    Edit any field and click <strong>Save edits</strong> before downloading.
    Blank fields mean unavailable or not recovered.
</p>
<form method="post" action="{{ url_for('edit',record_id=record['id']) }}"><div class="scroll"><table><thead><tr><th>Excel field</th><th>Extracted value (editable)</th></tr></thead><tbody>
{% for label,key in review_fields %}
{% set value = record['result']['data'].get(key, '') %}
<tr>
    <td>{{ label }}</td>
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
</tbody></table></div><div class="actions"><button type="submit">Save edits</button></div></form>
<form method="post" action="{{ url_for('remove',record_id=record['id']) }}" class="actions"><button class="danger" type="submit" onclick="return confirm('Remove this record?')">Remove selected record</button></form>
<details><summary>Extraction warnings ({{ record['result'].get('warnings',[])|length }})</summary>{% for w in record['result'].get('warnings',[]) %}<div class="warning">{{ w }}</div>{% endfor %}</details>
<details><summary>Source evidence ({{ record['result'].get('evidence',[])|length }})</summary><div class="scroll"><table><tr><th>Field</th><th>Value</th><th>Page</th><th>Source text</th></tr>{% for e in record['result'].get('evidence',[]) %}<tr><td>{{ e.get('field','') }}</td><td>{{ e.get('value','') }}</td><td>{{ e.get('page','') }}</td><td>{{ e.get('raw','') }}</td></tr>{% endfor %}</table></div></details>
<details><summary>Gemini AI review ({{ record['result'].get('ai_review',[])|length }})</summary><div class="scroll"><table><tr><th>Field</th><th>AI proposal</th><th>Decision</th><th>Page</th></tr>{% for e in record['result'].get('ai_review',[]) %}<tr><td>{{ e.get('field','') }}</td><td>{{ e.get('proposed','') }}</td><td>{{ e.get('decision','') }}</td><td>{{ e.get('page','') }}</td></tr>{% endfor %}</table></div></details>
{% endif %}</section>{% else %}<section class="panel"><p class="muted">Your extracted records will appear here.</p></section>{% endif %}
<section class="panel"><small>Records are stored in the shared project database. All website users can view, edit, export and remove saved records.</small></section>
<div class="actions">
    <form method="get" action="{{ url_for('download') }}">
        <input type="hidden" name="scope" value="session">
        <input type="hidden" name="visit_token"
               data-visit-token>

        <button id="session-download-button"
                type="submit" disabled>
            Download Excel — This session
        </button>
    </form>

    <a class="btn secondary"
       href="{{ url_for('download', scope='all') }}">
        Download Excel — All records
    </a>
</div>

<p class="status">
    This session includes records added during your current session.
    All records includes everything saved in the shared database.
    Save edits before downloading.
</p>
</main>

<div id="processing-indicator"
     role="status"
     aria-live="polite"
     hidden>
    <div class="processing-card">
        <div class="processing-spinner" aria-hidden="true"></div>
        <h2>Processing your PDFs…</h2>
        <p>Extracting data and running Gemini if selected.</p>
        <p class="status">
            This may take several minutes.
            Please keep this tab open.
        </p>
    </div>
</div>

<script>
(() => {
    const storageKey = 'epd-visit-v1';
    const uploadForm = document.getElementById('upload-form');
    const tokenInput = document.getElementById('visit-token');
    const extractButton = document.getElementById('extract-button');
    const newSessionButton = document.getElementById('new-session-button');
    const status = document.getElementById('session-status');

    let currentVisit = null;
    let processing = false;
    const processingIndicator =
        document.getElementById('processing-indicator');

    function showVisit(visit) {
        currentVisit = visit;
        tokenInput.value = visit.token;
        extractButton.disabled = false;

        const started = new Intl.DateTimeFormat('en-GB', {
            timeZone: 'Asia/Colombo',
            day: '2-digit',
            month: 'short',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit',
            hour12: false,
            timeZoneName: 'short'
        }).format(new Date(visit.started_at));

        status.textContent = 'Session started: ' + started;

        document.querySelectorAll('[data-visit-token]').forEach(input => {
            input.value = visit.token;
        });

        document.getElementById('session-download-button').disabled = false;
    }

    async function createVisit() {
        const response = await fetch(
            {{ url_for('start_session') | tojson }},
            {
                method: 'POST',
                credentials: 'same-origin',
                cache: 'no-store'
            }
        );

        if (!response.ok) {
            throw new Error('Could not start the session.');
        }

        const visit = await response.json();

        sessionStorage.setItem(storageKey, JSON.stringify(visit));
        showVisit(visit);
    }

    async function prepareVisit() {
        try {
            let stored = null;

            try {
                stored = JSON.parse(sessionStorage.getItem(storageKey));
            } catch {
                sessionStorage.removeItem(storageKey);
            }

            if (
                stored &&
                typeof stored.token === 'string' &&
                typeof stored.started_at === 'string' &&
                Number.isFinite(Date.parse(stored.started_at))
            ) {
                showVisit(stored);
            } else {
                await createVisit();
            }
        } catch (error) {
            currentVisit = null;
            extractButton.disabled = true;
            status.textContent =
                'Unable to prepare your session. Allow browser storage ' +
                'and reload the page, or click Start new session.';
        }
    }

    uploadForm.addEventListener('submit', event => {
        if (processing) {
            event.preventDefault();
            return;
        }

        if (!currentVisit) {
            event.preventDefault();
            status.textContent =
                'Please wait for your session to be ready.';
            return;
        }

        processing = true;
        processingIndicator.hidden = false;
        extractButton.disabled = true;
        extractButton.textContent = 'Processing…';
        newSessionButton.disabled = true;
        uploadForm.setAttribute('aria-busy', 'true');
    });

    window.addEventListener('pageshow', () => {
        processing = false;
        processingIndicator.hidden = true;
        extractButton.textContent = 'Extract uploaded PDFs';
        extractButton.disabled = !currentVisit;
        newSessionButton.disabled = false;
        uploadForm.removeAttribute('aria-busy');

        if (currentVisit) {
            showVisit(currentVisit);
        }
    });

    newSessionButton.addEventListener('click', async () => {
        newSessionButton.disabled = true;
        extractButton.disabled = true;
        currentVisit = null;
        status.textContent = 'Starting a new session…';
        document.getElementById('session-download-button').disabled = true;
        try {
            await createVisit();
            window.location.assign({{ url_for('index') | tojson }});
        } catch (error) {
            status.textContent =
                'Could not start a new session. Please try again.';
        } finally {
            newSessionButton.disabled = false;
        }
    });

    prepareVisit();
})();
</script>

</body></html>'''


def user_records():
    """Load saved records from the shared Turso database."""
    return list_records()

def visit_signer():
    """Sign visit details so the browser cannot alter them."""
    return URLSafeSerializer(
        app.secret_key,
        salt='epd-visit-v1',
    )


def load_visit(token):
    """Read and validate signed visit details."""
    if not token:
        return None

    try:
        visit = visit_signer().loads(token)
    except BadSignature:
        return None

    if not isinstance(visit, dict):
        return None

    if not isinstance(visit.get('id'), str):
        return None

    if not isinstance(visit.get('started_at'), str):
        return None

    return visit


@app.post('/session/start')
def start_session():
    """Create one identifier and start time for a new visit."""
    visit = {
        'id': uuid.uuid4().hex,
        'started_at': datetime.now(timezone.utc).isoformat(
            timespec='seconds'
        ),
    }

    return {
        'token': visit_signer().dumps(visit),
        'started_at': visit['started_at'],
    }
@app.get('/')
def index():
    records = user_records()
    selected_id = request.args.get('record', '')

    record = next(
        (r for r in records if r['id'] == selected_id),
        None,
    )

    return render_template_string(
        PAGE,
        records=records,
        record=record,
        selected_id=record['id'] if record else '',
        ai_ready=bool(API_KEY),
        categories=CATEGORIES,
        review_fields=excel_fields(),
        filled_columns=(
            filled_count(record['result']['data'])
            if record else 0
        ),
        total_columns=len(HEADERS),        
    )


@app.post('/upload')
def upload():
    visit = load_visit(request.form.get('visit_token', ''))

    if visit is None:
        flash(
            'Your session could not be identified. '
            'Reload the page and try again.',
            'error',
        )
        return redirect(url_for('index'))

    category = request.form.get('material_category', '')

    if category not in CATEGORIES:
        flash('Please select a valid material category.', 'error')
        return redirect(url_for('index'))

    uploads = [
        f for f in request.files.getlist('pdfs')
        if f.filename
    ]

    if not uploads:
        flash('Please select at least one PDF.', 'error')
        return redirect(url_for('index'))

    if len(uploads) > 5:
        flash('Please upload no more than 5 PDFs at a time.', 'error')
        return redirect(url_for('index'))
    last_saved_id = None
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
                        result = review_pdf(
                            path,
                            result,
                            HEADERS,
                            GROUPS,
                            IMPACT_FIELDS,
                            API_KEY,
                            MODEL,
                            Path(directory) / 'ai_cache',
                        )
                    except Exception as exc:
                        result.setdefault('warnings', []).insert(
                            0, 'AI PASS FAILED: ' + str(exc)
                        )
                        result['ai_failed'] = True

                result['data']['Material Category'] = category
                result['_visit'] = dict(visit)

                new_record_id = uuid.uuid4().hex
                insert_record({
                    'id': new_record_id,
                    'created': datetime.now().isoformat(
                        timespec='microseconds'
                    ),
                    'result': result,
                })
                last_saved_id = new_record_id
                message = f'Processed and saved {name}.'
                if result.get('ai_failed'):
                    message += ' Gemini failed; Python results retained.'
                flash(message)

            except Exception as exc:
                flash(f'Extraction or saving failed for {name}: {exc}', 'error')

    return redirect(
        url_for('index', record=last_saved_id)
        if last_saved_id
        else url_for('index')
    )


@app.post('/edit/<record_id>')
def edit(record_id):
    record = get_record(record_id)

    if record is None:
        abort(404)

    result = record['result']
    category_index = next(
        i for i, (_, key) in enumerate(excel_fields())
        if key == 'Material Category'
    )
    selected_category = request.form.get(f'f_{category_index}', '')

    if selected_category not in CATEGORIES:
        flash(
            'Please select a valid material category. No changes were saved.',
            'error',
        )
        return redirect(url_for('index'))

    numeric_fields = {
        HEADERS[j]
        for j in (15, 17, 19, 21, 40, 41, 42, 43, 110)
    }

    for i, (_, key) in enumerate(excel_fields()):
        val = request.form.get(f'f_{i}', '').strip()

        if val and (key.startswith('__') or key in numeric_fields):
            try:
                number = float(val)
                if math.isfinite(number):
                    val = number
            except ValueError:
                pass

        if val != result['data'].get(key, ''):
            result.setdefault('evidence', []).append({
                'field': key,
                'value': val,
                'page': None,
                'raw': str(val),
                'method': 'manual web review edit',
                'status': 'user supplied',
            })
            result['data'][key] = val

    result['reviewed'] = True
    update_record(record_id, result)
    flash('Record saved to the database.')
    return redirect(url_for('index', record=record_id))


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
    scope = request.args.get('scope', 'all')

    if scope not in ('all', 'session'):
        abort(400)

    visit = None

    if scope == 'session':
        visit = load_visit(
            request.args.get('visit_token', '')
        )

        if visit is None:
            flash(
                'Your session could not be identified. '
                'Reload the page and try again.',
                'error',
            )
            return redirect(url_for('index'))

    records = user_records()

    if visit is not None:
        records = [
            record
            for record in records
            if record['result'].get('_visit', {}).get('id')
            == visit['id']
        ]

        if not records:
            flash(
                'No saved records were found for this session. '
                'Extract a PDF first.',
                'error',
            )
            return redirect(url_for('index'))

    filename = (
        'EPD_This_Session.xlsx'
        if scope == 'session'
        else 'EPD_All_Records.xlsx'
    )

    with tempfile.TemporaryDirectory(
        prefix='epd_excel_'
    ) as directory:
        target = Path(directory) / filename

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
        download_name=filename,
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
