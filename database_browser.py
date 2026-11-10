"""Browse saved EPD records without extracting another PDF."""

from datetime import datetime

from flask import (
    request, render_template_string, url_for,
    redirect, flash, abort,
)

from database import list_records, get_record, delete_record


PAGE = """
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Browse EPD database</title>
<style>
body {
    font-family: system-ui, sans-serif;
    margin: 0;
    padding: 24px;
    background: #f3f6fa;
    color: #192b3d;
}
main {
    max-width: 1500px;
    margin: auto;
}
.panel {
    background: white;
    padding: 20px;
    border-radius: 12px;
    margin: 18px 0;
}
.filters {
    display: flex;
    flex-wrap: wrap;
    gap: 14px;
    align-items: end;
}
label {
    display: grid;
    gap: 6px;
}
input, select, button, .button {
    font: inherit;
    padding: 9px;
    border: 1px solid #aebdcc;
    border-radius: 6px;
}
button, .button {
    cursor: pointer;
    background: #1469a3;
    color: white;
    text-decoration: none;
    display: inline-block;
}
.danger {
    background: #b42318;
}
.scroll {
    overflow: auto;
}
table {
    width: 100%;
    border-collapse: collapse;
}
th, td {
    text-align: left;
    padding: 12px;
    border-bottom: 1px solid #dce4ec;
    vertical-align: top;
}
th {
    background: #edf3f8;
}
.actions {
    display: flex;
    gap: 8px;
    flex-wrap: wrap;
}
.notice {
    padding: 12px;
    background: #fff4dc;
    border-radius: 6px;
}
small {
    color: #526579;
}
</style>
</head>
<body>
<main>
<a href="{{ url_for('index') }}">← Back to extractor</a>
<h1>Browse database</h1>

{% for kind, message in get_flashed_messages(with_categories=true) %}
<p class="notice" role="alert">{{ message }}</p>
{% endfor %}

<section class="panel">
<form method="get" action="{{ url_for('browse_database') }}">
<div class="filters">
    <label>
        Search saved data
        <input type="search" name="q" value="{{ filters.q }}"
               placeholder="EPD number, product, manufacturer…">
    </label>

    <label>
        Material category
        <select name="category">
            <option value="">All categories</option>
            {% for category in categories %}
            <option value="{{ category }}"
                {% if filters.category == category %}selected{% endif %}>
                {{ category }}
            </option>
            {% endfor %}
        </select>
    </label>

    <label>
        Publication from
        <input type="date" name="published_from"
               value="{{ filters.published_from }}">
    </label>

    <label>
        Publication to
        <input type="date" name="published_to"
               value="{{ filters.published_to }}">
    </label>

    <label>
        Valid-until from
        <input type="date" name="valid_from"
               value="{{ filters.valid_from }}">
    </label>

    <label>
        Valid-until to
        <input type="date" name="valid_to"
               value="{{ filters.valid_to }}">
    </label>

    <button type="submit">Apply filters</button>
    <a href="{{ url_for('browse_database') }}">Clear filters</a>
</div>
</form>
<p><small>
    Date ranges include both endpoints. Records with missing or
    unrecognised dates are excluded when that date filter is used.
    Dates written as DD/MM/YYYY are interpreted day first.
</small></p>
</section>

<section class="panel">
<p><strong>{{ rows|length }} matching records</strong></p>
<div class="scroll">
<table>
<thead>
<tr>
    <th>EPD number</th>
    <th>Product</th>
    <th>Category</th>
    <th>Manufacturer</th>
    <th>Publication date</th>
    <th>Valid until</th>
    <th>Actions</th>
</tr>
</thead>
<tbody>
{% for record in rows %}
{% set data = record['result']['data'] %}
<tr>
    <td>{{ data.get('EPD Number') or '—' }}</td>
    <td>{{ data.get('Product name') or '—' }}</td>
    <td>{{ data.get('Material Category') or '—' }}</td>
    <td>{{ data.get('Manufacturer/EPD owner') or '—' }}</td>
    <td>{{ data.get('Publication date') or '—' }}</td>
    <td>{{ data.get('Valid until') or '—' }}</td>
    <td>
        <div class="actions">
            <a class="button"
               href="{{ url_for('index',
                                record=record['id'],
                                _anchor='review-results') }}">
                Edit
            </a>
            <a class="button danger"
               href="{{ url_for('confirm_database_delete',
                                record_id=record['id']) }}">
                Delete
            </a>
        </div>
    </td>
</tr>
{% else %}
<tr><td colspan="7">No matching records found.</td></tr>
{% endfor %}
</tbody>
</table>
</div>
</section>
</main>
</body>
</html>
"""


DELETE_PAGE = """
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Confirm EPD deletion</title>
<style>
body {
    font-family: system-ui, sans-serif;
    max-width: 700px;
    margin: 40px auto;
    padding: 20px;
}
button {
    background: #b42318;
    color: white;
    padding: 12px 18px;
    border: 0;
    border-radius: 6px;
    cursor: pointer;
    font: inherit;
}
</style>
</head>
<body>
<h1>Delete this entry?</h1>
<p><strong>EPD number:</strong>
    {{ data.get('EPD Number') or 'Unavailable' }}</p>
<p><strong>Product:</strong>
    {{ data.get('Product name') or 'Unavailable' }}</p>
<p><strong>Category:</strong>
    {{ data.get('Material Category') or 'Unavailable' }}</p>
<p><strong>Source file:</strong>
    {{ record['result'].get('file') or 'Unavailable' }}</p>
<p><strong>This permanently removes the entry from the shared
database for all users. Undo cannot restore it.</strong></p>

<form method="post">
    <button type="submit" name="confirm" value="delete">
        Permanently delete entry
    </button>
</form>

<p><a href="{{ url_for('browse_database') }}">Cancel</a></p>
</body>
</html>
"""


def parse_saved_date(value):
    """Read common saved date formats; return None if unrecognised."""
    if value is None:
        return None

    text = str(value).strip()

    if not text or text == '-':
        return None

    try:
        return datetime.fromisoformat(
            text.replace('Z', '+00:00')
        ).date()
    except ValueError:
        pass

    for pattern in (
        '%d/%m/%Y',
        '%d-%m-%Y',
        '%d.%m.%Y',
        '%Y/%m/%d',
        '%d %B %Y',
        '%d %b %Y',
        '%B %d, %Y',
    ):
        try:
            return datetime.strptime(text, pattern).date()
        except ValueError:
            continue

    return None


def register_database_browser(app, categories):
    @app.get('/database')
    def browse_database():
        names = (
            'q', 'category',
            'published_from', 'published_to',
            'valid_from', 'valid_to',
        )

        filters = {
            name: request.args.get(name, '').strip()
            for name in names
        }

        dates = {}

        for name in names[2:]:
            raw = filters[name]

            if not raw:
                dates[name] = None
                continue

            try:
                dates[name] = datetime.strptime(
                    raw, '%Y-%m-%d'
                ).date()
            except ValueError:
                flash('Please enter valid filter dates.', 'error')
                return redirect(url_for('browse_database'))

        for start, end in (
            ('published_from', 'published_to'),
            ('valid_from', 'valid_to'),
        ):
            if dates[start] and dates[end] and dates[start] > dates[end]:
                flash(
                    'A From date cannot be later than its To date.',
                    'error',
                )
                return redirect(url_for('browse_database'))

        rows = []
        query = filters['q'].casefold()

        for record in list_records():
            data = record['result']['data']

            if (
                filters['category']
                and data.get('Material Category') != filters['category']
            ):
                continue

            searchable = ' '.join(
                str(value)
                for value in data.values()
                if value is not None
            )
            searchable += ' ' + str(
                record['result'].get('file') or ''
            )

            if query and query not in searchable.casefold():
                continue

            matches_dates = True

            for field, start, end in (
                ('Publication date', 'published_from', 'published_to'),
                ('Valid until', 'valid_from', 'valid_to'),
            ):
                lower = dates[start]
                upper = dates[end]

                if lower is None and upper is None:
                    continue

                saved = parse_saved_date(data.get(field))

                if (
                    saved is None
                    or (lower is not None and saved < lower)
                    or (upper is not None and saved > upper)
                ):
                    matches_dates = False
                    break

            if matches_dates:
                rows.append(record)

        return render_template_string(
            PAGE,
            rows=rows,
            categories=categories,
            filters=filters,
        )

    @app.route(
        '/database/delete/<record_id>',
        methods=['GET', 'POST'],
    )
    def confirm_database_delete(record_id):
        record = get_record(record_id)

        if record is None:
            flash('This entry no longer exists.', 'error')
            return redirect(url_for('browse_database'))

        if request.method == 'POST':
            if request.form.get('confirm') != 'delete':
                abort(400)

            delete_record(record_id)
            flash('Entry permanently deleted.')
            return redirect(url_for('browse_database'))

        return render_template_string(
            DELETE_PAGE,
            record=record,
            data=record['result']['data'],
        )
