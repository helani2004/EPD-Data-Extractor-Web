"""Store EPD records in Turso/libSQL over HTTPS."""

import json
import os
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


class DatabaseError(RuntimeError):
    """Database configuration or connection failed."""


def execute(sql, args=()):
    """Execute a parameterised SQL statement in Turso."""
    url = os.environ.get("TURSO_DATABASE_URL", "").strip()
    token = os.environ.get("TURSO_AUTH_TOKEN", "").strip()

    if not url or not token:
        raise DatabaseError(
            "Set TURSO_DATABASE_URL and TURSO_AUTH_TOKEN "
            "in Vercel, then redeploy."
        )

    if url.startswith("libsql://"):
        url = "https://" + url[len("libsql://"):]

    if not url.startswith("https://"):
        raise DatabaseError(
            "The database URL must start with libsql:// or https://."
        )

    body = {
        "requests": [
            {
                "type": "execute",
                "stmt": {
                    "sql": sql,
                    "args": [
                        {"type": "text", "value": str(value)}
                        for value in args
                    ],
                },
            },
            {"type": "close"},
        ]
    }

    request = Request(
        url.rstrip("/") + "/v2/pipeline",
        data=json.dumps(body).encode("utf-8"),
        headers={
            "Authorization": "Bearer " + token,
            "Content-Type": "application/json",
        },
        method="POST",
    )

    try:
        with urlopen(request, timeout=30) as response:
            payload = json.load(response)
    except HTTPError as exc:
        raise DatabaseError(
            f"Turso returned HTTP {exc.code}. "
            "Check the database URL, token and permissions."
        ) from exc
    except (URLError, TimeoutError) as exc:
        raise DatabaseError(
            "Cannot connect to Turso. Please try again."
        ) from exc

    results = payload.get("results", [])
    if not results or results[0].get("type") != "ok":
        raise DatabaseError(
            "Turso rejected the query. Check that epd_records "
            "exists and the token allows reads and writes."
        )

    return results[0]["response"]["result"]


def decode_record(row):
    """Convert a Turso result row into an application record."""
    return {
        "id": row[0]["value"],
        "created": row[1]["value"],
        "result": json.loads(row[2]["value"]),
    }


def list_records():
    """Return all records in the shared database."""
    result = execute(
        "SELECT id, created, result_json "
        "FROM epd_records ORDER BY created, id"
    )
    return [decode_record(row) for row in result["rows"]]


def get_record(record_id):
    """Return one record, or None if it does not exist."""
    result = execute(
        "SELECT id, created, result_json "
        "FROM epd_records WHERE id = ?",
        (record_id,),
    )

    if not result["rows"]:
        return None

    return decode_record(result["rows"][0])

class DuplicateRecordError(ValueError):
    """The PDF or EPD number already exists."""


def duplicate_guard(result, record_id):
    """Build a database check that excludes the record being edited."""
    data = result.get("data", {})
    raw_epd = data.get("EPD Number")
    epd = "" if raw_epd is None else str(raw_epd)

    # Ignore ASCII whitespace and letter case; retain punctuation.
    for character in (" ", "\t", "\r", "\n"):
        epd = epd.replace(character, "")

    epd = epd.translate(str.maketrans(
        "abcdefghijklmnopqrstuvwxyz",
        "ABCDEFGHIJKLMNOPQRSTUVWXYZ",
    ))

    if epd == "-":
        epd = ""

    fingerprint = result.get("_pdf_sha256") or ""

    stored_epd = (
        "upper(replace(replace(replace(replace("
        "coalesce(json_extract(result_json, '$.data.\"EPD Number\"'), ''), "
        "' ', ''), char(9), ''), char(13), ''), char(10), ''))"
    )

    sql = (
        "NOT EXISTS ("
        "SELECT 1 FROM epd_records "
        "WHERE id <> ? AND ("
        f"(? <> '' AND {stored_epd} = ?) "
        "OR (? <> '' AND "
        "json_extract(result_json, '$._pdf_sha256') = ?)"
        "))"
    )

    return sql, (record_id, epd, epd, fingerprint, fingerprint)
def insert_record(record):
    """Save a record only if its PDF and EPD number are unique."""
    guard, guard_args = duplicate_guard(
        record["result"], record["id"]
    )

    response = execute(
        "INSERT INTO epd_records (id, created, result_json) "
        "SELECT ?, ?, ? WHERE " + guard,
        (
            record["id"],
            record["created"],
            json.dumps(record["result"], ensure_ascii=False),
        ) + guard_args,
    )

    if response.get("affected_row_count", 0) == 0:
        raise DuplicateRecordError(
            "This PDF or EPD number already exists. "
            "Open the existing record to edit it."
        )

def update_record(record_id, result):
    """Update a record without duplicating another EPD."""
    guard, guard_args = duplicate_guard(result, record_id)

    response = execute(
        "UPDATE epd_records SET result_json = ? "
        "WHERE id = ? AND " + guard,
        (
            json.dumps(result, ensure_ascii=False),
            record_id,
        ) + guard_args,
    )

    if response.get("affected_row_count", 0) == 0:
        if get_record(record_id) is None:
            raise DatabaseError(
                "This record no longer exists. Reload the page."
            )

        raise DuplicateRecordError(
            "Another record already has this EPD number "
            "or PDF. No changes were saved."
        )

def delete_record(record_id):
    """Permanently remove one record."""
    execute(
        "DELETE FROM epd_records WHERE id = ?",
        (record_id,),
    )
