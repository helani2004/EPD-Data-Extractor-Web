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


def insert_record(record):
    """Save a newly extracted record."""
    execute(
        "INSERT INTO epd_records "
        "(id, created, result_json) VALUES (?, ?, ?)",
        (
            record["id"],
            record["created"],
            json.dumps(record["result"], ensure_ascii=False),
        ),
    )


def update_record(record_id, result):
    """Save changes to an existing record."""
    response = execute(
        "UPDATE epd_records SET result_json = ? WHERE id = ?",
        (
            json.dumps(result, ensure_ascii=False),
            record_id,
        ),
    )

    if response.get("affected_row_count", 0) == 0:
        raise DatabaseError(
            "This record no longer exists. Reload the page."
        )


def delete_record(record_id):
    """Permanently remove one record."""
    execute(
        "DELETE FROM epd_records WHERE id = ?",
        (record_id,),
    )
