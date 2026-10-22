"""Match review fields and completion counts to the Excel columns."""

from extractor_core import HEADERS, GROUPS, STARTS, IMPACT_FIELDS


def excel_fields():
    """Return (display label, data key) in the 111-column Excel order."""
    fields = [(header, header) for header in HEADERS]

    for module, start in zip(GROUPS, STARTS):
        for offset, field in enumerate(IMPACT_FIELDS):
            fields[start - 1 + offset] = (
                f"{module} | {field}",
                f"__{module}__{field}",
            )

    fields[110] = (
        "A5 — Installation electricity (kWh per declared unit)",
        HEADERS[110],
    )

    return fields


def excel_values(data):
    """Return the extracted values in the original Excel column order."""
    return [data.get(key, "") for _, key in excel_fields()]


def is_filled(value):
    """Count zero as filled, while excluding None and blank text."""
    return value is not None and str(value).strip() != ""


def filled_count(data):
    """Count filled extraction columns; exclude session metadata."""
    return sum(is_filled(value) for value in excel_values(data))
