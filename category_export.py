"""Export Combined plus all 63 material category sheets."""

import re

from openpyxl import load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

from extractor_core import (
    TEMPLATE,
    HEADERS,
    GROUPS,
    STARTS,
    IMPACT_FIELDS,
)
from material_categories import CATEGORIES


def category_sheet_name(number, category):
    """Create a unique worksheet name within Excel's limits."""
    cleaned = re.sub(r"[\[\]:*?/\\]", "-", category)
    cleaned = cleaned.strip().strip("'")

    # The number uses three characters, leaving 28 for the name.
    return f"{number:02d} {cleaned[:28]}"


def append_values(sheet, row):
    """Write extracted text as values, never Excel formulas."""
    cleaned_row = [
        ILLEGAL_CHARACTERS_RE.sub("", value)
        if isinstance(value, str)
        else value
        for value in row
    ]

    sheet.append(cleaned_row)

    for cell, value in zip(sheet[sheet.max_row], cleaned_row):
        if isinstance(value, str):
            cell.data_type = "s"


def save_category_excel(records, destination):
    """Create exactly 64 sheets using the existing Excel template."""
    if len(CATEGORIES) != 63 or len(set(CATEGORIES)) != 63:
        raise ValueError("The category list must contain 63 unique items.")

    # Check categories before creating the workbook.
    for record in records:
        category = record["result"]["data"].get(
            "Material Category", ""
        )
        if category not in CATEGORIES:
            raise ValueError(
                "Every record must have a valid Material Category."
            )

    workbook = load_workbook(TEMPLATE)
    combined = workbook.active

    # Keep only the template's main sheet.
    for sheet in list(workbook):
        if sheet is not combined:
            workbook.remove(sheet)

    # Keep the two header rows; remove existing data.
    if combined.max_row >= 3:
        combined.delete_rows(3, combined.max_row - 2)

    combined.title = "Combined"

    # Create every category sheet before adding records.
    category_sheets = {}

    for number, category in enumerate(CATEGORIES, start=1):
        sheet = workbook.copy_worksheet(combined)
        sheet.title = category_sheet_name(number, category)
        category_sheets[category] = sheet

    for record in records:
        data = record["result"]["data"]
        category = data["Material Category"]

        # Preserve the existing 111-column order.
        row = [data.get(header, "") for header in HEADERS]

        # Preserve the existing lifecycle-module mapping.
        for module, start in zip(GROUPS, STARTS):
            for offset, field in enumerate(IMPACT_FIELDS):
                key = f"__{module}__{field}"
                row[start - 1 + offset] = data.get(key, "")

        append_values(combined, row)
        append_values(category_sheets[category], row)

    for sheet in workbook:
        sheet.freeze_panes = "C3"
        sheet.auto_filter.ref = (
            f"A2:DG{max(2, sheet.max_row)}"
        )

    workbook.save(destination)
    workbook.close()
