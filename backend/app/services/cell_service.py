from __future__ import annotations

from datetime import date
from datetime import datetime
from datetime import time
from pathlib import Path
from typing import Any


def normalize_value(value: Any) -> Any:
    if value is None:
        return None

    if isinstance(value, datetime):
        return value.isoformat()

    if isinstance(value, date):
        return value.isoformat()

    if isinstance(value, time):
        return value.isoformat()

    if isinstance(value, Path):
        return str(value)

    return value


def get_merged_parent_cell(sheet, row: int, column: int):
    cell = sheet.cell(row=row, column=column)

    for merged_range in sheet.merged_cells.ranges:
        if cell.coordinate in merged_range:
            return sheet.cell(row=merged_range.min_row, column=merged_range.min_col)

    return cell


def get_display_cells(formula_sheet, value_sheet, row: int, column: int):
    formula_cell = get_merged_parent_cell(formula_sheet, row, column)
    value_cell = get_merged_parent_cell(value_sheet, row, column)
    return formula_cell, value_cell


def extract_fill_color(cell) -> str | None:
    fill = cell.fill

    if fill is None or fill.fill_type is None:
        return None

    fg_color = fill.fgColor

    if fg_color is None:
        return None

    color_type = getattr(fg_color, "type", None)

    rgb = getattr(fg_color, "rgb", None)
    if rgb is not None:
        rgb_text = str(rgb).strip()
        if rgb_text and rgb_text not in {"00000000", "000000", "000000FF"}:
            if len(rgb_text) == 8:
                return f"#{rgb_text[2:]}"
            if len(rgb_text) == 6:
                return f"#{rgb_text}"

    indexed = getattr(fg_color, "indexed", None)
    if indexed is not None:
        return f"indexed:{indexed}"

    theme = getattr(fg_color, "theme", None)
    tint = getattr(fg_color, "tint", None)
    if theme is not None:
        if tint is not None:
            return f"theme:{theme}:tint:{tint}"
        return f"theme:{theme}"

    if color_type:
        return f"type:{color_type}"

    return None


def has_data_validation(sheet, coordinate: str) -> bool:
    data_validations = getattr(sheet, "data_validations", None)

    if not data_validations:
        return False

    validations = getattr(data_validations, "dataValidation", None)
    if not validations:
        return False

    for validation in validations:
        try:
            if coordinate in validation.cells:
                return True
        except Exception:
            pass

        try:
            for cell_range in validation.ranges.ranges:
                if coordinate in cell_range:
                    return True
        except Exception:
            pass

    return False


def classify_cell(formula_cell, value_cell, sheet) -> dict[str, bool | str | None]:
    fill_color = extract_fill_color(formula_cell)
    has_formula = formula_cell.data_type == "f"
    has_dropdown = has_data_validation(sheet, formula_cell.coordinate)

    formula_value = formula_cell.value
    cached_value = value_cell.value

    is_input_candidate = (
        not has_formula
        and (formula_value is not None or cached_value is not None)
        and (has_dropdown or fill_color is not None)
    )

    is_result_candidate = has_formula

    return {
        "hasFormula": has_formula,
        "hasDropdown": has_dropdown,
        "isInputCandidate": is_input_candidate,
        "isResultCandidate": is_result_candidate,
        "fillColor": fill_color,
    }


def build_cell_payload(
    formula_cell,
    value_cell,
    sheet,
    include_empty: bool = False,
) -> dict[str, Any] | None:
    formula_value = formula_cell.value
    cached_value = value_cell.value

    if not include_empty and formula_value is None and cached_value is None:
        return None

    classification = classify_cell(
        formula_cell=formula_cell,
        value_cell=value_cell,
        sheet=sheet,
    )

    return {
        "address": formula_cell.coordinate,
        "row": formula_cell.row,
        "column": formula_cell.column,
        "value": normalize_value(formula_value),
        "formula": formula_value if formula_cell.data_type == "f" else None,
        "cachedValue": normalize_value(cached_value),
        "dataType": formula_cell.data_type,
        "numberFormat": formula_cell.number_format,
        "fillColor": classification["fillColor"],
        "hasFormula": classification["hasFormula"],
        "hasDropdown": classification["hasDropdown"],
        "isInputCandidate": classification["isInputCandidate"],
        "isResultCandidate": classification["isResultCandidate"],
    }