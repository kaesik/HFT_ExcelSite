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
    fill = getattr(cell, "fill", None)

    if fill is None:
        return None

    if getattr(fill, "fill_type", None) is None:
        return None

    fg_color = getattr(fill, "fgColor", None)
    if fg_color is None:
        return None

    color_type = getattr(fg_color, "type", None)

    if color_type == "rgb":
        rgb = getattr(fg_color, "rgb", None)
        if isinstance(rgb, str):
            rgb_text = rgb.strip().upper()
            if rgb_text and rgb_text not in {"00000000", "000000", "000000FF"}:
                if len(rgb_text) == 8:
                    return f"#{rgb_text[2:]}"
                if len(rgb_text) == 6:
                    return f"#{rgb_text}"

    if color_type == "theme":
        theme = getattr(fg_color, "theme", None)
        tint = getattr(fg_color, "tint", None)

        if theme is not None:
            if tint is not None and float(tint) != 0:
                return f"theme:{theme}:tint:{tint}"
            return f"theme:{theme}"

    if color_type == "indexed":
        indexed = getattr(fg_color, "indexed", None)
        if isinstance(indexed, int):
            return f"indexed:{indexed}"

    rgb = getattr(fg_color, "rgb", None)
    if isinstance(rgb, str):
        rgb_text = rgb.strip().upper()
        if rgb_text and rgb_text not in {"00000000", "000000", "000000FF"}:
            if len(rgb_text) == 8:
                return f"#{rgb_text[2:]}"
            if len(rgb_text) == 6:
                return f"#{rgb_text}"

    theme = getattr(fg_color, "theme", None)
    if isinstance(theme, int):
        return f"theme:{theme}"

    indexed = getattr(fg_color, "indexed", None)
    if isinstance(indexed, int):
        return f"indexed:{indexed}"

    return None


def _coordinate_in_validation(validation, coordinate: str) -> bool:
    try:
        if coordinate in validation.cells:
            return True
    except Exception:
        pass

    try:
        for cell_range in validation.ranges:
            if coordinate in cell_range:
                return True
    except Exception:
        pass

    try:
        sqref = getattr(validation, "sqref", None)
        if sqref:
            for cell_range in sqref.ranges:
                if coordinate in cell_range:
                    return True
    except Exception:
        pass

    return False


def get_data_validation_info(sheet, requested_coordinate: str, source_coordinate: str | None = None) -> dict[str, Any] | None:
    data_validations = getattr(sheet, "data_validations", None)

    if not data_validations:
        return None

    validations = getattr(data_validations, "dataValidation", None)
    if not validations:
        return None

    coordinates_to_check = [requested_coordinate]
    if source_coordinate and source_coordinate not in coordinates_to_check:
        coordinates_to_check.append(source_coordinate)

    for validation in validations:
        for coordinate in coordinates_to_check:
            if _coordinate_in_validation(validation, coordinate):
                formula_1 = getattr(validation, "formula1", None)
                formula_2 = getattr(validation, "formula2", None)

                return {
                    "type": getattr(validation, "type", None),
                    "operator": getattr(validation, "operator", None),
                    "formula1": str(formula_1) if formula_1 is not None else None,
                    "formula2": str(formula_2) if formula_2 is not None else None,
                    "allowBlank": getattr(validation, "allowBlank", None),
                    "showDropDown": getattr(validation, "showDropDown", None),
                    "showInputMessage": getattr(validation, "showInputMessage", None),
                    "showErrorMessage": getattr(validation, "showErrorMessage", None),
                    "sqref": str(getattr(validation, "sqref", None)),
                }

    return None


def has_data_validation(sheet, requested_coordinate: str, source_coordinate: str | None = None) -> bool:
    validation_info = get_data_validation_info(
        sheet=sheet,
        requested_coordinate=requested_coordinate,
        source_coordinate=source_coordinate,
    )
    return validation_info is not None


def classify_cell(formula_cell, value_cell, sheet, requested_coordinate: str) -> dict[str, bool | str | None | dict[str, Any]]:
    fill_color = extract_fill_color(formula_cell)
    has_formula = formula_cell.data_type == "f"

    validation_info = get_data_validation_info(
        sheet=sheet,
        requested_coordinate=requested_coordinate,
        source_coordinate=formula_cell.coordinate,
    )

    has_dropdown = False
    if validation_info is not None:
        validation_type = validation_info.get("type")
        if validation_type == "list":
            has_dropdown = True

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
        "validationInfo": validation_info,
        "isInputCandidate": is_input_candidate,
        "isResultCandidate": is_result_candidate,
        "fillColor": fill_color,
    }


def build_cell_payload(
    formula_cell,
    value_cell,
    sheet,
    requested_coordinate: str,
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
        requested_coordinate=requested_coordinate,
    )

    return {
        "address": formula_cell.coordinate,
        "requestedAddress": requested_coordinate,
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
        "validationInfo": classification["validationInfo"],
        "isInputCandidate": classification["isInputCandidate"],
        "isResultCandidate": classification["isResultCandidate"],
    }