from __future__ import annotations

from typing import Any

from app.color_rules import COLOR_ROLE_MAP
from app.services.cell_service import build_cell_payload
from app.services.cell_service import get_display_cells
from app.services.cell_service import get_data_validation_info


RESULT_COLUMNS = {"H", "I", "K", "X", "Y", "Z"}
RESULT_STATUS_VALUES = {"OK", "!!!", "<1,0", ">1,0", "< 1,0", "> 1,0"}


def is_text_label(value: Any) -> bool:
    if value is None:
        return False

    if not isinstance(value, str):
        return False

    text = value.strip()
    if not text:
        return False

    if text.startswith("="):
        return False

    return True


def find_nearest_label(sheet, row: int, column: int) -> str | None:
    for left_column in range(column - 1, max(column - 4, 0), -1):
        candidate = sheet.cell(row=row, column=left_column).value
        if is_text_label(candidate):
            return str(candidate).strip()

    for upper_row in range(row - 1, max(row - 3, 0), -1):
        candidate = sheet.cell(row=upper_row, column=column).value
        if is_text_label(candidate):
            return str(candidate).strip()

    return None


def get_column_letters(cell_address: str) -> str:
    letters = ""

    for character in cell_address:
        if character.isalpha():
            letters += character
        else:
            break

    return letters


def get_color_role(fill_color: str | None) -> str | None:
    if fill_color is None:
        return None

    return COLOR_ROLE_MAP.get(fill_color)


def detect_input_type(color_role: str | None, cell_payload: dict[str, Any]) -> str:
    if color_role == "dropdown" or cell_payload.get("hasDropdown"):
        return "select"

    value = cell_payload.get("cachedValue")
    if value is None:
        value = cell_payload.get("value")

    if isinstance(value, (int, float)):
        return "number"

    return "text"


def is_result_candidate(item: dict[str, Any]) -> bool:
    if not item.get("hasFormula"):
        return False

    column_letters = get_column_letters(item["address"])
    if column_letters in RESULT_COLUMNS:
        return True

    cached_value = item.get("cachedValue")
    if isinstance(cached_value, str) and cached_value.strip() in RESULT_STATUS_VALUES:
        return True

    label = item.get("label") or ""
    lowered_label = label.lower()

    important_fragments = [
        "nośność",
        "docisk",
        "rozerwanie",
        "stateczność",
        "przekrój netto",
        "punkt",
        "maksymalne wytężenie",
    ]

    return any(fragment in lowered_label for fragment in important_fragments)


def build_sheet_candidates(
    workbook_path,
    formula_sheet,
    value_sheet,
) -> dict[str, Any]:
    inputs: list[dict[str, Any]] = []
    results: list[dict[str, Any]] = []

    for row_index in range(1, formula_sheet.max_row + 1):
        for column_index in range(1, formula_sheet.max_column + 1):
            formula_cell, value_cell = get_display_cells(
                formula_sheet=formula_sheet,
                value_sheet=value_sheet,
                row=row_index,
                column=column_index,
            )

            requested_address = formula_sheet.cell(row=row_index, column=column_index).coordinate

            payload = build_cell_payload(
                formula_cell=formula_cell,
                value_cell=value_cell,
                sheet=formula_sheet,
                requested_coordinate=requested_address,
                include_empty=False,
            )

            if payload is None:
                continue

            label = find_nearest_label(
                sheet=formula_sheet,
                row=row_index,
                column=column_index,
            )

            color_role = get_color_role(payload.get("fillColor"))

            item = {
                "address": requested_address,
                "sourceAddress": payload["address"],
                "label": label,
                "value": payload["value"],
                "cachedValue": payload["cachedValue"],
                "formula": payload["formula"],
                "dataType": payload["dataType"],
                "numberFormat": payload["numberFormat"],
                "fillColor": payload["fillColor"],
                "colorRole": color_role,
                "hasDropdown": payload["hasDropdown"],
                "validationInfo": payload.get("validationInfo"),
                "hasFormula": payload["hasFormula"],
                "isMergedDisplayValue": payload.get("isMergedDisplayValue", False),
            }

            if color_role in {"input_1", "input_2", "dropdown"} or payload.get("hasDropdown"):
                item["inputType"] = detect_input_type(color_role, payload)
                inputs.append(item)
                continue

            if is_result_candidate(item):
                results.append(item)

    return {
        "fileName": workbook_path.name,
        "sheet": formula_sheet.title,
        "inputsCount": len(inputs),
        "resultsCount": len(results),
        "inputs": inputs,
        "results": results,
    }


def build_sheet_color_statistics(
    workbook_path,
    formula_sheet,
    value_sheet,
) -> dict[str, Any]:
    colors_map: dict[str, dict[str, Any]] = {}

    for row_index in range(1, formula_sheet.max_row + 1):
        for column_index in range(1, formula_sheet.max_column + 1):
            formula_cell, value_cell = get_display_cells(
                formula_sheet=formula_sheet,
                value_sheet=value_sheet,
                row=row_index,
                column=column_index,
            )

            requested_address = formula_sheet.cell(row=row_index, column=column_index).coordinate

            payload = build_cell_payload(
                formula_cell=formula_cell,
                value_cell=value_cell,
                sheet=formula_sheet,
                requested_coordinate=requested_address,
                include_empty=False,
            )

            if payload is None:
                continue

            color = payload.get("fillColor")
            if not color:
                continue

            if color not in colors_map:
                colors_map[color] = {
                    "count": 0,
                    "addresses": [],
                    "hasFormulaCount": 0,
                    "hasDropdownCount": 0,
                }

            colors_map[color]["count"] += 1

            if len(colors_map[color]["addresses"]) < 10:
                colors_map[color]["addresses"].append(requested_address)

            if payload.get("hasFormula"):
                colors_map[color]["hasFormulaCount"] += 1

            if payload.get("hasDropdown"):
                colors_map[color]["hasDropdownCount"] += 1

    colors = []
    for color, data in colors_map.items():
        colors.append(
            {
                "color": color,
                "role": get_color_role(color),
                "count": data["count"],
                "addresses": data["addresses"],
                "hasFormulaCount": data["hasFormulaCount"],
                "hasDropdownCount": data["hasDropdownCount"],
            }
        )

    colors.sort(key=lambda item: item["count"], reverse=True)

    return {
        "fileName": workbook_path.name,
        "sheet": formula_sheet.title,
        "colorsCount": len(colors),
        "colors": colors,
    }


def build_sheet_style_diagnostics(
    workbook_path,
    formula_sheet,
    start_row: int = 1,
    end_row: int = 80,
    start_column: int = 1,
    end_column: int = 30,
) -> dict[str, Any]:
    cells: list[dict[str, Any]] = []

    for row_index in range(start_row, min(end_row, formula_sheet.max_row) + 1):
        for column_index in range(start_column, min(end_column, formula_sheet.max_column) + 1):
            cell = formula_sheet.cell(row=row_index, column=column_index)

            if cell.value is None:
                continue

            fill = cell.fill
            fg_color = fill.fgColor if fill else None
            bg_color = fill.bgColor if fill else None

            validation_info = get_data_validation_info(
                sheet=formula_sheet,
                requested_coordinate=cell.coordinate,
                source_coordinate=cell.coordinate,
            )

            cells.append(
                {
                    "address": cell.coordinate,
                    "value": str(cell.value),
                    "fillType": getattr(fill, "fill_type", None),
                    "fgColorType": getattr(fg_color, "type", None) if fg_color else None,
                    "fgColorRgb": str(getattr(fg_color, "rgb", None)) if fg_color else None,
                    "fgColorIndexed": str(getattr(fg_color, "indexed", None)) if fg_color else None,
                    "fgColorTheme": str(getattr(fg_color, "theme", None)) if fg_color else None,
                    "fgColorTint": str(getattr(fg_color, "tint", None)) if fg_color else None,
                    "bgColorType": getattr(bg_color, "type", None) if bg_color else None,
                    "bgColorRgb": str(getattr(bg_color, "rgb", None)) if bg_color else None,
                    "bgColorIndexed": str(getattr(bg_color, "indexed", None)) if bg_color else None,
                    "bgColorTheme": str(getattr(bg_color, "theme", None)) if bg_color else None,
                    "bgColorTint": str(getattr(bg_color, "tint", None)) if bg_color else None,
                    "hasFormula": cell.data_type == "f",
                    "hasDropdown": validation_info is not None and validation_info.get("type") == "list",
                    "validationInfo": validation_info,
                }
            )

    return {
        "fileName": workbook_path.name,
        "sheet": formula_sheet.title,
        "range": {
            "startRow": start_row,
            "endRow": min(end_row, formula_sheet.max_row),
            "startColumn": start_column,
            "endColumn": min(end_column, formula_sheet.max_column),
        },
        "cellsCount": len(cells),
        "cells": cells,
    }


def build_sheet_dropdown_diagnostics(
    workbook_path,
    formula_sheet,
) -> dict[str, Any]:
    dropdowns: list[dict[str, Any]] = []

    data_validations = getattr(formula_sheet, "data_validations", None)
    validations = getattr(data_validations, "dataValidation", None) if data_validations else None

    if validations:
        for validation in validations:
            if getattr(validation, "type", None) != "list":
                continue

            dropdowns.append(
                {
                    "type": getattr(validation, "type", None),
                    "formula1": str(getattr(validation, "formula1", None)),
                    "formula2": str(getattr(validation, "formula2", None)),
                    "allowBlank": getattr(validation, "allowBlank", None),
                    "showDropDown": getattr(validation, "showDropDown", None),
                    "showInputMessage": getattr(validation, "showInputMessage", None),
                    "showErrorMessage": getattr(validation, "showErrorMessage", None),
                    "sqref": str(getattr(validation, "sqref", None)),
                }
            )

    return {
        "fileName": workbook_path.name,
        "sheet": formula_sheet.title,
        "dropdownsCount": len(dropdowns),
        "dropdowns": dropdowns,
    }