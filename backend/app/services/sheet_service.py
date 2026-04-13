from fastapi import HTTPException

from app.config import MAX_RANGE_CELLS
from app.services.cell_service import build_cell_payload
from app.services.cell_service import get_display_cells


def clamp_range(
    sheet,
    start_row: int,
    end_row: int,
    start_column: int,
    end_column: int,
) -> tuple[int, int, int, int]:
    if start_row < 1:
        start_row = 1

    if start_column < 1:
        start_column = 1

    if end_row < start_row:
        end_row = start_row

    if end_column < start_column:
        end_column = start_column

    if sheet.max_row > 0:
        end_row = min(end_row, sheet.max_row)

    if sheet.max_column > 0:
        end_column = min(end_column, sheet.max_column)

    return start_row, end_row, start_column, end_column


def count_formula_cells(sheet) -> int:
    formula_count = 0

    for row_index in range(1, sheet.max_row + 1):
        for column_index in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=row_index, column=column_index)
            if cell.data_type == "f":
                formula_count += 1

    return formula_count


def build_sheet_summary(workbook_path, sheet) -> dict:
    formula_cells_count = count_formula_cells(sheet)

    return {
        "fileName": workbook_path.name,
        "sheet": {
            "title": sheet.title,
            "maxRow": sheet.max_row,
            "maxColumn": sheet.max_column,
            "dimensions": sheet.calculate_dimension(),
            "hasFormulas": formula_cells_count > 0,
            "formulaCellsCount": formula_cells_count,
        },
    }


def build_sheet_range(
    workbook_path,
    formula_sheet,
    value_sheet,
    start_row: int,
    end_row: int,
    start_column: int,
    end_column: int,
    include_empty: bool,
) -> dict:
    formula_cells_count = count_formula_cells(formula_sheet)

    if formula_cells_count == 0:
        raise HTTPException(
            status_code=400,
            detail=f"Sheet '{formula_sheet.title}' does not contain formulas",
        )

    start_row, end_row, start_column, end_column = clamp_range(
        sheet=formula_sheet,
        start_row=start_row,
        end_row=end_row,
        start_column=start_column,
        end_column=end_column,
    )

    requested_cells_count = (end_row - start_row + 1) * (end_column - start_column + 1)

    if requested_cells_count > MAX_RANGE_CELLS:
        raise HTTPException(
            status_code=400,
            detail=(
                f"Requested range contains {requested_cells_count} cells. "
                f"Maximum allowed is {MAX_RANGE_CELLS}."
            ),
        )

    cells = []

    for row_index in range(start_row, end_row + 1):
        for column_index in range(start_column, end_column + 1):
            formula_cell, value_cell = get_display_cells(
                formula_sheet=formula_sheet,
                value_sheet=value_sheet,
                row=row_index,
                column=column_index,
            )

            payload = build_cell_payload(
                formula_cell=formula_cell,
                value_cell=value_cell,
                sheet=formula_sheet,
                requested_coordinate=formula_sheet.cell(row=row_index, column=column_index).coordinate,
                include_empty=include_empty,
            )

            if payload is None:
                continue

            payload["requestedRow"] = row_index
            payload["requestedColumn"] = column_index
            payload["requestedAddress"] = formula_sheet.cell(row=row_index, column=column_index).coordinate
            payload["isMergedDisplayValue"] = payload["address"] != payload["requestedAddress"]

            cells.append(payload)

    return {
        "fileName": workbook_path.name,
        "sheet": formula_sheet.title,
        "range": {
            "startRow": start_row,
            "endRow": end_row,
            "startColumn": start_column,
            "endColumn": end_column,
            "requestedCellsCount": requested_cells_count,
            "returnedCellsCount": len(cells),
        },
        "cells": cells,
    }


def build_sheet_formulas(
    workbook_path,
    formula_sheet,
    value_sheet,
    limit: int,
) -> dict:
    formula_cells_count = count_formula_cells(formula_sheet)

    if formula_cells_count == 0:
        return {
            "fileName": workbook_path.name,
            "sheet": formula_sheet.title,
            "limit": limit,
            "returnedCellsCount": 0,
            "truncated": False,
            "cells": [],
        }

    cells = []

    for row_index in range(1, formula_sheet.max_row + 1):
        for column_index in range(1, formula_sheet.max_column + 1):
            formula_cell, value_cell = get_display_cells(
                formula_sheet=formula_sheet,
                value_sheet=value_sheet,
                row=row_index,
                column=column_index,
            )

            if formula_cell.data_type != "f":
                continue

            payload = build_cell_payload(
                formula_cell=formula_cell,
                value_cell=value_cell,
                sheet=formula_sheet,
                requested_coordinate=formula_sheet.cell(row=row_index, column=column_index).coordinate,
                include_empty=True,
            )

            if payload is None:
                continue

            payload["requestedRow"] = row_index
            payload["requestedColumn"] = column_index
            payload["requestedAddress"] = formula_sheet.cell(row=row_index, column=column_index).coordinate
            payload["isMergedDisplayValue"] = payload["address"] != payload["requestedAddress"]

            cells.append(payload)

            if len(cells) >= limit:
                return {
                    "fileName": workbook_path.name,
                    "sheet": formula_sheet.title,
                    "limit": limit,
                    "returnedCellsCount": len(cells),
                    "truncated": True,
                    "cells": cells,
                }

    return {
        "fileName": workbook_path.name,
        "sheet": formula_sheet.title,
        "limit": limit,
        "returnedCellsCount": len(cells),
        "truncated": False,
        "cells": cells,
    }