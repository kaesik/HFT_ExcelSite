from __future__ import annotations

from pathlib import Path
from threading import RLock
from typing import Any

from fastapi import HTTPException

from backend.config import MAX_RANGE_CELLS
from backend.services.cell_service import build_cell_payload
from backend.services.cell_service import get_display_cells

_sheet_cache_lock = RLock()
_formula_count_cache: dict[str, int] = {}
_sheet_summary_cache: dict[str, dict[str, Any]] = {}
_sheet_formulas_cache: dict[str, dict[str, Any]] = {}


def _get_sheet_cache_key(workbook_path: Path, sheet_name: str, suffix: str = "") -> str:
    try:
        modified_time = workbook_path.stat().st_mtime
    except OSError:
        modified_time = 0

    base_key = f"{workbook_path.resolve()}::{modified_time}::{sheet_name}"
    if suffix:
        return f"{base_key}::{suffix}"
    return base_key


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
    workbook_path = getattr(getattr(sheet, "parent", None), "_hft_workbook_path", None)

    if workbook_path is None:
        formula_count = 0
        for row_index in range(1, sheet.max_row + 1):
            for column_index in range(1, sheet.max_column + 1):
                cell = sheet.cell(row=row_index, column=column_index)
                if cell.data_type == "f":
                    formula_count += 1
        return formula_count

    cache_key = _get_sheet_cache_key(Path(workbook_path), sheet.title, "formula_count")

    with _sheet_cache_lock:
        cached = _formula_count_cache.get(cache_key)
        if cached is not None:
            return cached

    formula_count = 0

    for row_index in range(1, sheet.max_row + 1):
        for column_index in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=row_index, column=column_index)
            if cell.data_type == "f":
                formula_count += 1

    with _sheet_cache_lock:
        _formula_count_cache[cache_key] = formula_count

    return formula_count


def build_sheet_summary(workbook_path, sheet) -> dict:
    cache_key = _get_sheet_cache_key(workbook_path, sheet.title, "summary")

    with _sheet_cache_lock:
        cached = _sheet_summary_cache.get(cache_key)
        if cached is not None:
            return cached

    formula_cells_count = count_formula_cells(sheet)

    payload = {
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

    with _sheet_cache_lock:
        _sheet_summary_cache[cache_key] = payload

    return payload


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
            requested_coordinate = formula_sheet.cell(row=row_index, column=column_index).coordinate

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
                requested_coordinate=requested_coordinate,
                include_empty=include_empty,
            )

            if payload is None:
                continue

            payload["requestedRow"] = row_index
            payload["requestedColumn"] = column_index
            payload["requestedAddress"] = requested_coordinate
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
    cache_key = _get_sheet_cache_key(workbook_path, formula_sheet.title, f"formulas::{limit}")

    with _sheet_cache_lock:
        cached = _sheet_formulas_cache.get(cache_key)
        if cached is not None:
            return cached

    formula_cells_count = count_formula_cells(formula_sheet)

    if formula_cells_count == 0:
        payload = {
            "fileName": workbook_path.name,
            "sheet": formula_sheet.title,
            "limit": limit,
            "returnedCellsCount": 0,
            "truncated": False,
            "cells": [],
        }

        with _sheet_cache_lock:
            _sheet_formulas_cache[cache_key] = payload

        return payload

    cells = []

    for row_index in range(1, formula_sheet.max_row + 1):
        for column_index in range(1, formula_sheet.max_column + 1):
            requested_coordinate = formula_sheet.cell(row=row_index, column=column_index).coordinate

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
                requested_coordinate=requested_coordinate,
                include_empty=True,
            )

            if payload is None:
                continue

            payload["requestedRow"] = row_index
            payload["requestedColumn"] = column_index
            payload["requestedAddress"] = requested_coordinate
            payload["isMergedDisplayValue"] = payload["address"] != payload["requestedAddress"]

            cells.append(payload)

            if len(cells) >= limit:
                response = {
                    "fileName": workbook_path.name,
                    "sheet": formula_sheet.title,
                    "limit": limit,
                    "returnedCellsCount": len(cells),
                    "truncated": True,
                    "cells": cells,
                }

                with _sheet_cache_lock:
                    _sheet_formulas_cache[cache_key] = response

                return response

    response = {
        "fileName": workbook_path.name,
        "sheet": formula_sheet.title,
        "limit": limit,
        "returnedCellsCount": len(cells),
        "truncated": False,
        "cells": cells,
    }

    with _sheet_cache_lock:
        _sheet_formulas_cache[cache_key] = response

    return response