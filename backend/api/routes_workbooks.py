from __future__ import annotations

from typing import Any

from fastapi import APIRouter
from fastapi import HTTPException
from fastapi import Query

from backend.config import WORKBOOKS_DIR
from backend.services.mapping_service import build_sheet_candidates
from backend.services.mapping_service import build_sheet_color_statistics
from backend.services.mapping_service import build_sheet_dropdown_diagnostics
from backend.services.mapping_service import build_sheet_style_diagnostics
from backend.services.mapping_service import build_sheet_colored_cells
from backend.services.sheet_service import build_sheet_formulas
from backend.services.sheet_service import build_sheet_range
from backend.services.sheet_service import build_sheet_summary
from backend.services.sheet_service import count_formula_cells
from backend.services.workbook_cache_service import workbook_cache_service
from backend.services.workbook_service import get_workbook_path
from backend.services.workbook_service import list_workbook_files
from backend.services.workbook_service import validate_sheet_name

router = APIRouter(prefix="/api", tags=["workbooks"])


@router.get("/health")
def health() -> dict[str, str]:
    return {
        "message": "HFT Excel localhost is running",
        "workbooksDirectory": str(WORKBOOKS_DIR),
    }


@router.get("/workbooks")
def get_workbooks() -> dict[str, Any]:
    try:
        workbooks = list_workbook_files()

        return {
            "count": len(workbooks),
            "workbooks": [
                {
                    "fileName": workbook.name,
                    "extension": workbook.suffix.lower(),
                }
                for workbook in workbooks
            ],
        }

    except FileNotFoundError as error:
        raise HTTPException(status_code=404, detail=str(error)) from error
    except Exception as error:
        raise HTTPException(status_code=500, detail=str(error)) from error


@router.get("/workbooks/{file_name}/summary")
def get_workbook_summary(file_name: str) -> dict[str, Any]:
    try:
        workbook_path = get_workbook_path(file_name)
        workbook = workbook_cache_service.get_formula_workbook(workbook_path)

        sheets = []
        for sheet in workbook.worksheets:
            formula_cells_count = count_formula_cells(sheet)

            sheets.append(
                {
                    "title": sheet.title,
                    "maxRow": sheet.max_row,
                    "maxColumn": sheet.max_column,
                    "dimensions": sheet.calculate_dimension(),
                    "hasFormulas": formula_cells_count > 0,
                    "formulaCellsCount": formula_cells_count,
                }
            )

        return {
            "fileName": workbook_path.name,
            "sheetCount": len(workbook.sheetnames),
            "sheetNames": workbook.sheetnames,
            "sheets": sheets,
        }

    except HTTPException:
        raise
    except Exception as error:
        raise HTTPException(status_code=500, detail=str(error)) from error


@router.get("/workbooks/{file_name}/sheet/{sheet_name}/summary")
def get_sheet_summary(file_name: str, sheet_name: str) -> dict[str, Any]:
    try:
        workbook_path = get_workbook_path(file_name)
        workbook = workbook_cache_service.get_formula_workbook(workbook_path)
        validate_sheet_name(workbook, sheet_name)

        sheet = workbook[sheet_name]
        return build_sheet_summary(workbook_path, sheet)

    except HTTPException:
        raise
    except Exception as error:
        raise HTTPException(status_code=500, detail=str(error)) from error


@router.get("/workbooks/{file_name}/sheet/{sheet_name}/range")
def get_sheet_range(
    file_name: str,
    sheet_name: str,
    start_row: int = Query(1, ge=1),
    end_row: int = Query(20, ge=1),
    start_column: int = Query(1, ge=1),
    end_column: int = Query(10, ge=1),
    include_empty: bool = Query(False),
) -> dict[str, Any]:
    try:
        workbook_path = get_workbook_path(file_name)
        workbook_with_formulas, workbook_with_values = workbook_cache_service.get_two_workbooks(workbook_path)
        validate_sheet_name(workbook_with_formulas, sheet_name)

        formula_sheet = workbook_with_formulas[sheet_name]
        value_sheet = workbook_with_values[sheet_name]

        return build_sheet_range(
            workbook_path=workbook_path,
            formula_sheet=formula_sheet,
            value_sheet=value_sheet,
            start_row=start_row,
            end_row=end_row,
            start_column=start_column,
            end_column=end_column,
            include_empty=include_empty,
        )

    except HTTPException:
        raise
    except Exception as error:
        raise HTTPException(status_code=500, detail=str(error)) from error


@router.get("/workbooks/{file_name}/sheet/{sheet_name}/formulas")
def get_sheet_formulas(
    file_name: str,
    sheet_name: str,
    limit: int = Query(100, ge=1, le=2000),
) -> dict[str, Any]:
    try:
        workbook_path = get_workbook_path(file_name)
        workbook_with_formulas, workbook_with_values = workbook_cache_service.get_two_workbooks(workbook_path)
        validate_sheet_name(workbook_with_formulas, sheet_name)

        formula_sheet = workbook_with_formulas[sheet_name]
        value_sheet = workbook_with_values[sheet_name]

        return build_sheet_formulas(
            workbook_path=workbook_path,
            formula_sheet=formula_sheet,
            value_sheet=value_sheet,
            limit=limit,
        )

    except HTTPException:
        raise
    except Exception as error:
        raise HTTPException(status_code=500, detail=str(error)) from error


@router.get("/workbooks/{file_name}/sheet/{sheet_name}/candidates")
def get_sheet_candidates(file_name: str, sheet_name: str) -> dict[str, Any]:
    try:
        workbook_path = get_workbook_path(file_name)
        workbook_with_formulas, workbook_with_values = workbook_cache_service.get_two_workbooks(workbook_path)
        validate_sheet_name(workbook_with_formulas, sheet_name)

        formula_sheet = workbook_with_formulas[sheet_name]
        value_sheet = workbook_with_values[sheet_name]

        return build_sheet_candidates(
            workbook_path=workbook_path,
            formula_sheet=formula_sheet,
            value_sheet=value_sheet,
        )

    except HTTPException:
        raise
    except Exception as error:
        raise HTTPException(status_code=500, detail=str(error)) from error


@router.get("/workbooks/{file_name}/sheet/{sheet_name}/colors")
def get_sheet_colors(file_name: str, sheet_name: str) -> dict[str, Any]:
    try:
        workbook_path = get_workbook_path(file_name)
        workbook_with_formulas, workbook_with_values = workbook_cache_service.get_two_workbooks(workbook_path)
        validate_sheet_name(workbook_with_formulas, sheet_name)

        formula_sheet = workbook_with_formulas[sheet_name]
        value_sheet = workbook_with_values[sheet_name]

        return build_sheet_color_statistics(
            workbook_path=workbook_path,
            formula_sheet=formula_sheet,
            value_sheet=value_sheet,
        )

    except HTTPException:
        raise
    except Exception as error:
        raise HTTPException(status_code=500, detail=str(error)) from error


@router.get("/workbooks/{file_name}/sheet/{sheet_name}/styles")
def get_sheet_styles(
    file_name: str,
    sheet_name: str,
    start_row: int = Query(1, ge=1),
    end_row: int = Query(80, ge=1),
    start_column: int = Query(1, ge=1),
    end_column: int = Query(30, ge=1),
) -> dict[str, Any]:
    try:
        workbook_path = get_workbook_path(file_name)
        workbook = workbook_cache_service.get_formula_workbook(workbook_path)
        validate_sheet_name(workbook, sheet_name)

        formula_sheet = workbook[sheet_name]

        return build_sheet_style_diagnostics(
            workbook_path=workbook_path,
            formula_sheet=formula_sheet,
            start_row=start_row,
            end_row=end_row,
            start_column=start_column,
            end_column=end_column,
        )

    except HTTPException:
        raise
    except Exception as error:
        raise HTTPException(status_code=500, detail=str(error)) from error


@router.get("/workbooks/{file_name}/sheet/{sheet_name}/dropdowns")
def get_sheet_dropdowns(file_name: str, sheet_name: str) -> dict[str, Any]:
    try:
        workbook_path = get_workbook_path(file_name)
        workbook_with_formulas, workbook_with_values = workbook_cache_service.get_two_workbooks(workbook_path)
        validate_sheet_name(workbook_with_formulas, sheet_name)

        formula_sheet = workbook_with_formulas[sheet_name]
        value_sheet = workbook_with_values[sheet_name]

        return build_sheet_dropdown_diagnostics(
            workbook_path=workbook_path,
            formula_sheet=formula_sheet,
            value_sheet=value_sheet,
        )

    except HTTPException:
        raise
    except Exception as error:
        raise HTTPException(status_code=500, detail=str(error)) from error


@router.get("/workbooks/{file_name}/sheet/{sheet_name}/colored-cells")
def get_sheet_colored_cells(file_name: str, sheet_name: str) -> dict[str, Any]:
    try:
        workbook_path = get_workbook_path(file_name)
        workbook_with_formulas, workbook_with_values = workbook_cache_service.get_two_workbooks(workbook_path)
        validate_sheet_name(workbook_with_formulas, sheet_name)

        formula_sheet = workbook_with_formulas[sheet_name]
        value_sheet = workbook_with_values[sheet_name]

        return build_sheet_colored_cells(
            workbook_path=workbook_path,
            formula_sheet=formula_sheet,
            value_sheet=value_sheet,
        )

    except HTTPException:
        raise
    except Exception as error:
        raise HTTPException(status_code=500, detail=str(error)) from error


@router.post("/cache/clear")
def clear_workbook_cache() -> dict[str, str]:
    workbook_cache_service.clear()
    return {"message": "Workbook cache cleared"}


@router.post("/workbooks/{file_name}/invalidate-cache")
def invalidate_workbook_cache(file_name: str) -> dict[str, str]:
    workbook_path = get_workbook_path(file_name)
    workbook_cache_service.invalidate(workbook_path)
    return {"message": f"Cache invalidated for '{file_name}'"}