from pathlib import Path

from fastapi import HTTPException

from backend.config import ALLOWED_EXTENSIONS
from backend.config import WORKBOOKS_DIR


def ensure_workbooks_dir_exists() -> None:
    if not WORKBOOKS_DIR.exists():
        raise FileNotFoundError(f"Workbooks directory not found: {WORKBOOKS_DIR}")


def list_workbook_files() -> list[Path]:
    ensure_workbooks_dir_exists()

    files = [
        path
        for path in WORKBOOKS_DIR.iterdir()
        if path.is_file()
        and path.suffix.lower() in ALLOWED_EXTENSIONS
        and not path.name.startswith("~$")
    ]

    files.sort(key=lambda item: item.name.lower())
    return files


def get_workbook_path(file_name: str) -> Path:
    ensure_workbooks_dir_exists()

    workbook_path = (WORKBOOKS_DIR / file_name).resolve()
    workbooks_dir_resolved = WORKBOOKS_DIR.resolve()

    if workbook_path.name.startswith("~$"):
        raise HTTPException(status_code=400, detail="Temporary Excel lock files are not supported")

    if workbook_path.suffix.lower() not in ALLOWED_EXTENSIONS:
        raise HTTPException(status_code=400, detail="Unsupported workbook extension")

    if workbooks_dir_resolved not in workbook_path.parents and workbook_path != workbooks_dir_resolved:
        raise HTTPException(status_code=400, detail="Invalid workbook path")

    if not workbook_path.exists() or not workbook_path.is_file():
        raise HTTPException(status_code=404, detail=f"Workbook '{file_name}' not found")

    return workbook_path


def validate_sheet_name(workbook, sheet_name: str) -> None:
    if sheet_name not in workbook.sheetnames:
        raise HTTPException(
            status_code=404,
            detail=f"Sheet '{sheet_name}' not found",
        )