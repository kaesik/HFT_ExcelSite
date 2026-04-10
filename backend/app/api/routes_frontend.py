from pathlib import Path

from fastapi import APIRouter
from fastapi.responses import FileResponse

router = APIRouter(tags=["frontend"])

BASE_DIR = Path(__file__).resolve().parents[2]
STATIC_DIR = BASE_DIR / "static"


@router.get("/", response_class=FileResponse)
def root():
    return FileResponse(STATIC_DIR / "index.html")