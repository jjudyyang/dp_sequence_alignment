from pathlib import Path
import re
import shutil
import sqlite3
import uuid

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

from alignment import process_excel


app = FastAPI(title="Shiftline")

BASE_DIR = Path(__file__).resolve().parent
FRONTEND_DIR = BASE_DIR / "frontend"
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
DATA_DIR = BASE_DIR / "data"
USAGE_DB = DATA_DIR / "usage.db"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
DATA_DIR.mkdir(exist_ok=True)

app.mount("/assets", StaticFiles(directory=FRONTEND_DIR / "assets"), name="assets")
app.mount("/data", StaticFiles(directory=FRONTEND_DIR / "data"), name="sample-data")


def get_usage_count():
    """Return how many successful alignments have finished."""
    if not USAGE_DB.exists():
        return 0
    conn = sqlite3.connect(USAGE_DB)
    try:
        row = conn.execute("SELECT runs FROM usage WHERE id = 1").fetchone()
        return int(row[0]) if row else 0
    finally:
        conn.close()


def record_successful_run():
    """Increment the persistent successful-run counter."""
    conn = sqlite3.connect(USAGE_DB)
    try:
        conn.execute(
            "CREATE TABLE IF NOT EXISTS usage "
            "(id INTEGER PRIMARY KEY CHECK (id = 1), "
            "runs INTEGER NOT NULL DEFAULT 0)"
        )
        conn.execute("INSERT OR IGNORE INTO usage (id, runs) VALUES (1, 0)")
        conn.execute("UPDATE usage SET runs = runs + 1 WHERE id = 1")
        conn.commit()
        row = conn.execute("SELECT runs FROM usage WHERE id = 1").fetchone()
        return int(row[0]) if row else 1
    finally:
        conn.close()


def shifted_download_filename(original_name: str) -> str:
    """
    Return a browser-safe download name based on the uploaded workbook name.
    """
    base = Path(original_name or "").name.strip()
    if not base or base in (".", ".."):
        return "workbook_shifted.xlsx"
    stem = Path(base).stem.strip() or "workbook"
    stem = re.sub(r'[\x00-\x1f\\/:*?"<>|]', "_", stem)[:180]
    return f"{stem}_shifted.xlsx"


@app.get("/", response_class=HTMLResponse)
def home():
    index_path = FRONTEND_DIR / "index.html"
    return HTMLResponse(index_path.read_text(encoding="utf-8"))


@app.get("/usage")
def usage():
    return {"runs": get_usage_count()}


@app.post("/process")
async def process_file(
    file: UploadFile = File(...),
    input_sheet_name: str = Form(""),
    left_sheet_name: str = Form(""),
    right_sheet_name: str = Form(""),
    output_sheet_name: str = Form("Aligned results"),
    start_row: int = Form(2),
    header_first_row: int = Form(1),
    header_last_row: int = Form(1),
    left_input_col: str = Form("D"),
    left_block_start_col: str = Form("A"),
    left_block_end_col: str = Form("D"),
    left_output_start_col: str = Form("A"),
    right_input_col: str = Form("F"),
    right_block_start_col: str = Form("F"),
    right_block_end_col: str = Form("H"),
    right_output_start_col: str = Form("F"),
    threshold: float = Form(0.5),
    diff_output_col: str = Form("E"),
):
    upload_id = str(uuid.uuid4())
    uploaded_name = Path(file.filename or "workbook.xlsx").name
    input_path = UPLOAD_DIR / f"{upload_id}_{uploaded_name}"
    output_path = OUTPUT_DIR / f"{upload_id}_output.xlsx"

    with input_path.open("wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        process_excel(
            input_file=input_path,
            output_file=output_path,
            input_sheet_name=input_sheet_name,
            left_sheet_name=left_sheet_name,
            right_sheet_name=right_sheet_name,
            output_sheet_name=output_sheet_name,
            start_row=start_row,
            header_first_row=header_first_row,
            header_last_row=header_last_row,
            left_input_col=left_input_col,
            left_block_start_col=left_block_start_col,
            left_block_end_col=left_block_end_col,
            left_output_start_col=left_output_start_col,
            right_input_col=right_input_col,
            right_block_start_col=right_block_start_col,
            right_block_end_col=right_block_end_col,
            right_output_start_col=right_output_start_col,
            threshold=threshold,
            diff_output_col=diff_output_col,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    runs_total = record_successful_run()
    return FileResponse(
        output_path,
        filename=shifted_download_filename(file.filename or ""),
        media_type=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        headers={"X-Successful-Runs": str(runs_total)},
    )
