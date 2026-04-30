from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, FileResponse
from pathlib import Path
import shutil
import uuid

from version1 import process_excel

app = FastAPI(title="Excel Alignment")

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

HOME_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Align Excel</title>
  <style>
    :root { --border: #ddd; --btn: #2563eb; }
    * { box-sizing: border-box; }
    body { font-family: system-ui, sans-serif; font-size: 16px; margin: 0; padding: 20px; max-width: 520px; }
    h1 { font-size: 1.25rem; margin: 0 0 16px; font-weight: 600; }
    form { display: flex; flex-direction: column; gap: 14px; }
    label { display: block; font-size: 0.8rem; color: #444; margin-bottom: 4px; }
    input[type="text"], input[type="number"], input[type="file"] {
      width: 100%; padding: 8px; border: 1px solid var(--border); border-radius: 6px; font-size: 15px;
    }
    .inline2 { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
    .block-title { font-size: 0.85rem; font-weight: 600; margin: 4px 0 0; color: #222; }
    .col-row {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 8px;
      align-items: end;
    }
    .col-row .cell { min-width: 0; }
    .col-row input { text-align: center; }
    button[type="submit"] {
      margin-top: 6px; padding: 12px; font-size: 1rem; font-weight: 600;
      background: var(--btn); color: #fff; border: none; border-radius: 8px; cursor: pointer;
    }
    button[type="submit"]:hover { filter: brightness(1.06); }
  </style>
</head>
<body>
  <h1>Align spreadsheet</h1>
  <form action="/process" method="post" enctype="multipart/form-data">
    <div>
      <label for="file">File</label>
      <input id="file" type="file" name="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" required />
    </div>
    <div class="inline2">
      <div>
        <label for="input_sheet_name">Sheet in</label>
        <input id="input_sheet_name" type="text" name="input_sheet_name" value="sheet1" autocomplete="off" />
      </div>
      <div>
        <label for="output_sheet_name">Sheet out</label>
        <input id="output_sheet_name" type="text" name="output_sheet_name" value="Aligned results" autocomplete="off" />
      </div>
    </div>
    <div class="inline2">
      <div>
        <label for="header_row">Header row</label>
        <input id="header_row" type="number" name="header_row" value="1" min="1" />
      </div>
      <div>
        <label for="start_row">Data row</label>
        <input id="start_row" type="number" name="start_row" value="2" min="1" />
      </div>
    </div>

    <p class="block-title">Left</p>
    <div class="col-row">
      <div class="cell">
        <label for="left_block_start_col">From</label>
        <input id="left_block_start_col" type="text" name="left_block_start_col" value="A" maxlength="3" />
      </div>
      <div class="cell">
        <label for="left_block_end_col">To</label>
        <input id="left_block_end_col" type="text" name="left_block_end_col" value="D" maxlength="3" />
      </div>
      <div class="cell">
        <label for="left_input_col">Match</label>
        <input id="left_input_col" type="text" name="left_input_col" value="D" maxlength="3" />
      </div>
      <div class="cell">
        <label for="left_output_start_col">Out</label>
        <input id="left_output_start_col" type="text" name="left_output_start_col" value="A" maxlength="3" />
      </div>
    </div>

    <p class="block-title">Right</p>
    <div class="col-row">
      <div class="cell">
        <label for="right_block_start_col">From</label>
        <input id="right_block_start_col" type="text" name="right_block_start_col" value="F" maxlength="3" />
      </div>
      <div class="cell">
        <label for="right_block_end_col">To</label>
        <input id="right_block_end_col" type="text" name="right_block_end_col" value="H" maxlength="3" />
      </div>
      <div class="cell">
        <label for="right_input_col">Match</label>
        <input id="right_input_col" type="text" name="right_input_col" value="F" maxlength="3" />
      </div>
      <div class="cell">
        <label for="right_output_start_col">Out</label>
        <input id="right_output_start_col" type="text" name="right_output_start_col" value="F" maxlength="3" />
      </div>
    </div>

    <div>
      <label for="threshold">Max diff</label>
      <input id="threshold" type="number" step="0.001" name="threshold" value="0.5" />
    </div>

    <button type="submit">Download</button>
  </form>
</body>
</html>
"""


@app.get("/", response_class=HTMLResponse)
def home():
    return HTMLResponse(content=HOME_HTML)


@app.post("/process")
async def process_file(
    file: UploadFile = File(...),
    input_sheet_name: str = Form("sheet1"),
    output_sheet_name: str = Form("Aligned results"),
    start_row: int = Form(2),
    header_row: int = Form(1),
    left_input_col: str = Form("D"),
    left_block_start_col: str = Form("A"),
    left_block_end_col: str = Form("D"),
    left_output_start_col: str = Form("A"),
    right_input_col: str = Form("F"),
    right_block_start_col: str = Form("F"),
    right_block_end_col: str = Form("H"),
    right_output_start_col: str = Form("F"),
    threshold: float = Form(0.5),
):
    upload_id = str(uuid.uuid4())

    input_path = UPLOAD_DIR / f"{upload_id}_{file.filename}"
    output_path = OUTPUT_DIR / f"{upload_id}_output.xlsx"

    with input_path.open("wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        process_excel(
            input_file=input_path,
            output_file=output_path,
            input_sheet_name=input_sheet_name,
            output_sheet_name=output_sheet_name,
            start_row=start_row,
            header_row=header_row,
            left_input_col=left_input_col,
            left_block_start_col=left_block_start_col,
            left_block_end_col=left_block_end_col,
            left_output_start_col=left_output_start_col,
            right_input_col=right_input_col,
            right_block_start_col=right_block_start_col,
            right_block_end_col=right_block_end_col,
            right_output_start_col=right_output_start_col,
            threshold=threshold,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    return FileResponse(
        output_path,
        filename="aligned_output.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
