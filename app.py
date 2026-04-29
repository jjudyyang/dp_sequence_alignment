from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import HTMLResponse, FileResponse
from pathlib import Path
import shutil
import uuid

#dp program 
from version1 import process_excel 

app = FastAPI()

UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("outputs")

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


@app.get("/", response_class=HTMLResponse)
def home():
    return """
    <html>
        <head>
            <title>Excel Alignment Tool</title>
        </head>
        <body>
            <h1>Excel Alignment Tool</h1>

            <form action="/process" method="post" enctype="multipart/form-data">
                <p>
                    <label>Upload Excel file:</label><br>
                    <input type="file" name="file" required>
                </p>

                <p>
                    <label>Input sheet name:</label><br>
                    <input type="text" name="input_sheet_name" value="sheet1">
                </p>

                <p>
                    <label>Output sheet name:</label><br>
                    <input type="text" name="output_sheet_name" value="sheet 1 aligned result">
                </p>

                <p>
                    <label>Start row:</label><br>
                    <input type="number" name="start_row" value="2">
                </p>

                <p>
                    <label>Header row:</label><br>
                    <input type="number" name="header_row" value="1">
                </p>

                <h3>Left Side</h3>

                <p>
                    <label>Left matching column:</label><br>
                    <input type="text" name="left_input_col" value="D">
                </p>

                <p>
                    <label>Left block start column:</label><br>
                    <input type="text" name="left_block_start_col" value="A">
                </p>

                <p>
                    <label>Left block end column:</label><br>
                    <input type="text" name="left_block_end_col" value="D">
                </p>

                <p>
                    <label>Left output start column:</label><br>
                    <input type="text" name="left_output_start_col" value="A">
                </p>

                <h3>Right Side</h3>

                <p>
                    <label>Right matching column:</label><br>
                    <input type="text" name="right_input_col" value="F">
                </p>

                <p>
                    <label>Right block start column:</label><br>
                    <input type="text" name="right_block_start_col" value="F">
                </p>

                <p>
                    <label>Right block end column:</label><br>
                    <input type="text" name="right_block_end_col" value="H">
                </p>

                <p>
                    <label>Right output start column:</label><br>
                    <input type="text" name="right_output_start_col" value="F">
                </p>

                <h3>Matching Rule</h3>

                <p>
                    <label>Threshold:</label><br>
                    <input type="number" step="0.001" name="threshold" value="0.5">
                </p>

                <button type="submit">Process File</button>
            </form>
        </body>
    </html>
    """


@app.post("/process")
async def process_file(
    file: UploadFile = File(...),
    input_sheet_name: str = Form("sheet1"),
    output_sheet_name: str = Form("sheet 1 aligned result"),
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

    return FileResponse(
        output_path,
        filename="aligned_output.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )