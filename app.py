from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, FileResponse
from pathlib import Path
import re
import shutil
import uuid

from version1 import process_excel

app = FastAPI(title="Excel Alignment")

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


def shifted_download_filename(original_name: str) -> str:
    """
    Browser download name: input basename stem + '_shifted.xlsx'.
    Strips directories and characters that confuse Content-Disposition.
    """
    base = Path(original_name or "").name.strip()
    if not base or base in (".", ".."):
        return "workbook_shifted.xlsx"
    stem = Path(base).stem.strip() or "workbook"
    stem = re.sub(r'[\x00-\x1f\\/:*?"<>|]', "_", stem)[:180]
    return f"{stem}_shifted.xlsx"


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
    body { font-family: system-ui, sans-serif; font-size: 16px; margin: 0; padding: 20px; max-width: 760px; }
    h1 { font-size: 1.25rem; margin: 0 0 16px; font-weight: 600; }
    form { display: flex; flex-direction: column; gap: 14px; }
    label { display: block; font-size: 0.8rem; color: #444; margin-bottom: 4px; }
    input[type="text"], input[type="number"], input[type="file"] {
      width: 100%; padding: 8px; border: 1px solid var(--border); border-radius: 6px; font-size: 15px;
    }
    input::placeholder { color: #9ca3af; }
    .inline2 { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; align-items: end; }
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
    button[type="submit"][disabled] { opacity: 0.8; cursor: wait; }
    .submit-msg { font-size: 0.9rem; color: #374151; min-height: 1.2em; margin-top: 4px; }
    .preview-msg { font-size: 0.8rem; color: #555; margin: 4px 0 0; min-height: 1.25em; }
    .preview-holder {
      overflow: auto; max-height: 300px; border: 1px solid var(--border); border-radius: 6px;
      margin-top: 6px; background: #fafafa;
    }
    .preview-holder table { border-collapse: collapse; font-size: 11px; font-family: ui-monospace, monospace; }
    .preview-holder th,
    .preview-holder td {
      border: 1px solid #e8e8e8; padding: 3px 7px;
      white-space: nowrap; max-width: 120px; overflow: hidden; text-overflow: ellipsis;
    }
    .preview-holder th { background: #e5e7eb; font-weight: 600; text-align: center; }
    .preview-holder .row-num { background: #f3f4f6; text-align: right; }
    .preview-holder tr:nth-child(even) td { background: #f0f0f0; }
  </style>
</head>
<body>
  <h1>Align spreadsheet</h1>
  <form id="align-form" action="/process" method="post" enctype="multipart/form-data">
    <div>
      <label for="file">File</label>
      <input id="file" type="file" name="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" required />
      <p id="preview-msg" class="preview-msg"></p>
      <div id="preview-holder" class="preview-holder"></div>
    </div>
    <div class="inline2">
      <div>
        <label for="input_sheet_name">Sheet in</label>
        <input id="input_sheet_name" type="text" name="input_sheet_name" placeholder="sheet1" autocomplete="off" />
      </div>
      <div>
        <label for="output_sheet_name">Sheet out</label>
        <input id="output_sheet_name" type="text" name="output_sheet_name" placeholder="Aligned results" autocomplete="off" />
      </div>
    </div>
    <div class="inline2">
      <div>
        <label for="header_first_row">Header from</label>
        <input id="header_first_row" type="number" name="header_first_row" placeholder="1" min="1" />
      </div>
      <div>
        <label for="header_last_row">Header to</label>
        <input id="header_last_row" type="number" name="header_last_row" placeholder="1" min="1" />
      </div>
    </div>
    <div>
      <label for="start_row">Data row — first row with numbers to match</label>
      <input id="start_row" type="number" name="start_row" placeholder="2" min="1" />
    </div>

    <p class="block-title">Left</p>
    <div class="col-row">
      <div class="cell">
        <label for="left_block_start_col">From</label>
        <input id="left_block_start_col" type="text" name="left_block_start_col" placeholder="A" maxlength="3" />
      </div>
      <div class="cell">
        <label for="left_block_end_col">To</label>
        <input id="left_block_end_col" type="text" name="left_block_end_col" placeholder="D" maxlength="3" />
      </div>
      <div class="cell">
        <label for="left_input_col">Match</label>
        <input id="left_input_col" type="text" name="left_input_col" placeholder="D" maxlength="3" />
      </div>
      <div class="cell">
        <label for="left_output_start_col" title="First column where this block is pasted on the new sheet">Out</label>
        <input id="left_output_start_col" type="text" name="left_output_start_col" placeholder="A" maxlength="3"
          title="On the NEW sheet: column where the left block starts (often A)." />
      </div>
    </div>

    <p class="block-title">Right</p>
    <div class="col-row">
      <div class="cell">
        <label for="right_block_start_col">From</label>
        <input id="right_block_start_col" type="text" name="right_block_start_col" placeholder="F" maxlength="3" />
      </div>
      <div class="cell">
        <label for="right_block_end_col">To</label>
        <input id="right_block_end_col" type="text" name="right_block_end_col" placeholder="H" maxlength="3" />
      </div>
      <div class="cell">
        <label for="right_input_col">Match</label>
        <input id="right_input_col" type="text" name="right_input_col" placeholder="F" maxlength="3" />
      </div>
      <div class="cell">
        <label for="right_output_start_col" title="First column where this block is pasted on the new sheet">Out</label>
        <input id="right_output_start_col" type="text" name="right_output_start_col" placeholder="F" maxlength="3"
          title="On the NEW sheet: column where the right block starts. Auto-shifts right if it would overlap the left block." />
      </div>
    </div>

    <div class="inline2">
      <div>
        <label for="threshold" title="Match if amounts are closer than this">Max diff</label>
        <input id="threshold" type="number" step="0.001" name="threshold" placeholder="0.5" />
      </div>
      <div>
        <label for="diff_output_col"
          title="Column on the new sheet between the two pasted blocks — shows |left − right| for each aligned row">
          Diff col</label>
        <input id="diff_output_col" type="text" name="diff_output_col" placeholder="E" maxlength="3"
          title="New sheet column (e.g. E) between left and right data. Absolute difference after aligning." />
      </div>
    </div>

    <button id="submit-btn" type="submit">Download</button>
    <p id="submit-msg" class="submit-msg"></p>
  </form>

  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js" crossorigin="anonymous"></script>
  <script>
  (function () {
    var fileEl = document.getElementById("file");
    var sheetNameEl = document.getElementById("input_sheet_name");
    var previewMsg = document.getElementById("preview-msg");
    var previewHolder = document.getElementById("preview-holder");
    var MAX_R = 45;
    var MAX_C = 16;

    function pickSheet(wb, name) {
      var names = wb.SheetNames;
      if (!name) return names[0];
      if (names.indexOf(name) >= 0) return name;
      var lower = name.toLowerCase();
      for (var i = 0; i < names.length; i++) {
        if (names[i].toLowerCase() === lower) return names[i];
      }
      return names[0];
    }

    function toColLabel(idx) {
      var n = idx + 1;
      var s = "";
      while (n > 0) {
        var rem = (n - 1) % 26;
        s = String.fromCharCode(65 + rem) + s;
        n = Math.floor((n - 1) / 26);
      }
      return s;
    }

    function renderPreview(wb, sheetName) {
      var ws = wb.Sheets[sheetName];
      if (!ws) {
        previewMsg.textContent = "Preview: sheet not found.";
        previewHolder.innerHTML = "";
        return;
      }
      var rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
      var table = document.createElement("table");
      var thead = document.createElement("thead");
      var tbody = document.createElement("tbody");
      var rowCount = Math.min(rows.length, MAX_R);
      var maxCols = 0;
      for (var i = 0; i < rowCount; i++) {
        maxCols = Math.max(maxCols, (rows[i] || []).length);
      }
      var colCount = Math.min(Math.max(maxCols, 1), MAX_C);

      var headTr = document.createElement("tr");
      var corner = document.createElement("th");
      corner.textContent = "#";
      headTr.appendChild(corner);
      for (var hc = 0; hc < colCount; hc++) {
        var h = document.createElement("th");
        h.textContent = toColLabel(hc);
        headTr.appendChild(h);
      }
      thead.appendChild(headTr);

      for (var r = 0; r < rowCount; r++) {
        var tr = document.createElement("tr");
        var rowHead = document.createElement("th");
        rowHead.className = "row-num";
        rowHead.textContent = String(r + 1);
        tr.appendChild(rowHead);
        var row = rows[r] || [];
        for (var c = 0; c < colCount; c++) {
          var td = document.createElement("td");
          var v = row[c];
          td.textContent = v == null ? "" : String(v);
          tr.appendChild(td);
        }
        tbody.appendChild(tr);
      }
      table.appendChild(thead);
      table.appendChild(tbody);
      previewHolder.innerHTML = "";
      previewHolder.appendChild(table);
      var note = rows.length > MAX_R ? " (first " + MAX_R + " of " + rows.length + " rows)" : "";
      previewMsg.textContent = "Preview · " + sheetName + note;
    }

    function runPreview() {
      previewMsg.textContent = "";
      previewHolder.innerHTML = "";
      var f = fileEl.files && fileEl.files[0];
      if (!f) return;
      if (typeof XLSX === "undefined") {
        previewMsg.textContent = "Preview unavailable (spreadsheet parser did not load).";
        return;
      }
      var reader = new FileReader();
      reader.onload = function (e) {
        try {
          var wb = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
          var name = pickSheet(wb, (sheetNameEl.value || "").trim());
          renderPreview(wb, name);
        } catch (err) {
          previewMsg.textContent = "Could not preview this file.";
          previewHolder.innerHTML = "";
        }
      };
      reader.readAsArrayBuffer(f);
    }

    fileEl.addEventListener("change", runPreview);
    sheetNameEl.addEventListener("input", function () {
      if (fileEl.files && fileEl.files[0]) runPreview();
    });

    var formEl = document.getElementById("align-form");
    var submitBtn = document.getElementById("submit-btn");
    var submitMsg = document.getElementById("submit-msg");
    formEl.addEventListener("submit", function () {
      submitBtn.disabled = true;
      submitBtn.textContent = "Working...";
      submitMsg.textContent = "Shifting columns...";
    });
  })();
  </script>
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

    return FileResponse(
        output_path,
        filename=shifted_download_filename(file.filename or ""),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
