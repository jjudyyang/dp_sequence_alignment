# Shiftline Frontend

This folder contains the browser UI served by the FastAPI app at `/`.

## Open

Run the FastAPI app from the repo root:

```bash
python -m uvicorn app:app --reload
```

Then open `http://127.0.0.1:8000/`.

## Refresh Sample Data

Run this from the repo root after replacing or editing `input.xlsx`:

```bash
"/Users/judyyang/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3" frontend/tools/extract_workbook_preview.py input.xlsx --output frontend/data/sample-workbook.js --max-rows 40 --max-cols 8 --source-url "https://docs.google.com/spreadsheets/d/1LqQa7-rId7rYK-N7MViDUrTGGkrijHbg/edit?usp=drive_web&ouid=106529433934137787357&rtpof=true"
```

The generated fixture includes workbook metadata, sheet names, bounded preview
rows, and lightweight style keys for the default landing preview.
