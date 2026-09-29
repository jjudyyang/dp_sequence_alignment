from pathlib import Path
import asyncio
from contextlib import asynccontextmanager
from dataclasses import asdict
import hashlib
import hmac
import os
import re
import secrets
import shutil
import sqlite3
import time
import uuid

from fastapi import FastAPI, File, Form, HTTPException, Request, Response, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from starlette.concurrency import run_in_threadpool
from starlette.formparsers import MultiPartException

from alignment import AlignmentConfig
from jobs import JobManager, LOG
from workbook_limits import MAX_UPLOAD_BYTES, QUICK_ROWS, inspect_workbook


@asynccontextmanager
async def lifespan(app):
    app.state.jobs = JobManager(DATA_DIR / "jobs", record_successful_run)
    await app.state.jobs.start()
    try:
        yield
    finally:
        await app.state.jobs.stop()


app = FastAPI(title="Shiftline", lifespan=lifespan)

BASE_DIR = Path(__file__).resolve().parent
FRONTEND_DIR = BASE_DIR / "frontend"
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
DATA_DIR = Path(os.getenv("SHIFTLINE_DATA_DIR", BASE_DIR / "data"))
USAGE_DB = DATA_DIR / "usage.db"
APP_PASSWORD = os.getenv("SHIFTLINE_PASSWORD", "judy")
AUTH_SECRET = os.getenv("SHIFTLINE_AUTH_SECRET", secrets.token_hex(32))
AUTH_COOKIE_NAME = "shiftline_auth"
AUTH_COOKIE_MAX_AGE = 60 * 60 * 24 * 30

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
DATA_DIR.mkdir(parents=True, exist_ok=True)

app.mount("/assets", StaticFiles(directory=FRONTEND_DIR / "assets"), name="assets")
app.mount("/data", StaticFiles(directory=FRONTEND_DIR / "data"), name="sample-data")


def auth_cookie_value():
    """Return the signed value accepted for an authenticated browser session."""
    signature = hmac.new(
        AUTH_SECRET.encode("utf-8"),
        b"shiftline-authenticated",
        hashlib.sha256,
    ).hexdigest()
    return f"v1:{signature}"


def is_auth_cookie_valid(cookie_value):
    """Check a submitted auth cookie without leaking timing information."""
    if not cookie_value:
        return False
    return hmac.compare_digest(cookie_value, auth_cookie_value())


def is_authenticated(request: Request):
    """Return whether the incoming request already passed the password gate."""
    return is_auth_cookie_valid(request.cookies.get(AUTH_COOKIE_NAME))


def login_page(error_message: str = ""):
    """Render the password page without relying on protected static assets."""
    error_html = (
        f'<p class="error" role="alert">{error_message}</p>' if error_message else ""
    )
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Shiftline Password</title>
  <style>
    * {{ box-sizing: border-box; }}
    body {{
      min-height: 100vh;
      margin: 0;
      display: grid;
      place-items: center;
      background: #f3f5f7;
      color: #17212f;
      font: 14px/1.45 Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }}
    main {{
      width: min(100% - 32px, 360px);
      padding: 24px;
      border: 1px solid #d9e0e8;
      border-radius: 8px;
      background: #ffffff;
      box-shadow: 0 12px 28px rgba(23, 33, 47, 0.08);
    }}
    h1 {{
      margin: 0 0 6px;
      font-size: 20px;
      line-height: 1.2;
    }}
    p {{
      margin: 0 0 18px;
      color: #687586;
    }}
    label {{
      display: grid;
      gap: 6px;
      color: #4f5d6d;
      font-size: 12px;
      font-weight: 650;
    }}
    input {{
      width: 100%;
      min-height: 40px;
      border: 1px solid #aab6c5;
      border-radius: 6px;
      padding: 0 10px;
      font: inherit;
    }}
    button {{
      width: 100%;
      min-height: 40px;
      margin-top: 14px;
      border: 1px solid #2458a6;
      border-radius: 6px;
      background: #2458a6;
      color: #ffffff;
      font: inherit;
      font-weight: 700;
      cursor: pointer;
    }}
    .error {{
      margin: 0 0 12px;
      color: #b42318;
      font-weight: 650;
    }}
  </style>
</head>
<body>
  <main>
    <h1>Shiftline</h1>
    <p>Enter the password to continue.</p>
    {error_html}
    <form method="post" action="/login">
      <input name="username" type="text" value="shiftline" autocomplete="username" hidden>
      <label>
        Password
        <input name="password" type="password" autocomplete="current-password" autofocus required>
      </label>
      <button type="submit">Open Shiftline</button>
    </form>
  </main>
</body>
</html>"""


@app.middleware("http")
async def require_password(request: Request, call_next):
    """Protect every Shiftline route except the password form."""
    if request.url.path in {"/login", "/favicon.ico", "/healthz"} or is_authenticated(request):
        return await call_next(request)
    accepts_html = "text/html" in request.headers.get("accept", "")
    if request.method == "GET" and accepts_html:
        return RedirectResponse("/login", status_code=303)
    return JSONResponse({"detail": "Password required."}, status_code=401)


class UploadLimitMiddleware:
    """Bound multipart bytes before parsing/spooling, including chunked requests."""
    def __init__(self, app):
        self.app = app
        self.uploads = 0

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http" or scope["path"] != "/process" or scope["method"] != "POST":
            return await self.app(scope, receive, send)
        limit = MAX_UPLOAD_BYTES + 64 * 1024  # Allow a small multipart envelope.
        headers = dict(scope["headers"])
        try:
            size = int(headers.get(b"content-length", b"0"))
        except ValueError:
            size = limit + 1
        if size > limit:
            return await JSONResponse({"detail": "The upload limit is 25 MB."}, status_code=413)(scope, receive, send)
        if self.uploads >= 3:
            return await JSONResponse({"detail": "Other files are uploading. Please try again in a minute."}, status_code=429)(scope, receive, send)
        self.uploads += 1
        received = 0
        too_large = False

        async def bounded_receive():
            nonlocal received, too_large
            message = await receive()
            received += len(message.get("body", b""))
            if received > limit:
                too_large = True
                # Let the multipart parser close any temporary upload files.
                raise MultiPartException("The upload limit is 25 MB.")
            return message

        async def bounded_send(message):
            if too_large:
                # Multipart errors become 400 inside Starlette. Preserve the
                # specific size-limit status for clients, including chunked uploads.
                if message["type"] == "http.response.start":
                    await JSONResponse({"detail": "The upload limit is 25 MB."}, status_code=413)(scope, receive, send)
                return
            await send(message)

        try:
            await self.app(scope, bounded_receive, bounded_send)
        finally:
            self.uploads -= 1


app.add_middleware(UploadLimitMiddleware)


@app.get("/healthz")
def health():
    return {"status": "ok"}


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


@app.get("/login", response_class=HTMLResponse)
def login_form(request: Request):
    if is_authenticated(request):
        return RedirectResponse("/", status_code=303)
    return HTMLResponse(login_page())


@app.post("/login")
def login(request: Request, password: str = Form("")):
    if not hmac.compare_digest(password, APP_PASSWORD):
        return HTMLResponse(login_page("Incorrect password."), status_code=401)
    response = RedirectResponse("/", status_code=303)
    response.set_cookie(
        AUTH_COOKIE_NAME,
        auth_cookie_value(),
        max_age=AUTH_COOKIE_MAX_AGE,
        httponly=True,
        secure=request.url.scheme == "https",
        samesite="lax",
    )
    return response


@app.get("/favicon.ico")
def favicon():
    return Response(status_code=204)


@app.get("/", response_class=HTMLResponse)
def home():
    index_path = FRONTEND_DIR / "index.html"
    return HTMLResponse(
        index_path.read_text(encoding="utf-8").replace(
            "__USAGE_COUNT__", str(get_usage_count())
        )
    )


@app.get("/usage")
def usage():
    return {"runs": get_usage_count()}


@app.post("/process")
async def process_file(
    request: Request,
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
    workflow: str = Form("auto"),
):
    manager = request.app.state.jobs
    job_id = request.headers.get("X-Upload-ID", uuid.uuid4().hex)
    if not re.fullmatch(r"[a-f0-9]{32}", job_id):
        raise HTTPException(400, "Invalid upload identifier. Refresh the page and try again.")
    if workflow not in {"auto", "large"}:
        raise HTTPException(400, "Choose Automatic or Large workbook processing.")
    if not (file.filename or "").lower().endswith(".xlsx"):
        raise HTTPException(400, "Please choose an .xlsx workbook.")
    if file.size is not None and file.size > MAX_UPLOAD_BYTES:
        raise HTTPException(413, "The upload limit is 25 MB.")
    try:
        reserved = manager.reserve(job_id, shifted_download_filename(file.filename or ""))
    except OverflowError as exc:
        raise HTTPException(429, str(exc)) from exc
    if not reserved:
        await file.close()
        return JSONResponse(job_response(get_job(request, job_id)), status_code=202)

    directory = manager.directory(job_id)
    try:
        def save_upload():
            with (directory / "input.xlsx").open("wb") as buffer:
                shutil.copyfileobj(file.file, buffer, length=1024 * 1024)
        await run_in_threadpool(save_upload)
        config = AlignmentConfig(
            input_sheet_name=input_sheet_name, left_sheet_name=left_sheet_name,
            right_sheet_name=right_sheet_name, output_sheet_name=output_sheet_name,
            start_row=start_row, header_first_row=header_first_row, header_last_row=header_last_row,
            left_input_col=left_input_col, left_block_start_col=left_block_start_col,
            left_block_end_col=left_block_end_col, left_output_start_col=left_output_start_col,
            right_input_col=right_input_col, right_block_start_col=right_block_start_col,
            right_block_end_col=right_block_end_col, right_output_start_col=right_output_start_col,
            threshold=threshold, diff_output_col=diff_output_col,
        )
        info = await run_in_threadpool(inspect_workbook, directory / "input.xlsx", config)
        large = workflow == "large" or info["rows"] > QUICK_ROWS or (file.size or 0) > 2 * 1024 * 1024
        manager.submit(job_id, asdict(config), info, "large" if large else "quick")
        LOG.info("Upload accepted job=%s rows=%s workflow=%s request=%s", job_id, info["rows"], "large" if large else "quick", request.headers.get("Rndr-Id", "local"))
    except ValueError as exc:
        manager.fail(job_id, str(exc))
        raise HTTPException(400, str(exc)) from exc
    except asyncio.CancelledError:
        manager.fail(job_id, "The upload connection closed before the workbook was accepted. Please upload it again.")
        raise
    except Exception:
        manager.fail(job_id, "The upload could not be saved. Please try again.")
        raise
    finally:
        await file.close()

    # Small workbooks retain a direct download; never hold a processing request
    # open longer than four seconds. Larger/queued jobs return a polling URL.
    if not large:
        deadline = time.monotonic() + 4
        while time.monotonic() < deadline:
            state = manager.status(job_id)
            if state["status"] == "succeeded":
                return download_response(manager, state)
            if state["status"] == "failed":
                raise HTTPException(400, state["error"])
            await asyncio.sleep(0.1)
    return JSONResponse(job_response(manager.status(job_id)), status_code=202, headers={"Cache-Control": "no-store"})


def job_response(state):
    return {**state, "status_url": f"/jobs/{state['id']}", "download_url": f"/jobs/{state['id']}/download"}


def get_job(request, job_id):
    try:
        return request.app.state.jobs.status(job_id)
    except (FileNotFoundError, ValueError):
        raise HTTPException(404, "This job expired or the server restarted. Please upload the workbook again.") from None


@app.get("/jobs/{job_id}")
def job_status(request: Request, job_id: str):
    return JSONResponse(job_response(get_job(request, job_id)), headers={"Cache-Control": "no-store"})


def download_response(manager, state):
    path = manager.directory(state["id"]) / "output.xlsx"
    if not path.exists():
        raise HTTPException(410, "The download expired or the server restarted. Please upload the workbook again.")
    headers = {"Cache-Control": "no-store"}
    if "runs_total" in state:
        headers["X-Successful-Runs"] = str(state["runs_total"])
    return FileResponse(path, filename=state["filename"], media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers=headers)


@app.get("/jobs/{job_id}/download")
def download_job(request: Request, job_id: str):
    state = get_job(request, job_id)
    if state["status"] != "succeeded":
        raise HTTPException(409, state.get("error", "Your workbook is still processing."))
    return download_response(request.app.state.jobs, state)
