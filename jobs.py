"""A bounded, single-process worker queue for this single-instance deployment."""

import asyncio
import fcntl
from contextlib import suppress
import json
import logging
from pathlib import Path
import re
import shutil
import sys
import time

LOG = logging.getLogger("uvicorn.error")
TERMINAL = {"succeeded", "failed"}
RETENTION_SECONDS = 3600
JOB_TIMEOUT_SECONDS = 600
MAX_ACTIVE_JOBS = 3
MAX_RETAINED_BYTES = 256 * 1024 * 1024


def read_state(directory):
    return json.loads((directory / "status.json").read_text())


def write_state(directory, state):
    temporary = directory / "status.tmp"
    temporary.write_text(json.dumps(state))
    temporary.replace(directory / "status.json")


class JobManager:
    def __init__(self, root, on_success, *, timeout=JOB_TIMEOUT_SECONDS):
        self.root = Path(root)
        self.on_success = on_success
        self.timeout = timeout
        self.active = set()
        self.queue = asyncio.Queue()
        self.process = None
        self.runner = None
        self.cleaner = None
        self.current_id = None
        self.lock_file = None

    def directory(self, job_id):
        if not re.fullmatch(r"[a-f0-9]{32}", job_id):
            raise FileNotFoundError(job_id)
        return self.root / job_id

    def status(self, job_id):
        directory = self.directory(job_id)
        state = read_state(directory)
        if state["status"] in TERMINAL and time.time() - state["updated_at"] > RETENTION_SECONDS:
            raise FileNotFoundError(job_id)
        return state

    async def start(self):
        self.root.mkdir(parents=True, exist_ok=True)
        self.lock_file = (self.root / ".worker.lock").open("a")
        try:
            fcntl.flock(self.lock_file, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            self.lock_file.close()
            raise RuntimeError("Shiftline requires one Uvicorn worker per job directory. Start with --workers 1.") from None
        for directory in self.root.iterdir():
            if directory.is_dir() and (directory / "status.json").exists():
                state = read_state(directory)
                if state["status"] not in TERMINAL:
                    self.fail(directory.name, "The server restarted before this file finished. Please upload it again.")
        await self.cleanup()
        self.runner = asyncio.create_task(self.run())
        self.cleaner = asyncio.create_task(self.clean_periodically())

    async def stop(self):
        for task in (self.runner, self.cleaner):
            if task:
                task.cancel()
        if self.process and self.process.returncode is None:
            self.process.kill()
            await self.process.wait()
        for task in (self.runner, self.cleaner):
            if task:
                with suppress(asyncio.CancelledError):
                    await task
        for job_id in list(self.active):
            self.fail(job_id, "The server restarted before this file finished. Please upload it again.")
        self.active.clear()
        if self.lock_file:
            self.lock_file.close()

    def reserve(self, job_id, filename):
        directory = self.directory(job_id)
        if directory.exists():
            return False
        if len(self.active) >= MAX_ACTIVE_JOBS:
            raise OverflowError("The server is processing other workbooks. Please try again in a minute.")
        retained = 0
        for path in self.root.glob("*/*.xlsx"):
            with suppress(FileNotFoundError):  # Expiry cleanup can run concurrently.
                retained += path.stat().st_size
        if retained > MAX_RETAINED_BYTES:
            raise OverflowError("Temporary download storage is full. Please retry after older downloads expire (within one hour).")
        directory.mkdir()
        self.active.add(job_id)
        write_state(directory, {
            "id": job_id, "filename": filename, "status": "uploading",
            "stage": "Checking workbook", "created_at": time.time(), "updated_at": time.time(),
        })
        return True

    def submit(self, job_id, options, info, workflow):
        directory = self.directory(job_id)
        (directory / "options.json").write_text(json.dumps(options))
        state = read_state(directory)
        state.update(status="queued", stage="Waiting to start", workflow=workflow, **info)
        write_state(directory, state)
        self.queue.put_nowait(job_id)

    def fail(self, job_id, message):
        directory = self.directory(job_id)
        state = read_state(directory)
        state.update(status="failed", stage="Could not finish", error=message, updated_at=time.time())
        write_state(directory, state)
        for name in ("input.xlsx", "output.xlsx", "options.json"):
            (directory / name).unlink(missing_ok=True)
        self.active.discard(job_id)

    async def run(self):
        while True:
            job_id = await self.queue.get()
            self.current_id = job_id
            directory = self.directory(job_id)
            try:
                LOG.info("Alignment job %s started", job_id)
                self.process = await asyncio.create_subprocess_exec(
                    sys.executable, str(Path(__file__).with_name("worker.py")), str(directory),
                )
                try:
                    await asyncio.wait_for(self.process.wait(), self.timeout)
                except asyncio.TimeoutError:
                    self.process.kill()
                    await self.process.wait()
                    self.fail(job_id, "Processing exceeded 10 minutes. Please remove unused sheets or split this workbook and try again.")
                state = read_state(directory)
                if state["status"] == "finishing":
                    try:
                        state["runs_total"] = await asyncio.to_thread(self.on_success)
                    except Exception:
                        LOG.exception("Could not update usage for job %s; download is still available", job_id)
                    state.update(status="succeeded", stage="Ready to download", updated_at=time.time())
                    write_state(directory, state)
                elif state["status"] != "failed":
                    self.fail(job_id, "The workbook worker stopped unexpectedly. Please remove unused sheets or excess formatting and retry.")
                LOG.info("Alignment job %s %s", job_id, read_state(directory)["status"])
            except asyncio.CancelledError:
                self.fail(job_id, "The server restarted before this file finished. Please upload it again.")
                raise
            except Exception:
                LOG.exception("Alignment job %s failed", job_id)
                self.fail(job_id, "The server could not finish this workbook. Please try again.")
            finally:
                self.active.discard(job_id)
                self.current_id = None
                self.process = None
                self.queue.task_done()
                for name in ("input.xlsx", "options.json"):
                    (directory / name).unlink(missing_ok=True)

    async def cleanup(self):
        for directory in self.root.iterdir():
            if directory.is_dir() and directory.name not in self.active:
                with suppress(FileNotFoundError, ValueError):
                    state = read_state(directory)
                    if state["status"] in TERMINAL and time.time() - state["updated_at"] > RETENTION_SECONDS:
                        await asyncio.to_thread(shutil.rmtree, directory)

    async def clean_periodically(self):
        while True:
            await asyncio.sleep(60)
            await self.cleanup()
