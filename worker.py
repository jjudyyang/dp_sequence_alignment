"""Process exactly one workbook outside the web server's event loop and memory."""

import json
from pathlib import Path
import sys
import time
import traceback

from alignment import process_excel
from jobs import read_state, write_state


def run(directory):
    # Keep one complex file from exhausting the 512 MB Render instance. macOS
    # does not support a useful RLIMIT_AS here; the deployment and Linux CI do.
    if sys.platform == "linux":
        import resource
        resource.setrlimit(resource.RLIMIT_AS, (384 * 1024 * 1024, 384 * 1024 * 1024))
    state = read_state(directory)

    def progress(stage):
        state.update(status="processing", stage=stage, updated_at=time.time())
        write_state(directory, state)

    try:
        options = json.loads((directory / "options.json").read_text())
        result = process_excel(
            directory / "input.xlsx", directory / "output.xlsx", progress=progress, **options
        )
        result.pop("output_file", None)
        state.update(status="finishing", stage="Finishing download", result=result, updated_at=time.time())
    except MemoryError:
        state.update(status="failed", stage="Could not finish", error="This workbook needs more memory than available. Remove unused sheets or excess formatting, then try again.", updated_at=time.time())
    except ValueError as exc:
        state.update(status="failed", stage="Could not finish", error=str(exc), updated_at=time.time())
    except Exception:
        traceback.print_exc()
        state.update(status="failed", stage="Could not finish", error="The workbook could not be processed. Save a fresh .xlsx copy in Excel and try again.", updated_at=time.time())
    write_state(directory, state)
    if state["status"] == "failed":
        (directory / "output.xlsx").unlink(missing_ok=True)


if __name__ == "__main__":
    run(Path(sys.argv[1]))
