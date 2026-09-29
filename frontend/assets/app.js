(function () {
  "use strict";
  const $ = (id) => document.getElementById(id);
  const form = $("align-form");
  const fileInput = $("file");
  const submit = $("submit-btn");
  const message = $("submit-msg");
  const storageKey = "shiftline-active-job";
  const maxBytes = 25 * 1024 * 1024;
  let activeId = null;
  let previewWorker = null;
  let previewSheets = null;
  let pollTimer = null;
  let failures = 0;

  function saveJob(id) {
    activeId = id;
    try {
      if (id) sessionStorage.setItem(storageKey, id);
      else sessionStorage.removeItem(storageKey);
    } catch (_) { /* Processing still works when browser storage is unavailable. */ }
  }

  function busy(value) {
    submit.disabled = value;
    submit.textContent = value ? "Working…" : "Download";
  }

  function stage(text, hint, percent) {
    const newlyVisible = $("job-panel").hidden;
    $("job-panel").hidden = false;
    $("job-stage").textContent = text;
    $("job-hint").textContent = hint || "";
    $("job-progress").hidden = false;
    if (typeof percent === "number") $("job-progress").value = percent;
    else $("job-progress").removeAttribute("value");
    if (newlyVisible) $("job-panel").scrollIntoView({ block: "nearest", behavior: "smooth" });
  }

  function resetActions() {
    $("job-download").hidden = true;
    $("job-retry").hidden = true;
    $("job-login").hidden = true;
  }

  function stopWithError(text, retry) {
    stage(text, "");
    $("job-progress").hidden = true;
    $("job-retry").hidden = !retry;
    busy(false);
  }

  function schedulePoll() {
    clearTimeout(pollTimer);
    pollTimer = setTimeout(poll, Math.min(1500 * (failures + 1), 10000));
  }

  async function poll() {
    if (!activeId) return;
    const id = activeId;
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 15000);
    try {
      const response = await fetch("/jobs/" + id, { cache: "no-store", signal: controller.signal });
      if (id !== activeId) return;
      if (response.status === 401) {
        stopWithError("Please sign in again to check your workbook.", false);
        $("job-login").hidden = false;
        return;
      }
      if (response.status === 404) {
        saveJob(null);
        stopWithError("This job expired or the server restarted. Please upload the workbook again.", false);
        return;
      }
      if (!response.ok) throw new Error("Connection interrupted");
      const job = await response.json();
      if (id !== activeId) return;
      failures = 0;
      resetActions();
      if (job.status === "succeeded") {
        stage("Ready — your aligned workbook is available.",
          job.rows.toLocaleString() + " data rows processed · " + job.filename +
          ". Download is available for one hour, unless the server restarts.");
        $("job-progress").hidden = true;
        const link = $("job-download");
        link.href = job.download_url;
        link.download = job.filename;
        link.hidden = false;
        if (job.runs_total && $("usage-count")) $("usage-count").textContent = job.runs_total;
        busy(false);
      } else if (job.status === "failed") {
        saveJob(null);
        stopWithError(job.error || "Could not finish this workbook.", false);
      } else {
        busy(true);
        stage(job.stage, (job.workflow === "large" ? "Large workbook · " : "") +
          (job.rows ? job.rows.toLocaleString() + " data rows. " : "") +
          "Upload received. You can refresh this tab to reconnect; do not upload again.");
        schedulePoll();
      }
    } catch (_) {
      if (id !== activeId) return;
      failures += 1;
      if (failures < 5) {
        stage("Reconnecting to your workbook…", "Your upload has already been received. Checking its progress again.");
        schedulePoll();
      } else {
        stopWithError("We cannot reach the server. Your job may still be running.", true);
      }
    } finally {
      clearTimeout(timeout);
    }
  }

  $("job-retry").addEventListener("click", () => {
    failures = 0;
    resetActions();
    busy(true);
    poll();
  });

  function upload(data, id) {
    return new Promise((resolve, reject) => {
      const xhr = new XMLHttpRequest();
      xhr.open("POST", form.action);
      xhr.setRequestHeader("X-Upload-ID", id);
      xhr.responseType = "blob";
      xhr.timeout = 180000;
      xhr.upload.onprogress = (event) => {
        const percent = event.lengthComputable ? Math.round(100 * event.loaded / event.total) : undefined;
        stage(percent === 100 ? "Upload sent — checking workbook…" : "Uploading workbook…",
          "Keep this tab open until the server confirms receipt.", percent);
      };
      xhr.onload = () => resolve(xhr);
      xhr.onerror = xhr.ontimeout = () => reject(new Error("Upload connection interrupted."));
      xhr.send(data);
    });
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    if (submit.disabled) return;
    const file = fileInput.files[0];
    if (!file || file.size > maxBytes) {
      message.textContent = "Choose an .xlsx workbook no larger than 25 MB.";
      return;
    }
    clearTimeout(pollTimer);
    resetActions();
    message.textContent = "";
    busy(true);
    // Save the identifier before sending so a lost acknowledgement cannot lead
    // to a second job. The server uses the same ID for status and deduplication.
    const id = crypto.randomUUID().replace(/-/g, "");
    saveJob(id);
    stage("Uploading workbook…", "Keep this tab open until the server confirms receipt.", 0);
    try {
      const response = await upload(new FormData(form), id);
      if (response.status === 202) {
        const job = JSON.parse(await response.response.text());
        saveJob(job.id);
        failures = 0;
        await poll();
      } else if (response.status === 200 && response.response.type.includes("spreadsheetml")) {
        downloadBlob(response.response, file.name.replace(/\.xlsx$/i, "") + "_shifted.xlsx");
        await poll(); // Keep a retryable download link as well as the quick download.
      } else {
        let detail;
        try { detail = JSON.parse(await response.response.text()).detail; } catch (_) { /* Proxy errors may be HTML. */ }
        if (response.status === 401) {
          saveJob(null);
          stopWithError("Please sign in again, then upload your workbook.", false);
          $("job-login").hidden = false;
        } else if (response.status >= 500) {
          // The server may have accepted the job before the response was lost.
          await poll();
        } else {
          saveJob(null);
          stopWithError(typeof detail === "string" ? detail : "Check the workbook and field settings, then try again.", false);
        }
      }
    } catch (_) {
      stage("Checking whether your upload was received…", "Please do not upload a second copy yet.");
      failures = 0;
      await poll();
    }
  });

  // Preview at most 45 rows in a worker. Never convert all 10,000 rows into a
  // browser-side array just to display a small table.
  function renderPreview() {
    if (!previewSheets) return;
    const requested = $("input_sheet_name").value.trim().toLowerCase();
    const name = Object.keys(previewSheets).find((key) => key.toLowerCase() === requested) || Object.keys(previewSheets)[0];
    const rows = previewSheets[name] || [];
    const table = document.createElement("table");
    const header = document.createElement("tr");
    ["#", ..."ABCDEFGHIJKLMNOP"].forEach((label) => {
      const th = document.createElement("th");
      th.textContent = label;
      header.appendChild(th);
    });
    table.appendChild(header);
    rows.forEach((row, index) => {
      const tr = document.createElement("tr");
      const th = document.createElement("th");
      th.textContent = index + 1;
      tr.appendChild(th);
      row.slice(0, 16).forEach((value) => {
        const td = document.createElement("td");
        td.textContent = value == null ? "" : value;
        tr.appendChild(td);
      });
      table.appendChild(tr);
    });
    $("preview-holder").replaceChildren(table);
    $("preview-msg").textContent = "Preview · " + name + " · first 45 rows and 16 columns only";
  }

  fileInput.addEventListener("change", () => {
    if (previewWorker) previewWorker.terminate();
    previewSheets = null;
    $("preview-holder").replaceChildren();
    const file = fileInput.files[0];
    if (!file) { $("preview-msg").textContent = ""; return; }
    if (file.size > maxBytes) {
      $("preview-msg").textContent = "This file exceeds the 25 MB upload limit.";
      return;
    }
    if (file.size > 5 * 1024 * 1024) {
      $("preview-msg").textContent = "Preview skipped for this large file. You can still upload and process it.";
      return;
    }
    $("preview-msg").textContent = "Loading a small preview… You can upload while it loads.";
    previewWorker = new Worker("/assets/preview-worker.js");
    previewWorker.onmessage = (event) => {
      if (event.data.error) $("preview-msg").textContent = "Preview unavailable. You can still upload the workbook.";
      else { previewSheets = event.data.sheets; renderPreview(); }
      previewWorker.terminate();
      previewWorker = null;
    };
    previewWorker.onerror = () => { $("preview-msg").textContent = "Preview unavailable. You can still upload the workbook."; };
    previewWorker.postMessage(file);
  });
  $("input_sheet_name").addEventListener("input", renderPreview);

  try { activeId = sessionStorage.getItem(storageKey); } catch (_) { /* optional */ }
  if (activeId && /^[a-f0-9]{32}$/.test(activeId)) {
    busy(true);
    stage("Reconnecting to your workbook…", "");
    poll();
  }
})();
