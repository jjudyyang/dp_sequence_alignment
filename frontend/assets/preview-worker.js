self.onmessage = async function (event) {
  try {
    importScripts("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js");
    const workbook = XLSX.read(await event.data.arrayBuffer(), { type: "array", sheetRows: 45 });
    const sheets = {};
    workbook.SheetNames.forEach((name) => {
      sheets[name] = XLSX.utils.sheet_to_json(workbook.Sheets[name], {
        header: 1, defval: "", raw: false, range: "A1:P45",
      });
    });
    self.postMessage({ sheets });
  } catch (_) {
    self.postMessage({ error: true });
  }
};
