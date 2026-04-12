let excelRows = [];

function log(msg) {
  document.getElementById("log").textContent += msg + "\n";
}

function readExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const workbook = XLSX.read(e.target.result, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        resolve(rows);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

document.getElementById("readBtn").addEventListener("click", async () => {
  const file = document.getElementById("fileInput").files[0];
  if (!file) {
    log("Chưa chọn file Excel.");
    return;
  }

  excelRows = await readExcel(file);
  log(`Đọc xong ${excelRows.length} dòng.`);
  log("5 dòng đầu:");
  excelRows.slice(0, 5).forEach(r => {
    log(`${r.GUID} | ${r["PAINT CODE"]}`);
  });
});