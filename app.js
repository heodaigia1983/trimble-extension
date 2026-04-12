let API = null;
let excelRows = [];

const APPROVED_COLOR = "#4CAF50"; // xanh
const OTHER_COLOR = "#BDBDBD";    // xám

function log(msg) {
  document.getElementById("log").textContent += msg + "\n";
}

function clearLog() {
  document.getElementById("log").textContent = "";
}

async function initAPI() {
  if (API) return API;

  API = await TrimbleConnectWorkspace.connect(window.parent, (event, data) => {
    console.log("Trimble event:", event, data);
  });

  log("Đã kết nối Trimble API.");
  return API;
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

function normalizeRows(rows) {
  const seen = new Set();

  return rows
    .map(r => String(r.GUID || "").trim())
    .filter(guid => guid)
    .filter(guid => {
      if (seen.has(guid)) return false;
      seen.add(guid);
      return true;
    });
}

async function getLoadedModelId() {
  const api = await initAPI();

  const modelObjects = await api.viewer.getObjects();

  if (!modelObjects || !modelObjects.length) {
    throw new Error("Viewer chưa trả về object nào.");
  }

  const modelIds = [...new Set(modelObjects.map(x => x.modelId).filter(Boolean))];
  log("Loaded modelIds trong viewer: " + modelIds.join(", "));

  if (!modelIds.length) {
    throw new Error("Không lấy được modelId từ viewer.");
  }

  return modelIds[0];
}

function chunkArray(arr, size) {
  const result = [];
  for (let i = 0; i < arr.length; i += size) {
    result.push(arr.slice(i, i + size));
  }
  return result;
}

async function setGreenInBatches(api, modelId, runtimeIds) {
  const batches = chunkArray(runtimeIds, 1000);

  for (let i = 0; i < batches.length; i++) {
    const ids = batches[i];

    await api.viewer.setObjectState(
      {
        modelObjectIds: [
          {
            modelId: modelId,
            objectRuntimeIds: ids
          }
        ]
      },
      {
        color: APPROVED_COLOR
      }
    );

    log(`Tô xanh: batch ${i + 1}/${batches.length} (${ids.length} object)`);
  }
}

async function colorApprovedObjects() {
  const api = await initAPI();

  if (!excelRows.length) {
    log("Chưa có dữ liệu Excel.");
    return;
  }

  const modelId = await getLoadedModelId();
  const excelGuids = normalizeRows(excelRows);

  log("ModelId: " + modelId);
  log("Số GUID duy nhất trong Excel: " + excelGuids.length);

  // B1: tô xám TOÀN BỘ model
  log("Bắt đầu tô xám toàn bộ model...");
  await api.viewer.setObjectState(undefined, { color: OTHER_COLOR });

  // B2: đổi GUID trong Excel -> runtimeId
  log("Bắt đầu đổi GUID -> runtimeId...");
  log("Test GUID đầu tiên: " + excelGuids[0]);

  let testRuntimeIds;
  try {
    testRuntimeIds = await api.viewer.convertToObjectRuntimeIds(modelId, [excelGuids[0]]);
    log("Test runtimeIds[0]: " + JSON.stringify(testRuntimeIds));
  } catch (err) {
    throw new Error("Lỗi test convert GUID đầu tiên: " + (err?.message || String(err)));
  }

  let runtimeIds;
  try {
    runtimeIds = await api.viewer.convertToObjectRuntimeIds(modelId, excelGuids);
  } catch (err) {
    throw new Error("Lỗi convert full GUID list: " + (err?.message || String(err)));
  }

  const approvedRuntimeIds = [];
  let matched = 0;
  let unmatched = 0;

  for (let i = 0; i < excelGuids.length; i++) {
    const runtimeId = runtimeIds[i];

    if (runtimeId === undefined || runtimeId === null) {
      unmatched++;
      continue;
    }

    approvedRuntimeIds.push(runtimeId);
    matched++;
  }

  log("Match: " + matched);
  log("Không match: " + unmatched);

  // B3: tô xanh phần có trong Excel
  if (approvedRuntimeIds.length) {
    log("Bắt đầu tô xanh phần có trong Excel...");
    await setGreenInBatches(api, modelId, approvedRuntimeIds);
  }

  log("Hoàn tất.");
  log("Kết quả: Có trong Excel = xanh | Không có trong Excel = xám");
}

document.getElementById("readBtn").addEventListener("click", async () => {
  try {
    const file = document.getElementById("fileInput").files[0];

    if (!file) {
      log("Chưa chọn file Excel.");
      return;
    }

    clearLog();
    await initAPI();

    excelRows = await readExcel(file);

    log(`Đọc xong ${excelRows.length} dòng.`);
    log("5 GUID đầu:");
    excelRows.slice(0, 5).forEach(r => {
      log(String(r.GUID || "").trim());
    });

    await colorApprovedObjects();
  } catch (err) {
    console.error(err);
    log("Lỗi: " + (err?.message || JSON.stringify(err) || String(err)));
  }
});
