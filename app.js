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

  try {
    const viewerObjects = await api.viewer.getObjects();

    if (Array.isArray(viewerObjects) && viewerObjects.length) {
      const modelIds = [...new Set(viewerObjects.map(x => x.modelId).filter(Boolean))];
      if (modelIds.length) {
        log("Loaded modelIds trong viewer: " + modelIds.join(", "));
        return modelIds[0];
      }
    }

    if (
      viewerObjects &&
      Array.isArray(viewerObjects.modelObjectIds) &&
      viewerObjects.modelObjectIds.length
    ) {
      const modelIds = [
        ...new Set(viewerObjects.modelObjectIds.map(x => x.modelId).filter(Boolean))
      ];
      if (modelIds.length) {
        log("Loaded modelIds trong viewer: " + modelIds.join(", "));
        return modelIds[0];
      }
    }
  } catch (err) {
    log("getObjects fallback: " + (err?.message || String(err)));
  }

  const models = await api.viewer.getModels();

  if (!models || !models.length) {
    throw new Error("Không tìm thấy model đang load.");
  }

  log("viewer.getModels(): " + models.map(m => m.id).join(", "));
  return models[0].id;
}

function extractRuntimeIdsFromViewerObjects(raw, modelId) {
  const ids = [];

  function pushIds(item) {
    if (!item) return;
    if (item.modelId && modelId && item.modelId !== modelId) return;

    if (Array.isArray(item.objectRuntimeIds)) {
      ids.push(...item.objectRuntimeIds.filter(x => x !== undefined && x !== null));
    }

    if (item.objectRuntimeId !== undefined && item.objectRuntimeId !== null) {
      ids.push(item.objectRuntimeId);
    }

    if (item.runtimeId !== undefined && item.runtimeId !== null) {
      ids.push(item.runtimeId);
    }

    if (Array.isArray(item.ids)) {
      ids.push(...item.ids.filter(x => x !== undefined && x !== null));
    }
  }

  if (Array.isArray(raw)) {
    raw.forEach(pushIds);
  } else if (raw) {
    pushIds(raw);

    if (Array.isArray(raw.modelObjectIds)) {
      raw.modelObjectIds.forEach(pushIds);
    }

    if (Array.isArray(raw.objects)) {
      raw.objects.forEach(pushIds);
    }
  }

  return [...new Set(ids)];
}

function chunkArray(arr, size) {
  const result = [];
  for (let i = 0; i < arr.length; i += size) {
    result.push(arr.slice(i, i + size));
  }
  return result;
}

async function setColorInBatches(api, modelId, runtimeIds, color, label) {
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
        color: color
      }
    );

    log(`${label}: batch ${i + 1}/${batches.length} (${ids.length} object)`);
  }
}

async function grayAllObjects(api, modelId) {
  const viewerObjects = await api.viewer.getObjects();
  const allRuntimeIds = extractRuntimeIdsFromViewerObjects(viewerObjects, modelId);

  if (!allRuntimeIds.length) {
    throw new Error("Không lấy được danh sách object của model để tô xám.");
  }

  log("Tổng object trong viewer: " + allRuntimeIds.length);
  await setColorInBatches(api, modelId, allRuntimeIds, OTHER_COLOR, "Tô xám");
  return allRuntimeIds;
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

  // B1: tô xám toàn bộ model
  log("Bắt đầu tô xám toàn bộ model...");
  await grayAllObjects(api, modelId);

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
    await setColorInBatches(api, modelId, approvedRuntimeIds, APPROVED_COLOR, "Tô xanh");
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
