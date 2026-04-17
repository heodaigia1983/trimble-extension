let API = null;
let excelRows = [];

const colorMap = {
  "SYSTEM 2": "#4CAF50",
  "SYSTEM 5": "#F44336",
  "": "#BDBDBD"
};

function log(msg) {
  document.getElementById("log").textContent += msg + "\n";
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
  return rows
    .map(r => ({
      guid: String(r.GUID || "").trim(),
      paintCode: String(r["PAINT CODE"] || "").trim()
    }))
    .filter(r => r.guid);
}

async function getLoadedModelId() {
  const api = await initAPI();

  // thử lấy từ viewer objects trước
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
      const modelIds = [...new Set(viewerObjects.modelObjectIds.map(x => x.modelId).filter(Boolean))];
      if (modelIds.length) {
        log("Loaded modelIds trong viewer: " + modelIds.join(", "));
        return modelIds[0];
      }
    }
  } catch (err) {
    log("getObjects fallback: " + (err?.message || String(err)));
  }

  // fallback sang getModels
  const models = await api.viewer.getModels();

  if (!models || !models.length) {
    throw new Error("Không tìm thấy model đang load.");
  }

  log("viewer.getModels(): " + models.map(m => m.id).join(", "));
  return models[0].id;
}

async function colorByPaintCode() {
  const api = await initAPI();

  if (!excelRows.length) {
    log("Chưa có dữ liệu Excel.");
    return;
  }

  const modelId = await getLoadedModelId();
  const rows = normalizeRows(excelRows);

  log("ModelId: " + modelId);
  log("Bắt đầu đổi GUID -> runtimeId...");

  const guids = rows.map(r => r.guid);

  // test 1 GUID trước
  log("Test GUID đầu tiên: " + guids[0]);

  let testRuntimeIds;
  try {
    testRuntimeIds = await api.viewer.convertToObjectRuntimeIds(modelId, [guids[0]]);
    log("Test runtimeIds[0]: " + JSON.stringify(testRuntimeIds));
  } catch (err) {
    log("Lỗi test convert 1 GUID: " + (err?.message || JSON.stringify(err) || String(err)));
    throw err;
  }

  let runtimeIds;
  try {
    runtimeIds = await api.viewer.convertToObjectRuntimeIds(modelId, guids);
  } catch (err) {
    log("Lỗi convert full list: " + (err?.message || JSON.stringify(err) || String(err)));
    throw err;
  }

  const groups = {};
  let matched = 0;
  let unmatched = 0;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const runtimeId = runtimeIds[i];

    if (runtimeId === undefined || runtimeId === null) {
      unmatched++;
      continue;
    }

    const color = colorMap[row.paintCode] || "#2196F3";

    if (!groups[color]) groups[color] = [];
    groups[color].push(runtimeId);
    matched++;
  }

  log("Match: " + matched);
  log("Không match: " + unmatched);

  for (const color in groups) {
    const ids = groups[color];
    if (!ids.length) continue;

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

    log(`Đã tô ${ids.length} object -> ${color}`);
  }

  log("Hoàn tất tô màu.");
}

document.getElementById("readBtn").addEventListener("click", async () => {
  try {
    const file = document.getElementById("fileInput").files[0];
    if (!file) {
      log("Chưa chọn file Excel.");
      return;
    }

    document.getElementById("log").textContent = "";
    await initAPI();

    excelRows = await readExcel(file);
    log(`Đọc xong ${excelRows.length} dòng.`);
    log("5 dòng đầu:");

    excelRows.slice(0, 5).forEach(r => {
      log(`${r.GUID} | ${r["PAINT CODE"]}`);
    });

    await colorByPaintCode();
  } catch (err) {
    console.error(err);
    log("Lỗi: " + (err?.message || JSON.stringify(err) || String(err)));
  }
});
