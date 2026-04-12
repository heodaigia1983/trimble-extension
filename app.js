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

function updateStats({
  totalObjects = "-",
  excelGuidCount = "-",
  greenCount = "-",
  grayCount = "-",
  unmatchedCount = "-"
} = {}) {
  document.getElementById("stats").innerHTML =
    `Tổng object: ${totalObjects}<br />` +
    `GUID trong Excel: ${excelGuidCount}<br />` +
    `Xanh: ${greenCount}<br />` +
    `Xám: ${grayCount}<br />` +
    `Không match: ${unmatchedCount}`;
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

    reader.onload = (e) => {
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

function normalizeExcelGuids(rows) {
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

function chunkArray(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) {
    out.push(arr.slice(i, i + size));
  }
  return out;
}

/*
  getObjects() trả về:
  [
    {
      modelId: "...",
      objects: [
        { id: 123, ... },
        { id: 456, ... }
      ]
    }
  ]
*/
async function getLoadedModelGroups() {
  const api = await initAPI();
  const raw = await api.viewer.getObjects();

  if (!Array.isArray(raw) || !raw.length) {
    throw new Error("Viewer chưa trả về object nào.");
  }

  const groups = raw
    .map(group => {
      const modelId = group?.modelId;
      const runtimeIds = Array.isArray(group?.objects)
        ? [...new Set(
            group.objects
              .map(obj => obj?.id)
              .filter(id => typeof id === "number")
          )]
        : [];

      return { modelId, runtimeIds };
    })
    .filter(group => group.modelId && group.runtimeIds.length > 0);

  if (!groups.length) {
    throw new Error("Không đọc được runtime ids từ viewer.getObjects().");
  }

  log("Loaded modelIds trong viewer: " + groups.map(g => g.modelId).join(", "));
  return groups;
}

async function setObjectColorBatch(api, modelId, runtimeIds, color, label) {
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

async function applyGrayToAllObjects(api, modelGroups) {
  // Lớp 1: thử tô xám toàn cục
  try {
    await api.viewer.setObjectState(undefined, { color: OTHER_COLOR });
    log("Đã tô xám toàn cục bằng selector undefined.");
  } catch (err) {
    log("Tô xám toàn cục lỗi, sẽ fallback sang tô theo runtimeIds: " + (err?.message || String(err)));
  }

  // Lớp 2: tô xám theo toàn bộ runtimeIds lấy đúng từ getObjects()
  for (const group of modelGroups) {
    await setObjectColorBatch(api, group.modelId, group.runtimeIds, OTHER_COLOR, `Tô xám model ${group.modelId}`);
  }
}

async function convertGuidsAcrossModels(api, modelGroups, guids) {
  const matches = new Array(guids.length).fill(null);

  for (const group of modelGroups) {
    log(`Đổi GUID trên model ${group.modelId}...`);

    let runtimeIds;
    try {
      runtimeIds = await api.viewer.convertToObjectRuntimeIds(group.modelId, guids);
    } catch (err) {
      log(`Lỗi convert trên model ${group.modelId}: ${err?.message || String(err)}`);
      continue;
    }

    for (let i = 0; i < guids.length; i++) {
      if (matches[i]) continue;

      const runtimeId = runtimeIds[i];
      if (runtimeId !== undefined && runtimeId !== null) {
        matches[i] = {
          modelId: group.modelId,
          runtimeId: runtimeId,
          guid: guids[i]
        };
      }
    }
  }

  return matches;
}

function buildGreenGroups(matches) {
  const map = new Map();

  for (const item of matches) {
    if (!item) continue;

    if (!map.has(item.modelId)) {
      map.set(item.modelId, new Set());
    }

    map.get(item.modelId).add(item.runtimeId);
  }

  return map;
}

function countAllObjects(modelGroups) {
  return modelGroups.reduce((sum, g) => sum + g.runtimeIds.length, 0);
}

async function applyColorWorkflow() {
  const api = await initAPI();

  if (!excelRows.length) {
    log("Chưa có dữ liệu Excel.");
    return;
  }

  const grayOthers = document.getElementById("grayOthersCheckbox").checked;
  const excelGuids = normalizeExcelGuids(excelRows);
  const modelGroups = await getLoadedModelGroups();
  const totalObjects = countAllObjects(modelGroups);

  log("Tổng object trong viewer: " + totalObjects);
  log("Số GUID duy nhất trong Excel: " + excelGuids.length);

  updateStats({
    totalObjects,
    excelGuidCount: excelGuids.length,
    greenCount: 0,
    grayCount: grayOthers ? totalObjects : "-",
    unmatchedCount: "-"
  });

  if (!excelGuids.length) {
    throw new Error("Excel không có GUID hợp lệ.");
  }

  if (grayOthers) {
    log("Bắt đầu tô xám toàn bộ phần không nằm trong Excel...");
    await applyGrayToAllObjects(api, modelGroups);
  } else {
    log("Bỏ qua bước tô xám phần còn lại.");
  }

  log("Bắt đầu đổi GUID -> runtimeId...");
  log("Test GUID đầu tiên: " + excelGuids[0]);

  const firstModelId = modelGroups[0].modelId;
  const testRuntimeIds = await api.viewer.convertToObjectRuntimeIds(firstModelId, [excelGuids[0]]);
  log("Test runtimeIds[0]: " + JSON.stringify(testRuntimeIds));

  const matches = await convertGuidsAcrossModels(api, modelGroups, excelGuids);
  const greenGroups = buildGreenGroups(matches);

  let matchedCount = 0;
  for (const [, ids] of greenGroups) {
    matchedCount += ids.size;
  }

  const unmatchedCount = excelGuids.length - matchedCount;
  const grayCount = grayOthers ? Math.max(0, totalObjects - matchedCount) : "-";

  log("Match: " + matchedCount);
  log("Không match: " + unmatchedCount);

  for (const [modelId, idSet] of greenGroups.entries()) {
    const ids = Array.from(idSet);
    await setObjectColorBatch(api, modelId, ids, APPROVED_COLOR, `Tô xanh model ${modelId}`);
  }

  updateStats({
    totalObjects,
    excelGuidCount: excelGuids.length,
    greenCount: matchedCount,
    grayCount,
    unmatchedCount
  });

  log("Hoàn tất.");
  if (grayOthers) {
    log("Kết quả: Có trong Excel = xanh | Không có trong Excel = xám");
  } else {
    log("Kết quả: Chỉ tô xanh phần có trong Excel");
  }
}

async function resetColors() {
  const api = await initAPI();

  clearLog();
  log("Đang reset màu...");

  try {
    await api.viewer.setObjectState(undefined, { color: "reset" });
    log("Đã reset màu về mặc định.");
  } catch (err) {
    log("Reset màu lỗi: " + (err?.message || String(err)));
  }

  updateStats({
    totalObjects: "-",
    excelGuidCount: "-",
    greenCount: "-",
    grayCount: "-",
    unmatchedCount: "-"
  });
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

    await applyColorWorkflow();
  } catch (err) {
    console.error(err);
    log("Lỗi: " + (err?.message || JSON.stringify(err) || String(err)));
  }
});

document.getElementById("resetBtn").addEventListener("click", async () => {
  try {
    await resetColors();
  } catch (err) {
    console.error(err);
    log("Lỗi reset: " + (err?.message || JSON.stringify(err) || String(err)));
  }
});
