let API = null;
let excelRows = [];

const APPROVED_COLOR = "#4CAF50"; // xanh
const OTHER_COLOR = "#BDBDBD";    // xám

function log(msg) {
  const el = document.getElementById("log");
  if (el) el.textContent += msg + "\n";
  console.log(msg);
}

function clearLog() {
  const el = document.getElementById("log");
  if (el) el.textContent = "";
}

function updateStats({
  totalObjects = "-",
  excelGuidCount = "-",
  greenCount = "-",
  grayCount = "-",
  unmatchedCount = "-"
} = {}) {
  const el = document.getElementById("stats");
  if (!el) return;

  el.innerHTML =
    `Tổng object: ${totalObjects}<br />` +
    `GUID trong Excel: ${excelGuidCount}<br />` +
    `Xanh: ${greenCount}<br />` +
    `Xám: ${grayCount}<br />` +
    `Không match: ${unmatchedCount}`;
}

function shouldGrayOthers() {
  const checkbox = document.getElementById("grayOthersCheckbox");
  return checkbox ? checkbox.checked : true;
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

function extractRuntimeIdsFromGroup(group) {
  const ids = new Set();

  if (!group) return [];

  if (Array.isArray(group.objects)) {
    for (const obj of group.objects) {
      if (typeof obj?.id === "number") ids.add(obj.id);
    }
  }

  return Array.from(ids);
}

async function getLoadedModelGroups() {
  const api = await initAPI();
  const raw = await api.viewer.getObjects();

  if (!Array.isArray(raw) || !raw.length) {
    throw new Error("Viewer chưa trả về object nào.");
  }

  const groups = raw
    .map(group => {
      const modelId = group?.modelId;
      const runtimeIds = extractRuntimeIdsFromGroup(group);
      return { modelId, runtimeIds };
    })
    .filter(group => group.modelId && group.runtimeIds.length > 0);

  if (!groups.length) {
    throw new Error("Không đọc được runtime ids từ viewer.getObjects().");
  }

  log("Loaded modelIds trong viewer: " + groups.map(g => g.modelId).join(", "));
  return groups;
}

function countAllObjects(modelGroups) {
  return modelGroups.reduce((sum, g) => sum + g.runtimeIds.length, 0);
}

async function setObjectColorBatch(api, modelId, runtimeIds, color, label) {
  const batches = chunkArray(runtimeIds, 1000);

  for (let i = 0; i < batches.length; i++) {
    const ids = batches[i];
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

    log(`${label}: batch ${i + 1}/${batches.length} (${ids.length} object)`);
  }
}

function normalizeConvertedIds(value) {
  if (value === null || value === undefined) return [];

  if (typeof value === "number") return [value];

  if (Array.isArray(value)) {
    return value
      .flat(Infinity)
      .filter(v => typeof v === "number");
  }

  return [];
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

    if (!Array.isArray(runtimeIds)) continue;

    for (let i = 0; i < guids.length; i++) {
      if (matches[i]) continue;

      const ids = normalizeConvertedIds(runtimeIds[i]);

      if (ids.length > 0) {
        matches[i] = {
          modelId: group.modelId,
          runtimeIds: ids,
          guid: guids[i]
        };
      }
    }
  }

  return matches;
}

function buildApprovedSetsByModel(matches) {
  const map = new Map();

  for (const item of matches) {
    if (!item) continue;

    if (!map.has(item.modelId)) {
      map.set(item.modelId, new Set());
    }

    const set = map.get(item.modelId);
    item.runtimeIds.forEach(id => set.add(id));
  }

  return map;
}

function buildOtherIdsByModel(modelGroups, approvedSetsByModel) {
  const result = new Map();

  for (const group of modelGroups) {
    const approvedSet = approvedSetsByModel.get(group.modelId) || new Set();
    const others = group.runtimeIds.filter(id => !approvedSet.has(id));
    result.set(group.modelId, others);
  }

  return result;
}

async function applyColorWorkflow() {
  const api = await initAPI();

  if (!excelRows.length) {
    log("Chưa có dữ liệu Excel.");
    return;
  }

  const grayOthers = shouldGrayOthers();
  const excelGuids = normalizeExcelGuids(excelRows);
  const modelGroups = await getLoadedModelGroups();
  const totalObjects = countAllObjects(modelGroups);

  log("Tổng object trong viewer: " + totalObjects);
  log("Số GUID duy nhất trong Excel: " + excelGuids.length);

  if (!excelGuids.length) {
    throw new Error("Excel không có GUID hợp lệ.");
  }

  log("Bắt đầu đổi GUID -> runtimeId...");
  log("Test GUID đầu tiên: " + excelGuids[0]);

  try {
    const testRuntimeIds = await api.viewer.convertToObjectRuntimeIds(
      modelGroups[0].modelId,
      [excelGuids[0]]
    );
    log("Test runtimeIds[0]: " + JSON.stringify(testRuntimeIds));
  } catch (err) {
    log("Test convert lỗi: " + (err?.message || String(err)));
  }

  const matches = await convertGuidsAcrossModels(api, modelGroups, excelGuids);
  const approvedSetsByModel = buildApprovedSetsByModel(matches);

  let matchedGuidCount = 0;
  for (const item of matches) {
    if (item) matchedGuidCount++;
  }

  let greenObjectCount = 0;
  for (const [, ids] of approvedSetsByModel.entries()) {
    greenObjectCount += ids.size;
  }

  const unmatchedCount = excelGuids.length - matchedGuidCount;
  const otherIdsByModel = buildOtherIdsByModel(modelGroups, approvedSetsByModel);

  let grayCount = 0;
  for (const [, ids] of otherIdsByModel.entries()) {
    grayCount += ids.length;
  }

  log("GUID match: " + matchedGuidCount);
  log("GUID không match: " + unmatchedCount);
  log("Object xanh thực tế: " + greenObjectCount);
  log("Object xám thực tế: " + grayCount);

  if (grayOthers) {
    for (const [modelId, ids] of otherIdsByModel.entries()) {
      await setObjectColorBatch(api, modelId, ids, OTHER_COLOR, `Tô xám model ${modelId}`);
    }
  } else {
    grayCount = "-";
    log("Bỏ qua bước tô xám phần còn lại.");
  }

  for (const [modelId, idSet] of approvedSetsByModel.entries()) {
    const ids = Array.from(idSet);
    await setObjectColorBatch(api, modelId, ids, APPROVED_COLOR, `Tô xanh model ${modelId}`);
  }

  updateStats({
    totalObjects,
    excelGuidCount: excelGuids.length,
    greenCount: greenObjectCount,
    grayCount,
    unmatchedCount
  });

  log("Hoàn tất.");
  log(grayOthers
    ? "Kết quả: Có trong Excel = xanh | Không có trong Excel = xám"
    : "Kết quả: Chỉ tô xanh phần có trong Excel");
}

async function resetColors() {
  const api = await initAPI();
  clearLog();
  log("Đang reset viewer...");

  await api.viewer.reset();
  updateStats();

  log("Đã reset model/camera/tools về mặc định.");
}

document.getElementById("readBtn").addEventListener("click", async () => {
  try {
    const file = document.getElementById("fileInput").files[0];

    if (!file) {
      log("Chưa chọn file Excel.");
      return;
    }

    clearLog();
    updateStats();
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

updateStats();
