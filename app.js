let API = null;
let excelRows = [];
let workbookData = null;
let firstSheetName = "";

const APPROVED_COLOR = "#4CAF50";

const WEIGHT_KEYS = [
  "WEIGHT",
  "Weight",
  "weight",
  "MASS",
  "Mass",
  "mass",
  "PROFILE_WEIGHT",
  "ASSEMBLY_WEIGHT",
  "TOTAL_WEIGHT",
  "PART_WEIGHT",
  "MODEL_WEIGHT"
];

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
  totalWeight = "-",
  greenObjectCount = "-",
  greenWeight = "-"
} = {}) {
  const el = document.getElementById("stats");
  if (!el) return;

  el.innerHTML =
    `Tổng số cấu kiện dự án: <strong>${totalObjects}</strong><br />` +
    `Tổng khối lượng dự án: <strong>${totalWeight}</strong><br />` +
    `Số cấu kiện đang triển khai: <strong>${greenObjectCount}</strong><br />` +
    `Khối lượng đang triển khai: <strong>${greenWeight}</strong>`;
}

async function initAPI() {
  if (API) return API;

  API = await TrimbleConnectWorkspace.connect(window.parent, (event, data) => {
    console.log("Trimble event:", event, data);
  });

  log("Connected to Trimble Workspace API.");
  return API;
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const workbook = XLSX.read(e.target.result, { type: "array" });
        resolve(workbook);
      } catch (err) {
        reject(err);
      }
    };

    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function getFirstSheetName(workbook) {
  if (!workbook || !Array.isArray(workbook.SheetNames) || workbook.SheetNames.length === 0) {
    return "";
  }
  return String(workbook.SheetNames[0] || "").trim();
}

function buildSuggestedViewName(sheetName) {
  const now = new Date();
  const dateText =
    now.getFullYear() +
    "-" +
    String(now.getMonth() + 1).padStart(2, "0") +
    "-" +
    String(now.getDate()).padStart(2, "0") +
    " " +
    String(now.getHours()).padStart(2, "0") +
    ":" +
    String(now.getMinutes()).padStart(2, "0");

  return `${sheetName || "Approval"} - ${dateText}`;
}

function fillSuggestedViewName(force = false) {
  const input = document.getElementById("viewNameInput");
  if (!input) return;

  if (force || !input.value.trim() || input.dataset.userEdited !== "1") {
    input.value = buildSuggestedViewName(firstSheetName);
    input.dataset.userEdited = "0";
  }
}

function sheetToRows(workbook, sheetName) {
  if (!workbook) throw new Error("Workbook chưa được nạp.");
  if (!sheetName) throw new Error("Không tìm thấy sheet đầu tiên.");

  const sheet = workbook.Sheets[sheetName];
  if (!sheet) throw new Error(`Không tìm thấy sheet: ${sheetName}`);

  return XLSX.utils.sheet_to_json(sheet, { defval: "" });
}

function normalizeExcelGuids(rows) {
  const seen = new Set();

  return rows
    .map(r => String(r.GUID || "").trim())
    .filter(Boolean)
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
    throw new Error("Viewer returned no objects.");
  }

  const groups = raw
    .map(group => {
      const modelId = group?.modelId;
      const runtimeIds = extractRuntimeIdsFromGroup(group);
      return { modelId, runtimeIds };
    })
    .filter(group => group.modelId && group.runtimeIds.length > 0);

  if (!groups.length) {
    throw new Error("Could not read runtime IDs from viewer.getObjects().");
  }

  log("Loaded model IDs: " + groups.map(g => g.modelId).join(", "));
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

    log(`${label}: batch ${i + 1}/${batches.length} (${ids.length} objects)`);
  }
}

function normalizeConvertedIds(value) {
  if (value === null || value === undefined) return [];

  if (typeof value === "number") return [value];

  if (Array.isArray(value)) {
    return value.flat(Infinity).filter(v => typeof v === "number");
  }

  return [];
}

async function convertGuidsAcrossModels(api, modelGroups, guids) {
  const matches = new Array(guids.length).fill(null);

  for (const group of modelGroups) {
    log(`Converting GUIDs on model ${group.modelId}...`);

    let runtimeIds;
    try {
      runtimeIds = await api.viewer.convertToObjectRuntimeIds(group.modelId, guids);
    } catch (err) {
      log(`Convert error on model ${group.modelId}: ${err?.message || String(err)}`);
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

function parseWeightValue(value) {
  if (typeof value === "number") return value;

  if (typeof value !== "string") return null;

  let s = value.trim();
  if (!s) return null;

  s = s.replace(/,/g, "");
  const match = s.match(/-?\d+(\.\d+)?/);
  if (!match) return null;

  const n = Number(match[0]);
  return Number.isFinite(n) ? n : null;
}

function extractWeightFromObjectProperties(objectProps) {
  const sets = objectProps?.properties || [];
  for (const set of sets) {
    const props = set?.properties || [];
    for (const prop of props) {
      const name = String(prop?.name || "").trim();
      if (!WEIGHT_KEYS.includes(name)) continue;

      const weight = parseWeightValue(prop?.value);
      if (weight !== null) return weight;
    }
  }
  return null;
}

async function sumWeightForIds(api, modelId, runtimeIds, label) {
  let total = 0;

  const batches = chunkArray(runtimeIds, 300);

  for (let i = 0; i < batches.length; i++) {
    const ids = batches[i];
    if (!ids.length) continue;

    let objectProps = [];
    try {
      objectProps = await api.viewer.getObjectProperties(modelId, ids);
    } catch (err) {
      log(`${label} property batch ${i + 1}/${batches.length} failed: ${err?.message || String(err)}`);
      continue;
    }

    for (const obj of objectProps || []) {
      const w = extractWeightFromObjectProperties(obj);
      if (w !== null) total += w;
    }

    log(`${label} weight batch ${i + 1}/${batches.length} done.`);
  }

  return total;
}

async function calculateWeights(api, modelGroups, approvedSetsByModel) {
  let totalModelWeight = 0;
  let greenWeight = 0;

  for (const group of modelGroups) {
    const allIds = group.runtimeIds;
    const greenIds = Array.from(approvedSetsByModel.get(group.modelId) || new Set());

    totalModelWeight += await sumWeightForIds(api, group.modelId, allIds, `Total model ${group.modelId}`);

    if (greenIds.length) {
      greenWeight += await sumWeightForIds(api, group.modelId, greenIds, `Green model ${group.modelId}`);
    }
  }

  return {
    totalModelWeight,
    greenWeight
  };
}

function formatWeight(n) {
  if (!Number.isFinite(n)) return "-";
  return n.toLocaleString(undefined, {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

async function applyColorWorkflow() {
  const api = await initAPI();

  if (!excelRows.length) {
    log("No Excel data loaded.");
    return;
  }

  const excelGuids = normalizeExcelGuids(excelRows);
  const modelGroups = await getLoadedModelGroups();
  const totalObjects = countAllObjects(modelGroups);

  log("Using first sheet: " + firstSheetName);
  log("Total viewer objects: " + totalObjects);
  log("Unique Excel GUIDs: " + excelGuids.length);

  if (!excelGuids.length) {
    throw new Error("Excel does not contain valid GUIDs.");
  }

  const matches = await convertGuidsAcrossModels(api, modelGroups, excelGuids);
  const approvedSetsByModel = buildApprovedSetsByModel(matches);

  let greenObjectCount = 0;
  for (const [, ids] of approvedSetsByModel.entries()) {
    greenObjectCount += ids.size;
  }

  log("Approved objects (green): " + greenObjectCount);

  for (const [modelId, idSet] of approvedSetsByModel.entries()) {
    const ids = Array.from(idSet);
    await setObjectColorBatch(api, modelId, ids, APPROVED_COLOR, `Color green on model ${modelId}`);
  }

  log("Calculating weights...");
  const weightSummary = await calculateWeights(api, modelGroups, approvedSetsByModel);

  updateStats({
    totalObjects,
    totalWeight: formatWeight(weightSummary.totalModelWeight),
    greenObjectCount,
    greenWeight: formatWeight(weightSummary.greenWeight)
  });

  log("Done.");
  log("Result: Objects found in Excel are highlighted in green.");
}

async function resetViewer() {
  const api = await initAPI();
  clearLog();
  log("Resetting viewer...");

  try {
    await api.viewer.setObjectState(undefined, { color: "reset", visible: "reset" });
  } catch (err) {
    log("Color reset fallback: " + (err?.message || String(err)));
  }

  await api.viewer.reset();
  updateStats();
  log("Viewer reset completed.");
}

async function saveCurrentView() {
  const api = await initAPI();

  let viewName = String(document.getElementById("viewNameInput")?.value || "").trim();
  if (!viewName) {
    viewName = buildSuggestedViewName(firstSheetName);
    const input = document.getElementById("viewNameInput");
    if (input) input.value = viewName;
  }

  const description = `Saved from Model Approval Colorizer | Sheet: ${firstSheetName || "-"} | Developed by Le Van Thao`;

  const createdView = await api.view.createView({
    name: viewName,
    description: description
  });

  if (!createdView?.id) {
    throw new Error("Create view succeeded but no view ID was returned.");
  }

  await api.view.updateView({ id: createdView.id });
  await api.view.selectView(createdView.id);

  log(`View saved and opened: ${createdView.name || viewName}`);
}

document.getElementById("fileInput").addEventListener("change", async (event) => {
  try {
    const file = event.target.files?.[0];
    if (!file) return;

    workbookData = await readWorkbook(file);
    firstSheetName = getFirstSheetName(workbookData);

    if (!firstSheetName) {
      throw new Error("File Excel không có sheet.");
    }

    fillSuggestedViewName(true);

    log("Workbook loaded.");
    log("Using first sheet: " + firstSheetName);
  } catch (err) {
    console.error(err);
    log("Workbook load error: " + (err?.message || JSON.stringify(err) || String(err)));
  }
});

document.getElementById("viewNameInput").addEventListener("input", (e) => {
  e.target.dataset.userEdited = e.target.value.trim() ? "1" : "0";
});

document.getElementById("readBtn").addEventListener("click", async () => {
  try {
    const file = document.getElementById("fileInput").files[0];

    if (!file) {
      log("Please choose an Excel file.");
      return;
    }

    if (!workbookData) {
      workbookData = await readWorkbook(file);
      firstSheetName = getFirstSheetName(workbookData);
      fillSuggestedViewName(true);
    }

    if (!firstSheetName) {
      log("Could not find the first sheet.");
      return;
    }

    clearLog();
    updateStats();
    await initAPI();

    excelRows = sheetToRows(workbookData, firstSheetName);

    log(`Using first sheet: ${firstSheetName}`);
    log(`Excel rows loaded: ${excelRows.length}`);

    await applyColorWorkflow();
  } catch (err) {
    console.error(err);
    log("Error: " + (err?.message || JSON.stringify(err) || String(err)));
  }
});

document.getElementById("resetBtn").addEventListener("click", async () => {
  try {
    await resetViewer();
  } catch (err) {
    console.error(err);
    log("Reset error: " + (err?.message || JSON.stringify(err) || String(err)));
  }
});

document.getElementById("saveViewBtn").addEventListener("click", async () => {
  try {
    await saveCurrentView();
  } catch (err) {
    console.error(err);
    log("Save View error: " + (err?.message || JSON.stringify(err) || String(err)));
  }
});

updateStats();
