/* ════════════════════════════════════════════════════════════
   Jira Asset Manager — Office Add-in
   taskpane.js — v1.4 stream-write
   ════════════════════════════════════════════════════════════ */
"use strict";

// ── COLUMN MAP ────────────────────────────────────────────────
const COL = {
  ASSET_ID:    0,
  ASSET_KEY:   1,
  HOSTNAME:    2,
  SERIAL:      3,
  STATUS:      4,
  LOCATION:    5,
  REGION:      6,
  MANUFACTURER:7,
  MODEL:       8,
  OS:          9,
  OS_VERSION:  10,
  OS_BUILD:    11,
  CPU:         12,
  IP:          13,
  MAC:         14,
  NETWORK:     15,
  ANTIVIRUS:   16,
  USERNAME:    17,
  ASSIGNED:    18,
  FIRST_SEEN:  19,
  LAST_SEEN:   20,
  PURCHASE:    21,
  WARRANTY:    22,
  TENANT_ID:   23,
  LANSWEEPER:  24,
  DAYS:        25,
  NOTE:        26,
  CASE_JIRA:   27,
  VALIDATION:  28,
  SYNC_STATUS: 29,
  LAST_SYNC:   30,
  ACTION:      31,
};
const COL_COUNT = 32;
const HEADERS = [
  "Asset ID","Asset Key","Hostname","Serial Number",
  "Status","Location","Region","Manufacturer","Model",
  "Operating System","Windows Version","Windows Build","CPU",
  "IP Address","MAC Address","Network Name","Antivirus",
  "Username","Assigned User","First Seen","Last Seen",
  "Purchase Date","Warranty Expire","Tenant ID / Source ID","Lansweeper URL",
  "Days","Note","Case Jira","Validation","Sync Status","Last Sync","Action"
];

const CFG_KEYS = {
  JIRA_URL:     "jiraUrl",
  EMAIL:        "jiraEmail",
  TOKEN:        "jiraToken",
  CLOUD_ID:     "cloudId",
  WORKSPACE_ID: "workspaceId",
  PROJECT_KEY:  "projectKey",
  WORKER_URL:   "workerUrl",
  AQL_QUERY:    "aqlQuery",
  LAST_SYNC:    "lastSync",
};

let cfg = {};
let isSyncing = false;

// ══════════════════════════════════════════════════════════════
// INIT
// ══════════════════════════════════════════════════════════════
Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) return;
  loadConfig();
  populateConfigUI();
  wireEvents();
  await refreshDashboard();
});

// ══════════════════════════════════════════════════════════════
// CONFIG
// ══════════════════════════════════════════════════════════════
function loadConfig() {
  const s = Office.context.document.settings;
  cfg = {
    jiraUrl:     s.get(CFG_KEYS.JIRA_URL)     || "",
    email:       s.get(CFG_KEYS.EMAIL)         || "",
    token:       s.get(CFG_KEYS.TOKEN)         || "",
    cloudId:     s.get(CFG_KEYS.CLOUD_ID)      || "",
    workspaceId: s.get(CFG_KEYS.WORKSPACE_ID)  || "",
    projectKey:  s.get(CFG_KEYS.PROJECT_KEY)   || "IT",
    workerUrl:   s.get(CFG_KEYS.WORKER_URL)    || "",
    aqlQuery:    s.get(CFG_KEYS.AQL_QUERY)     || "objectTypeId IN (525,527,529)",
    lastSync:    s.get(CFG_KEYS.LAST_SYNC)     || null,
  };
  updateWorkspaceLabel();
}

function updateWorkspaceLabel() {
  const el = document.getElementById("ws-label");
  if (!el) return;
  try {
    const raw = cfg.jiraUrl.startsWith("http") ? cfg.jiraUrl : "https://" + cfg.jiraUrl;
    el.textContent = new URL(raw).hostname.split(".")[0].toUpperCase();
  } catch { el.textContent = "—"; }
}

function saveConfig() {
  const s = Office.context.document.settings;
  cfg.jiraUrl     = getVal("cfg-jira-url");
  cfg.email       = getVal("cfg-email");
  cfg.token       = getVal("cfg-token");
  cfg.cloudId     = getVal("cfg-cloud-id");
  cfg.workspaceId = getVal("cfg-workspace-id");
  cfg.projectKey  = getVal("cfg-project-key");
  cfg.workerUrl   = getVal("cfg-worker-url");
  cfg.aqlQuery    = getVal("cfg-aql-query");
  s.set(CFG_KEYS.JIRA_URL,     cfg.jiraUrl);
  s.set(CFG_KEYS.EMAIL,        cfg.email);
  s.set(CFG_KEYS.TOKEN,        cfg.token);
  s.set(CFG_KEYS.CLOUD_ID,     cfg.cloudId);
  s.set(CFG_KEYS.WORKSPACE_ID, cfg.workspaceId);
  s.set(CFG_KEYS.PROJECT_KEY,  cfg.projectKey);
  s.set(CFG_KEYS.WORKER_URL,   cfg.workerUrl);
  s.set(CFG_KEYS.AQL_QUERY,    cfg.aqlQuery);
  s.saveAsync(() => { updateWorkspaceLabel(); toast("Settings saved", "success"); });
}

function populateConfigUI() {
  setVal("cfg-jira-url",     cfg.jiraUrl);
  setVal("cfg-email",        cfg.email);
  setVal("cfg-token",        cfg.token);
  setVal("cfg-cloud-id",     cfg.cloudId);
  setVal("cfg-workspace-id", cfg.workspaceId);
  setVal("cfg-project-key",  cfg.projectKey);
  setVal("cfg-worker-url",   cfg.workerUrl);
  setVal("cfg-aql-query",    cfg.aqlQuery);
}

// ══════════════════════════════════════════════════════════════
// EVENTS
// ══════════════════════════════════════════════════════════════
function wireEvents() {
  document.querySelectorAll(".tab").forEach(t =>
    t.addEventListener("click", () => switchTab(t.dataset.tab))
  );
  on("btn-sync-now",       () => runSync());
  on("btn-validate-all",   () => runValidation());
  on("btn-open-jira",      openJira);
  on("btn-create-sheets",  createLocationSheets);
  on("btn-run-validate",   () => runValidation());
  on("btn-scan-pending",   scanPendingRows);
  on("btn-create-tickets", createTickets);
  on("btn-full-sync",      () => runSync());
  on("btn-sync-local",     matchLocalAssets);
  on("btn-save-cfg",       saveConfig);
  on("btn-test-conn",      testConnection);
}

function on(id, fn) {
  const el = document.getElementById(id);
  if (el) el.addEventListener("click", fn);
}

function switchTab(tab) {
  document.querySelectorAll(".tab").forEach(t =>
    t.classList.toggle("active", t.dataset.tab === tab)
  );
  document.querySelectorAll(".panel").forEach(p =>
    p.classList.toggle("active", p.id === `panel-${tab}`)
  );
  if (tab === "ticket") scanPendingRows();
}

// ══════════════════════════════════════════════════════════════
// API
// ══════════════════════════════════════════════════════════════
function jiraBase()   { return cfg.jiraUrl.replace(/\/+$/, ""); }
function assetsBase() {
  return `https://api.atlassian.com/ex/jira/${cfg.cloudId}/jsm/assets/workspace/${cfg.workspaceId}/v1`;
}
function proxyUrl(targetUrl) {
  return `${cfg.workerUrl.replace(/\/+$/, "")}/proxy?url=${encodeURIComponent(targetUrl)}`;
}
function jiraHeaders() {
  return {
    "Authorization": `Basic ${btoa(cfg.email + ":" + cfg.token)}`,
    "Content-Type":  "application/json",
    "Accept":        "application/json",
  };
}

async function jiraGet(path) {
  const res = await fetch(proxyUrl(`${jiraBase()}/rest${path}`), { headers: jiraHeaders() });
  if (!res.ok) throw new Error(`Jira ${res.status}: ${(await res.text().catch(()=>"")).slice(0,120)}`);
  return res.json();
}

async function jiraPost(path, body) {
  const res = await fetch(proxyUrl(`${jiraBase()}/rest${path}`), {
    method: "POST", headers: jiraHeaders(), body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`Jira ${res.status}: ${(await res.text().catch(()=>"")).slice(0,120)}`);
  return res.json();
}

async function assetsPost(path, body) {
  const res = await fetch(proxyUrl(`${assetsBase()}${path}`), {
    method: "POST", headers: jiraHeaders(), body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`Assets ${res.status}: ${(await res.text().catch(()=>"")).slice(0,120)}`);
  return res.json();
}

// ══════════════════════════════════════════════════════════════
// PARSE
// ══════════════════════════════════════════════════════════════
function parseAsset(obj) {
  const attrById = (id) => {
    const a = (obj.attributes || []).find(x => String(x.objectTypeAttributeId) === String(id));
    return a?.objectAttributeValues?.[0]?.displayValue || "";
  };
  return {
    id:           String(obj.id ?? obj.objectId ?? ""),
    key:          obj.objectKey || obj.key || "",
    hostname:     attrById(1737) || obj.label || "",
    serial:       attrById(5194),
    status:       attrById(5052),
    location:     attrById(30125),
    region:       attrById(27292),
    manufacturer: attrById(6608),
    model:        attrById(6609),
    os:           attrById(30345),
    osVersion:    attrById(27291),
    osBuild:      attrById(27290),
    cpu:          attrById(6610),
    ip:           attrById(5208),
    mac:          attrById(5209),
    network:      attrById(5210),
    antivirus:    attrById(6612),
    username:     attrById(5200),
    assigned:     attrById(26690),
    firstSeen:    attrById(5205),
    lastSeen:     attrById(5206),
    purchase:     attrById(5203),
    warranty:     attrById(6615),
    tenantId:     attrById(26398),
    lansweeper:   attrById(5207),
  };
}

function parseTypeIds() {
  const aql = (cfg.aqlQuery || "").trim();
  const m   = aql.match(/objectTypeId\s+IN\s*\(([^)]+)\)/i);
  if (m) return m[1].split(",").map(s => s.trim()).filter(Boolean);
  const m2  = aql.match(/objectTypeId\s*=\s*(\d+)/i);
  if (m2) return [m2[1]];
  return [];
}

// ══════════════════════════════════════════════════════════════
// SHEET HELPERS
// ══════════════════════════════════════════════════════════════
function locationSheetName(loc) {
  return "_" + loc.replace(/[^a-zA-Z0-9]/g, "_").toUpperCase();
}

async function ensureSheet(context, name) {
  let sheet = context.workbook.worksheets.getItemOrNullObject(name);
  await context.sync();
  if (sheet.isNullObject) {
    sheet = context.workbook.worksheets.add(name);
    await context.sync();
  }
  return sheet;
}

async function ensureHeaders(context, sheet) {
  const r = sheet.getRangeByIndexes(0, 0, 1, COL_COUNT);
  r.load("values");
  await context.sync();
  if (!r.values[0] || r.values[0][0] !== "Asset ID") {
    r.values = [HEADERS];
    r.format.font.bold  = true;
    r.format.fill.color = "#1a1d27";
    r.format.font.color = "#8892a4";
    sheet.freezePanes.freezeRows(1);
    await context.sync();
  }
}

async function readSheetRows(context, sheet) {
  const used = sheet.getUsedRangeOrNullObject(true);
  await context.sync();
  if (used.isNullObject) return [];
  used.load(["values","rowCount"]);
  await context.sync();
  if (used.rowCount <= 1) return [];
  return used.values.slice(1);
}

async function getLocationSheets(context) {
  context.workbook.worksheets.load("items/name");
  await context.sync();
  return context.workbook.worksheets.items
    .filter(s => s.name.startsWith("_"))
    .map(s => s.name);
}

// ══════════════════════════════════════════════════════════════
// STREAM WRITE — ghi ngay từng batch vào Excel
// Không cần load hết toàn bộ asset vào memory trước
// ══════════════════════════════════════════════════════════════

// Cache sheet state để tránh đọc lại mỗi lần ghi
const sheetCache = {}; // sheetName → { byId, bySerial, nextRow, jiraIdsSeen }

async function initSheetCache(context, sheetName) {
  if (sheetCache[sheetName]) return;
  const sheet    = await ensureSheet(context, sheetName);
  await ensureHeaders(context, sheet);
  const existing = await readSheetRows(context, sheet);
  const byId     = {}, bySerial = {};
  existing.forEach((row, idx) => {
    const id  = String(row[COL.ASSET_ID] || "").trim();
    const ser = String(row[COL.SERIAL]   || "").trim();
    if (id)  byId[id]      = idx;
    if (ser) bySerial[ser] = idx;
  });
  sheetCache[sheetName] = {
    byId,
    bySerial,
    existing,
    nextRow:     existing.length + 1, // +1 for header
    jiraIdsSeen: new Set(),            // track IDs đã sync lần này
  };
}

function assetToRow(asset, now, existingRow) {
  const row = existingRow ? [...existingRow] : Array(COL_COUNT).fill("");
  row[COL.ASSET_ID]     = asset.id;
  row[COL.ASSET_KEY]    = asset.key;
  row[COL.HOSTNAME]     = asset.hostname;
  row[COL.SERIAL]       = asset.serial;
  row[COL.STATUS]       = asset.status;
  row[COL.LOCATION]     = asset.location;
  row[COL.REGION]       = asset.region;
  row[COL.MANUFACTURER] = asset.manufacturer;
  row[COL.MODEL]        = asset.model;
  row[COL.OS]           = asset.os;
  row[COL.OS_VERSION]   = asset.osVersion;
  row[COL.OS_BUILD]     = asset.osBuild;
  row[COL.CPU]          = asset.cpu;
  row[COL.IP]           = asset.ip;
  row[COL.MAC]          = asset.mac;
  row[COL.NETWORK]      = asset.network;
  row[COL.ANTIVIRUS]    = asset.antivirus;
  row[COL.USERNAME]     = asset.username;
  row[COL.ASSIGNED]     = asset.assigned;
  row[COL.FIRST_SEEN]   = asset.firstSeen;
  row[COL.LAST_SEEN]    = asset.lastSeen;
  row[COL.PURCHASE]     = asset.purchase;
  row[COL.WARRANTY]     = asset.warranty;
  row[COL.TENANT_ID]    = asset.tenantId;
  row[COL.LANSWEEPER]   = asset.lansweeper;
  row[COL.SYNC_STATUS]  = "JIRA";
  row[COL.LAST_SYNC]    = now;
  return row;
}

// Ghi 1 batch assets vào đúng sheet của chúng
async function writeBatch(batch, now) {
  // Group batch theo sheet
  const bySheet = {};
  batch.forEach(asset => {
    const loc       = (asset.location || "UNKNOWN").trim();
    const sheetName = locationSheetName(loc);
    if (!bySheet[sheetName]) bySheet[sheetName] = [];
    bySheet[sheetName].push(asset);
  });

  await Excel.run(async (context) => {
    // Init cache cho các sheet chưa có
    for (const sheetName of Object.keys(bySheet)) {
      await initSheetCache(context, sheetName);
    }

    for (const [sheetName, assets] of Object.entries(bySheet)) {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const cache = sheetCache[sheetName];

      const toInsert = [];

      assets.forEach(asset => {
        cache.jiraIdsSeen.add(asset.id);

        const existingIdx =
          cache.byId[asset.id]                                    !== undefined ? cache.byId[asset.id] :
          asset.serial && cache.bySerial[asset.serial]            !== undefined ? cache.bySerial[asset.serial] :
          -1;

        if (existingIdx >= 0) {
          // UPDATE tại chỗ
          const newRow = assetToRow(asset, now, cache.existing[existingIdx]);
          sheet.getRangeByIndexes(existingIdx + 1, 0, 1, COL_COUNT).values = [newRow];
          // Xoá "Not in Jira" nếu asset quay lại
          if (String(cache.existing[existingIdx][COL.VALIDATION] || "") === "Not in Jira") {
            sheet.getRangeByIndexes(existingIdx + 1, COL.VALIDATION, 1, 1).values = [[""]];
          }
          // Cập nhật cache để tránh duplicate insert
          cache.existing[existingIdx] = newRow;
        } else {
          // INSERT mới — thêm vào batch insert
          const newRow = assetToRow(asset, now, null);
          toInsert.push(newRow);
          // Cập nhật cache ngay
          const newIdx = cache.existing.length;
          cache.byId[asset.id] = newIdx;
          if (asset.serial) cache.bySerial[asset.serial] = newIdx;
          cache.existing.push(newRow);
        }
      });

      // Batch insert tất cả rows mới cùng 1 lúc
      if (toInsert.length > 0) {
        sheet.getRangeByIndexes(cache.nextRow, 0, toInsert.length, COL_COUNT).values = toInsert;
        cache.nextRow += toInsert.length;
      }
    }

    await context.sync();
  });
}

// Sau khi sync xong: mark rows không còn trong Jira
async function markMissingAssets(now) {
  await Excel.run(async (context) => {
    for (const [sheetName, cache] of Object.entries(sheetCache)) {
      const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
      await context.sync();
      if (sheet.isNullObject) continue;

      cache.existing.forEach((row, idx) => {
        if (row[COL.SYNC_STATUS] === "LOCAL") return;
        const id = String(row[COL.ASSET_ID] || "").trim();
        if (!id || cache.jiraIdsSeen.has(id)) return;
        // Asset không còn trong Jira lần sync này
        const cell = sheet.getRangeByIndexes(idx + 1, COL.VALIDATION, 1, 1);
        cell.values = [["Not in Jira"]];
        cell.format.font.color = "#f59e0b";
      });
    }
    await context.sync();
  });
}

// ══════════════════════════════════════════════════════════════
// SYNC — stream fetch+write từng page
// ══════════════════════════════════════════════════════════════
async function runSync() {
  if (isSyncing) { toast("Sync already running", "warning"); return; }
  if (!cfg.jiraUrl || !cfg.token || !cfg.workerUrl) {
    toast("Kiểm tra lại Settings", "warning"); return;
  }

  isSyncing = true;
  setSyncIndicator("syncing", "Syncing...");

  // Reset cache mỗi lần sync
  Object.keys(sheetCache).forEach(k => delete sheetCache[k]);

  const syncPanel = document.getElementById("sync-location-list");
  syncPanel.innerHTML = `<div class="empty-state"><div class="spinner"></div>Đang sync...</div>`;

  const typeIds = parseTypeIds();
  if (!typeIds.length) {
    toast("AQL Query chưa đúng", "error");
    isSyncing = false;
    return;
  }

  const now         = new Date().toISOString();
  const globalSeen  = new Set(); // dedup across typeIds
  let   totalWritten = 0;

  try {
    for (let ti = 0; ti < typeIds.length; ti++) {
      const typeId    = typeIds[ti];
      const pageSize  = 25;
      const API_LIMIT = 1000;
      let   windowStart = 0;

      toast(`[${ti+1}/${typeIds.length}] Đang sync objectTypeId=${typeId}...`, "warning");

      // Sliding window: mỗi window tối đa 1000 records
      while (true) {
        let pageStart    = windowStart;
        let windowTotal  = null;
        let windowLoaded = 0;

        // Load từng page trong window, ghi ngay từng batch
        while (true) {
          const data   = await assetsPost("/object/aql", {
            qlQuery:           `objectTypeId = ${typeId}`,
            startAt:           pageStart,
            maxResults:        pageSize,
            includeAttributes: true,
          });

          const values = data.values || [];

          // Lấy total của window từ page đầu tiên
          if (windowTotal === null) {
            windowTotal = typeof data.total === "number" ? data.total : 0;
            console.log(`typeId=${typeId} window=${windowStart}: total=${windowTotal}`);
          }

          if (values.length === 0) break;

          // Parse + dedup + ghi ngay
          const batch = [];
          values.forEach(obj => {
            const asset = parseAsset(obj);
            const key   = asset.key || asset.id;
            if (key && globalSeen.has(key)) return;
            if (key) globalSeen.add(key);
            batch.push(asset);
          });

          if (batch.length > 0) {
            await writeBatch(batch, now);
            totalWritten += batch.length;
          }

          pageStart    += values.length;
          windowLoaded += values.length;

          // Update UI
          syncPanel.innerHTML = `<div class="info-box" style="font-size:11px">
            <strong>[${ti+1}/${typeIds.length}] typeId=${typeId}</strong><br>
            Window: ${windowStart}→${windowStart+windowTotal}<br>
            Page: ${pageStart} / ${windowStart+windowTotal}<br>
            Total ghi: <strong>${totalWritten}</strong>
          </div>`;

          console.log(`typeId=${typeId} pageStart=${pageStart}/${windowStart+windowTotal} written=${totalWritten}`);

          // Dừng khi đủ window
          if (windowTotal !== null && windowLoaded >= windowTotal) break;
          if (values.length < pageSize) break;
        }

        // Quyết định chạy window tiếp hay dừng
        if (windowTotal === null || windowTotal < API_LIMIT) break;
        windowStart += API_LIMIT;
        console.log(`typeId=${typeId}: window full, next windowStart=${windowStart}`);
      }
    }

    // Mark assets không còn trong Jira
    toast("Đang kiểm tra assets bị xoá khỏi Jira...", "warning");
    await markMissingAssets(now);

    // Push LOCAL rows lên Jira
    toast("Đang push LOCAL assets lên Jira...", "warning");
    await pushLocalAssets();

    cfg.lastSync = new Date().toISOString();
    Office.context.document.settings.set(CFG_KEYS.LAST_SYNC, cfg.lastSync);
    Office.context.document.settings.saveAsync();

    setSyncIndicator("ok", "Synced");
    setInner("last-sync-time", formatTime(cfg.lastSync));
    toast(`Sync hoàn tất — ${totalWritten} assets`, "success");
    await refreshDashboard();

  } catch(e) {
    setSyncIndicator("ok", "Sync failed");
    toast("Sync error: " + e.message, "error");
    console.error("runSync:", e);
  } finally {
    isSyncing = false;
  }
}

// ══════════════════════════════════════════════════════════════
// PUSH LOCAL → JIRA
// ══════════════════════════════════════════════════════════════
async function pushLocalAssets() {
  if (!cfg.cloudId || !cfg.workspaceId) return;
  const typeIds       = parseTypeIds();
  const defaultTypeId = typeIds[0] || "";
  if (!defaultTypeId) return;

  let pushed = 0, failed = 0;
  try {
    await Excel.run(async (context) => {
      const sheets = await getLocationSheets(context);
      for (const sheetName of sheets) {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const rows  = await readSheetRows(context, sheet);
        for (let i = 0; i < rows.length; i++) {
          const row = rows[i];
          if (row[COL.SYNC_STATUS] !== "LOCAL") continue;
          if (row[COL.ASSET_ID])               continue;
          const hostname = String(row[COL.HOSTNAME] || "").trim();
          const serial   = String(row[COL.SERIAL]   || "").trim();
          if (!hostname && !serial) continue;
          try {
            const attrs = [
              { objectTypeAttributeId: 1737,  objectAttributeValues: [{ value: hostname }] },
              { objectTypeAttributeId: 5194,  objectAttributeValues: [{ value: serial }] },
              { objectTypeAttributeId: 30125, objectAttributeValues: [{ value: row[COL.LOCATION] || "" }] },
              { objectTypeAttributeId: 5200,  objectAttributeValues: [{ value: row[COL.USERNAME] || "" }] },
            ].filter(a => a.objectAttributeValues[0].value);
            const res = await assetsPost("/object/create", { objectTypeId: defaultTypeId, attributes: attrs });
            if (res?.id) {
              const range = sheet.getRangeByIndexes(i + 1, 0, 1, COL_COUNT);
              range.load("values");
              await context.sync();
              const cur = range.values[0];
              cur[COL.ASSET_ID]    = String(res.id);
              cur[COL.ASSET_KEY]   = res.objectKey || "";
              cur[COL.SYNC_STATUS] = "JIRA";
              cur[COL.VALIDATION]  = "OK";
              cur[COL.LAST_SYNC]   = new Date().toISOString();
              range.values = [cur];
              await context.sync();
              pushed++;
            }
          } catch(e) { failed++; console.warn(`pushLocal row ${i+2}:`, e.message); }
        }
      }
    });
    if (pushed > 0 || failed > 0)
      toast(`LOCAL push: ${pushed} OK${failed ? ", "+failed+" lỗi" : ""}`, failed ? "warning" : "success");
  } catch(e) { console.warn("pushLocalAssets:", e.message); }
}

// ══════════════════════════════════════════════════════════════
// DASHBOARD
// ══════════════════════════════════════════════════════════════
async function refreshDashboard() {
  try {
    await Excel.run(async (context) => {
      const sheets = await getLocationSheets(context);
      let total = 0, pending = 0, mismatch = 0, local = 0;
      const locSummary = [];
      for (const sheetName of sheets) {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const rows  = await readSheetRows(context, sheet);
        total += rows.length;
        let locLocal = 0;
        rows.forEach(r => {
          if (r[COL.SYNC_STATUS] === "LOCAL")                         { local++; locLocal++; }
          if (r[COL.VALIDATION]  && r[COL.VALIDATION] !== "OK")       mismatch++;
          if (r[COL.ACTION] === "Create Ticket" && !r[COL.CASE_JIRA]) pending++;
        });
        locSummary.push({ name: sheetName, count: rows.length, local: locLocal });
      }
      setInner("stat-total",    total    || "0");
      setInner("stat-pending",  pending  || "0");
      setInner("stat-mismatch", mismatch || "0");
      setInner("stat-local",    local    || "0");
      if (cfg.lastSync) setInner("last-sync-time", formatTime(cfg.lastSync));
      const el = document.getElementById("location-summary");
      if (!locSummary.length) {
        el.innerHTML = `<div class="empty-state"><div class="icon">🗂</div>Sync to load locations</div>`;
      } else {
        el.innerHTML = locSummary.map(l => `
          <div class="location-item">
            <div class="loc-header">
              <span class="loc-name">${l.name}</span>
              <span class="loc-status done">${l.count} assets</span>
            </div>
            ${l.local > 0 ? `<div class="loc-count">⚠ ${l.local} LOCAL only</div>` : ""}
          </div>`).join("");
      }
    });
  } catch(e) { console.warn("refreshDashboard:", e.message); }
}

// ══════════════════════════════════════════════════════════════
// CREATE LOCATION SHEETS
// ══════════════════════════════════════════════════════════════
async function createLocationSheets() {
  if (!cfg.jiraUrl || !cfg.token) { toast("Configure Jira settings first", "warning"); return; }
  toast("Đang tạo sheets...", "warning");
  try {
    // Lấy 1 page đầu của mỗi typeId để discover locations
    const typeIds   = parseTypeIds();
    const locations = new Set();
    for (const typeId of typeIds) {
      const data = await assetsPost("/object/aql", {
        qlQuery: `objectTypeId = ${typeId}`, startAt: 0, maxResults: 100, includeAttributes: true,
      });
      (data.values || []).forEach(obj => {
        const a = (obj.attributes||[]).find(x => String(x.objectTypeAttributeId) === "30125");
        const loc = a?.objectAttributeValues?.[0]?.displayValue || "";
        if (loc) locations.add(loc.trim());
      });
    }
    if (!locations.size) { toast("Không tìm thấy location nào", "warning"); return; }
    await Excel.run(async (context) => {
      for (const loc of locations) {
        const sheet = await ensureSheet(context, locationSheetName(loc));
        await ensureHeaders(context, sheet);
      }
    });
    toast(`Đã tạo ${locations.size} sheet(s)`, "success");
    await refreshDashboard();
  } catch(e) { toast("Error: " + e.message, "error"); }
}

// ══════════════════════════════════════════════════════════════
// MATCH LOCAL ASSETS
// ══════════════════════════════════════════════════════════════
async function matchLocalAssets() {
  toast("Tính năng này chạy tự động trong Sync Now", "warning");
}

// ══════════════════════════════════════════════════════════════
// VALIDATION
// ══════════════════════════════════════════════════════════════
async function runValidation() {
  toast("Running validation…", "warning");
  const errors = {
    "Missing Asset ID": [], "Duplicate Serial": [],
    "Location Changed": [], "Not in Jira": [],
  };
  try {
    await Excel.run(async (context) => {
      const sheets     = await getLocationSheets(context);
      const serialSeen = {};
      for (const sheetName of sheets) {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const rows  = await readSheetRows(context, sheet);
        for (let i = 0; i < rows.length; i++) {
          const row      = rows[i];
          const serial   = String(row[COL.SERIAL]   || "").trim();
          const assetId  = String(row[COL.ASSET_ID] || "").trim();
          const locField = String(row[COL.LOCATION] || "").trim();
          let valid = String(row[COL.VALIDATION] || "").trim();
          if (valid === "Not in Jira") {
            errors["Not in Jira"].push({ sheet: sheetName, row: i + 2 }); continue;
          }
          valid = "OK";
          if (row[COL.SYNC_STATUS] === "JIRA" && !assetId) {
            valid = "Missing Asset ID";
            errors["Missing Asset ID"].push({ sheet: sheetName, row: i + 2 });
          }
          if (serial && valid === "OK") {
            if (serialSeen[serial]) {
              valid = "Duplicate Serial";
              errors["Duplicate Serial"].push({ sheet: sheetName, row: i + 2 });
              errors["Duplicate Serial"].push(serialSeen[serial]);
            } else { serialSeen[serial] = { sheet: sheetName, row: i + 2 }; }
          }
          if (locField && valid === "OK") {
            if (locationSheetName(locField) !== sheetName) {
              valid = "Location Changed";
              errors["Location Changed"].push({ sheet: sheetName, row: i + 2 });
            }
          }
          const cell = sheet.getRangeByIndexes(i + 1, COL.VALIDATION, 1, 1);
          cell.values = [[valid]];
          cell.format.font.color = valid === "OK" ? "#22c55e" : "#ef4444";
        }
        await context.sync();
      }
    });
    renderValidationList(errors);
    toast("Validation complete", "success");
    await refreshDashboard();
  } catch(e) { toast("Validation error: " + e.message, "error"); }
}

function renderValidationList(errors) {
  const el = document.getElementById("val-list");
  el.innerHTML = Object.entries(errors).map(([name, list]) => {
    const count  = list.length;
    const unique = list.filter((v,i,a) => i === a.findIndex(x => x.sheet===v.sheet && x.row===v.row));
    return `<div class="val-item" data-errors='${JSON.stringify(unique)}' onclick="jumpToError(this)">
      <span class="val-dot ${count > 0 ? "err" : "ok"}"></span>
      <span class="val-name">${name}</span>
      <span class="val-count ${count > 0 ? "has-err" : ""}">${count}</span>
    </div>`;
  }).join("") || `<div class="empty-state"><div class="icon">✓</div>No issues found</div>`;
}

async function jumpToError(el) {
  const errs = JSON.parse(el.dataset.errors || "[]");
  if (!errs.length) return;
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(errs[0].sheet);
      sheet.activate();
      sheet.getRangeByIndexes(errs[0].row - 1, 0, 1, 1).select();
      await context.sync();
    });
  } catch(e) { console.warn(e); }
}
window.jumpToError = jumpToError;

// ══════════════════════════════════════════════════════════════
// TICKET
// ══════════════════════════════════════════════════════════════
async function scanPendingRows() {
  let count = 0;
  try {
    await Excel.run(async (context) => {
      const sheets = await getLocationSheets(context);
      for (const sheetName of sheets) {
        const rows = await readSheetRows(context, context.workbook.worksheets.getItem(sheetName));
        rows.forEach(r => { if (r[COL.ACTION] === "Create Ticket" && !r[COL.CASE_JIRA]) count++; });
      }
    });
  } catch(e) { console.warn(e); }
  setInner("selected-count", count);
}

async function createTickets() {
  const issueType = document.getElementById("issue-type").value;
  const priority  = document.getElementById("issue-priority").value;
  const days      = parseInt(document.getElementById("default-days").value) || 30;
  if (!cfg.jiraUrl || !cfg.token) { toast("Configure Jira settings first", "warning"); return; }
  toast("Processing ticket queue…", "warning");
  let created = 0, skipped = 0, failed = 0;
  try {
    await Excel.run(async (context) => {
      const sheets = await getLocationSheets(context);
      for (const sheetName of sheets) {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const rows  = await readSheetRows(context, sheet);
        for (let i = 0; i < rows.length; i++) {
          const row = rows[i];
          if (row[COL.ACTION] !== "Create Ticket") continue;
          if (row[COL.CASE_JIRA]) { skipped++; continue; }
          const deviceName = row[COL.HOSTNAME] || row[COL.ASSET_KEY] || "Unknown";
          const email      = row[COL.ASSIGNED] || row[COL.USERNAME]  || "";
          const serial     = row[COL.SERIAL]   || "";
          const note       = row[COL.NOTE]     || "";
          const rowDays    = row[COL.DAYS]     || days;
          try {
            const res = await jiraPost("/api/3/issue", {
              fields: {
                project:     { key: cfg.projectKey },
                summary:     `[${issueType}] ${deviceName}${serial ? " – "+serial : ""}`,
                description: { type:"doc", version:1, content:[{ type:"paragraph",
                  content:[{ type:"text", text:`Asset: ${deviceName}\nSerial: ${serial}\nUser: ${email}\nDays: ${rowDays}\nNote: ${note}` }]
                }]},
                issuetype: { name: issueType },
                priority:  { name: priority },
              }
            });
            if (res.key) {
              sheet.getRangeByIndexes(i+1, COL.CASE_JIRA, 1, 1).values =
                [[`=HYPERLINK("${jiraBase()}/browse/${res.key}","${res.key}")`]];
              sheet.getRangeByIndexes(i+1, COL.ACTION, 1, 1).values = [[""]];
              await context.sync();
              created++;
            }
          } catch(e) { failed++; toast(`Row ${i+2}: ${e.message}`, "error"); }
        }
      }
    });
    toast(`Done — created:${created}, skipped:${skipped}${failed?", failed:"+failed:""}`,
      failed ? "warning" : "success");
    scanPendingRows();
  } catch(e) { toast("Ticket error: " + e.message, "error"); }
}

// ══════════════════════════════════════════════════════════════
// CONNECTION TEST
// ══════════════════════════════════════════════════════════════
async function testConnection() {
  const el = document.getElementById("conn-test-result");
  el.style.display = "block";
  el.style.borderColor = "var(--border)";
  el.innerHTML = `<div style="display:flex;gap:8px;align-items:center"><div class="spinner"></div><span>Testing…</span></div>`;
  saveConfig();
  if (!cfg.workerUrl) {
    el.innerHTML = "✗ Worker URL chưa điền";
    el.style.borderColor = "var(--red)"; return;
  }
  try {
    const me = await jiraGet("/api/3/myself");
    let assetsOk = "";
    try {
      await assetsPost("/object/aql", { qlQuery: "objectType != null", startAt: 0, maxResults: 1, includeAttributes: false });
      assetsOk = " · Assets API ✓";
    } catch(ae) { assetsOk = ` · Assets ✗ (${ae.message.slice(0,50)})`; }
    el.innerHTML = `✓ <strong>${me.displayName || me.emailAddress}</strong>${assetsOk}`;
    el.style.borderColor = "var(--green)";
    toast("Connection OK", "success");
  } catch(e) {
    el.innerHTML = `✗ ${e.message}`;
    el.style.borderColor = "var(--red)";
    toast("Connection failed", "error");
  }
}

// ══════════════════════════════════════════════════════════════
// MISC
// ══════════════════════════════════════════════════════════════
function openJira() {
  if (!cfg.jiraUrl) { toast("Set Jira URL in Settings first", "warning"); return; }
  window.open(cfg.jiraUrl, "_blank");
}

function updateSyncPanel(el, state) {
  el.innerHTML = state.map(l => `
    <div class="location-item">
      <div class="loc-header">
        <span class="loc-name">${l.name}</span>
        <span class="loc-status ${l.status}">${l.status === "done" ? "Done" : l.status === "running" ? "Running…" : "Waiting"}</span>
      </div>
      <div class="progress-track"><div class="progress-fill" style="width:${l.total?Math.round(l.done/l.total*100):0}%"></div></div>
      <div class="loc-count">${l.done} / ${l.total}</div>
    </div>`).join("");
}

function setSyncIndicator(state, text) {
  const dot = document.getElementById("sync-dot");
  const txt = document.getElementById("sync-status-text");
  if (dot) dot.className = "sync-dot" + (state === "syncing" ? " syncing" : "");
  if (txt) txt.textContent = text;
}

function toast(msg, type = "success") {
  const container = document.getElementById("toast-container");
  if (!container) return;
  const el = document.createElement("div");
  el.className = `toast ${type}`;
  el.innerHTML = `<span>${{success:"✓",error:"✗",warning:"⚠"}[type]||"ℹ"}</span><span>${msg}</span>`;
  container.appendChild(el);
  setTimeout(() => el.remove(), 4500);
}

function setInner(id, val) { const el=document.getElementById(id); if(el) el.textContent=val; }
function setVal(id, val)   { const el=document.getElementById(id); if(el) el.value=val||""; }
function getVal(id)        { const el=document.getElementById(id); return el?el.value.trim():""; }
function formatTime(iso) {
  if (!iso) return "—";
  try { return new Date(iso).toLocaleString([],{month:"2-digit",day:"2-digit",hour:"2-digit",minute:"2-digit"}); }
  catch { return iso; }
}