/* ════════════════════════════════════════════════════════════
   Jira Asset Manager — Office Add-in
   taskpane.js — Full logic layer  (v1.1 - fixed)
   ════════════════════════════════════════════════════════════ */

"use strict";

// ── COLUMN MAP ────────────────────────────────────────────────
const COL = {
  ASSET_ID:    0,   // Jira Object ID
  ASSET_KEY:   1,   // Jira Object Key
  HOSTNAME:    2,   // attr 1737
  SERIAL:      3,   // attr 5194
  STATUS:      4,   // attr 5052
  LOCATION:    5,   // attr 30125
  REGION:      6,   // attr 27292
  MANUFACTURER:7,   // attr 6608
  MODEL:       8,   // attr 6609
  OS:          9,   // attr 30345
  OS_VERSION:  10,  // attr 27291
  OS_BUILD:    11,  // attr 27290
  CPU:         12,  // attr 6610
  IP:          13,  // attr 5208
  MAC:         14,  // attr 5209
  NETWORK:     15,  // attr 5210
  ANTIVIRUS:   16,  // attr 6612
  USERNAME:    17,  // attr 5200
  ASSIGNED:    18,  // attr 26690
  FIRST_SEEN:  19,  // attr 5205
  LAST_SEEN:   20,  // attr 5206
  PURCHASE:    21,  // attr 5203
  WARRANTY:    22,  // attr 6615
  TENANT_ID:   23,  // attr 26398
  LANSWEEPER:  24,  // attr 5207
  // ── User-entered cols ──
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

// ── CONFIG KEYS ───────────────────────────────────────────────
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

// ── STATE ─────────────────────────────────────────────────────
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
  } catch {
    el.textContent = "—";
  }
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
  s.saveAsync(() => {
    updateWorkspaceLabel();
    toast("Settings saved", "success");
  });
}

function populateConfigUI() {
  setVal("cfg-jira-url",      cfg.jiraUrl);
  setVal("cfg-email",         cfg.email);
  setVal("cfg-token",         cfg.token);
  setVal("cfg-cloud-id",      cfg.cloudId);
  setVal("cfg-workspace-id",  cfg.workspaceId);
  setVal("cfg-project-key",   cfg.projectKey);
  setVal("cfg-worker-url",    cfg.workerUrl);
  setVal("cfg-aql-query",     cfg.aqlQuery);

}

// ══════════════════════════════════════════════════════════════
// EVENTS
// ══════════════════════════════════════════════════════════════
function wireEvents() {
  // Tabs
  document.querySelectorAll(".tab").forEach(t =>
    t.addEventListener("click", () => switchTab(t.dataset.tab))
  );
  // Dashboard
  on("btn-sync-now",       () => runSync());
  on("btn-validate-all",   () => runValidation());
  on("btn-open-jira",      openJira);
  on("btn-create-sheets",  createLocationSheets);
  // Validation
  on("btn-run-validate",   () => runValidation());
  // Ticket
  on("btn-scan-pending",   scanPendingRows);
  on("btn-create-tickets", createTickets);
  // Sync tab
  on("btn-full-sync",      () => runSync());
  on("btn-sync-local",     matchLocalAssets);
  // Settings
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
// JIRA API — gọi qua Cloudflare Worker proxy (giải quyết CORS)
//
// Mọi request đều đi qua:
//   https://YOUR_WORKER.workers.dev/proxy?url=TARGET_URL
// Worker forward lên Jira và trả về với CORS header đúng.
// ══════════════════════════════════════════════════════════════

function jiraBase() {
  return cfg.jiraUrl.replace(/\/+$/, "");
}

function assetsBase() {
  return `https://api.atlassian.com/ex/jira/${cfg.cloudId}/jsm/assets/workspace/${cfg.workspaceId}/v1`;
}

function proxyUrl(targetUrl) {
  const worker = cfg.workerUrl.replace(/\/+$/, "");
  return `${worker}/proxy?url=${encodeURIComponent(targetUrl)}`;
}

function jiraHeaders() {
  return {
    "Authorization": `Basic ${btoa(cfg.email + ":" + cfg.token)}`,
    "Content-Type":  "application/json",
    "Accept":        "application/json",
  };
}

// Gọi Jira REST API qua proxy
async function jiraGet(path) {
  const target = `${jiraBase()}/rest${path}`;
  const res = await fetch(proxyUrl(target), { headers: jiraHeaders() });
  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`Jira ${res.status}${txt ? ": " + txt.slice(0, 120) : ""}`);
  }
  return res.json();
}

async function jiraPost(path, body) {
  const target = `${jiraBase()}/rest${path}`;
  const res = await fetch(proxyUrl(target), {
    method:  "POST",
    headers: jiraHeaders(),
    body:    JSON.stringify(body),
  });
  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`Jira ${res.status}${txt ? ": " + txt.slice(0, 120) : ""}`);
  }
  return res.json();
}

// Gọi Jira Assets API qua proxy
async function assetsPost(path, body) {
  const target = `${assetsBase()}${path}`;
  const res = await fetch(proxyUrl(target), {
    method:  "POST",
    headers: jiraHeaders(),
    body:    JSON.stringify(body),
  });
  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`Assets API ${res.status}${txt ? ": " + txt.slice(0, 120) : ""}`);
  }
  return res.json();
}

// ── Fetch 1 page từ Jira Assets API ──────────────────────────
async function fetchPage(qlQuery, startAt, pageSize) {
  return assetsPost("/object/aql", {
    qlQuery,
    startAt,
    maxResults:        pageSize,
    includeAttributes: true,
  });
}

// ── Parse 1 object thành asset record ─────────────────────────
function parseAsset(obj) {
  const attrById = (id) => {
    const a = (obj.attributes || []).find(
      x => String(x.objectTypeAttributeId) === String(id)
    );
    return a?.objectAttributeValues?.[0]?.displayValue || "";
  };
  return {
    id:          String(obj.id || ""),
    key:         obj.objectKey || "",
    hostname:    attrById(1737) || obj.label || "",
    serial:      attrById(5194),
    status:      attrById(5052),
    location:    attrById(30125),
    region:      attrById(27292),
    manufacturer:attrById(6608),
    model:       attrById(6609),
    os:          attrById(30345),
    osVersion:   attrById(27291),
    osBuild:     attrById(27290),
    cpu:         attrById(6610),
    ip:          attrById(5208),
    mac:         attrById(5209),
    network:     attrById(5210),
    antivirus:   attrById(6612),
    username:    attrById(5200),
    assigned:    attrById(26690),
    firstSeen:   attrById(5205),
    lastSeen:    attrById(5206),
    purchase:    attrById(5203),
    warranty:    attrById(6615),
    tenantId:    attrById(26398),
    lansweeper:  attrById(5207),
  };
}

// ── Fetch tất cả assets của 1 AQL query ───────────────────────
// API giới hạn trả tối đa 1000 record/query (total <= 1000).
// Chiến lược:
//   - Mỗi "window": fetch từ startAt, dùng total của window làm limit
//   - Nếu total < 1000 → đây là window cuối, dừng sau khi load hết
//   - Nếu total = 1000 → còn data, dịch startAt += 1000 và fetch window tiếp
async function fetchByQuery(qlQuery) {
  const assets    = [];
  const pageSize  = 25;
  const API_LIMIT = 1000; // hard limit của Jira Assets API
  let windowStart = 0;    // startAt của window hiện tại

  while (true) {
    // Lấy total của window này từ page đầu tiên
    const firstPage   = await fetchPage(qlQuery, windowStart, pageSize);
    const windowTotal = typeof firstPage.total === "number" ? firstPage.total : 0;
    console.log(`  [${qlQuery}] window startAt=${windowStart}, total=${windowTotal}`);

    if (windowTotal === 0 || firstPage.values?.length === 0) break;

    // Load hết window này (windowStart → windowStart + windowTotal)
    firstPage.values.forEach(obj => assets.push(parseAsset(obj)));
    let pageStart = windowStart + firstPage.values.length;

    while (assets.length < windowStart + windowTotal) {
      const data   = await fetchPage(qlQuery, pageStart, pageSize);
      const values = data.values || [];
      if (values.length === 0) break;
      values.forEach(obj => assets.push(parseAsset(obj)));
      pageStart += values.length;
      console.log(`  [${qlQuery}] loaded=${assets.length}/${windowStart + windowTotal}`);
      if (values.length < pageSize) break;
    }

    console.log(`  [${qlQuery}] window done: ${assets.length} total so far`);

    // Nếu window này < 1000 → đã lấy hết, dừng
    if (windowTotal < API_LIMIT) break;

    // Nếu window = 1000 → còn data, dịch sang window tiếp
    windowStart += API_LIMIT;
  }

  console.log(`  [${qlQuery}] DONE: ${assets.length} assets`);
  return assets;
}

// ── Parse objectTypeIds từ AQL config ────────────────────────
function parseTypeIds() {
  const aqlRaw = (cfg.aqlQuery || "").trim();
  const m = aqlRaw.match(/objectTypeId\s+IN\s*\(([^)]+)\)/i);
  if (m) return m[1].split(",").map(s => s.trim()).filter(Boolean);
  const m2 = aqlRaw.match(/objectTypeId\s*=\s*(\d+)/i);
  if (m2) return [m2[1]];
  return [];
}

// ── Fetch ALL assets, từng typeId riêng để bypass 1000 limit ──
async function fetchJiraAssets() {
  const typeIds = parseTypeIds();
  if (!typeIds.length) {
    toast("AQL Query chưa được cấu hình đúng", "error");
    return [];
  }

  const seenIds = new Set();
  const allAssets = [];

  for (let i = 0; i < typeIds.length; i++) {
    const typeId = typeIds[i];
    toast(`Đang tải objectTypeId=${typeId} (${i+1}/${typeIds.length})...`, "warning");
    const assets = await fetchByQuery(`objectTypeId = ${typeId}`);
    let added = 0;
    assets.forEach(a => {
      if (!seenIds.has(a.id)) {
        seenIds.add(a.id);
        allAssets.push(a);
        added++;
      }
    });
    console.log(`typeId=${typeId}: fetched=${assets.length}, added=${added}, total=${allAssets.length}`);
  }

  console.log(`fetchJiraAssets complete: ${allAssets.length} assets from ${typeIds.length} typeId(s)`);
  return allAssets;
}

// (legacy parse block — replaced by parseAsset above, kept for ref)
function _legacyParseBlock(obj) {
      // Lookup by attribute ID (ổn định hơn tên)
}

// ══════════════════════════════════════════════════════════════
// SHEET HELPERS
// ══════════════════════════════════════════════════════════════
function locationSheetName(loc) {
  return "LOC_" + loc.replace(/[^a-zA-Z0-9]/g, "_").toUpperCase();
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

// FIX: only write headers if row 1 is empty (don't overwrite data)
async function ensureHeaders(context, sheet) {
  const firstRow = sheet.getRangeByIndexes(0, 0, 1, COL_COUNT);
  firstRow.load("values");
  await context.sync();
  const existing = firstRow.values[0];
  const hasHeaders = existing && existing[0] === "Asset ID";
  if (!hasHeaders) {
    firstRow.values = [HEADERS];
    firstRow.format.font.bold  = true;
    firstRow.format.fill.color = "#1a1d27";
    firstRow.format.font.color = "#8892a4";
    sheet.freezePanes.freezeRows(1);
    await context.sync();
  }
}

// FIX: robust read — handles empty sheet and header-only sheet
async function readSheetRows(context, sheet) {
  const used = sheet.getUsedRangeOrNullObject(true);
  await context.sync();
  if (used.isNullObject) return [];

  used.load(["values", "rowCount"]);
  await context.sync();

  if (used.rowCount <= 1) return [];          // header only
  return used.values.slice(1);               // skip header row
}

async function appendRow(context, sheet, rowData) {
  const used = sheet.getUsedRangeOrNullObject(true);
  await context.sync();
  let nextRow = 1; // default: row after header
  if (!used.isNullObject) {
    used.load("rowCount");
    await context.sync();
    nextRow = used.rowCount;
  }
  const range = sheet.getRangeByIndexes(nextRow, 0, 1, COL_COUNT);
  range.values = [rowData];
  await context.sync();
}

async function getLocationSheets(context) {
  context.workbook.worksheets.load("items/name");
  await context.sync();
  return context.workbook.worksheets.items
    .filter(s => s.name.startsWith("LOC_"))
    .map(s => s.name);
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
          if (r[COL.SYNC_STATUS] === "LOCAL")                                  { local++;  locLocal++; }
          if (r[COL.VALIDATION]  && r[COL.VALIDATION] !== "OK")                  mismatch++;
          if (r[COL.ACTION] === "Create Ticket" && !r[COL.CASE_JIRA])            pending++;
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
          </div>
        `).join("");
      }
    });
  } catch(e) {
    console.warn("refreshDashboard:", e.message);
  }
}

// ══════════════════════════════════════════════════════════════
// CREATE LOCATION SHEETS
// FIX: don't call non-existent statustype endpoint
// ══════════════════════════════════════════════════════════════
async function createLocationSheets() {
  if (!cfg.jiraUrl || !cfg.token) {
    toast("Configure Jira settings first", "warning"); return;
  }
  toast("Fetching assets from Jira to discover locations...", "warning");
  try {
    const assets    = await fetchJiraAssets();
    const locations = [...new Set(assets.map(a => a.location).filter(Boolean))];

    if (!locations.length) {
      toast("No locations found in Jira assets", "warning"); return;
    }

    await Excel.run(async (context) => {
      for (const loc of locations) {
        const sheetName = locationSheetName(loc);
        const sheet     = await ensureSheet(context, sheetName);
        await ensureHeaders(context, sheet);
      }
    });
    toast(`Created ${locations.length} location sheet(s)`, "success");
    await refreshDashboard();
  } catch(e) {
    toast("Error: " + e.message, "error");
  }
}

// ══════════════════════════════════════════════════════════════
// SYNC ENGINE
// Luồng:
//   1. Fetch từng objectTypeId riêng (bypass 1000 limit)
//   2. Group tất cả assets theo Location
//   3. Mỗi location: tạo sheet nếu chưa có
//      - Asset đã có trong sheet → UPDATE tại chỗ
//      - Asset mới                → INSERT cuối sheet
//      - Asset trong sheet ko còn trong Jira → "Not in Jira"
//   4. LOCAL rows → push lên Jira Assets API
// ══════════════════════════════════════════════════════════════
function assetToRow(asset, loc, now, existingRow) {
  // existingRow: dòng hiện có (để preserve DAYS, NOTE, ACTION, CASE_JIRA)
  const row = existingRow ? [...existingRow] : Array(COL_COUNT).fill("");
  row[COL.ASSET_ID]     = asset.id;
  row[COL.ASSET_KEY]    = asset.key;
  row[COL.HOSTNAME]     = asset.hostname;
  row[COL.SERIAL]       = asset.serial;
  row[COL.STATUS]       = asset.status;
  row[COL.LOCATION]     = loc;
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
  // COL.DAYS, COL.NOTE, COL.ACTION, COL.CASE_JIRA — preserved từ existingRow
  return row;
}

// ── Write 1 location sheet (dùng chung cho mọi typeId) ────────
async function writeLocationSheet(sheetName, assets, now) {
  await Excel.run(async (context) => {
    const sheet = await ensureSheet(context, sheetName);
    await ensureHeaders(context, sheet);

    const existing = await readSheetRows(context, sheet);
    const byId = {}, bySerial = {};
    existing.forEach((row, idx) => {
      const id  = String(row[COL.ASSET_ID] || "").trim();
      const ser = String(row[COL.SERIAL]   || "").trim();
      if (id)  byId[id]      = idx;
      if (ser) bySerial[ser] = idx;
    });

    const jiraIdSet   = new Set(assets.map(a => a.id).filter(Boolean));
    const updatedIdxs = new Set();

    // A: UPDATE existing rows
    assets.forEach(asset => {
      const idx =
        byId[asset.id] !== undefined                         ? byId[asset.id] :
        asset.serial && bySerial[asset.serial] !== undefined ? bySerial[asset.serial] :
        -1;
      if (idx >= 0) {
        const newRow = assetToRow(asset, asset.location || "", now, existing[idx]);
        if (newRow[COL.VALIDATION] === "Not in Jira") newRow[COL.VALIDATION] = "";
        sheet.getRangeByIndexes(idx + 1, 0, 1, COL_COUNT).values = [newRow];
        updatedIdxs.add(idx);
      }
    });

    // B: Mark rows không còn trong Jira
    existing.forEach((row, idx) => {
      if (updatedIdxs.has(idx))             return;
      if (row[COL.SYNC_STATUS] === "LOCAL") return;
      const id = String(row[COL.ASSET_ID] || "").trim();
      if (!id || jiraIdSet.has(id))         return;
      const cell = sheet.getRangeByIndexes(idx + 1, COL.VALIDATION, 1, 1);
      cell.values = [["Not in Jira"]];
      cell.format.font.color = "#f59e0b";
    });

    // C: INSERT rows mới
    const toInsert = assets.filter(a =>
      byId[a.id] === undefined &&
      !(a.serial && bySerial[a.serial] !== undefined)
    );
    if (toInsert.length > 0) {
      const used = sheet.getUsedRangeOrNullObject(true);
      await context.sync();
      let startRow = 1;
      if (!used.isNullObject) {
        used.load("rowCount");
        await context.sync();
        startRow = used.rowCount;
      }
      sheet.getRangeByIndexes(startRow, 0, toInsert.length, COL_COUNT).values =
        toInsert.map(a => assetToRow(a, a.location || "", now, null));
    }

    await context.sync();
    console.log(`${sheetName}: updated=${updatedIdxs.size}, inserted=${toInsert.length}`);
  });
}

async function runSync() {
  if (isSyncing) { toast("Sync already running", "warning"); return; }
  if (!cfg.jiraUrl || !cfg.token || !cfg.workerUrl) {
    toast("Kiểm tra lại Settings (URL, token, worker)", "warning"); return;
  }

  isSyncing = true;
  setSyncIndicator("syncing", "Syncing...");
  const syncPanel = document.getElementById("sync-location-list");
  syncPanel.innerHTML = `<div class="empty-state"><div class="spinner"></div>Đang tải từ Jira...</div>`;

  try {
    const typeIds = parseTypeIds();
    if (!typeIds.length) {
      toast("AQL Query chưa đúng — cần objectTypeId IN (...)", "error");
      return;
    }

    // byLocation tích lũy qua từng typeId
    // key = locationSheetName, value = Map<assetId, asset>
    // Dùng Map để tự dedup nếu asset xuất hiện ở nhiều typeId
    const locationMap = {}; // sheetName → Map<id, asset>
    const seenIds     = new Set();
    let totalFetched  = 0;
    const now         = new Date().toISOString();

    for (let i = 0; i < typeIds.length; i++) {
      const typeId = typeIds[i];
      toast(`[${i+1}/${typeIds.length}] Đang tải objectTypeId=${typeId}...`, "warning");

      const assets = await fetchByQuery(`objectTypeId = ${typeId}`);
      let added = 0;

      assets.forEach(a => {
        if (seenIds.has(a.id)) return;
        seenIds.add(a.id);
        added++;
        totalFetched++;

        const sheetName = locationSheetName((a.location || "UNKNOWN").trim());
        if (!locationMap[sheetName]) locationMap[sheetName] = new Map();
        locationMap[sheetName].set(a.id, a);
      });

      console.log(`typeId=${typeId}: fetched=${assets.length}, added=${added}, total=${totalFetched}`);

      // Ghi ngay vào Excel sau mỗi typeId — không chờ load hết
      const sheetNames = Object.keys(locationMap);
      const uiState = sheetNames.map(s => ({
        name: s.replace("LOC_", ""),
        done: locationMap[s].size,
        total: locationMap[s].size,
        status: i < typeIds.length - 1 ? "running" : "done",
      }));
      updateSyncPanel(syncPanel, uiState);

      for (const sheetName of sheetNames) {
        const sheetAssets = Array.from(locationMap[sheetName].values());
        await writeLocationSheet(sheetName, sheetAssets, now);
      }
    }

    // Push LOCAL rows lên Jira
    toast("Đang push LOCAL assets lên Jira...", "warning");
    await pushLocalAssets();

    cfg.lastSync = new Date().toISOString();
    Office.context.document.settings.set(CFG_KEYS.LAST_SYNC, cfg.lastSync);
    Office.context.document.settings.saveAsync();

    setSyncIndicator("ok", "Synced");
    setInner("last-sync-time", formatTime(cfg.lastSync));
    toast(`Sync hoàn tất — ${totalFetched} assets, ${Object.keys(locationMap).length} location(s)`, "success");
    await refreshDashboard();

  } catch(e) {
    setSyncIndicator("ok", "Sync failed");
    toast("Sync error: " + e.message, "error");
    console.error("runSync:", e);
  } finally {
    isSyncing = false;
  }
}

// ── Push LOCAL rows lên Jira Assets ───────────────────────────
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
          if (row[COL.ASSET_ID])               continue; // đã có ID = đã push

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

            const res = await assetsPost("/object/create", {
              objectTypeId: defaultTypeId,
              attributes:   attrs,
            });

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
          } catch(e) {
            failed++;
            console.warn(`pushLocal row ${i+2} in ${sheetName}:`, e.message);
          }
        }
      }
    });
    if (pushed > 0 || failed > 0) {
      toast(`LOCAL push: ${pushed} thành công${failed ? ", " + failed + " lỗi" : ""}`,
        failed ? "warning" : "success");
    }
  } catch(e) {
    console.warn("pushLocalAssets:", e.message);
  }
}


function updateSyncPanel(el, state) {
  el.innerHTML = state.map(l => `
    <div class="location-item">
      <div class="loc-header">
        <span class="loc-name">${locationSheetName(l.name)}</span>
        <span class="loc-status ${l.status}">
          ${l.status === "done" ? "Done" : l.status === "running" ? "Running…" : "Waiting"}
        </span>
      </div>
      <div class="progress-track">
        <div class="progress-fill" style="width:${l.total ? Math.round(l.done / l.total * 100) : 0}%"></div>
      </div>
      <div class="loc-count">${l.done} / ${l.total}</div>
    </div>
  `).join("");
}

// ══════════════════════════════════════════════════════════════
// MATCH LOCAL ASSETS
// ══════════════════════════════════════════════════════════════
async function matchLocalAssets() {
  if (!cfg.jiraUrl || !cfg.token) {
    toast("Configure Jira settings first", "warning"); return;
  }
  toast("Scanning LOCAL assets for Jira matches…", "warning");
  try {
    const jiraAssets  = await fetchJiraAssets();
    const jiraBySerial = {};
    jiraAssets.forEach(a => { if (a.serial) jiraBySerial[a.serial.trim()] = a; });

    let matched = 0;
    await Excel.run(async (context) => {
      const sheets = await getLocationSheets(context);
      for (const sheetName of sheets) {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const rows  = await readSheetRows(context, sheet);
        for (let i = 0; i < rows.length; i++) {
          if (rows[i][COL.SYNC_STATUS] !== "LOCAL") continue;
          const serial = String(rows[i][COL.SERIAL] || "").trim();
          if (!serial || !jiraBySerial[serial])       continue;

          const asset = jiraBySerial[serial];
          const range = sheet.getRangeByIndexes(i + 1, 0, 1, COL_COUNT);
          range.load("values");
          await context.sync();
          const cur = range.values[0];
          cur[COL.ASSET_ID]    = asset.id;
          cur[COL.ASSET_KEY]   = asset.key;
          cur[COL.HOSTNAME]    = asset.hostname;
          cur[COL.STATUS]      = asset.status;
          cur[COL.LOCATION]    = asset.location;
          cur[COL.MODEL]       = asset.model;
          cur[COL.SYNC_STATUS] = "JIRA";
          cur[COL.VALIDATION]  = "OK";
          cur[COL.LAST_SYNC]   = new Date().toISOString();
          range.values         = [cur];
          await context.sync();
          matched++;
        }
      }
    });

    toast(`Matched ${matched} LOCAL asset(s) to Jira`, "success");
    await refreshDashboard();
  } catch(e) {
    toast("Match error: " + e.message, "error");
  }
}

// ══════════════════════════════════════════════════════════════
// VALIDATION
// ══════════════════════════════════════════════════════════════
async function runValidation() {
  toast("Running validation…", "warning");

  const errors = {
    "Serial Mismatch":  [],
    "Owner Mismatch":   [],
    "Missing Asset ID": [],
    "Duplicate Serial": [],
    "Location Changed": [],
  };

  try {
    await Excel.run(async (context) => {
      const sheets     = await getLocationSheets(context);
      const serialSeen = {}; // serial → { sheet, row }

      for (const sheetName of sheets) {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const rows  = await readSheetRows(context, sheet);

        for (let i = 0; i < rows.length; i++) {
          const row     = rows[i];
          const serial  = String(row[COL.SERIAL]   || "").trim();
          const assetId = String(row[COL.ASSET_ID] || "").trim();
          const locField = String(row[COL.LOCATION] || "").trim();
          let valid = "OK";

          // Missing Asset ID for JIRA rows
          if (row[COL.SYNC_STATUS] === "JIRA" && !assetId) {
            valid = "Missing Asset ID";
            errors["Missing Asset ID"].push({ sheet: sheetName, row: i + 2 });
          }

          // Duplicate serial
          if (serial && valid === "OK") {
            if (serialSeen[serial]) {
              valid = "Duplicate Serial";
              errors["Duplicate Serial"].push({ sheet: sheetName, row: i + 2 });
              // Also mark the first occurrence
              errors["Duplicate Serial"].push(serialSeen[serial]);
            } else {
              serialSeen[serial] = { sheet: sheetName, row: i + 2 };
            }
          }

          // Location field doesn't match sheet name
          if (locField && valid === "OK") {
            const expected = locationSheetName(locField);
            if (expected !== sheetName) {
              valid = "Location Changed";
              errors["Location Changed"].push({ sheet: sheetName, row: i + 2 });
            }
          }

          // Write validation cell
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
  } catch(e) {
    toast("Validation error: " + e.message, "error");
  }
}

function renderValidationList(errors) {
  const el    = document.getElementById("val-list");
  const items = Object.entries(errors).map(([name, list]) => {
    const count = list.length;
    // De-duplicate list before storing
    const unique = list.filter((v, i, a) =>
      i === a.findIndex(x => x.sheet === v.sheet && x.row === v.row)
    );
    return `
      <div class="val-item" data-errors='${JSON.stringify(unique)}' onclick="jumpToError(this)">
        <span class="val-dot ${count > 0 ? "err" : "ok"}"></span>
        <span class="val-name">${name}</span>
        <span class="val-count ${count > 0 ? "has-err" : ""}">${count}</span>
      </div>
    `;
  });
  el.innerHTML = items.join("") ||
    `<div class="empty-state"><div class="icon">✓</div>No issues found</div>`;
}

async function jumpToError(el) {
  const errs = JSON.parse(el.dataset.errors || "[]");
  if (!errs.length) return;
  const first = errs[0];
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(first.sheet);
      sheet.activate();
      sheet.getRangeByIndexes(first.row - 1, 0, 1, 1).select();
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
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const rows  = await readSheetRows(context, sheet);
        rows.forEach(r => {
          if (r[COL.ACTION] === "Create Ticket" && !r[COL.CASE_JIRA]) count++;
        });
      }
    });
  } catch(e) { console.warn("scanPending:", e.message); }
  setInner("selected-count", count);
}

async function createTickets() {
  const issueType = document.getElementById("issue-type").value;
  const priority  = document.getElementById("issue-priority").value;
  const days      = parseInt(document.getElementById("default-days").value) || 30;

  if (!cfg.jiraUrl || !cfg.token) {
    toast("Configure Jira settings first", "warning"); return;
  }

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
          if (row[COL.CASE_JIRA]) { skipped++; continue; } // already has ticket

          const deviceName = row[COL.HOSTNAME] || row[COL.ASSET_KEY] || "Unknown";
          const email      = row[COL.ASSIGNED] || row[COL.USERNAME] || "";
          const serial     = row[COL.SERIAL]      || "";
          const note       = row[COL.NOTE]        || "";
          const rowDays    = row[COL.DAYS]        || days;

          const ticketBody = {
            fields: {
              project:     { key: cfg.projectKey },
              summary:     `[${issueType}] ${deviceName}${serial ? " – " + serial : ""}`,
              description: {
                type: "doc", version: 1,
                content: [{
                  type: "paragraph",
                  content: [{ type: "text", text:
                    `Asset: ${deviceName}\nSerial: ${serial}\nUser: ${email}\nDays: ${rowDays}\nNote: ${note}`
                  }]
                }]
              },
              issuetype: { name: issueType },
              priority:  { name: priority },
            }
          };

          try {
            let ticketKey = "";

            // Direct Jira API
            const res = await jiraPost("/api/3/issue", ticketBody);
            ticketKey = res.key || "";

            if (ticketKey) {
              const ticketUrl  = `${jiraBase()}/browse/${ticketKey}`;
              const caseCell   = sheet.getRangeByIndexes(i + 1, COL.CASE_JIRA, 1, 1);
              const actionCell = sheet.getRangeByIndexes(i + 1, COL.ACTION,    1, 1);
              caseCell.values   = [[`=HYPERLINK("${ticketUrl}","${ticketKey}")`]];
              actionCell.values = [[""]];
              await context.sync();
              created++;
            }
          } catch(ticketErr) {
            failed++;
            toast(`Row ${i + 2} failed: ${ticketErr.message}`, "error");
          }
        }
      }
    });

    toast(
      `Done — created: ${created}, skipped (duplicate): ${skipped}${failed ? ", failed: " + failed : ""}`,
      failed ? "warning" : "success"
    );
    scanPendingRows();
  } catch(e) {
    toast("Ticket error: " + e.message, "error");
  }
}

// ══════════════════════════════════════════════════════════════
// CONNECTION TEST  (FIX: correct endpoint path)
// ══════════════════════════════════════════════════════════════
async function testConnection() {
  const el = document.getElementById("conn-test-result");
  el.style.display = "block";
  el.style.borderColor = "var(--border)";
  el.innerHTML = `<div style="display:flex;gap:8px;align-items:center">
    <div class="spinner"></div><span>Testing connection…</span>
  </div>`;

  saveConfig(); // persist + reload cfg first

  if (!cfg.workerUrl) {
    el.innerHTML = "✗ Worker URL chưa được điền";
    el.style.display = "block";
    el.style.borderColor = "var(--red)";
    toast("Điền Cloudflare Worker URL trước", "error");
    return;
  }

  try {
    // 1. Test Jira REST API
    const me = await jiraGet("/api/3/myself");
    // 2. Test Assets API — fetch 1 object to confirm workspace access
    let assetsOk = "";
    try {
      await assetsPost(`/object/aql?startAt=0&maxResults=1`, { qlQuery: "objectType != null" });
      assetsOk = " · Assets API ✓";
    } catch(ae) {
      assetsOk = ` · Assets API ✗ (${ae.message.slice(0,60)})`;
    }
    el.innerHTML     = `✓ Connected as <strong>${me.displayName || me.emailAddress || "Unknown"}</strong>${assetsOk}`;
    el.style.borderColor = "var(--green)";
    toast("Connection successful", "success");
  } catch(e) {
    el.innerHTML     = `✗ ${e.message}`;
    el.style.borderColor = "var(--red)";
    toast("Connection failed — check URL, email, token", "error");
  }
}

// ══════════════════════════════════════════════════════════════
// OPEN JIRA
// ══════════════════════════════════════════════════════════════
function openJira() {
  if (!cfg.jiraUrl) { toast("Set Jira URL in Settings first", "warning"); return; }
  window.open(cfg.jiraUrl, "_blank");
}

// ══════════════════════════════════════════════════════════════
// SYNC INDICATOR
// ══════════════════════════════════════════════════════════════
function setSyncIndicator(state, text) {
  const dot = document.getElementById("sync-dot");
  const txt = document.getElementById("sync-status-text");
  if (dot) dot.className = "sync-dot" + (state === "syncing" ? " syncing" : "");
  if (txt) txt.textContent = text;
}

// ══════════════════════════════════════════════════════════════
// TOAST
// ══════════════════════════════════════════════════════════════
function toast(msg, type = "success") {
  const icons     = { success: "✓", error: "✗", warning: "⚠" };
  const container = document.getElementById("toast-container");
  if (!container) return;
  const el = document.createElement("div");
  el.className = `toast ${type}`;
  el.innerHTML = `<span>${icons[type] || "ℹ"}</span><span>${msg}</span>`;
  container.appendChild(el);
  setTimeout(() => el.remove(), 4500);
}

// ══════════════════════════════════════════════════════════════
// UTILS
// ══════════════════════════════════════════════════════════════
function setInner(id, val) {
  const el = document.getElementById(id);
  if (el) el.textContent = val;
}
function setVal(id, val) {
  const el = document.getElementById(id);
  if (el) el.value = val || "";
}
function getVal(id) {
  const el = document.getElementById(id);
  return el ? el.value.trim() : "";
}
function formatTime(iso) {
  if (!iso) return "—";
  try {
    return new Date(iso).toLocaleString([], {
      month: "2-digit", day: "2-digit",
      hour: "2-digit", minute: "2-digit"
    });
  } catch { return iso; }
}