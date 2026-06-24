/* ════════════════════════════════════════════════════════════
   Jira Asset Manager — Office Add-in
   taskpane.js — Full logic layer  (v1.1 - fixed)
   ════════════════════════════════════════════════════════════ */

"use strict";

// ── COLUMN MAP ────────────────────────────────────────────────
const COL = {
  ASSET_ID:    0,
  ASSET_KEY:   1,
  DEVICE_NAME: 2,
  SERIAL:      3,
  USER_EMAIL:  4,
  MODEL:       5,
  LOCATION:    6,
  STATUS:      7,
  DAYS:        8,
  NOTE:        9,
  CASE_JIRA:   10,
  VALIDATION:  11,
  SYNC_STATUS: 12,
  LAST_SYNC:   13,
  ACTION:      14,
};
const COL_COUNT = 15;

const HEADERS = [
  "Asset ID","Asset Key","Device Name","Serial Number",
  "User Email","Model","Location","Status",
  "Days","Note","Case Jira","Validation",
  "Sync Status","Last Sync","Action"
];

// ── CONFIG KEYS ───────────────────────────────────────────────
const CFG_KEYS = {
  JIRA_URL:     "jiraUrl",
  EMAIL:        "jiraEmail",
  TOKEN:        "jiraToken",
  CLOUD_ID:     "cloudId",
  WORKSPACE_ID: "workspaceId",
  PROJECT_KEY:  "projectKey",
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
  s.set(CFG_KEYS.JIRA_URL,     cfg.jiraUrl);
  s.set(CFG_KEYS.EMAIL,        cfg.email);
  s.set(CFG_KEYS.TOKEN,        cfg.token);
  s.set(CFG_KEYS.CLOUD_ID,     cfg.cloudId);
  s.set(CFG_KEYS.WORKSPACE_ID, cfg.workspaceId);
  s.set(CFG_KEYS.PROJECT_KEY,  cfg.projectKey);
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
// JIRA API  — Atlassian Cloud (api.atlassian.com)
// ══════════════════════════════════════════════════════════════

// https://yourorg.atlassian.net  (for /rest/api/3 calls)
function jiraBase() {
  return cfg.jiraUrl.replace(/\/+$/, "");
}

// https://api.atlassian.com/ex/jira/{cloudId}/jsm/assets/workspace/{workspaceId}/v1
function assetsBase() {
  return `https://api.atlassian.com/ex/jira/${cfg.cloudId}/jsm/assets/workspace/${cfg.workspaceId}/v1`;
}

function jiraHeaders() {
  return {
    "Authorization": `Basic ${btoa(cfg.email + ":" + cfg.token)}`,
    "Content-Type":  "application/json",
    "Accept":        "application/json",
  };
}

// Jira REST API (e.g. /api/3/myself, /api/3/issue)
async function jiraGet(path) {
  const url = `${jiraBase()}/rest${path}`;
  const res = await fetch(url, { headers: jiraHeaders() });
  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`Jira ${res.status}${txt ? ": " + txt.slice(0, 120) : ""}`);
  }
  return res.json();
}

async function jiraPost(path, body) {
  const url = `${jiraBase()}/rest${path}`;
  const res = await fetch(url, {
    method: "POST",
    headers: jiraHeaders(),
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`Jira ${res.status}${txt ? ": " + txt.slice(0, 120) : ""}`);
  }
  return res.json();
}

// Jira Assets API (api.atlassian.com)
async function assetsGet(path) {
  const url = `${assetsBase()}${path}`;
  const res = await fetch(url, { headers: jiraHeaders() });
  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`Assets API ${res.status}${txt ? ": " + txt.slice(0, 120) : ""}`);
  }
  return res.json();
}

async function assetsPost(path, body) {
  const url = `${assetsBase()}${path}`;
  const res = await fetch(url, {
    method: "POST",
    headers: jiraHeaders(),
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`Assets API ${res.status}${txt ? ": " + txt.slice(0, 120) : ""}`);
  }
  return res.json();
}

// ── Fetch all assets via AQL pagination ──────────────────────
// Endpoint: POST /object/aql?startAt=0&maxResults=50
async function fetchJiraAssets() {
  const assets = [];
  let startAt = 0;
  const pageSize = 50;

  while (true) {
    const data = await assetsPost(
      `/object/aql?startAt=${startAt}&maxResults=${pageSize}&includeAttributes=true`,
      { qlQuery: "objectType != null" }   // fetch all objects in workspace
    );

    const values = data.values || [];
    values.forEach(obj => {
      const attr = (key) => {
        const a = (obj.attributes || []).find(
          x => x.objectTypeAttributeName === key
        );
        return a?.objectAttributeValues?.[0]?.displayValue || "";
      };
      assets.push({
        id:       String(obj.id   || ""),
        key:      obj.objectKey   || "",
        name:     attr("Name")    || obj.label || "",
        serial:   attr("Serial Number") || attr("SerialNumber") || "",
        email:    attr("User")    || attr("Owner") || attr("Email") || "",
        model:    attr("Model")   || "",
        location: attr("Location")|| "",
        status:   attr("Status")  || "",
      });
    });

    startAt += values.length;
    if (values.length < pageSize || startAt >= (data.total || 0)) break;
  }
  return assets;
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
// SYNC
// ══════════════════════════════════════════════════════════════
async function runSync() {
  if (isSyncing) { toast("Sync already running", "warning"); return; }
  if (!cfg.jiraUrl || !cfg.token) {
    toast("Configure Jira settings first", "warning"); return;
  }

  isSyncing = true;
  setSyncIndicator("syncing", "Syncing...");

  const syncPanel = document.getElementById("sync-location-list");

  try {
    toast("Fetching assets from Jira...", "warning");
    const jiraAssets = await fetchJiraAssets();

    // Group by location
    const byLocation = {};
    jiraAssets.forEach(a => {
      const loc = (a.location || "UNKNOWN").trim();
      if (!byLocation[loc]) byLocation[loc] = [];
      byLocation[loc].push(a);
    });

    const locations  = Object.keys(byLocation);
    let syncState    = locations.map(l => ({
      name: l, done: 0, total: byLocation[l].length, status: "idle"
    }));
    updateSyncPanel(syncPanel, syncState);

    await Excel.run(async (context) => {
      for (let li = 0; li < locations.length; li++) {
        const loc       = locations[li];
        const assets    = byLocation[loc];
        const sheetName = locationSheetName(loc);

        syncState[li].status = "running";
        updateSyncPanel(syncPanel, syncState);

        const sheet = await ensureSheet(context, sheetName);
        await ensureHeaders(context, sheet);

        // Build lookup maps from existing rows
        const existing = await readSheetRows(context, sheet);
        const byId     = {};
        const bySerial = {};
        existing.forEach((row, idx) => {
          const id  = String(row[COL.ASSET_ID] || "").trim();
          const ser = String(row[COL.SERIAL]   || "").trim();
          if (id)  byId[id]   = idx;
          if (ser) bySerial[ser] = idx;
        });

        const now = new Date().toISOString();

        for (const asset of assets) {
          const existingIdx =
            byId[asset.id]         !== undefined ? byId[asset.id] :
            bySerial[asset.serial] !== undefined && asset.serial ? bySerial[asset.serial] :
            -1;

          if (existingIdx >= 0) {
            // UPDATE — preserve user-entered cols (DAYS, NOTE, ACTION)
            const updateRange = sheet.getRangeByIndexes(existingIdx + 1, 0, 1, COL_COUNT);
            updateRange.load("values");
            await context.sync();
            const cur = updateRange.values[0];
            cur[COL.ASSET_ID]    = asset.id;
            cur[COL.ASSET_KEY]   = asset.key;
            cur[COL.DEVICE_NAME] = asset.name;
            cur[COL.SERIAL]      = asset.serial;
            cur[COL.USER_EMAIL]  = asset.email;
            cur[COL.MODEL]       = asset.model;
            cur[COL.LOCATION]    = loc;
            cur[COL.STATUS]      = asset.status;
            cur[COL.SYNC_STATUS] = "JIRA";
            cur[COL.LAST_SYNC]   = now;
            // cur[COL.DAYS], cur[COL.NOTE], cur[COL.ACTION] — intentionally preserved
            updateRange.values   = [cur];
            await context.sync();
          } else {
            // INSERT new row
            const row             = Array(COL_COUNT).fill("");
            row[COL.ASSET_ID]    = asset.id;
            row[COL.ASSET_KEY]   = asset.key;
            row[COL.DEVICE_NAME] = asset.name;
            row[COL.SERIAL]      = asset.serial;
            row[COL.USER_EMAIL]  = asset.email;
            row[COL.MODEL]       = asset.model;
            row[COL.LOCATION]    = loc;
            row[COL.STATUS]      = asset.status;
            row[COL.SYNC_STATUS] = "JIRA";
            row[COL.LAST_SYNC]   = now;
            await appendRow(context, sheet, row);
            // Update lookup so duplicates within same sync don't insert twice
            byId[asset.id] = existing.length;
            if (asset.serial) bySerial[asset.serial] = existing.length;
            existing.push(row);
          }

          syncState[li].done++;
          if (syncState[li].done % 20 === 0) updateSyncPanel(syncPanel, syncState);
        }

        syncState[li].status = "done";
        syncState[li].done   = syncState[li].total;
        updateSyncPanel(syncPanel, syncState);
      }
    });

    // Persist last sync time
    cfg.lastSync = new Date().toISOString();
    Office.context.document.settings.set(CFG_KEYS.LAST_SYNC, cfg.lastSync);
    Office.context.document.settings.saveAsync();

    setSyncIndicator("ok", "Synced");
    setInner("last-sync-time", formatTime(cfg.lastSync));
    toast(`Sync complete — ${jiraAssets.length} assets across ${locations.length} location(s)`, "success");
    await refreshDashboard();

  } catch(e) {
    setSyncIndicator("ok", "Sync failed");
    toast("Sync error: " + e.message, "error");
  } finally {
    isSyncing = false;
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

          const deviceName = row[COL.DEVICE_NAME] || row[COL.ASSET_KEY] || "Unknown";
          const email      = row[COL.USER_EMAIL]  || "";
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
