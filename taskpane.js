/* ════════════════════════════════════════════════════════════
   Jira Asset Manager — Office Add-in
   taskpane.js — v3.0 (bypass-1000 via OS Version split)
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
// JIRA API — via Cloudflare Worker proxy (CORS)
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


// ══════════════════════════════════════════════════════════════
// FETCH ENGINE v3 — Bypass giới hạn 1000 record qua OS Version
//
// Quy trình cho mỗi objectTypeId:
//   Bước 1  — Lấy totalCount của toàn typeId
//   Bước 2  — Nếu < 1000: fetch thẳng, ghi Excel
//   Bước 3  — Nếu >= 1000: discover OS Versions thực tế từ data
//   Bước 4  — Với mỗi OS Version: check count, fetch nếu < 1000
//   Bước 5  — Merge toàn bộ kết quả
//   Bước 6  — Dedup theo objectKey
//   Bước 7  — Kiểm tra expectedTotal === actualTotal
//   Bước 8  — Ghi Excel (chỉ sau khi merge + check xong)
//   Bước 9  — Chuyển sang typeId tiếp theo
// ══════════════════════════════════════════════════════════════

// Attribute ID cho "Version OS" (Windows Version) = 27291
const ATTR_OS_VERSION = "27291";

/**
 * Bước 1 — Lấy totalCount của một AQL query.
 * POST /object/aql/totalcount — không tốn quota object.
 * Trả về số nguyên >= 0, hoặc -1 nếu API lỗi.
 */
async function fetchTotalCount(qlQuery) {
  try {
    const data = await assetsPost("/object/aql/totalcount", { qlQuery });
    return typeof data.totalCount === "number" ? data.totalCount : 0;
  } catch (e) {
    console.warn("fetchTotalCount error:", e.message);
    return -1;
  }
}

/**
 * Fetch một page từ Assets AQL (internal).
 */
async function fetchPage(qlQuery, startAt, pageSize = 25) {
  return assetsPost("/object/aql", {
    qlQuery,
    startAt,
    maxResults:        pageSize,
    includeAttributes: true,
  });
}

/**
 * Fetch toàn bộ object của một AQL query đã đảm bảo count < 1000.
 * Pagination trong phạm vi startAt [0, 999].
 *
 * @param {string} qlQuery  - AQL đảm bảo count < 1000
 * @param {string} label    - Nhãn cho console log
 * @returns {Array}         - Mảng asset đã parse
 */
async function fetchAllFromQuery(qlQuery, label = "") {
  const assets   = [];
  const pageSize = 25;
  let   startAt  = 0;

  while (true) {
    const data   = await fetchPage(qlQuery, startAt, pageSize);
    const values = data.values || [];
    values.forEach(obj => assets.push(parseAsset(obj)));

    if (values.length < pageSize) break;   // trang cuối

    startAt += values.length;
    if (startAt >= 1000) {
      // Query này count >= 1000 ngoài dự kiến — dừng để tránh miss data
      console.warn(`[${label}] Reached startAt=${startAt} >= 1000. Query too broad.`);
      break;
    }
  }

  console.log(`  [${label}] fetched=${assets.length}`);
  return assets;
}

/**
 * Bước 3 — Discover các giá trị OS Version thực tế của một typeId.
 *
 * Scan tối đa 1000 object đầu, collect unique values của attr 27291.
 * Luôn thêm bucket "__EMPTY__" để capture object không có OS Version.
 *
 * @param {string} typeId
 * @returns {string[]}  - Danh sách OS Version values + "__EMPTY__"
 */
async function discoverOsVersions(typeId) {
  const versions = new Set();
  const pageSize = 25;
  let   startAt  = 0;

  while (startAt < 1000) {
    const data   = await fetchPage(`objectTypeId = ${typeId}`, startAt, pageSize);
    const values = data.values || [];

    values.forEach(obj => {
      const attr = (obj.attributes || []).find(
        x => String(x.objectTypeAttributeId) === ATTR_OS_VERSION
      );
      const val = attr?.objectAttributeValues?.[0]?.displayValue;
      if (val && val.trim()) versions.add(val.trim());
    });

    if (values.length < pageSize) break;
    startAt += values.length;
  }

  const result = [...versions, "__EMPTY__"];
  console.log(`[typeId=${typeId}] OS Versions discovered (${result.length}):`, result);
  return result;
}

/**
 * Bước 2–7 — Core engine cho một objectTypeId.
 *
 * - totalCount < 1000  → fetch thẳng
 * - totalCount >= 1000 → chia theo OS Version, mỗi bucket < 1000 → fetch
 * - Merge tất cả → dedup theo objectKey → integrity check
 *
 * @param {string} typeId
 * @returns {{ assets: Array, expectedTotal: number, ok: boolean }}
 */
async function fetchOneTypeId(typeId) {
  const baseAql       = `objectTypeId = ${typeId}`;

  // Bước 1: lấy totalCount
  const expectedTotal = await fetchTotalCount(baseAql);
  console.log(`\n══ typeId=${typeId} ══ expectedTotal=${expectedTotal}`);
  toast(`[typeId=${typeId}] Tổng: ${expectedTotal} assets`, "warning");

  let rawAssets = [];

  if (expectedTotal >= 0 && expectedTotal < 1000) {
    // ── Bước 2: Dưới giới hạn → fetch thẳng ──────────────────
    console.log(`  → Fetch thẳng (count < 1000)`);
    rawAssets = await fetchAllFromQuery(baseAql, `typeId=${typeId}`);

  } else {
    // ── Bước 3: >= 1000 → chia theo OS Version ────────────────
    console.log(`  → count >= 1000, chia theo OS Version…`);
    toast(`[typeId=${typeId}] Đang discover OS Versions…`, "warning");

    const osVersions = await discoverOsVersions(typeId);

    // Bước 4: Với mỗi OS Version → check count → fetch
    for (const ver of osVersions) {
      const verAql = ver === "__EMPTY__"
        ? `${baseAql} AND "Version OS" IS EMPTY`
        : `${baseAql} AND "Version OS" = "${ver}"`;

      const verCount = await fetchTotalCount(verAql);
      console.log(`  [Version OS="${ver}"] count=${verCount}`);

      if (verCount === 0) continue;

      if (verCount < 1000) {
        // Bước 4a: An toàn → fetch trực tiếp
        const bucket = await fetchAllFromQuery(verAql, `typeId=${typeId},ver=${ver}`);
        rawAssets.push(...bucket);
        toast(`  [typeId=${typeId}] "${ver}": ${bucket.length} assets`, "warning");
      } else {
        // Bucket vẫn >= 1000 — cảnh báo, fetch tối đa có thể
        console.error(`  [WARN] "${ver}" count=${verCount} >= 1000. Fetch tối đa 999 records.`);
        toast(`⚠ [typeId=${typeId}] "${ver}" có ${verCount} records >= 1000! Có thể thiếu data.`, "warning");
        const bucket = await fetchAllFromQuery(verAql, `OVERFLOW:typeId=${typeId},ver=${ver}`);
        rawAssets.push(...bucket);
      }
    }
  }

  // ── Bước 6: Dedup theo objectKey ──────────────────────────
  // Dùng Map để giữ lại bản ghi cuối cùng của mỗi key (latest wins)
  const uniqueAssets = [
    ...new Map(rawAssets.map(a => [a.key || a.id, a])).values()
  ];

  // ── Bước 7: Integrity check ────────────────────────────────
  const actualTotal = uniqueAssets.length;
  const ok          = expectedTotal < 0 || actualTotal === expectedTotal;

  console.log(
    `  Expected: ${expectedTotal} | Raw fetched: ${rawAssets.length} | ` +
    `After dedup: ${actualTotal} | OK: ${ok}`
  );

  if (!ok) {
    console.error(
      `[typeId=${typeId}] ⚠ Missing records: expected=${expectedTotal}, actual=${actualTotal}, ` +
      `missing=${expectedTotal - actualTotal}`
    );
    toast(
      `⚠ [typeId=${typeId}] Thiếu ${expectedTotal - actualTotal} records (${actualTotal}/${expectedTotal})`,
      "warning"
    );
  } else if (expectedTotal >= 0) {
    console.log(`  ✓ Integrity OK: ${actualTotal} === ${expectedTotal}`);
  }

  return { assets: uniqueAssets, expectedTotal, actualTotal, ok };
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

/**
 * Fetch toàn bộ assets qua tất cả typeIds, xử lý tuần tự.
 * Mỗi typeId: fetch → dedup → check hoàn tất trước khi sang typeId tiếp.
 * Sau khi xong tất cả → merge global (dedup cross-typeId theo objectKey).
 *
 * @returns {{ allAssets: Array, stats: Array }}
 */
async function fetchJiraAssets() {
  const typeIds = parseTypeIds();
  if (!typeIds.length) {
    toast("AQL Query chưa được cấu hình đúng", "error");
    return { allAssets: [], stats: [] };
  }

  // Global dedup: objectKey → asset (xử lý asset xuất hiện ở nhiều typeId)
  const globalMap = new Map();
  const stats     = [];

  for (let i = 0; i < typeIds.length; i++) {
    const typeId = typeIds[i];
    toast(`[${i+1}/${typeIds.length}] Bắt đầu typeId=${typeId}…`, "warning");

    // Bước 2–7 cho typeId này (fetch + dedup + integrity check)
    const result = await fetchOneTypeId(typeId);

    // Merge vào global map
    let crossAdded = 0;
    result.assets.forEach(a => {
      const key = a.key || a.id;
      if (!globalMap.has(key)) {
        globalMap.set(key, a);
        crossAdded++;
      }
    });

    stats.push({
      typeId,
      expectedTotal: result.expectedTotal,
      actualTotal:   result.actualTotal,
      addedGlobal:   crossAdded,
      ok:            result.ok,
    });

    console.log(
      `[typeId=${typeId}] DONE. addedGlobal=${crossAdded}. Global total so far: ${globalMap.size}`
    );
    // Bước 9: vòng lặp tự động chuyển sang typeId tiếp theo
  }

  const allAssets = [...globalMap.values()];

  console.log(`\nfetchJiraAssets DONE: ${allAssets.length} assets (${typeIds.length} typeId(s))`);
  console.log("=== FETCH SUMMARY ===");
  stats.forEach(s =>
    console.log(
      `  typeId=${s.typeId}: expected=${s.expectedTotal}, actual=${s.actualTotal}, ` +
      `addedGlobal=${s.addedGlobal}, OK=${s.ok}`
    )
  );

  return { allAssets, stats };
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

async function readSheetRows(context, sheet) {
  const used = sheet.getUsedRangeOrNullObject(true);
  await context.sync();
  if (used.isNullObject) return [];
  used.load(["values", "rowCount"]);
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
          if (r[COL.SYNC_STATUS] === "LOCAL")                        { local++;  locLocal++; }
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
// ══════════════════════════════════════════════════════════════
async function createLocationSheets() {
  if (!cfg.jiraUrl || !cfg.token) {
    toast("Configure Jira settings first", "warning"); return;
  }
  toast("Fetching assets from Jira to discover locations...", "warning");
  try {
    const { allAssets } = await fetchJiraAssets();
    const locations = [...new Set(allAssets.map(a => a.location).filter(Boolean))];

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
// ASSET → ROW
// ══════════════════════════════════════════════════════════════
function assetToRow(asset, loc, now, existingRow) {
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
  // Preserve user-entered fields from existing row
  // COL.DAYS, COL.NOTE, COL.ACTION, COL.CASE_JIRA are already carried via existingRow spread
  return row;
}

// ══════════════════════════════════════════════════════════════
// WRITE LOCATION SHEET
// Upsert logic:
//   - Match by Asset ID (primary) hoặc Serial (fallback)
//   - UPDATE tại chỗ nếu đã tồn tại
//   - INSERT batch cuối sheet nếu mới
//   - Mark "Not in Jira" nếu row cũ không còn trong batch
//
// Ghi chú Excel batch:
//   - Không dùng appendRow từng dòng (quá chậm)
//   - Dùng getRangeByIndexes(startRow, 0, N, COL_COUNT).values = matrix
// ══════════════════════════════════════════════════════════════
async function writeLocationSheet(sheetName, assets, now) {
  await Excel.run(async (context) => {
    const sheet = await ensureSheet(context, sheetName);
    await ensureHeaders(context, sheet);

    // Load existing rows
    const existing = await readSheetRows(context, sheet);

    // Build lookup maps
    const byId     = {};
    const byKey    = {};
    const bySerial = {};
    existing.forEach((row, idx) => {
      const id     = String(row[COL.ASSET_ID]   || "").trim();
      const key    = String(row[COL.ASSET_KEY]  || "").trim();
      const serial = String(row[COL.SERIAL]     || "").trim();
      if (id)     byId[id]         = idx;
      if (key)    byKey[key]       = idx;
      if (serial) bySerial[serial] = idx;
    });

    // Track which existing rows were matched
    const updatedIdxs = new Set();
    // Track Asset IDs from Jira
    const jiraIdSet   = new Set(assets.map(a => a.id).filter(Boolean));
    const jiraKeySet  = new Set(assets.map(a => a.key).filter(Boolean));

    // ── A: UPDATE existing rows ──────────────────────────────
    // Collect updates as { rowIdx, newRow } for batch write
    const updates = [];

    assets.forEach(asset => {
      // Priority: ID → Key → Serial
      let idx = -1;
      if (asset.id   && byId[asset.id]       !== undefined) idx = byId[asset.id];
      else if (asset.key && byKey[asset.key] !== undefined) idx = byKey[asset.key];
      else if (asset.serial && bySerial[asset.serial] !== undefined) idx = bySerial[asset.serial];

      if (idx >= 0) {
        const newRow = assetToRow(asset, asset.location || "", now, existing[idx]);
        if (newRow[COL.VALIDATION] === "Not in Jira") newRow[COL.VALIDATION] = "";
        updates.push({ idx, newRow });
        updatedIdxs.add(idx);
      }
    });

    // Batch write updates (one call per row — Excel.run context batches these)
    for (const { idx, newRow } of updates) {
      sheet.getRangeByIndexes(idx + 1, 0, 1, COL_COUNT).values = [newRow];
    }

    // ── B: Mark rows no longer in Jira ───────────────────────
    existing.forEach((row, idx) => {
      if (updatedIdxs.has(idx))             return;
      if (row[COL.SYNC_STATUS] === "LOCAL") return;
      const id  = String(row[COL.ASSET_ID]  || "").trim();
      const key = String(row[COL.ASSET_KEY] || "").trim();
      // Only mark if this row previously had a Jira ID/Key that is now gone
      if (!id && !key) return;
      if (id && jiraIdSet.has(id))   return;
      if (key && jiraKeySet.has(key)) return;
      const cell = sheet.getRangeByIndexes(idx + 1, COL.VALIDATION, 1, 1);
      cell.values           = [["Not in Jira"]];
      cell.format.font.color = "#f59e0b";
    });

    // ── C: INSERT new assets (batch) ─────────────────────────
    const toInsert = assets.filter(a => {
      if (a.id     && byId[a.id]       !== undefined) return false;
      if (a.key    && byKey[a.key]     !== undefined) return false;
      if (a.serial && bySerial[a.serial] !== undefined) return false;
      return true;
    });

    if (toInsert.length > 0) {
      // Find next empty row after used range
      const used = sheet.getUsedRangeOrNullObject(true);
      await context.sync();
      let startRow = 1; // row after header
      if (!used.isNullObject) {
        used.load("rowCount");
        await context.sync();
        startRow = used.rowCount;
      }

      // Batch insert all new rows in ONE range write
      const matrix = toInsert.map(a => assetToRow(a, a.location || "", now, null));
      sheet.getRangeByIndexes(startRow, 0, matrix.length, COL_COUNT).values = matrix;
    }

    await context.sync();

    // ── D: Integrity check ───────────────────────────────────
    const afterRows = await readSheetRows(context, sheet);
    const sheetCount = afterRows.length;

    console.log(
      `[${sheetName}] updated=${updates.length} | inserted=${toInsert.length} | ` +
      `sheetRows=${sheetCount} | jiraAssets=${assets.length}`
    );

    if (sheetCount < assets.length) {
      console.error(
        `[${sheetName}] ⚠ INTEGRITY: sheetRows(${sheetCount}) < jiraAssets(${assets.length}). ` +
        `Missing ${assets.length - sheetCount} rows.`
      );
    }
  });
}

// ══════════════════════════════════════════════════════════════
// SYNC ENGINE — Main entry point
//
// Quy trình:
//   1. Với mỗi typeId: fetch (bypass-1000) → merge → dedup → check
//   2. Bước 8: Ghi Excel chỉ sau khi typeId đó hoàn tất hoàn toàn
//   3. Bước 9: Chuyển sang typeId tiếp theo
//   4. Sau tất cả typeIds: verify Excel row count
// ══════════════════════════════════════════════════════════════
async function runSync() {
  if (isSyncing) { toast("Sync đang chạy, vui lòng chờ…", "warning"); return; }
  if (!cfg.jiraUrl || !cfg.token || !cfg.workerUrl) {
    toast("Kiểm tra lại Settings (URL, token, worker)", "warning"); return;
  }

  isSyncing = true;
  setSyncIndicator("syncing", "Syncing...");
  const syncPanel = document.getElementById("sync-location-list");
  if (syncPanel) {
    syncPanel.innerHTML = `<div class="empty-state"><div class="spinner"></div>Đang tải từ Jira...</div>`;
  }

  try {
    const typeIds = parseTypeIds();
    if (!typeIds.length) {
      toast("AQL Query chưa đúng — cần objectTypeId IN (...) hoặc objectTypeId = N", "error");
      return;
    }

    const now        = new Date().toISOString();
    const syncedInfo = []; // { typeId, sheetName, expected, actual, sheetRows }

    // ══ Bước 2–9: Xử lý tuần tự từng typeId ══════════════════
    for (let i = 0; i < typeIds.length; i++) {
      const typeId    = typeIds[i];
      const sheetName = `_TYPE_${typeId}`;

      toast(`[${i+1}/${typeIds.length}] Đang xử lý typeId=${typeId}…`, "warning");

      if (syncPanel) {
        updateSyncPanel(syncPanel, typeIds.map((tid, j) => ({
          name:   `_TYPE_${tid}`,
          done:   0,
          total:  0,
          status: j < i ? "done" : (j === i ? "running" : "waiting"),
        })));
      }

      // Bước 2–7: fetch + dedup + integrity check cho typeId này
      const result = await fetchOneTypeId(typeId);
      const assets  = result.assets; // đã dedup, đã integrity-check

      // Bước 8: Ghi Excel SAU KHI merge + check xong
      // (không ghi từng OS-version bucket riêng lẻ)
      toast(
        `[typeId=${typeId}] Ghi ${assets.length} assets vào sheet ${sheetName}…`,
        "warning"
      );
      await writeLocationSheet(sheetName, assets, now);

      // Đếm rows thực tế trong sheet để verify
      let sheetRows = 0;
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const rows  = await readSheetRows(context, sheet);
        sheetRows   = rows.length;
      });

      // Log bước 8: so sánh jiraAssets vs sheetRows
      console.log(
        `\n=== EXCEL WRITE — typeId=${typeId} ===\n` +
        `  TotalCount Jira : ${result.expectedTotal}\n` +
        `  Total fetched   : ${result.actualTotal}\n` +
        `  Total inserted  : ${assets.length}\n` +
        `  Total rows sheet: ${sheetRows}`
      );

      if (sheetRows < assets.length) {
        console.error(
          `[typeId=${typeId}] ⚠ sheetRows(${sheetRows}) < jiraAssets(${assets.length}). ` +
          `Missing ${assets.length - sheetRows} rows in Excel.`
        );
        toast(`⚠ [typeId=${typeId}] Ghi thiếu ${assets.length - sheetRows} rows vào Excel!`, "warning");
      }

      syncedInfo.push({
        typeId,
        sheetName,
        expected: result.expectedTotal,
        actual:   result.actualTotal,
        ok:       result.ok,
        sheetRows,
      });

      // Bước 9: vòng lặp for tự động chuyển sang typeId tiếp theo
    }

    // ══ Tổng kết sau tất cả typeIds ══════════════════════════
    const totalExpected = syncedInfo.reduce((s, x) => s + (x.expected >= 0 ? x.expected : 0), 0);
    const totalActual   = syncedInfo.reduce((s, x) => s + x.actual,   0);
    const totalSheetRows = syncedInfo.reduce((s, x) => s + x.sheetRows, 0);

    console.log(`\n=== SYNC HOÀN TẤT ===`);
    console.log(`TotalCount Jira (sum) : ${totalExpected}`);
    console.log(`Total fetched (deduped): ${totalActual}`);
    console.log(`Total rows trong Excel : ${totalSheetRows}`);
    syncedInfo.forEach(s =>
      console.log(
        `  typeId=${s.typeId} [${s.sheetName}]: expected=${s.expected}, ` +
        `actual=${s.actual}, sheetRows=${s.sheetRows}, OK=${s.ok}`
      )
    );

    const allOk = syncedInfo.every(s => s.ok) && totalSheetRows >= totalActual;

    if (syncPanel) {
      updateSyncPanel(syncPanel, syncedInfo.map(s => ({
        name:   s.sheetName,
        done:   s.sheetRows,
        total:  s.expected >= 0 ? s.expected : s.actual,
        status: s.ok ? "done" : "error",
      })));
    }

    // Push LOCAL rows lên Jira
    toast("Đang push LOCAL assets lên Jira…", "warning");
    await pushLocalAssets();

    // Lưu timestamp
    cfg.lastSync = new Date().toISOString();
    Office.context.document.settings.set(CFG_KEYS.LAST_SYNC, cfg.lastSync);
    Office.context.document.settings.saveAsync();

    setSyncIndicator("ok", "Synced");
    setInner("last-sync-time", formatTime(cfg.lastSync));
    toast(
      `✓ Sync hoàn tất — ${totalActual} assets, ${typeIds.length} typeId(s), ${totalSheetRows} rows trong Excel` +
      (allOk ? "" : " ⚠ Có lỗi, kiểm tra console"),
      allOk ? "success" : "warning"
    );

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
// PUSH LOCAL ASSETS → JIRA
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
          if (row[COL.ASSET_ID])               continue; // already pushed

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
      toast(
        `LOCAL push: ${pushed} thành công${failed ? ", " + failed + " lỗi" : ""}`,
        failed ? "warning" : "success"
      );
    }
  } catch(e) {
    console.warn("pushLocalAssets:", e.message);
  }
}

function updateSyncPanel(el, state) {
  el.innerHTML = state.map(l => `
    <div class="location-item">
      <div class="loc-header">
        <span class="loc-name">${l.name}</span>
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
    const { allAssets } = await fetchJiraAssets();
    const jiraBySerial  = {};
    allAssets.forEach(a => { if (a.serial) jiraBySerial[a.serial.trim()] = a; });

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
      const serialSeen = {};

      for (const sheetName of sheets) {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const rows  = await readSheetRows(context, sheet);

        for (let i = 0; i < rows.length; i++) {
          const row      = rows[i];
          const serial   = String(row[COL.SERIAL]   || "").trim();
          const assetId  = String(row[COL.ASSET_ID] || "").trim();
          const locField = String(row[COL.LOCATION] || "").trim();
          let valid = "OK";

          if (row[COL.SYNC_STATUS] === "JIRA" && !assetId) {
            valid = "Missing Asset ID";
            errors["Missing Asset ID"].push({ sheet: sheetName, row: i + 2 });
          }

          if (serial && valid === "OK") {
            if (serialSeen[serial]) {
              valid = "Duplicate Serial";
              errors["Duplicate Serial"].push({ sheet: sheetName, row: i + 2 });
              errors["Duplicate Serial"].push(serialSeen[serial]);
            } else {
              serialSeen[serial] = { sheet: sheetName, row: i + 2 };
            }
          }

          if (locField && valid === "OK") {
            const expected = locationSheetName(locField);
            if (expected !== sheetName) {
              valid = "Location Changed";
              errors["Location Changed"].push({ sheet: sheetName, row: i + 2 });
            }
          }

          const cell = sheet.getRangeByIndexes(i + 1, COL.VALIDATION, 1, 1);
          cell.values            = [[valid]];
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
    const count  = list.length;
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
          if (row[COL.CASE_JIRA]) { skipped++; continue; }

          const deviceName = row[COL.HOSTNAME] || row[COL.ASSET_KEY] || "Unknown";
          const email      = row[COL.ASSIGNED] || row[COL.USERNAME] || "";
          const serial     = row[COL.SERIAL]   || "";
          const note       = row[COL.NOTE]     || "";
          const rowDays    = row[COL.DAYS]     || days;

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
            const res = await jiraPost("/api/3/issue", ticketBody);
            const ticketKey = res.key || "";

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
      `Done — created: ${created}, skipped: ${skipped}${failed ? ", failed: " + failed : ""}`,
      failed ? "warning" : "success"
    );
    scanPendingRows();
  } catch(e) {
    toast("Ticket error: " + e.message, "error");
  }
}

// ══════════════════════════════════════════════════════════════
// CONNECTION TEST
// ══════════════════════════════════════════════════════════════
async function testConnection() {
  const el = document.getElementById("conn-test-result");
  el.style.display = "block";
  el.style.borderColor = "var(--border)";
  el.innerHTML = `<div style="display:flex;gap:8px;align-items:center">
    <div class="spinner"></div><span>Testing connection…</span>
  </div>`;

  saveConfig();

  if (!cfg.workerUrl) {
    el.innerHTML = "✗ Worker URL chưa được điền";
    el.style.display = "block";
    el.style.borderColor = "var(--red)";
    toast("Điền Cloudflare Worker URL trước", "error");
    return;
  }

  try {
    const me = await jiraGet("/api/3/myself");
    let assetsOk = "";
    try {
      const countData = await assetsPost("/object/aql/totalcount", { qlQuery: "objectType != null" });
      assetsOk = ` · Assets API ✓ (${countData.totalCount ?? "?"} objects total)`;
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