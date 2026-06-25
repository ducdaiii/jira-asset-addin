

// ── COLUMN MAP ────────────────────────────────────────────────
const COL = {
  ASSET_ID:     0,  // Jira Object ID
  ASSET_KEY:    1,  // Jira Object Key
  HOSTNAME:     2,  // attr 1737
  SERIAL:       3,  // attr 5194
  STATUS:       4,  // attr 5052
  LOCATION:     5,  // attr 30125
  REGION:       6,  // attr 27292
  MANUFACTURER: 7,  // attr 6608
  MODEL:        8,  // attr 6609
  OS:           9,  // attr 30345
  OS_VERSION:   10, // attr 27291
  OS_BUILD:     11, // attr 27290
  CPU:          12, // attr 6610
  IP:           13, // attr 5208
  MAC:          14, // attr 5209
  NETWORK:      15, // attr 5210
  ANTIVIRUS:    16, // attr 6612
  USERNAME:     17, // attr 5200
  ASSIGNED:     18, // attr 26690
  FIRST_SEEN:   19, // attr 5205
  LAST_SEEN:    20, // attr 5206
  PURCHASE:     21, // attr 5203
  WARRANTY:     22, // attr 6615
  TENANT_ID:    23, // attr 26398
  LANSWEEPER:   24, // attr 5207
  // ── User-managed cols (never overwritten by sync) ──
  DAYS:         25,
  NOTE:         26,
  CASE_JIRA:    27,
  VALIDATION:   28,
  SYNC_STATUS:  29,
  LAST_SYNC:    30,
  ACTION:       31,
};
const COL_COUNT = 32;

// Cols that sync writes (indices 0–24 + 29–30); user cols 25–28,31 are preserved
const JIRA_COLS = [
  COL.ASSET_ID, COL.ASSET_KEY, COL.HOSTNAME, COL.SERIAL, COL.STATUS,
  COL.LOCATION, COL.REGION, COL.MANUFACTURER, COL.MODEL, COL.OS,
  COL.OS_VERSION, COL.OS_BUILD, COL.CPU, COL.IP, COL.MAC,
  COL.NETWORK, COL.ANTIVIRUS, COL.USERNAME, COL.ASSIGNED,
  COL.FIRST_SEEN, COL.LAST_SEEN, COL.PURCHASE, COL.WARRANTY,
  COL.TENANT_ID, COL.LANSWEEPER, COL.SYNC_STATUS, COL.LAST_SYNC,
];

const HEADERS = [
  "Asset ID","Asset Key","Hostname","Serial Number",
  "Status","Location","Region","Manufacturer","Model",
  "Operating System","Windows Version","Windows Build","CPU",
  "IP Address","MAC Address","Network Name","Antivirus",
  "Username","Assigned User","First Seen","Last Seen",
  "Purchase Date","Warranty Expire","Tenant ID / Source ID","Lansweeper URL",
  "Days","Note","Case Jira","Validation","Sync Status","Last Sync","Action"
];

// ── CONFIG ────────────────────────────────────────────────────
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

let cfg       = {};
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
    jiraUrl:     s.get(CFG_KEYS.JIRA_URL)    || "",
    email:       s.get(CFG_KEYS.EMAIL)        || "",
    token:       s.get(CFG_KEYS.TOKEN)        || "",
    cloudId:     s.get(CFG_KEYS.CLOUD_ID)     || "",
    workspaceId: s.get(CFG_KEYS.WORKSPACE_ID) || "",
    projectKey:  s.get(CFG_KEYS.PROJECT_KEY)  || "IT",
    workerUrl:   s.get(CFG_KEYS.WORKER_URL)   || "",
    aqlQuery:    s.get(CFG_KEYS.AQL_QUERY)    || "objectTypeId IN (525,527,529)",
    lastSync:    s.get(CFG_KEYS.LAST_SYNC)    || null,
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
    t.classList.toggle("active", t.dataset.tab === tab));
  document.querySelectorAll(".panel").forEach(p =>
    p.classList.toggle("active", p.id === `panel-${tab}`));
  if (tab === "ticket") scanPendingRows();
}

// ══════════════════════════════════════════════════════════════
// JIRA API — via Cloudflare Worker proxy
// ══════════════════════════════════════════════════════════════
function jiraBase() { return cfg.jiraUrl.replace(/\/+$/, ""); }

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
  if (!res.ok) throw new Error(`Jira ${res.status}: ${(await res.text().catch(() => "")).slice(0, 120)}`);
  return res.json();
}

async function jiraPost(path, body) {
  const res = await fetch(proxyUrl(`${jiraBase()}/rest${path}`), {
    method: "POST", headers: jiraHeaders(), body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`Jira ${res.status}: ${(await res.text().catch(() => "")).slice(0, 120)}`);
  return res.json();
}

async function assetsPost(path, body) {
  const res = await fetch(proxyUrl(`${assetsBase()}${path}`), {
    method: "POST", headers: jiraHeaders(), body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`Assets ${res.status}: ${(await res.text().catch(() => "")).slice(0, 120)}`);
  return res.json();
}

async function assetsPut(path, body) {
  const res = await fetch(proxyUrl(`${assetsBase()}${path}`), {
    method: "PUT", headers: jiraHeaders(), body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`Assets PUT ${res.status}: ${(await res.text().catch(() => "")).slice(0, 120)}`);
  return res.json();
}

// ══════════════════════════════════════════════════════════════
// FETCH — low level
// ══════════════════════════════════════════════════════════════

// Trả về tổng số record thực (không bị cap bởi /aql)
async function fetchTotalCount(qlQuery) {
  const res = await assetsPost("/object/aql/totalcount", { qlQuery });
  if (typeof res === "number")                   return res;
  if (res && typeof res.count      === "number") return res.count;
  if (res && typeof res.totalCount === "number") return res.totalCount;
  const v = parseInt(res, 10);
  return isNaN(v) ? 0 : v;
}

// Fetch 1 page
async function fetchPage(qlQuery, startAt, pageSize) {
  return assetsPost("/object/aql", {
    qlQuery, startAt, maxResults: pageSize, includeAttributes: true,
  });
}

// Parse 1 object từ API thành asset record
function parseAsset(obj) {
  const av = (attrId) => {
    const a = (obj.attributes || []).find(
      x => String(x.objectTypeAttributeId) === String(attrId)
    );
    return a?.objectAttributeValues?.[0]?.displayValue || "";
  };
  return {
    id:           String(obj.id ?? obj.objectId ?? ""),
    key:          String(obj.objectKey || obj.key || ""),
    hostname:     av(1737)  || obj.label || "",
    serial:       av(5194),
    status:       av(5052),
    location:     av(30125),
    region:       av(27292),
    manufacturer: av(6608),
    model:        av(6609),
    os:           av(30345),
    osVersion:    av(27291),
    osBuild:      av(27290),
    cpu:          av(6610),
    ip:           av(5208),
    mac:          av(5209),
    network:      av(5210),
    antivirus:    av(6612),
    username:     av(5200),
    assigned:     av(26690),
    firstSeen:    av(5205),
    lastSeen:     av(5206),
    purchase:     av(5203),
    warranty:     av(6615),
    tenantId:     av(26398),
    lansweeper:   av(5207),
  };
}

// ══════════════════════════════════════════════════════════════
// FETCH — pagination an toàn
//
// Điều kiện dừng: assets.length >= total (lấy từ firstPage.total)
// Hard stop kép:  startAt không vượt 1000 VÀ values rỗng
// Caller phải đảm bảo total < 1000 trước khi gọi hàm này
// ══════════════════════════════════════════════════════════════
async function fetchAllUnderLimit(qlQuery) {
  const PAGE     = 25;
  const HARD_CAP = 1000; // Jira Assets hard limit
  const assets   = [];

  const firstPage = await fetchPage(qlQuery, 0, PAGE);
  const total     = typeof firstPage.total === "number" ? firstPage.total : 0;
  const firstVals = firstPage.values || [];

  if (total === 0 || firstVals.length === 0) return assets;
  firstVals.forEach(o => assets.push(parseAsset(o)));
  console.log(`[fetchAll] "${qlQuery}": total=${total}, loaded=${assets.length}`);

  let startAt = firstVals.length;
  // Dừng khi: đã đủ total HOẶC startAt đã chạm hard cap HOẶC API trả rỗng
  while (assets.length < Math.min(total, HARD_CAP) && startAt < HARD_CAP) {
    const data   = await fetchPage(qlQuery, startAt, PAGE);
    const values = data.values || [];
    if (values.length === 0) break;
    values.forEach(o => assets.push(parseAsset(o)));
    startAt += values.length;
    console.log(`[fetchAll] "${qlQuery}": loaded=${assets.length}/${Math.min(total, HARD_CAP)}`);
  }

  return assets;
}

// ══════════════════════════════════════════════════════════════
// FETCH — discover OS Version values cho 1 typeId
//
// Vì /aql chỉ trả tối đa 1000 records, ta không thể dùng 1 page
// để discover tất cả versions nếu typeId > 1000.
// Giải pháp: dùng nhiều startAt cách nhau 1000 để lấy mẫu đại diện.
// Mỗi "mẫu" lấy 1 page nhỏ để collect thêm version values mới.
// Dừng khi không còn version mới xuất hiện sau 2 window liên tiếp.
// ══════════════════════════════════════════════════════════════
async function fetchOsVersions(typeId, totalCount) {
  const ATTR_OS_VERSION = "27291";
  const seen            = new Set();
  const PAGE            = 200; // đủ lớn để lấy nhiều version/page
  const windows         = Math.ceil(totalCount / 1000);

  for (let w = 0; w < windows; w++) {
    const startAt = w * 1000;
    const data    = await fetchPage(`objectTypeId = ${typeId}`, startAt, PAGE);
    const values  = data.values || [];
    if (values.length === 0) break;
    values.forEach(obj => {
      const a = (obj.attributes || []).find(
        x => String(x.objectTypeAttributeId) === ATTR_OS_VERSION
      );
      seen.add(a?.objectAttributeValues?.[0]?.displayValue || "");
    });
    console.log(`[fetchOsVersions] typeId=${typeId} window=${w}: versions so far=[${[...seen].join(", ")}]`);
  }

  const versions = [...seen];
  console.log(`[fetchOsVersions] typeId=${typeId}: ${versions.length} versions: [${versions.join(", ")}]`);
  return versions;
}

// ══════════════════════════════════════════════════════════════
// FETCH — discover OS Build values trong 1 sub-query (version)
// ══════════════════════════════════════════════════════════════
async function fetchOsBuilds(parentQuery, subTotal) {
  const ATTR_OS_BUILD = "27290";
  const seen          = new Set();
  const PAGE          = 200;
  const windows       = Math.ceil(subTotal / 1000);

  for (let w = 0; w < windows; w++) {
    const startAt = w * 1000;
    const data    = await fetchPage(parentQuery, startAt, PAGE);
    const values  = data.values || [];
    if (values.length === 0) break;
    values.forEach(obj => {
      const a = (obj.attributes || []).find(
        x => String(x.objectTypeAttributeId) === ATTR_OS_BUILD
      );
      seen.add(a?.objectAttributeValues?.[0]?.displayValue || "");
    });
  }

  const builds = [...seen];
  console.log(`[fetchOsBuilds] "${parentQuery}": ${builds.length} builds: [${builds.join(", ")}]`);
  return builds;
}

// Build AQL sub-query an toàn
function buildSubQuery(base, attrName, value) {
  if (value === "") return `${base} AND "${attrName}" is EMPTY`;
  return `${base} AND "${attrName}" = "${value.replace(/"/g, '\\"')}"`;
}

// ══════════════════════════════════════════════════════════════
// FETCH — toàn bộ assets của 1 objectTypeId
//
// Luồng:
//   totalCount < 1000  → fetchAllUnderLimit thẳng
//   totalCount >= 1000 → chia theo "Version OS"
//     subTotal < 1000  → fetchAllUnderLimit
//     subTotal >= 1000 → chia thêm theo "OS Build"
//       buildTotal < 1000  → fetchAllUnderLimit
//       buildTotal >= 1000 → log warning, fetch tối đa 1000
//   Dedup bằng objectKey, kiểm tra integrity vs totalCount
// ══════════════════════════════════════════════════════════════
async function fetchByTypeId(typeId) {
  const baseQuery  = `objectTypeId = ${typeId}`;
  const totalCount = await fetchTotalCount(baseQuery);
  console.log(`[fetchByTypeId] typeId=${typeId}, totalCount=${totalCount}`);
  toast(`objectTypeId=${typeId}: ${totalCount} records`, "warning");

  // ── Case 1: < 1000 → fetch thẳng ─────────────────────────
  if (totalCount < 1000) {
    const assets = await fetchAllUnderLimit(baseQuery);
    console.log(`[fetchByTypeId] typeId=${typeId}: done (under limit), got=${assets.length}`);
    return assets;
  }

  // ── Case 2: >= 1000 → phân chia theo Version OS ──────────
  toast(`objectTypeId=${typeId}: ${totalCount} records, đang phân chia theo Version OS...`, "warning");
  const versions  = await fetchOsVersions(typeId, totalCount);
  const allAssets = [];

  for (let vi = 0; vi < versions.length; vi++) {
    const ver      = versions[vi];
    const subQuery = buildSubQuery(baseQuery, "Version OS", ver);
    toast(`[${vi + 1}/${versions.length}] Version OS="${ver || "(blank)"}"...`, "warning");

    const subTotal = await fetchTotalCount(subQuery);
    console.log(`[fetchByTypeId] ver="${ver}": subTotal=${subTotal}`);

    if (subTotal === 0) continue;

    // ── Case 2a: sub < 1000 → fetch thẳng ─────────────────
    if (subTotal < 1000) {
      const assets = await fetchAllUnderLimit(subQuery);
      assets.forEach(a => allAssets.push(a));
      console.log(`[fetchByTypeId] ver="${ver}": fetched=${assets.length}`);
      continue;
    }

    // ── Case 2b: sub >= 1000 → chia tiếp theo OS Build ────
    toast(`  "${ver}": ${subTotal} records, đang phân chia theo OS Build...`, "warning");
    const builds = await fetchOsBuilds(subQuery, subTotal);

    for (const build of builds) {
      const buildQuery = buildSubQuery(subQuery, "OS Build", build);
      const buildTotal = await fetchTotalCount(buildQuery);
      console.log(`  [fetchByTypeId] build="${build}": buildTotal=${buildTotal}`);

      if (buildTotal === 0) continue;

      if (buildTotal >= 1000) {
        // Vẫn >= 1000 sau 2 tầng — fetch tối đa 1000, log warning
        console.error(`  [WARN] build="${build}": ${buildTotal} >= 1000, chỉ lấy được 1000 đầu`);
        toast(`⚠ Build "${build}": ${buildTotal} records, chỉ lấy 1000 đầu`, "warning");
      }

      const assets = await fetchAllUnderLimit(buildQuery);
      assets.forEach(a => allAssets.push(a));
      console.log(`  [fetchByTypeId] build="${build}": fetched=${assets.length}`);
    }
  }

  // ── Dedup bằng objectKey (key ưu tiên hơn id) ────────────
  const uniqueMap = new Map();
  allAssets.forEach(a => {
    const k = a.key || a.id;
    if (!k) {
      // Record không có key và id — vẫn giữ, dùng index làm key tạm
      uniqueMap.set(`__nokey_${uniqueMap.size}`, a);
    } else if (!uniqueMap.has(k)) {
      uniqueMap.set(k, a);
    }
  });
  const uniqueAssets = [...uniqueMap.values()];

  // ── Integrity check ───────────────────────────────────────
  if (uniqueAssets.length !== totalCount) {
    const msg = `[INTEGRITY] typeId=${typeId}: expected=${totalCount}, got=${uniqueAssets.length}`;
    console.error(msg);
    toast(`⚠ ${msg}`, "warning");
  } else {
    console.log(`[fetchByTypeId] typeId=${typeId}: integrity OK ${uniqueAssets.length}/${totalCount}`);
  }

  return uniqueAssets;
}

// ── Parse typeIds từ AQL config ──────────────────────────────
function parseTypeIds() {
  const raw = (cfg.aqlQuery || "").trim();
  const m1  = raw.match(/objectTypeId\s+IN\s*\(([^)]+)\)/i);
  if (m1) return m1[1].split(",").map(s => s.trim()).filter(Boolean);
  const m2  = raw.match(/objectTypeId\s*=\s*(\d+)/i);
  if (m2) return [m2[1]];
  return [];
}

// ── Fetch tất cả assets, dedup global bằng objectKey ─────────
async function fetchJiraAssets() {
  const typeIds = parseTypeIds();
  if (!typeIds.length) { toast("AQL Query chưa được cấu hình đúng", "error"); return []; }

  const seenKeys  = new Set();
  const allAssets = [];

  for (let i = 0; i < typeIds.length; i++) {
    const typeId = typeIds[i];
    toast(`[${i + 1}/${typeIds.length}] objectTypeId=${typeId}...`, "warning");

    const assets = await fetchByTypeId(typeId);
    let added    = 0;
    assets.forEach(a => {
      const k = a.key || a.id;
      if (k && seenKeys.has(k)) return;
      if (k) seenKeys.add(k);
      allAssets.push(a);
      added++;
    });
    console.log(`typeId=${typeId}: returned=${assets.length}, added=${added}, total=${allAssets.length}`);
  }

  console.log(`fetchJiraAssets: ${allAssets.length} unique assets`);
  return allAssets;
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

// Ghi header nếu sheet trống (row 0 chưa có "Asset ID")
async function ensureHeaders(context, sheet) {
  const r = sheet.getRangeByIndexes(0, 0, 1, COL_COUNT);
  r.load("values");
  await context.sync();
  if (r.values[0][0] !== "Asset ID") {
    r.values                 = [HEADERS];
    r.format.font.bold       = true;
    r.format.fill.color      = "#1a1d27";
    r.format.font.color      = "#8892a4";
    sheet.freezePanes.freezeRows(1);
    await context.sync();
  }
}

// Đọc tất cả data rows (bỏ qua header row 0)
async function readSheetRows(context, sheet) {
  const used = sheet.getUsedRangeOrNullObject(true);
  await context.sync();
  if (used.isNullObject) return [];
  used.load(["values", "rowCount"]);
  await context.sync();
  if (used.rowCount <= 1) return []; // chỉ có header
  return used.values.slice(1);       // bỏ header
}

// Lấy tất cả sheet có tên bắt đầu bằng "_"
async function getLocationSheets(context) {
  context.workbook.worksheets.load("items/name");
  await context.sync();
  return context.workbook.worksheets.items
    .filter(s => s.name.startsWith("_"))
    .map(s => s.name);
}

// ══════════════════════════════════════════════════════════════
// ROW BUILDER
//
// Tạo 1 row array từ asset.
// existingRow: nếu là UPDATE, truyền vào để giữ user-cols (DAYS, NOTE, ACTION, CASE_JIRA, VALIDATION)
// Nếu INSERT mới, existingRow = null → fill rỗng cho user-cols
// ══════════════════════════════════════════════════════════════
function assetToRow(asset, now, existingRow) {
  // Bắt đầu từ existingRow (giữ user-cols) hoặc array rỗng
  const row = existingRow ? [...existingRow] : Array(COL_COUNT).fill("");

  // Ghi tất cả Jira cols
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

  // user-cols (DAYS=25, NOTE=26, CASE_JIRA=27, VALIDATION=28, ACTION=31)
  // đã được giữ nguyên từ existingRow nếu là UPDATE
  // nếu INSERT mới → "" (đã fill ở Array(COL_COUNT).fill(""))

  return row;
}

// ══════════════════════════════════════════════════════════════
// WRITE LOCATION SHEET
//
// Logic:
//   A. Đọc toàn bộ rows hiện có, index theo Asset ID và Serial
//   B. Với mỗi asset từ Jira:
//      - Tìm row tương ứng (by ID → by Serial)
//      - Nếu tìm thấy → UPDATE toàn bộ Jira cols (giữ user-cols)
//      - Nếu không tìm thấy → INSERT cuối sheet
//   C. Rows hiện có không được match → đánh dấu "Not in Jira"
//      (chỉ với SYNC_STATUS=JIRA, không đụng vào LOCAL rows)
// ══════════════════════════════════════════════════════════════
async function writeLocationSheet(sheetName, assets, now) {
  await Excel.run(async (context) => {
    const sheet = await ensureSheet(context, sheetName);
    await ensureHeaders(context, sheet);

    // Đọc rows hiện có
    const existing = await readSheetRows(context, sheet);
    // Index: assetId → rowIndex (0-based trong existing array, tức Excel row = idx+1+1=idx+2)
    const byId     = {};
    const bySerial = {};
    existing.forEach((row, idx) => {
      const id  = String(row[COL.ASSET_ID] || "").trim();
      const ser = String(row[COL.SERIAL]   || "").trim();
      if (id)  byId[id]      = idx;
      if (ser) bySerial[ser] = idx;
    });

    const matchedIdxs = new Set(); // existing rows đã được match với Jira asset
    const toInsert    = [];        // assets mới chưa có trong sheet

    // ── B: UPDATE hoặc queue INSERT ─────────────────────────
    for (const asset of assets) {
      // Tìm row: ưu tiên match by Asset ID, fallback by Serial
      let idx = -1;
      if (asset.id && byId[asset.id] !== undefined) {
        idx = byId[asset.id];
      } else if (asset.serial && bySerial[asset.serial] !== undefined) {
        idx = bySerial[asset.serial];
      }

      if (idx >= 0) {
        // UPDATE: ghi toàn bộ Jira cols, giữ nguyên user-cols
        const newRow = assetToRow(asset, now, existing[idx]);
        // Xoá "Not in Jira" nếu asset đã quay lại Jira
        if (String(existing[idx][COL.VALIDATION] || "") === "Not in Jira") {
          newRow[COL.VALIDATION] = "";
        }
        // Ghi toàn bộ row (1 lần gọi API thay vì nhiều lần)
        sheet.getRangeByIndexes(idx + 1, 0, 1, COL_COUNT).values = [newRow];
        matchedIdxs.add(idx);
      } else {
        toInsert.push(asset);
      }
    }

    // ── C: Mark rows không còn trong Jira ───────────────────
    const jiraIdSet = new Set(assets.map(a => a.id).filter(Boolean));
    existing.forEach((row, idx) => {
      if (matchedIdxs.has(idx))              return; // đã match
      if (row[COL.SYNC_STATUS] === "LOCAL")  return; // không đụng LOCAL
      const id = String(row[COL.ASSET_ID] || "").trim();
      if (!id || jiraIdSet.has(id))          return; // chưa có ID hoặc vẫn còn trong Jira
      const cell = sheet.getRangeByIndexes(idx + 1, COL.VALIDATION, 1, 1);
      cell.values            = [["Not in Jira"]];
      cell.format.font.color = "#f59e0b";
    });

    // ── INSERT rows mới — batch 1 lần ───────────────────────
    if (toInsert.length > 0) {
      // Tìm row tiếp theo sau dữ liệu hiện có
      // getUsedRangeOrNullObject trả rowCount tính cả header
      const used = sheet.getUsedRangeOrNullObject(true);
      await context.sync();
      let nextRow = 1; // mặc định: ngay sau header (row index 1)
      if (!used.isNullObject) {
        used.load("rowCount");
        await context.sync();
        nextRow = used.rowCount; // rowCount = số row đang dùng (bao gồm header)
        // nên nextRow là index của row trống đầu tiên (0-based)
      }
      // Ghi tất cả rows mới trong 1 lần
      sheet.getRangeByIndexes(nextRow, 0, toInsert.length, COL_COUNT).values =
        toInsert.map(a => assetToRow(a, now, null));
    }

    await context.sync();
    console.log(`[writeLocationSheet] ${sheetName}: updated=${matchedIdxs.size}, inserted=${toInsert.length}, notInJira=${existing.length - matchedIdxs.size}`);
  });
}

// ══════════════════════════════════════════════════════════════
// CREATE LOCATION SHEETS
// ══════════════════════════════════════════════════════════════
async function createLocationSheets() {
  if (!cfg.jiraUrl || !cfg.token) { toast("Configure Jira settings first", "warning"); return; }
  toast("Đang tải assets từ Jira...", "warning");
  try {
    const assets    = await fetchJiraAssets();
    const locations = [...new Set(assets.map(a => a.location).filter(Boolean))];
    if (!locations.length) { toast("Không tìm thấy Location nào", "warning"); return; }
    await Excel.run(async (context) => {
      for (const loc of locations) {
        const sheet = await ensureSheet(context, locationSheetName(loc));
        await ensureHeaders(context, sheet);
      }
    });
    toast(`Đã tạo ${locations.length} sheet`, "success");
    await refreshDashboard();
  } catch(e) { toast("Error: " + e.message, "error"); }
}

// ══════════════════════════════════════════════════════════════
// SYNC ENGINE
//
// Luồng:
//   1. Với mỗi typeId: fetchByTypeId → assets[]
//   2. Group assets theo location sheet
//   3. writeLocationSheet cho mỗi sheet (UPDATE + INSERT)
//   4. Sau tất cả typeId: pushLocalAssets
//
// Lưu ý quan trọng:
//   - KHÔNG reset locationMap giữa các typeId
//     → typeId sau có thể ghi vào cùng sheet với typeId trước
//   - writeLocationSheet được gọi sau KHI đã tích lũy TẤT CẢ typeId
//     → tránh ghi đè lẫn nhau
// ══════════════════════════════════════════════════════════════
async function runSync() {
  if (isSyncing) { toast("Sync đang chạy", "warning"); return; }
  if (!cfg.jiraUrl || !cfg.token || !cfg.workerUrl) {
    toast("Kiểm tra Settings (URL, token, worker)", "warning"); return;
  }

  isSyncing = true;
  setSyncIndicator("syncing", "Syncing...");
  const syncPanel = document.getElementById("sync-location-list");
  if (syncPanel) syncPanel.innerHTML =
    `<div class="empty-state"><div class="spinner"></div>Đang tải từ Jira...</div>`;

  try {
    const typeIds = parseTypeIds();
    if (!typeIds.length) { toast("AQL Query chưa đúng", "error"); return; }

    const now = new Date().toISOString();

    // locationMap tích lũy qua tất cả typeId: sheetName → Map<key, asset>
    // KHÔNG reset giữa các typeId
    const locationMap = {};
    let totalFetched  = 0;
    const seenKeys    = new Set();

    // ── Bước 1: fetch tất cả typeId ─────────────────────────
    for (let i = 0; i < typeIds.length; i++) {
      const typeId = typeIds[i];
      toast(`[${i + 1}/${typeIds.length}] Đang tải objectTypeId=${typeId}...`, "warning");

      const assets = await fetchByTypeId(typeId);
      let added    = 0;

      assets.forEach(a => {
        const k = a.key || a.id;
        if (k && seenKeys.has(k)) return; // dedup global
        if (k) seenKeys.add(k);

        added++;
        totalFetched++;

        const sheetName = locationSheetName((a.location || "UNKNOWN").trim());
        if (!locationMap[sheetName]) locationMap[sheetName] = new Map();
        locationMap[sheetName].set(k || `__nokey_${totalFetched}`, a);
      });

      console.log(`typeId=${typeId}: fetched=${assets.length}, added=${added}, total=${totalFetched}`);
    }

    // ── Bước 2: ghi tất cả location sheets ──────────────────
    const sheetNames = Object.keys(locationMap);
    toast(`Đang ghi ${sheetNames.length} sheet(s)...`, "warning");

    for (const sheetName of sheetNames) {
      const sheetAssets = Array.from(locationMap[sheetName].values());
      toast(`Ghi sheet ${sheetName} (${sheetAssets.length} assets)...`, "warning");
      await writeLocationSheet(sheetName, sheetAssets, now);
    }

    // Update UI panel
    if (syncPanel) {
      updateSyncPanel(syncPanel, sheetNames.map(s => ({
        name:   s.replace(/^_/, ""),
        count:  locationMap[s].size,
        status: "done",
      })));
    }

    // ── Bước 3: push LOCAL rows lên Jira ─────────────────────
    toast("Đang push LOCAL assets lên Jira...", "warning");
    await pushLocalAssets();

    // ── Lưu lastSync ─────────────────────────────────────────
    cfg.lastSync = new Date().toISOString();
    Office.context.document.settings.set(CFG_KEYS.LAST_SYNC, cfg.lastSync);
    Office.context.document.settings.saveAsync();

    setSyncIndicator("ok", "Synced");
    setInner("last-sync-time", formatTime(cfg.lastSync));
    toast(`Sync hoàn tất — ${totalFetched} assets, ${sheetNames.length} location(s)`, "success");
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
// PUSH LOCAL ASSETS LÊN JIRA
//
// Tìm tất cả rows có SYNC_STATUS=LOCAL và chưa có ASSET_ID.
// Tạo object mới trên Jira Assets, ghi lại ID và KEY vào Excel.
// ══════════════════════════════════════════════════════════════
async function pushLocalAssets() {
  if (!cfg.cloudId || !cfg.workspaceId) return;
  const typeIds = parseTypeIds();
  if (!typeIds.length) return;
  const defaultTypeId = typeIds[0];

  let pushed = 0, failed = 0;
  try {
    await Excel.run(async (context) => {
      const sheetNames = await getLocationSheets(context);

      for (const sheetName of sheetNames) {
        const sheet    = context.workbook.worksheets.getItem(sheetName);
        const existing = await readSheetRows(context, sheet);

        for (let i = 0; i < existing.length; i++) {
          const row = existing[i];
          if (row[COL.SYNC_STATUS] !== "LOCAL") continue;
          if (row[COL.ASSET_ID])                continue; // đã push rồi

          const hostname = String(row[COL.HOSTNAME] || "").trim();
          const serial   = String(row[COL.SERIAL]   || "").trim();
          if (!hostname && !serial) continue;

          try {
            const attrs = [
              { objectTypeAttributeId: 1737,  objectAttributeValues: [{ value: hostname }] },
              { objectTypeAttributeId: 5194,  objectAttributeValues: [{ value: serial }] },
              { objectTypeAttributeId: 30125, objectAttributeValues: [{ value: String(row[COL.LOCATION] || "") }] },
              { objectTypeAttributeId: 5200,  objectAttributeValues: [{ value: String(row[COL.USERNAME] || "") }] },
              { objectTypeAttributeId: 5052,  objectAttributeValues: [{ value: String(row[COL.STATUS] || "") }] },
            ].filter(a => a.objectAttributeValues[0].value);

            const res = await assetsPost("/object/create", {
              objectTypeId: defaultTypeId,
              attributes:   attrs,
            });

            if (res?.id) {
              // Đọc lại row hiện tại trước khi ghi (tránh race condition)
              const cell = sheet.getRangeByIndexes(i + 1, 0, 1, COL_COUNT);
              cell.load("values");
              await context.sync();
              const cur             = cell.values[0];
              cur[COL.ASSET_ID]    = String(res.id);
              cur[COL.ASSET_KEY]   = res.objectKey || "";
              cur[COL.SYNC_STATUS] = "JIRA";
              cur[COL.LAST_SYNC]   = new Date().toISOString();
              cell.values          = [cur];
              await context.sync();
              pushed++;
            }
          } catch(err) {
            failed++;
            console.warn(`[pushLocal] row ${i + 2} in ${sheetName}: ${err.message}`);
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

// ══════════════════════════════════════════════════════════════
// UPDATE JIRA ASSET từ Excel row
//
// Gọi khi user chỉnh sửa 1 hoặc nhiều rows và muốn push lên Jira.
// Đọc row(s) có SYNC_STATUS=JIRA và ASSET_ID, gọi PUT /object/{id}
// ══════════════════════════════════════════════════════════════
async function updateJiraAsset(assetId, row) {
  const attrs = [
    { objectTypeAttributeId: 1737,  objectAttributeValues: [{ value: String(row[COL.HOSTNAME]  || "") }] },
    { objectTypeAttributeId: 5194,  objectAttributeValues: [{ value: String(row[COL.SERIAL]    || "") }] },
    { objectTypeAttributeId: 30125, objectAttributeValues: [{ value: String(row[COL.LOCATION]  || "") }] },
    { objectTypeAttributeId: 5200,  objectAttributeValues: [{ value: String(row[COL.USERNAME]  || "") }] },
    { objectTypeAttributeId: 5052,  objectAttributeValues: [{ value: String(row[COL.STATUS]    || "") }] },
    { objectTypeAttributeId: 26690, objectAttributeValues: [{ value: String(row[COL.ASSIGNED]  || "") }] },
  ].filter(a => a.objectAttributeValues[0].value);

  return assetsPut(`/object/${assetId}`, { attributes: attrs });
}

// Push các rows đã chỉnh sửa (SYNC_STATUS=JIRA, có ASSET_ID) lên Jira
async function pushUpdatedRows() {
  if (!cfg.cloudId || !cfg.workspaceId) { toast("Chưa cấu hình Cloud ID / Workspace ID", "warning"); return; }
  let updated = 0, failed = 0;
  try {
    await Excel.run(async (context) => {
      const sheetNames = await getLocationSheets(context);
      for (const sheetName of sheetNames) {
        const sheet    = context.workbook.worksheets.getItem(sheetName);
        const existing = await readSheetRows(context, sheet);
        for (let i = 0; i < existing.length; i++) {
          const row     = existing[i];
          const assetId = String(row[COL.ASSET_ID] || "").trim();
          if (row[COL.SYNC_STATUS] !== "JIRA") continue;
          if (!assetId)                         continue;
          // Chỉ update những row có ACTION = "Update Jira"
          if (row[COL.ACTION] !== "Update Jira") continue;
          try {
            await updateJiraAsset(assetId, row);
            const cell   = sheet.getRangeByIndexes(i + 1, COL.ACTION, 1, 2);
            const cur    = cell.values[0] || ["", ""];
            cur[0]       = ""; // xoá ACTION
            cur[1]       = new Date().toISOString(); // LAST_SYNC
            cell.values  = [cur];
            await context.sync();
            updated++;
          } catch(err) {
            failed++;
            console.warn(`[pushUpdated] row ${i + 2}: ${err.message}`);
          }
        }
      }
    });
    toast(`Update Jira: ${updated} thành công${failed ? ", " + failed + " lỗi" : ""}`,
      failed ? "warning" : "success");
  } catch(e) {
    toast("Update error: " + e.message, "error");
  }
}

// ══════════════════════════════════════════════════════════════
// DASHBOARD
// ══════════════════════════════════════════════════════════════
async function refreshDashboard() {
  try {
    await Excel.run(async (context) => {
      const sheetNames = await getLocationSheets(context);
      let total = 0, pending = 0, mismatch = 0, local = 0;
      const locSummary = [];

      for (const name of sheetNames) {
        const sheet = context.workbook.worksheets.getItem(name);
        const rows  = await readSheetRows(context, sheet);
        total += rows.length;
        let locLocal = 0;
        rows.forEach(r => {
          if (r[COL.SYNC_STATUS] === "LOCAL")                         { local++; locLocal++; }
          if (r[COL.VALIDATION]  && r[COL.VALIDATION] !== "OK")         mismatch++;
          if (r[COL.ACTION] === "Create Ticket" && !r[COL.CASE_JIRA])   pending++;
        });
        locSummary.push({ name, count: rows.length, local: locLocal });
      }

      setInner("stat-total",    total    || "0");
      setInner("stat-pending",  pending  || "0");
      setInner("stat-mismatch", mismatch || "0");
      setInner("stat-local",    local    || "0");
      if (cfg.lastSync) setInner("last-sync-time", formatTime(cfg.lastSync));

      const el = document.getElementById("location-summary");
      if (!el) return;
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
// MATCH LOCAL ASSETS với Jira (by Serial)
// ══════════════════════════════════════════════════════════════
async function matchLocalAssets() {
  if (!cfg.jiraUrl || !cfg.token) { toast("Configure Jira settings first", "warning"); return; }
  toast("Đang tìm kiếm match từ Jira...", "warning");
  try {
    const jiraAssets   = await fetchJiraAssets();
    const bySerial     = {};
    jiraAssets.forEach(a => { if (a.serial) bySerial[a.serial.trim()] = a; });

    let matched = 0;
    await Excel.run(async (context) => {
      const sheetNames = await getLocationSheets(context);
      for (const sheetName of sheetNames) {
        const sheet    = context.workbook.worksheets.getItem(sheetName);
        const existing = await readSheetRows(context, sheet);
        for (let i = 0; i < existing.length; i++) {
          const row    = existing[i];
          if (row[COL.SYNC_STATUS] !== "LOCAL") continue;
          const serial = String(row[COL.SERIAL] || "").trim();
          if (!serial || !bySerial[serial])       continue;

          const asset = bySerial[serial];
          const cell  = sheet.getRangeByIndexes(i + 1, 0, 1, COL_COUNT);
          cell.load("values");
          await context.sync();
          const cur             = cell.values[0];
          cur[COL.ASSET_ID]    = asset.id;
          cur[COL.ASSET_KEY]   = asset.key;
          cur[COL.HOSTNAME]    = asset.hostname;
          cur[COL.STATUS]      = asset.status;
          cur[COL.LOCATION]    = asset.location;
          cur[COL.REGION]      = asset.region;
          cur[COL.MODEL]       = asset.model;
          cur[COL.OS]          = asset.os;
          cur[COL.OS_VERSION]  = asset.osVersion;
          cur[COL.OS_BUILD]    = asset.osBuild;
          cur[COL.SYNC_STATUS] = "JIRA";
          cur[COL.LAST_SYNC]   = new Date().toISOString();
          cell.values          = [cur];
          await context.sync();
          matched++;
        }
      }
    });

    toast(`Matched ${matched} LOCAL asset(s)`, "success");
    await refreshDashboard();
  } catch(e) { toast("Match error: " + e.message, "error"); }
}

// ══════════════════════════════════════════════════════════════
// VALIDATION
// ══════════════════════════════════════════════════════════════
async function runValidation() {
  toast("Running validation…", "warning");
  const errors = {
    "Missing Asset ID": [],
    "Duplicate Serial": [],
    "Location Changed": [],
    "Serial Mismatch":  [],
    "Owner Mismatch":   [],
  };

  try {
    await Excel.run(async (context) => {
      const sheetNames = await getLocationSheets(context);
      const serialSeen = {};

      for (const sheetName of sheetNames) {
        const sheet    = context.workbook.worksheets.getItem(sheetName);
        const existing = await readSheetRows(context, sheet);

        for (let i = 0; i < existing.length; i++) {
          const row      = existing[i];
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
              if (!errors["Duplicate Serial"].find(
                e => e.sheet === serialSeen[serial].sheet && e.row === serialSeen[serial].row
              )) {
                errors["Duplicate Serial"].push(serialSeen[serial]);
              }
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
  } catch(e) { toast("Validation error: " + e.message, "error"); }
}

function renderValidationList(errors) {
  const el = document.getElementById("val-list");
  if (!el) return;
  el.innerHTML = Object.entries(errors).map(([name, list]) => {
    const unique = [...new Map(list.map(e => [`${e.sheet}:${e.row}`, e])).values()];
    return `
      <div class="val-item" data-errors='${JSON.stringify(unique)}' onclick="jumpToError(this)">
        <span class="val-dot ${unique.length > 0 ? "err" : "ok"}"></span>
        <span class="val-name">${name}</span>
        <span class="val-count ${unique.length > 0 ? "has-err" : ""}">${unique.length}</span>
      </div>`;
  }).join("") || `<div class="empty-state"><div class="icon">✓</div>No issues</div>`;
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
      const sheetNames = await getLocationSheets(context);
      for (const name of sheetNames) {
        const rows = await readSheetRows(context, context.workbook.worksheets.getItem(name));
        rows.forEach(r => { if (r[COL.ACTION] === "Create Ticket" && !r[COL.CASE_JIRA]) count++; });
      }
    });
  } catch(e) { console.warn("scanPending:", e.message); }
  setInner("selected-count", count);
}

async function createTickets() {
  const issueType = document.getElementById("issue-type")?.value    || "Task";
  const priority  = document.getElementById("issue-priority")?.value || "Medium";
  const days      = parseInt(document.getElementById("default-days")?.value) || 30;

  if (!cfg.jiraUrl || !cfg.token) { toast("Configure Jira settings first", "warning"); return; }
  toast("Processing ticket queue…", "warning");
  let created = 0, skipped = 0, failed = 0;

  try {
    await Excel.run(async (context) => {
      const sheetNames = await getLocationSheets(context);
      for (const sheetName of sheetNames) {
        const sheet    = context.workbook.worksheets.getItem(sheetName);
        const existing = await readSheetRows(context, sheet);

        for (let i = 0; i < existing.length; i++) {
          const row = existing[i];
          if (row[COL.ACTION] !== "Create Ticket") continue;
          if (row[COL.CASE_JIRA]) { skipped++; continue; }

          const device  = row[COL.HOSTNAME] || row[COL.ASSET_KEY] || "Unknown";
          const serial  = row[COL.SERIAL]   || "";
          const email   = row[COL.ASSIGNED] || row[COL.USERNAME] || "";
          const note    = row[COL.NOTE]     || "";
          const rowDays = row[COL.DAYS]     || days;

          try {
            const res = await jiraPost("/api/3/issue", {
              fields: {
                project:     { key: cfg.projectKey },
                summary:     `[${issueType}] ${device}${serial ? " – " + serial : ""}`,
                description: {
                  type: "doc", version: 1,
                  content: [{ type: "paragraph", content: [{ type: "text",
                    text: `Asset: ${device}\nSerial: ${serial}\nUser: ${email}\nDays: ${rowDays}\nNote: ${note}`
                  }]}],
                },
                issuetype: { name: issueType },
                priority:  { name: priority },
              },
            });

            if (res.key) {
              const url = `${jiraBase()}/browse/${res.key}`;
              sheet.getRangeByIndexes(i + 1, COL.CASE_JIRA, 1, 1).values =
                [[`=HYPERLINK("${url}","${res.key}")`]];
              sheet.getRangeByIndexes(i + 1, COL.ACTION, 1, 1).values = [[""]];
              await context.sync();
              created++;
            }
          } catch(err) { failed++; toast(`Row ${i + 2}: ${err.message}`, "error"); }
        }
      }
    });

    toast(`Done — created:${created}, skipped:${skipped}${failed ? ", failed:"+failed : ""}`,
      failed ? "warning" : "success");
    scanPendingRows();
  } catch(e) { toast("Ticket error: " + e.message, "error"); }
}

// ══════════════════════════════════════════════════════════════
// CONNECTION TEST
// ══════════════════════════════════════════════════════════════
async function testConnection() {
  const el = document.getElementById("conn-test-result");
  if (!el) return;
  el.style.display     = "block";
  el.style.borderColor = "var(--border)";
  el.innerHTML = `<div style="display:flex;gap:8px;align-items:center"><div class="spinner"></div><span>Testing…</span></div>`;

  saveConfig();

  if (!cfg.workerUrl) {
    el.innerHTML         = "✗ Worker URL chưa được điền";
    el.style.borderColor = "var(--red)";
    toast("Điền Cloudflare Worker URL trước", "error");
    return;
  }

  try {
    const me = await jiraGet("/api/3/myself");
    let assetsMsg = "";
    try {
      await assetsPost("/object/aql", { qlQuery: "objectType != null", startAt: 0, maxResults: 1 });
      assetsMsg = " · Assets API ✓";
    } catch(ae) {
      assetsMsg = ` · Assets API ✗ (${ae.message.slice(0, 60)})`;
    }
    el.innerHTML         = `✓ Connected as <strong>${me.displayName || me.emailAddress || "?"}</strong>${assetsMsg}`;
    el.style.borderColor = "var(--green)";
    toast("Connection successful", "success");
  } catch(e) {
    el.innerHTML         = `✗ ${e.message}`;
    el.style.borderColor = "var(--red)";
    toast("Connection failed", "error");
  }
}

// ══════════════════════════════════════════════════════════════
// SYNC INDICATOR & PANEL
// ══════════════════════════════════════════════════════════════
function setSyncIndicator(state, text) {
  const dot = document.getElementById("sync-dot");
  const txt = document.getElementById("sync-status-text");
  if (dot) dot.className = "sync-dot" + (state === "syncing" ? " syncing" : "");
  if (txt) txt.textContent = text;
}

function updateSyncPanel(el, state) {
  el.innerHTML = state.map(l => `
    <div class="location-item">
      <div class="loc-header">
        <span class="loc-name">_${l.name}</span>
        <span class="loc-status ${l.status}">
          ${l.status === "done" ? "Done" : l.status === "running" ? "Running…" : "Waiting"}
        </span>
      </div>
      <div class="loc-count">${l.count} assets</div>
    </div>`).join("");
}

// ══════════════════════════════════════════════════════════════
// OPEN JIRA
// ══════════════════════════════════════════════════════════════
function openJira() {
  if (!cfg.jiraUrl) { toast("Set Jira URL in Settings first", "warning"); return; }
  window.open(cfg.jiraUrl, "_blank");
}

// ══════════════════════════════════════════════════════════════
// TOAST
// ══════════════════════════════════════════════════════════════
function toast(msg, type = "success") {
  const icons     = { success: "✓", error: "✗", warning: "⚠" };
  const container = document.getElementById("toast-container");
  if (!container) return;
  const el      = document.createElement("div");
  el.className  = `toast ${type}`;
  el.innerHTML  = `<span>${icons[type] || "ℹ"}</span><span>${msg}</span>`;
  container.appendChild(el);
  setTimeout(() => el.remove(), 4500);
}

// ══════════════════════════════════════════════════════════════
// UTILS
// ══════════════════════════════════════════════════════════════
function setInner(id, val) { const e = document.getElementById(id); if (e) e.textContent = val; }
function setVal(id, val)   { const e = document.getElementById(id); if (e) e.value = val || ""; }
function getVal(id)        { const e = document.getElementById(id); return e ? e.value.trim() : ""; }
function formatTime(iso) {
  if (!iso) return "—";
  try {
    return new Date(iso).toLocaleString([], { month:"2-digit", day:"2-digit", hour:"2-digit", minute:"2-digit" });
  } catch { return iso; }
}
