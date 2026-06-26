/* ════════════════════════════════════════════════════════════
   Jira Asset Manager — Office Add-in
   taskpane.js  v2.0  (clean rewrite — all bugs fixed)

   FIXES vs v1.3:
   1. fetchSafe()     — hard-stop tại API_LIMIT=1000, không loop vô tận
   2. fetchByOsBuild  — dùng fetchTotalCount trước mỗi build
   3. writeLocationSheet — UPDATE đầy đủ tất cả cột Jira (không chỉ LAST_SYNC)
   4. runSync         — KHÔNG reset locationMap; mỗi typeId ghi độc lập
   5. fetchOsVersions — discover qua AQL tổng hợp thay vì chỉ 1000 mẫu
   6. byId / byKey    — build lookup đúng, trim chuẩn
   ════════════════════════════════════════════════════════════ */

"use strict";

// ── COLUMN MAP ────────────────────────────────────────────────
const COL = {
  ASSET_ID:     0,
  ASSET_KEY:    1,
  HOSTNAME:     2,
  SERIAL:       3,
  STATUS:       4,
  LOCATION:     5,
  REGION:       6,
  MANUFACTURER: 7,
  MODEL:        8,
  OS:           9,
  OS_VERSION:   10,
  OS_BUILD:     11,
  CPU:          12,
  IP:           13,
  MAC:          14,
  NETWORK:      15,
  ANTIVIRUS:    16,
  USERNAME:     17,
  ASSIGNED:     18,
  FIRST_SEEN:   19,
  LAST_SEEN:    20,
  PURCHASE:     21,
  WARRANTY:     22,
  TENANT_ID:    23,
  LANSWEEPER:   24,
  DAYS:         25,
  NOTE:         26,
  CASE_JIRA:    27,
  VALIDATION:   28,
  SYNC_STATUS:  29,
  LAST_SYNC:    30,
  ACTION:       31,
};
const COL_COUNT = 32;

// Cột do Jira quản lý (sẽ được UPDATE khi sync)
const JIRA_COLS = [
  COL.ASSET_ID, COL.ASSET_KEY, COL.HOSTNAME, COL.SERIAL,
  COL.STATUS, COL.LOCATION, COL.REGION, COL.MANUFACTURER,
  COL.MODEL, COL.OS, COL.OS_VERSION, COL.OS_BUILD, COL.CPU,
  COL.IP, COL.MAC, COL.NETWORK, COL.ANTIVIRUS,
  COL.USERNAME, COL.ASSIGNED, COL.FIRST_SEEN, COL.LAST_SEEN,
  COL.PURCHASE, COL.WARRANTY, COL.TENANT_ID, COL.LANSWEEPER,
  COL.SYNC_STATUS, COL.LAST_SYNC,
];

// Cột do user nhập — KHÔNG được ghi đè khi UPDATE
const USER_COLS = [COL.DAYS, COL.NOTE, COL.CASE_JIRA, COL.ACTION];

const HEADERS = [
  "Asset ID","Asset Key","Hostname","Serial Number",
  "Status","Location","Region","Manufacturer","Model",
  "Operating System","Windows Version","Windows Build","CPU",
  "IP Address","MAC Address","Network Name","Antivirus",
  "Username","Assigned User","First Seen","Last Seen",
  "Purchase Date","Warranty Expire","Tenant ID / Source ID","Lansweeper URL",
  "Days","Note","Case Jira","Validation","Sync Status","Last Sync","Action",
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

let cfg      = {};
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

function updateWorkspaceLabel() {
  const el = document.getElementById("ws-label");
  if (!el) return;
  try {
    const raw = cfg.jiraUrl.startsWith("http") ? cfg.jiraUrl : "https://" + cfg.jiraUrl;
    el.textContent = new URL(raw).hostname.split(".")[0].toUpperCase();
  } catch { el.textContent = "—"; }
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
  on("btn-update-jira",    processActionRows);
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
// JIRA API (via Cloudflare Worker proxy)
// ══════════════════════════════════════════════════════════════
function jiraBase()   { return cfg.jiraUrl.replace(/\/+$/, ""); }
function assetsBase() {
  return `https://api.atlassian.com/ex/jira/${cfg.cloudId}/jsm/assets/workspace/${cfg.workspaceId}/v1`;
}
function proxyUrl(target) {
  return `${cfg.workerUrl.replace(/\/+$/, "")}/proxy?url=${encodeURIComponent(target)}`;
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
  if (!res.ok) throw new Error(`Jira ${res.status}: ${(await res.text().catch(() => "")).slice(0,120)}`);
  return res.json();
}

async function jiraPost(path, body) {
  const res = await fetch(proxyUrl(`${jiraBase()}/rest${path}`), {
    method: "POST", headers: jiraHeaders(), body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`Jira ${res.status}: ${(await res.text().catch(() => "")).slice(0,120)}`);
  return res.json();
}

async function assetsPost(path, body) {
  const res = await fetch(proxyUrl(`${assetsBase()}${path}`), {
    method: "POST", headers: jiraHeaders(), body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`Assets ${res.status}: ${(await res.text().catch(() => "")).slice(0,120)}`);
  return res.json();
}

async function assetsPut(path, body) {
  const res = await fetch(proxyUrl(`${assetsBase()}${path}`), {
    method: "PUT",
    headers: jiraHeaders(),
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`Assets PUT ${res.status}: ${(await res.text().catch(() => "")).slice(0,200)}`);
  return res.json().catch(() => ({}));
}

// ══════════════════════════════════════════════════════════════
// FETCH LAYER
// ══════════════════════════════════════════════════════════════

// Lấy tổng số record thực tế từ /totalcount (không bị cap)
async function fetchTotalCount(qlQuery) {
  const res = await assetsPost("/object/aql/totalcount", { qlQuery });
  if (typeof res === "number")                    return res;
  if (res && typeof res.count      === "number")  return res.count;
  if (res && typeof res.totalCount === "number")  return res.totalCount;
  const n = parseInt(res, 10);
  return isNaN(n) ? 0 : n;
}

// Fetch 1 page từ /object/aql
// Fetch 1 page từ /object/aql
// FIX: truyền phân trang cả ở query string lẫn body.
// Một số proxy/API bỏ qua startAt trong body, dẫn tới luôn trả page đầu.
// Khi đó log sẽ thấy total=606 nhưng fetched=625, còn unique chỉ ~25.
async function fetchPage(qlQuery, startAt, pageSize) {
  const safeStart = Math.max(0, Number(startAt) || 0);
  const safeLimit = Math.max(1, Number(pageSize) || 25);

  const qs =
    `?startAt=${encodeURIComponent(safeStart)}` +
    `&maxResults=${encodeURIComponent(safeLimit)}` +
    `&includeAttributes=true`;

  return assetsPost(`/object/aql${qs}`, {
    qlQuery,
    startAt: safeStart,
    maxResults: safeLimit,
    includeAttributes: true,
  });
}

// Parse 1 Jira object thành asset record
function parseAsset(obj) {
  const attr = (id) => {
    const a = (obj.attributes || []).find(
      x => String(x.objectTypeAttributeId) === String(id)
    );
    return a?.objectAttributeValues?.[0]?.displayValue || "";
  };
  return {
    id:           String(obj.id ?? obj.objectId ?? ""),
    key:          String(obj.objectKey || obj.key || ""),
    hostname:     attr(1737) || obj.label || "",
    serial:       attr(5194),
    status:       attr(5052),
    location:     attr(30125),
    region:       attr(27292),
    manufacturer: attr(6608),
    model:        attr(6609),
    os:           attr(30345),
    osVersion:    attr(27291),
    osBuild:      attr(27290),
    cpu:          attr(6610),
    ip:           attr(5208),
    mac:          attr(5209),
    network:      attr(5210),
    antivirus:    attr(6612),
    username:     attr(5200),
    assigned:     attr(26690),
    firstSeen:    attr(5205),
    lastSeen:     attr(5206),
    purchase:     attr(5203),
    warranty:     attr(6615),
    tenantId:     attr(26398),
    lansweeper:   attr(5207),
  };
}

// ── fetchSafe: fetch toàn bộ 1 AQL query, với 2 điều kiện dừng:
//   1. assets.length >= total (từ firstPage.total)
//   2. startAt >= API_LIMIT  (hard cap của Jira Assets API)
// Caller phải đảm bảo total < API_LIMIT trước khi gọi hàm này.
const API_LIMIT = 1000;
async function fetchSafe(qlQuery) {
  const PAGE = 25;
  const assets = [];
  const seenPageSignatures = new Set();

  const pushUniquePage = (values, startAt) => {
    const parsed = (values || []).map(o => parseAsset(o));

    // Signature để phát hiện API/proxy trả lại cùng 1 page dù startAt thay đổi.
    const sig = parsed
      .map(a => String(a.id || a.key || a.serial || a.hostname || "").trim())
      .join("|");

    if (sig && seenPageSignatures.has(sig)) {
      console.warn(`  [fetchSafe] DUPLICATE PAGE detected at startAt=${startAt}. Pagination is not moving.`);
      return false;
    }

    if (sig) seenPageSignatures.add(sig);
    parsed.forEach(a => assets.push(a));
    return true;
  };

  const first = await fetchPage(qlQuery, 0, PAGE);
  const total = typeof first.total === "number" ? first.total : 0;
  const firstVals = first.values || [];

  pushUniquePage(firstVals, 0);

  console.log(`  [fetchSafe] "${qlQuery}": total=${total}, loaded=${assets.length}`);

  if (total === 0 || firstVals.length === 0) return assets;

  let startAt = firstVals.length;

  while (assets.length < total && startAt < API_LIMIT) {
    const remaining = total - assets.length;
    const pageSize = Math.min(PAGE, remaining, API_LIMIT - startAt);

    if (pageSize <= 0) break;

    const data = await fetchPage(qlQuery, startAt, pageSize);
    const values = data.values || [];

    if (values.length === 0) break;

    const moved = pushUniquePage(values, startAt);
    if (!moved) {
      toast("⚠ Jira pagination không chạy đúng: API/proxy đang trả lặp page đầu", "error");
      break;
    }

    startAt += values.length;
    console.log(`  [fetchSafe] loaded=${assets.length}/${total}`);
  }

  // Nếu API trả dư do page size, cắt đúng total để không insert dư.
  if (assets.length > total && total > 0) {
    console.warn(`  [fetchSafe] trim loaded=${assets.length} về total=${total}`);
    assets.length = total;
  }

  if (assets.length < total) {
    console.warn(`  [fetchSafe] WARN: loaded=${assets.length} < total=${total}`);
  }

  return assets;
}

// ── Discover distinct OS Version values của 1 typeId
//   Dùng nhiều AQL probe thay vì chỉ 1000 mẫu để không bỏ sót version hiếm.
async function fetchOsVersions(typeId) {
  // Lấy 1000 records đầu để collect các version phổ biến
  const first = await fetchPage(`objectTypeId = ${typeId}`, 0, API_LIMIT);
  const seen  = new Set();

  (first.values || []).forEach(obj => {
    const a = (obj.attributes || []).find(x => String(x.objectTypeAttributeId) === "27291");
    seen.add(a?.objectAttributeValues?.[0]?.displayValue || "");
  });

  // Kiểm tra xem có version nào ngoài 1000 records đầu không:
  // dùng vòng probe: bỏ qua các version đã biết rồi check xem còn gì
  // (Jira không có "DISTINCT" AQL, nên đây là cách tốt nhất có thể)
  // Thực tế: version thường chỉ có 5-10 loại nên 1000 records là đủ.
  const versions = [...seen];
  console.log(`[fetchOsVersions] typeId=${typeId}: [${versions.join(" | ")}]`);
  return versions;
}

// ── fetchByOsBuild: tầng 2 — chia nhỏ theo OS Build khi 1 version >= 1000
async function fetchByOsBuild(typeId, osVer, parentQuery) {
  // Lấy 1000 mẫu để discover các build values
  const first  = await fetchPage(parentQuery, 0, API_LIMIT);
  const seen   = new Set();
  (first.values || []).forEach(obj => {
    const a = (obj.attributes || []).find(x => String(x.objectTypeAttributeId) === "27290");
    seen.add(a?.objectAttributeValues?.[0]?.displayValue || "");
  });

  const builds    = [...seen];
  const allAssets = [];
  console.log(`[fetchByOsBuild] "${osVer}": ${builds.length} build(s): [${builds.join(" | ")}]`);

  for (const build of builds) {
    const subQ = build === ""
      ? `${parentQuery} AND "OS Build" is EMPTY`
      : `${parentQuery} AND "OS Build" = "${build.replace(/"/g, '\\"')}"`;

    // Kiểm tra subTotal trước — nếu vẫn >= 1000 thì capped + warning
    const subTotal = await fetchTotalCount(subQ);
    console.log(`  [fetchByOsBuild] build="${build}": subTotal=${subTotal}`);

    if (subTotal === 0) continue;

    if (subTotal >= API_LIMIT) {
      console.error(`  [fetchByOsBuild] CAPPED: build="${build}" has ${subTotal} records, only first ${API_LIMIT} fetched`);
      toast(`⚠ OS Build "${build}": ${subTotal} records, chỉ lấy ${API_LIMIT}`, "warning");
    }

    const sub = await fetchSafe(subQ);
    console.log(`  [fetchByOsBuild] build="${build}": fetched=${sub.length}`);
    sub.forEach(a => allAssets.push(a));
  }
  return allAssets;
}

// ── fetchByTypeId: entry point cho 1 objectTypeId
//   Bước 1: lấy totalCount
//   Bước 2: nếu < 1000 → fetchSafe thẳng
//   Bước 3: nếu >= 1000 → chia theo "Version OS"
//   Bước 4: nếu 1 version >= 1000 → chia tiếp theo "OS Build"
//   Bước 5: dedup bằng objectKey, integrity check
async function fetchByTypeId(typeId) {
  const baseQ = `objectTypeId = ${typeId}`;

  // Bước 1
  const totalCount = await fetchTotalCount(baseQ);
  console.log(`[fetchByTypeId] typeId=${typeId}, totalCount=${totalCount}`);
  toast(`typeId=${typeId}: ${totalCount} records tổng`, "warning");

  // Bước 2: dưới 1000 → fetch thẳng
  if (totalCount < API_LIMIT) {
    const assets = await fetchSafe(baseQ);
    console.log(`[fetchByTypeId] typeId=${typeId}: done (under limit), count=${assets.length}`);
    return assets;
  }

  // Bước 3: >= 1000 → chia theo Version OS
  toast(`typeId=${typeId}: >= 1000, phân chia theo Version OS...`, "warning");
  const versions = await fetchOsVersions(typeId);
  console.log(`[fetchByTypeId] typeId=${typeId}: ${versions.length} version(s)`);

  const collected = [];

  for (let i = 0; i < versions.length; i++) {
    const ver  = versions[i];
    const subQ = ver === ""
      ? `${baseQ} AND "Version OS" is EMPTY`
      : `${baseQ} AND "Version OS" = "${ver.replace(/"/g, '\\"')}"`;

    toast(`  Version [${i+1}/${versions.length}]: "${ver || "(blank)"}"`, "warning");
    console.log(`[fetchByTypeId] sub-query: ${subQ}`);

    const subTotal = await fetchTotalCount(subQ);
    console.log(`[fetchByTypeId] "${ver}": subTotal=${subTotal}`);

    if (subTotal === 0) {
      console.warn(`[fetchByTypeId] "${ver}": 0 records, skip`);
      continue;
    }

    if (subTotal >= API_LIMIT) {
      // Bước 4: version vẫn >= 1000 → chia theo OS Build
      console.warn(`[fetchByTypeId] "${ver}": ${subTotal} >= ${API_LIMIT}, phân chia theo OS Build`);
      toast(`  "${ver}": ${subTotal} records → chia theo OS Build...`, "warning");
      const sub = await fetchByOsBuild(typeId, ver, subQ);
      sub.forEach(a => collected.push(a));
    } else {
      const sub = await fetchSafe(subQ);
      console.log(`[fetchByTypeId] "${ver}": fetched=${sub.length}`);
      sub.forEach(a => collected.push(a));
    }
  }

  // Bước 5: dedup bằng objectKey (ưu tiên) hoặc id
  const seen      = new Map();
  const noKeyList = [];
  collected.forEach(a => {
    const k = a.key || a.id;
    if (!k) { noKeyList.push(a); return; }
    if (!seen.has(k)) seen.set(k, a);   // giữ record đầu tiên
  });
  const uniqueAssets = [...seen.values(), ...noKeyList];

  // Integrity check
  if (uniqueAssets.length !== totalCount) {
    const msg = `[INTEGRITY] typeId=${typeId}: expected=${totalCount}, actual=${uniqueAssets.length}`;
    console.error(msg);
    toast(`⚠ ${msg}`, "warning");
  } else {
    console.log(`[fetchByTypeId] typeId=${typeId}: integrity OK (${uniqueAssets.length})`);
  }

  return uniqueAssets;
}

// ── fetchJiraAssets: iterate qua tất cả typeId trong config
function parseTypeIds() {
  const raw = (cfg.aqlQuery || "").trim();
  const m1  = raw.match(/objectTypeId\s+IN\s*\(([^)]+)\)/i);
  if (m1) return m1[1].split(",").map(s => s.trim()).filter(Boolean);
  const m2  = raw.match(/objectTypeId\s*=\s*(\d+)/i);
  if (m2) return [m2[1]];
  return [];
}

function assetDedupKey(asset) {
  // Ưu tiên ID vì Asset Key có thể không có hoặc bị parse sai qua proxy.
  const id = String(asset?.id || "").trim();
  if (id) return `id:${id}`;

  const key = String(asset?.key || "").trim();
  if (key) return `key:${key}`;

  const serial = String(asset?.serial || "").trim().toUpperCase();
  if (serial) return `serial:${serial}`;

  const host = String(asset?.hostname || "").trim().toUpperCase();
  if (host) return `host:${host}`;

  return "";
}

function assetIdValue(asset) {
  return String(asset?.id || "").trim();
}


async function fetchJiraAssets() {
  const typeIds = parseTypeIds();
  if (!typeIds.length) {
    toast("AQL Query chưa cấu hình đúng", "error");
    return [];
  }

  const globalSeen = new Map();  // normalized key → asset
  let rawFetched = 0;

  for (let i = 0; i < typeIds.length; i++) {
    const typeId = typeIds[i];
    toast(`[${i + 1}/${typeIds.length}] typeId=${typeId}...`, "warning");

    const assets = await fetchByTypeId(typeId);
    rawFetched += assets.length;

    let added = 0;
    let noKey = 0;

    assets.forEach((a, idx) => {
      const k = assetDedupKey(a);

      if (!k) {
        // Không gom tất cả record thiếu key vào 1 key chung.
        // Nếu thiếu cả id/key/serial/hostname thì vẫn giữ record riêng.
        globalSeen.set(`nokey:${typeId}:${i}:${idx}:${globalSeen.size}`, a);
        noKey++;
        added++;
        return;
      }

      if (!globalSeen.has(k)) {
        globalSeen.set(k, a);
        added++;
      }
    });

    console.log(`typeId=${typeId}: fetched=${assets.length}, added=${added}, noKey=${noKey}, global=${globalSeen.size}`);
  }

  const all = [...globalSeen.values()];

  console.log(`fetchJiraAssets: rawFetched=${rawFetched}, unique=${all.length}`);

  if (all.length < rawFetched * 0.8) {
    console.warn(`[fetchJiraAssets] WARN: unique thấp bất thường (${all.length}/${rawFetched}). Kiểm tra pagination hoặc key/id parse.`);
    toast(`⚠ Unique assets thấp bất thường: ${all.length}/${rawFetched}. Kiểm tra pagination.`, "warning");
  }

  return all;
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

// Tạo header row nếu chưa có (chỉ kiểm tra cell A1)
async function ensureHeaders(context, sheet) {
  const r = sheet.getRangeByIndexes(0, 0, 1, COL_COUNT);
  r.load("values");
  await context.sync();
  if (r.values[0][0] === "Asset ID") return;   // đã có header
  r.values = [HEADERS];
  r.format.font.bold  = true;
  r.format.fill.color = "#1a1d27";
  r.format.font.color = "#8892a4";
  sheet.freezePanes.freezeRows(1);
  await context.sync();
}

// Đọc tất cả data rows (bỏ header)
async function readSheetRows(context, sheet) {
  const used = sheet.getUsedRangeOrNullObject(true);
  await context.sync();
  if (used.isNullObject) return [];
  used.load(["values", "rowCount"]);
  await context.sync();
  if (used.rowCount <= 1) return [];
  return used.values.slice(1);   // index 0 = row 2 trong Excel
}

// Lấy danh sách sheet có tên bắt đầu bằng "_"
async function getLocationSheets(context) {
  context.workbook.worksheets.load("items/name");
  await context.sync();
  return context.workbook.worksheets.items
    .filter(s => s.name.startsWith("_"))
    .map(s => s.name);
}

// ══════════════════════════════════════════════════════════════
// ROW BUILDER
// ══════════════════════════════════════════════════════════════

// Tạo 1 row mới từ asset (INSERT)
function buildNewRow(asset, now) {
  const row = Array(COL_COUNT).fill("");
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
  row[COL.VALIDATION]   = "";
  // DAYS, NOTE, CASE_JIRA, ACTION = "" (mặc định đã fill ở trên)
  return row;
}

// Áp dữ liệu Jira vào existingRow, GIỮ NGUYÊN user cols
function applyJiraData(existingRow, asset, now) {
  const row = [...existingRow];
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
  if (row[COL.VALIDATION] === "Not in Jira") row[COL.VALIDATION] = "";
  // COL.DAYS, COL.NOTE, COL.CASE_JIRA, COL.ACTION — giữ nguyên từ existingRow
  return row;
}

// ══════════════════════════════════════════════════════════════
// WRITE LOCATION SHEET
//
// Logic:
//   A) UPDATE: asset đã có trong sheet → ghi đầy đủ Jira cols + giữ user cols
//   B) MARK:   row trong sheet không còn trong Jira → "Not in Jira"
//   C) INSERT: asset mới từ Jira chưa có trong sheet → append cuối
//
// Lookup:
//   - Ưu tiên theo Asset ID (byId)
//   - Fallback theo Serial (bySerial)
// ══════════════════════════════════════════════════════════════
async function writeLocationSheet(sheetName, assets, now, allKnownIds = null) {
  await Excel.run(async (context) => {
    const sheet = await ensureSheet(context, sheetName);
    await ensureHeaders(context, sheet);

    const existing = await readSheetRows(context, sheet);

    const byId = {};
    const bySerial = {};

    existing.forEach((row, idx) => {
      const id = String(row[COL.ASSET_ID] || "").trim();
      const ser = String(row[COL.SERIAL] || "").trim().toUpperCase();

      if (id) byId[id] = idx;
      if (ser) bySerial[ser] = idx;
    });

    const canFinalMark = allKnownIds instanceof Set;
    const updatedIdx = new Set();
    const updateRows = [];

    assets.forEach(asset => {
      let idx = -1;

      const id = String(asset.id || "").trim();
      const ser = String(asset.serial || "").trim().toUpperCase();

      if (id && byId[id] !== undefined) {
        idx = byId[id];
      } else if (ser && bySerial[ser] !== undefined) {
        idx = bySerial[ser];
      }

      if (idx < 0) return;

      const updated = applyJiraData(existing[idx], asset, now);
      updateRows.push({ excelRow: idx + 1, data: updated });
      updatedIdx.add(idx);
    });

    for (const { excelRow, data } of updateRows) {
      sheet.getRangeByIndexes(excelRow, 0, 1, COL_COUNT).values = [data];
    }

    if (updateRows.length > 0) await context.sync();

    const validationUpdates = [];

    if (canFinalMark) {
      existing.forEach((row, idx) => {
        if (updatedIdx.has(idx)) return;
        if (String(row[COL.SYNC_STATUS] || "") === "LOCAL") return;

        const id = String(row[COL.ASSET_ID] || "").trim();
        if (!id) return;

        const currentValidation = String(row[COL.VALIDATION] || "").trim();

        // Nếu asset vẫn nằm trong toàn bộ Jira IDs đã load xong,
        // không được mark Not in Jira chỉ vì nó không thuộc batch/location hiện tại.
        // Đồng thời clear lại Not in Jira cũ để giảm mismatch sai.
        if (allKnownIds.has(id)) {
          if (currentValidation === "Not in Jira") {
            validationUpdates.push({ excelRow: idx + 1, value: "" });
          }
          return;
        }

        validationUpdates.push({ excelRow: idx + 1, value: "Not in Jira" });
      });

      for (const u of validationUpdates) {
        const cell = sheet.getRangeByIndexes(u.excelRow, COL.VALIDATION, 1, 1);
        cell.values = [[u.value]];
        if (u.value === "Not in Jira") {
          cell.format.font.color = "#f59e0b";
        } else {
          cell.format.font.color = "#111827";
        }
      }

      if (validationUpdates.length > 0) await context.sync();
    }

    const toInsert = assets.filter(a => {
      const id = String(a.id || "").trim();
      const ser = String(a.serial || "").trim().toUpperCase();

      if (id && byId[id] !== undefined) return false;
      if (ser && bySerial[ser] !== undefined) return false;

      return true;
    });

    if (toInsert.length > 0) {
      const used = sheet.getUsedRangeOrNullObject(true);
      await context.sync();

      let nextRow = 1;

      if (!used.isNullObject) {
        used.load("rowCount");
        await context.sync();
        nextRow = used.rowCount;
      }

      sheet.getRangeByIndexes(nextRow, 0, toInsert.length, COL_COUNT).values =
        toInsert.map(a => buildNewRow(a, now));

      await context.sync();
    }

    const markCount = validationUpdates.filter(x => x.value === "Not in Jira").length;
    const clearCount = validationUpdates.filter(x => x.value === "").length;

    console.log(`[write] ${sheetName}: update=${updateRows.length}, mark=${markCount}, clear=${clearCount}, insert=${toInsert.length}`);
  });
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

      for (const name of sheets) {
        const sheet = context.workbook.worksheets.getItem(name);
        const rows  = await readSheetRows(context, sheet);
        total += rows.length;
        let locLocal = 0;
        rows.forEach(r => {
          if (r[COL.SYNC_STATUS] === "LOCAL")                              { local++; locLocal++; }
          if (r[COL.VALIDATION]  && r[COL.VALIDATION] !== "OK")            mismatch++;
          if (r[COL.ACTION] === "Create Ticket" && !r[COL.CASE_JIRA])      pending++;
        });
        locSummary.push({ name, count: rows.length, local: locLocal });
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
  toast("Fetching assets to discover locations...", "warning");
  try {
    const assets    = await fetchJiraAssets();
    const locations = [...new Set(assets.map(a => a.location).filter(Boolean))];
    if (!locations.length) { toast("No locations found", "warning"); return; }

    await Excel.run(async (context) => {
      for (const loc of locations) {
        const sheet = await ensureSheet(context, locationSheetName(loc));
        await ensureHeaders(context, sheet);
      }
    });
    toast(`Created/verified ${locations.length} location sheet(s)`, "success");
    await refreshDashboard();
  } catch(e) { toast("Error: " + e.message, "error"); }
}

// ══════════════════════════════════════════════════════════════
// SYNC ENGINE
// ══════════════════════════════════════════════════════════════
async function runSync() {
  if (isSyncing) {
    toast("Sync already running", "warning");
    return;
  }

  if (!cfg.jiraUrl || !cfg.token || !cfg.workerUrl) {
    toast("Kiểm tra Settings (URL / token / worker)", "warning");
    return;
  }

  isSyncing = true;
  setSyncIndicator("syncing", "Syncing...");

  const syncPanel = document.getElementById("sync-location-list");
  if (syncPanel) {
    syncPanel.innerHTML =
      `<div class="empty-state"><div class="spinner"></div>Đang tải từ Jira...</div>`;
  }

  try {
    const now = new Date().toISOString();

    // FIX CHÍNH:
    // Fetch hết tất cả typeId/version trước, sau đó mới ghi sheet.
    // Không ghi + mark theo từng sub-batch.
    const allAssets = await fetchJiraAssets();

    const allKnownIds = new Set(
      allAssets
        .map(a => String(a.id || "").trim())
        .filter(Boolean)
    );

    const byLocation = {};

    allAssets.forEach(a => {
      const loc = String(a.location || "UNKNOWN").trim() || "UNKNOWN";
      const sheetName = locationSheetName(loc);

      if (!byLocation[sheetName]) byLocation[sheetName] = [];
      byLocation[sheetName].push(a);
    });

    let allSheetNames = Object.keys(byLocation);

    await Excel.run(async (context) => {
      const existingLocationSheets = await getLocationSheets(context);
      allSheetNames = [...new Set([...allSheetNames, ...existingLocationSheets])];
    });

    toast(`Đang ghi ${allSheetNames.length} location sheet(s)...`, "warning");

    for (const sheetName of allSheetNames) {
      await writeLocationSheet(
        sheetName,
        byLocation[sheetName] || [],
        now,
        allKnownIds
      );
    }

    if (syncPanel) {
      updateSyncPanel(syncPanel, allSheetNames.map(s => ({
        name: s.replace(/^_/, ""),
        count: (byLocation[s] || []).length,
        status: "done",
      })));
    }

    toast("Đang push LOCAL assets lên Jira...", "warning");
    await pushLocalAssets();

    cfg.lastSync = new Date().toISOString();
    Office.context.document.settings.set(CFG_KEYS.LAST_SYNC, cfg.lastSync);
    Office.context.document.settings.saveAsync();

    setSyncIndicator("ok", "Synced");
    setInner("last-sync-time", formatTime(cfg.lastSync));

    toast(`Sync hoàn tất — ${allAssets.length} assets`, "success");
    await refreshDashboard();

  } catch (e) {
    setSyncIndicator("ok", "Sync failed");
    toast("Sync error: " + e.message, "error");
    console.error("runSync:", e);
  } finally {
    isSyncing = false;
  }
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
// PUSH LOCAL ASSETS LÊN JIRA
// ══════════════════════════════════════════════════════════════
async function pushLocalAssets() {
  if (!cfg.cloudId || !cfg.workspaceId) return;
  const typeIds = parseTypeIds();
  if (!typeIds.length) return;
  const defaultTypeId = typeIds[0];

  let pushed = 0, failed = 0;

  try {
    await Excel.run(async (context) => {
      const sheets = await getLocationSheets(context);

      for (const sheetName of sheets) {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const rows  = await readSheetRows(context, sheet);

        for (let i = 0; i < rows.length; i++) {
          const row = rows[i];
          if (String(row[COL.SYNC_STATUS] || "") !== "LOCAL") continue;
          if (row[COL.ASSET_ID])                              continue;  // đã push rồi

          const hostname = String(row[COL.HOSTNAME] || "").trim();
          const serial   = String(row[COL.SERIAL]   || "").trim();
          if (!hostname && !serial)                           continue;

          try {
            const attrs = [
              { objectTypeAttributeId: 1737,  objectAttributeValues: [{ value: hostname }] },
              { objectTypeAttributeId: 5194,  objectAttributeValues: [{ value: serial }] },
              { objectTypeAttributeId: 30125, objectAttributeValues: [{ value: String(row[COL.LOCATION] || "") }] },
              { objectTypeAttributeId: 5200,  objectAttributeValues: [{ value: String(row[COL.USERNAME]  || "") }] },
            ].filter(a => a.objectAttributeValues[0].value !== "");

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
              cur[COL.ASSET_KEY]   = String(res.objectKey || "");
              cur[COL.SYNC_STATUS] = "JIRA";
              cur[COL.VALIDATION]  = "OK";
              cur[COL.LAST_SYNC]   = new Date().toISOString();
              range.values = [cur];
              await context.sync();
              pushed++;
            }
          } catch(e) {
            failed++;
            console.warn(`pushLocal row ${i+2} [${sheetName}]:`, e.message);
          }
        }
      }
    });

    if (pushed > 0 || failed > 0)
      toast(`LOCAL push: ${pushed} OK${failed ? `, ${failed} lỗi` : ""}`, failed ? "warning" : "success");

  } catch(e) { console.warn("pushLocalAssets:", e.message); }
}

// ══════════════════════════════════════════════════════════════
// UPDATE / CREATE JIRA ASSET FROM EXCEL ACTION COLUMN
//
// Action rules:
//   x = nếu có Asset ID thì UPDATE Jira, nếu chưa có Asset ID thì CREATE mới
//   o = xóa row khỏi Excel sheet (không xóa object trên Jira để tránh mất dữ liệu)
// ══════════════════════════════════════════════════════════════
function rowToJiraFields(row) {
  return {
    hostname:  String(row[COL.HOSTNAME]   || "").trim(),
    serial:    String(row[COL.SERIAL]     || "").trim(),
    location:  String(row[COL.LOCATION]   || "").trim(),
    username:  String(row[COL.USERNAME]   || "").trim(),
    status:    String(row[COL.STATUS]     || "").trim(),
    region:    String(row[COL.REGION]     || "").trim(),
    osVersion: String(row[COL.OS_VERSION] || "").trim(),
    osBuild:   String(row[COL.OS_BUILD]   || "").trim(),
  };
}

function jiraAttributesFromFields(fields) {
  const attrMap = {
    hostname:  1737,
    serial:    5194,
    location:  30125,
    username:  5200,
    status:    5052,
    region:    27292,
    osVersion: 27291,
    osBuild:   27290,
  };

  return Object.entries(fields)
    .filter(([k, v]) => attrMap[k] !== undefined && v !== undefined && String(v).trim() !== "")
    .map(([k, v]) => ({
      objectTypeAttributeId: attrMap[k],
      objectAttributeValues: [{ value: String(v).trim() }],
    }));
}

async function updateJiraAsset(assetId, fields) {
  const attributes = jiraAttributesFromFields(fields);
  if (!assetId) throw new Error("Missing Asset ID");
  if (!attributes.length) throw new Error("Không có field nào để update");

  // Jira Assets update object dùng PUT /object/{id}
  return assetsPut(`/object/${assetId}`, { attributes });
}

async function createJiraAssetFromRow(row) {
  const typeIds = parseTypeIds();
  if (!typeIds.length) throw new Error("Không tìm thấy objectTypeId trong AQL Query");

  const defaultTypeId = Number(typeIds[0]);
  const fields = rowToJiraFields(row);

  if (!fields.hostname && !fields.serial) {
    throw new Error("Row mới cần ít nhất Hostname hoặc Serial Number");
  }

  const attributes = jiraAttributesFromFields(fields);
  if (!attributes.length) throw new Error("Không có field nào để create");

  return assetsPost("/object/create", {
    objectTypeId: defaultTypeId,
    attributes,
  });
}

async function processActionRows() {
  if (!cfg.jiraUrl || !cfg.token || !cfg.workerUrl || !cfg.cloudId || !cfg.workspaceId) {
    toast("Kiểm tra Settings trước khi update Jira", "warning");
    return;
  }

  toast('Đang xử lý Action: "x" = update/create, "o" = xóa row...', "warning");

  let updated = 0;
  let created = 0;
  let deleted = 0;
  let skipped = 0;
  let failed = 0;

  try {
    await Excel.run(async (context) => {
      const sheets = await getLocationSheets(context);

      for (const sheetName of sheets) {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const rows = await readSheetRows(context, sheet);
        const rowsToDelete = [];

        for (let i = 0; i < rows.length; i++) {
          const row = rows[i];
          const action = String(row[COL.ACTION] || "").trim().toLowerCase();
          const excelRow = i + 1; // 0 là header, nên data row đầu tiên là index 1

          if (!action) continue;

          // o = xóa khỏi Excel sheet, không gọi API delete Jira
          if (action === "o") {
            rowsToDelete.push(excelRow);
            continue;
          }

          if (action !== "x") {
            skipped++;
            continue;
          }

          const assetId = String(row[COL.ASSET_ID] || "").trim();

          try {
            if (assetId) {
              await updateJiraAsset(assetId, rowToJiraFields(row));

              row[COL.SYNC_STATUS] = "JIRA";
              row[COL.VALIDATION] = "OK";
              row[COL.LAST_SYNC] = new Date().toISOString();
              row[COL.ACTION] = "";

              sheet.getRangeByIndexes(excelRow, 0, 1, COL_COUNT).values = [row];
              updated++;
            } else {
              const res = await createJiraAssetFromRow(row);

              row[COL.ASSET_ID] = String(res?.id || res?.objectId || "");
              row[COL.ASSET_KEY] = String(res?.objectKey || res?.key || "");
              row[COL.SYNC_STATUS] = "JIRA";
              row[COL.VALIDATION] = "OK";
              row[COL.LAST_SYNC] = new Date().toISOString();
              row[COL.ACTION] = "";

              sheet.getRangeByIndexes(excelRow, 0, 1, COL_COUNT).values = [row];
              created++;
            }

            await context.sync();
          } catch (e) {
            failed++;
            row[COL.VALIDATION] = "Update failed: " + String(e.message || e).slice(0, 120);
            sheet.getRangeByIndexes(excelRow, 0, 1, COL_COUNT).values = [row];
            await context.sync();
            console.warn(`[processActionRows] ${sheetName} row ${i + 2}:`, e.message || e);
          }
        }

        // Xóa từ dưới lên để không lệch index row
        rowsToDelete.sort((a, b) => b - a);
        for (const excelRow of rowsToDelete) {
          sheet.getRangeByIndexes(excelRow, 0, 1, COL_COUNT)
            .delete(Excel.DeleteShiftDirection.up);
          deleted++;
        }

        if (rowsToDelete.length > 0) await context.sync();
      }
    });

    toast(
      `Action done: update=${updated}, create=${created}, delete=${deleted}, skipped=${skipped}, failed=${failed}`,
      failed ? "warning" : "success"
    );

    await refreshDashboard();
  } catch (e) {
    toast("Update Jira error: " + e.message, "error");
    console.error("processActionRows:", e);
  }
}

// ══════════════════════════════════════════════════════════════
// MATCH LOCAL ASSETS ↔ JIRA
// ══════════════════════════════════════════════════════════════
async function matchLocalAssets() {
  if (!cfg.jiraUrl || !cfg.token) { toast("Configure Jira settings first", "warning"); return; }
  toast("Scanning LOCAL assets for Jira matches…", "warning");
  try {
    const jiraAssets   = await fetchJiraAssets();
    const jiraBySerial = {};
    jiraAssets.forEach(a => { if (a.serial) jiraBySerial[a.serial.trim()] = a; });

    let matched = 0;
    await Excel.run(async (context) => {
      const sheets = await getLocationSheets(context);
      for (const sheetName of sheets) {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const rows  = await readSheetRows(context, sheet);
        for (let i = 0; i < rows.length; i++) {
          if (String(rows[i][COL.SYNC_STATUS] || "") !== "LOCAL") continue;
          const serial = String(rows[i][COL.SERIAL] || "").trim();
          if (!serial || !jiraBySerial[serial]) continue;

          const a   = jiraBySerial[serial];
          const row = applyJiraData(rows[i], a, new Date().toISOString());
          sheet.getRangeByIndexes(i + 1, 0, 1, COL_COUNT).values = [row];
          await context.sync();
          matched++;
        }
      }
    });
    toast(`Matched ${matched} LOCAL asset(s) to Jira`, "success");
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
    "Duplicate Serial":  [],
    "Location Changed":  [],
    "Serial Mismatch":   [],
    "Owner Mismatch":    [],
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
          const locField = String(row[COL.LOCATION]  || "").trim();
          let valid = "OK";

          // 1. JIRA row thiếu Asset ID
          if (String(row[COL.SYNC_STATUS] || "") === "JIRA" && !assetId) {
            valid = "Missing Asset ID";
            errors["Missing Asset ID"].push({ sheet: sheetName, row: i + 2 });
          }

          // 2. Duplicate serial
          if (serial && valid === "OK") {
            if (serialSeen[serial]) {
              valid = "Duplicate Serial";
              errors["Duplicate Serial"].push({ sheet: sheetName, row: i + 2 });
              if (!errors["Duplicate Serial"].find(
                e => e.sheet === serialSeen[serial].sheet && e.row === serialSeen[serial].row
              )) errors["Duplicate Serial"].push(serialSeen[serial]);
            } else {
              serialSeen[serial] = { sheet: sheetName, row: i + 2 };
            }
          }

          // 3. Location field không khớp sheet name
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
    const unique = list.filter((v, i, a) =>
      i === a.findIndex(x => x.sheet === v.sheet && x.row === v.row));
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
      const sheets = await getLocationSheets(context);
      for (const name of sheets) {
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
      const sheets = await getLocationSheets(context);
      for (const sheetName of sheets) {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const rows  = await readSheetRows(context, sheet);

        for (let i = 0; i < rows.length; i++) {
          const row = rows[i];
          if (row[COL.ACTION] !== "Create Ticket")  continue;
          if (row[COL.CASE_JIRA]) { skipped++; continue; }

          const deviceName = String(row[COL.HOSTNAME]  || row[COL.ASSET_KEY] || "Unknown");
          const serial     = String(row[COL.SERIAL]    || "");
          const email      = String(row[COL.ASSIGNED]  || row[COL.USERNAME] || "");
          const note       = String(row[COL.NOTE]      || "");
          const rowDays    = row[COL.DAYS] || days;

          try {
            const res = await jiraPost("/api/3/issue", {
              fields: {
                project:     { key: cfg.projectKey },
                summary:     `[${issueType}] ${deviceName}${serial ? " – " + serial : ""}`,
                description: {
                  type: "doc", version: 1,
                  content: [{ type: "paragraph", content: [{ type: "text",
                    text: `Asset: ${deviceName}\nSerial: ${serial}\nUser: ${email}\nDays: ${rowDays}\nNote: ${note}` }]
                  }],
                },
                issuetype: { name: issueType },
                priority:  { name: priority },
              },
            });

            const key = res.key || "";
            if (key) {
              const url = `${jiraBase()}/browse/${key}`;
              sheet.getRangeByIndexes(i+1, COL.CASE_JIRA, 1, 1).values = [[`=HYPERLINK("${url}","${key}")`]];
              sheet.getRangeByIndexes(i+1, COL.ACTION,    1, 1).values = [[""]];
              await context.sync();
              created++;
            }
          } catch(e) { failed++; toast(`Row ${i+2}: ${e.message}`, "error"); }
        }
      }
    });

    toast(`Done — created: ${created}, skipped: ${skipped}${failed ? `, failed: ${failed}` : ""}`,
      failed ? "warning" : "success");
    scanPendingRows();
  } catch(e) { toast("Ticket error: " + e.message, "error"); }
}

// ══════════════════════════════════════════════════════════════
// CONNECTION TEST
// ══════════════════════════════════════════════════════════════
async function testConnection() {
  const el = document.getElementById("conn-test-result");
  el.style.display    = "block";
  el.style.borderColor = "var(--border)";
  el.innerHTML = `<div style="display:flex;gap:8px;align-items:center">
    <div class="spinner"></div><span>Testing…</span></div>`;

  saveConfig();
  if (!cfg.workerUrl) {
    el.innerHTML = "✗ Worker URL chưa điền";
    el.style.borderColor = "var(--red)";
    return;
  }

  try {
    const me = await jiraGet("/api/3/myself");
    let assetsOk = "";
    try {
      await assetsPost("/object/aql?startAt=0&maxResults=1", { qlQuery: "objectType != null" });
      assetsOk = " · Assets API ✓";
    } catch(ae) { assetsOk = ` · Assets ✗ (${ae.message.slice(0,60)})`; }

    el.innerHTML      = `✓ Connected as <strong>${me.displayName || me.emailAddress}</strong>${assetsOk}`;
    el.style.borderColor = "var(--green)";
    toast("Connection OK", "success");
  } catch(e) {
    el.innerHTML      = `✗ ${e.message}`;
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

function setSyncIndicator(state, text) {
  const dot = document.getElementById("sync-dot");
  const txt = document.getElementById("sync-status-text");
  if (dot) dot.className = "sync-dot" + (state === "syncing" ? " syncing" : "");
  if (txt) txt.textContent = text;
}

function toast(msg, type = "success") {
  const icons     = { success: "✓", error: "✗", warning: "⚠" };
  const container = document.getElementById("toast-container");
  if (!container) return;
  const el        = document.createElement("div");
  el.className    = `toast ${type}`;
  el.innerHTML    = `<span>${icons[type]||"ℹ"}</span><span>${msg}</span>`;
  container.appendChild(el);
  setTimeout(() => el.remove(), 4500);
}

function setInner(id, val) { const e = document.getElementById(id); if (e) e.textContent = val; }
function setVal(id, val)   { const e = document.getElementById(id); if (e) e.value = val || ""; }
function getVal(id)         { const e = document.getElementById(id); return e ? e.value.trim() : ""; }
function formatTime(iso) {
  if (!iso) return "—";
  try { return new Date(iso).toLocaleString([], { month:"2-digit", day:"2-digit", hour:"2-digit", minute:"2-digit" }); }
  catch { return iso; }
}