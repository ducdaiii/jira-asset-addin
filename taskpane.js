/* ════════════════════════════════════════════════════════════
   Jira Asset Manager — Office Add-in
   PATCH: auto-move rows to correct Location sheet on sync
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

// Các cột kỹ thuật nên ẩn bớt cho user khi tạo/ghi sheet.
// Hide column không ảnh hưởng index nên add-in vẫn chạy bình thường.
const SYSTEM_HIDDEN_COLS = [
  COL.OS,           // Operating System
  COL.OS_VERSION,   // Windows Version
  COL.OS_BUILD,     // Windows Build
  COL.CPU,          // CPU
  COL.IP,           // IP Address
  COL.NETWORK,      // Network Name
  COL.ANTIVIRUS,    // Antivirus
  COL.TENANT_ID,    // Tenant ID / Source ID
  COL.LANSWEEPER,   // Lansweeper URL
];

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
  JIRA_TOTAL:   "jiraTotal",
};

let cfg      = {};
let isSyncing = false;

// Audit note: ghi người chỉnh sửa vào cột Note khi user sửa row trong sheet location.
let auditNoteHandlersRegistered = false;
let isAuditNoteWriting = false;
let cachedCurrentUserEmail = "";

// Status attribute trong Jira Assets không phải text.
// Excel hiển thị tên như "IN USE", nhưng API cần status.id, ví dụ IN USE -> 31.
const STATUS_ATTR_ID = "5052";
let statusNameToId = {};
let statusIdToName = {};
let statusOptionsLoaded = false;

// Owner/Assigned User là reference object.
// Assigned User dùng attribute 26690 và trỏ tới objectTypeId=269 (Users).
const OWNER_ATTR_ID = "26690";
const OWNER_OBJECT_TYPE_ID = "269";
const OWNER_METADATA_SHEET = "_SYS_OWNER_METADATA";

let ownerNameToObject = {};
let ownerKeyToObject = {};
let ownerOptionsLoaded = false;
let ownerChangeHandlersRegistered = false;

// Cache attribute hợp lệ theo objectTypeId để tránh gửi field không thuộc schema.
// Ví dụ: một số objectType không có Assigned User attribute 26690.
const objectTypeAttributeCache = new Map();

function normalizeStatusName(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[\s_-]+/g, " ");
}

function rememberStatusOption(name, id) {
  const rawName = String(name || "").trim();
  const normalized = normalizeStatusName(rawName);
  const sid = String(id || "").trim();
  if (!normalized || !sid) return;

  // Lưu cả 2 chiều để vừa map API vừa tạo dropdown Excel.
  statusNameToId[normalized] = sid;
  statusIdToName[sid] = rawName.toUpperCase();
}

function getStatusIdFromCache(name) {
  return statusNameToId[normalizeStatusName(name)] || "";
}

function getCachedStatusNamesArray() {
  return Object.values(statusIdToName)
    .filter(Boolean)
    .filter((v, i, arr) => arr.findIndex(x => normalizeStatusName(x) === normalizeStatusName(v)) === i)
    .sort((a, b) => a.localeCompare(b));
}

function listCachedStatusNames() {
  return getCachedStatusNamesArray().join(", ");
}

function normalizeOwnerName(value) {
  return String(value || "")
    .trim()
    .replace(/\s*\([^)]*\)\s*/g, " ")     // bỏ "(Le Thi Hai Thu | Human Resources)"
    .replace(/\s*\[[^\]]*\]\s*/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function ownerNameAliases(value) {
  const raw = String(value || "").trim();
  if (!raw) return [];

  const noParen = raw
    .replace(/\s*\([^)]*\)\s*/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const aliases = new Set([raw, noParen]);

  // "Thu, Lisa" -> "Lisa Thu"
  const commaName = noParen.match(/^([^,]+),\s*(.+)$/);
  if (commaName) {
    aliases.add(`${commaName[2]} ${commaName[1]}`.trim());
  }

  // "Thu, Lisa (Le Thi Hai Thu | Human Resources)" -> "Le Thi Hai Thu"
  const inside = raw.match(/\(([^|)]+)(?:\|[^)]*)?\)/);
  if (inside && inside[1]) {
    aliases.add(inside[1].trim());
  }

  return [...aliases]
    .map(x => String(x || "").trim())
    .filter(Boolean);
}

function rememberOwnerOption(owner) {
  const id = String(owner?.id || owner?.objectId || "").trim();
  const key = String(owner?.key || owner?.objectKey || "").trim();
  const label = String(owner?.label || owner?.name || owner?.displayValue || "").trim();
  if (!id || !label) return;

  const username = String(owner?.username || owner?.userName || owner?.searchValue || "").trim() || key || label;

  const item = { id, key, label, username };

  ownerNameAliases(label).forEach(alias => {
    ownerNameToObject[normalizeOwnerName(alias)] = item;
  });

  ownerNameAliases(username).forEach(alias => {
    ownerNameToObject[normalizeOwnerName(alias)] = item;
  });

  if (key) ownerKeyToObject[key.toUpperCase()] = item;
}

function getOwnerFromCache(value) {
  const raw = String(value || "").trim();
  if (!raw) return null;

  const byKey = ownerKeyToObject[raw.toUpperCase()];
  if (byKey) return byKey;

  for (const alias of ownerNameAliases(raw)) {
    const found = ownerNameToObject[normalizeOwnerName(alias)];
    if (found) return found;
  }

  // Fallback mềm: Excel value chứa label hoặc ngược lại.
  const needle = normalizeOwnerName(raw);
  for (const [k, owner] of Object.entries(ownerNameToObject)) {
    if (!k || !needle) continue;
    if (k.includes(needle) || needle.includes(k)) return owner;
  }

  return null;
}

function getCachedOwnerNamesArray() {
  const seen = new Set();
  return Object.values(ownerNameToObject)
    .map(o => String(o.label || "").trim())
    .filter(Boolean)
    .filter(name => {
      const k = normalizeOwnerName(name);
      if (seen.has(k)) return false;
      seen.add(k);
      return true;
    })
    .sort((a, b) => a.localeCompare(b));
}

function parseOwnerObject(obj) {
  const owner = {
    id: String(obj?.id ?? obj?.objectId ?? ""),
    key: String(obj?.objectKey || obj?.key || ""),
    label: String(obj?.label || obj?.name || ""),
    username: "",
  };

  const attrs = obj?.attributes || [];
  for (const a of attrs) {
    const vals = a?.objectAttributeValues || [];
    const v = vals[0] || {};
    const display = String(v.displayValue || v.value || v.searchValue || "").trim();
    const attrName = String(a.name || a.label || "").toLowerCase();

    if (!owner.username && display && (
      attrName.includes("username") ||
      attrName.includes("user name") ||
      attrName.includes("email") ||
      attrName.includes("mail") ||
      attrName.includes("login")
    )) {
      owner.username = display;
    }
  }

  if (!owner.username) owner.username = owner.key || owner.label;

  rememberOwnerOption(owner);

  [
    obj?.displayValue,
    obj?.searchValue,
    obj?.name,
    obj?.label,
    obj?.objectKey,
    obj?.key,
  ].forEach(v => {
    ownerNameAliases(v).forEach(alias => {
      ownerNameToObject[normalizeOwnerName(alias)] = owner;
    });
  });

  return owner;
}

// ══════════════════════════════════════════════════════════════
// INIT
// ══════════════════════════════════════════════════════════════
Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) return;
  loadConfig();
  populateConfigUI();
  wireEvents();
  await refreshDashboard();
  await registerOwnerChangeHandlers();
  await registerAuditNoteHandlers();
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
    jiraTotal:   Number(s.get(CFG_KEYS.JIRA_TOTAL) || 0),
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
  on("btn-refresh-status", applyStatusDropdownAllSheets);
  on("btn-refresh-status-dropdown", refreshMetadataDropdowns);
  on("btn-refresh-owner", applyOwnerDropdownAllSheets);
  on("btn-refresh-owner-dropdown", applyOwnerDropdownAllSheets);
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

async function assetsGet(path) {
  const res = await fetch(proxyUrl(`${assetsBase()}${path}`), {
    method: "GET",
    headers: jiraHeaders(),
  });

  if (!res.ok) {
    throw new Error(`Assets GET ${res.status}: ${(await res.text().catch(() => "")).slice(0,200)}`);
  }

  return res.json();
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

async function assetsDelete(path) {
  const res = await fetch(proxyUrl(`${assetsBase()}${path}`), {
    method: "DELETE",
    headers: jiraHeaders(),
  });

  if (!res.ok) {
    throw new Error(`Assets DELETE ${res.status}: ${(await res.text().catch(() => "")).slice(0,200)}`);
  }

  return true;
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
  const getAttrObj = (id) => (obj.attributes || []).find(
    x => String(x.objectTypeAttributeId) === String(id)
  );

  const attr = (id) => {
    const a = getAttrObj(id);
    return a?.objectAttributeValues?.[0]?.displayValue || "";
  };

  const statusAttr = getAttrObj(STATUS_ATTR_ID);
  const statusVal  = statusAttr?.objectAttributeValues?.[0] || null;
  const statusName = statusVal?.displayValue || statusVal?.status?.name || "";
  const statusId   = statusVal?.status?.id || "";
  rememberStatusOption(statusName, statusId);

  const ownerAttr = getAttrObj(OWNER_ATTR_ID);
  const ownerVal = ownerAttr?.objectAttributeValues?.[0] || null;
  const ownerRef = ownerVal?.referencedObject || null;
  if (ownerRef) {
    rememberOwnerOption({
      id: ownerRef.id,
      key: ownerRef.objectKey,
      label: ownerRef.label || ownerRef.name || ownerVal.displayValue,
      username: ownerVal.searchValue || ownerRef.objectKey || "",
    });
  }

  return {
    id:           String(obj.id ?? obj.objectId ?? ""),
    key:          String(obj.objectKey || obj.key || ""),
    hostname:     attr(1737) || obj.label || "",
    serial:       attr(5194),
    status:       statusName || attr(5052),
    statusId:     String(statusId || ""),
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
      : `${parentQuery} AND "OS Build" = "${build.replace(/"/g, '\"')}"`;

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
      : `${baseQ} AND "Version OS" = "${ver.replace(/"/g, '\"')}"`;

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
  if (r.values[0][0] === "Asset ID") {
    await hideSystemColumns(context, sheet);
    return;   // đã có header
  }
  r.values = [HEADERS];
  r.format.font.bold  = true;
  r.format.fill.color = "#1a1d27";
  r.format.font.color = "#8892a4";
  sheet.freezePanes.freezeRows(1);
  await hideSystemColumns(context, sheet);
  await context.sync();
}

async function hideSystemColumns(context, sheet) {
  // Ẩn cột theo index, không xóa cột.
  // Add-in vẫn đọc/ghi đúng vì COL map không đổi.
  // Dùng columnHidden trực tiếp ổn định hơn format.columnHidden trên Excel Online.
  for (const colIdx of SYSTEM_HIDDEN_COLS) {
    try {
      const col = sheet.getRangeByIndexes(0, colIdx, 1, 1).getEntireColumn();

      // ExcelApi hỗ trợ trực tiếp property columnHidden trên Range.
      col.columnHidden = true;

      // Fallback cho một số host cũ.
      try {
        col.format.columnHidden = true;
      } catch (_) {}

    } catch (e) {
      console.warn("hideSystemColumns:", e.message || e);
    }
  }

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
    .filter(s => s.name.startsWith("_") && !s.name.startsWith("_SYS_"))
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
async function writeLocationSheet(sheetName, assets, now, allKnownIds = null, allAssetById = null) {
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
    const rowsToDeleteBecauseMoved = [];
    const rowsToDeleteBecauseMissing = [];

    if (canFinalMark) {
      existing.forEach((row, idx) => {
        if (updatedIdx.has(idx)) return;
        if (String(row[COL.SYNC_STATUS] || "") === "LOCAL") return;

        const id = String(row[COL.ASSET_ID] || "").trim();
        if (!id) return;

        const currentValidation = String(row[COL.VALIDATION] || "").trim();

        // MOVE FIX:
        // Nếu asset vẫn còn trên Jira nhưng Location hiện tại của Jira thuộc sheet khác,
        // xóa row cũ khỏi sheet này. Row đúng sẽ được INSERT/UPDATE ở sheet đích.
        // Tránh tình trạng sau khi update Location, asset nằm cả ở sheet cũ hoặc không tự chuyển sheet.
        if (allKnownIds.has(id) && allAssetById instanceof Map) {
          const canonical = allAssetById.get(id);
          if (canonical) {
            const targetLoc = String(canonical.location || "UNKNOWN").trim() || "UNKNOWN";
            const targetSheet = locationSheetName(targetLoc);
            if (targetSheet !== sheetName) {
              rowsToDeleteBecauseMoved.push(idx + 1);
              return;
            }
          }
        }

        // Nếu asset vẫn nằm trong toàn bộ Jira IDs đã load xong,
        // không được mark Not in Jira chỉ vì nó không thuộc batch/location hiện tại.
        // Đồng thời clear lại Not in Jira cũ để giảm mismatch sai.
        if (allKnownIds.has(id)) {
          if (currentValidation === "Not in Jira") {
            validationUpdates.push({ excelRow: idx + 1, value: "" });
          }
          return;
        }

        // Nếu row không còn trong Jira full sync nữa thì xóa khỏi sheet.
        // Không áp dụng cho LOCAL vì LOCAL đã được skip ở trên.
        rowsToDeleteBecauseMissing.push(idx + 1);
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

      // Delete stale rows from old location sheets or rows removed from Jira.
      // Delete from bottom to top so Excel row indexes do not shift.
      const rowsToDelete = [...new Set([
        ...rowsToDeleteBecauseMoved,
        ...rowsToDeleteBecauseMissing,
      ])].sort((a, b) => b - a);

      for (const excelRow of rowsToDelete) {
        sheet.getRangeByIndexes(excelRow, 0, 1, COL_COUNT)
          .delete(Excel.DeleteShiftDirection.up);
      }
      if (rowsToDelete.length > 0) await context.sync();
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

    await hideSystemColumns(context, sheet);
    await context.sync();

    const markCount = validationUpdates.filter(x => x.value === "Not in Jira").length;
    const clearCount = validationUpdates.filter(x => x.value === "").length;
    const movedDeleteCount = rowsToDeleteBecauseMoved.length;
    const missingDeleteCount = rowsToDeleteBecauseMissing.length;

    console.log(`[write] ${sheetName}: update=${updateRows.length}, mark=${markCount}, clear=${clearCount}, movedDelete=${movedDeleteCount}, missingDelete=${missingDeleteCount}, insert=${toInsert.length}`);
  });
}


// ══════════════════════════════════════════════════════════════
// STATUS DROPDOWN IN EXCEL
// ══════════════════════════════════════════════════════════════
async function applyStatusDropdownToSheet(context, sheet) {
  const names = getCachedStatusNamesArray();
  if (!names.length) return;

  // Status hiện tại không có dấu phẩy. Nếu sau này có dấu phẩy, nên chuyển sang hidden metadata sheet.
  const source = names.join(",");

  // Áp dropdown cho 5000 dòng dưới header.
  const range = sheet.getRangeByIndexes(1, COL.STATUS, 5000, 1);
  range.dataValidation.rule = {
    list: {
      inCellDropDown: true,
      source,
    },
  };

  range.dataValidation.errorAlert = {
    showAlert: true,
    style: Excel.DataValidationAlertStyle.stop,
    title: "Invalid Status",
    message: `Chỉ chọn Status từ danh sách Jira: ${names.join(", ")}`,
  };
}

async function applyStatusDropdownAllSheets() {
  try {
    await ensureStatusMap();

    await Excel.run(async (context) => {
      const sheets = await getLocationSheets(context);
      for (const sheetName of sheets) {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        await applyStatusDropdownToSheet(context, sheet);
      }
      await context.sync();
    });

    toast(`Đã cập nhật dropdown Status (${getCachedStatusNamesArray().length} giá trị)`, "success");
  } catch (e) {
    console.warn("applyStatusDropdownAllSheets:", e.message || e);
    toast("Không cập nhật được dropdown Status: " + (e.message || e), "warning");
  }
}

// ══════════════════════════════════════════════════════════════
// OWNER DROPDOWN IN EXCEL
// ══════════════════════════════════════════════════════════════
function ownerDedupKey(owner) {
  const id = String(owner?.id || "").trim();
  if (id) return `id:${id}`;

  const key = String(owner?.key || "").trim().toUpperCase();
  if (key) return `key:${key}`;

  const label = normalizeOwnerName(owner?.label || "");
  if (label) return `label:${label}`;

  return "";
}

async function fetchOwnersUnderLimit(qlQuery) {
  const PAGE = 100;
  const owners = [];
  const seenPageSignatures = new Set();

  const first = await fetchPage(qlQuery, 0, PAGE);
  const total = typeof first.total === "number" ? first.total : 0;
  const firstVals = first.values || [];

  const pushPage = (values, startAt) => {
    const parsed = (values || []).map(o => parseOwnerObject(o));

    const sig = parsed
      .map(o => String(o.id || o.key || o.label || "").trim())
      .join("|");

    if (sig && seenPageSignatures.has(sig)) {
      console.warn(`[ownerMap] duplicate page at startAt=${startAt}. Pagination not moving.`);
      return false;
    }

    if (sig) seenPageSignatures.add(sig);
    parsed.forEach(o => owners.push(o));
    return true;
  };

  pushPage(firstVals, 0);

  if (total === 0 || firstVals.length === 0) return owners;

  let startAt = firstVals.length;

  while (owners.length < total && startAt < API_LIMIT) {
    const remaining = total - owners.length;
    const pageSize = Math.min(PAGE, remaining, API_LIMIT - startAt);
    if (pageSize <= 0) break;

    const data = await fetchPage(qlQuery, startAt, pageSize);
    const values = data.values || [];
    if (!values.length) break;

    const ok = pushPage(values, startAt);
    if (!ok) break;

    startAt += values.length;
  }

  if (owners.length > total && total > 0) {
    owners.length = total;
  }

  return owners;
}

function buildOwnerSplitQueries(baseQ, attrName) {
  const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");
  const digits = "0123456789".split("");
  const queries = [];

  // Jira Assets AQL thường hỗ trợ LIKE với wildcard.
  // Với Users, label/name thường bắt đầu bằng Last name, ví dụ "Thu, Lisa...".
  letters.forEach(ch => {
    queries.push(`${baseQ} AND "${attrName}" LIKE "${ch}%"`);
  });

  digits.forEach(ch => {
    queries.push(`${baseQ} AND "${attrName}" LIKE "${ch}%"`);
  });

  // Bucket cho tên không bắt đầu bằng chữ/số.
  // Nếu Jira không support nhiều NOT LIKE, query này sẽ fail nhẹ và bị bỏ qua.
  const notLike = [...letters, ...digits]
    .map(ch => `"${attrName}" NOT LIKE "${ch}%"`)
    .join(" AND ");

  queries.push(`${baseQ} AND ${notLike}`);

  return queries;
}

async function fetchOwnersSplitByAttribute(attrName) {
  const baseQ = `objectTypeId = ${OWNER_OBJECT_TYPE_ID}`;
  const queries = buildOwnerSplitQueries(baseQ, attrName);

  const seen = new Map();
  let rawFetched = 0;
  let successBuckets = 0;

  for (const q of queries) {
    try {
      const total = await fetchTotalCount(q);
      if (total === 0) continue;

      if (total >= API_LIMIT) {
        console.warn(`[ownerMap] bucket still >= ${API_LIMIT}: ${q} total=${total}`);
        toast(`⚠ Owner bucket "${attrName}" vẫn >= ${API_LIMIT}; có thể thiếu user trong bucket này`, "warning");
      }

      const list = await fetchOwnersUnderLimit(q);
      rawFetched += list.length;
      successBuckets++;

      list.forEach(o => {
        const k = ownerDedupKey(o);
        if (!k) return;
        if (!seen.has(k)) seen.set(k, o);
      });

      console.log(`[ownerMap] bucket attr="${attrName}" total=${total}, fetched=${list.length}, unique=${seen.size}`);
    } catch (e) {
      console.warn(`[ownerMap] split query failed attr="${attrName}":`, e.message || e);
    }
  }

  return {
    owners: [...seen.values()],
    rawFetched,
    successBuckets,
  };
}

async function fetchOwnersFromJira() {
  // Reset cache để Refresh Owner luôn lấy mới.
  ownerNameToObject = {};
  ownerKeyToObject = {};
  ownerOptionsLoaded = false;

  const baseQ = `objectTypeId = ${OWNER_OBJECT_TYPE_ID}`;
  const total = await fetchTotalCount(baseQ).catch(() => 0);

  console.log(`[ownerMap] total Users=${total}`);

  let owners = [];

  if (total > 0 && total < API_LIMIT) {
    owners = await fetchOwnersUnderLimit(baseQ);
  } else {
    // Users >= 1000: chia theo attribute tên giống cách device chia theo Version OS.
    // Thử nhiều field phổ biến vì schema Users có thể đặt tên attribute khác nhau.
    const splitAttrs = [
      "Name",
      "Display Name",
      "Full Name",
      "Email",
      "User principal name",
      "Username",
    ];

    let best = { owners: [], rawFetched: 0, successBuckets: 0, attr: "" };

    for (const attrName of splitAttrs) {
      const result = await fetchOwnersSplitByAttribute(attrName);

      if (result.owners.length > best.owners.length) {
        best = { ...result, attr: attrName };
      }

      // Nếu đã lấy gần đủ total thì dừng.
      if (total > 0 && result.owners.length >= total * 0.98) {
        best = { ...result, attr: attrName };
        break;
      }
    }

    owners = best.owners;

    console.log(
      `[ownerMap] split done attr="${best.attr}", total=${total}, unique=${owners.length}, raw=${best.rawFetched}, buckets=${best.successBuckets}`
    );

    if (total > 0 && owners.length < total) {
      toast(`⚠ Owner loaded ${owners.length}/${total}. Nếu thiếu user, cần chỉnh split attribute cho Users schema.`, "warning");
    }
  }

  owners.forEach(o => rememberOwnerOption(o));

  ownerOptionsLoaded = owners.length > 0;
  console.log(`[ownerMap] loaded ${owners.length} owners`);

  return owners;
}


async function loadOwnerMapFromMetadataSheet() {
  // Đọc cache Owner từ sheet _SYS_OWNER_METADATA nếu đã refresh trước đó.
  // Hàm này KHÔNG gọi Jira API.
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItemOrNullObject(OWNER_METADATA_SHEET);
      await context.sync();

      if (sheet.isNullObject) return;

      const used = sheet.getUsedRangeOrNullObject(true);
      await context.sync();

      if (used.isNullObject) return;

      used.load(["values", "rowCount"]);
      await context.sync();

      if (used.rowCount <= 1) return;

      const rows = used.values.slice(1);
      rows.forEach(r => {
        const label = String(r[0] || "").trim();
        const id = String(r[1] || "").trim();
        const key = String(r[2] || "").trim();
        const username = String(r[3] || "").trim();

        if (label && id) {
          rememberOwnerOption({ id, key, label, username });
        }
      });
    });

    ownerOptionsLoaded = getCachedOwnerNamesArray().length > 0;
    return ownerOptionsLoaded;
  } catch (e) {
    console.warn("loadOwnerMapFromMetadataSheet:", e.message || e);
    return false;
  }
}

async function ensureOwnerMap(forceApiLoad = false) {
  if (ownerOptionsLoaded && getCachedOwnerNamesArray().length > 0) return;

  // Ưu tiên đọc cache từ _SYS_OWNER_METADATA để Update Jira không tự tải 13k users.
  const loadedFromSheet = await loadOwnerMapFromMetadataSheet();
  if (loadedFromSheet) return;

  if (!forceApiLoad) {
    throw new Error('Owner list chưa được tải. Hãy bấm "Refresh Metadata" trước khi Update/Create Owner.');
  }

  toast("Đang tải Owner list từ Jira...", "warning");
  await fetchOwnersFromJira();

  if (!ownerOptionsLoaded) {
    throw new Error("Không load được Owner list từ Jira. Kiểm tra objectTypeId Users hoặc quyền Assets API.");
  }
}

async function ensureOwnerMetadataSheet(context) {
  let sheet = context.workbook.worksheets.getItemOrNullObject(OWNER_METADATA_SHEET);
  await context.sync();

  if (sheet.isNullObject) {
    sheet = context.workbook.worksheets.add(OWNER_METADATA_SHEET);
    await context.sync();
  }

  try { sheet.visibility = Excel.SheetVisibility.hidden; } catch (_) {}
  return sheet;
}

async function writeOwnerMetadataSheet(context) {
  const names = getCachedOwnerNamesArray();
  if (!names.length) return null;

  const sheet = await ensureOwnerMetadataSheet(context);
  const used = sheet.getUsedRangeOrNullObject(true);
  await context.sync();

  if (!used.isNullObject) used.clear(Excel.ClearApplyTo.all);

  sheet.getRangeByIndexes(0, 0, 1, 4).values = [["Owner Name", "Owner ID", "Owner Key", "Username"]];

  const rows = names.map(name => {
    const o = getOwnerFromCache(name) || {};
    return [name, o.id || "", o.key || "", o.username || ""];
  });

  sheet.getRangeByIndexes(1, 0, rows.length, 4).values = rows;

  try { sheet.visibility = Excel.SheetVisibility.hidden; } catch (_) {}
  await context.sync();

  return `=${OWNER_METADATA_SHEET}!$A$2:$A$${rows.length + 1}`;
}

async function applyOwnerDropdownToSheet(context, sheet, ownerSourceRangeFormula) {
  if (!ownerSourceRangeFormula) return;

  const assignedRange = sheet.getRangeByIndexes(1, COL.ASSIGNED, 5000, 1);
  assignedRange.dataValidation.rule = {
    list: { inCellDropDown: true, source: ownerSourceRangeFormula },
  };

  assignedRange.dataValidation.errorAlert = {
    showAlert: true,
    style: Excel.DataValidationAlertStyle.warning,
    title: "Owner not in Jira list",
    message: "Nên chọn Owner từ danh sách Jira để update/create đúng reference.",
  };
}

async function applyOwnerDropdownAllSheets() {
  try {
    await ensureOwnerMap();

    await Excel.run(async (context) => {
      const ownerSource = await writeOwnerMetadataSheet(context);
      const sheets = await getLocationSheets(context);

      for (const sheetName of sheets) {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        await applyOwnerDropdownToSheet(context, sheet, ownerSource);
      }

      await context.sync();
    });

    await registerOwnerChangeHandlers();
    toast(`Đã cập nhật dropdown Owner (${getCachedOwnerNamesArray().length} users)`, "success");
  } catch (e) {
    console.warn("applyOwnerDropdownAllSheets:", e.message || e);
    toast("Không cập nhật được dropdown Owner: " + (e.message || e), "warning");
  }
}

async function refreshMetadataDropdowns() {
  // Refresh Metadata chỉ tải Status + Owner và áp dụng dropdown.
  // Không sync asset, không ghi dữ liệu asset.
  try {
    statusOptionsLoaded = false;
    ownerOptionsLoaded = false;

    await ensureStatusMap();
    await ensureOwnerMap(true);

    await Excel.run(async (context) => {
      const sheets = await getLocationSheets(context);
      const ownerSource = await writeOwnerMetadataSheet(context);

      for (const sheetName of sheets) {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        await applyStatusDropdownToSheet(context, sheet);
        await applyOwnerDropdownToSheet(context, sheet, ownerSource);
      }

      await context.sync();
    });

    await registerOwnerChangeHandlers();

    toast(
      `Refresh Metadata xong: Status=${getCachedStatusNamesArray().length}, Owner=${getCachedOwnerNamesArray().length}`,
      "success"
    );
  } catch (e) {
    console.warn("refreshMetadataDropdowns:", e.message || e);
    toast("Refresh Metadata lỗi: " + (e.message || e), "warning");
  }
}


async function syncUsernameForOwnerSelection(context, sheet, rowIndex, ownerName) {
  const owner = getOwnerFromCache(ownerName);
  if (!owner) return false;
  const username = owner.username || owner.key || "";
  if (!username) return false;
  sheet.getRangeByIndexes(rowIndex, COL.USERNAME, 1, 1).values = [[username]];
  return true;
}

async function registerOwnerChangeHandlers() {
  if (ownerChangeHandlersRegistered) return;

  try {
    await Excel.run(async (context) => {
      const sheets = await getLocationSheets(context);

      for (const sheetName of sheets) {
        const sheet = context.workbook.worksheets.getItem(sheetName);

        sheet.onChanged.add(async (event) => {
          try {
            if (!event || !event.address) return;

            await Excel.run(async (ctx) => {
              const ws = ctx.workbook.worksheets.getItem(sheetName);
              const changed = ws.getRange(event.address);
              changed.load(["rowIndex", "columnIndex", "rowCount", "columnCount", "values"]);
              await ctx.sync();

              if (changed.columnIndex > COL.ASSIGNED ||
                  changed.columnIndex + changed.columnCount - 1 < COL.ASSIGNED) {
                return;
              }

              for (let r = 0; r < changed.rowCount; r++) {
                const absoluteRow = changed.rowIndex + r;
                if (absoluteRow === 0) continue;

                const relativeOwnerCol = COL.ASSIGNED - changed.columnIndex;
                const ownerName = changed.values?.[r]?.[relativeOwnerCol] || "";
                await syncUsernameForOwnerSelection(ctx, ws, absoluteRow, ownerName);
              }

              await ctx.sync();
            });
          } catch (e) {
            console.warn("[ownerChange] failed:", e.message || e);
          }
        });
      }

      await context.sync();
      ownerChangeHandlersRegistered = true;
    });
  } catch (e) {
    console.warn("registerOwnerChangeHandlers:", e.message || e);
  }
}

// ══════════════════════════════════════════════════════════════
// AUDIT NOTE: ghi "Standardized by ..." vào cột Note khi user chỉnh row
// ══════════════════════════════════════════════════════════════
function decodeJwtPayload(token) {
  try {
    const part = String(token || "").split(".")[1];
    if (!part) return {};
    const normalized = part.replace(/-/g, "+").replace(/_/g, "/");
    const json = atob(normalized.padEnd(normalized.length + (4 - normalized.length % 4) % 4, "="));
    return JSON.parse(json);
  } catch {
    return {};
  }
}

async function getCurrentUserEmailForAudit() {
  if (cachedCurrentUserEmail) return cachedCurrentUserEmail;

  // Ưu tiên lấy email Microsoft account đang đăng nhập nếu SSO khả dụng.
  try {
    if (OfficeRuntime?.auth?.getAccessToken) {
      const token = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: false });
      const payload = decodeJwtPayload(token);
      const email =
        payload.preferred_username ||
        payload.upn ||
        payload.email ||
        payload.unique_name ||
        "";

      if (email) {
        cachedCurrentUserEmail = String(email).trim();
        return cachedCurrentUserEmail;
      }
    }
  } catch (e) {
    console.warn("[auditNote] SSO user unavailable, fallback to cfg.email:", e.message || e);
  }

  // Fallback: email trong Settings. Nếu team dùng chung token thì giá trị này có thể là email Jira config.
  cachedCurrentUserEmail = String(cfg.email || "unknown-user").trim();
  return cachedCurrentUserEmail;
}

function buildAuditNote(userEmail) {
  const now = new Date().toLocaleString();
  return `Standardized by ${userEmail} at ${now}`;
}

function shouldAuditChangedRange(changed) {
  if (!changed) return false;
  if (changed.rowIndex === 0) return false; // header
  if (isSyncing || isAuditNoteWriting) return false;

  const startCol = changed.columnIndex;
  const endCol = changed.columnIndex + changed.columnCount - 1;

  // Không audit khi chính cột Note/Last Sync/Validation bị ghi để tránh loop.
  if (startCol <= COL.NOTE && endCol >= COL.NOTE) return false;
  if (startCol <= COL.LAST_SYNC && endCol >= COL.LAST_SYNC) return false;
  if (startCol <= COL.VALIDATION && endCol >= COL.VALIDATION) return false;

  // Chỉ audit các thay đổi trong vùng cột của asset.
  return startCol < COL_COUNT;
}

async function registerAuditNoteHandlers() {
  if (auditNoteHandlersRegistered) return;

  try {
    await Excel.run(async (context) => {
      const sheets = await getLocationSheets(context);

      for (const sheetName of sheets) {
        const sheet = context.workbook.worksheets.getItem(sheetName);

        sheet.onChanged.add(async (event) => {
          try {
            // Chỉ ghi audit cho thao tác local của user đang mở workbook.
            // Nếu Excel không trả source thì vẫn xử lý.
            if (event?.source && String(event.source).toLowerCase() !== "local") return;
            if (!event?.address) return;

            await Excel.run(async (ctx) => {
              const ws = ctx.workbook.worksheets.getItem(sheetName);
              const changed = ws.getRange(event.address);
              changed.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
              await ctx.sync();

              if (!shouldAuditChangedRange(changed)) return;

              const userEmail = await getCurrentUserEmailForAudit();
              const noteText = buildAuditNote(userEmail);

              isAuditNoteWriting = true;

              for (let r = 0; r < changed.rowCount; r++) {
                const absoluteRow = changed.rowIndex + r;
                if (absoluteRow === 0) continue;

                // Chỉ ghi note nếu row có dữ liệu asset hoặc LOCAL.
                const rowRange = ws.getRangeByIndexes(absoluteRow, 0, 1, COL_COUNT);
                rowRange.load("values");
                await ctx.sync();

                const row = rowRange.values[0] || [];
                const hasRowData = Boolean(
                  String(row[COL.ASSET_ID] || "").trim() ||
                  String(row[COL.ASSET_KEY] || "").trim() ||
                  String(row[COL.HOSTNAME] || "").trim() ||
                  String(row[COL.SERIAL] || "").trim() ||
                  String(row[COL.SYNC_STATUS] || "").trim().toUpperCase() === "LOCAL"
                );

                if (!hasRowData) continue;

                ws.getRangeByIndexes(absoluteRow, COL.NOTE, 1, 1).values = [[noteText]];
              }

              await ctx.sync();
              isAuditNoteWriting = false;
            });
          } catch (e) {
            isAuditNoteWriting = false;
            console.warn("[auditNote] failed:", e.message || e);
          }
        });
      }

      await context.sync();
      auditNoteHandlersRegistered = true;
    });
  } catch (e) {
    console.warn("registerAuditNoteHandlers:", e.message || e);
  }
}


// ══════════════════════════════════════════════════════════════
// DASHBOARD
// ══════════════════════════════════════════════════════════════
async function refreshDashboard() {
  try {
    await Excel.run(async (context) => {
      const sheets = await getLocationSheets(context);

      let total = 0;
      let pending = 0;
      let mismatch = 0;
      let local = 0;

      const locSummary = [];

      const norm = (v) => String(v || "").trim();
      const upper = (v) => norm(v).toUpperCase();

      function isRealAssetRow(row) {
        const assetId = norm(row[COL.ASSET_ID]);
        const assetKey = norm(row[COL.ASSET_KEY]);
        const hostname = norm(row[COL.HOSTNAME]);
        const serial = norm(row[COL.SERIAL]);
        const syncStatus = upper(row[COL.SYNC_STATUS]);

        // Chỉ tính row asset thật:
        // - có Asset ID / Asset Key / Hostname / Serial
        // - hoặc row LOCAL có nhập dữ liệu
        return Boolean(
          assetId ||
          assetKey ||
          hostname ||
          serial ||
          syncStatus === "LOCAL"
        );
      }

      function isMismatch(row) {
        const validation = norm(row[COL.VALIDATION]);
        if (!validation) return false;

        const v = validation.toUpperCase();

        // OK thì không tính mismatch
        if (v === "OK") return false;

        // Row LOCAL không tính mismatch, vì nó chưa thuộc Jira
        if (upper(row[COL.SYNC_STATUS]) === "LOCAL") return false;

        return true;
      }

      function isPendingTicket(row) {
        const action = upper(row[COL.ACTION]);
        const caseJira = norm(row[COL.CASE_JIRA]);

        return action === "CREATE TICKET" && !caseJira;
      }

      for (const name of sheets) {
        const sheet = context.workbook.worksheets.getItem(name);
        const rows = await readSheetRows(context, sheet);

        let locTotal = 0;
        let locLocal = 0;
        let locMismatch = 0;

        rows.forEach(row => {
          if (!isRealAssetRow(row)) return;

          locTotal++;
          total++;

          if (upper(row[COL.SYNC_STATUS]) === "LOCAL") {
            locLocal++;
            local++;
          }

          if (isMismatch(row)) {
            locMismatch++;
            mismatch++;
          }

          if (isPendingTicket(row)) {
            pending++;
          }
        });

        // Chỉ hiện sheet có asset thật
        if (locTotal > 0) {
          locSummary.push({
            name,
            count: locTotal,
            local: locLocal,
            mismatch: locMismatch,
          });
        }
      }

      // Total Assets = tổng asset Jira từ lần sync gần nhất + LOCAL rows trong Excel.
      // Không dùng mỗi row count vì Excel có thể chưa ghi đủ hoặc row bị bỏ qua bởi filter.
      const jiraTotal = Number(cfg.jiraTotal || 0);
      const dashboardTotal = jiraTotal > 0 ? jiraTotal + local : total;

      setInner("stat-total", dashboardTotal || "0");
      setInner("stat-pending", pending || "0");
      setInner("stat-mismatch", mismatch || "0");
      setInner("stat-local", local || "0");

      if (cfg.lastSync) {
        setInner("last-sync-time", formatTime(cfg.lastSync));
      }

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
            ${l.mismatch > 0 ? `<div class="loc-count">⚠ ${l.mismatch} mismatch</div>` : ""}
          </div>
        `).join("");
      }
    });
  } catch (e) {
    console.warn("refreshDashboard:", e.message);
  }
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
        await hideSystemColumns(context, sheet);
        await applyStatusDropdownToSheet(context, sheet);
      }
    });
    toast(`Created/verified ${locations.length} location sheet(s)`, "success");
    auditNoteHandlersRegistered = false;
    ownerChangeHandlersRegistered = false;
    await registerAuditNoteHandlers();
    await registerOwnerChangeHandlers();
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

    // Lưu tổng Jira asset từ lần sync gần nhất.
    // Dashboard sẽ dùng số này + LOCAL rows để tránh lệch khi Excel chưa ghi đủ row.
    cfg.jiraTotal = allAssets.length;
    Office.context.document.settings.set(CFG_KEYS.JIRA_TOTAL, String(cfg.jiraTotal));

    const allKnownIds = new Set(
      allAssets
        .map(a => String(a.id || "").trim())
        .filter(Boolean)
    );

    const allAssetById = new Map();
    allAssets.forEach(a => {
      const id = String(a.id || "").trim();
      if (id) allAssetById.set(id, a);
    });

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
        allKnownIds,
        allAssetById
      );
    }

    if (syncPanel) {
      updateSyncPanel(syncPanel, allSheetNames.map(s => ({
        name: s.replace(/^_/, ""),
        count: (byLocation[s] || []).length,
        status: "done",
      })));
    }

    // Full Sync chỉ sync asset + áp lại Status dropdown nhẹ.
    // Owner list >1000 records nên KHÔNG load trong Full Sync.
    // Muốn cập nhật Owner dropdown thì bấm nút "Refresh Metadata".
    await applyStatusDropdownAllSheets();

    // Không tự động tạo LOCAL asset khi Full Sync.
    // Tạo mới chỉ được thực hiện thủ công bằng Action = "+"
    // và chỉ khi SYNC_STATUS = LOCAL + Asset ID trống.
    cfg.lastSync = new Date().toISOString();
    Office.context.document.settings.set(CFG_KEYS.LAST_SYNC, cfg.lastSync);
    Office.context.document.settings.set(CFG_KEYS.JIRA_TOTAL, String(cfg.jiraTotal || allAssets.length || 0));
    Office.context.document.settings.saveAsync();

    setSyncIndicator("ok", "Synced");
    setInner("last-sync-time", formatTime(cfg.lastSync));

    toast(`Sync hoàn tất — ${allAssets.length} assets`, "success");
    auditNoteHandlersRegistered = false;
    ownerChangeHandlersRegistered = false;
    await registerAuditNoteHandlers();
    await registerOwnerChangeHandlers();
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
//   x = UPDATE Jira asset hiện có, bắt buộc có Asset ID
//   + = CREATE mới Jira asset, chỉ khi SYNC_STATUS = LOCAL và Asset ID trống
//   o = xóa object trên Jira nếu có Asset ID, sau đó xóa row khỏi Excel sheet
//
// Create/Update fields gửi lên Jira:
//   Hostname, Serial, Location, Status, Purchase Date
// ══════════════════════════════════════════════════════════════
function normalizeAction(value) {
  return String(value || "").trim().toLowerCase();
}

function isLocalRow(row) {
  return String(row[COL.SYNC_STATUS] || "").trim().toUpperCase() === "LOCAL";
}

function rowToJiraFields(row) {
  return {
    hostname:     String(row[COL.HOSTNAME]     || "").trim(),
    serial:       String(row[COL.SERIAL]       || "").trim(),
    status:       String(row[COL.STATUS]       || "").trim(),
    location:     String(row[COL.LOCATION]     || "").trim(),
    region:       String(row[COL.REGION]       || "").trim(),
    manufacturer: String(row[COL.MANUFACTURER] || "").trim(),
    model:        String(row[COL.MODEL]        || "").trim(),
    os:           String(row[COL.OS]           || "").trim(),
    osVersion:    String(row[COL.OS_VERSION]   || "").trim(),
    osBuild:      String(row[COL.OS_BUILD]     || "").trim(),
    cpu:          String(row[COL.CPU]          || "").trim(),
    ip:           String(row[COL.IP]           || "").trim(),
    mac:          String(row[COL.MAC]          || "").trim(),
    network:      String(row[COL.NETWORK]      || "").trim(),
    antivirus:    String(row[COL.ANTIVIRUS]    || "").trim(),
    username:     String(row[COL.USERNAME]     || "").trim(),
    owner:        String(row[COL.ASSIGNED]     || "").trim(),
    firstSeen:    normalizeJiraDateTime(row[COL.FIRST_SEEN]),
    lastSeen:     normalizeJiraDateTime(row[COL.LAST_SEEN]),
    purchase:     normalizeJiraDate(row[COL.PURCHASE]),
    warranty:     normalizeJiraDate(row[COL.WARRANTY]),
    tenantId:     String(row[COL.TENANT_ID]    || "").trim(),
    lansweeper:   String(row[COL.LANSWEEPER]   || "").trim(),
  };
}

function normalizeJiraDate(value) {
  if (!value) return "";

  // Excel date serial number -> YYYY-MM-DD
  if (typeof value === "number") {
    const d = new Date(Math.round((value - 25569) * 86400 * 1000));
    if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  }

  const raw = String(value || "").trim();
  if (!raw) return "";

  // Đã đúng format Jira date
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;

  const d = new Date(raw);
  if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);

  return raw;
}

function normalizeJiraDateTime(value) {
  if (!value) return "";

  // Excel date serial number -> ISO datetime
  if (typeof value === "number") {
    const d = new Date(Math.round((value - 25569) * 86400 * 1000));
    if (!isNaN(d.getTime())) return d.toISOString();
  }

  const raw = String(value || "").trim();
  if (!raw) return "";

  // Đã là ISO datetime
  if (/^\d{4}-\d{2}-\d{2}T/.test(raw)) return raw;

  const d = new Date(raw);
  if (!isNaN(d.getTime())) return d.toISOString();

  // Nếu Jira đang nhận display string kiểu "24/Jun/25 2:14 AM" thì giữ nguyên.
  return raw;
}


function ingestStatusResponse(data) {
  const list = Array.isArray(data)
    ? data
    : Array.isArray(data?.values)
      ? data.values
      : Array.isArray(data?.statusTypes)
        ? data.statusTypes
        : Array.isArray(data?.statuses)
          ? data.statuses
          : [];

  let count = 0;
  list.forEach(x => {
    const id = x?.id ?? x?.statusId ?? x?.globalId;
    const name = x?.name ?? x?.label ?? x?.displayValue;
    if (id && name) {
      rememberStatusOption(name, id);
      count++;
    }
  });
  return count;
}

async function fetchObjectSchemaIdsFromTypeIds() {
  const schemaIds = new Set();
  const typeIds = parseTypeIds();

  for (const typeId of typeIds) {
    try {
      const ot = await assetsGet(`/objecttype/${encodeURIComponent(typeId)}`);
      const schemaId = ot?.objectSchemaId || ot?.schemaId;
      if (schemaId) schemaIds.add(String(schemaId));
    } catch (e) {
      console.warn(`[statusMap] cannot read objecttype ${typeId}:`, e.message || e);
    }
  }

  return [...schemaIds];
}


function extractArrayFromMaybeResponse(data) {
  if (Array.isArray(data)) return data;
  if (Array.isArray(data?.values)) return data.values;
  if (Array.isArray(data?.items)) return data.items;
  if (Array.isArray(data?.attributes)) return data.attributes;
  if (Array.isArray(data?.objectTypeAttributes)) return data.objectTypeAttributes;
  return [];
}

function extractStatusIdsFromAttribute(attr) {
  const ids = new Set();

  const add = (v) => {
    const s = String(v ?? "").trim();
    if (s) ids.add(s);
  };

  if (!attr) return ids;

  // Jira Assets status attribute thường chứa allowed status IDs tại typeValueMulti.
  if (Array.isArray(attr.typeValueMulti)) {
    attr.typeValueMulti.forEach(add);
  }

  if (Array.isArray(attr.typeValues)) {
    attr.typeValues.forEach(x => add(x?.id ?? x?.statusId ?? x));
  }

  if (Array.isArray(attr.options)) {
    attr.options.forEach(x => add(x?.id ?? x?.statusId ?? x));
  }

  if (attr.typeValue) {
    add(attr.typeValue?.id ?? attr.typeValue?.statusId ?? attr.typeValue);
  }

  if (attr.defaultType) {
    add(attr.defaultType?.id ?? attr.defaultType?.statusId);
  }

  return ids;
}

async function loadStatusTypesFromObjectTypeAttributes() {
  const typeIds = parseTypeIds();
  const statusIds = new Set();

  for (const typeId of typeIds) {
    const candidates = [
      `/objecttype/${encodeURIComponent(typeId)}/attributes`,
      `/objecttype/${encodeURIComponent(typeId)}/attributes?includeChildren=true`,
    ];

    for (const path of candidates) {
      try {
        const data = await assetsGet(path);
        const attrs = extractArrayFromMaybeResponse(data);

        attrs.forEach(attr => {
          const attrId = String(attr?.id ?? attr?.objectTypeAttributeId ?? "").trim();
          const attrName = String(attr?.name ?? attr?.label ?? "").trim().toLowerCase();

          // Ưu tiên đúng attribute ID 5052, fallback theo tên Status.
          if (attrId === String(STATUS_ATTR_ID) || attrName === "status") {
            extractStatusIdsFromAttribute(attr).forEach(id => statusIds.add(id));
          }
        });

        if (statusIds.size > 0) break;
      } catch (e) {
        console.warn(`[statusMap] cannot read ${path}:`, e.message || e);
      }
    }
  }

  let added = 0;

  for (const id of statusIds) {
    const candidates = [
      `/config/statustype/${encodeURIComponent(id)}`,
      `/config/status/${encodeURIComponent(id)}`,
    ];

    for (const path of candidates) {
      try {
        const data = await assetsGet(path);
        const before = Object.keys(statusNameToId).length;
        ingestStatusResponse([data]);
        const after = Object.keys(statusNameToId).length;

        if (after > before) added++;
        break;
      } catch (e) {
        console.warn(`[statusMap] cannot load status id=${id} by ${path}:`, e.message || e);
      }
    }
  }

  console.log(`[statusMap] attribute config IDs=${[...statusIds].join(",")} added=${added}`);
  return added;
}


async function loadStatusTypesFromApi() {
  let total = 0;

  // Cách chuẩn nhất cho dropdown Status:
  // đọc object type attribute 5052 để lấy allowed status IDs trong typeValueMulti,
  // sau đó gọi /config/statustype/{id}. Cách này lấy đủ cả status chưa asset nào đang dùng.
  total += await loadStatusTypesFromObjectTypeAttributes();

  const schemaIds = await fetchObjectSchemaIdsFromTypeIds();

  for (const schemaId of schemaIds) {
    const candidates = [
      `/config/statustype?objectSchemaId=${encodeURIComponent(schemaId)}`,
      `/config/statustype?objectSchemaId=${encodeURIComponent(schemaId)}&maxResults=100`,
      `/config/status?objectSchemaId=${encodeURIComponent(schemaId)}`,
      `/config/status?objectSchemaId=${encodeURIComponent(schemaId)}&maxResults=100`,
    ];

    for (const path of candidates) {
      try {
        const data = await assetsGet(path);
        const added = ingestStatusResponse(data);
        total += added;
        if (added > 0) break;
      } catch (e) {
        console.warn(`[statusMap] ${path}:`, e.message || e);
      }
    }
  }

  if (total === 0) {
    const candidates = ["/config/statustype", "/config/status"];
    for (const path of candidates) {
      try {
        const data = await assetsGet(path);
        const added = ingestStatusResponse(data);
        total += added;
        if (added > 0) break;
      } catch (e) {
        console.warn(`[statusMap] ${path}:`, e.message || e);
      }
    }
  }

  return total;
}

async function ensureStatusMap() {
  if (statusOptionsLoaded && Object.keys(statusNameToId).length > 0) return;

  toast("Đang tải Status list từ Jira...", "warning");

  // Ưu tiên lấy từ API config/schema để có đủ cả status chưa asset nào đang dùng.
  await loadStatusTypesFromApi();

  // Fallback: parse từ asset đang sync. Cách này chỉ lấy được status đã được dùng trong asset.
  // Nếu API/config chỉ trả một phần, fetchJiraAssets() vẫn có thể bổ sung thêm status đang có trong dữ liệu.
  if (Object.keys(statusNameToId).length < 3) {
    await fetchJiraAssets();
  }

  statusOptionsLoaded = Object.keys(statusNameToId).length > 0;

  if (!statusOptionsLoaded) {
    throw new Error("Không load được Status map từ Jira. Hãy Sync trước hoặc kiểm tra quyền Assets API.");
  }

  console.log(`[statusMap] loaded ${getCachedStatusNamesArray().length} status: ${listCachedStatusNames()}`);
}

async function getObjectTypeIdForAsset(assetId) {
  const id = String(assetId || "").trim();
  if (!id) return "";

  try {
    const data = await assetsGet(`/object/${encodeURIComponent(id)}`);
    return String(
      data?.objectType?.id ||
      data?.objectTypeId ||
      data?.objectType?.objectTypeId ||
      ""
    ).trim();
  } catch (e) {
    console.warn(`[schema] cannot read object ${id}:`, e.message || e);
    return "";
  }
}

async function getAllowedAttributeIdsForObjectType(objectTypeId) {
  const typeId = String(objectTypeId || "").trim();
  if (!typeId) return null;

  if (objectTypeAttributeCache.has(typeId)) {
    return objectTypeAttributeCache.get(typeId);
  }

  try {
    const data = await assetsGet(`/objecttype/${encodeURIComponent(typeId)}/attributes`);

    const attrs = Array.isArray(data)
      ? data
      : Array.isArray(data?.values)
        ? data.values
        : Array.isArray(data?.attributes)
          ? data.attributes
          : Array.isArray(data?.objectTypeAttributes)
            ? data.objectTypeAttributes
            : [];

    const ids = new Set(
      attrs
        .map(a => String(a?.id ?? a?.objectTypeAttributeId ?? "").trim())
        .filter(Boolean)
    );

    objectTypeAttributeCache.set(typeId, ids);
    return ids;
  } catch (e) {
    console.warn(`[schema] cannot read attributes for objectTypeId=${typeId}:`, e.message || e);
    return null;
  }
}

function filterAttributesByAllowedIds(attributes, allowedIds) {
  if (!(allowedIds instanceof Set)) return attributes;

  return attributes.filter(a =>
    allowedIds.has(String(a?.objectTypeAttributeId || "").trim())
  );
}

async function jiraAttributesFromFields(fields) {
  const attributes = [];

  const addValue = (objectTypeAttributeId, value) => {
    const v = String(value || "").trim();
    if (!v) return;

    attributes.push({
      objectTypeAttributeId: Number(objectTypeAttributeId),
      objectAttributeValues: [{ value: v }],
    });
  };

  // Text/value attributes theo schema object hiện tại
  addValue(1737,  fields.hostname);      // Hostname / Name
  addValue(5194,  fields.serial);        // Serial Number
  addValue(30125, fields.location);      // Location
  addValue(27292, fields.region);        // Region
  addValue(6608,  fields.manufacturer);  // Manufacturer
  addValue(6609,  fields.model);         // Model
  addValue(30345, fields.os);            // Operating System
  addValue(27291, fields.osVersion);     // Windows Version
  addValue(27290, fields.osBuild);       // Windows Build
  addValue(6610,  fields.cpu);           // CPU
  addValue(5208,  fields.ip);            // IP Address
  addValue(5209,  fields.mac);           // MAC Address
  addValue(5210,  fields.network);       // Network Name
  addValue(6612,  fields.antivirus);     // Antivirus
  addValue(5200,  fields.username);      // Username
  addValue(5205,  fields.firstSeen);     // First Seen
  addValue(5206,  fields.lastSeen);      // Last Seen
  addValue(5203,  fields.purchase);      // Purchase Date
  addValue(6615,  fields.warranty);      // Warranty Expire
  addValue(26398, fields.tenantId);      // Tenant ID / Source ID
  addValue(5207,  fields.lansweeper);    // Lansweeper URL

  // Status 5052 là status type, không gửi text. Phải map name -> status.id.
  if (fields.status) {
    await ensureStatusMap();
    const statusId = getStatusIdFromCache(fields.status);

    if (!statusId) {
      throw new Error(
        `Status "${fields.status}" không hợp lệ. Status hợp lệ đang cache: ${listCachedStatusNames() || "chưa có"}`
      );
    }

    attributes.push({
      objectTypeAttributeId: Number(STATUS_ATTR_ID),
      objectAttributeValues: [{ value: String(statusId) }],
    });
  }

  // Assigned User / Owner 26690 là reference object.
  // Payload đúng: { referencedObjectBeanId: 82854 }
  if (fields.owner) {
    await ensureOwnerMap(false);
    const owner = getOwnerFromCache(fields.owner);

    if (!owner?.id) {
      throw new Error(`Owner "${fields.owner}" không có trong cache Owner. Hãy bấm "Refresh Metadata" rồi chọn lại Owner từ dropdown.`);
    }

    attributes.push({
      objectTypeAttributeId: Number(OWNER_ATTR_ID),
      objectAttributeValues: [{ referencedObjectBeanId: Number(owner.id) }],
    });
  }

  return attributes;
}

async function updateJiraAsset(assetId, fields) {
  const id = String(assetId || "").trim();
  if (!id) throw new Error("Action x cần có Asset ID để update");

  let attributes = await jiraAttributesFromFields(fields);
  if (!attributes.length) throw new Error("Không có field nào để update");

  // Object type nào không có attribute nào thì bỏ attribute đó.
  // Fix lỗi: Object Type Attribute not valid (id: 26690).
  const objectTypeId = await getObjectTypeIdForAsset(id);
  const allowedIds = await getAllowedAttributeIdsForObjectType(objectTypeId);
  const before = attributes.length;

  attributes = filterAttributesByAllowedIds(attributes, allowedIds);

  const skipped = before - attributes.length;
  if (skipped > 0) {
    console.warn(`[updateJiraAsset] skipped ${skipped} invalid attribute(s) for asset=${id}, objectTypeId=${objectTypeId}`);
  }

  if (!attributes.length) {
    throw new Error(`Không có field hợp lệ để update cho objectTypeId=${objectTypeId || "unknown"}`);
  }

  try {
    return await assetsPut(`/object/${id}`, { attributes });
  } catch (e) {
    // Fallback: nếu Jira vẫn báo attribute không hợp lệ, bỏ attr đó và thử lại.
    const msg = String(e?.message || e || "");
    const m = msg.match(/Object Type Attribute not valid \(id:\s*(\d+)\)/i);

    if (m && m[1]) {
      const badId = String(m[1]);
      const filtered = attributes.filter(a => String(a.objectTypeAttributeId) !== badId);

      if (filtered.length && filtered.length < attributes.length) {
        console.warn(`[updateJiraAsset] retry without invalid attr ${badId}`);
        return await assetsPut(`/object/${id}`, { attributes: filtered });
      }
    }

    throw e;
  }
}

function escapeAqlString(value) {
  return String(value || "").replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}

function normalizeSerial(value) {
  return String(value || "").trim().toUpperCase();
}

function baseAqlForSerialSearch() {
  const raw = String(cfg.aqlQuery || "").trim();
  if (raw) return `(${raw})`;

  const typeIds = parseTypeIds();
  if (typeIds.length) return `objectTypeId IN (${typeIds.join(",")})`;

  throw new Error("AQL Query chưa cấu hình, không thể kiểm tra trùng Serial Number");
}

async function findJiraAssetBySerial(serial) {
  const ser = String(serial || "").trim();
  if (!ser) return null;

  const qlQuery = `${baseAqlForSerialSearch()} AND "Serial Number" = "${escapeAqlString(ser)}"`;
  const data = await fetchPage(qlQuery, 0, 10);
  const values = data?.values || [];

  if (!values.length) return null;

  const exact = values
    .map(o => parseAsset(o))
    .find(a => normalizeSerial(a.serial) === normalizeSerial(ser));

  return exact || parseAsset(values[0]);
}

async function createOrUpdateJiraAssetFromLocalRow(row) {
  const assetId = String(row[COL.ASSET_ID] || "").trim();

  if (assetId) {
    throw new Error("Action + chỉ dùng cho row chưa có Asset ID");
  }

  if (!isLocalRow(row)) {
    throw new Error('Action + chỉ xử lý khi Sync Status = "LOCAL"');
  }

  const fields = rowToJiraFields(row);
  const serial = String(fields.serial || "").trim();

  if (!serial) {
    throw new Error("Action + cần Serial Number để kiểm tra trùng trước khi tạo mới");
  }

  // Kiểm tra toàn bộ asset theo AQL đang cấu hình.
  // Nếu Serial đã tồn tại thì UPDATE object đó, không CREATE duplicate.
  const existing = await findJiraAssetBySerial(serial);

  if (existing?.id) {
    await updateJiraAsset(existing.id, fields);
    return {
      mode: "updatedExisting",
      asset: existing,
      response: existing,
    };
  }

  const res = await createJiraAssetFromRow(row);
  return {
    mode: "created",
    asset: null,
    response: res,
  };
}

async function createJiraAssetFromRow(row) {
  const assetId = String(row[COL.ASSET_ID] || "").trim();

  if (assetId) {
    throw new Error("Action + chỉ dùng cho row chưa có Asset ID");
  }

  if (!isLocalRow(row)) {
    throw new Error('Action + chỉ tạo mới khi Sync Status = "LOCAL"');
  }

  const typeIds = parseTypeIds();
  if (!typeIds.length) {
    throw new Error("Không tìm thấy objectTypeId trong AQL Query");
  }

  const defaultTypeId = Number(typeIds[0]);
  if (!defaultTypeId) {
    throw new Error("objectTypeId không hợp lệ");
  }

  const fields = rowToJiraFields(row);

  if (!fields.hostname && !fields.serial) {
    throw new Error("Row mới cần ít nhất Hostname hoặc Serial Number");
  }

  let attributes = await jiraAttributesFromFields(fields);

  // Chỉ gửi attribute hợp lệ cho objectType create mặc định.
  const allowedIds = await getAllowedAttributeIdsForObjectType(defaultTypeId);
  attributes = filterAttributesByAllowedIds(attributes, allowedIds);

  if (!attributes.length) {
    throw new Error("Không có field hợp lệ để create");
  }

  return assetsPost("/object/create", {
    objectTypeId: defaultTypeId,
    attributes,
  });
}

async function deleteJiraAsset(assetId) {
  const id = String(assetId || "").trim();
  if (!id) throw new Error("Missing Asset ID để xóa Jira asset");

  return assetsDelete(`/object/${id}`);
}

async function processActionRows() {
  if (!cfg.jiraUrl || !cfg.token || !cfg.workerUrl || !cfg.cloudId || !cfg.workspaceId) {
    toast("Kiểm tra Settings trước khi update Jira", "warning");
    return;
  }

  toast('Đang xử lý Action trong Excel: x=update, +=create/match serial, o=delete. Không full sync asset.', "warning");

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
          const action = normalizeAction(row[COL.ACTION]);
          const excelRow = i + 1; // 0 là header, data row đầu tiên là index 1

          if (!action) continue;

          // o = xóa Jira asset trước nếu có Asset ID, sau đó xóa row khỏi Excel
          if (action === "o") {
            const assetId = String(row[COL.ASSET_ID] || "").trim();

            try {
              if (assetId) {
                await deleteJiraAsset(assetId);
              }

              rowsToDelete.push(excelRow);
              deleted++;
            } catch (e) {
              failed++;
              row[COL.VALIDATION] = "Delete failed: " + String(e.message || e).slice(0, 120);
              sheet.getRangeByIndexes(excelRow, 0, 1, COL_COUNT).values = [row];
              await context.sync();
              console.warn(`[processActionRows] delete ${sheetName} row ${i + 2}:`, e.message || e);
            }

            continue;
          }

          const assetId = String(row[COL.ASSET_ID] || "").trim();

          try {
            if (action === "x") {
              // UPDATE chỉ cho asset đã có trên Jira
              if (!assetId) {
                throw new Error('Action x chỉ update row đã có Asset ID. Muốn tạo mới hãy nhập Action = "+"');
              }

              const fields = rowToJiraFields(row);
              const owner = getOwnerFromCache(fields.owner);
              if (owner?.username) row[COL.USERNAME] = owner.username;

              await updateJiraAsset(assetId, fields);

              row[COL.SYNC_STATUS] = "JIRA";
              row[COL.VALIDATION] = "OK";
              row[COL.LAST_SYNC] = new Date().toISOString();
              row[COL.ACTION] = "";

              sheet.getRangeByIndexes(excelRow, 0, 1, COL_COUNT).values = [row];
              updated++;
              await context.sync();
              continue;
            }

            if (action === "+") {
              // CREATE chỉ khi LOCAL + chưa có Asset ID
              if (assetId) {
                throw new Error('Action + chỉ dùng để tạo mới row chưa có Asset ID');
              }

              if (!isLocalRow(row)) {
                throw new Error('Action + chỉ tạo mới khi Sync Status = LOCAL');
              }

              const localFields = rowToJiraFields(row);
              const owner = getOwnerFromCache(localFields.owner);
              if (owner?.username) row[COL.USERNAME] = owner.username;

              const result = await createOrUpdateJiraAssetFromLocalRow(row);
              const res = result.response || {};
              const existingAsset = result.asset || {};

              row[COL.ASSET_ID] = String(res?.id || res?.objectId || existingAsset.id || "");
              row[COL.ASSET_KEY] = String(res?.objectKey || res?.key || existingAsset.key || "");
              row[COL.SYNC_STATUS] = "JIRA";
              row[COL.VALIDATION] = result.mode === "updatedExisting" ? "OK - Matched by Serial" : "OK";
              row[COL.LAST_SYNC] = new Date().toISOString();
              row[COL.ACTION] = "";

              sheet.getRangeByIndexes(excelRow, 0, 1, COL_COUNT).values = [row];

              if (result.mode === "updatedExisting") updated++;
              else created++;

              await context.sync();
              continue;
            }

            skipped++;
          } catch (e) {
            failed++;
            row[COL.VALIDATION] = "Action failed: " + String(e.message || e).slice(0, 120);
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