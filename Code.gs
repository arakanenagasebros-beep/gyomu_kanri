var DATA_FILE_NAME = "業務管理アプリ_data.json";
var FILES_FOLDER_NAME = "業務管理アプリ_files";
var FILES_ROOT_FOLDER_ID = "1-eGu7Sti6TrcM-U8fBJgcr1OfwZmL1A8";
var COMPLETED_FILES_FOLDER_NAME = "完了済";

var TOKEN_TTL_SECONDS = 6 * 60 * 60;
var DATA_CACHE_TTL = 600;
var DATA_CACHE_KEY = "DATA_JSON_V2";
var DATA_CACHE_CHUNK = 90 * 1024;
var LOGIN_RATE_LIMIT = 5;
var LOGIN_RATE_WINDOW = 5 * 60;
var PROP = PropertiesService.getScriptProperties();

// 初期プロファイル（PW は含めない / GitHub 公開でも安全）
// PW は管理者が UI / bootstrapAdmin で設定する
var DEFAULT_STAFF_PROFILES = {
  shakai_test: { name: "テスト社会人", userType: "社会人" },
  ogasawara: { name: "小笠原", userType: "社会人" },
  morotomi: { name: "諸富", userType: "社会人" },
  osawa: { name: "大澤", userType: "社会人" },
  yoneoka: { name: "米岡", userType: "社会人" },
  hosaka: { name: "保坂", userType: "社会人" },
  gakusei_test: { name: "テスト学生", userType: "学生" },
  miko: { name: "神子", userType: "学生" },
  shirakawa: { name: "白川", userType: "学生" },
  matsumoto: { name: "松本", userType: "学生" },
  mizutani: { name: "水谷", userType: "学生" },
  takeuchi: { name: "竹内", userType: "学生" },
  fujikawa: { name: "藤川", userType: "学生" },
  kobayashi: { name: "小林", userType: "学生" }
};
var DEFAULT_USER_HOURLY_RATES = {
  shakai_test: 1300, ogasawara: 1300, morotomi: 1300, osawa: 1300, yoneoka: 1300,
  hosaka: 1300, gakusei_test: 1300, miko: 1300, shirakawa: 1300, matsumoto: 1300,
  mizutani: 1300, takeuchi: 1300, fujikawa: 1300, kobayashi: 1300
};
var DEFAULT_DAILY_PASSWORDS_LEN = 34;

// ====== 一回だけ実行する初期化系 ======
// GAS エディタから一度だけ実行：管理者の ID/PW を設定
function bootstrapAdmin(newId, newPw) {
  if (!newId || !newPw) throw new Error("usage: bootstrapAdmin(id, pw)");
  setAdminCredsHashed_(String(newId), String(newPw));
  Logger.log("管理者を設定しました: " + newId);
}
// GAS エディタから：合言葉リストを設定（カンマ区切り or 配列）
function bootstrapDailyPasswords(list) {
  var arr = Array.isArray(list) ? list : String(list || "").split(/[,\s]+/).filter(Boolean);
  if (!arr.length) throw new Error("list required");
  saveJsonProperty_("DAILY_PASSWORDS_JSON", arr);
  Logger.log("合言葉を保存しました（" + arr.length + "件）");
}
// 既存の平文 PW を一括ハッシュ化（既にハッシュなら何もしない）
function migrateAllStaffPasswordsToHash() {
  var staff = loadJsonProperty_("STAFF_ACCOUNTS_JSON", {});
  var changed = 0;
  Object.keys(staff || {}).forEach(function(id) {
    var acc = staff[id] || {};
    if (acc.pwHash && acc.salt) return;
    if (acc.pw) {
      var rec = makePasswordRecord_(acc.pw);
      delete acc.pw;
      acc.salt = rec.salt;
      acc.pwHash = rec.pwHash;
      staff[id] = acc;
      changed++;
    }
  });
  if (changed) saveJsonProperty_("STAFF_ACCOUNTS_JSON", staff);
  Logger.log("ハッシュ化: " + changed + " 件");
}

// 復元用：Drive 上のファイルから読み直し、_version を Script Property に同期 + キャッシュ破棄
function syncVersionFromDriveData() {
  var data = readData_();
  var v = parseInt((data && data._version) || 0, 10);
  var ua = (data && data._updatedAt) || new Date().toISOString();
  PROP.setProperty("DATA_VERSION", String(v));
  PROP.setProperty("DATA_UPDATED_AT", ua);
  // データキャッシュも破棄（壊れた版がキャッシュされている可能性を排除）
  try {
    var cache = CacheService.getScriptCache();
    cache.remove(DATA_CACHE_KEY);
    for (var i = 0; i < 20; i++) cache.remove(DATA_CACHE_KEY + ":" + i);
  } catch (e) {}
  var users = (data && data.users) || {};
  var totalReports = 0, totalStamps = 0;
  Object.keys(users).forEach(function(uid) {
    var u = users[uid] || {};
    totalReports += (u.reports || []).length;
    totalStamps += Object.keys(u.stamps || {}).length;
  });
  Logger.log("同期完了: _version=" + v + " / _updatedAt=" + ua);
  Logger.log("reports合計=" + totalReports + " / stamps合計=" + totalStamps);
}

// GAS エディタから実行：スタッフ別のスタンプ件数をログ表示（読取専用・診断用）
function inspectStaffStamps() {
  var data = readData_();
  var users = (data && data.users) || {};
  var ids = Object.keys(users).sort();
  Logger.log("=== スタッフ別 スタンプ件数 ===");
  Logger.log("総ユーザー数: " + ids.length);
  ids.forEach(function(id) {
    var u = users[id] || {};
    var stamps = u.stamps || {};
    var keys = Object.keys(stamps);
    var sample = keys.slice(0, 5).map(function(k) { return k + "=" + stamps[k]; }).join(", ");
    Logger.log(
      id + " (" + (u.name || "") + ") : " +
      keys.length + " 件" +
      (keys.length ? " | 例: " + sample : "")
    );
  });
  var v = PROP.getProperty("DATA_VERSION") || "0";
  var ua = PROP.getProperty("DATA_UPDATED_AT") || "(unknown)";
  Logger.log("DATA_VERSION = " + v + " / DATA_UPDATED_AT = " + ua);
}

// ====== 共通ユーティリティ ======
function loadJsonProperty_(key, fallback) {
  var raw = PROP.getProperty(key);
  if (!raw) return fallback;
  try { return JSON.parse(raw); } catch (e) { return fallback; }
}
function saveJsonProperty_(key, obj) { PROP.setProperty(key, JSON.stringify(obj)); }

// ====== パスワードハッシュ ======
function bytesToHex_(bytes) {
  var s = "";
  for (var i = 0; i < bytes.length; i++) {
    var b = bytes[i] < 0 ? bytes[i] + 256 : bytes[i];
    s += ("0" + b.toString(16)).slice(-2);
  }
  return s;
}
function hashPassword_(salt, pw) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(salt) + ":" + String(pw),
    Utilities.Charset.UTF_8
  );
  return bytesToHex_(bytes);
}
function makePasswordRecord_(pw) {
  var salt = Utilities.getUuid();
  return { salt: salt, pwHash: hashPassword_(salt, pw) };
}
function verifyPasswordRecord_(rec, pw) {
  if (!rec) return false;
  if (rec.pwHash && rec.salt) return hashPassword_(rec.salt, pw) === rec.pwHash;
  if (rec.pw != null) return String(rec.pw) === String(pw);
  return false;
}

// ====== ログインレート制限 ======
function getLoginAttemptKey_(scope, key) { return "loginAttempts:" + scope + ":" + key; }
function checkLoginRateLimit_(scope, key) {
  var cache = CacheService.getScriptCache();
  var ck = getLoginAttemptKey_(scope, key);
  var cnt = parseInt(cache.get(ck) || "0", 10);
  if (cnt >= LOGIN_RATE_LIMIT) return false;
  return true;
}
function recordLoginFailure_(scope, key) {
  var cache = CacheService.getScriptCache();
  var ck = getLoginAttemptKey_(scope, key);
  var cnt = parseInt(cache.get(ck) || "0", 10) + 1;
  cache.put(ck, String(cnt), LOGIN_RATE_WINDOW);
}
function clearLoginFailures_(scope, key) {
  CacheService.getScriptCache().remove(getLoginAttemptKey_(scope, key));
}

// ====== Staff アカウント / プロファイル ======
function getStaffAccounts_() {
  var obj = loadJsonProperty_("STAFF_ACCOUNTS_JSON", null);
  if (!obj || typeof obj !== "object") {
    // 初回：プロファイルだけ用意（PW は未設定）
    return seedStaffAccountsFromProfiles_();
  }
  return obj;
}
function seedStaffAccountsFromProfiles_() {
  var seed = {};
  Object.keys(DEFAULT_STAFF_PROFILES).forEach(function(id) {
    var p = DEFAULT_STAFF_PROFILES[id];
    seed[id] = { name: p.name || id, userType: p.userType || "学生" };
  });
  saveJsonProperty_("STAFF_ACCOUNTS_JSON", seed);
  return seed;
}
function getDailyPasswords_() {
  var arr = loadJsonProperty_("DAILY_PASSWORDS_JSON", null);
  return Array.isArray(arr) && arr.length ? arr : [];
}
function getUserDisplayName_(user) {
  return user ? (user.name || user.id || "") : "";
}
function getTaskAssignedUserId_(task, data) {
  if (!task) return "";
  if (task.staffUserId) return String(task.staffUserId);
  var ref = String(task.staff || "");
  if (!ref) return "";
  var users = (data && data.users) || {};
  if (users[ref]) return ref;
  for (var uid in users) {
    if (!Object.prototype.hasOwnProperty.call(users, uid)) continue;
    var u = users[uid];
    if (u && getUserDisplayName_(u) === ref) return uid;
  }
  return "";
}
function isTaskAccessibleByAuth_(task, auth, data) {
  if (!auth) return false;
  if (auth.role === "admin") return true;
  return getTaskAssignedUserId_(task, data) === auth.userId;
}
function isWorkTypeChangeAllowed_(task) {
  return task && (task.status === "依頼中" || task.status === "期限超過");
}
function buildStaffWorkTypeChangeRequest_(existingTask, incomingReq, auth) {
  if (!incomingReq) return null;
  if (!isWorkTypeChangeAllowed_(existingTask)) {
    throw new Error("workType change is only allowed while task is in progress");
  }
  var requested = String(incomingReq.requestedWorkType || "").trim();
  if (requested !== "出勤" && requested !== "在宅") {
    throw new Error("invalid requested workType");
  }
  var current = String(existingTask.workType || "");
  if (!current || requested === current) {
    throw new Error("requested workType must differ from current workType");
  }
  return {
    status: "pending",
    currentWorkType: current,
    requestedWorkType: requested,
    requestedAt: Date.now(),
    requestedBy: auth && auth.userId ? String(auth.userId) : ""
  };
}
function sanitizeUserForStaffRead_(user, isSelf) {
  if (!user) return null;
  if (!isSelf) {
    return {
      id: user.id || "",
      name: user.name || "",
      userType: user.userType || "",
      createdAt: user.createdAt || null
    };
  }
  var clean = JSON.parse(JSON.stringify(user));
  delete clean.pw; delete clean.pwHash; delete clean.salt;
  return clean;
}
function filterDataForAuth_(data, auth) {
  if (!auth || auth.role === "admin") return data;
  var filtered = JSON.parse(JSON.stringify(data || {}));
  var users = filtered.users || {};
  var nextUsers = {};
  Object.keys(users).forEach(function(uid) {
    var sanitized = sanitizeUserForStaffRead_(users[uid], uid === auth.userId);
    if (sanitized) nextUsers[uid] = sanitized;
  });
  filtered.users = nextUsers;
  filtered.tasks = (filtered.tasks || []).filter(function(task) {
    return isTaskAccessibleByAuth_(task, auth, data);
  });
  filtered.userHourlyRates = {};
  if (data.userHourlyRates && data.userHourlyRates[auth.userId] != null) {
    filtered.userHourlyRates[auth.userId] = data.userHourlyRates[auth.userId];
  }
  filtered.staffWorkStatus = {};
  var selfUser = (data.users || {})[auth.userId];
  if (selfUser) {
    var selfLabel = getUserDisplayName_(selfUser);
    if (data.staffWorkStatus && data.staffWorkStatus[auth.userId] != null) filtered.staffWorkStatus[auth.userId] = data.staffWorkStatus[auth.userId];
    if (data.staffWorkStatus && data.staffWorkStatus[selfLabel] != null) filtered.staffWorkStatus[selfLabel] = data.staffWorkStatus[selfLabel];
  }
  return filtered;
}
function sanitizeReportForStaffWrite_(report) {
  var src = report || {};
  return {
    date: src.date || "", workType: src.workType || "",
    startH: src.startH || "", startM: src.startM || "",
    endH: src.endH || "", endM: src.endM || "",
    breakTime: src.breakTime || "", workTime: src.workTime || "",
    taskType: src.taskType || "", manHours: src.manHours || "",
    transport: src.transport || "", bizId: src.bizId || "",
    productId: src.productId || "", serviceId: src.serviceId || "",
    textCode: src.textCode || "", year: src.year || "", content: src.content || ""
  };
}
function sanitizeStampMap_(stamps) {
  var src = stamps || {};
  var clean = {};
  Object.keys(src).forEach(function(key) {
    var val = src[key];
    if (val === true || val === false || val === "emergency") clean[key] = val;
  });
  return clean;
}
function sanitizePendingStampRequestForStaff_(req) {
  if (!req) return null;
  return {
    status: "pending",
    requestedAt: req.requestedAt || Date.now(),
    stamps: sanitizeStampMap_(req.stamps),
    originalStamps: sanitizeStampMap_(req.originalStamps)
  };
}
function mergeAllowedStaffUserUpdate_(currentUser, incoming) {
  var src = incoming || {};
  var next = JSON.parse(JSON.stringify(currentUser || {}));
  next.id = currentUser && currentUser.id ? currentUser.id : (src.id || "");
  next.name = currentUser && currentUser.name ? currentUser.name : (src.name || "");
  next.userType = currentUser && currentUser.userType ? currentUser.userType : (src.userType || "");
  next.createdAt = currentUser && currentUser.createdAt ? currentUser.createdAt : (src.createdAt || Date.now());
  next.stamps = sanitizeStampMap_(src.stamps || next.stamps);
  next.reports = Array.isArray(src.reports) ? src.reports.map(sanitizeReportForStaffWrite_) : (next.reports || []);
  next.pendingStampRequest = sanitizePendingStampRequestForStaff_(src.pendingStampRequest);
  if (src.stampScreenVisitedToday != null) next.stampScreenVisitedToday = src.stampScreenVisitedToday;
  if (src.stampFailed != null) next.stampFailed = src.stampFailed;
  return next;
}
function ensureUserShapeServer_(user, userId) {
  var next = JSON.parse(JSON.stringify(user || {}));
  next.id = next.id || userId || "";
  next.stamps = sanitizeStampMap_(next.stamps || {});
  next.reports = Array.isArray(next.reports) ? next.reports : [];
  next.pendingStampRequest = next.pendingStampRequest || null;
  return next;
}
function issueReportId_() { return "report_" + Utilities.getUuid(); }
function sanitizeReportForWrite_(report, existing) {
  var clean = sanitizeReportForStaffWrite_(report || {});
  var base = existing || {};
  clean.reportId = String((report && report.reportId) || base.reportId || issueReportId_());
  if (base.proofCount != null) clean.proofCount = base.proofCount;
  if (base.incentiveAmount != null) clean.incentiveAmount = base.incentiveAmount;
  if (report && report.proofCount != null) clean.proofCount = Math.max(0, parseInt(report.proofCount, 10) || 0);
  if (report && report.incentiveAmount != null) clean.incentiveAmount = Math.max(0, parseInt(report.incentiveAmount, 10) || 0);
  return clean;
}
function ensureReportIdsOnUser_(user) {
  var next = ensureUserShapeServer_(user, user && user.id);
  next.reports = (next.reports || []).map(function(r) { return sanitizeReportForWrite_(r, r); });
  return next;
}
function findReportIndex_(reports, reportId, reportIndex) {
  if (!Array.isArray(reports)) return -1;
  if (reportId) {
    for (var i = 0; i < reports.length; i++) {
      if (String((reports[i] && reports[i].reportId) || "") === String(reportId)) return i;
    }
  }
  var idx = parseInt(reportIndex, 10);
  if (!isNaN(idx) && idx >= 0 && idx < reports.length) return idx;
  return -1;
}
function assertUserWritableByAuth_(auth, targetUserId) {
  if (!auth) return { ok: false, error: "unauthorized" };
  if (auth.role === "admin") return { ok: true };
  if (String(auth.userId || "") === String(targetUserId || "")) return { ok: true };
  return { ok: false, error: "forbidden" };
}
function sanitizeStampValue_(value) {
  if (value === true || value === false || value === "emergency") return value;
  return null;
}
function applyStampMetaPatch_(user, patch) {
  var next = ensureUserShapeServer_(user, user && user.id);
  var src = patch || {};
  if (Object.prototype.hasOwnProperty.call(src, "stampScreenVisitedToday")) {
    if (src.stampScreenVisitedToday) next.stampScreenVisitedToday = src.stampScreenVisitedToday;
    else delete next.stampScreenVisitedToday;
  }
  if (Object.prototype.hasOwnProperty.call(src, "stampFailed")) {
    if (src.stampFailed) next.stampFailed = src.stampFailed;
    else delete next.stampFailed;
  }
  if (Object.prototype.hasOwnProperty.call(src, "bonusPoints")) next.bonusPoints = Math.max(0, parseInt(src.bonusPoints, 10) || 0);
  if (Object.prototype.hasOwnProperty.call(src, "lastCongrats50")) next.lastCongrats50 = Math.max(0, parseInt(src.lastCongrats50, 10) || 0);
  if (Object.prototype.hasOwnProperty.call(src, "lastMonthFirstStamp")) next.lastMonthFirstStamp = src.lastMonthFirstStamp || "";
  return next;
}

// ====== トークン ======
function issueToken_(payload) {
  var token = Utilities.getUuid();
  CacheService.getScriptCache().put("t:" + token, JSON.stringify(payload), TOKEN_TTL_SECONDS);
  return token;
}
function getTokenPayload_(token) {
  if (!token) return null;
  var raw = CacheService.getScriptCache().get("t:" + token);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}
function safeParse_(s) { try { return JSON.parse(s); } catch (e) { return {}; } }
function requireAuth_(e) {
  var token = (e && e.parameter && e.parameter.token) ||
              (e && e.postData && e.postData.contents && safeParse_(e.postData.contents).token) || null;
  var payload = getTokenPayload_(token);
  if (!payload) return { ok: false, error: "unauthorized" };
  return { ok: true, auth: payload };
}
function requireAdmin_(e) {
  var gate = requireAuth_(e);
  if (!gate.ok) return gate;
  if (gate.auth.role !== "admin") return { ok: false, error: "forbidden" };
  return gate;
}

function getTodayPassword_() {
  var tz = "Asia/Tokyo";
  var d = new Date();
  var m = parseInt(Utilities.formatDate(d, tz, "M"), 10);
  var day = parseInt(Utilities.formatDate(d, tz, "d"), 10);
  var list = getDailyPasswords_();
  if (!list || !list.length) return "";
  var idx = ((m - 1) * 31 + day) % list.length;
  return list[idx];
}

// ====== Drive フォルダ ======
function getFilesFolder_() {
  if (FILES_ROOT_FOLDER_ID) {
    try {
      var root = DriveApp.getFolderById(FILES_ROOT_FOLDER_ID);
      PROP.setProperty("FILES_FOLDER_ID", root.getId());
      return root;
    } catch (e0) {}
  }
  var folderId = PROP.getProperty("FILES_FOLDER_ID");
  if (folderId) { try { return DriveApp.getFolderById(folderId); } catch (e) {} }
  var folders = DriveApp.getFoldersByName(FILES_FOLDER_NAME);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FILES_FOLDER_NAME);
  PROP.setProperty("FILES_FOLDER_ID", folder.getId());
  return folder;
}
function safeDriveFolderName_(name) {
  return String(name || "").replace(/[\\/:*?"<>|]/g, "_").trim() || "未指定";
}
function getOrCreateChildFolder_(parent, name) {
  var safeName = safeDriveFolderName_(name);
  var folders = parent.getFoldersByName(safeName);
  return folders.hasNext() ? folders.next() : parent.createFolder(safeName);
}
function ensureStaffDriveFolder_(user) {
  var root = getFilesFolder_();
  var folderName = safeDriveFolderName_(getUserDisplayName_(user) || (user && user.id) || "");
  var staffFolder = getOrCreateChildFolder_(root, folderName);
  getOrCreateChildFolder_(staffFolder, COMPLETED_FILES_FOLDER_NAME);
  return staffFolder;
}
function ensureAllStaffDriveFolders_(data) {
  var users = (data && data.users) || {};
  var count = 0;
  Object.keys(users).forEach(function(uid) {
    var user = users[uid];
    if (!user) return;
    ensureStaffDriveFolder_(user);
    count++;
  });
  var staff = getStaffAccounts_();
  Object.keys(staff).forEach(function(uid) {
    if (users[uid]) return;
    ensureStaffDriveFolder_({ id: uid, name: staff[uid].name || uid, userType: staff[uid].userType || "" });
    count++;
  });
  return count;
}
// 1日1回だけフォルダ整備を行う
function maybeRunDailyMaintenance_(data) {
  var today = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd");
  if (PROP.getProperty("LAST_FOLDER_ENSURE") === today) return;
  try { ensureAllStaffDriveFolders_(data); } catch (e) {}
  PROP.setProperty("LAST_FOLDER_ENSURE", today);
}
function getTaskStaffDriveFolder_(task, data) {
  var uid = getTaskAssignedUserId_(task, data);
  var user = uid ? ((data.users || {})[uid] || { id: uid }) : null;
  var label = user ? getUserDisplayName_(user) : "";
  if (!label) label = task ? String(task.staff || task.staffUserId || "未指定") : "未指定";
  return ensureStaffDriveFolder_(Object.assign({}, user || {}, { name: label }));
}
function getTaskUploadFolder_(task, data, auth, uploadMode) {
  var staffFolder = getTaskStaffDriveFolder_(task, data);
  if ((auth && auth.role === "staff") || uploadMode === "staff") {
    return getOrCreateChildFolder_(staffFolder, COMPLETED_FILES_FOLDER_NAME);
  }
  return staffFolder;
}
function isFileUnderFolder_(file, rootFolder) {
  var rootId = rootFolder.getId();
  var queue = [];
  var parents = file.getParents();
  while (parents.hasNext()) queue.push(parents.next());
  var seen = {};
  while (queue.length) {
    var folder = queue.shift();
    var id = folder.getId();
    if (seen[id]) continue;
    seen[id] = true;
    if (id === rootId) return true;
    var nextParents = folder.getParents();
    while (nextParents.hasNext()) queue.push(nextParents.next());
  }
  return false;
}

// ====== データファイル + キャッシュ ======
function getDataFile_() {
  var fileId = PROP.getProperty("DATA_FILE_ID");
  if (fileId) { try { return DriveApp.getFileById(fileId); } catch (e) {} }
  var files = DriveApp.getFilesByName(DATA_FILE_NAME);
  var file = files.hasNext() ? files.next() : DriveApp.createFile(
    DATA_FILE_NAME,
    JSON.stringify({ _version: 0, users: {}, session: {}, tasks: [], notices: [] }),
    "application/json"
  );
  PROP.setProperty("DATA_FILE_ID", file.getId());
  return file;
}
function cachePutChunked_(key, value) {
  try {
    var cache = CacheService.getScriptCache();
    var n = Math.ceil(value.length / DATA_CACHE_CHUNK);
    var pairs = { };
    pairs[key + ":n"] = String(n);
    for (var i = 0; i < n; i++) pairs[key + ":" + i] = value.substr(i * DATA_CACHE_CHUNK, DATA_CACHE_CHUNK);
    cache.putAll(pairs, DATA_CACHE_TTL);
  } catch (e) {}
}
function cacheGetChunked_(key) {
  try {
    var cache = CacheService.getScriptCache();
    var n = parseInt(cache.get(key + ":n") || "0", 10);
    if (!n) return null;
    var keys = [];
    for (var i = 0; i < n; i++) keys.push(key + ":" + i);
    var got = cache.getAll(keys);
    var s = "";
    for (var j = 0; j < n; j++) {
      var part = got[key + ":" + j];
      if (part == null) return null;
      s += part;
    }
    return s;
  } catch (e) { return null; }
}
function cacheRemoveChunked_(key) {
  try {
    var cache = CacheService.getScriptCache();
    var n = parseInt(cache.get(key + ":n") || "0", 10);
    var keys = [key + ":n"];
    for (var i = 0; i < n; i++) keys.push(key + ":" + i);
    cache.removeAll(keys);
  } catch (e) {}
}

function readDataRaw_() {
  var cached = cacheGetChunked_(DATA_CACHE_KEY);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }
  var file = getDataFile_();
  var raw = file.getBlob().getDataAsString();
  cachePutChunked_(DATA_CACHE_KEY, raw);
  try { return JSON.parse(raw); }
  catch (e) { return { _version: 0, users: {}, session: {}, tasks: [], notices: [] }; }
}
function readData_() {
  return mergeStaffAccountsIntoData_(readDataRaw_());
}

function mergeStaffAccountsIntoData_(data) {
  var next = data || {};
  next.users = next.users || {};
  next.userHourlyRates = next.userHourlyRates || {};
  var staff = getStaffAccounts_();
  Object.keys(staff).forEach(function(uid, index) {
    if (!uid) return;
    var acc = staff[uid] || {};
    var current = next.users[uid] || {};
    next.users[uid] = Object.assign({}, current, {
      id: uid,
      name: acc.name || current.name || uid,
      userType: acc.userType || current.userType || "学生",
      createdAt: current.createdAt || (Date.now() + index)
    });
    if (Object.prototype.hasOwnProperty.call(next.users[uid], "pw")) delete next.users[uid].pw;
    if (Object.prototype.hasOwnProperty.call(next.users[uid], "pwHash")) delete next.users[uid].pwHash;
    if (Object.prototype.hasOwnProperty.call(next.users[uid], "salt")) delete next.users[uid].salt;
    if (next.userHourlyRates[uid] == null && DEFAULT_USER_HOURLY_RATES[uid] != null) {
      next.userHourlyRates[uid] = DEFAULT_USER_HOURLY_RATES[uid];
    }
  });
  return next;
}

function stripSensitiveFromData_(data) {
  var next = data || {};
  next.users = next.users || {};
  Object.keys(next.users).forEach(function(uid) {
    var u = next.users[uid];
    if (!u) return;
    if (Object.prototype.hasOwnProperty.call(u, "pw")) delete u.pw;
    if (Object.prototype.hasOwnProperty.call(u, "pwHash")) delete u.pwHash;
    if (Object.prototype.hasOwnProperty.call(u, "salt")) delete u.salt;
  });
  return next;
}

function writeData_(data) {
  var file = getDataFile_();
  stripSensitiveFromData_(data);
  data._version = (data._version || 0) + 1;
  data._updatedAt = new Date().toISOString();
  var json = JSON.stringify(data);
  file.setContent(json);
  PROP.setProperty("DATA_VERSION", String(data._version));
  PROP.setProperty("DATA_UPDATED_AT", data._updatedAt);
  cachePutChunked_(DATA_CACHE_KEY, json);
  return { ok: true, version: data._version, updatedAt: data._updatedAt };
}

function withLock_(callback) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); return callback(); }
  finally { lock.releaseLock(); }
}
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function checkBaseVersion_(body, currentData) {
  if (!body || body._baseVersion == null) return { ok: true };
  var clientVersion = parseInt(body._baseVersion, 10);
  var serverVersion = parseInt((currentData && currentData._version) || 0, 10);
  if (clientVersion !== serverVersion) return { ok: false, error: "version_conflict", version: serverVersion };
  return { ok: true };
}
function getMonthLockKeyFromDate_(value) {
  var text = String(value || "");
  if (/^\d{4}-\d{2}$/.test(text)) return text;
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text.slice(0, 7);
  return "";
}
function isLockedMonth_(data, value) {
  var key = getMonthLockKeyFromDate_(value);
  return !!(key && data && data.lockedMonths && data.lockedMonths[key]);
}
function assertUnlockedMonth_(data, value) {
  return isLockedMonth_(data, value) ? { ok: false, error: "locked_month" } : { ok: true };
}

// ====== 管理者認証 ======
function getAdminCredsRecord_() {
  var rec = loadJsonProperty_("ADMIN_CRED_JSON", null);
  if (rec && rec.id && rec.pwHash && rec.salt) return rec;
  // 旧形式（プレーンPW）からの後方互換
  var legacyId = PROP.getProperty("ADMIN_ID");
  var legacyPw = PROP.getProperty("ADMIN_PW");
  if (legacyId && legacyPw) {
    var migrated = makePasswordRecord_(legacyPw);
    migrated.id = legacyId;
    saveJsonProperty_("ADMIN_CRED_JSON", migrated);
    PROP.deleteProperty("ADMIN_PW");
    return migrated;
  }
  return null;
}
function setAdminCredsHashed_(id, pw) {
  var rec = makePasswordRecord_(pw);
  rec.id = String(id);
  saveJsonProperty_("ADMIN_CRED_JSON", rec);
  PROP.setProperty("ADMIN_ID", String(id));
  PROP.deleteProperty("ADMIN_PW");
}
function verifyAdminLogin_(id, pw) {
  var rec = getAdminCredsRecord_();
  if (!rec) return false;
  if (String(rec.id) !== String(id)) return false;
  return verifyPasswordRecord_(rec, pw);
}

// ====== Staff 認証（ハッシュ + 自動アップグレード） ======
function verifyStaffLogin_(id, pw) {
  var staff = getStaffAccounts_();
  var acc = staff[id];
  if (!acc) return { ok: false };
  if (!verifyPasswordRecord_(acc, pw)) return { ok: false };
  // 平文 → ハッシュへ自動アップグレード
  if (acc.pw && (!acc.pwHash || !acc.salt)) {
    var rec = makePasswordRecord_(pw);
    delete acc.pw;
    acc.salt = rec.salt;
    acc.pwHash = rec.pwHash;
    staff[id] = acc;
    saveJsonProperty_("STAFF_ACCOUNTS_JSON", staff);
  }
  return { ok: true, account: acc };
}

// ====== doGet ======
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || "read";

  if (action === "ping") {
    return out_({ ok: true, ping: true, now: new Date().toISOString() });
  }

  var gate = requireAuth_(e);
  if (!gate.ok) return out_(gate);

  if (action === "todayPassword") {
    if (gate.auth.role !== "admin") return out_({ ok: false, error: "forbidden" });
    return out_({ ok: true, password: getTodayPassword_() });
  }

  if (action === "version") {
    var v = PROP.getProperty("DATA_VERSION") || "0";
    var u = PROP.getProperty("DATA_UPDATED_AT") || null;
    return out_({ ok: true, version: parseInt(v, 10), updatedAt: u });
  }

  if (action === "listStaffAccounts") {
    var gateAdminList = requireAdmin_(e);
    if (!gateAdminList.ok) return out_(gateAdminList);
    var staff = getStaffAccounts_();
    // PW は絶対に返さない（hashSet フラグのみ）
    var result = Object.keys(staff).sort().map(function(id) {
      return {
        id: id,
        name: staff[id].name || "",
        userType: staff[id].userType || "学生",
        hashSet: !!(staff[id].pwHash && staff[id].salt) || !!staff[id].pw
      };
    });
    return out_({ ok: true, staffAccounts: result });
  }

  if (action === "download") {
    var fileId = e.parameter.fileId;
    if (!fileId) return out_({ ok: false, error: "fileId required" });
    try {
      var dataForDownload = readData_();
      if (gate.auth.role !== "admin") {
        var taskId = String(e.parameter.taskId || "");
        var task = (dataForDownload.tasks || []).find(function(t) { return String(t.id) === taskId; });
        if (!task || !isTaskAccessibleByAuth_(task, gate.auth, dataForDownload)) return out_({ ok: false, error: "forbidden" });
        if (!(task.fileIds || []).some(function(id) { return String(id) === String(fileId); })) return out_({ ok: false, error: "forbidden" });
      }
      var f = DriveApp.getFileById(fileId);
      var folder = getFilesFolder_();
      if (!isFileUnderFolder_(f, folder)) return out_({ ok: false, error: "access denied: file not in app folder" });
      var b64 = Utilities.base64Encode(f.getBlob().getBytes());
      return out_({ ok: true, name: f.getName(), mimeType: f.getMimeType(), data: b64 });
    } catch (err) { return out_({ ok: false, error: err.message }); }
  }

  if (action === "listFiles") {
    var gateAdminFiles = requireAdmin_(e);
    if (!gateAdminFiles.ok) return out_(gateAdminFiles);
    var folder2 = getFilesFolder_();
    var iter = folder2.getFiles();
    var list = [];
    while (iter.hasNext()) {
      var fi = iter.next();
      list.push({ id: fi.getId(), name: fi.getName(), size: fi.getSize(), date: fi.getDateCreated().toISOString() });
    }
    return out_({ ok: true, files: list });
  }

  // ====== read（304 相当の最適化） ======
  var clientBaseVersion = e && e.parameter && e.parameter._baseVersion != null
    ? parseInt(e.parameter._baseVersion, 10) : NaN;
  var currentVersion = parseInt(PROP.getProperty("DATA_VERSION") || "0", 10);
  var currentUpdated = PROP.getProperty("DATA_UPDATED_AT") || null;
  if (!isNaN(clientBaseVersion) && clientBaseVersion === currentVersion && currentVersion > 0) {
    return out_({ ok: true, unchanged: true, version: currentVersion, updatedAt: currentUpdated });
  }

  var data = readData_();
  if (gate.auth.role === "admin") maybeRunDailyMaintenance_(data);
  data.session = {};
  return out_({ ok: true, data: filterDataForAuth_(data, gate.auth), version: data._version || currentVersion, updatedAt: data._updatedAt || currentUpdated });
}

// ====== doPost ======
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || "{}");
    var action = body._action || "";

    if (action === "loginStaff") {
      var id = String(body.id || "").trim();
      var pw = String(body.pw || "");
      if (!id || !pw) return out_({ ok: false, error: "invalid credentials" });
      if (!checkLoginRateLimit_("staff", id)) return out_({ ok: false, error: "too_many_attempts" });
      var v = verifyStaffLogin_(id, pw);
      if (!v.ok) {
        recordLoginFailure_("staff", id);
        return out_({ ok: false, error: "invalid credentials" });
      }
      clearLoginFailures_("staff", id);
      var token = issueToken_({ role: "staff", userId: id });
      return out_({ ok: true, token: token, user: { id: id, name: v.account.name, userType: v.account.userType } });
    }

    if (action === "loginAdmin") {
      var aid = String(body.id || "").trim();
      var apw = String(body.pw || "");
      if (!aid || !apw) return out_({ ok: false, error: "invalid credentials" });
      if (!checkLoginRateLimit_("admin", aid)) return out_({ ok: false, error: "too_many_attempts" });
      if (!verifyAdminLogin_(aid, apw)) {
        recordLoginFailure_("admin", aid);
        return out_({ ok: false, error: "invalid credentials" });
      }
      clearLoginFailures_("admin", aid);
      var tokenA = issueToken_({ role: "admin", adminId: aid });
      return out_({ ok: true, token: tokenA, admin: { id: aid, name: "管理者" } });
    }

    if (action === "verifyDailyPassword") {
      var gateV = requireAuth_(e);
      if (!gateV.ok) return out_(gateV);
      var ans = String(body.answer || "").trim();
      var ok = ans && ans === getTodayPassword_();
      return out_({ ok: true, match: !!ok });
    }

    var gate = requireAuth_(e);
    if (!gate.ok) return out_(gate);

    if (action === "setAdminCreds") {
      var gateAdmin = requireAdmin_(e);
      if (!gateAdmin.ok) return out_(gateAdmin);
      var oldRec = getAdminCredsRecord_();
      if (!oldRec || !verifyPasswordRecord_(oldRec, String(body.oldPw || ""))) {
        return out_({ ok: false, error: "旧パスワードが違います" });
      }
      var nid = String(body.newId || "").trim();
      var npw = String(body.newPw || "").trim();
      if (!nid || !npw) return out_({ ok: false, error: "ID/PWは必須です" });
      setAdminCredsHashed_(nid, npw);
      return out_({ ok: true });
    }

    if (action === "changeOwnStaffPassword") {
      if (gate.auth.role !== "staff") return out_({ ok: false, error: "forbidden" });
      var currentPw = String(body.currentPw || "");
      var newPw = String(body.newPw || "").trim();
      if (!currentPw || !newPw) return out_({ ok: false, error: "current/new password required" });
      return out_(withLock_(function() {
        var dataPw = readData_();
        var pwCheck = checkBaseVersion_(body, dataPw);
        if (!pwCheck.ok) return pwCheck;
        var staffPw = getStaffAccounts_();
        var selfId = String(gate.auth.userId || "");
        var selfAcc = staffPw[selfId];
        if (!selfAcc || !verifyPasswordRecord_(selfAcc, currentPw)) {
          return { ok: false, error: "current password mismatch" };
        }
        var rec = makePasswordRecord_(newPw);
        delete selfAcc.pw;
        selfAcc.salt = rec.salt;
        selfAcc.pwHash = rec.pwHash;
        staffPw[selfId] = selfAcc;
        saveJsonProperty_("STAFF_ACCOUNTS_JSON", staffPw);
        dataPw.users = dataPw.users || {};
        dataPw.users[selfId] = Object.assign({}, dataPw.users[selfId] || {}, {
          id: selfId, name: selfAcc.name || selfId,
          userType: selfAcc.userType || ((dataPw.users[selfId] || {}).userType || "学生")
        });
        var pwWrite = writeData_(dataPw);
        pwWrite.user = dataPw.users[selfId];
        return pwWrite;
      }));
    }

    if (action === "upsertStaffUser") {
      var gateA = requireAdmin_(e);
      if (!gateA.ok) return out_(gateA);
      var sid = String(body.id || "").trim();
      var spw = String(body.pw || "");
      var name = String(body.name || "").trim() || sid;
      var userType = String(body.userType || "学生").trim();
      if (!sid) return out_({ ok: false, error: "id required" });
      return out_(withLock_(function() {
        var data2 = readData_();
        var baseCheck2 = checkBaseVersion_(body, data2);
        if (!baseCheck2.ok) return baseCheck2;
        var staff2 = getStaffAccounts_();
        var old = staff2[sid] || {};
        var nextAcc = { name: name, userType: userType };
        if (spw) {
          var rec = makePasswordRecord_(spw);
          nextAcc.salt = rec.salt;
          nextAcc.pwHash = rec.pwHash;
        } else if (old.pwHash && old.salt) {
          nextAcc.salt = old.salt;
          nextAcc.pwHash = old.pwHash;
        } else if (old.pw) {
          // 旧プレーン保持を hash 化
          var rec2 = makePasswordRecord_(old.pw);
          nextAcc.salt = rec2.salt;
          nextAcc.pwHash = rec2.pwHash;
        } else {
          return { ok: false, error: "pw required" };
        }
        staff2[sid] = nextAcc;
        saveJsonProperty_("STAFF_ACCOUNTS_JSON", staff2);
        data2.users = data2.users || {};
        data2.users[sid] = Object.assign({}, data2.users[sid] || {}, { id: sid, name: name, userType: userType });
        ensureStaffDriveFolder_(data2.users[sid]);
        return writeData_(data2);
      }));
    }

    if (action === "deleteStaffUser") {
      var gateDel = requireAdmin_(e);
      if (!gateDel.ok) return out_(gateDel);
      var deleteId = String(body.id || "").trim();
      if (!deleteId) return out_({ ok: false, error: "id required" });
      return out_(withLock_(function() {
        var data3 = readData_();
        var baseCheck3 = checkBaseVersion_(body, data3);
        if (!baseCheck3.ok) return baseCheck3;
        var staff3 = getStaffAccounts_();
        if (staff3[deleteId]) { delete staff3[deleteId]; saveJsonProperty_("STAFF_ACCOUNTS_JSON", staff3); }
        data3.users = data3.users || {};
        if (data3.users[deleteId]) delete data3.users[deleteId];
        return writeData_(data3);
      }));
    }

    if (action === "uploadFile") {
      var uploadTaskId = String(body.taskId || "");
      var uploadData = readData_();
      var uploadTask = uploadTaskId ? (uploadData.tasks || []).find(function(t) { return String(t.id) === uploadTaskId; }) : null;
      if (gate.auth.role !== "admin") {
        if (!uploadTaskId) return out_({ ok: false, error: "taskId required" });
        if (!uploadTask || !isTaskAccessibleByAuth_(uploadTask, gate.auth, uploadData)) return out_({ ok: false, error: "forbidden" });
      }
      var folder3 = uploadTask ? getTaskUploadFolder_(uploadTask, uploadData, gate.auth, String(body.uploadMode || "")) : getFilesFolder_();
      var bytes = Utilities.base64Decode(body.data);
      var blob = Utilities.newBlob(bytes, body.mimeType || "application/octet-stream", body.fileName);
      var uploaded = folder3.createFile(blob);
      return out_({ ok: true, fileId: uploaded.getId(), fileName: uploaded.getName(), url: uploaded.getUrl() });
    }

    if (action === "setStamp") {
      var stampTargetId = String(body.targetUserId || "");
      var stampGate = assertUserWritableByAuth_(gate.auth, stampTargetId);
      if (!stampGate.ok) return out_(stampGate);
      var stampDate = String(body.date || "");
      if (!stampDate) return out_({ ok: false, error: "date required" });
      return out_(withLock_(function() {
        var currentDataStamp = readData_();
        var stampCheck = checkBaseVersion_(body, currentDataStamp);
        if (!stampCheck.ok) return stampCheck;
        var stampLock = assertUnlockedMonth_(currentDataStamp, stampDate);
        if (!stampLock.ok) return stampLock;
        currentDataStamp.users = currentDataStamp.users || {};
        var stampUser = ensureUserShapeServer_(currentDataStamp.users[stampTargetId] || {}, stampTargetId);
        var stampValue = sanitizeStampValue_(body.value);
        if (body.value == null || body.value === "") delete stampUser.stamps[stampDate];
        else if (stampValue == null) return { ok: false, error: "invalid stamp value" };
        else stampUser.stamps[stampDate] = stampValue;
        currentDataStamp.users[stampTargetId] = applyStampMetaPatch_(stampUser, body.metaPatch);
        var stampWrite = writeData_(currentDataStamp);
        stampWrite.user = currentDataStamp.users[stampTargetId];
        return stampWrite;
      }));
    }

    if (action === "setStampMeta") {
      var metaTargetId = String(body.targetUserId || "");
      var metaGate = assertUserWritableByAuth_(gate.auth, metaTargetId);
      if (!metaGate.ok) return out_(metaGate);
      return out_(withLock_(function() {
        var currentMetaData = readData_();
        var metaCheck = checkBaseVersion_(body, currentMetaData);
        if (!metaCheck.ok) return metaCheck;
        currentMetaData.users = currentMetaData.users || {};
        var metaUser = ensureUserShapeServer_(currentMetaData.users[metaTargetId] || {}, metaTargetId);
        currentMetaData.users[metaTargetId] = applyStampMetaPatch_(metaUser, body.metaPatch);
        var metaWrite = writeData_(currentMetaData);
        metaWrite.user = currentMetaData.users[metaTargetId];
        return metaWrite;
      }));
    }

    if (action === "requestStampCorrection") {
      var reqTargetId = String(body.targetUserId || "");
      var reqGate = assertUserWritableByAuth_(gate.auth, reqTargetId);
      if (!reqGate.ok) return out_(reqGate);
      return out_(withLock_(function() {
        var currentReqData = readData_();
        var reqCheck = checkBaseVersion_(body, currentReqData);
        if (!reqCheck.ok) return reqCheck;
        currentReqData.users = currentReqData.users || {};
        var reqUser = ensureUserShapeServer_(currentReqData.users[reqTargetId] || {}, reqTargetId);
        reqUser.pendingStampRequest = {
          status: "pending",
          requestedAt: Date.now(),
          stamps: sanitizeStampMap_(body.stamps),
          originalStamps: sanitizeStampMap_(reqUser.stamps)
        };
        currentReqData.users[reqTargetId] = reqUser;
        var reqWrite = writeData_(currentReqData);
        reqWrite.user = currentReqData.users[reqTargetId];
        return reqWrite;
      }));
    }

    if (action === "updateStampRequestDraft") {
      var gateDraft = requireAdmin_(e);
      if (!gateDraft.ok) return out_(gateDraft);
      var draftTargetId = String(body.targetUserId || "");
      if (!draftTargetId) return out_({ ok: false, error: "targetUserId required" });
      return out_(withLock_(function() {
        var currentDraftData = readData_();
        var draftCheck = checkBaseVersion_(body, currentDraftData);
        if (!draftCheck.ok) return draftCheck;
        currentDraftData.users = currentDraftData.users || {};
        var draftUser = ensureUserShapeServer_(currentDraftData.users[draftTargetId] || {}, draftTargetId);
        if (!draftUser.pendingStampRequest || draftUser.pendingStampRequest.status !== "pending") {
          return { ok: false, error: "no pending request" };
        }
        draftUser.pendingStampRequest.stamps = sanitizeStampMap_(body.stamps);
        currentDraftData.users[draftTargetId] = draftUser;
        var draftWrite = writeData_(currentDraftData);
        draftWrite.user = currentDraftData.users[draftTargetId];
        return draftWrite;
      }));
    }

    if (action === "resolveStampCorrection") {
      var gateResolve = requireAdmin_(e);
      if (!gateResolve.ok) return out_(gateResolve);
      var resolveTargetId = String(body.targetUserId || "");
      var decision = String(body.decision || "");
      if (!resolveTargetId || !decision) return out_({ ok: false, error: "targetUserId and decision required" });
      return out_(withLock_(function() {
        var currentResolveData = readData_();
        var resolveCheck = checkBaseVersion_(body, currentResolveData);
        if (!resolveCheck.ok) return resolveCheck;
        currentResolveData.users = currentResolveData.users || {};
        var resolveUser = ensureUserShapeServer_(currentResolveData.users[resolveTargetId] || {}, resolveTargetId);
        var pending = resolveUser.pendingStampRequest;
        if (!pending || pending.status !== "pending") return { ok: false, error: "no pending request" };
        if (decision === "approved") {
          resolveUser.stamps = sanitizeStampMap_(pending.stamps || body.stamps || {});
          resolveUser.pendingStampRequest = { status: "approved", resolvedAt: Date.now() };
        } else if (decision === "rejected") {
          resolveUser.pendingStampRequest = { status: "rejected", resolvedAt: Date.now() };
        } else {
          return { ok: false, error: "invalid decision" };
        }
        currentResolveData.users[resolveTargetId] = resolveUser;
        var resolveWrite = writeData_(currentResolveData);
        resolveWrite.user = currentResolveData.users[resolveTargetId];
        return resolveWrite;
      }));
    }

    if (action === "clearStampRequestState") {
      var clearTargetId = String(body.targetUserId || "");
      var clearGate = assertUserWritableByAuth_(gate.auth, clearTargetId);
      if (!clearGate.ok) return out_(clearGate);
      return out_(withLock_(function() {
        var currentClearData = readData_();
        var clearCheck = checkBaseVersion_(body, currentClearData);
        if (!clearCheck.ok) return clearCheck;
        currentClearData.users = currentClearData.users || {};
        var clearUser = ensureUserShapeServer_(currentClearData.users[clearTargetId] || {}, clearTargetId);
        clearUser.pendingStampRequest = null;
        currentClearData.users[clearTargetId] = clearUser;
        var clearWrite = writeData_(currentClearData);
        clearWrite.user = currentClearData.users[clearTargetId];
        return clearWrite;
      }));
    }

    if (action === "addReport") {
      var addTargetId = String(body.targetUserId || "");
      var addGate = assertUserWritableByAuth_(gate.auth, addTargetId);
      if (!addGate.ok) return out_(addGate);
      return out_(withLock_(function() {
        var currentAddData = readData_();
        var addCheck = checkBaseVersion_(body, currentAddData);
        if (!addCheck.ok) return addCheck;
        currentAddData.users = currentAddData.users || {};
        var addUser = ensureReportIdsOnUser_(currentAddData.users[addTargetId] || {});
        addUser.id = addUser.id || addTargetId;
        var nextReport = sanitizeReportForWrite_(body.report || {}, null);
        var addLock = assertUnlockedMonth_(currentAddData, nextReport.date);
        if (!addLock.ok) return addLock;
        addUser.reports.push(nextReport);
        currentAddData.users[addTargetId] = addUser;
        var addWrite = writeData_(currentAddData);
        addWrite.user = currentAddData.users[addTargetId];
        addWrite.report = nextReport;
        return addWrite;
      }));
    }

    if (action === "addReportsBatch") {
      var addBatchTargetId = String(body.targetUserId || "");
      var addBatchGate = assertUserWritableByAuth_(gate.auth, addBatchTargetId);
      if (!addBatchGate.ok) return out_(addBatchGate);
      return out_(withLock_(function() {
        var currentAddBatchData = readData_();
        var addBatchCheck = checkBaseVersion_(body, currentAddBatchData);
        if (!addBatchCheck.ok) return addBatchCheck;
        currentAddBatchData.users = currentAddBatchData.users || {};
        var addBatchUser = ensureReportIdsOnUser_(currentAddBatchData.users[addBatchTargetId] || {});
        addBatchUser.id = addBatchUser.id || addBatchTargetId;
        var incomingReports = Array.isArray(body.reports) ? body.reports : [];
        if (!incomingReports.length) return { ok: false, error: "reports required" };
        var addedReports = [];
        incomingReports.forEach(function(report) {
          var nextBatchReport = sanitizeReportForWrite_(report || {}, null);
          var reportLock = assertUnlockedMonth_(currentAddBatchData, nextBatchReport.date);
          if (!reportLock.ok) throw new Error("locked_month");
          addBatchUser.reports.push(nextBatchReport);
          addedReports.push(nextBatchReport);
        });
        currentAddBatchData.users[addBatchTargetId] = addBatchUser;
        var addBatchWrite = writeData_(currentAddBatchData);
        addBatchWrite.user = currentAddBatchData.users[addBatchTargetId];
        addBatchWrite.reports = addedReports;
        return addBatchWrite;
      }));
    }

    if (action === "updateReport") {
      var updateTargetId = String(body.targetUserId || "");
      var updateGate = assertUserWritableByAuth_(gate.auth, updateTargetId);
      if (!updateGate.ok) return out_(updateGate);
      return out_(withLock_(function() {
        var currentUpdateData = readData_();
        var updateCheck = checkBaseVersion_(body, currentUpdateData);
        if (!updateCheck.ok) return updateCheck;
        currentUpdateData.users = currentUpdateData.users || {};
        var updateUser = ensureReportIdsOnUser_(currentUpdateData.users[updateTargetId] || {});
        updateUser.id = updateUser.id || updateTargetId;
        var updateIndex = findReportIndex_(updateUser.reports, body.reportId, body.reportIndex);
        if (updateIndex < 0) return { ok: false, error: "report not found" };
        var existingReport = updateUser.reports[updateIndex] || {};
        var mergedReport = Object.assign({}, existingReport, body.patch || {});
        var updateLock = assertUnlockedMonth_(currentUpdateData, mergedReport.date || existingReport.date);
        if (!updateLock.ok) return updateLock;
        updateUser.reports[updateIndex] = sanitizeReportForWrite_(mergedReport, existingReport);
        currentUpdateData.users[updateTargetId] = updateUser;
        var updateWrite = writeData_(currentUpdateData);
        updateWrite.user = currentUpdateData.users[updateTargetId];
        updateWrite.report = updateUser.reports[updateIndex];
        return updateWrite;
      }));
    }

    if (action === "deleteReport") {
      var deleteReportTargetId = String(body.targetUserId || "");
      var deleteReportGate = assertUserWritableByAuth_(gate.auth, deleteReportTargetId);
      if (!deleteReportGate.ok) return out_(deleteReportGate);
      return out_(withLock_(function() {
        var currentDeleteReportData = readData_();
        var deleteReportCheck = checkBaseVersion_(body, currentDeleteReportData);
        if (!deleteReportCheck.ok) return deleteReportCheck;
        currentDeleteReportData.users = currentDeleteReportData.users || {};
        var deleteReportUser = ensureReportIdsOnUser_(currentDeleteReportData.users[deleteReportTargetId] || {});
        deleteReportUser.id = deleteReportUser.id || deleteReportTargetId;
        var deleteIndex = findReportIndex_(deleteReportUser.reports, body.reportId, body.reportIndex);
        if (deleteIndex < 0) return { ok: false, error: "report not found" };
        var deleteLock = assertUnlockedMonth_(currentDeleteReportData, (deleteReportUser.reports[deleteIndex] || {}).date);
        if (!deleteLock.ok) return deleteLock;
        deleteReportUser.reports.splice(deleteIndex, 1);
        currentDeleteReportData.users[deleteReportTargetId] = deleteReportUser;
        var deleteWrite = writeData_(currentDeleteReportData);
        deleteWrite.user = currentDeleteReportData.users[deleteReportTargetId];
        return deleteWrite;
      }));
    }

    if (action === "setReportReview") {
      var gateReview = requireAdmin_(e);
      if (!gateReview.ok) return out_(gateReview);
      var reviewTargetId = String(body.targetUserId || "");
      if (!reviewTargetId) return out_({ ok: false, error: "targetUserId required" });
      return out_(withLock_(function() {
        var currentReviewData = readData_();
        var reviewCheck = checkBaseVersion_(body, currentReviewData);
        if (!reviewCheck.ok) return reviewCheck;
        currentReviewData.users = currentReviewData.users || {};
        var reviewUser = ensureReportIdsOnUser_(currentReviewData.users[reviewTargetId] || {});
        reviewUser.id = reviewUser.id || reviewTargetId;
        var reviewIndex = findReportIndex_(reviewUser.reports, body.reportId, body.reportIndex);
        if (reviewIndex < 0) return { ok: false, error: "report not found" };
        var reviewReport = reviewUser.reports[reviewIndex];
        if (Object.prototype.hasOwnProperty.call(body, "proofCount")) reviewReport.proofCount = Math.max(0, parseInt(body.proofCount, 10) || 0);
        if (Object.prototype.hasOwnProperty.call(body, "incentiveAmount")) reviewReport.incentiveAmount = Math.max(0, parseInt(body.incentiveAmount, 10) || 0);
        currentReviewData.users[reviewTargetId] = reviewUser;
        var reviewWrite = writeData_(currentReviewData);
        reviewWrite.user = currentReviewData.users[reviewTargetId];
        reviewWrite.report = reviewReport;
        return reviewWrite;
      }));
    }

    if (action === "updateUserFull") {
      var targetId = body.targetUserId;
      if (gate.auth.role !== "admin" && gate.auth.userId !== targetId) return out_({ ok: false, error: "forbidden" });
      return out_(withLock_(function() {
        var currentData = readData_();
        var userCheck = checkBaseVersion_(body, currentData);
        if (!userCheck.ok) return userCheck;
        currentData.users = currentData.users || {};
        if (body.userObj) {
          if (gate.auth.role === "admin") currentData.users[targetId] = body.userObj;
          else currentData.users[targetId] = mergeAllowedStaffUserUpdate_(currentData.users[targetId] || {}, body.userObj);
        }
        return writeData_(currentData);
      }));
    }

    if (action === "updateTask") {
      return out_(withLock_(function() {
        var currentData2 = readData_();
        var taskCheck = checkBaseVersion_(body, currentData2);
        if (!taskCheck.ok) return taskCheck;
        currentData2.tasks = currentData2.tasks || [];
        var tIdx = currentData2.tasks.findIndex(function(t) { return t.id === body.task.id; });
        var existingTask = tIdx >= 0 ? currentData2.tasks[tIdx] : null;
        var taskLock = assertUnlockedMonth_(currentData2, (body.task && body.task.requestDate) || (existingTask && existingTask.requestDate));
        if (!taskLock.ok) return taskLock;
        if (gate.auth.role !== "admin") {
          if (!existingTask || !isTaskAccessibleByAuth_(existingTask, gate.auth, currentData2)) return { ok: false, error: "forbidden" };
          if (body.isDelete) return { ok: false, error: "forbidden" };
          var nextTask = JSON.parse(JSON.stringify(existingTask));
          if (body.task && Object.prototype.hasOwnProperty.call(body.task, "notes")) nextTask.notes = body.task.notes;
          if (body.task && Array.isArray(body.task.fileNames)) nextTask.fileNames = body.task.fileNames;
          if (body.task && Array.isArray(body.task.fileIds)) nextTask.fileIds = body.task.fileIds;
          if (body.task && (body.task.status === "提出中" || body.task.status === "依頼中")) nextTask.status = body.task.status;
          if (body.task && Object.prototype.hasOwnProperty.call(body.task, "completionDate")) nextTask.completionDate = body.task.completionDate;
          if (body.task && Object.prototype.hasOwnProperty.call(body.task, "workTypeChangeRequest")) {
            var incomingWorkTypeReq = body.task.workTypeChangeRequest;
            if (incomingWorkTypeReq && incomingWorkTypeReq.status === "pending") {
              try {
                nextTask.workTypeChangeRequest = buildStaffWorkTypeChangeRequest_(existingTask, incomingWorkTypeReq, gate.auth);
              } catch (reqErr) {
                return { ok: false, error: reqErr.message };
              }
            } else {
              nextTask.workTypeChangeRequest = existingTask.workTypeChangeRequest || null;
            }
          }
          currentData2.tasks[tIdx] = nextTask;
          return writeData_(currentData2);
        }
        if (body.isDelete) { if (tIdx >= 0) currentData2.tasks.splice(tIdx, 1); }
        else { if (tIdx >= 0) currentData2.tasks[tIdx] = body.task; else currentData2.tasks.push(body.task); }
        return writeData_(currentData2);
      }));
    }

    if (action === "updateMaster") {
      var gateMaster = requireAdmin_(e);
      if (!gateMaster.ok) return out_(gateMaster);
      return out_(withLock_(function() {
        var currentData3 = readData_();
        var masterCheck = checkBaseVersion_(body, currentData3);
        if (!masterCheck.ok) return masterCheck;
        if (body.taskTypes) currentData3.taskTypes = body.taskTypes;
        if (body.taskPrices) currentData3.taskPrices = body.taskPrices;
        if (body.employees) currentData3.employees = body.employees;
        if (body.userHourlyRates) currentData3.userHourlyRates = body.userHourlyRates;
        if (body.staffWorkStatus) currentData3.staffWorkStatus = body.staffWorkStatus;
        if (body.lockedMonths) currentData3.lockedMonths = body.lockedMonths;
        if (body.stampIncentiveRules) currentData3.stampIncentiveRules = body.stampIncentiveRules;
        if (body.notices) currentData3.notices = body.notices;
        if (body.deleteUserId) {
          var ids = Array.isArray(body.deleteUserId) ? body.deleteUserId : [body.deleteUserId];
          var staff4 = getStaffAccounts_();
          currentData3.users = currentData3.users || {};
          ids.forEach(function(uid) {
            if (!uid) return;
            if (currentData3.users[uid]) delete currentData3.users[uid];
            if (staff4[uid]) delete staff4[uid];
          });
          saveJsonProperty_("STAFF_ACCOUNTS_JSON", staff4);
        }
        return writeData_(currentData3);
      }));
    }

    return out_({ ok: false, error: "unknown action" });
  } catch (err) {
    return out_({ ok: false, error: err.message });
  }
}
