var DATA_FILE_NAME = "業務管理アプリ_data.json";
var FILES_FOLDER_NAME = "業務管理アプリ_files";

var TOKEN_TTL_SECONDS = 6 * 60 * 60;
var PROP = PropertiesService.getScriptProperties();

// 本番では Script Properties を正本にする
var DEFAULT_STAFF_ACCOUNTS = {
  shakai_test: { pw: "shakai_test", name: "テスト社会人", userType: "社会人" },
  ogasawara: { pw: "ogasawara", name: "小笠原", userType: "社会人" },
  morotomi: { pw: "morotomi", name: "諸富", userType: "社会人" },
  osawa: { pw: "osawa", name: "大澤", userType: "社会人" },
  yoneoka: { pw: "yoneoka", name: "米岡", userType: "社会人" },
  hosaka: { pw: "hosaka", name: "保坂", userType: "社会人" },
  gakusei_test: { pw: "gakusei_test", name: "テスト学生", userType: "学生" },
  miko: { pw: "miko", name: "神子", userType: "学生" },
  shirakawa: { pw: "shirakawa", name: "白川", userType: "学生" },
  matsumoto: { pw: "matsumoto", name: "松本", userType: "学生" },
  mizutani: { pw: "mizutani", name: "水谷", userType: "学生" },
  takeuchi: { pw: "takeuchi", name: "竹内", userType: "学生" },
  fujikawa: { pw: "fujikawa", name: "藤川", userType: "学生" },
  kobayashi: { pw: "kobayashi", name: "小林", userType: "学生" }
};
var DEFAULT_USER_HOURLY_RATES = {
  shakai_test: 1300,
  ogasawara: 1300,
  morotomi: 1300,
  osawa: 1300,
  yoneoka: 1300,
  hosaka: 1300,
  gakusei_test: 1300,
  miko: 1300,
  shirakawa: 1300,
  matsumoto: 1300,
  mizutani: 1300,
  takeuchi: 1300,
  fujikawa: 1300,
  kobayashi: 1300
};
var DEFAULT_DAILY_PASSWORDS = [
  "さくら","ひかり","かぜ","うみ","そら","ほし","つき","やま","はな","にじ",
  "ゆめ","たね","もり","かわ","くも","きり","しお","すな","いわ","いし",
  "つる","かめ","まつ","たけ","うめ","きく","ばら","ゆり","すみ","もも",
  "りんご","みかん","かき","なし"
];

function loadJsonProperty_(key, fallbackObj) {
  var raw = PROP.getProperty(key);
  if (!raw) return fallbackObj;
  try { return JSON.parse(raw); } catch (e) { return fallbackObj; }
}
function saveJsonProperty_(key, obj) {
  PROP.setProperty(key, JSON.stringify(obj));
}
function getStaffAccounts_() {
  var obj = loadJsonProperty_("STAFF_ACCOUNTS_JSON", DEFAULT_STAFF_ACCOUNTS);
  if (!obj || typeof obj !== "object") return {};
  return obj;
}
function getDailyPasswords_() {
  return loadJsonProperty_("DAILY_PASSWORDS_JSON", DEFAULT_DAILY_PASSWORDS);
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
  delete clean.pw;
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
    date: src.date || "",
    workType: src.workType || "",
    startH: src.startH || "",
    startM: src.startM || "",
    endH: src.endH || "",
    endM: src.endM || "",
    breakTime: src.breakTime || "",
    workTime: src.workTime || "",
    taskType: src.taskType || "",
    manHours: src.manHours || "",
    transport: src.transport || "",
    bizId: src.bizId || "",
    productId: src.productId || "",
    serviceId: src.serviceId || "",
    textCode: src.textCode || "",
    year: src.year || "",
    content: src.content || ""
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
function issueReportId_() {
  return "report_" + Utilities.getUuid();
}
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
  next.reports = (next.reports || []).map(function(report) {
    return sanitizeReportForWrite_(report, report);
  });
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
function applyStampMetaPatch_(user, patch, auth) {
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
  if (Object.prototype.hasOwnProperty.call(src, "bonusPoints")) {
    next.bonusPoints = Math.max(0, parseInt(src.bonusPoints, 10) || 0);
  }
  if (Object.prototype.hasOwnProperty.call(src, "lastCongrats50")) {
    next.lastCongrats50 = Math.max(0, parseInt(src.lastCongrats50, 10) || 0);
  }
  if (Object.prototype.hasOwnProperty.call(src, "lastMonthFirstStamp")) {
    next.lastMonthFirstStamp = src.lastMonthFirstStamp || "";
  }
  return next;
}

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
function safeParse_(s) {
  try { return JSON.parse(s); } catch (e) { return {}; }
}
function requireAuth_(e) {
  var token =
    (e && e.parameter && e.parameter.token) ||
    (e && e.postData && e.postData.contents && safeParse_(e.postData.contents).token) ||
    null;

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

function getFilesFolder_() {
  var folderId = PROP.getProperty("FILES_FOLDER_ID");
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch (e) {}
  }
  var folders = DriveApp.getFoldersByName(FILES_FOLDER_NAME);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FILES_FOLDER_NAME);
  PROP.setProperty("FILES_FOLDER_ID", folder.getId());
  return folder;
}

function getDataFile_() {
  var fileId = PROP.getProperty("DATA_FILE_ID");
  if (fileId) {
    try { return DriveApp.getFileById(fileId); } catch (e) {}
  }
  var files = DriveApp.getFilesByName(DATA_FILE_NAME);
  var file = files.hasNext()
    ? files.next()
    : DriveApp.createFile(
        DATA_FILE_NAME,
        JSON.stringify({ _version: 0, users: {}, session: {}, tasks: [], notices: [] }),
        "application/json"
      );
  PROP.setProperty("DATA_FILE_ID", file.getId());
  return file;
}

function collectRecoverableStaffAccountsFromUsers_(users) {
  var source = users || {};
  var recovered = {};
  Object.keys(source).forEach(function(key) {
    var user = source[key] || {};
    var id = String(user.id || key || "").trim();
    var pw = String(user.pw || "").trim();
    if (!id || !pw) return;
    recovered[id] = {
      pw: pw,
      name: String(user.name || id),
      userType: String(user.userType || "")
    };
  });
  return recovered;
}

function ensureRecoverableStaffAccounts_(data) {
  var current = getStaffAccounts_();
  var recovered = collectRecoverableStaffAccountsFromUsers_((data && data.users) || {});
  var next = {};
  var changed = false;

  Object.keys(current || {}).forEach(function(uid) {
    if (!uid) return;
    var acc = current[uid] || {};
    next[uid] = {
      pw: String(acc.pw || ""),
      name: String(acc.name || uid),
      userType: String(acc.userType || "")
    };
  });

  Object.keys(recovered).forEach(function(uid) {
    var existing = next[uid] || {};
    if (!existing.pw) {
      next[uid] = {
        pw: recovered[uid].pw,
        name: recovered[uid].name || existing.name || uid,
        userType: recovered[uid].userType || existing.userType || ""
      };
      changed = true;
      return;
    }
    if (!existing.name && recovered[uid].name) {
      next[uid].name = recovered[uid].name;
      changed = true;
    }
    if (!existing.userType && recovered[uid].userType) {
      next[uid].userType = recovered[uid].userType;
      changed = true;
    }
  });

  if (!Object.keys(next).length && !Object.keys(recovered).length) {
    Object.keys(DEFAULT_STAFF_ACCOUNTS).forEach(function(uid) {
      var acc = DEFAULT_STAFF_ACCOUNTS[uid] || {};
      next[uid] = {
        pw: String(acc.pw || ""),
        name: String(acc.name || uid),
        userType: String(acc.userType || "")
      };
    });
    changed = Object.keys(next).length > 0;
  }

  if (changed) saveJsonProperty_("STAFF_ACCOUNTS_JSON", next);
  return changed ? next : current;
}

function readData_() {
  var file = getDataFile_();
  try {
    return mergeStaffAccountsIntoData_(JSON.parse(file.getBlob().getDataAsString()));
  }
  catch (e) {
    return mergeStaffAccountsIntoData_({ _version: 0, users: {}, session: {}, tasks: [], notices: [] });
  }
}

function mergeStaffAccountsIntoData_(data) {
  var next = data || {};
  next.users = next.users || {};
  next.userHourlyRates = next.userHourlyRates || {};
  var staff = ensureRecoverableStaffAccounts_(next);
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
    if (next.userHourlyRates[uid] == null && DEFAULT_USER_HOURLY_RATES[uid] != null) {
      next.userHourlyRates[uid] = DEFAULT_USER_HOURLY_RATES[uid];
    }
  });
  return next;
}

function stripPasswordsFromData_(data) {
  var next = data || {};
  next.users = next.users || {};
  Object.keys(next.users).forEach(function(uid) {
    if (next.users[uid] && Object.prototype.hasOwnProperty.call(next.users[uid], "pw")) {
      delete next.users[uid].pw;
    }
  });
  return next;
}

function writeData_(data) {
  var file = getDataFile_();
  stripPasswordsFromData_(data);
  data._version = (data._version || 0) + 1;
  data._updatedAt = new Date().toISOString();
  file.setContent(JSON.stringify(data));

  PROP.setProperty("DATA_VERSION", String(data._version));
  PROP.setProperty("DATA_UPDATED_AT", data._updatedAt);

  return { ok: true, version: data._version, updatedAt: data._updatedAt };
}

function withLock_(callback) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function out_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function checkBaseVersion_(body, currentData) {
  if (!body || body._baseVersion == null) return { ok: true };
  var clientVersion = parseInt(body._baseVersion, 10);
  var serverVersion = parseInt((currentData && currentData._version) || 0, 10);
  if (clientVersion !== serverVersion) {
    return { ok: false, error: "version_conflict", version: serverVersion };
  }
  return { ok: true };
}

function getAdminCreds_() {
  var id = PROP.getProperty("ADMIN_ID") || "suugaku";
  var pw = PROP.getProperty("ADMIN_PW") || "stamp";
  return { id: id, pw: pw };
}

function setAdminCreds_(id, pw) {
  PROP.setProperty("ADMIN_ID", id);
  PROP.setProperty("ADMIN_PW", pw);
}

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
    var result = Object.keys(staff).sort().map(function(id) {
      return {
        id: id,
        name: staff[id].name || "",
        userType: staff[id].userType || "学生",
        pw: staff[id].pw || ""
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
        if (!task || !isTaskAccessibleByAuth_(task, gate.auth, dataForDownload)) {
          return out_({ ok: false, error: "forbidden" });
        }
        if (!(task.fileIds || []).some(function(id) { return String(id) === String(fileId); })) {
          return out_({ ok: false, error: "forbidden" });
        }
      }
      var f = DriveApp.getFileById(fileId);
      var folder = getFilesFolder_();
      var parents = f.getParents();
      var allowed = false;
      while (parents.hasNext()) {
        if (parents.next().getId() === folder.getId()) {
          allowed = true;
          break;
        }
      }
      if (!allowed) return out_({ ok: false, error: "access denied: file not in app folder" });

      var b64 = Utilities.base64Encode(f.getBlob().getBytes());
      return out_({ ok: true, name: f.getName(), mimeType: f.getMimeType(), data: b64 });
    } catch (err) {
      return out_({ ok: false, error: err.message });
    }
  }

  if (action === "listFiles") {
    var gateAdminFiles = requireAdmin_(e);
    if (!gateAdminFiles.ok) return out_(gateAdminFiles);
    var folder2 = getFilesFolder_();
    var iter = folder2.getFiles();
    var list = [];
    while (iter.hasNext()) {
      var fi = iter.next();
      list.push({
        id: fi.getId(),
        name: fi.getName(),
        size: fi.getSize(),
        date: fi.getDateCreated().toISOString()
      });
    }
    return out_({ ok: true, files: list });
  }

  var data = readData_();
  data.session = {};
  return out_({ ok: true, data: filterDataForAuth_(data, gate.auth) });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || "{}");
    var action = body._action || "";

    if (action === "loginStaff") {
      var id = String(body.id || "").trim();
      var pw = String(body.pw || "");
      var staff = getStaffAccounts_();
      var acc = staff[id];

      if (!acc || acc.pw !== pw) {
        return out_({ ok: false, error: "invalid credentials" });
      }

      var token = issueToken_({ role: "staff", userId: id });
      return out_({
        ok: true,
        token: token,
        user: { id: id, name: acc.name, userType: acc.userType }
      });
    }

    if (action === "loginAdmin") {
      var aid = String(body.id || "").trim();
      var apw = String(body.pw || "");
      var creds = getAdminCreds_();

      if (aid !== creds.id || apw !== creds.pw) {
        return out_({ ok: false, error: "invalid credentials" });
      }

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

      var creds2 = getAdminCreds_();
      var oldOk = String(body.oldPw || "") === creds2.pw;
      if (!oldOk) return out_({ ok: false, error: "旧パスワードが違います" });

      var nid = String(body.newId || "").trim();
      var npw = String(body.newPw || "").trim();
      if (!nid || !npw) return out_({ ok: false, error: "ID/PWは必須です" });

      setAdminCreds_(nid, npw);
      return out_({ ok: true });
    }

    if (action === "changeOwnStaffPassword") {
      if (gate.auth.role !== "staff") {
        return out_({ ok: false, error: "forbidden" });
      }

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
        if (!selfAcc || String(selfAcc.pw || "") !== currentPw) {
          return { ok: false, error: "current password mismatch" };
        }

        selfAcc.pw = newPw;
        staffPw[selfId] = selfAcc;
        saveJsonProperty_("STAFF_ACCOUNTS_JSON", staffPw);

        dataPw.users = dataPw.users || {};
        dataPw.users[selfId] = Object.assign({}, dataPw.users[selfId] || {}, {
          id: selfId,
          name: selfAcc.name || selfId,
          userType: selfAcc.userType || ((dataPw.users[selfId] || {}).userType || "学生"),
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
      var spw = String(body.pw || "").trim();
      var name = String(body.name || "").trim() || sid;
      var userType = String(body.userType || "学生").trim();

      if (!sid) return out_({ ok: false, error: "id required" });

      return out_(withLock_(function() {
        var data2 = readData_();
        var baseCheck2 = checkBaseVersion_(body, data2);
        if (!baseCheck2.ok) return baseCheck2;
        var staff2 = getStaffAccounts_();
        var old = staff2[sid] || {};
        var nextPw = spw || old.pw || "";
        if (!nextPw) return { ok: false, error: "pw required" };

        staff2[sid] = {
          pw: nextPw,
          name: name,
          userType: userType
        };
        saveJsonProperty_("STAFF_ACCOUNTS_JSON", staff2);

        data2.users = data2.users || {};
        data2.users[sid] = Object.assign({}, data2.users[sid] || {}, {
          id: sid,
          name: name,
          userType: userType,
        });
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
        if (staff3[deleteId]) {
          delete staff3[deleteId];
          saveJsonProperty_("STAFF_ACCOUNTS_JSON", staff3);
        }

        data3.users = data3.users || {};
        if (data3.users[deleteId]) delete data3.users[deleteId];
        return writeData_(data3);
      }));
    }

    if (action === "uploadFile") {
      if (gate.auth.role !== "admin") {
        var uploadTaskId = String(body.taskId || "");
        if (!uploadTaskId) return out_({ ok: false, error: "taskId required" });
        var uploadData = readData_();
        var uploadTask = (uploadData.tasks || []).find(function(t) { return String(t.id) === uploadTaskId; });
        if (!uploadTask || !isTaskAccessibleByAuth_(uploadTask, gate.auth, uploadData)) {
          return out_({ ok: false, error: "forbidden" });
        }
      }
      var folder3 = getFilesFolder_();
      var bytes = Utilities.base64Decode(body.data);
      var blob = Utilities.newBlob(bytes, body.mimeType || "application/octet-stream", body.fileName);
      var uploaded = folder3.createFile(blob);
      return out_({
        ok: true,
        fileId: uploaded.getId(),
        fileName: uploaded.getName(),
        url: uploaded.getUrl()
      });
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
        currentDataStamp.users = currentDataStamp.users || {};
        var stampUser = ensureUserShapeServer_(currentDataStamp.users[stampTargetId] || {}, stampTargetId);
        var stampValue = sanitizeStampValue_(body.value);
        if (body.value == null || body.value === "") {
          delete stampUser.stamps[stampDate];
        } else if (stampValue == null) {
          return { ok: false, error: "invalid stamp value" };
        } else {
          stampUser.stamps[stampDate] = stampValue;
        }
        currentDataStamp.users[stampTargetId] = applyStampMetaPatch_(stampUser, body.metaPatch, gate.auth);
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
        currentMetaData.users[metaTargetId] = applyStampMetaPatch_(metaUser, body.metaPatch, gate.auth);
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
        if (Object.prototype.hasOwnProperty.call(body, "proofCount")) {
          reviewReport.proofCount = Math.max(0, parseInt(body.proofCount, 10) || 0);
        }
        if (Object.prototype.hasOwnProperty.call(body, "incentiveAmount")) {
          reviewReport.incentiveAmount = Math.max(0, parseInt(body.incentiveAmount, 10) || 0);
        }
        currentReviewData.users[reviewTargetId] = reviewUser;
        var reviewWrite = writeData_(currentReviewData);
        reviewWrite.user = currentReviewData.users[reviewTargetId];
        reviewWrite.report = reviewReport;
        return reviewWrite;
      }));
    }

    if (action === "updateUserFull") {
      var targetId = body.targetUserId;
      if (gate.auth.role !== "admin" && gate.auth.userId !== targetId) {
        return out_({ ok: false, error: "forbidden" });
      }

      return out_(withLock_(function() {
        var currentData = readData_();
        var userCheck = checkBaseVersion_(body, currentData);
        if (!userCheck.ok) return userCheck;
        currentData.users = currentData.users || {};
        if (body.userObj) {
          if (gate.auth.role === "admin") {
            currentData.users[targetId] = body.userObj;
          } else {
            currentData.users[targetId] = mergeAllowedStaffUserUpdate_(currentData.users[targetId] || {}, body.userObj);
          }
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
        var tIdx = currentData2.tasks.findIndex(function(t) {
          return t.id === body.task.id;
        });
        var existingTask = tIdx >= 0 ? currentData2.tasks[tIdx] : null;
        if (gate.auth.role !== "admin") {
          if (!existingTask || !isTaskAccessibleByAuth_(existingTask, gate.auth, currentData2)) {
            return { ok: false, error: "forbidden" };
          }
          if (body.isDelete) {
            return { ok: false, error: "forbidden" };
          }
          var nextTask = JSON.parse(JSON.stringify(existingTask));
          if (body.task && Object.prototype.hasOwnProperty.call(body.task, "notes")) nextTask.notes = body.task.notes;
          if (body.task && Array.isArray(body.task.fileNames)) nextTask.fileNames = body.task.fileNames;
          if (body.task && Array.isArray(body.task.fileIds)) nextTask.fileIds = body.task.fileIds;
          if (body.task && (body.task.status === "完了" || body.task.status === "依頼中")) nextTask.status = body.task.status;
          if (body.task && Object.prototype.hasOwnProperty.call(body.task, "completionDate")) nextTask.completionDate = body.task.completionDate;
          currentData2.tasks[tIdx] = nextTask;
          return writeData_(currentData2);
        }

        if (body.isDelete) {
          if (tIdx >= 0) currentData2.tasks.splice(tIdx, 1);
        } else {
          if (tIdx >= 0) currentData2.tasks[tIdx] = body.task;
          else currentData2.tasks.push(body.task);
        }
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
