var DATA_FILE_NAME = "業務管理アプリ_data.json";
var FILES_FOLDER_NAME = "業務管理アプリ_files";

var TOKEN_TTL_SECONDS = 6 * 60 * 60;
var PROP = PropertiesService.getScriptProperties();

// 本番では Script Properties を正本にする
var DEFAULT_STAFF_ACCOUNTS = {};
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
        JSON.stringify({ _version: 0, users: {}, session: {}, tasks: [] }),
        "application/json"
      );
  PROP.setProperty("DATA_FILE_ID", file.getId());
  return file;
}

function readData_() {
  var file = getDataFile_();
  try { return JSON.parse(file.getBlob().getDataAsString()); }
  catch (e) { return { _version: 0, users: {}, session: {}, tasks: [] }; }
}

function writeData_(data) {
  var file = getDataFile_();
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
        userType: staff[id].userType || "学生"
      };
    });
    return out_({ ok: true, staffAccounts: result });
  }

  if (action === "download") {
    var fileId = e.parameter.fileId;
    if (!fileId) return out_({ ok: false, error: "fileId required" });
    try {
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
  return out_({ ok: true, data: data });
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

      var data = readData_();
      data.users = data.users || {};
      data.users[id] = Object.assign({}, data.users[id] || {}, {
        id: id,
        name: acc.name,
        userType: acc.userType
      });
      writeData_(data);

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

    if (action === "upsertStaffUser") {
      var gateA = requireAdmin_(e);
      if (!gateA.ok) return out_(gateA);

      var sid = String(body.id || "").trim();
      var spw = String(body.pw || "").trim();
      var name = String(body.name || "").trim() || sid;
      var userType = String(body.userType || "学生").trim();

      if (!sid) return out_({ ok: false, error: "id required" });

      return out_(withLock_(function() {
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

        var data2 = readData_();
        data2.users = data2.users || {};
        data2.users[sid] = Object.assign({}, data2.users[sid] || {}, {
          id: sid,
          name: name,
          userType: userType
        });
        writeData_(data2);

        return { ok: true };
      }));
    }

    if (action === "deleteStaffUser") {
      var gateDel = requireAdmin_(e);
      if (!gateDel.ok) return out_(gateDel);

      var deleteId = String(body.id || "").trim();
      if (!deleteId) return out_({ ok: false, error: "id required" });

      return out_(withLock_(function() {
        var staff3 = getStaffAccounts_();
        if (staff3[deleteId]) {
          delete staff3[deleteId];
          saveJsonProperty_("STAFF_ACCOUNTS_JSON", staff3);
        }

        var data3 = readData_();
        data3.users = data3.users || {};
        if (data3.users[deleteId]) delete data3.users[deleteId];
        writeData_(data3);

        return { ok: true };
      }));
    }

    if (action === "uploadFile") {
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

    if (action === "updateUserFull") {
      var targetId = body.targetUserId;
      if (gate.auth.role !== "admin" && gate.auth.userId !== targetId) {
        return out_({ ok: false, error: "forbidden" });
      }

      return out_(withLock_(function() {
        var currentData = readData_();
        currentData.users = currentData.users || {};
        if (body.userObj) {
          currentData.users[targetId] = body.userObj;
        }
        return writeData_(currentData);
      }));
    }

    if (action === "updateTask") {
      return out_(withLock_(function() {
        var currentData2 = readData_();
        currentData2.tasks = currentData2.tasks || [];
        var tIdx = currentData2.tasks.findIndex(function(t) {
          return t.id === body.task.id;
        });

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
        if (body.taskTypes) currentData3.taskTypes = body.taskTypes;
        if (body.taskPrices) currentData3.taskPrices = body.taskPrices;
        if (body.employees) currentData3.employees = body.employees;
        if (body.userHourlyRates) currentData3.userHourlyRates = body.userHourlyRates;
        if (body.staffWorkStatus) currentData3.staffWorkStatus = body.staffWorkStatus;
        if (body.lockedMonths) currentData3.lockedMonths = body.lockedMonths;

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