/* ================================================================ */
/* common.js - 業務管理アプリ 共通コード                              */
/* ================================================================ */

const LS_KEY="stampcard_v7_clean";

/* ================================================================ */
/* === GOOGLE DRIVE API SYNC LAYER === */
/* ================================================================ */
const API_URL_KEY = "stampcard_api_url";
const DEFAULT_API_URL = window.APP_CONFIG?.DEFAULT_API_URL || "";
const SYNC_VERSION_KEY = "stampcard_sync_version_v1";
const LAST_SYNCED_KEY = "stampcard_last_synced_v1";
const ACTION_QUEUE_KEY = "stampcard_action_queue_v1";
const MAX_QUEUE_RETRY_DELAY_MS = 30000;
const DEFAULT_STAMP_INCENTIVE_RULES = [
  { every: 25, amount: 5000 },
  { every: 50, amount: 5000 },
  { every: 250, amount: 40000 }
];
let API_URL = DEFAULT_API_URL || localStorage.getItem(API_URL_KEY) || ""; 
localStorage.setItem(API_URL_KEY, API_URL);

const TOKEN_KEY = "stampcard_api_token";
function getToken(){ return sessionStorage.getItem(TOKEN_KEY) || ""; }
function setToken(t){ sessionStorage.setItem(TOKEN_KEY, t); invalidateStaffAccountsCache(); setTimeout(processQueue, 0); }
function clearToken(){ sessionStorage.removeItem(TOKEN_KEY); invalidateStaffAccountsCache(); }

async function downloadDriveFile(fileId, fileName, taskId){
  if(!API_URL) { showModal({title:"API未接続です（⚙で設定）",big:"⚠️"}); return; }
  const t=getToken();
  if(!t){ showModal({title:"未ログインです",big:"⚠️"}); return; }
  const url = API_URL + "?action=download&token=" + encodeURIComponent(t) + "&fileId=" + encodeURIComponent(fileId) + (taskId!=null?"&taskId="+encodeURIComponent(taskId):"");
  const resp = await fetch(url, { redirect:"follow" });
  const r = await resp.json();
  if(!r.ok){ showModal({title:"ダウンロード失敗",sub:(r.error||"unknown"),big:"⚠️"}); return; }
  const bin = Uint8Array.from(atob(r.data), c=>c.charCodeAt(0));
  const blob = new Blob([bin], {type: r.mimeType || "application/octet-stream"});
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = (fileName || r.name || "download");
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(()=>URL.revokeObjectURL(a.href), 2000);
}

async function fetchTodayPasswordForAdmin(forceRefresh=false){
  if(!API_URL) return null;
  const t = getToken();
  if(!t) return null;
  const cacheKey = "todayPassword_" + ymd(new Date());
  if(!forceRefresh){
    const cached = sessionStorage.getItem(cacheKey);
    if(cached != null) return cached || null;
  }
  try{
    const resp = await fetch(API_URL + "?action=todayPassword&token=" + encodeURIComponent(t), { redirect:"follow" });
    const r = await resp.json();
    if(!r || !r.ok) return null;
    const password = r.password || null;
    sessionStorage.setItem(cacheKey, password || "");
    return password;
  }catch(e){
    return null;
  }
}

let _syncVersion = parseInt(localStorage.getItem(SYNC_VERSION_KEY) || "0", 10) || 0;
let _syncTimer = null;
let _isSyncing = false;
let _lastSyncTime = null;
const POLL_INTERVAL = 30000;

function updateSyncUI(status, msg) {
  const dot = document.getElementById("syncDot");
  const msgEl = document.getElementById("syncMsg");
  if (!dot || !msgEl) return;
  dot.className = "sync-dot " + status;
  msgEl.textContent = msg;
  if (_lastSyncTime) {
    const te = document.getElementById("syncTime");
    if (te) te.textContent = "最終同期: " + new Date(_lastSyncTime).toLocaleTimeString("ja-JP");
  }
}

async function testApiConnection(){
  if(!API_URL) return false;
  try{
    const resp = await fetch(API_URL + "?action=ping", { redirect:"follow" });
    const r = await resp.json();
    return !!(r && r.ok && r.ping);
  }catch(e){
    return false;
  }
}

function promptApiUrl() {
  var ov = document.getElementById("apiSetupOverlay");
  var inp = document.getElementById("apiUrlInput");
  var st = document.getElementById("apiSetupStatus");
  if (!ov) return;
  inp.value = API_URL || "";
  st.textContent = API_URL ? "接続中" : "未設定";
  st.style.color = API_URL ? "#6bcb77" : "var(--muted)";
  ov.style.display = "flex";
}

function setupApiModal() {
  var ov = document.getElementById("apiSetupOverlay");
  if (!ov) return;
  document.getElementById("apiSetupClose").addEventListener("click", function(){ ov.style.display="none"; });
  ov.addEventListener("click", function(e){ if(e.target===ov) ov.style.display="none"; });
  
  document.getElementById("apiSetupSave").addEventListener("click", async function(){
    var inp = document.getElementById("apiUrlInput");
    var st = document.getElementById("apiSetupStatus");
    var url = inp.value.trim();
    if (!url) { st.textContent = "URLを入力してください"; st.style.color = "#ff4757"; return; }
    API_URL = url;
    localStorage.setItem(API_URL_KEY, API_URL);
    st.textContent = "接続テスト中..."; st.style.color = "#4d96ff";
    updateSyncUI("loading", "接続テスト中...");
    
    var ok = await testApiConnection();
    if (ok) {
      st.textContent = "接続成功！"; st.style.color = "#6bcb77";
      updateSyncUI("ok", "接続成功");
      startSyncPolling();
      setTimeout(function(){ ov.style.display="none"; }, 1000);
    } else {
      st.textContent = "接続失敗 - URLを確認してください"; st.style.color = "#ff4757";
      updateSyncUI("err", "接続失敗");
    }
  });

  document.getElementById("apiSetupClear").addEventListener("click", function(){
    API_URL = "";
    localStorage.removeItem(API_URL_KEY);
    stopSyncPolling();
    document.getElementById("apiUrlInput").value = "";
    var st = document.getElementById("apiSetupStatus");
    st.textContent = "切断しました"; st.style.color = "var(--muted)";
    updateSyncUI("warn", "オフライン（⚙でAPI設定）");
  });
}

// ==========================================
// ▼▼▼ スマート同期システム（差分自動検知） ▼▼▼
// ==========================================
let lastSyncedDataStr = localStorage.getItem(LAST_SYNCED_KEY) || "";
let actionQueue = loadPersistedActionQueue();
let isSending = false;
let skipNextSaveDataSync = false;
let needsQueueRebuild = false;
let _staffAccountsCache = null;

function stripSensitiveFieldsFromData(clean) {
  const users = (clean && clean.users) || {};
  Object.keys(users).forEach(uid => {
    const user = users[uid];
    if (user && Object.prototype.hasOwnProperty.call(user, "pw")) delete user.pw;
  });
  return clean;
}

function cloneForStorage(src) {
  return stripSensitiveFieldsFromData(JSON.parse(JSON.stringify(src || {})));
}

function normalizeStampIncentiveRules(rules) {
  const source = Array.isArray(rules) ? rules : [];
  const normalized = source.map(rule => ({
    every: Math.max(1, parseInt(rule && rule.every, 10) || 0),
    amount: Math.max(0, parseInt(rule && rule.amount, 10) || 0)
  })).filter(rule => rule.every > 0 && rule.amount > 0);
  if (!normalized.length) {
    return DEFAULT_STAMP_INCENTIVE_RULES.map(rule => ({ every: rule.every, amount: rule.amount }));
  }
  normalized.sort((a, b) => a.every - b.every || a.amount - b.amount);
  return normalized;
}

function getStampIncentiveRules() {
  return normalizeStampIncentiveRules(data && data.stampIncentiveRules);
}

function formatStampIncentiveRule(rule) {
  if (!rule) return "";
  return `${rule.every}ptごとに${rule.amount.toLocaleString()}円`;
}

function describeStampIncentiveRules() {
  return getStampIncentiveRules().map(formatStampIncentiveRule).join(" / ");
}

function serializeDataForSync(src) {
  const clean = cloneForStorage(src);
  delete clean.session;
  delete clean._version;
  delete clean._updatedAt;
  return clean;
}

function loadPersistedActionQueue() {
  try {
    const parsed = JSON.parse(localStorage.getItem(ACTION_QUEUE_KEY) || "[]");
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

function persistActionQueue() {
  localStorage.setItem(ACTION_QUEUE_KEY, JSON.stringify(actionQueue));
}

function persistSyncVersion() {
  localStorage.setItem(SYNC_VERSION_KEY, String(_syncVersion || 0));
}

function persistSyncedBaseline(serialized) {
  lastSyncedDataStr = serialized || "{}";
  localStorage.setItem(LAST_SYNCED_KEY, lastSyncedDataStr);
}

function syncMetaFromResult(result) {
  if (!result || typeof result !== "object") return null;
  return { version: result.version, updatedAt: result.updatedAt };
}

function applySyncMeta(syncMeta) {
  if (syncMeta && syncMeta.version != null) {
    _syncVersion = parseInt(syncMeta.version, 10) || 0;
    persistSyncVersion();
  }
  if (syncMeta && syncMeta.updatedAt) _lastSyncTime = Date.parse(syncMeta.updatedAt) || Date.now();
}

function updateSyncedBaseline(mutator, syncMeta) {
  let baseline;
  try {
    baseline = JSON.parse(lastSyncedDataStr || "{}");
  } catch (e) {
    baseline = {};
  }
  mutator(baseline);
  persistSyncedBaseline(JSON.stringify(serializeDataForSync(baseline)));
  applySyncMeta(syncMeta);
}

function saveLocalOnly(d) {
  localStorage.setItem(LS_KEY, JSON.stringify(cloneForStorage(d)));
}

function saveLocalAsSynced(d, syncMeta) {
  saveLocalOnly(d);
  persistSyncedBaseline(JSON.stringify(serializeDataForSync(d)));
  applySyncMeta(syncMeta);
}

if (!lastSyncedDataStr) {
  try {
    const localRaw = localStorage.getItem(LS_KEY) || "{}";
    persistSyncedBaseline(JSON.stringify(serializeDataForSync(JSON.parse(localRaw))));
  } catch (e) {
    persistSyncedBaseline("{}");
  }
}


function queueAction(action, payload) {
  actionQueue.push({ action, payload, retryCount: 0, nextAttemptAt: 0 });
  persistActionQueue();
  processQueue();
}

function scheduleQueueRetry(req, message) {
  req.retryCount = (req.retryCount || 0) + 1;
  const delay = Math.min(1000 * Math.pow(2, Math.min(req.retryCount, 5)), MAX_QUEUE_RETRY_DELAY_MS);
  req.nextAttemptAt = Date.now() + delay;
  persistActionQueue();
  updateSyncUI("err", (message || "送信エラー") + `（${Math.round(delay / 1000)}秒後に再送）`);
  setTimeout(processQueue, delay);
}

async function handleQueueVersionConflict() {
  updateSyncUI("warn", "別端末の更新を取り込み中...");
  const desiredLocal = JSON.parse(JSON.stringify(data));
  actionQueue = [];
  persistActionQueue();
  const ok = await forceSyncPull();
  if (!ok) {
    updateSyncUI("err", "競合解決の最新取得に失敗");
    return;
  }
  data = desiredLocal;
  saveLocalOnly(data);
  saveData(data);
  updateSyncUI("loading", "最新状態へ差分を再送中...");
}

async function processQueue() {
  if (isSending || actionQueue.length === 0 || !API_URL) return;
  const req = actionQueue[0];
  if (req.nextAttemptAt && req.nextAttemptAt > Date.now()) {
    setTimeout(processQueue, Math.min(req.nextAttemptAt - Date.now(), 1000));
    return;
  }
  isSending = true;
  updateSyncUI("loading", "保存中...");
  try {
    const body = Object.assign({}, req.payload, {
      _action: req.action,
      token: getToken(),
      _baseVersion: _syncVersion
    });
    const resp = await fetch(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain" },
      body: JSON.stringify(body),
      redirect: "follow"
    });
    const result = await resp.json();
    if (result.ok) {
      actionQueue.shift();
      persistActionQueue();
      _syncVersion = result.version || _syncVersion + 1;
      persistSyncVersion();
      _lastSyncTime = Date.now();
      updateSyncUI("ok", "同期済み ✓");
    } else {
      updateSyncUI("err", result.error || "保存エラー");
    }
  } catch (e) {
    updateSyncUI("err", "通信エラー");
  }
  isSending = false;
  if (needsQueueRebuild) {
    needsQueueRebuild = false;
    saveData(data);
    return;
  }
  processQueue();
}

async function syncPush() { return true; }

function sanitizeRemoteData(remoteData) {
  const clean = JSON.parse(JSON.stringify(remoteData || {}));
  delete clean._version;
  delete clean._updatedAt;
  const localSession = (data && data.session)
    ? JSON.parse(JSON.stringify(data.session))
    : { userId: "", adminAuthed: false, adminEditingUserId: "", adminReportEditingUserId: "" };
  clean.session = localSession;
  return clean;
}

function waitForQueueIdle(timeoutMs = 5000) {
  return new Promise(resolve => {
    const start = Date.now();
    (function check() {
      if (!isSending && actionQueue.length === 0) {
        resolve(true);
        return;
      }
      if (Date.now() - start >= timeoutMs) {
        resolve(false);
        return;
      }
      setTimeout(check, 50);
    })();
  });
}

function downloadTaskFiles(task) {
  if (!task || !task.fileNames || !task.fileNames.length || task.fileNames[0] === "（ファイルなし）") return;
  if (task.fileIds && task.fileIds.length) {
    for (let i = 0; i < task.fileIds.length; i++) {
      if (task.fileIds[i]) {
        downloadDriveFile(task.fileIds[i], (task.fileNames && task.fileNames[i]) || "download", task.id);
      }
    }
    return;
  }
  task.fileNames.forEach(fn => showModal({ title: "ダウンロード", sub: fn, big: "📥" }));
}

async function syncPullLegacy() {
  return syncPull();
}

function saveDataLegacy(d) {
  return saveData(d);
}

// ==========================================
// ▲▲▲ スマート同期システムここまで ▲▲▲
// ==========================================

async function syncCheckVersionLegacy() {
  return syncCheckVersion();
}

function startSyncPolling() {
  stopSyncPolling();
  _syncTimer = setInterval(syncCheckVersion, POLL_INTERVAL);
}
function stopSyncPolling() {
  if (_syncTimer) { clearInterval(_syncTimer); _syncTimer = null; }
}

// 強制同期: 重要な画面遷移前に最新データを取得
async function forceSyncPullLegacy() {
  return forceSyncPull();
}

if (actionQueue.length && getToken()) setTimeout(processQueue, 0);

// スタンプ申請の件数を取得
function countPendingStampRequests() {
  let count = 0;
  Object.values(data.users || {}).forEach(function(u) {
    if (u && u.pendingStampRequest && u.pendingStampRequest.status === "pending") count++;
  });
  return count;
}

function migrateData() {
  Object.keys(data.users||{}).forEach(id=>{
    let u = data.users[id]; if(!u) return;
    if(!u.id) u.id = id;
    if(u.bonusPoints==null)u.bonusPoints=0;if(u.lastCongrats50==null)u.lastCongrats50=0;
    if(!u.lastMonthFirstStamp)u.lastMonthFirstStamp="";if(!u.reports)u.reports=[];
    if(!u.proofingIncentives)u.proofingIncentives={};if(!u.userType)u.userType="学生";
    if(!u.pendingStampRequest)u.pendingStampRequest=null;
    u.reports.forEach(r=>{if(!r.workTime||r.workTime===""){
      const sh=parseInt(r.startH)||0,sm=parseInt(r.startM)||0,eh=parseInt(r.endH)||0,em=parseInt(r.endM)||0,brk=parseInt(r.breakTime)||0;
      let d=(eh*60+em)-(sh*60+sm)-brk;if(d<0)d=0;r.workTime=`${Math.floor(d/60)}時間${d%60>0?d%60+"分":""}`}});
  });
  if(!data.tasks)data.tasks=[];
  if(!data.employees)data.employees=[...DEFAULT_EMPLOYEES];
  if(!data.taskTypes)data.taskTypes=[...DEFAULT_TASK_TYPES];
  if(!data.taskPrices)data.taskPrices={...TASK_PRICES};
  if(!data.userHourlyRates)data.userHourlyRates={};
  if(!data.staffWorkStatus)data.staffWorkStatus={};
  data.stampIncentiveRules = normalizeStampIncentiveRules(data.stampIncentiveRules);
  if(!data.session)data.session={userId:"",adminAuthed:false,adminEditingUserId:"",adminReportEditingUserId:""};
  if(!data.session.adminReportEditingUserId)data.session.adminReportEditingUserId="";
  Object.keys(data.staffWorkStatus).forEach(key=>{const u=findUserByStaffRef(key);if(u&&u.id!==key&&data.staffWorkStatus[u.id]==null){data.staffWorkStatus[u.id]=data.staffWorkStatus[key];delete data.staffWorkStatus[key]}});
  data.tasks.forEach(t=>{if(!t.fileNames){t.fileNames=t.fileName?[t.fileName]:[];if(t.fileName)delete t.fileName}});
  data.tasks.forEach(t=>{if(!t.staffUserId){const u=findUserByStaffRef(t.staff);if(u)t.staffUserId=u.id}});
  if(data.taskTypes){data.taskTypes=data.taskTypes.map(t=>t==="その他（時給）"?"時給":t)}
  data.tasks.forEach(t=>{if(t.taskType==="その他（時給）")t.taskType="時給"});
  if(data.taskPrices&&data.taskPrices["その他（時給）"]!=null){data.taskPrices["時給"]=data.taskPrices["その他（時給）"];delete data.taskPrices["その他（時給）"]}
}

function initSync() {
  setupApiModal();
  const setupBtn = document.getElementById("syncSetup");
  if (setupBtn) setupBtn.addEventListener("click", promptApiUrl);
  const manualBtn = document.getElementById("syncManual");
  if (manualBtn) manualBtn.addEventListener("click", async () => {
    updateSyncUI("loading", "手動同期中...");
    await syncPull();
  });

  if (API_URL) {
    updateSyncUI("loading", "接続中...");
    testApiConnection().then(ok => {
      if (ok) startSyncPolling();
      else updateSyncUI("err", "接続エラー - ⚙で設定確認");
    });
  } else {
    updateSyncUI("warn", "オフライン（⚙でAPI設定）");
  }
}

/* === DRIVE FILE UPLOAD === */
async function uploadFileToDrive(file, taskId) {
  if (!API_URL) return null;
  return new Promise(function(resolve) {
    var reader = new FileReader();
    reader.onload = async function() {
      try {
        var b64 = reader.result.split(",")[1];
        var resp = await fetch(API_URL, {
          method: "POST",
          headers: { "Content-Type": "text/plain" },
          body: JSON.stringify({ _action: "uploadFile", token: getToken(), taskId: taskId || "", fileName: file.name, mimeType: file.type || "application/octet-stream", data: b64 }),
          redirect: "follow"
        });
        var result = await resp.json();
        if (result.ok) { resolve({ fileId: result.fileId, fileName: result.fileName, url: result.url }); }
        else { console.warn("Upload error:", result.error); resolve(null); }
      } catch(e) { console.warn("Upload error:", e); resolve(null); }
    };
    reader.readAsDataURL(file);
  });
}
/* === END SYNC LAYER === */

function applyRemoteSyncData(remoteData) {
  const clean = sanitizeRemoteData(remoteData || {});
  data = JSON.parse(JSON.stringify(clean));
  saveLocalAsSynced(data, {
    version: remoteData && remoteData._version != null ? remoteData._version : _syncVersion,
    updatedAt: remoteData && remoteData._updatedAt ? remoteData._updatedAt : null
  });
  migrateData();
  _lastSyncTime = Date.now();
  return clean;
}

async function processQueue() {
  if (isSending || actionQueue.length === 0 || !API_URL) return;
  const req = actionQueue[0];
  if (req.nextAttemptAt && req.nextAttemptAt > Date.now()) {
    setTimeout(processQueue, Math.min(req.nextAttemptAt - Date.now(), 1000));
    return;
  }
  isSending = true;
  updateSyncUI("loading", "同期送信中...");
  try {
    const body = Object.assign({}, req.payload, {
      _action: req.action,
      token: getToken(),
      _baseVersion: _syncVersion
    });
    const resp = await fetch(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain" },
      body: JSON.stringify(body),
      redirect: "follow"
    });
    const result = await resp.json();
    if (result.ok) {
      actionQueue.shift();
      persistActionQueue();
      _syncVersion = result.version || (_syncVersion + 1);
      persistSyncVersion();
      _lastSyncTime = Date.now();
      updateSyncUI("ok", "同期済み");
      if (actionQueue.length === 0 && !needsQueueRebuild) {
        saveLocalAsSynced(data, syncMetaFromResult(result));
      }
    } else if (result.error === "version_conflict") {
      await handleQueueVersionConflict();
    } else if (result.error === "locked_month") {
      actionQueue.shift();
      persistActionQueue();
      showModal({ title: "確定済みの月です", sub: "この操作は月次確定後のため実行できません", big: "NG" });
    } else {
      scheduleQueueRetry(req, result.error || "同期エラー");
    }
  } catch (e) {
    scheduleQueueRetry(req, "通信エラー");
  }
  isSending = false;
  processQueue();
}

async function syncPull() {
  if (!API_URL || _isSyncing) return false;
  const idle = await waitForQueueIdle();
  if (!idle) return false;
  _isSyncing = true;
  try {
    const resp = await fetch(API_URL + "?action=read&token=" + encodeURIComponent(getToken()), { redirect: "follow" });
    const result = await resp.json();
    if (result.ok && result.data) {
      const remoteVer = result.data._version || 0;
      if (remoteVer > _syncVersion) {
        applyRemoteSyncData(result.data);
        updateSyncUI("ok", "同期済み");
        return true;
      }
      _lastSyncTime = Date.now();
      updateSyncUI("ok", "最新");
      return false;
    }
    updateSyncUI("err", "読み込みエラー");
    return false;
  } catch (e) {
    updateSyncUI("err", "通信エラー");
    return false;
  } finally {
    _isSyncing = false;
  }
}

function saveData(d) {
  if (skipNextSaveDataSync) {
    skipNextSaveDataSync = false;
    saveLocalAsSynced(d);
    return;
  }

  saveLocalOnly(d);
  if (isSending) {
    needsQueueRebuild = true;
    return;
  }
  const oldD = JSON.parse(lastSyncedDataStr || "{}");
  const newD = serializeDataForSync(d);
  const nextQueue = [];

  const oldTasksById = new Map((oldD.tasks || []).map(t => [t.id, t]));
  const newTasksById = new Map((newD.tasks || []).map(t => [t.id, t]));
  newTasksById.forEach((nT, id) => {
    const oT = oldTasksById.get(id);
    if (!oT || JSON.stringify(oT) !== JSON.stringify(nT)) {
      nextQueue.push({ action: "updateTask", payload: { task: nT, isDelete: false }, retryCount: 0, nextAttemptAt: 0 });
    }
  });
  oldTasksById.forEach((oT, id) => {
    if (!newTasksById.has(id)) {
      nextQueue.push({ action: "updateTask", payload: { task: { id: oT.id }, isDelete: true }, retryCount: 0, nextAttemptAt: 0 });
    }
  });

  const oldUsers = oldD.users || {};
  const newUsers = newD.users || {};
  Object.keys(newUsers).forEach(uid => {
    if (JSON.stringify(oldUsers[uid]) !== JSON.stringify(newUsers[uid])) {
      nextQueue.push({ action: "updateUserFull", payload: { targetUserId: uid, userObj: newUsers[uid] }, retryCount: 0, nextAttemptAt: 0 });
    }
  });

  if (JSON.stringify(oldD.taskTypes) !== JSON.stringify(newD.taskTypes)) nextQueue.push({ action: "updateMaster", payload: { taskTypes: newD.taskTypes }, retryCount: 0, nextAttemptAt: 0 });
  if (JSON.stringify(oldD.taskPrices) !== JSON.stringify(newD.taskPrices)) nextQueue.push({ action: "updateMaster", payload: { taskPrices: newD.taskPrices }, retryCount: 0, nextAttemptAt: 0 });
  if (JSON.stringify(oldD.employees) !== JSON.stringify(newD.employees)) nextQueue.push({ action: "updateMaster", payload: { employees: newD.employees }, retryCount: 0, nextAttemptAt: 0 });
  if (JSON.stringify(oldD.userHourlyRates) !== JSON.stringify(newD.userHourlyRates)) nextQueue.push({ action: "updateMaster", payload: { userHourlyRates: newD.userHourlyRates }, retryCount: 0, nextAttemptAt: 0 });
  if (JSON.stringify(oldD.staffWorkStatus) !== JSON.stringify(newD.staffWorkStatus)) nextQueue.push({ action: "updateMaster", payload: { staffWorkStatus: newD.staffWorkStatus }, retryCount: 0, nextAttemptAt: 0 });
  if (JSON.stringify(oldD.lockedMonths) !== JSON.stringify(newD.lockedMonths)) nextQueue.push({ action: "updateMaster", payload: { lockedMonths: newD.lockedMonths }, retryCount: 0, nextAttemptAt: 0 });
  if (JSON.stringify(oldD.stampIncentiveRules) !== JSON.stringify(newD.stampIncentiveRules)) nextQueue.push({ action: "updateMaster", payload: { stampIncentiveRules: newD.stampIncentiveRules }, retryCount: 0, nextAttemptAt: 0 });
  if (JSON.stringify(oldD.notices) !== JSON.stringify(newD.notices)) nextQueue.push({ action: "updateMaster", payload: { notices: newD.notices }, retryCount: 0, nextAttemptAt: 0 });

  const deletedUids = Object.keys(oldUsers).filter(uid => !newUsers[uid]);
  if (deletedUids.length) nextQueue.push({ action: "updateMaster", payload: { deleteUserId: deletedUids }, retryCount: 0, nextAttemptAt: 0 });

  actionQueue = nextQueue;
  persistActionQueue();
  processQueue();
}

async function syncCheckVersion() {
  if (!API_URL || _isSyncing || isSending || actionQueue.length) return;
  try {
    const resp = await fetch(API_URL + "?action=version&token=" + encodeURIComponent(getToken()), { redirect: "follow" });
    const result = await resp.json();
    if (result.ok && (result.version || 0) > _syncVersion) {
      await syncPull();
    }
  } catch (e) {
  }
}

async function forceSyncPull() {
  if (!API_URL) return false;
  await waitForQueueIdle();
  const wasSyncing = _isSyncing;
  _isSyncing = true;
  try {
    const resp = await fetch(API_URL + "?action=read&token=" + encodeURIComponent(getToken()), { redirect: "follow" });
    const result = await resp.json();
    if (result.ok && result.data) {
      applyRemoteSyncData(result.data);
      updateSyncUI("ok", "同期済み");
      return true;
    }
    return false;
  } catch (e) {
    return false;
  } finally {
    _isSyncing = wasSyncing;
  }
}

async function postDirectAction(action, payload) {
  if (!API_URL) throw new Error("API not configured");
  const token = getToken();
  if (!token) throw new Error("not logged in");
  const resp = await fetch(API_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain" },
    body: JSON.stringify(Object.assign({
      _action: action,
      token,
      _baseVersion: _syncVersion
    }, payload || {})),
    redirect: "follow"
  });
  const result = await resp.json();
  if (result && result.ok) {
    if (result.version != null) {
      _syncVersion = parseInt(result.version, 10) || _syncVersion;
      persistSyncVersion();
    }
    if (result.updatedAt) _lastSyncTime = Date.parse(result.updatedAt) || Date.now();
    return result;
  }
  if (result && result.error === "version_conflict") {
    await forceSyncPull();
    throw new Error("version_conflict");
  }
  throw new Error((result && result.error) || "request_failed");
}

function cloneDeep(value) {
  return JSON.parse(JSON.stringify(value));
}

function ensureLocalUser(userId) {
  data.users = data.users || {};
  data.users[userId] = data.users[userId] || { id: userId, stamps: {}, reports: [], pendingStampRequest: null };
  ensureUserShape(data.users[userId]);
  return data.users[userId];
}

function ensureLocalReportId(report) {
  if (!report) return "";
  if (!report.reportId) report.reportId = "report_local_" + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 8);
  return report.reportId;
}

function findLocalReportIndex(user, reportId, fallbackIndex) {
  if (!user || !Array.isArray(user.reports)) return -1;
  if (reportId) {
    const idx = user.reports.findIndex(r => String((r && r.reportId) || "") === String(reportId));
    if (idx >= 0) return idx;
  }
  return typeof fallbackIndex === "number" ? fallbackIndex : -1;
}

function applyDirectUserSync(userId, nextUser, result) {
  data.users = data.users || {};
  data.users[userId] = cloneDeep(nextUser);
  saveLocalOnly(data);
  updateSyncedBaseline(baseline => {
    baseline.users = baseline.users || {};
    baseline.users[userId] = cloneDeep(nextUser);
  }, syncMetaFromResult(result));
}

async function fetchStaffAccountsForAdmin(forceRefresh = false) {
  if (!API_URL) throw new Error("API not configured");
  const token = getToken();
  if (!token) throw new Error("not logged in");
  if (!forceRefresh && Array.isArray(_staffAccountsCache)) return cloneDeep(_staffAccountsCache);
  const resp = await fetch(API_URL + "?action=listStaffAccounts&token=" + encodeURIComponent(token), { redirect: "follow" });
  const result = await resp.json();
  if (!result || !result.ok || !Array.isArray(result.staffAccounts)) {
    throw new Error((result && result.error) || "staff_accounts_fetch_failed");
  }
  _staffAccountsCache = result.staffAccounts.map(acc => Object.assign({}, acc));
  return cloneDeep(_staffAccountsCache);
}

function invalidateStaffAccountsCache() {
  _staffAccountsCache = null;
}

async function fetchStaffPasswordForAdmin(userId, forceRefresh = false) {
  const accounts = await fetchStaffAccountsForAdmin(forceRefresh);
  const found = accounts.find(acc => String(acc.id || "") === String(userId || ""));
  return found ? String(found.pw || "") : "";
}

function fillStaffPasswordField(fieldId, userId, forceRefresh = false) {
  const input = document.getElementById(fieldId);
  if (!input) return;
  input.value = "";
  if (!userId) return;
  fetchStaffPasswordForAdmin(userId, forceRefresh)
    .then(pw => {
      if (document.getElementById(fieldId) !== input) return;
      input.value = pw || "";
    })
    .catch(() => {
      if (document.getElementById(fieldId) !== input) return;
      input.value = "";
    });
}

function isUnsupportedDirectActionError(error) {
  const message = String(error && error.message || "").toLowerCase();
  return message.indexOf("unknown action") >= 0;
}

function saveLegacyDirectUser(targetUserId, mutator) {
  const user = ensureLocalUser(targetUserId);
  mutator(user);
  saveData(data);
  return cloneDeep(user);
}

async function setStampRemote(targetUserId, date, value, metaPatch) {
  try {
    const result = await postDirectAction("setStamp", { targetUserId, date, value, metaPatch: metaPatch || null });
    if (result.user) applyDirectUserSync(targetUserId, result.user, result);
    return result;
  } catch (error) {
    if (!isUnsupportedDirectActionError(error)) throw error;
    const user = saveLegacyDirectUser(targetUserId, u => {
      if (value == null || value === "") delete u.stamps[date];
      else u.stamps[date] = value;
      if (metaPatch) {
        Object.keys(metaPatch).forEach(key => {
          const nextValue = metaPatch[key];
          if (nextValue == null || nextValue === "") delete u[key];
          else u[key] = nextValue;
        });
      }
    });
    return { ok: true, user, legacyFallback: true };
  }
}

async function setStampMetaRemote(targetUserId, metaPatch) {
  try {
    const result = await postDirectAction("setStampMeta", { targetUserId, metaPatch: metaPatch || {} });
    if (result.user) applyDirectUserSync(targetUserId, result.user, result);
    return result;
  } catch (error) {
    if (!isUnsupportedDirectActionError(error)) throw error;
    const user = saveLegacyDirectUser(targetUserId, u => {
      Object.keys(metaPatch || {}).forEach(key => {
        const nextValue = metaPatch[key];
        if (nextValue == null || nextValue === "") delete u[key];
        else u[key] = nextValue;
      });
    });
    return { ok: true, user, legacyFallback: true };
  }
}

async function requestStampCorrectionRemote(targetUserId, stamps) {
  try {
    const result = await postDirectAction("requestStampCorrection", { targetUserId, stamps: cloneDeep(stamps || {}) });
    if (result.user) applyDirectUserSync(targetUserId, result.user, result);
    return result;
  } catch (error) {
    if (!isUnsupportedDirectActionError(error)) throw error;
    const user = saveLegacyDirectUser(targetUserId, u => {
      u.pendingStampRequest = { stamps: cloneDeep(stamps || {}), status: "pending", createdAt: Date.now() };
    });
    return { ok: true, user, legacyFallback: true };
  }
}

async function updateStampRequestDraftRemote(targetUserId, stamps) {
  try {
    const result = await postDirectAction("updateStampRequestDraft", { targetUserId, stamps: cloneDeep(stamps || {}) });
    if (result.user) applyDirectUserSync(targetUserId, result.user, result);
    return result;
  } catch (error) {
    if (!isUnsupportedDirectActionError(error)) throw error;
    const user = saveLegacyDirectUser(targetUserId, u => {
      u.pendingStampRequest = u.pendingStampRequest || {};
      u.pendingStampRequest.stamps = cloneDeep(stamps || {});
      if (!u.pendingStampRequest.status) u.pendingStampRequest.status = "pending";
    });
    return { ok: true, user, legacyFallback: true };
  }
}

async function resolveStampCorrectionRemote(targetUserId, decision) {
  try {
    const result = await postDirectAction("resolveStampCorrection", { targetUserId, decision });
    if (result.user) applyDirectUserSync(targetUserId, result.user, result);
    return result;
  } catch (error) {
    if (!isUnsupportedDirectActionError(error)) throw error;
    const user = saveLegacyDirectUser(targetUserId, u => {
      if (decision === "approved" && u.pendingStampRequest && u.pendingStampRequest.stamps) {
        u.stamps = cloneDeep(u.pendingStampRequest.stamps);
        u.pendingStampRequest = { status: "approved", resolvedAt: Date.now() };
      } else if (decision === "rejected") {
        u.pendingStampRequest = { status: "rejected", resolvedAt: Date.now() };
      }
    });
    return { ok: true, user, legacyFallback: true };
  }
}

async function clearStampRequestStateRemote(targetUserId) {
  try {
    const result = await postDirectAction("clearStampRequestState", { targetUserId });
    if (result.user) applyDirectUserSync(targetUserId, result.user, result);
    return result;
  } catch (error) {
    if (!isUnsupportedDirectActionError(error)) throw error;
    const user = saveLegacyDirectUser(targetUserId, u => { u.pendingStampRequest = null; });
    return { ok: true, user, legacyFallback: true };
  }
}

async function addReportRemote(targetUserId, report) {
  try {
    const result = await postDirectAction("addReport", { targetUserId, report: cloneDeep(report || {}) });
    if (result.user) applyDirectUserSync(targetUserId, result.user, result);
    return result.report || null;
  } catch (error) {
    if (!isUnsupportedDirectActionError(error)) throw error;
    const nextReport = cloneDeep(report || {});
    ensureLocalReportId(nextReport);
    const user = saveLegacyDirectUser(targetUserId, u => {
      u.reports = Array.isArray(u.reports) ? u.reports : [];
      u.reports.push(nextReport);
    });
    return cloneDeep((user.reports || []).find(r => r.reportId === nextReport.reportId) || nextReport);
  }
}

async function addReportsBatchRemote(targetUserId, reports) {
  const nextReports = Array.isArray(reports) ? reports.filter(Boolean).map(report => cloneDeep(report)) : [];
  if (!nextReports.length) return [];
  try {
    const result = await postDirectAction("addReportsBatch", { targetUserId, reports: nextReports });
    if (result.user) applyDirectUserSync(targetUserId, result.user, result);
    return Array.isArray(result.reports) ? result.reports : [];
  } catch (error) {
    if (!isUnsupportedDirectActionError(error)) throw error;
    const added = [];
    for (const report of nextReports) added.push(await addReportRemote(targetUserId, report));
    return added;
  }
}

async function updateReportRemote(targetUserId, reportId, reportIndex, patch) {
  try {
    const result = await postDirectAction("updateReport", {
      targetUserId,
      reportId: reportId || "",
      reportIndex,
      patch: cloneDeep(patch || {})
    });
    if (result.user) applyDirectUserSync(targetUserId, result.user, result);
    return result.report || null;
  } catch (error) {
    if (!isUnsupportedDirectActionError(error)) throw error;
    let updated = null;
    saveLegacyDirectUser(targetUserId, u => {
      u.reports = Array.isArray(u.reports) ? u.reports : [];
      const idx = findLocalReportIndex(u, reportId, reportIndex);
      if (idx < 0) return;
      const nextReport = Object.assign({}, u.reports[idx] || {}, cloneDeep(patch || {}));
      ensureLocalReportId(nextReport);
      u.reports[idx] = nextReport;
      updated = cloneDeep(nextReport);
    });
    return updated;
  }
}

async function deleteReportRemote(targetUserId, reportId, reportIndex) {
  try {
    const result = await postDirectAction("deleteReport", {
      targetUserId,
      reportId: reportId || "",
      reportIndex
    });
    if (result.user) applyDirectUserSync(targetUserId, result.user, result);
    return true;
  } catch (error) {
    if (!isUnsupportedDirectActionError(error)) throw error;
    saveLegacyDirectUser(targetUserId, u => {
      u.reports = Array.isArray(u.reports) ? u.reports : [];
      const idx = findLocalReportIndex(u, reportId, reportIndex);
      if (idx >= 0) u.reports.splice(idx, 1);
    });
    return true;
  }
}

async function setReportReviewRemote(targetUserId, reportId, reportIndex, reviewPatch) {
  try {
    const result = await postDirectAction("setReportReview", Object.assign({
      targetUserId,
      reportId: reportId || "",
      reportIndex
    }, cloneDeep(reviewPatch || {})));
    if (result.user) applyDirectUserSync(targetUserId, result.user, result);
    return result.report || null;
  } catch (error) {
    if (!isUnsupportedDirectActionError(error)) throw error;
    let updated = null;
    saveLegacyDirectUser(targetUserId, u => {
      u.reports = Array.isArray(u.reports) ? u.reports : [];
      const idx = findLocalReportIndex(u, reportId, reportIndex);
      if (idx < 0) return;
      const nextReport = Object.assign({}, u.reports[idx] || {}, cloneDeep(reviewPatch || {}));
      ensureLocalReportId(nextReport);
      u.reports[idx] = nextReport;
      updated = cloneDeep(nextReport);
    });
    return updated;
  }
}

async function changeOwnStaffPasswordRemote(currentPw, newPw) {
  try {
    const result = await postDirectAction("changeOwnStaffPassword", { currentPw, newPw });
    if (result.user && data.session && data.session.userId) {
      applyDirectUserSync(data.session.userId, result.user, result);
    }
    return result;
  } catch (error) {
    if (isUnsupportedDirectActionError(error)) {
      throw new Error("password_change_requires_gas_update");
    }
    throw error;
  }
}

function buildReportPayloadFromForm(transportValueOverride) {
  const digitsOnly = value => String(value == null ? "" : value).replace(/[^0-9]/g, "");
  const padTime = (value, max) => {
    const num = Math.max(0, Math.min(max, parseInt(digitsOnly(value), 10) || 0));
    return pad2(num);
  };
  const wt = $("rpWorkType").value;
  const report = {
    date: $("rpDate").value,
    workType: wt,
    startH: padTime($("rpStartH").value, 23),
    startM: padTime($("rpStartM").value, 59),
    endH: padTime($("rpEndH").value, 23),
    endM: padTime($("rpEndM").value, 59),
    breakTime: digitsOnly($("rpBreak").value),
    workTime: $("rpWorkTime").value,
    content: $("rpContent").value
  };
  if (wt === "在宅") {
    report.taskType = $("rpTaskType").value;
    report.manHours = $("rpManHours").value;
  } else {
    report.transport = transportValueOverride != null ? transportValueOverride : $("rpTransport").value;
    report.bizId = $("rpBizId").value;
    report.productId = $("rpProductId").value;
    report.serviceId = $("rpServiceId").value;
    report.textCode = $("rpTextCode").value;
    report.year = $("rpYear").value;
  }
  return report;
}

function handleDirectActionError(error, fallbackMessage) {
  if (error && error.message === "version_conflict") {
    showModal({ title: "最新データに更新しました", sub: "もう一度操作してください", big: "🔄" });
    return;
  }
  showModal({ title: "通信エラー", sub: (error && error.message) || fallbackMessage || "保存に失敗しました", big: "📡" });
}


/* ================================================================ */
/* === ユーティリティ・ヘルパー === */
/* ================================================================ */
// ⑮修正: fileIds の初期化を修正
function ensureUserShape(u){
  if(!u) return;
  u.id = u.id || data.session.userId || u.id;
  u.stamps = u.stamps || {};
  u.reports = u.reports || [];
  u.tasks = u.tasks || [];
  u.pendingStampRequest = u.pendingStampRequest || null;
  u.fileIds = u.fileIds || [];
}

const RANK_EMOJIS=["⭐","🌟","💫","🔥","💎","👑","🏆","🎖️","💜","🌈","🚀","✨"];
function getMilestoneCount(total){if(total<200)return Math.floor(total/25);return 8+Math.floor((total-200)/50)}
function getNextMilestone(total){if(total<200)return Math.ceil((total+1)/25)*25;return 200+Math.ceil((total-200+1)/50)*50}
function getRank(total){const mc=getMilestoneCount(total);const emoji=RANK_EMOJIS[Math.min(mc,RANK_EMOJIS.length-1)];const yen=300+mc*50;return{rank:mc+1,yen,label:`ランク${mc+1}`,emoji}}
const MONTHLY_COMMENTS=[{min:0,max:0,msg:"今月出勤してくれてありがとう。"},{min:1,max:3,msg:"毎月出勤してくれてありがとう。"},{min:4,max:10,msg:"たくさん出勤してくれてありがとうございます。"},{min:11,max:16,msg:"いつも出勤していただき非常に助かっております。"},{min:17,max:999,msg:"数学科一同大変感謝しております。"}];
const DEFAULT_TASK_TYPES=["模試校正(1問)","テキスト校正(1問)","テキスト校正(1講)","確認テストチェック(1講)","確認テスト作題(A1講)","確認テスト作題(BC1講)","修了判定テストチェック(1セット)","修了判定テスト作題(A1セット)","修了判定テスト作題(BC1セット)","テキスト入力(解答解説あり)","テキスト入力(解答解説なし)","添削(1問)","web採点基準作成(1問)","全体概観作成(1試験種)","時給","共通テスト模試校正","東大・京大模試校正","早慶模試校正","国公立・関関同立・明青立法中模試校正","全国統一中学生テスト校正"];
const TASK_PRICES={"模試校正(1問)":500,"テキスト校正(1問)":500,"テキスト校正(1講)":4000,"確認テストチェック(1講)":500,"確認テスト作題(A1講)":1500,"確認テスト作題(BC1講)":1000,"修了判定テストチェック(1セット)":2000,"修了判定テスト作題(A1セット)":5000,"修了判定テスト作題(BC1セット)":3000,"テキスト入力(解答解説あり)":3000,"テキスト入力(解答解説なし)":1000,"添削(1問)":300,"web採点基準作成(1問)":300,"全体概観作成(1試験種)":500,"時給":0,"共通テスト模試校正":4000,"東大・京大模試校正":6000,"早慶模試校正":8000,"国公立・関関同立・明青立法中模試校正":11000,"全国統一中学生テスト校正":2000};
const HOURLY_RATE=1300;
const BIZ_IDS=["01 企画立案","02 番組構成","03 制作","04 収録","05 ナレーション収録","08 分析","09 検証","10 メンテナンス","11 運営","17 打合せ・ミーティング","25 添削","26 採点","27 成績処理"];
const PRODUCT_IDS=["00 全般","01 講座","02 テキスト","03 確認テスト","04 講座修了判定テスト","05 授業制作","07 模試","08 その他教材・コンテンツ","11 データベース","12 資料","19 答案","22 収録立会い","23 版下管理"];
const SERVICE_IDS=["000 全般","001 講座（HS）","002 講座（中等部）","003 講座（四谷大塚）","004 模試（HS）","005 模試（中等部）","006 模試（四谷大塚）","007 バックアップサービス","008 東大特進","009 答案練習講座","010 解答速報","011 過去問データベース","012 公開授業","013 リメディアル講座","014 千題テスト","015 夏期合宿","016 冬期合宿","017 正月特訓","018 高速マスター","019 ビジネススクール","020 パンフレット","021 東進タイムズ","022 講座提案シート","023 講座系統図","024 東進進学情報","025 研修・全国大会","026 過去問演習センター","027 過去問演習国立","028 探求・リーダー","029 中学部","030 JFA福島","031 四谷（復習ナビ）","032 四谷全国統一小学生テスト","033 四谷（その他）","034 TOEIC","035 模試部関連","036 教務部関連","037 ビジネススクール関連","038 ハイスクール開発","039 衛星関連","040 研修関連","041 イトマン","042 その他部署関連"];

const DAILY_PASSWORDS = [];
function getTodayPassword(){return ""}

const DEFAULT_EMPLOYEES=["荒金","貝沼","勝原","高橋","田中"];

const _noop=document.createElement('div');const $=id=>document.getElementById(id)||_noop;
const pad2=n=>String(n).padStart(2,"0");
const ymd=d=>`${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`;
const ym=d=>`${d.getFullYear()}-${pad2(d.getMonth()+1)}`;
function monthLabelJa(d){return`${d.getFullYear()}年${d.getMonth()+1}月`}
function startOfMonth(d){return new Date(d.getFullYear(),d.getMonth(),1)}
function endOfMonth(d){return new Date(d.getFullYear(),d.getMonth()+1,0)}
function addMonths(d,n){return new Date(d.getFullYear(),d.getMonth()+n,1)}
function isSameDay(a,b){return a.getFullYear()===b.getFullYear()&&a.getMonth()===b.getMonth()&&a.getDate()===b.getDate()}
function dowJa(d){return["日","月","火","水","木","金","土"][d.getDay()]}
function startOfWeekMon(d){const x=new Date(d.getFullYear(),d.getMonth(),d.getDate());const day=(x.getDay()+6)%7;x.setDate(x.getDate()-day);x.setHours(0,0,0,0);return x}
function endOfWeekMon(d){const s=startOfWeekMon(d);const e=new Date(s);e.setDate(e.getDate()+6);e.setHours(23,59,59,999);return e}
function between(d,a,b){const t=d.getTime();return t>=a.getTime()&&t<=b.getTime()}
function escapeHtml(s){return String(s).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[c]))}
function addDays(d,n){const r=new Date(d);r.setDate(r.getDate()+n);return r}

/* === DATA === */
function loadData(){
  const raw=localStorage.getItem(LS_KEY);if(raw){try{return JSON.parse(raw)}catch(e){}}
  const mkU=(id,name,type,stamps,bp,reports)=>({id,name,userType:type,stamps:stamps||{},incentives:{},bonusPoints:0,lastCongrats50:0,lastMonthFirstStamp:"",reports:reports||[],createdAt:Date.now()-Math.random()*86400000*30,proofingIncentives:{},pendingStampRequest:null});
  const users= {};
  const tasks=[];
  return{users,session:{userId:"",adminAuthed:false,adminEditingUserId:"",adminReportEditingUserId:""},
    tasks,employees:[...DEFAULT_EMPLOYEES],taskTypes:[...DEFAULT_TASK_TYPES],taskPrices:{...TASK_PRICES},staffWorkStatus:{},stampIncentiveRules:DEFAULT_STAMP_INCENTIVE_RULES.map(rule=>({ every: rule.every, amount: rule.amount }))};
}
var data=loadData();
migrateData();

function getUserHourlyRate(userId){return data.userHourlyRates&&data.userHourlyRates[userId]!=null?data.userHourlyRates[userId]:HOURLY_RATE}
// ⑩修正: 初回はlocalStorageのみ保存（APIへの差分送信はスキップ）
localStorage.setItem(LS_KEY, JSON.stringify(data));
if (!lastSyncedDataStr) persistSyncedBaseline(JSON.stringify(serializeDataForSync(data)));

function getTaskPrice(name){return data.taskPrices&&data.taskPrices[name]!=null?data.taskPrices[name]:(TASK_PRICES[name]!=null?TASK_PRICES[name]:null)}
function getTaskTypes(){return data.taskTypes||DEFAULT_TASK_TYPES}
function getEmployees(){return data.employees||DEFAULT_EMPLOYEES}
function getUserDisplayName(u){return u?(u.name||u.id||""):""}
function getStaffUsers(){return Object.values(data.users||{}).filter(Boolean)}
function getStaffNames(){return getStaffUsers().map(getUserDisplayName)}
function findUserByStaffRef(staffRef){
  if(staffRef==null)return null;
  const ref=String(staffRef);
  return getStaffUsers().find(u=>String(u.id||"")===ref||getUserDisplayName(u)===ref)||null
}
function getUserTypeByStaffName(staffName){
  const u=findUserByStaffRef(staffName);
  return u?u.userType:"";
}
function getTaskStaffUserId(task){
  if(!task)return"";
  if(task.staffUserId!=null&&task.staffUserId!=="")return String(task.staffUserId);
  const u=findUserByStaffRef(task.staff);
  return u?String(u.id):"";
}
function getTaskStaffLabel(task){
  const u=findUserByStaffRef(task&&(task.staffUserId||task.staff));
  if(u)return getUserDisplayName(u);
  return task&&task.staff?task.staff:"";
}
function taskMatchesStaffRef(task,staffRef){
  if(!staffRef||staffRef==="全て")return true;
  const ref=String(staffRef);
  return getTaskStaffUserId(task)===ref||getTaskStaffLabel(task)===ref||String(task&&task.staff||"")===ref
}
function setTaskStaffRef(task,staffRef){
  if(!task)return;
  const ref=staffRef==null?"":String(staffRef);
  if(!ref||ref==="未指定"){task.staff="未指定";delete task.staffUserId;return}
  const u=findUserByStaffRef(ref);
  if(u){task.staffUserId=u.id;task.staff=getUserDisplayName(u);return}
  task.staff=ref;delete task.staffUserId;
}
function normalizeTaskTextCodes(textCodes){
  if(Array.isArray(textCodes))return textCodes.map(v=>String(v||"").trim()).filter(Boolean);
  return String(textCodes||"").split(",").map(v=>v.trim()).filter(Boolean);
}
/*
function applyTaskDraft(task,values){
  if(!task)task={};
  values=values||{};
  task.workType=String(values.workType||task.workType||"蜃ｺ蜍､");
  task.status=String(values.status||task.status||"萓晞ｼ蜑・);
  task.requestDate=String(values.requestDate||"");
  task.deadline=String(values.deadline||"");
  task.completionDate=String(values.completionDate||"");
  task.manHours=Math.max(1,parseInt(values.manHours,10)||1);
  task.textCodes=normalizeTaskTextCodes(values.textCodes);
  task.taskType=String(values.taskType||"");
  task.content=String(values.content||"");
  task.employee=String(values.employee||"");
  task.notes=String(values.notes||"");
  if(task.validPointCount==null)task.validPointCount=0;
  if(!Array.isArray(task.vpEditHistory))task.vpEditHistory=[];
  if(!Array.isArray(task.fileNames))task.fileNames=[];
  if(!Array.isArray(task.fileIds))task.fileIds=[];
  setTaskStaffRef(task,values.staff);
  return task;
}
function createTaskDraft(values){
  values=values||{};
  const workType=String(values.workType||"蜃ｺ蜍､");
  const task={
    id:values.id!=null?values.id:Date.now(),
    seqNum:values.seqNum!=null?values.seqNum:nextSeqNum(workType)
  };
  return applyTaskDraft(task,Object.assign({},values,{workType:workType}));
}
*/
function applyTaskDraft(task,values){
  if(!task)task={};
  values=values||{};
  task.workType=String(values.workType||task.workType||"出勤");
  task.status=String(values.status||task.status||"依頼前");
  task.requestDate=String(values.requestDate||"");
  task.deadline=String(values.deadline||"");
  task.completionDate=String(values.completionDate||"");
  task.manHours=Math.max(1,parseInt(values.manHours,10)||1);
  task.textCodes=normalizeTaskTextCodes(values.textCodes);
  task.taskType=String(values.taskType||"");
  task.content=String(values.content||"");
  task.employee=String(values.employee||"");
  task.notes=String(values.notes||"");
  if(task.validPointCount==null)task.validPointCount=0;
  if(!Array.isArray(task.vpEditHistory))task.vpEditHistory=[];
  if(!Array.isArray(task.fileNames))task.fileNames=[];
  if(!Array.isArray(task.fileIds))task.fileIds=[];
  setTaskStaffRef(task,values.staff);
  return task;
}
function createTaskDraft(values){
  values=values||{};
  const workType=String(values.workType||"出勤");
  const task={
    id:values.id!=null?values.id:Date.now(),
    seqNum:values.seqNum!=null?values.seqNum:nextSeqNum(workType)
  };
  return applyTaskDraft(task,Object.assign({},values,{workType:workType}));
}
function getStaffWorkStatusForRef(staffRef){
  if(!data.staffWorkStatus)return"";
  const u=findUserByStaffRef(staffRef);
  if(!u)return data.staffWorkStatus[staffRef]||"";
  const label=getUserDisplayName(u);
  return data.staffWorkStatus[u.id]||data.staffWorkStatus[label]||"";
}
function setStaffWorkStatusForRef(staffRef,value){
  data.staffWorkStatus=data.staffWorkStatus||{};
  const u=findUserByStaffRef(staffRef);
  if(!u){data.staffWorkStatus[staffRef]=value;return}
  const label=getUserDisplayName(u);
  data.staffWorkStatus[u.id]=value;
  if(label!==u.id&&data.staffWorkStatus[label]!=null)delete data.staffWorkStatus[label];
}
function renameStaffReferences(oldId,oldName,newId,newName){
  const oldRefs=[String(oldId||""),String(oldName||"")].filter(Boolean);
  const nextLabel=newName||newId||oldName||oldId||"";
  (data.tasks||[]).forEach(t=>{
    if(oldRefs.includes(String(t.staffUserId||""))||oldRefs.includes(String(t.staff||""))){
      if(newId)t.staffUserId=newId;
      t.staff=nextLabel;
    }
  });
  if(!data.staffWorkStatus)return;
  let migrated;
  oldRefs.forEach(ref=>{
    if(data.staffWorkStatus[ref]!=null&&migrated==null)migrated=data.staffWorkStatus[ref];
    if(ref!==newId)delete data.staffWorkStatus[ref];
  });
  if(migrated!=null&&newId)data.staffWorkStatus[newId]=migrated;
}

/* === CORE FUNCTIONS === */
function getNextRank(t){const mc=getMilestoneCount(t);const nextMile=getNextMilestone(t);const nextMc=getMilestoneCount(nextMile);if(nextMc<=mc)return null;return getRank(nextMile)}
function countTotal(u){let c=0;for(const k of Object.keys(u.stamps||{})){const v=u.stamps[k];if(v==="emergency")c+=3;else if(v)c+=1;}return c+(u.bonusPoints||0)}
function countRange(u,d1,d2){let c=0;for(const k of Object.keys(u.stamps||{})){const d=new Date(k+"T00:00:00");if(between(d,d1,d2)){const v=u.stamps[k];if(v==="emergency")c+=3;else if(v)c+=1;}}return c}
function countRangeDays(u,d1,d2){let c=0;for(const k of Object.keys(u.stamps||{})){const d=new Date(k+"T00:00:00");if(between(d,d1,d2)&&u.stamps[k])c++;}return c}
function countThisMonth(u,b){return countRange(u,startOfMonth(b),endOfMonth(b))}
function countThisWeek(u,b){return countRange(u,startOfWeekMon(b),endOfWeekMon(b))}
function calcStampIncentive(totalPt){const total=Math.max(0,parseInt(totalPt,10)||0);let inc=0;getStampIncentiveRules().forEach(rule=>{inc+=Math.floor(total/rule.every)*rule.amount});return inc}
function getNextStampIncentiveTarget(totalPt){const total=Math.max(0,parseInt(totalPt,10)||0);let nextTarget=null;getStampIncentiveRules().forEach(rule=>{const next=Math.ceil((total+1)/rule.every)*rule.every;if(!nextTarget||next<nextTarget)nextTarget=next});return nextTarget}
function calcMonthInc(u,mk){const total=countTotal(u);const rank=getRank(total);const d=new Date(mk+"-01");return countRange(u,startOfMonth(d),endOfMonth(d))*rank.yen}
function getMonthlyComment(c){for(const m of MONTHLY_COMMENTS)if(c>=m.min&&c<=m.max)return m.msg;return MONTHLY_COMMENTS[4].msg}
function calcReportSalary(r,userId){const hr=userId?getUserHourlyRate(userId):HOURLY_RATE;if(r.workType==="在宅"){const tp=r.taskType||"";const price=getTaskPrice(tp);if(tp==="時給"||tp==="その他（時給）")return calcWorkMinutes(r)/60*hr;if(price!=null)return price*(parseInt(r.manHours)||1);return 0;}return Math.round(calcWorkMinutes(r)/60*hr)}
function calcWorkMinutes(r){const sh=parseInt(r.startH)||0,sm=parseInt(r.startM)||0,eh=parseInt(r.endH)||0,em=parseInt(r.endM)||0;const brk=parseInt(r.breakTime)||0;let d=(eh*60+em)-(sh*60+sm)-brk;return d<0?0:d}
function getUserDateRange(u){let mn=null,mx=null;(u.reports||[]).forEach(r=>{if(!r.date)return;const d=new Date(r.date+"T00:00:00");if(!mn||d<mn)mn=d;if(!mx||d>mx)mx=d;});if(!mn){const n=new Date();mn=n;mx=n;}return{min:mn,max:mx}}
function getAllUsersDateRange(){let mn=null,mx=null;Object.values(data.users).forEach(u=>{(u.reports||[]).forEach(r=>{if(!r.date)return;const d=new Date(r.date+"T00:00:00");if(!mn||d<mn)mn=d;if(!mx||d>mx)mx=d;});});if(!mn){const n=new Date();mn=n;mx=n;}return{min:mn,max:mx}}
function getMonthLockKey(value){
  if(!value)return"";
  const text=String(value);
  if(/^\d{4}-\d{2}$/.test(text))return text;
  if(/^\d{4}-\d{2}-\d{2}$/.test(text))return text.slice(0,7);
  return"";
}
function isLockedMonth(value){
  const key=getMonthLockKey(value);
  return !!(key&&data.lockedMonths&&data.lockedMonths[key]);
}
function buildYearMonthOpts(ySel,mSel,dr,def){const now=new Date();const minY=dr.min.getFullYear();const maxY=Math.max(dr.max.getFullYear(),now.getFullYear())+1;
  ySel.innerHTML="";const a=document.createElement("option");a.value="全て";a.textContent="全て";ySel.appendChild(a);for(let y=minY;y<=maxY;y++){const o=document.createElement("option");o.value=y;o.textContent=y+"年";ySel.appendChild(o);}if(def)ySel.value=String(now.getFullYear());
  mSel.innerHTML="";const b=document.createElement("option");b.value="全て";b.textContent="全て";mSel.appendChild(b);for(let m=1;m<=12;m++){const o=document.createElement("option");o.value=m;o.textContent=m+"月";mSel.appendChild(o);}if(def)mSel.value=String(now.getMonth()+1)}
function filterReports(reps,y,m,wt){reps=reps||[];return reps.filter(r=>{if(!r.date)return false;const d=new Date(r.date+"T00:00:00");if(y!=="全て"&&d.getFullYear()!==parseInt(y))return false;if(m!=="全て"&&(d.getMonth()+1)!==parseInt(m))return false;if(wt!=="全て"&&r.workType!==wt)return false;return true})}

/* === WORKLOAD DISPLAY === */
function renderWorkload(container,staffFilter){
  if(!container)return;
  container.innerHTML="";
  const staffRefs=staffFilter?[staffFilter]:getStaffUsers().map(u=>String(u.id||getUserDisplayName(u)));
  const grid=document.createElement("div");grid.className="workload-grid";
  function applyWlColor(sel){sel.classList.remove("wl-want","wl-ok","wl-busy");
    if(sel.value==="業務が欲しい")sel.classList.add("wl-want");
    else if(sel.value==="まだ余裕あり")sel.classList.add("wl-ok");
    else if(sel.value==="厳しい")sel.classList.add("wl-busy")}
  staffRefs.forEach(staffRef=>{
    const user=findUserByStaffRef(staffRef);
    const label=user?getUserDisplayName(user):String(staffRef);
    const name=label;
    const active=data.tasks.filter(t=>taskMatchesStaffRef(t,staffRef)&&(t.status==="依頼中"||t.status==="期限超過"));
    const irai=active.filter(t=>t.status==="依頼中").length;
    const kigen=active.filter(t=>t.status==="期限超過").length;
    const autoSt=autoWorkloadStatus(staffRef);
    const card=document.createElement("div");card.className="workload-card";
    card.innerHTML=`<div class="wl-name">${escapeHtml(name)}</div><div class="wl-counts">依頼中: <span style="color:var(--blue)">${irai}</span>　期限超過: <span style="color:var(--red)">${kigen}</span></div>`;
    const sel=document.createElement("select");
    const tp=getUserTypeByStaffName(staffRef);
    const opts=(tp==="社会人")?["空いている","まだ余裕あり","厳しい"]:["業務が欲しい","まだ余裕あり","厳しい"];
    opts.forEach(v=>{const o=document.createElement("option");o.value=v;o.textContent=v;sel.appendChild(o)});
    sel.value=autoSt;applyWlColor(sel);
    if(tp==="社会人"){
      sel.disabled=true;
    } else {
      sel.addEventListener("change",()=>{setStaffWorkStatusForRef(staffRef,sel.value);saveData(data);applyWlColor(sel)});
    }
    card.appendChild(sel);grid.appendChild(card);
  });
  container.appendChild(grid);
}

/* Auto-check overdue tasks */
function checkOverdue(){const today=ymd(new Date());let changed=false;data.tasks.forEach(t=>{if(t.status==="依頼中"&&t.deadline&&t.deadline<today){t.status="期限超過";changed=true}});if(changed)saveData(data)}

/* Auto workload status */
function autoWorkloadStatus(staffName){
  const active=data.tasks.filter(t=>taskMatchesStaffRef(t,staffName)&&(t.status==="依頼中"||t.status==="期限超過")).length;
  const tp=getUserTypeByStaffName(staffName);
  if(tp==="社会人"){
    if(active>=3)return"厳しい";
    if(active>=1)return"まだ余裕あり";
    return"空いている";
  }
  if(active>=2)return"厳しい";
  if(active===0)return"まだ余裕あり";
  return getStaffWorkStatusForRef(staffName)||"業務が欲しい";
}

/* Task sequence number */
function nextSeqNum(workType){const existing=data.tasks.filter(t=>t.workType===workType);return existing.length+1}

/* === MODAL === */
function showModal(o){$("mTitle").textContent=o.title||"";$("mSub").textContent=o.sub||"";$("mBody").textContent=o.body||"";$("mBig").textContent=o.big||"🎉";$("mSmall").textContent=o.small||"";$("overlay").style.display="flex"}
function hideModal(){$("overlay").style.display="none";if(modalCb){const cb=modalCb;modalCb=null;cb()}}
var modalCb=null;function showModalCb(o,cb){showModal(o);modalCb=cb}

function showConfetti(){const c=document.createElement("div");c.className="confetti-container";document.body.appendChild(c);const cols=["#ff6b9d","#ff9a56","#ffd93d","#6bcb77","#4d96ff","#9b59b6"];for(let i=0;i<50;i++){const p=document.createElement("div");p.className="confetti-piece";p.style.left=Math.random()*100+"%";p.style.background=cols[~~(Math.random()*cols.length)];p.style.width=(6+Math.random()*8)+"px";p.style.height=(6+Math.random()*8)+"px";p.style.borderRadius=Math.random()>.5?"50%":"2px";p.style.animationDelay=Math.random()*1.5+"s";p.style.animationDuration=2+Math.random()*2+"s";c.appendChild(p)}setTimeout(()=>c.remove(),5000)}

/* === LOTTERY === */
var lotteryCb=null;
function startLottery(cb){lotteryCb=cb;var lo=$("lotteryOverlay");lo.style.display="flex";$("lotteryResult").textContent="";$("lotteryClose").classList.add("hidden");$("lotteryCards").innerHTML="";
const roll=Math.random();let prize=roll<.01?5:roll<.11?2:1;const vals=[5,2,1,1,1];for(let i=vals.length-1;i>0;i--){const j=~~(Math.random()*(i+1));[vals[i],vals[j]]=[vals[j],vals[i]]}
let chosen=false;lo.dataset.prize=prize;
vals.forEach(dv=>{const card=document.createElement("div");card.className="lottery-card";const vd=document.createElement("div");vd.className="card-val";vd.innerHTML=`<span class="pt-num">${dv}</span><span>pt</span>`;card.appendChild(vd);
card.addEventListener("click",()=>{if(chosen)return;chosen=true;vd.innerHTML=`<span class="pt-num">${prize}</span><span>pt</span>`;card.classList.add("selected","revealed");
$("lotteryCards").querySelectorAll(".lottery-card").forEach(c=>{if(c!==card)c.classList.add("disabled")});$("lotteryResult").textContent=`🎉 ${prize}pt ゲット！`;
setTimeout(()=>{$("lotteryCards").querySelectorAll(".lottery-card").forEach(c=>{if(c!==card){c.classList.remove("disabled");c.classList.add("revealed")}});$("lotteryClose").classList.remove("hidden")},1200)});
$("lotteryCards").appendChild(card)})}

/* === COMMON NAV HELPERS === */
function doLogout(){data.session.userId="";clearToken();saveData(data);location.hash="#user-login"}
function doAdminLogout(){data.session.adminAuthed=false;clearToken();data.session.adminEditingUserId="";data.session.adminReportEditingUserId="";saveData(data);location.hash="#admin-login"}

function renderRankBadge(el,total){const r=getRank(total);el.innerHTML=`<span class="rank-badge rank-${r.rank}">${r.emoji} ${r.label}</span>`}
function renderProgress(el,total){
  const nextMile=getNextMilestone(total);
  const prevMile=total<200?nextMile-25:nextMile-50;
  const range=nextMile-prevMile;
  const pct=range===0?100:Math.min(100,((total-prevMile)/range)*100);
  const currentInc=calcStampIncentive(total);
  const nextInc=calcStampIncentive(nextMile);
  const bonus=nextInc-currentInc;
  el.innerHTML=`<div class="progress-wrap"><div class="progress-label"><span>次のｲﾝｾﾝﾃｨﾌﾞ ${nextMile}pt</span><span>あと <b>${nextMile-total}pt</b>（+${bonus.toLocaleString()}円）</span></div><div class="progress-bar"><div class="progress-fill" style="width:${pct}%"></div></div></div>`}
