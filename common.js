/* ================================================================ */
/* common.js - 業務管理アプリ 共通コード                              */
/* ================================================================ */

const LS_KEY="stampcard_v7_clean";

/* ================================================================ */
/* === GOOGLE DRIVE API SYNC LAYER === */
/* ================================================================ */
const API_URL_KEY = "stampcard_api_url";
const DEFAULT_API_URL = window.APP_CONFIG?.DEFAULT_API_URL || "";
let API_URL = DEFAULT_API_URL || localStorage.getItem(API_URL_KEY) || ""; 
localStorage.setItem(API_URL_KEY, API_URL);

const TOKEN_KEY = "stampcard_api_token";
function getToken(){ return sessionStorage.getItem(TOKEN_KEY) || ""; }
function setToken(t){ sessionStorage.setItem(TOKEN_KEY, t); }
function clearToken(){ sessionStorage.removeItem(TOKEN_KEY); }

async function downloadDriveFile(fileId, fileName){
  if(!API_URL) { showModal({title:"API未接続です（⚙で設定）",big:"⚠️"}); return; }
  const t=getToken();
  if(!t){ showModal({title:"未ログインです",big:"⚠️"}); return; }
  const url = API_URL + "?action=download&token=" + encodeURIComponent(t) + "&fileId=" + encodeURIComponent(fileId);
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

async function fetchTodayPasswordForAdmin(){
  if(!API_URL) return null;
  const t = getToken();
  if(!t) return null;
  try{
    const resp = await fetch(API_URL + "?action=todayPassword&token=" + encodeURIComponent(t), { redirect:"follow" });
    const r = await resp.json();
    if(!r || !r.ok) return null;
    return r.password || null;
  }catch(e){
    return null;
  }
}

let _syncVersion = 0;
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
  ov.style.display = "block";
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
let lastSyncedDataStr = localStorage.getItem(LS_KEY) || "{}";
let actionQueue = [];
let isSending = false;

function queueAction(action, payload) {
  actionQueue.push({ action, payload });
  processQueue();
}

async function processQueue() {
  if (isSending || actionQueue.length === 0 || !API_URL) return;
  isSending = true;
  const req = actionQueue.shift();
  updateSyncUI("loading", "保存中...");
  try {
    req.payload._action = req.action;
    req.payload.token = getToken();
    const resp = await fetch(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain" },
      body: JSON.stringify(req.payload),
      redirect: "follow"
    });
    const result = await resp.json();
    if (result.ok) {
      _syncVersion = result.version || _syncVersion + 1;
      _lastSyncTime = Date.now();
      updateSyncUI("ok", "同期済み ✓");
    } else {
      updateSyncUI("err", result.error || "保存エラー");
    }
  } catch (e) {
    updateSyncUI("err", "通信エラー");
  }
  isSending = false;
  processQueue();
}

async function syncPush() { return true; }

async function syncPull() {
  if (!API_URL || _isSyncing) return false;
  _isSyncing = true;
  try {
    const resp = await fetch(API_URL + "?action=read&token=" + encodeURIComponent(getToken()), { redirect: "follow" });
    const result = await resp.json();
    if (result.ok && result.data) {
      const remoteVer = result.data._version || 0;
      if (remoteVer > _syncVersion) {
        _syncVersion = remoteVer;
        const clean = { ...result.data };
        delete clean._version;
        delete clean._updatedAt;
        localStorage.setItem(LS_KEY, JSON.stringify(clean));
        data = JSON.parse(JSON.stringify(clean));
        lastSyncedDataStr = JSON.stringify(clean);
        migrateData();
        _lastSyncTime = Date.now();
        updateSyncUI("ok", "同期済み ✓");
        try { route(); } catch(e) {}
        return true;
      }
      _lastSyncTime = Date.now();
      updateSyncUI("ok", "最新 ✓");
      return true;
    }
    updateSyncUI("err", "読込エラー");
    return false;
  } catch (e) {
    updateSyncUI("err", "通信エラー");
    return false;
  } finally {
    _isSyncing = false;
  }
}

function saveData(d) {
  localStorage.setItem(LS_KEY, JSON.stringify(d));
  const oldD = JSON.parse(lastSyncedDataStr);
  const newD = d;

  const oldTasks = oldD.tasks || [];
  const newTasks = newD.tasks || [];
  newTasks.forEach(nT => {
    const oT = oldTasks.find(t => t.id === nT.id);
    if (!oT || JSON.stringify(oT) !== JSON.stringify(nT)) {
      queueAction("updateTask", { task: nT, isDelete: false });
    }
  });
  oldTasks.forEach(oT => {
    if (!newTasks.find(t => t.id === oT.id)) {
      queueAction("updateTask", { task: { id: oT.id }, isDelete: true });
    }
  });

  const oldUsers = oldD.users || {};
  const newUsers = newD.users || {};
  Object.keys(newUsers).forEach(uid => {
    if (JSON.stringify(oldUsers[uid]) !== JSON.stringify(newUsers[uid])) {
      queueAction("updateUserFull", { targetUserId: uid, userObj: newUsers[uid] });
    }
  });

  let masterChanged = false;
  let deletedUid = null;
  if (JSON.stringify(oldD.taskTypes) !== JSON.stringify(newD.taskTypes)) masterChanged = true;
  if (JSON.stringify(oldD.taskPrices) !== JSON.stringify(newD.taskPrices)) masterChanged = true;
  if (JSON.stringify(oldD.employees) !== JSON.stringify(newD.employees)) masterChanged = true;
  if (JSON.stringify(oldD.userHourlyRates) !== JSON.stringify(newD.userHourlyRates)) masterChanged = true;
  if (JSON.stringify(oldD.staffWorkStatus) !== JSON.stringify(newD.staffWorkStatus)) masterChanged = true;
  if (JSON.stringify(oldD.lockedMonths) !== JSON.stringify(newD.lockedMonths)) masterChanged = true;
  Object.keys(oldUsers).forEach(uid => {
    if (!newUsers[uid]) { masterChanged = true; deletedUid = uid; }
  });

  if (masterChanged) {
    queueAction("updateMaster", {
      taskTypes: newD.taskTypes,
      taskPrices: newD.taskPrices,
      employees: newD.employees,
      userHourlyRates: newD.userHourlyRates,
      staffWorkStatus: newD.staffWorkStatus,
      lockedMonths: newD.lockedMonths,
      deleteUserId: deletedUid
    });
  }

  lastSyncedDataStr = JSON.stringify(d);
}
// ==========================================
// ▲▲▲ スマート同期システムここまで ▲▲▲
// ==========================================

async function syncCheckVersion() {
  if (!API_URL || _isSyncing) return;
  try {
    const resp = await fetch(API_URL + "?action=version&token=" + encodeURIComponent(getToken()), { redirect: "follow" });
    const result = await resp.json();
    if (result.ok && (result.version || 0) > _syncVersion) {
      await syncPull();
    }
  } catch (e) {
  }
}

function startSyncPolling() {
  stopSyncPolling();
  _syncTimer = setInterval(syncCheckVersion, POLL_INTERVAL);
}
function stopSyncPolling() {
  if (_syncTimer) { clearInterval(_syncTimer); _syncTimer = null; }
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
  if(!data.session)data.session={userId:"",adminAuthed:false,adminEditingUserId:"",adminReportEditingUserId:""};
  if(!data.session.adminReportEditingUserId)data.session.adminReportEditingUserId="";
  data.tasks.forEach(t=>{if(!t.fileNames){t.fileNames=t.fileName?[t.fileName]:[];if(t.fileName)delete t.fileName}});
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
async function uploadFileToDrive(file) {
  if (!API_URL) return null;
  return new Promise(function(resolve) {
    var reader = new FileReader();
    reader.onload = async function() {
      try {
        var b64 = reader.result.split(",")[1];
        var resp = await fetch(API_URL, {
          method: "POST",
          headers: { "Content-Type": "text/plain" },
          body: JSON.stringify({ _action: "uploadFile", token: getToken(), fileName: file.name, mimeType: file.type || "application/octet-stream", data: b64 }),
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
  const users={
    "morotomi"  :mkU("morotomi","諸富","社会人"),
    "osawa"     :mkU("osawa","大澤","社会人"),
    "ogasawara" :mkU("ogasawara","小笠原","社会人"),
    "shirakawa" :mkU("shirakawa","白川","学生"),
    "matsumoto" :mkU("matsumoto","松本","学生"),
    "kamiko"    :mkU("kamiko","神子","学生"),
    "tanaka_s"  :mkU("tanaka_s","田中（学生）","学生"),
    "suzuki"    :mkU("suzuki","鈴木","社会人"),
    "yamada"    :mkU("yamada","山田","学生"),
    "ito"       :mkU("ito","伊藤","学生")
  };
  const tasks=[];
  return{users,session:{userId:"",adminAuthed:false,adminEditingUserId:"",adminReportEditingUserId:""},
    tasks,employees:[...DEFAULT_EMPLOYEES],taskTypes:[...DEFAULT_TASK_TYPES],taskPrices:{...TASK_PRICES},staffWorkStatus:{}};
}
var data=loadData();
migrateData();

function getUserHourlyRate(userId){return data.userHourlyRates&&data.userHourlyRates[userId]!=null?data.userHourlyRates[userId]:HOURLY_RATE}
// ⑩修正: 初回はlocalStorageのみ保存（APIへの差分送信はスキップ）
localStorage.setItem(LS_KEY, JSON.stringify(data));
lastSyncedDataStr = JSON.stringify(data);

function getTaskPrice(name){return data.taskPrices&&data.taskPrices[name]!=null?data.taskPrices[name]:(TASK_PRICES[name]!=null?TASK_PRICES[name]:null)}
function getTaskTypes(){return data.taskTypes||DEFAULT_TASK_TYPES}
function getEmployees(){return data.employees||DEFAULT_EMPLOYEES}
function getStaffNames(){return Object.values(data.users).map(u=>u.name||u.id)}
function getUserTypeByStaffName(staffName){
  const u=Object.values(data.users).find(x=>(x.name||x.id)===staffName);
  return u?u.userType:"";
}

/* === CORE FUNCTIONS === */
function getNextRank(t){const mc=getMilestoneCount(t);const nextMile=getNextMilestone(t);const nextMc=getMilestoneCount(nextMile);if(nextMc<=mc)return null;return getRank(nextMile)}
function countTotal(u){let c=0;for(const k of Object.keys(u.stamps||{})){const v=u.stamps[k];if(v==="emergency")c+=3;else if(v)c+=1;}return c+(u.bonusPoints||0)}
function countRange(u,d1,d2){let c=0;for(const k of Object.keys(u.stamps||{})){const d=new Date(k+"T00:00:00");if(between(d,d1,d2)){const v=u.stamps[k];if(v==="emergency")c+=3;else if(v)c+=1;}}return c}
function countRangeDays(u,d1,d2){let c=0;for(const k of Object.keys(u.stamps||{})){const d=new Date(k+"T00:00:00");if(between(d,d1,d2)&&u.stamps[k])c++;}return c}
function countThisMonth(u,b){return countRange(u,startOfMonth(b),endOfMonth(b))}
function countThisWeek(u,b){return countRange(u,startOfWeekMon(b),endOfWeekMon(b))}
function calcStampIncentive(totalPt){let inc=0;inc+=Math.floor(totalPt/25)*5000;if(totalPt>=50){const s=Math.floor(totalPt/50);for(let i=1;i<=s;i++)inc+=(i*50)*100;}if(totalPt>=250){inc+=Math.floor((totalPt-200)/50)*40000;}return inc}
function calcMonthInc(u,mk){const total=countTotal(u);const rank=getRank(total);const d=new Date(mk+"-01");return countRange(u,startOfMonth(d),endOfMonth(d))*rank.yen}
function getMonthlyComment(c){for(const m of MONTHLY_COMMENTS)if(c>=m.min&&c<=m.max)return m.msg;return MONTHLY_COMMENTS[4].msg}
function calcReportSalary(r,userId){const hr=userId?getUserHourlyRate(userId):HOURLY_RATE;if(r.workType==="在宅"){const tp=r.taskType||"";const price=getTaskPrice(tp);if(tp==="時給"||tp==="その他（時給）")return calcWorkMinutes(r)/60*hr;if(price!=null)return price*(parseInt(r.manHours)||1);return 0;}return Math.round(calcWorkMinutes(r)/60*hr)}
function calcWorkMinutes(r){const sh=parseInt(r.startH)||0,sm=parseInt(r.startM)||0,eh=parseInt(r.endH)||0,em=parseInt(r.endM)||0;const brk=parseInt(r.breakTime)||0;let d=(eh*60+em)-(sh*60+sm)-brk;return d<0?0:d}
function getUserDateRange(u){let mn=null,mx=null;(u.reports||[]).forEach(r=>{if(!r.date)return;const d=new Date(r.date+"T00:00:00");if(!mn||d<mn)mn=d;if(!mx||d>mx)mx=d;});if(!mn){const n=new Date();mn=n;mx=n;}return{min:mn,max:mx}}
function getAllUsersDateRange(){let mn=null,mx=null;Object.values(data.users).forEach(u=>{(u.reports||[]).forEach(r=>{if(!r.date)return;const d=new Date(r.date+"T00:00:00");if(!mn||d<mn)mn=d;if(!mx||d>mx)mx=d;});});if(!mn){const n=new Date();mn=n;mx=n;}return{min:mn,max:mx}}
function buildYearMonthOpts(ySel,mSel,dr,def){const now=new Date();const minY=dr.min.getFullYear();const maxY=Math.max(dr.max.getFullYear(),now.getFullYear())+1;
  ySel.innerHTML="";const a=document.createElement("option");a.value="全て";a.textContent="全て";ySel.appendChild(a);for(let y=minY;y<=maxY;y++){const o=document.createElement("option");o.value=y;o.textContent=y+"年";ySel.appendChild(o);}if(def)ySel.value=String(now.getFullYear());
  mSel.innerHTML="";const b=document.createElement("option");b.value="全て";b.textContent="全て";mSel.appendChild(b);for(let m=1;m<=12;m++){const o=document.createElement("option");o.value=m;o.textContent=m+"月";mSel.appendChild(o);}if(def)mSel.value=String(now.getMonth()+1)}
function filterReports(reps,y,m,wt){reps=reps||[];return reps.filter(r=>{if(!r.date)return false;const d=new Date(r.date+"T00:00:00");if(y!=="全て"&&d.getFullYear()!==parseInt(y))return false;if(m!=="全て"&&(d.getMonth()+1)!==parseInt(m))return false;if(wt!=="全て"&&r.workType!==wt)return false;return true})}

/* Auto-check overdue tasks */
function checkOverdue(){const today=ymd(new Date());data.tasks.forEach(t=>{if(t.status==="依頼中"&&t.deadline&&t.deadline<today)t.status="期限超過";});saveData(data)}

/* Auto workload status */
function autoWorkloadStatus(staffName){
  const active=data.tasks.filter(t=>(t.staff===staffName)&&(t.status==="依頼中"||t.status==="期限超過")).length;
  const tp=getUserTypeByStaffName(staffName);
  if(tp==="社会人"){
    if(active>=3)return"厳しい";
    if(active>=1)return"まだ余裕あり";
    return"空いている";
  }
  if(active>=2)return"厳しい";
  if(active===0)return"まだ余裕あり";
  return data.staffWorkStatus[staffName]||"業務が欲しい";
}

/* Task sequence number */
function nextSeqNum(workType){const existing=data.tasks.filter(t=>t.workType===workType);return existing.length+1}

/* === MODAL === */
function showModal(o){$("mTitle").textContent=o.title||"";$("mSub").textContent=o.sub||"";$("mBody").textContent=o.body||"";$("mBig").textContent=o.big||"🎉";$("mSmall").textContent=o.small||"";$("overlay").style.display="block"}
function hideModal(){$("overlay").style.display="none";if(modalCb){const cb=modalCb;modalCb=null;cb()}}
var modalCb=null;function showModalCb(o,cb){showModal(o);modalCb=cb}

function showConfetti(){const c=document.createElement("div");c.className="confetti-container";document.body.appendChild(c);const cols=["#ff6b9d","#ff9a56","#ffd93d","#6bcb77","#4d96ff","#9b59b6"];for(let i=0;i<50;i++){const p=document.createElement("div");p.className="confetti-piece";p.style.left=Math.random()*100+"%";p.style.background=cols[~~(Math.random()*cols.length)];p.style.width=(6+Math.random()*8)+"px";p.style.height=(6+Math.random()*8)+"px";p.style.borderRadius=Math.random()>.5?"50%":"2px";p.style.animationDelay=Math.random()*1.5+"s";p.style.animationDuration=2+Math.random()*2+"s";c.appendChild(p)}setTimeout(()=>c.remove(),5000)}

/* === LOTTERY === */
var lotteryCb=null;
function startLottery(cb){lotteryCb=cb;var lo=$("lotteryOverlay");lo.style.display="block";$("lotteryResult").textContent="";$("lotteryClose").classList.add("hidden");$("lotteryCards").innerHTML="";
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
