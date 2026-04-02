/* === VIEWS (admin-specific) === */
const views={userAuth:_noop,userStamp:_noop,userHome:_noop,reportInput:_noop,reportConfirm:_noop,staffTaskList:_noop,adminAuth:$("adminAuth"),adminReportMgmt:$("adminReportMgmt"),adminReportDetail:$("adminReportDetail"),adminTaskList:$("adminTaskList"),adminDropdownEdit:$("adminDropdownEdit"),adminHome:$("adminHome"),adminEdit:$("adminEdit"),adminMonthCheck:$("adminMonthCheck")};

const DEFAULT_BOOTSTRAP_STAFF = [
  { id: "shakai_test", pw: "shakai_test", name: "テスト社会人", userType: "社会人" },
  { id: "ogasawara", pw: "ogasawara", name: "小笠原", userType: "社会人" },
  { id: "morotomi", pw: "morotomi", name: "諸富", userType: "社会人" },
  { id: "osawa", pw: "osawa", name: "大澤", userType: "社会人" },
  { id: "yoneoka", pw: "yoneoka", name: "米岡", userType: "社会人" },
  { id: "hosaka", pw: "hosaka", name: "保坂", userType: "社会人" },
  { id: "gakusei_test", pw: "gakusei_test", name: "テスト学生", userType: "学生" },
  { id: "miko", pw: "miko", name: "神子", userType: "学生" },
  { id: "shirakawa", pw: "shirakawa", name: "白川", userType: "学生" },
  { id: "matsumoto", pw: "matsumoto", name: "松本", userType: "学生" },
  { id: "mizutani", pw: "mizutani", name: "水谷", userType: "学生" },
  { id: "takeuchi", pw: "takeuchi", name: "竹内", userType: "学生" },
  { id: "fujikawa", pw: "fujikawa", name: "藤川", userType: "学生" },
  { id: "kobayashi", pw: "kobayashi", name: "小林", userType: "学生" }
];
let _defaultStaffBootstrapPromise = null;

/* === MODAL/ESCAPE SETUP === */
$("mClose").addEventListener("click",hideModal);$("overlay").addEventListener("click",e=>{if(e.target===$("overlay"))hideModal()});
document.addEventListener("keydown",e=>{if(e.key==="Escape"){hideModal();var fo=document.getElementById("fileOverlay");if(fo)fo.style.display="none";var ta=document.getElementById("taskAddOverlay");if(ta)ta.style.display="none";var dd=document.getElementById("ddEditOverlay");if(dd)dd.style.display="none";var _ao=document.getElementById("apiSetupOverlay");if(_ao)_ao.style.display="none";var se=document.getElementById("staffEditOverlay");if(se)se.style.display="none"}});
$("lotteryClose").addEventListener("click",()=>{$("lotteryOverlay").style.display="none";if(lotteryCb){const cb=lotteryCb;lotteryCb=null;cb(parseInt($("lotteryOverlay").dataset.prize)||1)}});

/* === ROUTER === */
function showOnly(v){Object.values(views).forEach(x=>x.classList.add("hidden"));views[v].classList.remove("hidden")}
function route(){checkOverdue();const h=location.hash||"#admin-login";
if(h==="#user-login"||h==="#user-stamp"||h==="#user"||h==="#report-input"||h==="#report-confirm"||h==="#staff-task-list"){window.location.href="staff.html"+h;return}
if(h==="#admin-login"){showOnly("adminAuth");$("adminAuthErr").style.display="none";$("adminLoginPw").value="";return}
if(h==="#admin-report-mgmt"){if(!data.session.adminAuthed){location.hash="#admin-login";return}showOnly("adminReportMgmt");renderAdminReportMgmt();return}
if(h==="#admin-report-detail"){if(!data.session.adminAuthed||!data.session.adminReportEditingUserId){location.hash="#admin-report-mgmt";return}showOnly("adminReportDetail");renderAdminReportDetail();return}
if(h==="#admin-task-list"){if(!data.session.adminAuthed){location.hash="#admin-login";return}showOnly("adminTaskList");renderAdminTaskList();return}
if(h==="#admin-dropdown-edit"){if(!data.session.adminAuthed){location.hash="#admin-login";return}showOnly("adminDropdownEdit");renderDropdownEdit();syncPull().then(changed=>{if(changed && location.hash==="#admin-dropdown-edit" && data.session.adminAuthed) renderDropdownEdit()});return}
if(h==="#admin-month-check"){if(!data.session.adminAuthed){location.hash="#admin-login";return}showOnly("adminMonthCheck");renderMonthCheck();return}
if(h==="#admin"){if(!data.session.adminAuthed){location.hash="#admin-login";return}showOnly("adminHome");renderAdminHome();syncPull().then(changed=>{if(changed && location.hash==="#admin" && data.session.adminAuthed) renderAdminHome()});return}
if(h==="#admin-edit"){if(!data.session.adminAuthed||!data.session.adminEditingUserId){location.hash="#admin";return}showOnly("adminEdit");renderAdminEdit();syncPull().then(changed=>{if(changed && location.hash==="#admin-edit" && data.session.adminAuthed && data.session.adminEditingUserId) renderAdminEdit()});return}
location.hash="#admin-login"}
window.addEventListener("hashchange",route);

/* === NAV === */
function doLogout(){data.session.userId="";clearToken();saveLocalOnly(data);location.hash="#user-login"}
function doAdminLogout(){data.session.adminAuthed=false;clearToken();data.session.adminEditingUserId="";data.session.adminReportEditingUserId="";saveLocalOnly(data);location.hash="#admin-login"}
$("adminLogout").addEventListener("click",doAdminLogout);$("armLogout").addEventListener("click",doAdminLogout);$("atlLogout").addEventListener("click",doAdminLogout);$("ddeLogout").addEventListener("click",doAdminLogout);

async function bootstrapDefaultStaffIfNeeded(){
  if (_defaultStaffBootstrapPromise) return _defaultStaffBootstrapPromise;
  if (!data.session.adminAuthed || !API_URL || !getToken()) return false;
  const missingStaff = DEFAULT_BOOTSTRAP_STAFF.filter(staff => !((data.users || {})[staff.id]));
  if (!missingStaff.length) return false;

  _defaultStaffBootstrapPromise = (async () => {
    const created = [];
    const failed = [];
    try {
      for (const staff of missingStaff) {
        try {
          const resp = await fetch(API_URL, {
            method: "POST",
            headers: {"Content-Type":"text/plain"},
            body: JSON.stringify({
              _action: "upsertStaffUser",
              token: getToken(),
              id: staff.id,
              pw: staff.pw,
              name: staff.name,
              userType: staff.userType
            }),
            redirect: "follow"
          });
          const result = await resp.json();
          applySyncMeta(syncMetaFromResult(result));
          if (!result.ok) throw new Error(result.error || "bootstrap failed");
          created.push(staff.id);
        } catch (error) {
          failed.push(`${staff.id}: ${error && error.message ? error.message : "error"}`);
        }
      }

      data.users = data.users || {};
      data.userHourlyRates = data.userHourlyRates || {};
      DEFAULT_BOOTSTRAP_STAFF.forEach((staff, index) => {
        data.users[staff.id] = Object.assign({}, data.users[staff.id] || {}, {
          id: staff.id,
          name: staff.name,
          userType: staff.userType,
          createdAt: (data.users[staff.id] && data.users[staff.id].createdAt) || (Date.now() + index)
        });
        data.userHourlyRates[staff.id] = 1300;
      });
      saveData(data);
      try { await syncPull(); } catch (_error) {}
      if (failed.length) {
        showModal({title:"一部スタッフ未作成",sub:`${failed.length}件失敗しました`,big:"⚠️"});
      }
      return created.length > 0;
    } catch (error) {
      if (created.length) {
        try { await syncPull(); } catch (_error) {}
      }
      showModal({title:"初期スタッフ作成に失敗",sub:error && error.message ? error.message : "error",big:"⚠️"});
      return false;
    } finally {
      _defaultStaffBootstrapPromise = null;
    }
  })();

  return _defaultStaffBootstrapPromise;
}

// Login
$("btnUserLogin").addEventListener("click", async ()=>{const id=$("userLoginId").value.trim(),pw=$("userLoginPw").value;$("userAuthErr").style.display="none";
if(!API_URL){$("userAuthErr").textContent="API未接続です（⚙で設定）";$("userAuthErr").style.display="block";return}
try{
  const resp=await fetch(API_URL,{method:"POST",headers:{"Content-Type":"text/plain"},body:JSON.stringify({_action:"loginStaff",id,pw}),redirect:"follow"});
  const result=await resp.json();
  if(!result.ok){$("userAuthErr").textContent="ログインに失敗しました。";$("userAuthErr").style.display="block";return}
  setToken(result.token);
  data.session.userId=result.user.id;
  data.users=data.users||{};
  data.users[result.user.id]=Object.assign({},data.users[result.user.id]||{},{name:result.user.name,userType:result.user.userType});
  saveLocalOnly(data);
  userMonthCursor=startOfMonth(new Date());
  syncPull().then(changed=>{
    if(!changed || data.session.userId!==result.user.id) return;
    if(location.hash==="#user-stamp") renderStampScreen();
    else if(location.hash==="#report-confirm") renderReportConfirm();
  });
  if(result.user.userType==="社会人"){location.hash="#report-confirm";}else{
    const today=ymd(new Date());
    const visited=data.users[result.user.id].stampScreenVisitedToday===today;
    if(!visited){data.users[result.user.id].stampScreenVisitedToday=today;saveLocalOnly(data)}
    location.hash="#user-stamp";
  }
}catch(e){$("userAuthErr").textContent="通信エラー";$("userAuthErr").style.display="block";}
});
$("userLoginPw").addEventListener("keydown",e=>{if(e.key==="Enter")$("btnUserLogin").click()});
$("btnAdminLogin").addEventListener("click", async ()=>{const id=$("adminLoginId").value.trim(),pw=$("adminLoginPw").value;$("adminAuthErr").style.display="none";
if(!API_URL){$("adminAuthErr").textContent="API未接続です（⚙で設定）";$("adminAuthErr").style.display="block";return}
try{
  const resp=await fetch(API_URL,{method:"POST",headers:{"Content-Type":"text/plain"},body:JSON.stringify({_action:"loginAdmin",id,pw}),redirect:"follow"});
  const result=await resp.json();
  if(!result.ok){$("adminAuthErr").textContent="ログインに失敗しました。";$("adminAuthErr").style.display="block";return}
  setToken(result.token);
  data.session.adminAuthed=true;
  data.session.adminEditingUserId="";
  saveLocalOnly(data);
  syncPull().then(changed=>{
    if(changed && location.hash==="#admin-task-list" && data.session.adminAuthed) renderAdminTaskList();
  });
  location.hash="#admin-task-list";
}catch(e){$("adminAuthErr").textContent="通信エラー";$("adminAuthErr").style.display="block";}
});
$("adminLoginPw").addEventListener("keydown",e=>{if(e.key==="Enter")$("btnAdminLogin").click()});

// Nav buttons
let rpTransportLocked=false;
$("goToCalendarFromStamp").addEventListener("click",()=>{userMonthCursor=startOfMonth(new Date());location.hash="#user"});
$("goToStamp").addEventListener("click",()=>location.hash="#user-stamp");
$("stampGoReport").addEventListener("click",()=>{editingReportIdx=-1;rpTransportLocked=false;location.hash="#report-input"});
$("stampGoReportList").addEventListener("click",()=>location.hash="#report-confirm");
$("stampGoTaskList").addEventListener("click",()=>location.hash="#staff-task-list");
$("btnGoReport").addEventListener("click",()=>{editingReportIdx=-1;rpTransportLocked=false;location.hash="#report-input"});
$("btnGoReportList").addEventListener("click",()=>location.hash="#report-confirm");
$("btnGoTaskListFromCal").addEventListener("click",()=>location.hash="#staff-task-list");
$("reportBackToCal").addEventListener("click",()=>{
  if(adminEditingReportMode){adminEditingReportMode=false;data.session.userId=adminEditOrigUserId;saveLocalOnly(data);editingReportIdx=-1;location.hash="#admin-report-detail";return}
  const u=data.users[data.session.userId];window.location.href="staff.html"+(u&&u.userType==="社会人"?"#report-confirm":"#user")});
$("confirmBackToCal").addEventListener("click",()=>{const u=data.users[data.session.userId];if(u&&u.userType==="社会人")return;location.hash="#user"});
$("confirmGoTask").addEventListener("click",()=>location.hash="#staff-task-list");
$("btnNewReport").addEventListener("click",()=>{editingReportIdx=-1;rpTransportLocked=false;location.hash="#report-input"});
$("stlBack").addEventListener("click",()=>{const u=data.users[data.session.userId];
if(u&&u.userType==="社会人"){location.hash="#report-confirm";}else{location.hash="#user";}});
$("stlBackToReport").addEventListener("click",()=>location.hash="#report-confirm");

// Admin tabs
const tabNav=(rm,sm,tl,dd,mc)=>{$(rm).addEventListener("click",()=>location.hash="#admin-report-mgmt");$(sm).addEventListener("click",()=>location.hash="#admin");$(tl).addEventListener("click",()=>location.hash="#admin-task-list");$(dd).addEventListener("click",()=>location.hash="#admin-dropdown-edit");if(mc)$(mc).addEventListener("click",()=>location.hash="#admin-month-check")};
tabNav("tabRM","tabSM","tabTL","tabDD","tabMC");tabNav("tabRM2","tabSM2","tabTL2","tabDD2","tabMC2");tabNav("tabRM3","tabSM3","tabTL3","tabDD3","tabMC3");tabNav("tabRM4","tabSM4","tabTL4","tabDD4","tabMC4");

/* === STAMP SCREEN === */
function renderStampScreen(){const u=data.users[data.session.userId];if(!u){location.hash="#user-login";return}
const now=new Date(),total=countTotal(u),stamped=!!u.stamps[ymd(now)];
$("stampUserName").textContent=u.name||u.id;$("stampDate").textContent=`${now.getMonth()+1}月${now.getDate()}日（${dowJa(now)}）`;
renderRankBadge($("stampRankBadge"),total);renderProgress($("stampRankInfo"),total);
const btn=$("bigStampBtn");const failed=u.stampFailed===ymd(now);
if(stamped){btn.classList.add("done");btn.querySelector(".emoji").textContent="✅";btn.querySelector(".label").textContent="スタンプ済み";btn.disabled=true;$("stampAlreadyMsg").classList.remove("hidden")}
else if(failed){btn.classList.add("done");btn.querySelector(".emoji").textContent="❌";btn.querySelector(".label").textContent="本日不可";btn.disabled=true;$("stampAlreadyMsg").classList.add("hidden")}
else{btn.classList.remove("done");btn.querySelector(".emoji").textContent="👆";btn.querySelector(".label").textContent="出勤スタンプ";btn.disabled=false;$("stampAlreadyMsg").classList.add("hidden")}}

$("bigStampBtn").addEventListener("click",async ()=>{const u=data.users[data.session.userId];if(!u)return;const now=new Date(),key=ymd(now);if(u.stamps[key])return;
if(u.stampFailed===key){showModal({title:"本日のスタンプ不可",sub:"合言葉を間違えたため、今日はスタンプを押せません。",big:"🚫"});return}
const ans=prompt("合言葉は？");if(ans===null)return;
try{const resp=await fetch(API_URL,{method:"POST",headers:{"Content-Type":"text/plain"},body:JSON.stringify({_action:"verifyDailyPassword",token:getToken(),answer:ans}),redirect:"follow"});const r=await resp.json();if(!(r.ok&&r.match)){u.stampFailed=key;saveData(data);const btn=$("bigStampBtn");btn.classList.add("done");btn.querySelector(".emoji").textContent="❌";btn.querySelector(".label").textContent="本日不可";btn.disabled=true;showModal({title:"合言葉が違います",sub:"今日はスタンプを押せません。",big:"❌"});return}}catch(e){showModal({title:"通信エラー",sub:"合言葉確認に失敗しました。",big:"📡"});return}

u.stamps[key]=true;saveData(data);const mk=ym(now);let monthFirstMsg=null;
if(u.lastMonthFirstStamp!==mk){u.lastMonthFirstStamp=mk;const ps=addMonths(startOfMonth(now),-1),pe=endOfMonth(ps),pc=countRangeDays(u,ps,pe);
monthFirstMsg={title:"先月の振り返り",sub:`先月の出勤回数：${pc}回`,body:getMonthlyComment(pc),big:"👏",small:`${monthLabelJa(ps)}おつかれさまでした！`};saveData(data)}
const lotteryTrigger=Math.random()<.33;
const afterStamp=bp=>{if(bp&&bp>1){u.bonusPoints=(u.bonusPoints||0)+(bp-1);saveData(data)}
const total=countTotal(u);const m50=Math.floor(total/50);
if(m50>0&&m50>(u.lastCongrats50||0)){u.lastCongrats50=m50;saveData(data);showConfetti();showModalCb({title:"🎊 おめでとう！",sub:`累計 ${m50*50}pt`,body:"お礼の1万円！",big:"💰",small:"これからもよろしく！"},()=>{userMonthCursor=startOfMonth(new Date());location.hash="#user"});return}
if(monthFirstMsg){showModalCb(monthFirstMsg,()=>{userMonthCursor=startOfMonth(new Date());location.hash="#user"});return}
showConfetti();showModalCb({title:"スタンプ！",sub:"出勤おつかれ！",body:`累計 ${total}pt`,big:"🎉",small:"ナイス出勤！"},()=>{userMonthCursor=startOfMonth(new Date());location.hash="#user"})};
if(lotteryTrigger){startLottery(p=>afterStamp(p))}else{afterStamp(0)}});

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

/* === USER CALENDAR === */
let userMonthCursor=startOfMonth(new Date());
$("uPrev").addEventListener("click",()=>{userMonthCursor=addMonths(userMonthCursor,-1);renderUserHome()});
$("uNext").addEventListener("click",()=>{userMonthCursor=addMonths(userMonthCursor,+1);renderUserHome()});
$("uThis").addEventListener("click",()=>{userMonthCursor=startOfMonth(new Date());renderUserHome()});
let stampEditMode=false;
let stampEditStamps={};
let stampEditEmergencyMode=false;
function renderUserHome(){const u=data.users[data.session.userId];if(!u){location.hash="#user-login";return}
$("userNameLabel").textContent=u.name||u.id;const now=new Date(),total=countTotal(u),rank=getRank(total);
$("uTotal").textContent=total;$("uMonth").textContent=countThisMonth(u,now);
renderRankBadge($("userRankArea"),total);renderProgress($("userProgressArea"),total);
if(u.stamps[ymd(now)])$("stampDoneBanner").classList.remove("hidden");else $("stampDoneBanner").classList.add("hidden");
$("userMonthLabel").textContent=monthLabelJa(userMonthCursor);
// Stamp request area (student only)
const sra=$("stampRequestArea");sra.innerHTML="";
if(u.userType==="学生"){
  if(u.pendingStampRequest&&u.pendingStampRequest.status==="pending"){
    sra.innerHTML=`<div style="margin-top:12px;"><span class="stamp-request-badge">📨 スタンプ修正申請中</span></div>`;
    stampEditMode=false;
  } else if(u.pendingStampRequest&&u.pendingStampRequest.status==="approved"){
    const wrap=document.createElement("div");wrap.style.cssText="margin-top:12px;padding:12px;border-radius:12px;background:rgba(107,203,119,.12);border:1.5px solid rgba(107,203,119,.3);text-align:center;";
    wrap.innerHTML=`<div style="font-weight:900;color:var(--mint);font-size:14px;">✅ スタンプ修正が承認されました！</div><div style="font-size:11px;color:var(--muted);margin-top:4px;">カレンダーに反映済みです</div>`;
    const dismissBtn=document.createElement("button");dismissBtn.className="btn ghost small";dismissBtn.style.cssText="margin-top:8px;font-size:11px;";dismissBtn.textContent="確認しました";
    dismissBtn.addEventListener("click",()=>{u.pendingStampRequest=null;saveData(data);renderUserHome()});
    wrap.appendChild(dismissBtn);sra.appendChild(wrap);
    stampEditMode=false;
  } else if(u.pendingStampRequest&&u.pendingStampRequest.status==="rejected"){
    const wrap=document.createElement("div");wrap.style.cssText="margin-top:12px;padding:12px;border-radius:12px;background:rgba(255,71,87,.08);border:1.5px solid rgba(255,71,87,.2);text-align:center;";
    wrap.innerHTML=`<div style="font-weight:900;color:var(--red);font-size:14px;">❌ スタンプ修正申請が却下されました</div><div style="font-size:11px;color:var(--muted);margin-top:4px;">必要に応じて再度申請してください</div>`;
    const dismissBtn=document.createElement("button");dismissBtn.className="btn ghost small";dismissBtn.style.cssText="margin-top:8px;font-size:11px;";dismissBtn.textContent="確認しました";
    dismissBtn.addEventListener("click",()=>{u.pendingStampRequest=null;saveData(data);renderUserHome()});
    wrap.appendChild(dismissBtn);sra.appendChild(wrap);
    stampEditMode=false;
  } else if(stampEditMode){
    const cancelBtn=document.createElement("button");cancelBtn.className="btn ghost";cancelBtn.style.cssText="width:100%;margin-top:12px;font-size:13px;padding:10px;";
    cancelBtn.textContent="✕ 修正モードを終了";
    cancelBtn.addEventListener("click",()=>{stampEditMode=false;stampEditStamps={};stampEditEmergencyMode=false;renderUserHome()});
    sra.appendChild(cancelBtn);
  } else {
    const reqBtn=document.createElement("button");reqBtn.className="btn stamp-request-btn";reqBtn.textContent="📝 スタンプ修正申請";
    reqBtn.addEventListener("click",()=>{stampEditMode=true;stampEditEmergencyMode=false;stampEditStamps=JSON.parse(JSON.stringify(u.stamps));renderUserHome()});
    sra.appendChild(reqBtn);
  }
}
// Calendar rendering
if(stampEditMode){
  renderCalendar({mount:$("userCal"),monthCursor:userMonthCursor,stampedMap:stampEditStamps,clickable:true,
    onDayClick:d=>{const k=ymd(d);const cur=stampEditStamps[k];
      if(stampEditEmergencyMode){
        if(!cur)stampEditStamps[k]="emergency";
        else if(cur==="emergency")delete stampEditStamps[k];
        else{stampEditStamps[k]="emergency"}
      } else {
        if(!cur)stampEditStamps[k]=true;else if(cur===true)delete stampEditStamps[k];else if(cur==="emergency")delete stampEditStamps[k];
      }
      renderUserHome()},
    pendingChanges:stampEditStamps,originalStamps:u.stamps});
  const bar=$("stampApplyBar");bar.classList.remove("hidden");bar.innerHTML="";
  bar.className="stamp-apply-bar";bar.style.flexWrap="wrap";
  // Legend
  const legendDiv=document.createElement("div");legendDiv.style.cssText="width:100%;font-size:11px;color:var(--muted);margin-bottom:6px;display:flex;gap:12px;align-items:center;flex-wrap:wrap;";
  legendDiv.innerHTML=`<span>✓ 通常出勤</span><span>⚡ 緊急出勤（3pt）</span><span style="color:var(--orange);">＋ 追加申請</span><span style="color:var(--red);">− 削除申請</span>`;
  bar.appendChild(legendDiv);
  // Mode toggle
  const modeDiv=document.createElement("div");modeDiv.style.cssText="display:flex;gap:6px;align-items:center;flex:1;";
  const normalBtn=document.createElement("button");normalBtn.className="btn small"+(stampEditEmergencyMode?"":" primary");normalBtn.textContent="✓ 通常";
  normalBtn.addEventListener("click",()=>{stampEditEmergencyMode=false;renderUserHome()});
  const emgBtn=document.createElement("button");emgBtn.className="btn small"+(stampEditEmergencyMode?" primary":"");emgBtn.style.cssText=stampEditEmergencyMode?"background:linear-gradient(135deg,#ff9a56,#ffd93d);border-color:rgba(255,154,86,.3);color:#fff;":"";
  emgBtn.textContent="⚡ 緊急出勤";
  emgBtn.addEventListener("click",()=>{stampEditEmergencyMode=true;renderUserHome()});
  modeDiv.appendChild(normalBtn);modeDiv.appendChild(emgBtn);bar.appendChild(modeDiv);
  const applyBtn=document.createElement("button");applyBtn.className="btn primary small";applyBtn.textContent="📨 申請する";
  applyBtn.addEventListener("click",()=>{
    u.pendingStampRequest={stamps:JSON.parse(JSON.stringify(stampEditStamps)),status:"pending",createdAt:Date.now()};
    saveData(data);stampEditMode=false;stampEditStamps={};stampEditEmergencyMode=false;
    showModal({title:"申請完了",sub:"管理者の承認をお待ちください",big:"📨"});renderUserHome()});
  bar.appendChild(applyBtn);
} else {
  renderCalendar({mount:$("userCal"),monthCursor:userMonthCursor,stampedMap:u.stamps,clickable:false,onDayClick:null});
  $("stampApplyBar").classList.add("hidden");
}}

/* === REPORT FORM === */
let editingReportIdx=-1;
let adminEditingReportMode=false;
let adminEditOrigUserId="";
function popSel(sel,opts,s0){sel.innerHTML="";opts.forEach(o=>{const op=document.createElement("option");op.value=o;op.textContent=o;sel.appendChild(op)});if(s0!=null)sel.value=s0}
function initReportForm(){const u=data.users[data.session.userId];if(!u)return;
$("reportUserName").textContent=(adminEditingReportMode?"【管理編集】":"")+( u.name||u.id);
if(!rpTransportLocked){$("rpTransport").readOnly=false;$("rpTransport").style.opacity="1"}
const hrs=[];for(let i=0;i<24;i++)hrs.push(pad2(i));const mins=["00","15","30","45"];
popSel($("rpStartH"),hrs);popSel($("rpStartM"),mins);popSel($("rpEndH"),hrs);popSel($("rpEndM"),mins);
const now=new Date();$("rpStartH").value=pad2(now.getHours());const rm=Math.round(now.getMinutes()/15)*15;$("rpStartM").value=pad2(rm>=60?45:rm);
$("rpEndH").value=pad2(Math.min(23,now.getHours()+1));$("rpEndM").value=$("rpStartM").value;
const brks=[];for(let i=0;i<=90;i+=15)brks.push(i+"分");popSel($("rpBreak"),brks,"60分");
popSel($("rpTaskType"),getTaskTypes());const mh=[];for(let i=1;i<=100;i++)mh.push(String(i));popSel($("rpManHours"),mh,"1");
popSel($("rpBizId"),BIZ_IDS);popSel($("rpProductId"),PRODUCT_IDS);popSel($("rpServiceId"),SERVICE_IDS);
const years=[];for(let y=2025;y<=2030;y++)years.push(String(y));popSel($("rpYear"),years,"2026");
$("rpDate").value=ymd(new Date());$("rpWorkType").value="出勤";$("rpTransport").value="";$("rpTextCode").value="";$("rpContent").value="";toggleWorkType();calcWorkTime();
if(editingReportIdx>=0&&u.reports[editingReportIdx]){const r=u.reports[editingReportIdx];$("rpDate").value=r.date||ymd(new Date());$("rpWorkType").value=r.workType||"出勤";if(r.startH)$("rpStartH").value=r.startH;if(r.startM)$("rpStartM").value=r.startM;if(r.endH)$("rpEndH").value=r.endH;if(r.endM)$("rpEndM").value=r.endM;if(r.breakTime)$("rpBreak").value=r.breakTime;if(r.taskType)$("rpTaskType").value=r.taskType;if(r.manHours)$("rpManHours").value=r.manHours;$("rpTransport").value=r.transport||"";if(r.bizId)$("rpBizId").value=r.bizId;if(r.productId)$("rpProductId").value=r.productId;if(r.serviceId)$("rpServiceId").value=r.serviceId;$("rpTextCode").value=r.textCode||"";if(r.year)$("rpYear").value=r.year;$("rpContent").value=r.content||"";toggleWorkType();calcWorkTime()}}
function toggleWorkType(){$("taskSection").classList.toggle("hidden",$("rpWorkType").value==="出勤");$("officeSection").classList.toggle("hidden",$("rpWorkType").value==="在宅")}
$("rpWorkType").addEventListener("change",toggleWorkType);
function calcWorkTime(){const sh=parseInt($("rpStartH").value)||0,sm=parseInt($("rpStartM").value)||0,eh=parseInt($("rpEndH").value)||0,em=parseInt($("rpEndM").value)||0,brk=parseInt($("rpBreak").value)||0;let d=(eh*60+em)-(sh*60+sm)-brk;if(d<0)d=0;$("rpWorkTime").value=`${Math.floor(d/60)}時間${d%60>0?d%60+"分":""}`}
["rpStartH","rpStartM","rpEndH","rpEndM","rpBreak"].forEach(id=>$(id).addEventListener("change",calcWorkTime));
["rpTransport","rpTextCode"].forEach(id=>$(id).addEventListener("input",function(){this.value=this.value.replace(/[^0-9]/g,"")}));
$("btnAddReport").addEventListener("click",()=>{const u=data.users[data.session.userId];if(!u)return;const wt=$("rpWorkType").value;
const e={date:$("rpDate").value,workType:wt,startH:$("rpStartH").value,startM:$("rpStartM").value,endH:$("rpEndH").value,endM:$("rpEndM").value,breakTime:$("rpBreak").value,workTime:$("rpWorkTime").value};
if(wt==="在宅"){e.taskType=$("rpTaskType").value;e.manHours=$("rpManHours").value}else{e.transport=$("rpTransport").value;e.bizId=$("rpBizId").value;e.productId=$("rpProductId").value;e.serviceId=$("rpServiceId").value;e.textCode=$("rpTextCode").value;e.year=$("rpYear").value}
e.content=$("rpContent").value;e.proofCount=0;
if(editingReportIdx>=0){e.proofCount=u.reports[editingReportIdx].proofCount||0;u.reports[editingReportIdx]=e}else u.reports.push(e);
saveData(data);editingReportIdx=-1;
if(adminEditingReportMode){adminEditingReportMode=false;data.session.userId=adminEditOrigUserId;saveLocalOnly(data);location.hash="#admin-report-detail"}
else{location.hash="#report-confirm"}});

/* Add and continue same-day */
$("btnAddAndContinue").addEventListener("click",()=>{
  const u=data.users[data.session.userId];if(!u)return;const wt=$("rpWorkType").value;
  const e={date:$("rpDate").value,workType:wt,startH:$("rpStartH").value,startM:$("rpStartM").value,endH:$("rpEndH").value,endM:$("rpEndM").value,breakTime:$("rpBreak").value,workTime:$("rpWorkTime").value};
  if(wt==="在宅"){e.taskType=$("rpTaskType").value;e.manHours=$("rpManHours").value}else{e.transport=rpTransportLocked?"0":$("rpTransport").value;e.bizId=$("rpBizId").value;e.productId=$("rpProductId").value;e.serviceId=$("rpServiceId").value;e.textCode=$("rpTextCode").value;e.year=$("rpYear").value}
  e.content=$("rpContent").value;e.proofCount=0;
  if(editingReportIdx>=0){e.proofCount=u.reports[editingReportIdx].proofCount||0;u.reports[editingReportIdx]=e}else u.reports.push(e);
  saveData(data);
  const prevDate=$("rpDate").value,prevEndH=$("rpEndH").value,prevEndM=$("rpEndM").value;
  editingReportIdx=-1;rpTransportLocked=true;
  initReportForm();
  $("rpDate").value=prevDate;$("rpStartH").value=prevEndH;$("rpStartM").value=prevEndM;
  $("rpTransport").value="0";$("rpTransport").readOnly=true;$("rpTransport").style.opacity="0.5";
  calcWorkTime();
  showModal({title:"保存して次へ",sub:"同日の次の業務を入力してください",big:"✅"});
});

/* === REPORT CONFIRM === */
let rcInit=false;
function renderReportConfirm(){const u=data.users[data.session.userId];if(!u)return;$("confirmUserName").textContent=u.name||u.id;
u.userType==="社会人"?$("confirmBackToCal").classList.add("hidden"):$("confirmBackToCal").classList.remove("hidden");
const dr=getUserDateRange(u);buildYearMonthOpts($("rcFilterYear"),$("rcFilterMonth"),dr,!rcInit);if(!rcInit){$("rcFilterType").value="出勤";rcInit=true}doRenderReportList()}
function doRenderReportList(){const u=data.users[data.session.userId];if(!u)return;const y=$("rcFilterYear").value,m=$("rcFilterMonth").value,wt=$("rcFilterType").value;
const filtered=filterReports(u.reports,y,m,wt);$("reportThead").innerHTML=`<tr><th>#</th><th>日付</th><th>形態</th><th>開始</th><th>終了</th><th>休憩</th><th>勤務</th><th>業務ID</th><th>商品ID</th><th>サービスID</th><th>テキストコード</th><th>年度</th><th>内容</th><th>給与</th><th>詳細</th><th>操作</th></tr>`;
const tb=$("reportTbody");tb.innerHTML="";let tMin=0,tSal=0,tTr=0,tInc=0;
if(!filtered.length){tb.innerHTML=`<tr><td colspan="16" style="text-align:center;padding:20px;color:var(--muted);">データなし</td></tr>`}else{
filtered.forEach((r,fi)=>{const oi=u.reports.indexOf(r);const sal=Math.round(calcReportSalary(r,u.id));const mins=calcWorkMinutes(r);const tr_t=r.workType==="出勤"?parseInt(r.transport)||0:0;tMin+=mins;tSal+=sal;tTr+=tr_t;
const isShakaijin=(u.userType||"学生")==="社会人";
if(isShakaijin){tInc+=r.incentiveAmount||0}else{tInc+=(r.proofCount||0)*500}
const tr=document.createElement("tr");const extra=r.workType==="在宅"?`${r.taskType||""} ×${r.manHours||""}`:`交通費:${r.transport||0}円`;
[String(fi+1),r.date,r.workType,`${r.startH}:${r.startM}`,`${r.endH}:${r.endM}`,r.breakTime,r.workTime].forEach(c=>{const td=document.createElement("td");td.textContent=c;tr.appendChild(td)});
// 業務ID, 商品ID, サービスID, テキストコード, 年度, 内容
[r.bizId||"",r.productId||"",r.serviceId||"",r.textCode||"",r.year||"",(r.content||"").substring(0,30)].forEach(c=>{const td=document.createElement("td");td.textContent=c;tr.appendChild(td)});
// 給与
const tdS=document.createElement("td");tdS.textContent=sal.toLocaleString()+"円";tdS.style.fontWeight="700";tr.appendChild(tdS);
// 詳細
const tdE=document.createElement("td");tdE.textContent=extra;tr.appendChild(tdE);
// 操作
const tdA=document.createElement("td");const b1=document.createElement("button");b1.className="btn primary small";b1.textContent="編集";b1.addEventListener("click",()=>{editingReportIdx=oi;location.hash="#report-input"});tdA.appendChild(b1);
const b2=document.createElement("button");b2.className="btn danger small";b2.textContent="削除";b2.style.marginLeft="4px";b2.addEventListener("click",()=>{u.reports.splice(oi,1);saveData(data);renderReportConfirm()});tdA.appendChild(b2);tr.appendChild(tdA);
tb.appendChild(tr)})}
const tH=Math.floor(tMin/60),tM=tMin%60;
$("rcSummaryBar").innerHTML=`<div class="summary-chip"><div><div class="sk">勤務時間</div><div class="sv">${tH}h${tM>0?tM+"m":""}</div></div></div><div class="summary-chip"><div><div class="sk">インセンティブ</div><div class="sv">${tInc.toLocaleString()}円</div></div></div><div class="summary-chip"><div><div class="sk">給料（ｲﾝｾﾝﾃｨﾌﾞあり）</div><div class="sv">${(tSal+tInc).toLocaleString()}円</div></div></div><div class="summary-chip"><div><div class="sk">交通費</div><div class="sv">${tTr.toLocaleString()}円</div></div></div>`}
["rcFilterYear","rcFilterMonth","rcFilterType"].forEach(id=>$(id).addEventListener("change",doRenderReportList));

/* === ADMIN REPORT MGMT === */
let armInit=false;
function renderAdminReportMgmt(){renderAdminNotifications();const now=new Date();$("armMonthInfo").textContent=monthLabelJa(now);const dr=getAllUsersDateRange();buildYearMonthOpts($("armFilterYear"),$("armFilterMonth"),dr,!armInit);if(!armInit){$("armFilterType").value="出勤";armInit=true}doRenderARM()}
function doRenderARM(){const y=$("armFilterYear").value,m=$("armFilterMonth").value,wt=$("armFilterType").value,ut=$("armFilterUserType").value;
let users=Object.values(data.users);
// Filter by user type
if(ut!=="全て")users=users.filter(u=>(u.userType||"学生")===ut);
// Sort: 社会人 first, then 学生, within each group by createdAt
users.sort((a,b)=>{const aType=(a.userType||"学生")==="社会人"?0:1;const bType=(b.userType||"学生")==="社会人"?0:1;if(aType!==bType)return aType-bType;return(a.createdAt||0)-(b.createdAt||0)});
const tb=$("armTbody");tb.innerHTML="";users.forEach((u,idx)=>{const filtered=filterReports(u.reports,y,m,wt);let tMin=0,tSal=0,tTr=0,cnt=0,pInc=0;
filtered.forEach(r=>{tMin+=calcWorkMinutes(r);tSal+=Math.round(calcReportSalary(r,u.id));if(r.workType==="出勤")tTr+=parseInt(r.transport)||0;cnt++;
  if(u.userType==="社会人"){pInc+=r.incentiveAmount||0}else{pInc+=(r.proofCount||0)*500}});
const total=countTotal(u);const sInc=calcStampIncentive(total);const pay=tSal+pInc+sInc;const tH=Math.floor(tMin/60),tMm=tMin%60;
const tr=document.createElement("tr");
[String(idx+1),null,u.name||u.id,null,`${tH}h${tMm>0?tMm+"m":""}`,cnt+"回",tSal.toLocaleString()+"円",pInc.toLocaleString()+"円",sInc.toLocaleString()+"円",pay.toLocaleString()+"円",tTr.toLocaleString()+"円"].forEach((c,i)=>{
const td=document.createElement("td");if(i===1)td.innerHTML=`<span class="tag">${escapeHtml(u.id)}</span>`;
else if(i===3){const cls=u.userType==="社会人"?"tag shakaijin":"tag student";td.innerHTML=`<span class="${cls}">${escapeHtml(u.userType||"学生")}</span>`}
else{td.textContent=c;if(i===9){td.style.fontWeight="900";td.style.color="var(--pink)"}}tr.appendChild(td)});
const tdE=document.createElement("td");const btn=document.createElement("button");btn.className="btn primary small";btn.textContent="詳細";
btn.addEventListener("click",()=>{data.session.adminReportEditingUserId=u.id;saveLocalOnly(data);ardInit=false;location.hash="#admin-report-detail"});
tdE.appendChild(btn);tr.appendChild(tdE);tb.appendChild(tr)});$("armMeta").textContent=`${users.length}名`}
["armFilterYear","armFilterMonth","armFilterType","armFilterUserType"].forEach(id=>$(id).addEventListener("change",doRenderARM));

// ARM sub-tabs: サマリー / ユーザー追加
$("armSubTabSummary").addEventListener("click",()=>{
  $("armSubTabSummary").classList.add("active");$("armSubTabAddUser").classList.remove("active");
  $("armSummaryCard").classList.remove("hidden");$("armAddUserCard").classList.add("hidden");
});
$("armSubTabAddUser").addEventListener("click",()=>{
  $("armSubTabAddUser").classList.add("active");$("armSubTabSummary").classList.remove("active");
  $("armAddUserCard").classList.remove("hidden");$("armSummaryCard").classList.add("hidden");
});

// ARM Excel Export
$("armExportExcel").addEventListener("click",()=>{
  const y=$("armFilterYear").value,m=$("armFilterMonth").value,wt=$("armFilterType").value,ut=$("armFilterUserType").value;
  let users=Object.values(data.users);
  if(ut!=="全て")users=users.filter(u=>(u.userType||"学生")===ut);
  users.sort((a,b)=>{const aT=(a.userType||"学生")==="社会人"?0:1;const bT=(b.userType||"学生")==="社会人"?0:1;return aT!==bT?aT-bT:(a.createdAt||0)-(b.createdAt||0)});
  const label=[y!=="全て"?y+"年":"",m!=="全て"?m+"月":"",wt!=="全て"?wt:""].filter(x=>x).join("_")||"全期間";
  // Build Excel-compatible HTML table
  let html=`<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head><meta charset="utf-8"><style>td,th{mso-number-format:'\\@';border:1px solid #ccc;padding:4px 8px;font-size:11pt;font-family:'Yu Gothic',sans-serif;}
th{background:#4472C4;color:#fff;font-weight:bold;text-align:center;}
.num{mso-number-format:'#\\,##0';text-align:right;}
.total{font-weight:bold;background:#FFF2CC;}.highlight{font-weight:bold;color:#C00000;}
tr:nth-child(even){background:#F2F2F2;}
</style></head><body><table>`;
  html+=`<tr><th>#</th><th>ID</th><th>名前</th><th>種別</th><th>勤務</th><th>回数</th><th>給与</th><th>校正/ｲﾝｾﾝﾃｨﾌﾞ</th><th>ｽﾀﾝﾌﾟ</th><th>合計</th><th>交通費</th></tr>`;
  let gSal=0,gPInc=0,gSInc=0,gPay=0,gTr=0,gCnt=0;
  users.forEach((u,idx)=>{
    const filtered=filterReports(u.reports,y,m,wt);let tMin=0,tSal=0,tTr=0,cnt=0,pInc=0;
    filtered.forEach(r=>{tMin+=calcWorkMinutes(r);tSal+=Math.round(calcReportSalary(r,u.id));if(r.workType==="出勤")tTr+=parseInt(r.transport)||0;cnt++;
      if(u.userType==="社会人"){pInc+=r.incentiveAmount||0}else{pInc+=(r.proofCount||0)*500}});
    const total=countTotal(u);const sInc=calcStampIncentive(total);const pay=tSal+pInc+sInc;
    const tH=Math.floor(tMin/60),tMm=tMin%60;
    gSal+=tSal;gPInc+=pInc;gSInc+=sInc;gPay+=pay;gTr+=tTr;gCnt+=cnt;
    html+=`<tr><td>${idx+1}</td><td>${escapeHtml(u.id)}</td><td>${escapeHtml(u.name||u.id)}</td><td>${escapeHtml(u.userType||"学生")}</td><td>${tH}h${tMm>0?tMm+"m":""}</td><td class="num">${cnt}</td><td class="num">${tSal}</td><td class="num">${pInc}</td><td class="num">${sInc}</td><td class="num highlight">${pay}</td><td class="num">${tTr}</td></tr>`;
  });
  html+=`<tr class="total"><td colspan="5"></td><td class="num total">${gCnt}</td><td class="num total">${gSal}</td><td class="num total">${gPInc}</td><td class="num total">${gSInc}</td><td class="num total highlight">${gPay}</td><td class="num total">${gTr}</td></tr>`;
  html+=`</table></body></html>`;
  const blob=new Blob([html],{type:"application/vnd.ms-excel;charset=utf-8;"});
  const a=document.createElement("a");a.href=URL.createObjectURL(blob);
  a.download=`業務日報サマリー_${label}.xls`;document.body.appendChild(a);a.click();a.remove();
  showModal({title:"エクスポート完了",sub:`業務日報サマリー_${label}.xls`,big:"📥"});
});

// User add (moved here)
$("btnAddUser").addEventListener("click", async ()=>{const id=$("newUserId").value.trim(),pw=$("newUserPw").value.trim(),name=$("newUserName").value.trim()||id,utype=$("newUserType").value;
if(!id||!pw){showModal({title:"入力不足",big:"⚠️"});return}
if(!API_URL){showModal({title:"API未接続",sub:"⚙でURLを設定してください",big:"🔌"});return}
try{
  const resp=await fetch(API_URL,{method:"POST",headers:{"Content-Type":"text/plain"},body:JSON.stringify({_action:"upsertStaffUser",token:getToken(),id,pw,name,userType:utype}),redirect:"follow"});
  const r=await resp.json();
  applySyncMeta(syncMetaFromResult(r));
  if(!r.ok){showModal({title:"追加失敗",sub:r.error||"エラー",big:"🚫"});return}
// ローカルデータにも PW を保持して、編集オーバーレイから確認できるようにする
  data.users=data.users||{};
  const existing = data.users[id] || {};
  data.users[id]=Object.assign({id:id, stamps:{}, incentives:{}, bonusPoints:0, lastCongrats50:0, lastMonthFirstStamp:"", reports:[], createdAt:Date.now(), proofingIncentives:{}, pendingStampRequest:null}, existing, {name:name, userType:utype});
  invalidateStaffAccountsCache();
  saveLocalOnly(data);
  $("newUserId").value="";$("newUserPw").value="";$("newUserName").value="";
  showModal({title:"追加しました",sub:"ログイン情報はサーバ側に保存されました。",big:"✅"});
}catch(e){showModal({title:"通信エラー",sub:"ユーザー追加に失敗しました。",big:"📡"});}
});

/* === ADMIN REPORT DETAIL === */
let ardInit=false;
function renderAdminReportDetail(){const u=data.users[data.session.adminReportEditingUserId];if(!u){location.hash="#admin-report-mgmt";return}
$("ardUserName").textContent=u.name||u.id;$("ardUserId").textContent=u.id;const dr=getUserDateRange(u);buildYearMonthOpts($("ardFilterYear"),$("ardFilterMonth"),dr,!ardInit);if(!ardInit){$("ardFilterType").value="出勤";ardInit=true}
$("ardProofLabel").textContent=u.userType==="社会人"?"ｲﾝｾﾝﾃｨﾌﾞ合計":"校正ｲﾝｾﾝﾃｨﾌﾞ合計";
doRenderARD()}
function doRenderARD(){const u=data.users[data.session.adminReportEditingUserId];if(!u)return;const y=$("ardFilterYear").value,m=$("ardFilterMonth").value,wt=$("ardFilterType").value;
const filtered=filterReports(u.reports,y,m,wt);
const isShakaijin=u.userType==="社会人";
const lastColLabel=isShakaijin?"ｲﾝｾﾝﾃｨﾌﾞ":"校正";
const isZaitaku=wt==="在宅";
const extraHeaders=isZaitaku?"<th>業務種類</th><th>工数</th>":"<th>業務ID</th><th>商品ID</th><th>サービスID</th><th>テキストコード</th><th>交通費</th>";
$("ardThead").innerHTML=`<tr><th>#</th><th>日付</th><th>形態</th><th>開始</th><th>終了</th><th>休憩</th><th>勤務</th><th>給与</th>${extraHeaders}<th>内容</th><th>${lastColLabel}</th><th>操作</th></tr>`;
const tb=$("ardTbody");tb.innerHTML="";let tMin=0,tSal=0,tTr=0,tProof=0;
if(!filtered.length){tb.innerHTML=`<tr><td colspan="${isZaitaku?13:17}" style="text-align:center;padding:20px;color:var(--muted);">データなし</td></tr>`}else{
filtered.forEach((r,fi)=>{const oi=u.reports.indexOf(r);const sal=Math.round(calcReportSalary(r,u.id));const mins=calcWorkMinutes(r);const tr_a=r.workType==="出勤"?parseInt(r.transport)||0:0;
tMin+=mins;tSal+=sal;tTr+=tr_a;
const tr=document.createElement("tr");const mkCell=(v,f,ed)=>{const td=document.createElement("td");td.textContent=v;if(ed){td.classList.add("editable");td.addEventListener("dblclick",()=>startEdit(td,u,oi,f))}return td};
tr.appendChild(mkCell(String(fi+1),null,false));tr.appendChild(mkCell(r.date,"date",true));tr.appendChild(mkCell(r.workType,"workType",true));
tr.appendChild(mkCell(`${r.startH}:${r.startM}`,"startTime",true));tr.appendChild(mkCell(`${r.endH}:${r.endM}`,"endTime",true));tr.appendChild(mkCell(r.breakTime,"breakTime",true));
tr.appendChild(mkCell(r.workTime,null,false));const tdSal=mkCell(sal.toLocaleString()+"円",null,false);tdSal.style.fontWeight="700";tr.appendChild(tdSal);
if(r.workType==="在宅"){
tr.appendChild(mkCell(r.taskType||"","taskType",true));tr.appendChild(mkCell(r.manHours||"","manHours",true));
}else{
tr.appendChild(mkCell(r.bizId||"","bizId",true));tr.appendChild(mkCell(r.productId||"","productId",true));tr.appendChild(mkCell(r.serviceId||"","serviceId",true));tr.appendChild(mkCell(r.textCode||"","textCode",true));tr.appendChild(mkCell(r.transport||"0","transport",true));
}
tr.appendChild(mkCell((r.content||"").substring(0,30),"content",true));
// Proof / Incentive column
const tdP=document.createElement("td");
if(isShakaijin){
  // インセンティブ for 社会人
  const workHrs=mins/60;const hr=getUserHourlyRate(u.id);
  const defaultInc=Math.ceil(hr*workHrs/10);
  if(r.incentiveAmount==null){r.incentiveAmount=defaultInc;saveData(data)}
  tProof+=r.incentiveAmount;
  const pD=document.createElement("div");pD.className="proof-inline";
  const pI=document.createElement("input");pI.type="number";pI.min="0";pI.value=r.incentiveAmount;pI.style.width="70px";
  const pA=document.createElement("span");pA.className="proof-amt";pA.textContent=r.incentiveAmount>0?`=${r.incentiveAmount.toLocaleString()}円`:"";
  pI.addEventListener("change",()=>{const c=Math.max(0,parseInt(pI.value)||0);u.reports[oi].incentiveAmount=c;saveData(data);pA.textContent=c>0?`=${c.toLocaleString()}円`:"";updateARDTotals()});
  pD.appendChild(pI);pD.appendChild(pA);tdP.appendChild(pD);
} else {
  // 校正 for 学生
  const pD=document.createElement("div");pD.className="proof-inline";
  const pI=document.createElement("input");pI.type="number";pI.min="0";pI.value=r.proofCount||0;
  const pA=document.createElement("span");pA.className="proof-amt";const pc=r.proofCount||0;pA.textContent=pc>0?`=${(pc*500).toLocaleString()}円`:"";
  pI.addEventListener("change",()=>{const c=Math.max(0,parseInt(pI.value)||0);u.reports[oi].proofCount=c;saveData(data);pA.textContent=c>0?`=${(c*500).toLocaleString()}円`:"";updateARDTotals()});
  pD.appendChild(pI);pD.appendChild(pA);tdP.appendChild(pD);
  tProof+=(r.proofCount||0)*500;
}
tr.appendChild(tdP);
// 操作 (edit/delete)
const tdAct=document.createElement("td");tdAct.style.whiteSpace="nowrap";
// 編集ボタン削除（ダブルクリックで直接編集）
const bDel=document.createElement("button");bDel.className="btn danger small";bDel.style.marginLeft="4px";bDel.textContent="削除";
bDel.addEventListener("click",()=>{
  if(!confirm("この日報を削除しますか？"))return;
  u.reports.splice(oi,1);saveData(data);doRenderARD();
});tdAct.appendChild(bDel);tr.appendChild(tdAct);
tb.appendChild(tr)})}updateARDTotals()}
function updateARDTotals(){const u=data.users[data.session.adminReportEditingUserId];if(!u)return;const y=$("ardFilterYear").value,m=$("ardFilterMonth").value,wt=$("ardFilterType").value;
const isShakaijin=u.userType==="社会人";
const f=filterReports(u.reports,y,m,wt);let tM=0,tS=0,tT=0,tP=0;f.forEach(r=>{tM+=calcWorkMinutes(r);tS+=Math.round(calcReportSalary(r,u.id));if(r.workType==="出勤")tT+=parseInt(r.transport)||0;
  if(isShakaijin){tP+=r.incentiveAmount||0}else{tP+=(r.proofCount||0)*500}});
const h=Math.floor(tM/60),mm=tM%60;
$("ardSummaryBar").innerHTML=`<div class="summary-chip"><div><div class="sk">勤務</div><div class="sv">${h}h${mm>0?mm+"m":""}</div></div></div><div class="summary-chip"><div><div class="sk">ｲﾝｾﾝﾃｨﾌﾞ</div><div class="sv">${tP.toLocaleString()}円</div></div></div><div class="summary-chip"><div><div class="sk">給料(ｲﾝｾﾝﾃｨﾌﾞ含)</div><div class="sv">${(tS+tP).toLocaleString()}円</div></div></div><div class="summary-chip"><div><div class="sk">交通費</div><div class="sv">${tT.toLocaleString()}円</div></div></div>`;
$("ardProofTotal").textContent=tP.toLocaleString()+"円"}

function startEdit(td,u,oi,field){if(td.classList.contains("editing"))return;td.classList.add("editing");const r=u.reports[oi];let inp;
if(field==="workType"){inp=document.createElement("select");["出勤","在宅"].forEach(v=>{const o=document.createElement("option");o.value=v;o.textContent=v;inp.appendChild(o)});inp.value=r.workType}
else if(field==="date"){inp=document.createElement("input");inp.type="date";inp.value=r.date||""}
else if(field==="startTime"){inp=document.createElement("input");inp.type="text";inp.value=`${r.startH}:${r.startM}`}
else if(field==="endTime"){inp=document.createElement("input");inp.type="text";inp.value=`${r.endH}:${r.endM}`}
else if(field==="breakTime"){inp=document.createElement("select");for(let i=0;i<=90;i+=15){const o=document.createElement("option");o.value=i+"分";o.textContent=i+"分";inp.appendChild(o)}inp.value=r.breakTime||"0分"}
else if(field==="content"){inp=document.createElement("textarea");inp.value=r.content||""}
else if(field==="taskType"){inp=document.createElement("select");getTaskTypes().forEach(v=>{const o=document.createElement("option");o.value=v;o.textContent=v;inp.appendChild(o)});inp.value=r.taskType||""}
else if(field==="manHours"){inp=document.createElement("input");inp.type="number";inp.min="0";inp.value=r.manHours||""}
else if(field==="bizId"){inp=document.createElement("input");inp.type="text";inp.value=r.bizId||""}
else if(field==="productId"){inp=document.createElement("input");inp.type="text";inp.value=r.productId||""}
else if(field==="serviceId"){inp=document.createElement("input");inp.type="text";inp.value=r.serviceId||""}
else if(field==="textCode"){inp=document.createElement("input");inp.type="text";inp.value=r.textCode||""}
else if(field==="transport"){inp=document.createElement("input");inp.type="text";inp.value=r.transport||"0"}
else{inp=document.createElement("input");inp.type="text";inp.value=td.textContent}
td.textContent="";td.appendChild(inp);inp.focus();
const save=()=>{td.classList.remove("editing");
if(field==="date")r.date=inp.value;else if(field==="workType")r.workType=inp.value;
else if(field==="startTime"){const p=inp.value.split(":");r.startH=pad2(parseInt(p[0])||0);r.startM=pad2(parseInt(p[1])||0)}
else if(field==="endTime"){const p=inp.value.split(":");r.endH=pad2(parseInt(p[0])||0);r.endM=pad2(parseInt(p[1])||0)}
else if(field==="breakTime")r.breakTime=inp.value;else if(field==="content")r.content=inp.value;
else if(field==="taskType")r.taskType=inp.value;else if(field==="manHours")r.manHours=inp.value;
else if(field==="bizId")r.bizId=inp.value;else if(field==="productId")r.productId=inp.value;else if(field==="serviceId")r.serviceId=inp.value;else if(field==="textCode")r.textCode=inp.value;else if(field==="transport")r.transport=inp.value;
else if(field==="extra"){/* complex, skip */}
const sh=parseInt(r.startH)||0,sm=parseInt(r.startM)||0,eh=parseInt(r.endH)||0,em=parseInt(r.endM)||0,brk=parseInt(r.breakTime)||0;
let d=(eh*60+em)-(sh*60+sm)-brk;if(d<0)d=0;r.workTime=`${Math.floor(d/60)}時間${d%60>0?d%60+"分":""}`;saveData(data);doRenderARD()};
inp.addEventListener("blur",save);inp.addEventListener("keydown",e=>{if(e.key==="Enter"&&field!=="content"){e.preventDefault();inp.blur()}})};
["ardFilterYear","ardFilterMonth","ardFilterType"].forEach(id=>$(id).addEventListener("change",doRenderARD));
$("ardBack").addEventListener("click",()=>{data.session.adminReportEditingUserId="";ardInit=false;saveLocalOnly(data);location.hash="#admin-report-mgmt"});
$("ardExport").addEventListener("click",()=>{const u=data.users[data.session.adminReportEditingUserId];if(!u)return;const y=$("ardFilterYear").value,m=$("ardFilterMonth").value,wt=$("ardFilterType").value;const f=filterReports(u.reports,y,m,wt);
let csv="\uFEFF#,日付,形態,開始,終了,休憩,勤務,給与,交通費,業務ID,商品ID,ｻｰﾋﾞｽID,ﾃｷｽﾄｺｰﾄﾞ,年度,業務種類,工数,内容,校正回数,校正金額\n";
f.forEach((r,i)=>{const sal=Math.round(calcReportSalary(r,u.id));const pc=r.proofCount||0;csv+=[i+1,r.date,r.workType,`${r.startH}:${r.startM}`,`${r.endH}:${r.endM}`,r.breakTime,r.workTime,sal,r.transport||0,r.bizId||"",r.productId||"",r.serviceId||"",r.textCode||"",r.year||"",r.taskType||"",r.manHours||"",`"${(r.content||"").replace(/"/g,'""')}"`,pc,pc*500].join(",")+"\n"});
const blob=new Blob([csv],{type:"text/csv;charset=utf-8;"});const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=`業務日報_${u.name||u.id}.csv`;document.body.appendChild(a);a.click();a.remove();showModal({title:"エクスポート完了",big:"📥"})});

/* === ADMIN STAMP HOME === */
$("btnExport").addEventListener("click",()=>{const b=new Blob([JSON.stringify(data,null,2)],{type:"application/json"});const a=document.createElement("a");a.href=URL.createObjectURL(b);a.download="stampcard_export.json";document.body.appendChild(a);a.click();a.remove();showModal({title:"エクスポート完了",big:"📦"})});
$("btnResetAll").addEventListener("click",()=>{if(!confirm("全ユーザー初期化？"))return;Object.values(data.users).forEach(u=>{u.stamps={};u.incentives={};u.bonusPoints=0;u.lastCongrats50=0;u.lastMonthFirstStamp="";u.reports=[];u.proofingIncentives={};delete u.stampScreenVisitedToday;delete u.stampFailed});saveData(data);renderAdminHome();showModal({title:"全初期化完了",big:"🧼"})});

function renderAdminHome(){renderAdminNotifications();const now=new Date();$("adminMonthInfo").textContent=monthLabelJa(now);$("adminTodayPassword").textContent="取得中…";fetchTodayPasswordForAdmin(false).then(p=>{$("adminTodayPassword").textContent=p||"―";}).catch(()=>{$("adminTodayPassword").textContent="―";});const users=Object.entries(data.users||{}).map(([id,u])=>{u=u||{}; if(!u.id) u.id=id; return u;}).filter(u=>u.userType==="学生").sort((a,b)=>(a.createdAt||0)-(b.createdAt||0));const tb=$("adminTbody");tb.innerHTML="";
users.forEach((u,idx)=>{const total=countTotal(u);const rank=getRank(total);const sInc=calcStampIncentive(total);
const tr=document.createElement("tr");[String(idx+1),null,u.name||u.id,null,countThisWeek(u,now)+"回",countThisMonth(u,now)+"回",total+"pt",null,rank.yen+"円",sInc.toLocaleString()+"円"].forEach((c,i)=>{
const td=document.createElement("td");if(i===1)td.innerHTML=`<span class="tag">${escapeHtml(u.id)}</span>`;
else if(i===3){const cls=u.userType==="社会人"?"tag shakaijin":"tag student";td.innerHTML=`<span class="${cls}">${escapeHtml(u.userType||"学生")}</span>`}
else if(i===7)td.innerHTML=`<span class="rank-badge rank-${rank.rank}" style="font-size:9px;padding:2px 8px;">${rank.emoji}R${rank.rank}</span>`;
else td.textContent=c;tr.appendChild(td)});
// 申請 column
const tdReq=document.createElement("td");
if(u.pendingStampRequest&&u.pendingStampRequest.status==="pending"){
  const reqBadge=document.createElement("span");reqBadge.className="stamp-request-badge";reqBadge.style.cssText="font-size:9px;cursor:pointer;";reqBadge.textContent="📨 申請中";
  reqBadge.addEventListener("click",()=>{data.session.adminEditingUserId=u.id;saveLocalOnly(data);editMonthCursor=startOfMonth(new Date());location.hash="#admin-edit"});
  tdReq.appendChild(reqBadge);
} else {tdReq.textContent="―"}
tr.appendChild(tdReq);
const tdE=document.createElement("td");const btn=document.createElement("button");btn.className="btn primary small";btn.textContent="編集";
btn.addEventListener("click",()=>{data.session.adminEditingUserId=u.id;saveLocalOnly(data);editMonthCursor=startOfMonth(new Date());location.hash="#admin-edit"});
tdE.appendChild(btn);tr.appendChild(tdE);tb.appendChild(tr)});$("adminMeta").textContent=`${users.length}名`}

/* === ADMIN EDIT === */
let editMonthCursor=startOfMonth(new Date());
$("ePrev").addEventListener("click",()=>{editMonthCursor=addMonths(editMonthCursor,-1);renderAdminEdit()});
$("eNext").addEventListener("click",()=>{editMonthCursor=addMonths(editMonthCursor,+1);renderAdminEdit()});
$("eThis").addEventListener("click",()=>{editMonthCursor=startOfMonth(new Date());renderAdminEdit()});
$("backToAdminHome").addEventListener("click",()=>{data.session.adminEditingUserId="";saveLocalOnly(data);location.hash="#admin"});
$("resetThisUser").addEventListener("click",()=>{const u=data.users[data.session.adminEditingUserId];if(!u)return;if(!confirm(`${u.name||u.id}を初期化？`))return;u.stamps={};u.incentives={};u.bonusPoints=0;u.lastCongrats50=0;u.lastMonthFirstStamp="";u.reports=[];u.proofingIncentives={};delete u.stampScreenVisitedToday;delete u.stampFailed;saveData(data);renderAdminEdit();showModal({title:"初期化完了",big:"🧼"})});
$("btnSaveUserInfo").addEventListener("click", async ()=>{
  const oldId=data.session.adminEditingUserId;const u=data.users[oldId];if(!u)return;
  const nid=$("editUid").value.trim(),nn=$("editUname").value.trim(),np=$("editUpw").value.trim(),nt=$("editUserType").value;
  if(!nid){showModal({title:"入力不足",sub:"IDは必須です",big:"⚠️"});return}
  if(nid!==oldId&&data.users[nid]){showModal({title:"ID重複",big:"🧩"});return}

  try {
  const resp = await fetch(API_URL, {
    method: "POST",
    headers: {"Content-Type": "text/plain"},
    body: JSON.stringify({
      _action: "upsertStaffUser",
      token: getToken(),
      id: nid,
      pw: np,
      name: nn||nid,
      userType: nt
    }),
    redirect: "follow"
  });
  const r = await resp.json();
  applySyncMeta(syncMetaFromResult(r));
  if (!r.ok) {
    showModal({title: "エラー", sub: r.error, big: "🚫"});
    return;
  }

  // ID変更時は旧IDをGAS側からも削除
  if (nid !== oldId) {
    const delResp = await fetch(API_URL, {
      method: "POST",
      headers: {"Content-Type": "text/plain"},
      body: JSON.stringify({
        _action: "deleteStaffUser",
        token: getToken(),
        id: oldId
      }),
      redirect: "follow"
    });
    const delResult = await delResp.json();
    applySyncMeta(syncMetaFromResult(delResult));
    if (!delResult.ok) {
      showModal({title: "旧ID削除失敗", sub: delResult.error || "エラー", big: "🚫"});
      return;
    }
  }
} catch(e) {
  showModal({title: "通信エラー", big: "📡"});
  return;
}

u.name=nn||nid;
u.userType=nt;
if(nid!==oldId){
  u.id=nid;
  data.users[nid]=u;
  delete data.users[oldId];

  if (data.userHourlyRates && data.userHourlyRates[oldId] != null) {
    data.userHourlyRates[nid] = data.userHourlyRates[oldId];
    delete data.userHourlyRates[oldId];
  }

  data.session.adminEditingUserId=nid;
}
saveData(data);
renderAdminEdit();
showModal({title:"更新完了",big:"✅"});
});
function renderAdminEdit(){const u=data.users[data.session.adminEditingUserId];if(!u){location.hash="#admin";return}
$("editUserName").textContent=u.name||u.id;$("editUserId").textContent=u.id;$("editUid").value=u.id;$("editUname").value=u.name||"";$("editUpw").value=u.pw||"";$("editUserType").value=u.userType||"学生";
fillStaffPasswordField("editUpw", u.id);
$("editUpw").value="";
const now=new Date();const total=countTotal(u);$("eTotal").textContent=total;$("eMonth").textContent=countThisMonth(u,now);$("eWeek").textContent=countThisWeek(u,now);$("eMonthKey").textContent=ym(editMonthCursor);
const sInc=calcStampIncentive(total);$("incentiveDisplay").innerHTML=`<div class="incentive-box"><div class="ib-title">💰 ｲﾝｾﾝﾃｨﾌﾞ（自動）</div><div style="font-family:var(--font-display);font-size:20px;font-weight:900;color:var(--pink);">${sInc.toLocaleString()}円</div><div style="font-size:11px;color:var(--muted);margin-top:4px;">累計${total}pt</div></div>`;
$("editMonthLabel").textContent=monthLabelJa(editMonthCursor);
// Show pending stamp request info
const hasPending=u.pendingStampRequest&&u.pendingStampRequest.status==="pending";
if(hasPending){
  renderCalendar({mount:$("editCal"),monthCursor:editMonthCursor,stampedMap:u.pendingStampRequest.stamps,clickable:true,
    onDayClick:d=>{const k=ymd(d);const target=u.pendingStampRequest.stamps||(u.pendingStampRequest.stamps={});const cur=target[k];if(!cur)target[k]=true;else if(cur===true)target[k]="emergency";else delete target[k];saveData(data);renderAdminEdit()},
    pendingChanges:u.pendingStampRequest.stamps,originalStamps:u.stamps});
} else {
  renderCalendar({mount:$("editCal"),monthCursor:editMonthCursor,stampedMap:u.stamps,clickable:true,onDayClick:d=>{const k=ymd(d);const cur=u.stamps[k];if(!cur)u.stamps[k]=true;else if(cur===true)u.stamps[k]="emergency";else delete u.stamps[k];saveData(data);renderAdminEdit()}});
}
// Approval bar
let reqBar=document.getElementById("adminStampReqBar");
if(!reqBar){reqBar=document.createElement("div");reqBar.id="adminStampReqBar";$("editCal").parentNode.appendChild(reqBar)}
reqBar.innerHTML="";
if(hasPending){
  reqBar.className="admin-request-info";
  reqBar.innerHTML=`<div class="ari-title">📨 スタンプ修正申請あり</div><div style="font-size:11px;color:var(--muted);margin-bottom:8px;">ハイライト部分が申請された変更です</div>`;
  const btnWrap=document.createElement("div");btnWrap.style.cssText="display:flex;gap:8px;";
  const approveBtn=document.createElement("button");approveBtn.className="btn success small";approveBtn.textContent="✅ 承認";
  approveBtn.addEventListener("click",()=>{
    u.stamps=JSON.parse(JSON.stringify(u.pendingStampRequest.stamps));
    u.pendingStampRequest={status:"approved",resolvedAt:Date.now()};saveData(data);
    showModal({title:"承認しました",sub:`${u.name||u.id}のスタンプを更新しました`,big:"✅"});renderAdminEdit()});
  const rejectBtn=document.createElement("button");rejectBtn.className="btn danger small";rejectBtn.textContent="❌ 却下";
  rejectBtn.addEventListener("click",()=>{
    u.pendingStampRequest={status:"rejected",resolvedAt:Date.now()};saveData(data);
    showModal({title:"却下しました",sub:`${u.name||u.id}の申請を却下しました`,big:"❌"});renderAdminEdit()});
  btnWrap.appendChild(approveBtn);btnWrap.appendChild(rejectBtn);reqBar.appendChild(btnWrap);
} else {reqBar.className="";}
}

/* === CALENDAR === */
function renderCalendar({mount,monthCursor,stampedMap,clickable,onDayClick,pendingChanges,originalStamps}){
  const monthKey = `${monthCursor.getFullYear()}-${monthCursor.getMonth()}`;
  const today = new Date();
  const state = mount._calendarState || {};

  function getStampState(key){
    const sv = stampedMap && stampedMap[key];
    const isEmg = sv === "emergency";
    const stamped = !!sv;
    const isPendingAdd = pendingChanges && originalStamps && pendingChanges[key] && !originalStamps[key];
    const isPendingRemove = pendingChanges && originalStamps && !pendingChanges[key] && originalStamps[key];
    const isPendingChange = pendingChanges && originalStamps && pendingChanges[key] && originalStamps[key] && pendingChanges[key] !== originalStamps[key];
    if (isPendingAdd) return { cls: "stamp pending", text: "＋" };
    if (isPendingRemove) return { cls: "stamp pending-remove", text: "−" };
    if (isPendingChange) return { cls: "stamp pending", text: pendingChanges[key] === "emergency" ? "⚡" : "✓" };
    if (isEmg) return { cls: "stamp emergency", text: "⚡" };
    return { cls: "stamp" + (stamped ? " on" : ""), text: stamped ? "✓" : "" };
  }

  function getPendingDayClass(key){
    const isPendingAdd = pendingChanges && originalStamps && pendingChanges[key] && !originalStamps[key];
    const isPendingRemove = pendingChanges && originalStamps && !pendingChanges[key] && originalStamps[key];
    const isPendingChange = pendingChanges && originalStamps && pendingChanges[key] && originalStamps[key] && pendingChanges[key] !== originalStamps[key];
    if (isPendingAdd) return " pending-add";
    if (isPendingRemove) return " pending-remove";
    if (isPendingChange) return " pending-change";
    return "";
  }

  const rebuild = state.monthKey !== monthKey || !state.cells;
  if (rebuild) {
    mount.innerHTML = "";
    const frag = document.createDocumentFragment();
    ["月","火","水","木","金","土","日"].forEach(w => {
      const d = document.createElement("div");
      d.className = "dow";
      d.textContent = w;
      frag.appendChild(d);
    });
    const first = startOfMonth(monthCursor);
    const firstDow = (first.getDay() + 6) % 7;
    const start = new Date(first);
    start.setDate(start.getDate() - firstDow);
    const cells = {};
    for (let i = 0; i < 42; i++) {
      const d = new Date(start);
      d.setDate(start.getDate() + i);
      const inM = d.getMonth() === monthCursor.getMonth();
      const key = ymd(d);
      const cell = document.createElement("div");
      cell.className = "day" + (inM ? "" : " muted") + (clickable ? " clickable" : "") + getPendingDayClass(key);
      const top = document.createElement("div");
      top.className = "n";
      const num = document.createElement("span");
      num.textContent = d.getDate();
      top.appendChild(num);
      if (isSameDay(d, today)) {
        const b = document.createElement("span");
        b.className = "badgeToday";
        b.textContent = "TODAY";
        top.appendChild(b);
      }
      const meta = document.createElement("div");
      meta.className = "dayMeta";
      meta.textContent = dowJa(d) + "曜";
      const st = document.createElement("div");
      const s = getStampState(key);
      st.className = s.cls;
      st.textContent = s.text;
      cell.appendChild(top);
      cell.appendChild(meta);
      cell.appendChild(st);
      if (clickable && onDayClick) cell.addEventListener("click", () => onDayClick(d));
      frag.appendChild(cell);
      cells[key] = { cell, stampEl: st };
    }
    mount.appendChild(frag);
    mount._calendarState = { monthKey, cells };
    return;
  }

  Object.keys(state.cells).forEach(key => {
    const ref = state.cells[key];
    const s = getStampState(key);
    if (ref.stampEl.className !== s.cls) ref.stampEl.className = s.cls;
    if (ref.stampEl.textContent !== s.text) ref.stampEl.textContent = s.text;
    ref.cell.classList.toggle("clickable", !!clickable);
    ref.cell.classList.toggle("pending-add", getPendingDayClass(key) === " pending-add");
    ref.cell.classList.toggle("pending-remove", getPendingDayClass(key) === " pending-remove");
    ref.cell.classList.toggle("pending-change", getPendingDayClass(key) === " pending-change");
    if (!clickable) ref.cell.onclick = null;
  });
}


// Build task filter selects
function buildTaskFilterSelects(ySel,mSel,staffSel,empSel){
  const now=new Date();
  ySel.innerHTML="<option value='全て'>全て</option>";
  for(let y=2024;y<=now.getFullYear()+1;y++){const o=document.createElement("option");o.value=y;o.textContent=y+"年";ySel.appendChild(o)}
  ySel.value=String(now.getFullYear());
  mSel.innerHTML="<option value='全て'>全て</option>";
  for(let m=1;m<=12;m++){const o=document.createElement("option");o.value=m;o.textContent=m+"月";mSel.appendChild(o)}
  mSel.value=String(now.getMonth()+1);
  if(staffSel){staffSel.innerHTML="<option value='全て'>全て</option>";getStaffNames().forEach(n=>{const o=document.createElement("option");o.value=n;o.textContent=n;staffSel.appendChild(o)})}
  if(empSel){empSel.innerHTML="<option value='全て'>全て</option>";getEmployees().forEach(n=>{const o=document.createElement("option");o.value=n;o.textContent=n;empSel.appendChild(o)})}
}

function filterTasks(dateType,y,m,staff,employee,status,subTabStaff,hideStaffs){
  return data.tasks.filter(t=>{
    if(subTabStaff&&t.staff!==subTabStaff)return false;
    // Hide specific staff from 全体 view
    if(!subTabStaff&&hideStaffs&&hideStaffs.length){
      for(const hs of hideStaffs){if(t.staff&&t.staff.includes(hs))return false}
    }
    const dateVal=t[dateType];
    if(y!=="全て"&&dateVal){const d=new Date(dateVal+"T00:00:00");if(d.getFullYear()!==parseInt(y))return false}
    if(m!=="全て"&&dateVal){const d=new Date(dateVal+"T00:00:00");if((d.getMonth()+1)!==parseInt(m))return false}
    if(staff!=="全て"&&t.staff!==staff)return false;
    if(employee&&employee!=="全て"&&t.employee!==employee)return false;
    if(status!=="全て"&&t.status!==status)return false;
    return true;
  });
}

function renderTaskTable(theadEl,tbodyEl,tasks,isAdmin){
  theadEl.innerHTML=`<tr><th>No</th><th>状況</th><th>形態</th><th>依頼日</th><th>期限日</th><th>完了日</th><th>工数</th><th>ﾃｷｽﾄｺｰﾄﾞ</th><th>業務種類</th><th>内容</th><th>担当社員</th><th>担当ｽﾀｯﾌ</th><th>備考</th><th>有効指摘</th><th>提出</th>${isAdmin?"<th>削除</th>":""}</tr>`;
  tbodyEl.innerHTML="";
  if(!tasks.length){tbodyEl.innerHTML=`<tr><td colspan="${isAdmin?16:15}" style="text-align:center;padding:20px;color:var(--muted);">データなし</td></tr>`;return}
  tasks.forEach(t=>{
    const tr=document.createElement("tr");
    const mkTd=(v)=>{const td=document.createElement("td");td.textContent=v;return td};
    // No
    tr.appendChild(mkTd(t.seqNum||"-"));
    // Status
    const tdSt=document.createElement("td");tdSt.innerHTML=`<span class="status-${t.status}">${t.status}</span>`;tr.appendChild(tdSt);
    // WorkType
    tr.appendChild(mkTd(t.workType));
    tr.appendChild(mkTd(t.requestDate||""));tr.appendChild(mkTd(t.deadline||""));tr.appendChild(mkTd(t.completionDate||""));
    tr.appendChild(mkTd(t.manHours||0));
    tr.appendChild(mkTd((t.textCodes||[]).join(", ")));
    // TaskType - read-only for staff
    tr.appendChild(mkTd(t.taskType||""));
    tr.appendChild(mkTd((t.content||"").substring(0,20)));tr.appendChild(mkTd(t.employee||""));tr.appendChild(mkTd(t.staff||""));
    // Notes - editable for staff
    if(!isAdmin){
      const tdN=document.createElement("td");const inp=document.createElement("input");inp.type="text";inp.value=t.notes||"";inp.style.cssText="font-size:10px;padding:2px 4px;border-radius:6px;width:80px;";
      inp.addEventListener("change",()=>{t.notes=inp.value;saveData(data)});tdN.appendChild(inp);tr.appendChild(tdN);
    } else tr.appendChild(mkTd(t.notes||""));
    // Valid point count
    if(isAdmin){
      const tdV=document.createElement("td");tdV.style.whiteSpace="nowrap";
      const vpCount=t.validPointCount||0;
      if(vpCount>0){const sp=document.createElement("span");sp.textContent=vpCount;sp.style.cssText="font-weight:700;margin-right:4px;";tdV.appendChild(sp);}
      const vpBtn=document.createElement("button");vpBtn.className="btn small ghost";vpBtn.textContent="編集";vpBtn.style.cssText="font-size:10px;padding:2px 6px;";
      vpBtn.addEventListener("click",()=>{
        const val=prompt("有効指摘回数を入力",String(t.validPointCount||0));if(val===null)return;
        const n=Math.max(0,parseInt(val)||0);
        const today=ymd(new Date());
        if(!t.vpEditHistory)t.vpEditHistory=[];
        const todayEdits=t.vpEditHistory.filter(e=>e.date===today).length;
        if(t.vpEditHistory.length>0){
          const lastDate=t.vpEditHistory[t.vpEditHistory.length-1].date;
          if(lastDate!==today||todayEdits>=2){if(!confirm("本当に再編集しますか？"))return;}}
        t.validPointCount=n;t.vpEditHistory.push({date:today,value:n});saveData(data);
        if(data.session.adminAuthed)renderAdminTaskList();else renderStaffTaskList();
      });
      tdV.appendChild(vpBtn);tr.appendChild(tdV);
    } else {
      tr.appendChild(mkTd(t.validPointCount||0));
    }
    // Submit / DL / status change column
    const tdSub=document.createElement("td");tdSub.style.whiteSpace="nowrap";
    if(isAdmin){
      // === ADMIN side ===
      if(t.status==="依頼前"){
        // 依頼前: show "依頼中に変更" button (opens file upload in admin-irai mode)
        const btn=document.createElement("button");btn.className="btn primary small";btn.textContent="📨 依頼中に変更";
        btn.addEventListener("click",()=>openFileUpload(t,"admin-irai"));tdSub.appendChild(btn);
      } else if(t.status==="依頼中"||t.status==="期限超過"){
        // 依頼中/期限超過: show DL button (if files exist) + file attach button
        const wrap=document.createElement("div");wrap.style.display="flex";wrap.style.alignItems="center";wrap.style.gap="4px";wrap.style.flexWrap="wrap";
        if(t.fileNames&&t.fileNames.length&&t.fileNames[0]!=="（ファイルなし）"){
          const dlBtn=document.createElement("button");dlBtn.className="btn primary small";dlBtn.textContent="📥 DL";
          dlBtn.addEventListener("click",()=>{
            if(t.fileIds && t.fileIds.length){
              for(let i=0;i<t.fileIds.length;i++){
                downloadDriveFile(t.fileIds[i], (t.fileNames && t.fileNames[i]) || "download");
              }
            }else if(t.fileNames && t.fileNames.length){
              downloadTaskFiles(t);
            }
          });
          wrap.appendChild(dlBtn);
        }
        const attBtn=document.createElement("button");attBtn.className="btn ghost small";attBtn.textContent="📎添付";
        attBtn.addEventListener("click",()=>openFileUpload(t,"admin-attach"));wrap.appendChild(attBtn);
        tdSub.appendChild(wrap);
      } else if(t.status==="完了"){
        // 完了: DL + 依頼中に戻す
        const wrap=document.createElement("div");wrap.style.display="flex";wrap.style.alignItems="center";wrap.style.gap="4px";wrap.style.flexWrap="wrap";
        const span=document.createElement("span");span.style.fontSize="10px";span.style.color="var(--mint)";span.textContent="✅ 完了";wrap.appendChild(span);
        if(t.fileNames&&t.fileNames.length&&t.fileNames[0]!=="（ファイルなし）"){
          const dlBtn=document.createElement("button");dlBtn.className="btn primary small";dlBtn.textContent="📥 DL";
          dlBtn.addEventListener("click",()=>{
            if(t.fileIds && t.fileIds.length){
              for(let i=0;i<t.fileIds.length;i++){
                downloadDriveFile(t.fileIds[i], (t.fileNames && t.fileNames[i]) || "download");
              }
            }else if(t.fileNames && t.fileNames.length){
              downloadTaskFiles(t);
            }
          });
          wrap.appendChild(dlBtn);
        }
        const revertBtn=document.createElement("button");revertBtn.className="btn danger small";revertBtn.textContent="↩ 依頼中に戻す";
        revertBtn.addEventListener("click",()=>{
          if(!confirm("依頼中に戻しますか？"))return;
          t.status="依頼中";t.completionDate="";saveData(data);renderAdminTaskList();
          showModal({title:"戻しました",sub:"ステータスを依頼中に戻しました。",big:"↩️"});
        });wrap.appendChild(revertBtn);
        tdSub.appendChild(wrap);
      } else if(t.status==="キャンセル"){
        tdSub.textContent="―";
      }
    } else {
      // === STAFF side ===
      if(t.status==="依頼中"||t.status==="期限超過"){
        const wrap=document.createElement("div");wrap.style.display="flex";wrap.style.alignItems="center";wrap.style.gap="4px";wrap.style.flexWrap="wrap";
        // DL button for staff (only if admin attached files)
        if(t.fileNames&&t.fileNames.length&&t.fileNames[0]!=="（ファイルなし）"){
          const dlBtn=document.createElement("button");dlBtn.className="btn primary small";dlBtn.textContent="📥 DL";
          dlBtn.addEventListener("click",()=>{
            downloadTaskFiles(t);
          });wrap.appendChild(dlBtn);
        }
        // Submit button
        const btn=document.createElement("button");btn.className="btn success small";btn.textContent="📎提出";
        btn.addEventListener("click",()=>openFileUpload(t,"staff"));wrap.appendChild(btn);
        tdSub.appendChild(wrap);
      } else if(t.status==="完了"){
        const wrap=document.createElement("div");wrap.style.display="flex";wrap.style.alignItems="center";wrap.style.gap="4px";wrap.style.flexWrap="wrap";
        const span=document.createElement("span");span.style.fontSize="10px";span.style.color="var(--mint)";
        const fnames=(t.fileNames&&t.fileNames.length)?t.fileNames.join(", "):"完了";
        span.textContent="✅ "+fnames;wrap.appendChild(span);
        const cancelBtn=document.createElement("button");cancelBtn.className="btn danger small";cancelBtn.textContent="取り消し";
        cancelBtn.addEventListener("click",()=>{
          if(!confirm("依頼中に戻しますか？"))return;
          t.status="依頼中";t.completionDate="";t.fileNames=[];saveData(data);renderStaffTaskList();
          showModal({title:"戻しました",sub:"ステータスを依頼中に戻しました。",big:"↩️"});
        });wrap.appendChild(cancelBtn);
        tdSub.appendChild(wrap);
      } else if(t.status==="依頼前"){
        tdSub.textContent="―";
      } else {
        if(t.fileNames&&t.fileNames.length){tdSub.innerHTML=`<span style="font-size:10px;color:var(--mint);">✅ ${escapeHtml(t.fileNames.join(", "))}</span>`}
      }
    }
    tr.appendChild(tdSub);
    // Admin: inline click-to-edit on cells
    if(isAdmin){
      tr.querySelectorAll("td").forEach((td,ci)=>{
        // Skip status(1), submit(13), valid point(12)
        if([0,1,12,13].includes(ci))return;
        td.style.cursor="pointer";
        td.addEventListener("dblclick",()=>{
          if(td.classList.contains("editing"))return;
          td.classList.add("editing");
          const fields=["seqNum","","workType","requestDate","deadline","completionDate","manHours","textCodes","taskType","content","employee","staff","notes"];
          const f=fields[ci];if(!f)return;
          let inp;
          if(f==="workType"){inp=document.createElement("select");["出勤","在宅"].forEach(v=>{const o=document.createElement("option");o.value=v;o.textContent=v;inp.appendChild(o)});inp.value=t[f]||"";}
          else if(f==="requestDate"||f==="deadline"||f==="completionDate"){inp=document.createElement("input");inp.type="date";inp.value=t[f]||"";}
          else if(f==="taskType"){inp=document.createElement("select");getTaskTypes().forEach(v=>{const o=document.createElement("option");o.value=v;o.textContent=v;inp.appendChild(o)});inp.value=t[f]||"";}
          else if(f==="employee"){inp=document.createElement("select");getEmployees().forEach(v=>{const o=document.createElement("option");o.value=v;o.textContent=v;inp.appendChild(o)});inp.value=t[f]||"";}
          else if(f==="staff"){inp=document.createElement("select");const _uo=document.createElement("option");_uo.value="未指定";_uo.textContent="未指定";inp.appendChild(_uo);getStaffNames().forEach(v=>{const o=document.createElement("option");o.value=v;o.textContent=v;inp.appendChild(o)});inp.value=t[f]||"未指定";}
          else if(f==="textCodes"){inp=document.createElement("input");inp.type="text";inp.value=(t.textCodes||[]).join(", ");}
          else{inp=document.createElement("input");inp.type=f==="manHours"||f==="seqNum"?"number":"text";inp.value=t[f]||"";}
          inp.style.cssText="font-size:11px;padding:2px 4px;border-radius:6px;width:100%;min-width:60px;";
          td.textContent="";td.appendChild(inp);inp.focus();
          const save=()=>{td.classList.remove("editing");
            if(f==="textCodes"){t.textCodes=inp.value.split(",").map(x=>x.trim()).filter(x=>x);}
            else if(f==="manHours"||f==="seqNum"){t[f]=parseInt(inp.value)||0;}
            else{t[f]=inp.value;}
            saveData(data);renderAdminTaskList();};
          inp.addEventListener("blur",save);inp.addEventListener("keydown",e=>{if(e.key==="Enter"){e.preventDefault();inp.blur()}});
        });
      });
      // Delete button
      const tdDel=document.createElement("td");
const bD=document.createElement("button");bD.className="btn danger small";bD.textContent="削除";
bD.addEventListener("click",()=>{if(!confirm("この業務を削除しますか？"))return;data.tasks=data.tasks.filter(x=>x.id!==t.id);saveData(data);renderAdminTaskList()});
      tdDel.appendChild(bD);tr.appendChild(tdDel);
    }
    tbodyEl.appendChild(tr);
  });
}

/* File Upload (Multiple files) */
let fileUploadTask=null;
let pendingFiles=[];
let fileUploadMode="staff"; // "staff"=提出(完了にする), "admin-attach"=ファイル添付のみ, "admin-irai"=依頼前→依頼中+ファイル添付
function renderFileList(){
  const fl=$("fileList");fl.innerHTML="";
  const allFiles=[...(fileUploadTask.fileNames||[]),...pendingFiles.map(f=>f.name)];
  allFiles.forEach((name,i)=>{
    const item=document.createElement("div");item.className="file-item";
    item.innerHTML=`<span>📎 ${escapeHtml(name)}</span>`;
    const rm=document.createElement("span");rm.className="file-remove";rm.textContent="✕";
    rm.addEventListener("click",()=>{
      if(!confirm("このファイルを削除しますか？"))return;
      const existCount=(fileUploadTask.fileNames||[]).length;
      if(i<existCount){fileUploadTask.fileNames.splice(i,1);if(fileUploadTask.fileIds&&fileUploadTask.fileIds.length>i)fileUploadTask.fileIds.splice(i,1);saveData(data)}
      else{pendingFiles.splice(i-existCount,1)}
      renderFileList();
    });
    item.appendChild(rm);fl.appendChild(item);
  });
  $("fileSubmitBtn").disabled=allFiles.length===0;
  $("dzFileName").textContent=allFiles.length>0?`${allFiles.length}件のファイル`:"";
}
function openFileUpload(task,mode){fileUploadMode=mode||"staff";fileUploadTask=task;pendingFiles=[];renderFileList();
$("fileOverlay").style.display="flex";$("fileInput").value="";
// Adjust button labels based on mode
if(fileUploadMode==="admin-attach"){$("fileSubmitBtn").textContent="ファイル添付 ✅";$("fileSubmitDirectBtn").textContent="添付せず閉じる"}
else if(fileUploadMode==="admin-irai"){$("fileSubmitBtn").textContent="依頼中に変更 ✅";$("fileSubmitDirectBtn").textContent="ファイルなしで依頼中に変更"}
else{$("fileSubmitBtn").textContent="完了 ✅";$("fileSubmitDirectBtn").textContent="提出（ファイルなしでも完了）"}}
$("fileOverlayClose").addEventListener("click",()=>{$("fileOverlay").style.display="none"});
const dz=$("dropZone");
dz.addEventListener("dragover",e=>{e.preventDefault();dz.classList.add("over")});
dz.addEventListener("dragleave",()=>dz.classList.remove("over"));
dz.addEventListener("drop",e=>{e.preventDefault();dz.classList.remove("over");const files=e.dataTransfer.files;if(files.length)handleFiles(files)});
dz.addEventListener("click",()=>$("fileInput").click());
$("fileInput").addEventListener("change",e=>{if(e.target.files.length)handleFiles(e.target.files)});
function handleFiles(files){if(!fileUploadTask)return;for(let i=0;i<files.length;i++){pendingFiles.push(files[i])}renderFileList()}
$("fileSubmitBtn").addEventListener("click",async ()=>{if(!fileUploadTask)return;
  // Upload files to Drive if API is configured
  for(var i=0;i<pendingFiles.length;i++){
    var f=pendingFiles[i];
    if(!fileUploadTask.fileNames)fileUploadTask.fileNames=[];
    if(API_URL){
      var result=await uploadFileToDrive(f, fileUploadTask.id);
      if(result){fileUploadTask.fileNames.push(result.fileName);if(!fileUploadTask.fileIds)fileUploadTask.fileIds=[];fileUploadTask.fileIds.push(result.fileId);}
      else{fileUploadTask.fileNames.push(f.name+"(アップロード失敗)");}
    }else{fileUploadTask.fileNames.push(f.name);}
  }
  pendingFiles=[];
  if(fileUploadMode==="admin-attach"){
    // Just attach files, don't change status
    saveData(data);$("fileOverlay").style.display="none";
    const names=(fileUploadTask.fileNames||[]).join(", ");fileUploadTask=null;
    showModalCb({title:"ファイル添付完了",sub:names,big:"📎"},()=>renderAdminTaskList());
  } else if(fileUploadMode==="admin-irai"){
    // Change 依頼前→依頼中 + attach files
    fileUploadTask.status="依頼中";saveData(data);$("fileOverlay").style.display="none";
    const names=(fileUploadTask.fileNames||[]).join(", ");fileUploadTask=null;
    showModalCb({title:"依頼中に変更しました",sub:names||"ファイルなし",big:"📨"},()=>renderAdminTaskList());
  } else {
    fileUploadTask.status="完了";fileUploadTask.completionDate=ymd(new Date());saveData(data);$("fileOverlay").style.display="none";
    const names=(fileUploadTask.fileNames||[]).join(", ");fileUploadTask=null;
    showModalCb({title:"提出完了！",sub:names,big:"✅"},()=>{
      if(data.session.adminAuthed)renderAdminTaskList();else renderStaffTaskList()});
  }});

$("fileSubmitDirectBtn").addEventListener("click",async ()=>{
  if(!fileUploadTask)return;
  const baseTask = data.tasks.find(x=>x.id===fileUploadTask.id) || fileUploadTask;
  for(var i=0;i<pendingFiles.length;i++){
    var f=pendingFiles[i];
    if(!baseTask.fileNames)baseTask.fileNames=[];
    if(API_URL){
      var result=await uploadFileToDrive(f, baseTask.id);
      if(result){baseTask.fileNames.push(result.fileName);if(!baseTask.fileIds)baseTask.fileIds=[];baseTask.fileIds.push(result.fileId);}
      else{baseTask.fileNames.push(f.name);}
    }else{baseTask.fileNames.push(f.name);}
  }
  pendingFiles=[];
  if(fileUploadMode==="admin-attach"){
    // Close without doing anything extra
    saveData(data);$("fileOverlay").style.display="none";fileUploadTask=null;
    renderAdminTaskList();
  } else if(fileUploadMode==="admin-irai"){
    baseTask.status="依頼中";saveData(data);$("fileOverlay").style.display="none";fileUploadTask=null;
    showModalCb({title:"依頼中に変更しました",sub:"ファイルなし",big:"📨"},()=>renderAdminTaskList());
  } else {
    baseTask.status="完了";baseTask.completionDate=ymd(new Date());
    if(!baseTask.fileNames||baseTask.fileNames.length===0) baseTask.fileNames=["（ファイルなし）"];
    saveData(data);$("fileOverlay").style.display="none";fileUploadTask=null;
    const names=(baseTask.fileNames||[]).join(", ");
    showModalCb({title:"提出完了！",sub:names,big:"✅"},()=>{
      if(data.session.adminAuthed) renderAdminTaskList(); else renderStaffTaskList()});
  }
});

/* Task Add/Edit Overlay */
let editingTaskId=null;
function populateStaffSelect(sel, selectedValue){
  sel.innerHTML="";
  const unspecified=document.createElement("option");unspecified.value="未指定";unspecified.textContent="未指定";sel.appendChild(unspecified);
  getStaffNames().forEach(n=>{const o=document.createElement("option");o.value=n;o.textContent=n;sel.appendChild(o)});
  sel.value=selectedValue||"未指定";
}
function isStaffShakaijin(staffName){
  const u=Object.values(data.users).find(x=>(x.name||x.id)===staffName);
  return u&&u.userType==="社会人";
}
function applyTaskTypeLogic(){
  const wt=$("taWorkType").value;
  const staffName=$("taStaff").value;
  const isShakaijin=isStaffShakaijin(staffName);
  // 出勤 or 社会人 → force 時給 and disable
  if(wt==="出勤"||isShakaijin){
    $("taTaskType").value="時給";
    if(!getTaskTypes().includes("時給"))$("taTaskType").value="その他（時給）";
    $("taTaskType").disabled=true;
  } else {
    $("taTaskType").disabled=false;
  }
}
function openTaskAdd(){editingTaskId=null;$("taskAddTitle").textContent="📑 新規業務追加";
const now=new Date();const week=addDays(now,7);
$("taWorkType").value="出勤";$("taStatus").value="依頼前";$("taRequestDate").value=ymd(now);$("taDeadline").value=ymd(week);$("taCompletionDate").value="";$("taManHours").value="1";$("taContent").value="";$("taNote").value="";
popSel($("taTaskType"),getTaskTypes());popSel($("taEmployee"),getEmployees());
populateStaffSelect($("taStaff"),"未指定");
applyTaskTypeLogic();
renderTextCodeInputs([""]);$("taskAddOverlay").style.display="flex"}

$("taStaff").addEventListener("change",()=>{applyTaskTypeLogic()});
$("taWorkType").addEventListener("change",()=>{applyTaskTypeLogic()});
$("taTaskType").addEventListener("change",()=>{
  const staffName=$("taStaff").value;
  const wt=$("taWorkType").value;
  const tt=$("taTaskType").value;
  if((wt==="出勤"||isStaffShakaijin(staffName))&&tt!=="時給"&&tt!=="その他（時給）"){
    $("taTaskType").value="時給";
    if(!getTaskTypes().includes("時給"))$("taTaskType").value="その他（時給）";
  }
});

function openTaskEdit(t){editingTaskId=t.id;$("taskAddTitle").textContent="📑 業務編集";
$("taWorkType").value=t.workType;$("taStatus").value=t.status;$("taRequestDate").value=t.requestDate||"";$("taDeadline").value=t.deadline||"";$("taCompletionDate").value=t.completionDate||"";$("taManHours").value=t.manHours||1;$("taContent").value=t.content||"";$("taNote").value=t.notes||"";
popSel($("taTaskType"),getTaskTypes(),t.taskType);popSel($("taEmployee"),getEmployees(),t.employee);
populateStaffSelect($("taStaff"),t.staff||"未指定");
applyTaskTypeLogic();
renderTextCodeInputs(t.textCodes&&t.textCodes.length?t.textCodes:[""]);$("taskAddOverlay").style.display="flex"}
$("taskAddClose").addEventListener("click",()=>{$("taskAddOverlay").style.display="none"});
function renderTextCodeInputs(codes){
  const area=$("taTextCodesArea");area.innerHTML="";
  codes.forEach((c,i)=>{
    const row=document.createElement("div");row.className="tc-row";
    const inp=document.createElement("input");inp.type="text";inp.inputMode="numeric";inp.value=c;inp.placeholder="テキストコード";inp.dataset.idx=i;
    inp.addEventListener("input",function(){this.value=this.value.replace(/[^0-9]/g,"")});
    row.appendChild(inp);
    if(i===codes.length-1){const ab=document.createElement("button");ab.className="btn small primary";ab.textContent="+";ab.type="button";
      ab.addEventListener("click",()=>{const cur=getTextCodes();cur.push("");renderTextCodeInputs(cur)});row.appendChild(ab)}
    if(codes.length>1){const rb=document.createElement("button");rb.className="btn small danger";rb.textContent="×";rb.type="button";
      rb.addEventListener("click",()=>{const cur=getTextCodes();cur.splice(i,1);renderTextCodeInputs(cur)});row.appendChild(rb)}
    area.appendChild(row);
  });
}
function getTextCodes(){return Array.from($("taTextCodesArea").querySelectorAll("input")).map(i=>i.value)}
$("taskAddSave").addEventListener("click",()=>{
  const wt=$("taWorkType").value;const tc=getTextCodes().filter(x=>x);
  if(editingTaskId){
    const t=data.tasks.find(x=>x.id===editingTaskId);if(!t)return;
    t.workType=wt;t.status=$("taStatus").value;t.requestDate=$("taRequestDate").value;t.deadline=$("taDeadline").value;t.completionDate=$("taCompletionDate").value;
    t.manHours=parseInt($("taManHours").value)||1;t.textCodes=tc;t.taskType=$("taTaskType").value;t.content=$("taContent").value;t.employee=$("taEmployee").value;t.staff=$("taStaff").value;t.notes=$("taNote").value;
  } else {
    data.tasks.push({id:Date.now(),seqNum:nextSeqNum(wt),workType:wt,status:$("taStatus").value,
      requestDate:$("taRequestDate").value,deadline:$("taDeadline").value,completionDate:$("taCompletionDate").value,
      manHours:parseInt($("taManHours").value)||1,textCodes:tc,taskType:$("taTaskType").value,content:$("taContent").value,
      employee:$("taEmployee").value,staff:$("taStaff").value,notes:$("taNote").value,validPointCount:0,fileNames:[]});
  }
  saveData(data);$("taskAddOverlay").style.display="none";
  // If admin and status is 依頼中, open file upload for attaching files
  if(data.session.adminAuthed&&$("taStatus").value==="依頼中"){
    const newT=editingTaskId?data.tasks.find(x=>x.id===editingTaskId):data.tasks[data.tasks.length-1];
    if(newT){renderAdminTaskList();openFileUpload(newT,"admin-attach");return}
  }
  if(data.session.adminAuthed)renderAdminTaskList();else renderStaffTaskList()});
$("atlAddTask").addEventListener("click",openTaskAdd);

function getTaskImportHeaders() {
  return [
    { key: "workType", label: "業務形態(workType)" },
    { key: "status", label: "状態(status)" },
    { key: "requestDate", label: "依頼日(requestDate)" },
    { key: "deadline", label: "期限(deadline)" },
    { key: "completionDate", label: "完了日(completionDate)" },
    { key: "manHours", label: "工数(manHours)" },
    { key: "taskType", label: "業務種類(taskType)" },
    { key: "employee", label: "担当社員(employee)" },
    { key: "staff", label: "担当スタッフ(staff)" },
    { key: "textCodes", label: "テキストコード(textCodes)" },
    { key: "content", label: "内容(content)" },
    { key: "notes", label: "備考(notes)" }
  ];
}

function getTaskImportDefaults() {
  const workTypeSel = $("taWorkType");
  const statusSel = $("taStatus");
  return {
    workType: workTypeSel && workTypeSel.value ? workTypeSel.value : "制作",
    status: statusSel && statusSel.value ? statusSel.value : "依頼中",
    requestDate: ymd(new Date()),
    deadline: ymd(addDays(new Date(), 7)),
    completionDate: "",
    manHours: "1",
    taskType: (getTaskTypes()[0] || ""),
    employee: (getEmployees()[0] || ""),
    staff: "未指定",
    textCodes: "",
    content: "内容を入力",
    notes: ""
  };
}

function escapeTaskImportHtml(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;");
}

function buildTaskImportTemplateHtml() {
  const headers = getTaskImportHeaders();
  const sample = getTaskImportDefaults();
  const statusValues = Array.from(($("taStatus") && $("taStatus").options) || []).map(o => o.value).filter(Boolean);
  const workTypeValues = Array.from(($("taWorkType") && $("taWorkType").options) || []).map(o => o.value).filter(Boolean);
  const employeeValues = getEmployees();
  const staffValues = ["未指定"].concat(getStaffUsers().map(getUserDisplayName));

  return `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>業務追加テンプレート</title>
</head>
<body>
<table border="1" data-template="task-import">
  <thead>
    <tr>${headers.map(h => `<th>${escapeTaskImportHtml(h.label)}</th>`).join("")}</tr>
  </thead>
  <tbody>
    <tr>${headers.map(h => `<td>${escapeTaskImportHtml(sample[h.key] || "")}</td>`).join("")}</tr>
  </tbody>
</table>
<br>
<table border="1">
  <tbody>
    <tr><th>使い方</th><td>1行につき1件です。不要な行は削除してください。</td></tr>
    <tr><th>業務形態</th><td>${escapeTaskImportHtml(workTypeValues.join(" / "))}</td></tr>
    <tr><th>状態</th><td>${escapeTaskImportHtml(statusValues.join(" / "))}</td></tr>
    <tr><th>担当社員</th><td>${escapeTaskImportHtml(employeeValues.join(" / "))}</td></tr>
    <tr><th>担当スタッフ</th><td>${escapeTaskImportHtml(staffValues.join(" / "))}</td></tr>
    <tr><th>テキストコード</th><td>複数ある場合はカンマ区切りで入力してください</td></tr>
  </tbody>
</table>
</body>
</html>`;
}

function downloadTaskImportTemplate() {
  const html = buildTaskImportTemplateHtml();
  const blob = new Blob([html], { type: "application/vnd.ms-excel;charset=utf-8;" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "業務追加テンプレート.xls";
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}

function normalizeTaskImportHeader(header) {
  const raw = String(header || "").trim();
  const match = raw.match(/\(([^)]+)\)/);
  return (match ? match[1] : raw).trim();
}

function parseDelimitedTaskImport(text) {
  const normalized = String(text || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const lines = normalized.split("\n").filter(line => line.trim() !== "");
  if (!lines.length) return [];
  const separator = lines[0].includes("\t") ? "\t" : ",";
  return lines.map(line => {
    const row = [];
    let current = "";
    let inQuotes = false;
    for (let i = 0; i < line.length; i += 1) {
      const ch = line[i];
      if (ch === "\"") {
        if (inQuotes && line[i + 1] === "\"") {
          current += "\"";
          i += 1;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (ch === separator && !inQuotes) {
        row.push(current);
        current = "";
      } else {
        current += ch;
      }
    }
    row.push(current);
    return row.map(cell => cell.trim());
  });
}

function decodeTaskImportBytes(bytes) {
  return new TextDecoder("utf-8").decode(bytes);
}

function colLabelToIndex(label) {
  let index = 0;
  const text = String(label || "").toUpperCase();
  for (let i = 0; i < text.length; i += 1) index = index * 26 + (text.charCodeAt(i) - 64);
  return Math.max(0, index - 1);
}

function parseTaskImportSharedStringsXml(xmlText) {
  const doc = new DOMParser().parseFromString(String(xmlText || ""), "application/xml");
  return Array.from(doc.getElementsByTagName("si")).map(si =>
    Array.from(si.getElementsByTagName("t")).map(node => node.textContent || "").join("")
  );
}

function parseTaskImportSheetXml(xmlText, sharedStrings) {
  const doc = new DOMParser().parseFromString(String(xmlText || ""), "application/xml");
  return Array.from(doc.getElementsByTagName("row")).map(row => {
    const cells = [];
    Array.from(row.getElementsByTagName("c")).forEach(cell => {
      const ref = cell.getAttribute("r") || "";
      const colRef = ref.replace(/[0-9]/g, "");
      const idx = colLabelToIndex(colRef);
      const type = cell.getAttribute("t") || "";
      let value = "";
      if (type === "s") {
        const v = cell.getElementsByTagName("v")[0];
        value = sharedStrings[parseInt(v && v.textContent, 10) || 0] || "";
      } else if (type === "inlineStr") {
        value = Array.from(cell.getElementsByTagName("t")).map(node => node.textContent || "").join("");
      } else {
        const v = cell.getElementsByTagName("v")[0];
        value = v ? String(v.textContent || "") : "";
      }
      cells[idx] = String(value || "").trim();
    });
    return cells;
  });
}

async function inflateTaskImportZipEntry(method, bytes) {
  if (method === 0) return new Uint8Array(bytes);
  if (method === 8) {
    if (typeof DecompressionStream === "undefined") throw new Error("xlsx inflate unsupported");
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }
  throw new Error(`unsupported zip method: ${method}`);
}

async function readTaskImportZipEntries(arrayBuffer) {
  const view = new DataView(arrayBuffer);
  let eocdOffset = -1;
  for (let pos = arrayBuffer.byteLength - 22; pos >= 0; pos -= 1) {
    if (view.getUint32(pos, true) === 0x06054b50) {
      eocdOffset = pos;
      break;
    }
  }
  if (eocdOffset < 0) throw new Error("zip eocd not found");

  const centralDirectorySize = view.getUint32(eocdOffset + 12, true);
  const centralDirectoryOffset = view.getUint32(eocdOffset + 16, true);
  const decoder = new TextDecoder("utf-8");
  const entries = {};
  let pos = centralDirectoryOffset;
  const end = centralDirectoryOffset + centralDirectorySize;

  while (pos < end) {
    if (view.getUint32(pos, true) !== 0x02014b50) throw new Error("invalid central directory");
    const compressionMethod = view.getUint16(pos + 10, true);
    const compressedSize = view.getUint32(pos + 20, true);
    const fileNameLength = view.getUint16(pos + 28, true);
    const extraLength = view.getUint16(pos + 30, true);
    const commentLength = view.getUint16(pos + 32, true);
    const localHeaderOffset = view.getUint32(pos + 42, true);
    const fileNameBytes = new Uint8Array(arrayBuffer, pos + 46, fileNameLength);
    const fileName = decoder.decode(fileNameBytes);
    pos += 46 + fileNameLength + extraLength + commentLength;

    if (view.getUint32(localHeaderOffset, true) !== 0x04034b50) throw new Error("invalid local header");
    const localNameLength = view.getUint16(localHeaderOffset + 26, true);
    const localExtraLength = view.getUint16(localHeaderOffset + 28, true);
    const dataStart = localHeaderOffset + 30 + localNameLength + localExtraLength;
    const compressed = new Uint8Array(arrayBuffer.slice(dataStart, dataStart + compressedSize));
    entries[fileName] = await inflateTaskImportZipEntry(compressionMethod, compressed);
  }

  return entries;
}

async function parseTaskImportRowsFromXlsx(file) {
  const entries = await readTaskImportZipEntries(await file.arrayBuffer());
  const sharedStrings = entries["xl/sharedStrings.xml"]
    ? parseTaskImportSharedStringsXml(decodeTaskImportBytes(entries["xl/sharedStrings.xml"]))
    : [];
  const sheetPath = Object.keys(entries).filter(name => /^xl\/worksheets\/sheet\d+\.xml$/i.test(name)).sort()[0];
  if (!sheetPath) return [];
  return parseTaskImportSheetXml(decodeTaskImportBytes(entries[sheetPath]), sharedStrings);
}

function normalizeTaskImportValue(key, value) {
  const raw = String(value == null ? "" : value).trim();
  if (!raw) return "";
  if ((key === "requestDate" || key === "deadline" || key === "completionDate") && /^\d+(\.\d+)?$/.test(raw)) {
    const serial = Number(raw);
    if (isFinite(serial) && serial > 0) {
      const utc = Date.UTC(1899, 11, 30) + Math.round(serial * 86400000);
      const date = new Date(utc);
      const y = date.getUTCFullYear();
      const m = pad2(date.getUTCMonth() + 1);
      const d = pad2(date.getUTCDate());
      return `${y}-${m}-${d}`;
    }
  }
  return raw;
}

function parseTaskImportTableRows(raw) {
  const doc = new DOMParser().parseFromString(String(raw || ""), "text/html");
  const table = doc.querySelector('table[data-template="task-import"]') || doc.querySelector("table");
  if (!table) return [];
  return Array.from(table.querySelectorAll("tr")).map(tr => Array.from(tr.children).map(cell => String(cell.textContent || "").trim()));
}

function extractTaskImportSheetRefs(raw) {
  const refs = [];
  const text = String(raw || "");
  const patterns = [
    /<x:WorksheetSource[^>]+HRef="([^"]+)"/gi,
    /<frame[^>]+src="([^"]+sheet[^"]*\.htm[^"]*)"/gi,
    /<link[^>]+id="?shLink"?[^>]+href="([^"]+)"/gi
  ];
  patterns.forEach(pattern => {
    let match;
    while ((match = pattern.exec(text))) {
      const href = String(match[1] || "").trim();
      if (href && refs.indexOf(href) < 0) refs.push(href);
    }
  });
  return refs;
}

async function fetchTaskImportSheetRows(sheetRef) {
  const url = new URL(sheetRef, window.location.href.split("#")[0]);
  const response = await fetch(url.toString(), { cache: "no-store" });
  if (!response.ok) throw new Error(`sheet fetch failed: ${response.status}`);
  const html = await response.text();
  return parseTaskImportTableRows(html);
}

async function parseTaskImportRows(text) {
  const raw = String(text || "");
  if (/<table/i.test(raw)) {
    const rows = parseTaskImportTableRows(raw);
    if (rows.length) return rows;
  }
  const sheetRefs = extractTaskImportSheetRefs(raw);
  for (let i = 0; i < sheetRefs.length; i += 1) {
    try {
      const rows = await fetchTaskImportSheetRows(sheetRefs[i]);
      if (rows.length) return rows;
    } catch (_error) {}
  }
  return parseDelimitedTaskImport(raw);
}

function convertImportRowsToTasks(rows) {
  if (!rows.length) return [];
  const headerMap = rows[0].map(normalizeTaskImportHeader);
  const validHeaders = getTaskImportHeaders().map(header => header.key);
  const matchedHeaderCount = headerMap.filter(key => validHeaders.indexOf(key) >= 0).length;
  if (matchedHeaderCount < 2) return [];
  const tasks = [];
  const defaults = getTaskImportDefaults();

  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
    const record = {};
    headerMap.forEach((key, idx) => { record[key] = normalizeTaskImportValue(key, (row && row[idx]) || ""); });
    const hasAnyValue = Object.values(record).some(value => String(value || "").trim() !== "");
    if (!hasAnyValue) continue;

    const workType = record.workType || defaults.workType;
    const task = {
      id: Date.now() + i,
      seqNum: nextSeqNum(workType),
      workType,
      status: record.status || defaults.status,
      requestDate: record.requestDate || defaults.requestDate,
      deadline: record.deadline || defaults.deadline,
      completionDate: record.completionDate || "",
      manHours: Math.max(1, parseInt(record.manHours, 10) || 1),
      textCodes: (record.textCodes || "").split(",").map(v => v.trim()).filter(Boolean),
      taskType: record.taskType || defaults.taskType,
      content: record.content || "",
      employee: record.employee || defaults.employee,
      staff: defaults.staff,
      notes: record.notes || "",
      validPointCount: 0,
      fileNames: [],
      fileIds: []
    };
    setTaskStaffRef(task, record.staff || defaults.staff);
    tasks.push(task);
    data.tasks.push(task);
  }
  return tasks;
}

async function importTasksFromFile(file) {
  if (!file) return;
  const isXlsx = /\.xlsx$/i.test(String(file.name || ""));
  const text = isXlsx ? "" : await file.text();
  const rows = isXlsx ? await parseTaskImportRowsFromXlsx(file) : await parseTaskImportRows(text);
  const imported = convertImportRowsToTasks(rows);
  if (!imported.length) {
    showModal({ title: "取込対象がありません", sub: "テンプレートの1行目は見出しです", big: "NG" });
    return;
  }
  saveData(data);
  renderAdminTaskList();
  showModal({ title: "Excel取込完了", sub: `${imported.length}件追加しました`, big: "OK" });
}

function ensureTaskImportControls() {
  const addBtn = $("atlAddTask");
  if (!addBtn || !addBtn.parentNode) return;
  if (!document.getElementById("atlDownloadTemplate")) {
    const templateBtn = document.createElement("button");
    templateBtn.id = "atlDownloadTemplate";
    templateBtn.className = "btn ghost small";
    templateBtn.textContent = "テンプレート";
    templateBtn.addEventListener("click", () => {
      downloadTaskImportTemplate().catch(error => handleDirectActionError(error, "テンプレート取得に失敗しました"));
    });
    addBtn.parentNode.insertBefore(templateBtn, addBtn);
  }
  if (!document.getElementById("atlImportExcel")) {
    const importBtn = document.createElement("button");
    importBtn.id = "atlImportExcel";
    importBtn.className = "btn small";
    importBtn.textContent = "Excel取込";
    importBtn.addEventListener("click", () => {
      const input = document.getElementById("atlImportFile");
      if (input) {
        input.value = "";
        input.click();
      }
    });
    addBtn.parentNode.insertBefore(importBtn, addBtn);
  }
  if (!document.getElementById("atlImportFile")) {
    const input = document.createElement("input");
    input.type = "file";
    input.id = "atlImportFile";
    input.accept = ".xlsx,.xls,.htm,.html,.csv,.tsv,.txt";
    input.style.display = "none";
    input.addEventListener("change", event => {
      const file = event.target && event.target.files && event.target.files[0];
      importTasksFromFile(file).catch(error => handleDirectActionError(error, "Excel取込に失敗しました"));
    });
    addBtn.parentNode.appendChild(input);
  }
}

importTasksFromFile = async function(file) {
  if (!file) return;
  const text = await file.text();
  const rows = await parseTaskImportRows(text);
  const imported = convertImportRowsToTasks(rows);
  if (!imported.length) {
    const looksLikeExcelFrameset = /WorksheetSource|ExcelWorkbook|sheet001\.htm|File-List/i.test(text);
    const sub = looksLikeExcelFrameset
      ? "この .xls は分割保存形式です。テンプレートを再ダウンロードするか、.files 内の sheet001.htm を取り込んでください。"
      : "テンプレートの1行目は見出し、2行目以降にデータを入れてください。";
    showModal({ title: "Excel取込できません", sub, big: "NG" });
    return;
  }
  saveData(data);
  renderAdminTaskList();
  showModal({ title: "Excel取込完了", sub: `${imported.length}件追加しました`, big: "OK" });
};

downloadTaskImportTemplate = async function() {
  const response = await fetch("./業務追加テンプレート.xlsx", { cache: "no-store" });
  if (!response.ok) throw new Error("template fetch failed");
  const blob = await response.blob();
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "業務追加テンプレート.xlsx";
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
};

importTasksFromFile = async function(file) {
  if (!file) return;
  const isXlsx = /\.xlsx$/i.test(String(file.name || ""));
  const text = isXlsx ? "" : await file.text();
  const rows = isXlsx ? await parseTaskImportRowsFromXlsx(file) : await parseTaskImportRows(text);
  const imported = convertImportRowsToTasks(rows);
  if (!imported.length) {
    const looksLikeExcelFrameset = !isXlsx && /WorksheetSource|ExcelWorkbook|sheet001\.htm|File-List/i.test(text);
    const sub = looksLikeExcelFrameset
      ? "この .xls は分割保存形式です。.xlsx テンプレートを使うか、.files 内の sheet001.htm を取り込んでください。"
      : "テンプレートの1行目は見出しです。2行目以降にデータを入れてください。";
    showModal({ title: "Excel取込できません", sub, big: "NG" });
    return;
  }
  saveData(data);
  renderAdminTaskList();
  showModal({ title: "Excel取込完了", sub: `${imported.length}件追加しました`, big: "OK" });
};

/* === STAFF TASK LIST === */
let stlInit=false;
let stlPriceListOpen=false;
function renderPriceList(container){
  container.innerHTML="";
  const panel=document.createElement("div");panel.className="price-list-panel";
  panel.innerHTML=`<div class="pl-title">💰 業務種類 単価一覧</div>`;
  const grid=document.createElement("div");grid.className="price-list-grid";
  getTaskTypes().forEach(t=>{
    const price=getTaskPrice(t);
    const priceStr=t==="時給"||t==="その他（時給）"?"時給":(price!=null?price.toLocaleString()+"円":"―");
    const n=document.createElement("div");n.className="pl-name";n.textContent=t;
    const p=document.createElement("div");p.className="pl-price";p.textContent=priceStr;
    grid.appendChild(n);grid.appendChild(p);
  });
  panel.appendChild(grid);container.appendChild(panel);
}
$("stlPriceListToggle").addEventListener("click",()=>{
  stlPriceListOpen=!stlPriceListOpen;
  const area=$("stlPriceListArea");
  if(stlPriceListOpen){area.classList.remove("hidden");renderPriceList(area);$("stlPriceListToggle").textContent="💰 単価一覧を閉じる"}
  else{area.classList.add("hidden");$("stlPriceListToggle").textContent="💰 業務単価一覧"}
});
function renderStaffTaskList(){
  const u=data.users[data.session.userId];if(!u)return;
  $("stlUserName").textContent=u.name||u.id;
  // Show/hide nav buttons based on user type
  if(u.userType==="社会人"){$("stlBackToReport").style.display="";$("stlBack").style.display="none";}
  else{$("stlBackToReport").style.display="none";$("stlBack").style.display="";}
  // Hide price list for 社会人
  if(u.userType==="社会人"){$("stlPriceListToggle").classList.add("hidden");$("stlPriceListArea").classList.add("hidden")}
  else{$("stlPriceListToggle").classList.remove("hidden")}
  renderWorkload($("stlWorkloadArea"),u.name||u.id);
  if(!stlInit){buildTaskFilterSelects($("stlYear"),$("stlMonth"),$("stlStaff"),null);$("stlStatus").value="全て";stlInit=true}
  // Always default to logged-in user
  $("stlStaff").value=u.name||u.id;
  doRenderSTL();
}
function doRenderSTL(){
  const tasks=filterTasks($("stlDateType").value,$("stlYear").value,$("stlMonth").value,$("stlStaff").value,null,$("stlStatus").value,null);
  renderTaskTable($("stlThead"),$("stlTbody"),tasks,false);
}
["stlDateType","stlYear","stlMonth","stlStaff","stlStatus"].forEach(id=>$(id).addEventListener("change",doRenderSTL));

/* === ADMIN TASK LIST === */
let atlInit=false,atlSubTab="全体";
function renderAdminTaskList(){
  renderAdminNotifications();
  renderWorkload($("atlWorkloadArea"),null);
  if(!atlInit){buildTaskFilterSelects($("atlYear"),$("atlMonth"),$("atlStaff"),$("atlEmployee"));$("atlStatus").value="全て";atlInit=true}
  // Sub tabs
  const subTabs=$("atlSubTabs");subTabs.innerHTML="";
  const tabs=["全体","小笠原さん業務","諸富さん業務"];
  tabs.forEach(t=>{const btn=document.createElement("div");btn.className="sub-tab"+(atlSubTab===t?" active":"");btn.textContent=t;
    btn.addEventListener("click",()=>{atlSubTab=t;renderAdminTaskList()});subTabs.appendChild(btn)});
  doRenderATL();
}
function doRenderATL(){
  let subStaff=null;
  if(atlSubTab==="小笠原さん業務")subStaff="小笠原";
  else if(atlSubTab==="諸富さん業務")subStaff="諸富";
  // partial match
  if(subStaff){
    const matchName=getStaffNames().find(n=>n.includes(subStaff));
    if(matchName)subStaff=matchName;
  }
  // Hide 小笠原/諸富 from 全体 tab
  const hideStaffs=(!subStaff&&atlSubTab==="全体")?["小笠原","諸富"]:null;
  const tasks=filterTasks($("atlDateType").value,$("atlYear").value,$("atlMonth").value,$("atlStaff").value,$("atlEmployee").value,$("atlStatus").value,subStaff,hideStaffs);
  renderTaskTable($("atlThead"),$("atlTbody"),tasks,true);
}
["atlDateType","atlYear","atlMonth","atlStaff","atlEmployee","atlStatus"].forEach(id=>$(id).addEventListener("change",doRenderATL));

/* === DROPDOWN EDIT === */
let ddEditIdx=-1;
let ddEditType=""; // "taskType" or ""
function renderDropdownEdit(){
  renderAdminNotifications();
  renderAdminCreds();
  if (DEFAULT_BOOTSTRAP_STAFF.some(staff => !((data.users || {})[staff.id]))) {
    bootstrapDefaultStaffIfNeeded().then(created => {
      if (created && location.hash === "#admin-dropdown-edit" && data.session.adminAuthed) renderDropdownEdit();
    });
  }
  // Task Types with prices
  const ttList=$("ddTaskTypeList");ttList.innerHTML="";
  getTaskTypes().forEach((t,i)=>{
    const price=getTaskPrice(t);
    const priceStr=t==="時給"||t==="その他（時給）"?"時給":(price!=null?price.toLocaleString()+"円":"―");
    const div=document.createElement("div");div.className="dd-item";
    div.innerHTML=`<div class="dd-info"><span>${escapeHtml(t)}</span><div class="dd-price">${priceStr}</div></div><div class="dd-btns"></div>`;
    const btns=div.querySelector(".dd-btns");
    const eBtn=document.createElement("button");eBtn.className="btn small ghost";eBtn.textContent="編集";
    eBtn.addEventListener("click",()=>{openDdEdit(i)});btns.appendChild(eBtn);
    const btn=document.createElement("button");btn.className="btn danger small";btn.textContent="削除";
    btn.addEventListener("click",()=>{if(!confirm(`「${data.taskTypes[i]}」を削除しますか？`))return;const name=data.taskTypes[i];data.taskTypes.splice(i,1);if(data.taskPrices&&data.taskPrices[name]!=null)delete data.taskPrices[name];saveData(data);renderDropdownEdit()});btns.appendChild(btn);
    ttList.appendChild(div);
  });
  // Employees
  const eList=$("ddEmployeeList");eList.innerHTML="";
  getEmployees().forEach((e,i)=>{const div=document.createElement("div");div.className="dd-item";div.innerHTML=`<div class="dd-info"><span>${escapeHtml(e)}</span></div><div class="dd-btns"></div>`;
    const btns=div.querySelector(".dd-btns");
    const btn=document.createElement("button");btn.className="btn danger small";btn.textContent="削除";
    btn.addEventListener("click",()=>{if(!confirm(`「${data.employees[i]}」を削除しますか？`))return;data.employees.splice(i,1);saveData(data);renderDropdownEdit()});btns.appendChild(btn);eList.appendChild(div)});
  // Staff list with full edit & delete
  const sList=$("ddStaffList");sList.innerHTML="";
  const addWrap=document.createElement("div");addWrap.style.marginBottom="10px";
  const addBtn=document.createElement("button");addBtn.className="btn primary";addBtn.textContent="スタッフ追加";
  addBtn.addEventListener("click",()=>openStaffCreate());addWrap.appendChild(addBtn);sList.appendChild(addWrap);
  Object.values(data.users).sort((a,b)=>{const aT=(a.userType||"学生")==="社会人"?0:1;const bT=(b.userType||"学生")==="社会人"?0:1;return aT!==bT?aT-bT:(a.createdAt||0)-(b.createdAt||0)}).forEach(u=>{
    const hr=getUserHourlyRate(u.id);const cls=u.userType==="社会人"?"tag shakaijin":"tag student";
    const div=document.createElement("div");div.className="dd-item";
    div.innerHTML=`<div class="dd-info"><span>${escapeHtml(u.name||u.id)}</span> <span class="${cls}" style="font-size:10px;">${escapeHtml(u.userType||"学生")}</span><div class="dd-price">${hr.toLocaleString()}円/h</div><div style="font-size:10px;color:var(--muted);margin-top:2px;">ID: ${escapeHtml(u.id)}</div></div><div class="dd-btns"></div>`;
    const btns=div.querySelector(".dd-btns");
    // Edit button - opens overlay with full staff fields
    const eBtn=document.createElement("button");eBtn.className="btn small ghost";eBtn.textContent="編集";
    eBtn.addEventListener("click",()=>openStaffEdit(u.id));
    btns.appendChild(eBtn);
    // Delete button
    const dBtn=document.createElement("button");dBtn.className="btn small danger";dBtn.textContent="削除";
dBtn.addEventListener("click", async ()=>{
  if(!confirm(`「${u.name||u.id}」を削除しますか？\nこのスタッフの全データ（日報・スタンプ等）も削除されます。`))return;
  if(!API_URL){showModal({title:"API未接続",sub:"⚙でURLを設定してください",big:"🔌"});return}

  try{
    const resp = await fetch(API_URL, {
      method: "POST",
      headers: {"Content-Type": "text/plain"},
      body: JSON.stringify({
        _action: "deleteStaffUser",
        token: getToken(),
        id: u.id
      }),
      redirect: "follow"
    });
    const r = await resp.json();
    applySyncMeta(syncMetaFromResult(r));
    if(!r.ok){
      showModal({title:"削除失敗",sub:r.error||"エラー",big:"🚫"});
      return;
    }
  }catch(e){
    showModal({title:"通信エラー",sub:"スタッフ削除に失敗しました。",big:"📡"});
    return;
  }

  delete data.users[u.id];
  if(data.userHourlyRates&&data.userHourlyRates[u.id]!=null)delete data.userHourlyRates[u.id];
  if(data.staffWorkStatus){
    const nm=u.name||u.id;
    if(data.staffWorkStatus[nm])delete data.staffWorkStatus[nm];
  }
  saveData(data);
  renderDropdownEdit();
  showModal({title:"削除完了",sub:`${u.name||u.id} を削除しました`,big:"🗑️"});
});
    btns.appendChild(dBtn);
    sList.appendChild(div);
  });
}
/* === STAFF EDIT OVERLAY === */
let _staffEditId = null;
let _staffEditMode = "edit";

function setStaffEditOverlayTitle(text) {
  const titleEl = document.querySelector("#staffEditOverlay h3");
  if (titleEl) titleEl.textContent = text;
}

function openStaffCreate() {
  _staffEditId = "__new__";
  _staffEditMode = "create";
  setStaffEditOverlayTitle("スタッフ新規追加");
  $("seId").value = "";
  $("sePw").value = "";
  $("seName").value = "";
  $("seType").value = "学生";
  $("seRate").value = "1300";
  $("staffEditOverlay").style.display = "flex";
}
async function openStaffEdit(userId) {
  _staffEditId = userId;
  _staffEditMode = "edit";
  try {
    await syncPull();
  } catch (_error) {}
  const u = data.users[userId]; if (!u) return;
  setStaffEditOverlayTitle("スタッフ情報編集");
  $("seId").value = u.id;
  $("sePw").value = "";
  $("seName").value = u.name || "";
  $("seType").value = u.userType || "学生";
  $("seRate").value = String(getUserHourlyRate(u.id));
  try { $("sePw").value = await fetchStaffPasswordForAdmin(u.id, true); } catch (_error) {}
  $("staffEditOverlay").style.display = "flex";
}
$("staffEditClose").addEventListener("click", () => { $("staffEditOverlay").style.display = "none"; });
$("staffEditOverlay").addEventListener("click", e => { if (e.target === $("staffEditOverlay")) $("staffEditOverlay").style.display = "none"; });
$("seRate").addEventListener("input", function(){ this.value = this.value.replace(/[^0-9]/g, ""); });
$("staffEditSave").addEventListener("click", async () => {
  if (!_staffEditId) return;
  const u = data.users[_staffEditId]; if (!u) return;
  const newId = $("seId").value.trim();
  const newPw = $("sePw").value.trim();
  const newName = $("seName").value.trim();
  const newType = $("seType").value;
  const newRate = parseInt($("seRate").value);
  if (!newId) { showModal({ title: "IDは必須です", big: "⚠️" }); return; }
  if (newId !== _staffEditId && data.users[newId]) { showModal({ title: "IDが重複しています", big: "🧩" }); return; }
  
  // ★APIを叩いてGASの隠し金庫（プロパティ）にパスワードを保存する処理を追加
  try {
  const oldId = _staffEditId;

  const resp = await fetch(API_URL, {
    method: "POST",
    headers: {"Content-Type": "text/plain"},
    body: JSON.stringify({
      _action: "upsertStaffUser",
      token: getToken(),
      id: newId,
      pw: newPw,
      name: newName||newId,
      userType: newType
    }),
    redirect: "follow"
  });
  const r = await resp.json();
  applySyncMeta(syncMetaFromResult(r));
  if (!r.ok) {
    showModal({title: "エラー", sub: r.error, big: "🚫"});
    return;
  }

  // ID変更時は旧IDもGAS側から削除
  if (newId !== oldId) {
    const delResp = await fetch(API_URL, {
      method: "POST",
      headers: {"Content-Type": "text/plain"},
      body: JSON.stringify({
        _action: "deleteStaffUser",
        token: getToken(),
        id: oldId
      }),
      redirect: "follow"
    });
    const delResult = await delResp.json();
    applySyncMeta(syncMetaFromResult(delResult));
    if (!delResult.ok) {
      showModal({title: "旧ID削除失敗", sub: delResult.error || "エラー", big: "🚫"});
      return;
    }
  }
} catch(e) {
  showModal({title: "通信エラー", big: "📡"});
  return;
}

// Update fields
u.name = newName || newId;
u.userType = newType;

// Update hourly rate
if (!isNaN(newRate) && newRate >= 0) {
  data.userHourlyRates = data.userHourlyRates || {};
  data.userHourlyRates[_staffEditId] = newRate;
}

// Handle ID change
if (newId !== _staffEditId) {
  const oldId = _staffEditId;
  u.id = newId;
  data.users[newId] = u;
  delete data.users[oldId];

  if (data.userHourlyRates && data.userHourlyRates[oldId] != null) {
    data.userHourlyRates[newId] = data.userHourlyRates[oldId];
    delete data.userHourlyRates[oldId];
  }

  _staffEditId = newId;
}

saveData(data);
$("staffEditOverlay").style.display = "none";
renderDropdownEdit();
showModal({ title: "更新完了", sub: `${u.name || u.id}`, big: "✅" });
});

/* function installStaffEditSaveOverride() {
  const oldBtn = $("staffEditSave");
  if (!oldBtn || oldBtn.dataset.overrideInstalled === "1") return;
  const newBtn = oldBtn.cloneNode(true);
  newBtn.dataset.overrideInstalled = "1";
  oldBtn.parentNode.replaceChild(newBtn, oldBtn);
  newBtn.addEventListener("click", async () => {
    const isCreate = _staffEditMode === "create";
    const oldId = _staffEditId;
    const existingUser = !isCreate && oldId ? data.users[oldId] : null;
    if (!isCreate && !existingUser) return;

    const newId = $("seId").value.trim();
    const newPw = $("sePw").value.trim();
    const newName = $("seName").value.trim();
    const newType = $("seType").value;
    const newRate = parseInt($("seRate").value, 10);

    if (!newId) { showModal({ title: "ID縺ｯ蠢・医〒縺・, big: "笞・・ }); return; }
    if ((isCreate || newId !== oldId) && data.users[newId]) { showModal({ title: "ID縺碁㍾隍・＠縺ｦ縺・∪縺・, big: "ｧｩ" }); return; }
    if (!API_URL) { showModal({title:"API譛ｪ謗･邯・,sub:"笞吶〒URL繧定ｨｭ螳壹＠縺ｦ縺上□縺輔＞",big:"伯"}); return; }
    if (!newPw && isCreate) { showModal({ title: "PW縺ｯ蠢・医〒縺・, big: "笞・・ }); return; }

    try {
      const resp = await fetch(API_URL, {
        method: "POST",
        headers: {"Content-Type": "text/plain"},
        body: JSON.stringify({
          _action: "upsertStaffUser",
          token: getToken(),
          id: newId,
          pw: newPw || (existingUser && existingUser.pw) || "",
          name: newName || newId,
          userType: newType
        }),
        redirect: "follow"
      });
      const r = await resp.json();
      applySyncMeta(syncMetaFromResult(r));
  if (!r.ok) {
    showModal({title: "繧ｨ繝ｩ繝ｼ", sub: r.error, big: "圻"});
    return;
  }
  invalidateStaffAccountsCache();

      if (!isCreate && newId !== oldId) {
        const delResp = await fetch(API_URL, {
          method: "POST",
          headers: {"Content-Type": "text/plain"},
          body: JSON.stringify({
            _action: "deleteStaffUser",
            token: getToken(),
            id: oldId
          }),
          redirect: "follow"
        });
        const delResult = await delResp.json();
        applySyncMeta(syncMetaFromResult(delResult));
        if (!delResult.ok) {
          showModal({title: "譌ｧID蜑企勁螟ｱ謨・, sub: delResult.error || "繧ｨ繝ｩ繝ｼ", big: "圻"});
          return;
        }
      }
    } catch (e) {
      showModal({title: "騾壻ｿ｡繧ｨ繝ｩ繝ｼ", big: "藤"});
      return;
    }

    const targetUser = isCreate ? {} : existingUser;
    targetUser.id = newId;
    targetUser.name = newName || newId;
    targetUser.userType = newType;
    targetUser.createdAt = targetUser.createdAt || Date.now();

    data.users = data.users || {};
    data.userHourlyRates = data.userHourlyRates || {};
    data.users[newId] = targetUser;
    if (!isNaN(newRate) && newRate >= 0) data.userHourlyRates[newId] = newRate;

    if (!isCreate && newId !== oldId) {
      delete data.users[oldId];
      if (data.userHourlyRates && data.userHourlyRates[oldId] != null) {
        if (data.userHourlyRates[newId] == null) data.userHourlyRates[newId] = data.userHourlyRates[oldId];
        delete data.userHourlyRates[oldId];
      }
    }

    _staffEditId = newId;
    _staffEditMode = "edit";
    setStaffEditOverlayTitle("スタッフ情報編集");
    saveData(data);
    $("staffEditOverlay").style.display = "none";
    renderDropdownEdit();
    showModal({ title: isCreate ? "スタッフ追加完了" : "譖ｴ譁ｰ螳御ｺ・, sub: `${targetUser.name || targetUser.id}`, big: "笨・ });
  });
}

installStaffEditSaveOverride(); */

function installStaffEditSaveOverride() {
  const oldBtn = $("staffEditSave");
  if (!oldBtn || oldBtn.dataset.overrideInstalled === "1") return;
  const newBtn = oldBtn.cloneNode(true);
  newBtn.dataset.overrideInstalled = "1";
  oldBtn.parentNode.replaceChild(newBtn, oldBtn);
  newBtn.addEventListener("click", async () => {
    const isCreate = _staffEditMode === "create";
    const oldId = _staffEditId;
    const existingUser = !isCreate && oldId && oldId !== "__new__" ? data.users[oldId] : null;
    if (!isCreate && !existingUser) return;

    const newId = $("seId").value.trim();
    const newPw = $("sePw").value.trim();
    const newName = $("seName").value.trim();
    const newType = $("seType").value;
    const newRate = parseInt($("seRate").value, 10);

    if (!newId) { showModal({ title: "ID is required", big: "NG" }); return; }
    if ((isCreate || newId !== oldId) && data.users[newId]) { showModal({ title: "ID already exists", big: "NG" }); return; }
    if (!API_URL) { showModal({ title: "API not connected", sub: "Check URL in settings", big: "NG" }); return; }
    let passwordForSave = newPw;

    if (!passwordForSave && isCreate) { showModal({ title: "PW is required", big: "NG" }); return; }
    if (!passwordForSave && !isCreate) {
      try {
        passwordForSave = await fetchStaffPasswordForAdmin(oldId, true);
      } catch (_error) {}
      if (!passwordForSave) {
        showModal({ title: "PW is required", sub: "Current password could not be loaded", big: "NG" });
        return;
      }
    }

    try {
      const resp = await fetch(API_URL, {
        method: "POST",
        headers: {"Content-Type": "text/plain"},
        body: JSON.stringify({
          _action: "upsertStaffUser",
          token: getToken(),
          id: newId,
          pw: passwordForSave,
          name: newName || newId,
          userType: newType
        }),
        redirect: "follow"
      });
      const r = await resp.json();
      applySyncMeta(syncMetaFromResult(r));
      if (!r.ok) {
        showModal({ title: "Save failed", sub: r.error || "error", big: "NG" });
        return;
      }
      invalidateStaffAccountsCache();

      if (!isCreate && newId !== oldId) {
        const delResp = await fetch(API_URL, {
          method: "POST",
          headers: {"Content-Type": "text/plain"},
          body: JSON.stringify({
            _action: "deleteStaffUser",
            token: getToken(),
            id: oldId
          }),
          redirect: "follow"
        });
        const delResult = await delResp.json();
        applySyncMeta(syncMetaFromResult(delResult));
        if (!delResult.ok) {
          showModal({ title: "Delete old ID failed", sub: delResult.error || "error", big: "NG" });
          return;
        }
      }
    } catch (e) {
      showModal({ title: "Network error", big: "NG" });
      return;
    }

    const targetUser = isCreate ? {} : existingUser;
    targetUser.id = newId;
    targetUser.name = newName || newId;
    targetUser.userType = newType;
    targetUser.createdAt = targetUser.createdAt || Date.now();

    data.users = data.users || {};
    data.userHourlyRates = data.userHourlyRates || {};
    data.users[newId] = targetUser;
    if (!isNaN(newRate) && newRate >= 0) data.userHourlyRates[newId] = newRate;

    if (!isCreate && newId !== oldId) {
      delete data.users[oldId];
      if (data.userHourlyRates && data.userHourlyRates[oldId] != null) {
        if (data.userHourlyRates[newId] == null) data.userHourlyRates[newId] = data.userHourlyRates[oldId];
        delete data.userHourlyRates[oldId];
      }
    }

    _staffEditId = newId;
    _staffEditMode = "edit";
    setStaffEditOverlayTitle("スタッフ情報編集");
    saveData(data);
    $("staffEditOverlay").style.display = "none";
    renderDropdownEdit();
    showModal({ title: isCreate ? "Staff added" : "Updated", sub: `${targetUser.name || targetUser.id}`, big: "OK" });
  });
}

installStaffEditSaveOverride();

function openDdEdit(idx){
  ddEditIdx=idx;
  const name=data.taskTypes[idx];
  const price=getTaskPrice(name);
  $("ddEditName").value=name;
  $("ddEditPrice").value=(name==="時給"||name==="その他（時給）"?"":price!=null?String(price):"");
  $("ddEditOverlay").style.display="flex";
}
$("ddEditClose").addEventListener("click",()=>{$("ddEditOverlay").style.display="none";var _ao=document.getElementById("apiSetupOverlay");if(_ao)_ao.style.display="none"});
$("ddEditOverlay").addEventListener("click",e=>{if(e.target===$("ddEditOverlay"))$("ddEditOverlay").style.display="none";var _ao=document.getElementById("apiSetupOverlay");if(_ao)_ao.style.display="none"});
$("ddEditPrice").addEventListener("input",function(){this.value=this.value.replace(/[^0-9]/g,"")});
$("ddEditSave").addEventListener("click",()=>{
  if(ddEditIdx<0)return;
  const oldName=data.taskTypes[ddEditIdx];
  const newName=$("ddEditName").value.trim();
  const newPrice=parseInt($("ddEditPrice").value);
  if(!newName){showModal({title:"業務種類名を入力してください",big:"⚠️"});return}
  // Update name
  data.taskTypes[ddEditIdx]=newName;
  // Update price
  if(!data.taskPrices)data.taskPrices={};
  if(oldName!==newName&&data.taskPrices[oldName]!=null){delete data.taskPrices[oldName]}
  if(newName!=="時給"&&newName!=="その他（時給）"&&!isNaN(newPrice)){data.taskPrices[newName]=newPrice}
  // Update task references
  data.tasks.forEach(t=>{if(t.taskType===oldName)t.taskType=newName});
  saveData(data);
  $("ddEditOverlay").style.display="none";var _ao=document.getElementById("apiSetupOverlay");if(_ao)_ao.style.display="none";
  renderDropdownEdit();
  showModal({title:"更新完了",big:"✅"});
});
$("ddAddTaskType").addEventListener("click",()=>{const v=$("ddNewTaskType").value.trim();if(!v)return;data.taskTypes.push(v);const p=parseInt($("ddNewTaskPrice").value);if(!isNaN(p)){if(!data.taskPrices)data.taskPrices={};data.taskPrices[v]=p}saveData(data);$("ddNewTaskType").value="";$("ddNewTaskPrice").value="";renderDropdownEdit()});
$("ddNewTaskPrice").addEventListener("input",function(){this.value=this.value.replace(/[^0-9]/g,"")});
$("ddAddEmployee").addEventListener("click",()=>{const v=$("ddNewEmployee").value.trim();if(!v)return;data.employees.push(v);saveData(data);$("ddNewEmployee").value="";renderDropdownEdit()});

/* === INIT === */

/* === MONTH CHECK === */
$("mcLogout").addEventListener("click",doAdminLogout);
tabNav("tabRM5","tabSM5","tabTL5","tabDD5","tabMC5");
$("mcDetailClose").addEventListener("click",()=>$("mcDetailCard").classList.add("hidden"));
let mcInit=false;
function renderMonthCheck(){
  renderAdminNotifications();
  if(!mcInit){
    const now=new Date();
    const ySel=$("mcYear");ySel.innerHTML="";
    for(let y=2024;y<=now.getFullYear()+1;y++){const o=document.createElement("option");o.value=y;o.textContent=y+"年";ySel.appendChild(o)}
    ySel.value=String(now.getFullYear());
    const mSel=$("mcMonth");mSel.innerHTML="";
    for(let m=1;m<=12;m++){const o=document.createElement("option");o.value=m;o.textContent=m+"月";mSel.appendChild(o)}
    mSel.value=String(now.getMonth()+1);
    mcInit=true;
    $("mcYear").addEventListener("change",doRenderMC);
    $("mcMonth").addEventListener("change",doRenderMC);
  }
  doRenderMC();
}
function doRenderMC(){
  const y=parseInt($("mcYear").value),m=parseInt($("mcMonth").value);
  $("mcThead").innerHTML=`<tr><th>#</th><th>ID</th><th>名前</th><th>種別</th><th>状態</th><th>スタンプ申請</th><th>出勤回数</th><th>ｲﾝｾﾝﾃｨﾌﾞ</th><th>給与(ｲﾝｾﾝﾃｨﾌﾞ含)</th><th>交通費</th><th>操作</th></tr>`;
  const tb=$("mcTbody");tb.innerHTML="";
  let allMatch=true;
  const users=Object.values(data.users).sort((a,b)=>{const aT=(a.userType||"学生")==="社会人"?0:1;const bT=(b.userType||"学生")==="社会人"?0:1;return aT!==bT?aT-bT:(a.createdAt||0)-(b.createdAt||0)});
  users.forEach((u,idx)=>{
    const result=checkUserMatch(u,y,m);
    if(result.status!=="一致")allMatch=false;
    const isShakaijin=u.userType==="社会人";
    const reps=filterReports(u.reports,String(y),String(m),"全て");
    let tSal=0,tTr=0,tP=0;
    reps.forEach(r=>{tSal+=Math.round(calcReportSalary(r,u.id));if(r.workType==="出勤")tTr+=parseInt(r.transport)||0;
      if(isShakaijin){tP+=r.incentiveAmount||0}else{tP+=(r.proofCount||0)*500}});
    const uniqueDays=new Set(reps.filter(r=>r.workType==="出勤").map(r=>r.date)).size;
    const hasStampReq=u.pendingStampRequest&&u.pendingStampRequest.status==="pending";
    const tr=document.createElement("tr");
    const statusCls=result.status==="一致"?"color:var(--mint)":result.status==="日報不足"?"color:var(--blue)":result.status==="業務過多"?"color:var(--orange)":"color:var(--red)";
    const cls=u.userType==="社会人"?"tag shakaijin":"tag student";
    [String(idx+1),null,u.name||u.id,null,null,hasStampReq?"有":"無",uniqueDays+"回",tP.toLocaleString()+"円",(tSal+tP).toLocaleString()+"円",tTr.toLocaleString()+"円",null].forEach((c,i)=>{
      const td=document.createElement("td");
      if(i===1)td.innerHTML=`<span class="tag">${escapeHtml(u.id)}</span>`;
      else if(i===3)td.innerHTML=`<span class="${cls}">${escapeHtml(u.userType||"学生")}</span>`;
      else if(i===4)td.innerHTML=`<span style="${statusCls};font-weight:900;">${result.status}</span>`;
      else if(i===10){
        const btn=document.createElement("button");btn.className="btn primary small";btn.textContent="詳細";
        btn.addEventListener("click",()=>showMCDetail(u,y,m));td.appendChild(btn);
      }
      else if(c!==null)td.textContent=c;
      tr.appendChild(td);
    });
    tb.appendChild(tr);
  });
  // Confirm button
  const ca=$("mcConfirmArea");ca.innerHTML="";
  if(allMatch){
    const btn=document.createElement("button");btn.className="btn success";btn.style.cssText="width:100%;font-size:15px;padding:14px;";
    btn.textContent="✅ この月の業務を確定する";
    btn.addEventListener("click",()=>{
      if(!confirm(y+"年"+m+"月の業務を確定しますか？確定すると日報・業務管理の指摘回数以外は変更不可になります。"))return;
      if(!data.lockedMonths)data.lockedMonths={};
      data.lockedMonths[y+"-"+String(m).padStart(2,"0")]=true;
      saveData(data);doRenderMC();showModal({title:"確定完了",sub:y+"年"+m+"月",big:"🔒"});
    });
    ca.appendChild(btn);
  }
  const mk=y+"-"+String(m).padStart(2,"0");
  if(data.lockedMonths&&data.lockedMonths[mk]){
    ca.innerHTML=`<div style="text-align:center;padding:14px;border-radius:14px;background:rgba(107,203,119,.1);border:2px solid rgba(107,203,119,.2);font-weight:900;color:var(--mint);">🔒 ${y}年${m}月は確定済みです</div>`;
  }
}
function checkUserMatch(u,y,m){
  const zaitakuReports=filterReports(u.reports,String(y),String(m),"在宅");
  const staffName=u.name||u.id;
  const staffTasks=data.tasks.filter(t=>{
    if(t.staff!==staffName)return false;
    const rd=t.requestDate?new Date(t.requestDate+"T00:00:00"):null;
    if(!rd)return false;
    if(rd.getFullYear()!==y||(rd.getMonth()+1)!==m)return false;
    return true;
  });
  // Group by taskType
  const reportByType={};
  zaitakuReports.forEach(r=>{const tt=r.taskType||"不明";reportByType[tt]=(reportByType[tt]||0)+(parseInt(r.manHours)||0)});
  const taskByType={};
  staffTasks.forEach(t=>{const tt=t.taskType||"不明";taskByType[tt]=(taskByType[tt]||0)+(parseInt(t.manHours)||0)});
  const allTypes=new Set([...Object.keys(reportByType),...Object.keys(taskByType)]);
  let hasShortage=false,hasExcess=false;
  const details=[];
  allTypes.forEach(tt=>{
    const rh=reportByType[tt]||0;
    const th=taskByType[tt]||0;
    let st="一致";
    if(rh<th){st="日報不足";hasShortage=true;}
    else if(rh>th){st="業務過多";hasExcess=true;}
    details.push({taskType:tt,reportHours:rh,taskHours:th,status:st});
  });
  let status="一致";
  if(hasShortage&&hasExcess)status="双方不一致";
  else if(hasShortage)status="日報不足";
  else if(hasExcess)status="業務過多";
  return{status,details};
}
function showMCDetail(u,y,m){
  const result=checkUserMatch(u,y,m);
  $("mcDetailTitle").textContent=`${u.name||u.id} の業務種類別照合`;
  $("mcDetailCard").classList.remove("hidden");
  $("mcDetailThead").innerHTML=`<tr><th>業務種類</th><th>日報工数</th><th>業務管理工数</th><th>状態</th></tr>`;
  const tb=$("mcDetailTbody");tb.innerHTML="";
  if(!result.details.length){tb.innerHTML=`<tr><td colspan="4" style="text-align:center;padding:20px;color:var(--muted);">データなし</td></tr>`;return}
  result.details.forEach(d=>{
    const tr=document.createElement("tr");
    const stCls=d.status==="一致"?"color:var(--mint)":d.status==="日報不足"?"color:var(--blue)":"color:var(--orange)";
    [d.taskType,String(d.reportHours),String(d.taskHours),null].forEach((c,i)=>{
      const td=document.createElement("td");
      if(i===3)td.innerHTML=`<span style="${stCls};font-weight:900;">${d.status}</span>`;
      else td.textContent=c;
      tr.appendChild(td);
    });
    tb.appendChild(tr);
  });
}


/* === ADMIN NOTIFICATION: completed tasks === */
function renderAdminNotifications(){
  const completed=data.tasks.filter(t=>t.status==="完了");
  const count=completed.length;
  const stampReqCount=countPendingStampRequests();
  const notifIds=["adminNotif1","adminNotif2","adminNotif3","adminNotif4","adminNotif5"];
  notifIds.forEach(id=>{
    const el=document.getElementById(id);if(!el)return;
    const msgs=[];
    if(count>0){
      msgs.push(`<span class="notif-icon">🔔</span><span>完了になった業務が <span class="notif-count" id="${id}Count">${count}件</span> あります</span>`);
    }
    if(stampReqCount>0){
      msgs.push(`<span class="notif-icon">📨</span><span>スタンプ修正申請が <span class="notif-count notif-stamp-req" id="${id}StampReq">${stampReqCount}件</span> あります</span>`);
    }
    if(msgs.length>0){
      el.classList.remove("hidden");
      el.innerHTML=msgs.join('<br style="margin:4px 0;">');
      const countEl=document.getElementById(id+"Count");
      if(countEl){countEl.addEventListener("click",()=>{
        location.hash="#admin-task-list";
        setTimeout(()=>{$("atlStatus").value="完了";doRenderATL()},100);
      })}
      const stampEl=document.getElementById(id+"StampReq");
      if(stampEl){stampEl.addEventListener("click",()=>{
        // スタンプ申請がある最初のユーザーの編集画面へ遷移
        const reqUser=Object.values(data.users).find(u=>u&&u.pendingStampRequest&&u.pendingStampRequest.status==="pending");
        if(reqUser){
          data.session.adminEditingUserId=reqUser.id;saveLocalOnly(data);
          editMonthCursor=startOfMonth(new Date());location.hash="#admin-edit";
        } else { location.hash="#admin"; }
      })}
    } else {
      el.classList.add("hidden");el.innerHTML="";
    }
  });
}
function renderAdminCreds() {
  $("adminCredId").value = "";
  $("adminCredOldPw").value = "";
  $("adminCredPw").value = "";
}
$("adminCredSave").addEventListener("click", async function(){
  const nid = $("adminCredId").value.trim();
  const oldPw = ($("adminCredOldPw")?$("adminCredOldPw").value:"").trim();
  const npw = $("adminCredPw").value.trim();

  if(!nid || !oldPw || !npw){
    showModal({title:"ID・旧PW・新PWは必須です",big:"⚠️"});
    return;
  }
  if(!API_URL){
    showModal({title:"API未接続です（⚙で設定）",big:"⚠️"});
    return;
  }
  try{
    const resp = await fetch(API_URL,{
      method:"POST",
      headers:{ "Content-Type":"text/plain" },
      body: JSON.stringify({_action:"setAdminCreds", token:getToken(), oldPw:oldPw, newId:nid, newPw:npw}),
      redirect:"follow"
    });
    const r = await resp.json();
    applySyncMeta(syncMetaFromResult(r));
    if(!r.ok){
      showModal({title:"更新失敗",sub:(r.error||"unauthorized"),big:"⚠️"});
      return;
    }
    showModal({title:"更新しました",sub:"次回ログインから新ID/PWです",big:"✅"});
    $("adminCredOldPw").value="";
    $("adminCredPw").value="";
  }catch(e){
    showModal({title:"通信エラー",sub:String(e),big:"📡"});
  }
});

(function(){
  if(!location.hash || location.hash==="#user-login") location.hash="#admin-login";
  route();
  if(document.readyState === "loading") document.addEventListener("DOMContentLoaded", initSync, {once:true});
  else initSync();
})();

function adminRebindNodeById(id, binder) {
  const node = document.getElementById(id);
  if (!node || !node.parentNode) return null;
  const clone = node.cloneNode(true);
  node.parentNode.replaceChild(clone, node);
  if (binder) binder(clone);
  return clone;
}

function installAdminDirectBindings() {
  ["ardFilterYear", "ardFilterMonth", "ardFilterType"].forEach(id => {
    adminRebindNodeById(id, node => {
      node.addEventListener("change", () => doRenderARD());
    });
  });
}

const _doRenderARDDirectBase = doRenderARD;
doRenderARD = function() {
  const originalSaveData = saveData;
  saveData = function() {};
  try {
    _doRenderARDDirectBase();
  } finally {
    saveData = originalSaveData;
  }

  const u = data.users[data.session.adminReportEditingUserId];
  const tbody = $("ardTbody");
  if (!u || !tbody) return;
  const filtered = filterReports(u.reports, $("ardFilterYear").value, $("ardFilterMonth").value, $("ardFilterType").value);
  const isShakaijin = u.userType === "遉ｾ莨壻ｺｺ";

  Array.from(tbody.children).forEach((row, idx) => {
    const report = filtered[idx];
    const originalIndex = report ? u.reports.indexOf(report) : -1;
    if (!report || originalIndex < 0) return;

    const reviewInput = row.querySelector('input[type="number"]');
    if (reviewInput && reviewInput.parentNode) {
      const cloneInput = reviewInput.cloneNode(true);
      reviewInput.parentNode.replaceChild(cloneInput, reviewInput);
      cloneInput.addEventListener("change", () => {
        const value = Math.max(0, parseInt(cloneInput.value, 10) || 0);
        const patch = isShakaijin ? { incentiveAmount: value } : { proofCount: value };
        setReportReviewRemote(u.id, report.reportId, originalIndex, patch)
          .then(() => doRenderARD())
          .catch(error => handleDirectActionError(error, "査定値の保存に失敗しました"));
      });
    }

    const deleteBtn = row.querySelector("button.btn.danger.small");
    if (deleteBtn && deleteBtn.parentNode) {
      const cloneDelete = deleteBtn.cloneNode(true);
      deleteBtn.parentNode.replaceChild(cloneDelete, deleteBtn);
      cloneDelete.addEventListener("click", () => {
        if (!confirm("この日報を削除しますか？")) return;
        deleteReportRemote(u.id, report.reportId, originalIndex)
          .then(() => doRenderARD())
          .catch(error => handleDirectActionError(error, "日報削除に失敗しました"));
      });
    }
  });
};

startEdit = function(td, u, oi, field) {
  if (td.classList.contains("editing")) return;
  td.classList.add("editing");
  const report = u.reports[oi];
  let inp;

  if (field === "workType") {
    inp = document.createElement("select");
    [report.workType || ""].filter(Boolean).forEach(v => {
      const o = document.createElement("option");
      o.value = v;
      o.textContent = v;
      inp.appendChild(o);
    });
    inp.value = report.workType || "";
  } else if (field === "date") {
    inp = document.createElement("input");
    inp.type = "date";
    inp.value = report.date || "";
  } else if (field === "startTime") {
    inp = document.createElement("input");
    inp.type = "text";
    inp.value = `${report.startH}:${report.startM}`;
  } else if (field === "endTime") {
    inp = document.createElement("input");
    inp.type = "text";
    inp.value = `${report.endH}:${report.endM}`;
  } else if (field === "breakTime") {
    inp = document.createElement("select");
    for (let i = 0; i <= 90; i += 15) {
      const o = document.createElement("option");
      o.value = String(i);
      o.textContent = String(i);
      inp.appendChild(o);
    }
    inp.value = String(parseInt(report.breakTime, 10) || 0);
  } else if (field === "content") {
    inp = document.createElement("textarea");
    inp.value = report.content || "";
  } else if (field === "taskType") {
    inp = document.createElement("select");
    getTaskTypes().forEach(v => {
      const o = document.createElement("option");
      o.value = v;
      o.textContent = v;
      inp.appendChild(o);
    });
    inp.value = report.taskType || "";
  } else if (field === "manHours") {
    inp = document.createElement("input");
    inp.type = "number";
    inp.min = "0";
    inp.value = report.manHours || "";
  } else if (field === "bizId") {
    inp = document.createElement("input");
    inp.type = "text";
    inp.value = report.bizId || "";
  } else if (field === "productId") {
    inp = document.createElement("input");
    inp.type = "text";
    inp.value = report.productId || "";
  } else if (field === "serviceId") {
    inp = document.createElement("input");
    inp.type = "text";
    inp.value = report.serviceId || "";
  } else if (field === "textCode") {
    inp = document.createElement("input");
    inp.type = "text";
    inp.value = report.textCode || "";
  } else if (field === "transport") {
    inp = document.createElement("input");
    inp.type = "text";
    inp.value = report.transport || "0";
  } else {
    inp = document.createElement("input");
    inp.type = "text";
    inp.value = td.textContent;
  }

  td.textContent = "";
  td.appendChild(inp);
  inp.focus();

  let handled = false;
  const save = async () => {
    if (handled) return;
    handled = true;
    td.classList.remove("editing");

    const next = Object.assign({}, report);
    if (field === "date") next.date = inp.value;
    else if (field === "workType") next.workType = inp.value;
    else if (field === "startTime") {
      const p = inp.value.split(":");
      next.startH = pad2(parseInt(p[0], 10) || 0);
      next.startM = pad2(parseInt(p[1], 10) || 0);
    } else if (field === "endTime") {
      const p = inp.value.split(":");
      next.endH = pad2(parseInt(p[0], 10) || 0);
      next.endM = pad2(parseInt(p[1], 10) || 0);
    } else if (field === "breakTime") next.breakTime = inp.value;
    else if (field === "content") next.content = inp.value;
    else if (field === "taskType") next.taskType = inp.value;
    else if (field === "manHours") next.manHours = inp.value;
    else if (field === "bizId") next.bizId = inp.value;
    else if (field === "productId") next.productId = inp.value;
    else if (field === "serviceId") next.serviceId = inp.value;
    else if (field === "textCode") next.textCode = inp.value;
    else if (field === "transport") next.transport = inp.value;

    const sh = parseInt(next.startH, 10) || 0;
    const sm = parseInt(next.startM, 10) || 0;
    const eh = parseInt(next.endH, 10) || 0;
    const em = parseInt(next.endM, 10) || 0;
    const brk = parseInt(next.breakTime, 10) || 0;
    let minutes = (eh * 60 + em) - (sh * 60 + sm) - brk;
    if (minutes < 0) minutes = 0;
    next.workTime = `${Math.floor(minutes / 60)}h${minutes % 60 > 0 ? minutes % 60 + "m" : ""}`;

    try {
      await updateReportRemote(u.id, report.reportId, oi, next);
      doRenderARD();
    } catch (error) {
      handleDirectActionError(error, "日報更新に失敗しました");
      doRenderARD();
    }
  };

  inp.addEventListener("blur", () => { save(); });
  inp.addEventListener("keydown", e => {
    if (e.key === "Enter" && field !== "content") {
      e.preventDefault();
      inp.blur();
    }
  });
};

renderAdminEdit = function() {
  const u = data.users[data.session.adminEditingUserId];
  if (!u) {
    location.hash = "#admin";
    return;
  }

  $("editUserName").textContent = u.name || u.id;
  $("editUserId").textContent = u.id;
  $("editUid").value = u.id;
  $("editUname").value = u.name || "";
  $("editUpw").value = "";
  $("editUserType").value = u.userType || "";
  fillStaffPasswordField("editUpw", u.id);

  const now = new Date();
  const total = countTotal(u);
  $("eTotal").textContent = total;
  $("eMonth").textContent = countThisMonth(u, now);
  $("eWeek").textContent = countThisWeek(u, now);
  $("eMonthKey").textContent = ym(editMonthCursor);

  const sInc = calcStampIncentive(total);
  $("incentiveDisplay").innerHTML = `<div class="incentive-box"><div class="ib-title">Stamp Incentive</div><div style="font-family:var(--font-display);font-size:20px;font-weight:900;color:var(--pink);">${sInc.toLocaleString()}円</div><div style="font-size:11px;color:var(--muted);margin-top:4px;">累計 ${total}pt</div></div>`;
  $("editMonthLabel").textContent = monthLabelJa(editMonthCursor);

  const hasPending = u.pendingStampRequest && u.pendingStampRequest.status === "pending";
  if (hasPending) {
    renderCalendar({
      mount: $("editCal"),
      monthCursor: editMonthCursor,
      stampedMap: u.pendingStampRequest.stamps,
      clickable: true,
      onDayClick: async d => {
        const key = ymd(d);
        const nextStamps = cloneDeep((u.pendingStampRequest && u.pendingStampRequest.stamps) || {});
        const cur = nextStamps[key];
        if (!cur) nextStamps[key] = true;
        else if (cur === true) nextStamps[key] = "emergency";
        else delete nextStamps[key];
        try {
          await updateStampRequestDraftRemote(u.id, nextStamps);
          renderAdminEdit();
        } catch (error) {
          handleDirectActionError(error, "申請下書きの更新に失敗しました");
        }
      },
      pendingChanges: u.pendingStampRequest.stamps,
      originalStamps: u.stamps
    });
  } else {
    renderCalendar({
      mount: $("editCal"),
      monthCursor: editMonthCursor,
      stampedMap: u.stamps,
      clickable: true,
      onDayClick: async d => {
        const key = ymd(d);
        const cur = u.stamps[key];
        const nextValue = !cur ? true : (cur === true ? "emergency" : null);
        try {
          await setStampRemote(u.id, key, nextValue, null);
          renderAdminEdit();
        } catch (error) {
          handleDirectActionError(error, "スタンプ更新に失敗しました");
        }
      }
    });
  }

  let reqBar = document.getElementById("adminStampReqBar");
  if (!reqBar) {
    reqBar = document.createElement("div");
    reqBar.id = "adminStampReqBar";
    $("editCal").parentNode.appendChild(reqBar);
  }
  reqBar.innerHTML = "";

  if (hasPending) {
    reqBar.className = "admin-request-info";
    reqBar.innerHTML = `<div class="ari-title">スタンプ修正申請</div><div style="font-size:11px;color:var(--muted);margin-bottom:8px;">内容を確認して承認または却下してください</div>`;
    const btnWrap = document.createElement("div");
    btnWrap.style.cssText = "display:flex;gap:8px;";

    const approveBtn = document.createElement("button");
    approveBtn.className = "btn success small";
    approveBtn.textContent = "承認";
    approveBtn.addEventListener("click", () => {
      resolveStampCorrectionRemote(u.id, "approved")
        .then(() => {
          showModal({ title: "承認しました", sub: `${u.name || u.id} のスタンプを更新しました`, big: "OK" });
          renderAdminEdit();
        })
        .catch(error => handleDirectActionError(error, "申請承認に失敗しました"));
    });

    const rejectBtn = document.createElement("button");
    rejectBtn.className = "btn danger small";
    rejectBtn.textContent = "却下";
    rejectBtn.addEventListener("click", () => {
      resolveStampCorrectionRemote(u.id, "rejected")
        .then(() => {
          showModal({ title: "却下しました", sub: `${u.name || u.id} の申請を却下しました`, big: "OK" });
          renderAdminEdit();
        })
        .catch(error => handleDirectActionError(error, "申請却下に失敗しました"));
    });

    btnWrap.appendChild(approveBtn);
    btnWrap.appendChild(rejectBtn);
    reqBar.appendChild(btnWrap);
  } else {
    reqBar.className = "";
  }
};

function getAdminOverviewStats() {
  const tasks = data.tasks || [];
  const today = ymd(new Date());
  return {
    staffCount: Object.values(data.users || {}).filter(u => u && (u.userType || "") !== "遉ｾ莨壻ｺｺ").length,
    pendingStamp: countPendingStampRequests(),
    overdueTasks: tasks.filter(task => task.status === "譛滄剞雜・℃").length,
    workingTasks: tasks.filter(task => task.status === "萓晞ｼ蜑・" || task.status === "萓晞ｼ荳ｭ").length,
    completedToday: tasks.filter(task => task.status === "螳御ｺ・" && task.completionDate === today).length
  };
}

function openPendingStampSummary() {
  data.session = data.session || {};
  data.session.adminHomeFilter = "pending-stamp";
  saveLocalOnly(data);
  location.hash = "#admin";
}

function openOverdueTaskSummary() {
  if (data.session && data.session.adminHomeFilter) {
    delete data.session.adminHomeFilter;
    saveLocalOnly(data);
  }
  location.hash = "#admin-task-list";
  setTimeout(() => {
    if ($("atlStatus")) $("atlStatus").value = "期限超過";
    if (typeof doRenderATL === "function") doRenderATL();
  }, 120);
}

function clearAdminHomeSummaryFilter() {
  data.session = data.session || {};
  if (!data.session.adminHomeFilter) return;
  delete data.session.adminHomeFilter;
  saveLocalOnly(data);
}

function applyAdminHomeSummaryFilter() {
  const section = document.getElementById("adminHome");
  const tbody = document.getElementById("adminTbody");
  if (!section || !tbody) return;

  const filterKey = (data.session && data.session.adminHomeFilter) || "";
  let notice = document.getElementById("adminHomeSummaryFilter");
  if (filterKey !== "pending-stamp") {
    if (notice) notice.remove();
    Array.from(tbody.querySelectorAll("tr")).forEach(row => { row.style.display = ""; });
    return;
  }

  const pendingIds = new Set(
    Object.entries(data.users || {})
      .filter(([, user]) => user && user.pendingStampRequest && user.pendingStampRequest.status === "pending")
      .map(([id, user]) => String((user && user.id) || id || ""))
  );

  let visibleCount = 0;
  Array.from(tbody.querySelectorAll("tr")).forEach(row => {
    const idCell = row.children[1];
    const rowId = String((idCell && idCell.textContent) || "").trim();
    const visible = pendingIds.has(rowId);
    row.style.display = visible ? "" : "none";
    if (visible) visibleCount += 1;
  });

  if (!notice) {
    notice = document.createElement("div");
    notice.id = "adminHomeSummaryFilter";
    notice.className = "card";
    const cards = section.querySelectorAll(".card");
    const anchor = cards[cards.length - 1] || null;
    if (anchor && anchor.parentNode === section) section.insertBefore(notice, anchor);
    else section.appendChild(notice);
  }

  notice.innerHTML = `
    <div class="bd" style="display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;">
      <div>
        <div style="font-weight:900;">Pending stamp requests</div>
        <div class="sub">${visibleCount} user(s) are shown</div>
      </div>
      <button type="button" class="btn small" id="adminHomeSummaryFilterClear">Show all</button>
    </div>
  `;

  const clearBtn = document.getElementById("adminHomeSummaryFilterClear");
  if (clearBtn) {
    clearBtn.addEventListener("click", () => {
      clearAdminHomeSummaryFilter();
      renderAdminHome();
    });
  }
}

function renderAdminOverviewDashboard() {
  const section = document.getElementById("adminHome");
  if (!section) return;
  const firstCard = section.querySelector(".card");
  let mount = document.getElementById("adminOverviewDashboard");
  if (!mount) {
    mount = document.createElement("div");
    mount.id = "adminOverviewDashboard";
    if (firstCard && firstCard.parentNode === section) section.insertBefore(mount, firstCard.nextSibling);
    else section.appendChild(mount);
  }
  mount.className = "dash-shell";

  const stats = getAdminOverviewStats();
  const alert = stats.pendingStamp > 0 || stats.overdueTasks > 0
    ? `要確認: スタンプ申請 ${stats.pendingStamp} 件 / 期限超過タスク ${stats.overdueTasks} 件`
    : "";

  mount.innerHTML = `
    <div class="dash-head">
      <div>
        <div class="dash-title">Admin Dashboard</div>
        <div class="dash-sub">最初に確認したい状況をまとめています</div>
      </div>
      <div class="dash-chip">今日の完了 ${stats.completedToday} 件</div>
    </div>
    <div class="dash-grid">
      <div class="dash-card">
        <div class="dash-label">スタッフ数</div>
        <div class="dash-value">${stats.staffCount}</div>
        <div class="dash-note">学生スタッフを集計</div>
      </div>
      <div class="dash-card">
        <div class="dash-label">スタンプ申請</div>
        <div class="dash-value">${stats.pendingStamp}</div>
        <div class="dash-note">確認待ち件数</div>
      </div>
      <div class="dash-card">
        <div class="dash-label">期限超過タスク</div>
        <div class="dash-value">${stats.overdueTasks}</div>
        <div class="dash-note">優先対応が必要</div>
      </div>
      <div class="dash-card">
        <div class="dash-label">進行中タスク</div>
        <div class="dash-value">${stats.workingTasks}</div>
        <div class="dash-note">依頼中 / 作業中</div>
      </div>
    </div>
    ${alert ? `<div class="dash-alert">${alert}</div>` : ""}
    <div class="dash-actions">
      <button type="button" class="dash-action admin" data-admin-nav="pending-stamp">申請を確認 <span>></span></button>
      <button type="button" class="dash-action" data-admin-nav="overdue-task">期限超過を見る <span>></span></button>
      <button type="button" class="dash-action" data-admin-nav="report-mgmt">日報管理 <span>></span></button>
      <button type="button" class="dash-action" data-admin-nav="month-check">月次チェック <span>></span></button>
    </div>
  `;

  mount.querySelectorAll("[data-admin-nav]").forEach(btn => {
    btn.addEventListener("click", () => {
      const action = btn.getAttribute("data-admin-nav");
      if (action === "pending-stamp") { openPendingStampSummary(); return; }
      if (action === "overdue-task") { openOverdueTaskSummary(); return; }
      if (action === "pending-stamp") {
        const reqUser = Object.values(data.users || {}).find(user => user && user.pendingStampRequest && user.pendingStampRequest.status === "pending");
        if (reqUser) {
          data.session.adminEditingUserId = reqUser.id;
          saveLocalOnly(data);
          editMonthCursor = startOfMonth(new Date());
          location.hash = "#admin-edit";
        } else {
          location.hash = "#admin";
        }
      } else if (action === "overdue-task") {
        location.hash = "#admin-task-list";
        setTimeout(() => {
          if ($("atlStatus")) $("atlStatus").value = "譛滄剞雜・℃";
          if (typeof doRenderATL === "function") doRenderATL();
        }, 120);
      } else if (action === "report-mgmt") {
        location.hash = "#admin-report-mgmt";
      } else if (action === "month-check") {
        location.hash = "#admin-month-check";
      }
    });
  });
}

const _renderAdminHomeDashboardBase = renderAdminHome;
renderAdminHome = function() {
  _renderAdminHomeDashboardBase();
  const oldOverview = document.getElementById("adminOverviewDashboard");
  if (oldOverview) oldOverview.remove();
};

function hideAdminDuplicateNav(sectionId) {
  const section = document.getElementById(sectionId);
  if (!section) return;
  section.querySelectorAll(".admin-tabs").forEach(el => { el.style.display = "none"; });
}

function renderAdminGlobalDashboard(sectionId, mountId, currentKey) {
  const section = document.getElementById(sectionId);
  if (!section) return;
  hideAdminDuplicateNav(sectionId);
  const topbar = section.querySelector(".topbar");
  let mount = document.getElementById(mountId);
  if (!mount) {
    mount = document.createElement("div");
    mount.id = mountId;
    if (topbar && topbar.parentNode === section) section.insertBefore(mount, topbar.nextSibling);
    else section.appendChild(mount);
  }
  mount.className = "dash-shell";
  const stats = getAdminOverviewStats();
  const items = [
    { key: "task", label: "業務管理" },
    { key: "report", label: "日報管理" },
    { key: "stamp", label: "スタンプ管理" },
    { key: "master", label: "設定" },
    { key: "month", label: "月次確認" }
  ];
  mount.innerHTML = `
    <div class="dash-head">
      <div>
        <div class="dash-title">業務管理ナビ</div>
        <div class="dash-sub">どの管理画面からでも主要メニューと今日の合言葉を確認できます</div>
      </div>
    </div>
    <div class="dash-grid" style="margin-bottom:12px;">
      <div class="dash-card">
        <div class="dash-label">今日の合言葉</div>
        <div class="dash-value small" id="${mountId}DailyPassword">読み込み中...</div>
        <div class="dash-note">管理画面のどこからでも確認できます</div>
      </div>
    </div>
    <div class="dash-actions">
      ${items.map(item => `<button type="button" class="dash-action${item.key === currentKey ? " admin" : ""}" data-admin-global="${item.key}">${item.label}<span>${item.key === currentKey ? "●" : ">"}</span></button>`).join("")}
    </div>
  `;
  const pwEl = document.getElementById(`${mountId}DailyPassword`);
  if (pwEl) {
    pwEl.textContent = "読み込み中...";
    fetchTodayPasswordForAdmin(false)
      .then(pw => { pwEl.textContent = pw || "---"; })
      .catch(() => { pwEl.textContent = "---"; });
  }
  mount.querySelectorAll("[data-admin-global]").forEach(btn => {
    btn.addEventListener("click", () => {
      const action = btn.getAttribute("data-admin-global");
      if (action === "task") location.hash = "#admin-task-list";
      else if (action === "report") location.hash = "#admin-report-mgmt";
      else if (action === "stamp") {
        clearAdminHomeSummaryFilter();
        location.hash = "#admin";
      }
      else if (action === "master") location.hash = "#admin-dropdown-edit";
      else if (action === "month") location.hash = "#admin-month-check";
    });
  });
}

renderAdminGlobalDashboard = function(sectionId, mountId, currentKey) {
  const section = document.getElementById(sectionId);
  if (!section) return;
  hideAdminDuplicateNav(sectionId);
  const topbar = section.querySelector(".topbar");
  let mount = document.getElementById(mountId);
  if (!mount) {
    mount = document.createElement("div");
    mount.id = mountId;
    if (topbar && topbar.parentNode === section) section.insertBefore(mount, topbar.nextSibling);
    else section.appendChild(mount);
  }
  mount.className = "dash-shell";

  const stats = getAdminOverviewStats();
  const items = [
    { key: "task", label: "業務管理" },
    { key: "report", label: "日報管理" },
    { key: "stamp", label: "スタンプ管理" },
    { key: "master", label: "設定" },
    { key: "month", label: "月次確認" }
  ];

  mount.innerHTML = `
    <div class="dash-head">
      <div>
        <div class="dash-title">業務管理ナビ</div>
        <div class="dash-sub">どの管理画面からでも主要メニューと今日の合言葉を確認できます</div>
      </div>
    </div>
    <div class="dash-grid" style="margin-bottom:12px;">
      <div class="dash-card">
        <div class="dash-label">今日の合言葉</div>
        <div class="dash-value small" id="${mountId}DailyPassword">読み込み中...</div>
        <div class="dash-note">管理画面のどこからでも確認できます</div>
      </div>
      <div class="dash-card">
        <div class="dash-label">スタンプ申請</div>
        <button type="button" class="dash-value-link" data-admin-summary="pending-stamp">${stats.pendingStamp}</button>
        <div class="dash-note">数字を押すと申請一覧へ移動</div>
      </div>
      <div class="dash-card">
        <div class="dash-label">期限超過タスク</div>
        <button type="button" class="dash-value-link" data-admin-summary="overdue-task">${stats.overdueTasks}</button>
        <div class="dash-note">数字を押すと業務一覧へ移動</div>
      </div>
    </div>
    <div class="dash-actions">
      ${items.map(item => `<button type="button" class="dash-action${item.key === currentKey ? " admin" : ""}" data-admin-global="${item.key}">${item.label}<span>${item.key === currentKey ? "●" : ">"}</span></button>`).join("")}
    </div>
  `;

  const pwEl = document.getElementById(`${mountId}DailyPassword`);
  if (pwEl) {
    pwEl.textContent = "読み込み中...";
    fetchTodayPasswordForAdmin(false)
      .then(pw => { pwEl.textContent = pw || "---"; })
      .catch(() => { pwEl.textContent = "---"; });
  }

  mount.querySelectorAll("[data-admin-global]").forEach(btn => {
    btn.addEventListener("click", () => {
      const action = btn.getAttribute("data-admin-global");
      if (action === "task") location.hash = "#admin-task-list";
      else if (action === "report") location.hash = "#admin-report-mgmt";
      else if (action === "stamp") location.hash = "#admin";
      else if (action === "master") location.hash = "#admin-dropdown-edit";
      else if (action === "month") location.hash = "#admin-month-check";
    });
  });

  mount.querySelectorAll("[data-admin-summary]").forEach(btn => {
    btn.addEventListener("click", () => {
      const action = btn.getAttribute("data-admin-summary");
      if (action === "pending-stamp") openPendingStampSummary();
      else if (action === "overdue-task") openOverdueTaskSummary();
    });
  });
};

const _renderAdminHomeGlobalBase = renderAdminHome;
renderAdminHome = function() {
  _renderAdminHomeGlobalBase();
  if (DEFAULT_BOOTSTRAP_STAFF.some(staff => !((data.users || {})[staff.id]))) {
    bootstrapDefaultStaffIfNeeded().then(created => {
      if (created && location.hash === "#admin" && data.session.adminAuthed) renderAdminHome();
    });
  }
  const legacyTodayPassword = document.getElementById("adminTodayPassword");
  const legacyTodayPasswordCard = legacyTodayPassword && legacyTodayPassword.closest(".card");
  if (legacyTodayPasswordCard) legacyTodayPasswordCard.remove();
  const title = document.querySelector("#adminHome .topbar .brand h1");
  if (title) title.textContent = "スタンプ管理";
  renderAdminGlobalDashboard("adminHome", "adminHomeGlobalNav", "stamp");
  applyAdminHomeSummaryFilter();
};

renderAdminOverviewDashboard = function() {};

const _renderAdminReportMgmtGlobalBase = renderAdminReportMgmt;
renderAdminReportMgmt = function() {
  _renderAdminReportMgmtGlobalBase();
  const subTabs = document.getElementById("armSubTabs");
  if (subTabs) subTabs.remove();
  const addUserCard = document.getElementById("armAddUserCard");
  if (addUserCard) addUserCard.remove();
  const summaryCard = document.getElementById("armSummaryCard");
  if (summaryCard) summaryCard.classList.remove("hidden");
  renderAdminGlobalDashboard("adminReportMgmt", "adminReportMgmtNav", "report");
};

const _renderAdminTaskListGlobalBase = renderAdminTaskList;
renderAdminTaskList = function() {
  _renderAdminTaskListGlobalBase();
  const title = document.querySelector("#adminTaskList .topbar .brand h1");
  if (title) title.textContent = "業務管理";
  ensureTaskImportControls();
  renderAdminGlobalDashboard("adminTaskList", "adminTaskListNav", "task");
};

const _renderDropdownEditGlobalBase = renderDropdownEdit;
renderDropdownEdit = function() {
  _renderDropdownEditGlobalBase();
  const section = document.getElementById("adminDropdownEdit");
  const grid = section && section.querySelector(".grid");
  const employeeList = document.getElementById("ddEmployeeList");
  const staffList = document.getElementById("ddStaffList");
  const employeeCard = employeeList && employeeList.closest(".card");
  const staffCard = staffList && staffList.closest(".card");
  if (grid && employeeCard && staffCard && employeeCard !== staffCard) {
    grid.insertBefore(staffCard, employeeCard);
  }
  renderAdminGlobalDashboard("adminDropdownEdit", "adminDropdownEditNav", "master");
};

const _renderMonthCheckGlobalBase = renderMonthCheck;
renderMonthCheck = function() {
  _renderMonthCheckGlobalBase();
  renderAdminGlobalDashboard("adminMonthCheck", "adminMonthCheckNav", "month");
};

const _renderAdminReportDetailGlobalBase = renderAdminReportDetail;
renderAdminReportDetail = function() {
  _renderAdminReportDetailGlobalBase();
  renderAdminGlobalDashboard("adminReportDetail", "adminReportDetailNav", "report");
};

const _renderAdminEditGlobalBase = renderAdminEdit;
renderAdminEdit = function() {
  _renderAdminEditGlobalBase();
  renderAdminGlobalDashboard("adminEdit", "adminEditNav", "stamp");
};

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", installAdminDirectBindings, { once: true });
} else {
  installAdminDirectBindings();
}

route();
