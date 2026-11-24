// 使用人＋密碼登入：打開網站要先登入才能使用。
// 帳號清單放在公司 Google 試算表的「使用者」分頁（和雲端同步共用同一組帳號）。
// 第一次登入需要網路；登入後記在這台手機，之後沒網路也能用。
// 有網路時每次開啟會在背景再確認一次，帳號被刪除或停用就會被登出。
(function(){
  "use strict";
  var CFG_KEY="jiaoliao_sync_cfg";
  var SYNC_URL="https://script.google.com/macros/s/AKfycbwf1DwO6ARo_ePfa3TT3ZPecLmNnqjksK3sPxcIN-5_Obn3mHSQerYo1rRzQRJyNtM/exec";
  // 管理員（可以新增、刪除使用人）。後台 Code.gs 也有同一份名單，以後台為準。
  var ADMINS=["李羿萱","蔡鴻基"];
  // 超過這段時間沒有操作就自動登出（5 分鐘）
  var IDLE_MS=5*60*1000;
  var ACT_KEY="jiaoliao_last_act";

  function jget(k,d){try{var s=localStorage.getItem(k);return s?JSON.parse(s):d}catch(e){return d}}
  function jset(k,v){try{localStorage.setItem(k,JSON.stringify(v))}catch(e){}}
  function cfg(){var c=jget(CFG_KEY,{});c.url=c.url||SYNC_URL;return c}
  function loggedIn(){var c=cfg();return !!(c.who&&c.code)}
  function unlocked(){return loggedIn()&&!cfg().locked}
  function lastAct(){return +localStorage.getItem(ACT_KEY)||0}
  function touch(){try{localStorage.setItem(ACT_KEY,String(Date.now()))}catch(e){}}
  function isAdmin(){var c=cfg();return c.role==="admin"||ADMINS.indexOf(c.who)>=0}
  function esc(s){return String(s==null?"":s).replace(/[&<>"]/g,function(ch){return{"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[ch]})}

  function api(body){
    var c=cfg();
    return fetch(c.url||SYNC_URL,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify(body)})
      .then(function(r){return r.json()});
  }

  // ---------- 樣式 ----------
  var css=document.createElement("style");
  css.textContent=
    "#loginGate{position:fixed;inset:0;z-index:9999;background:#f3f4f8;display:flex;align-items:center;justify-content:center;padding:16px;overflow:auto}"+
    "#loginGate .lg-card{background:#fff;border-radius:20px;box-shadow:0 10px 40px rgba(0,0,0,.12);width:100%;max-width:420px;padding:28px 22px}"+
    "#loginGate .lg-logo{display:flex;align-items:center;gap:12px;margin-bottom:22px}"+
    "#loginGate .lg-logo img{width:56px;height:56px;border-radius:14px}"+
    "#loginGate h1{font-size:24px;margin:0;color:#2f5a1c}"+
    "#loginGate .lg-logo small{display:block;font-size:15px;color:#6b7280;margin-top:2px}"+
    "#loginGate label{display:block;font-size:18px;font-weight:600;color:#1f2937;margin:14px 0 6px}"+
    "#loginGate input{display:block;width:100%;font-size:22px;padding:12px 14px;border:2px solid #d1d5db;border-radius:12px;background:#fff;color:#1f2937}"+
    "#loginGate input:focus{outline:none;border-color:#4d7c2a}"+
    "#loginGate .lg-show{display:flex;align-items:center;gap:8px;font-size:16px;font-weight:400;margin:10px 0 0;color:#4b5563}"+
    "#loginGate .lg-show input{width:22px;height:22px;display:inline-block;padding:0}"+
    "#loginGate .lg-go{display:block;width:100%;margin-top:22px;font-size:22px;font-weight:700;padding:14px;border:0;border-radius:14px;background:#2f5a1c;color:#fff}"+
    "#loginGate .lg-go:disabled{opacity:.6}"+
    "#loginGate .lg-msg{min-height:26px;margin:14px 0 0;font-size:17px;color:#dc2626;text-align:center}"+
    "#loginGate .lg-tip{font-size:14px;color:#6b7280;margin:16px 0 0;text-align:center;line-height:1.6}"+
    "body.lg-locked{overflow:hidden}"+
    "#userDlg .ur{display:grid;grid-template-columns:1fr 1fr auto;gap:6px;align-items:center;padding:8px 0;border-bottom:1px solid #e5e7eb}"+
    "#userDlg .ur b{font-size:16px}#userDlg .ur span{font-size:15px;color:#4b5563}"+
    "#userDlg .ur .tag{font-size:12px;border-radius:999px;padding:1px 8px;margin-left:4px;background:#eef6e4;color:#2f5a1c}"+
    "#userDlg .ur .tag.off{background:#fee2e2;color:#dc2626}"+
    "#userDlg .ur-act{display:flex;gap:4px;flex-wrap:wrap;justify-content:flex-end}"+
    "#userDlg .add{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:12px}"+
    "#userDlg .add input{font-size:17px;padding:10px;border:1px solid #d1d5db;border-radius:10px;width:100%}"+
    "#userDlg .add .btn{grid-column:1/-1;font-size:17px;padding:10px}";
  document.head.appendChild(css);

  // ---------- 登入畫面 ----------
  var gate=null;
  function showGate(msg){
    if(gate){gate.querySelector(".lg-msg").textContent=msg||"";return}
    var c=cfg();
    gate=document.createElement("div");
    gate.id="loginGate";
    gate.innerHTML='<form class="lg-card" autocomplete="on">'+
      '<div class="lg-logo"><img src="app-icon-192.png" alt=""><div><h1>工程叫料小幫手</h1><small>請先登入</small></div></div>'+
      '<label for="lgWho">使用人（姓名）</label><input id="lgWho" name="username" autocomplete="username" placeholder="例如：王小明" value="'+esc(c.who||"")+'">'+
      '<label for="lgPw">密碼</label><input id="lgPw" name="password" type="password" inputmode="tel" autocomplete="current-password" placeholder="請輸入密碼">'+
      '<label class="lg-show"><input type="checkbox" id="lgShow"> 顯示密碼</label>'+
      '<button class="lg-go" type="submit">登入</button>'+
      '<p class="lg-msg">'+esc(msg||"")+'</p>'+
      '<p class="lg-tip">第一次登入需要連上網路。<br>忘記密碼請找管理員 李羿萱。</p>'+
      '</form>';
    document.body.appendChild(gate);
    document.body.classList.add("lg-locked");
    var form=gate.querySelector("form"),who=gate.querySelector("#lgWho"),pw=gate.querySelector("#lgPw"),go=gate.querySelector(".lg-go"),m=gate.querySelector(".lg-msg");
    gate.querySelector("#lgShow").onchange=function(){pw.type=this.checked?"text":"password"};
    form.onsubmit=function(ev){
      ev.preventDefault();
      var w=who.value.trim(),p=pw.value.trim();
      if(!w||!p){m.textContent="請輸入使用人和密碼";return}
      var old=cfg();
      if(old.who===w&&old.code===p&&(!navigator.onLine)){old.locked=false;jset(CFG_KEY,old);touch();hideGate();return}
      if(!navigator.onLine){m.textContent=old.who===w&&old.code?"密碼不正確":"目前沒有網路，第一次登入要連上網路";return}
      go.disabled=true;go.textContent="登入中…";m.textContent="";
      api({action:"login",who:w,code:p}).then(function(j){
        if(!j||!j.ok){m.textContent=(j&&j.error)||"使用人或密碼不正確";return}
        jset(CFG_KEY,{who:w,code:p,url:SYNC_URL,role:j.role||(ADMINS.indexOf(w)>=0?"admin":"user"),at:Date.now()});
        touch();
        hideGate();
        if(window.jiaoliaoSyncRescan)window.jiaoliaoSyncRescan();
      }).catch(function(){m.textContent="連不上網路，請稍後再試"})
        .then(function(){go.disabled=false;go.textContent="登入"});
    };
    setTimeout(function(){(c.who?pw:who).focus()},50);
  }
  function hideGate(){
    if(gate){gate.remove();gate=null}
    document.documentElement.classList.remove("lg-pre");
    document.body.classList.remove("lg-locked");
    paintMenu();
  }

  // 閒置自動登出：畫面鎖住，要重新輸入密碼；資料和雲端同步都不受影響，沒網路時也能用原密碼解鎖
  function idleLock(){
    if(!unlocked())return;
    var c=cfg();c.locked=true;jset(CFG_KEY,c);
    var dlg=document.querySelectorAll("dialog[open]");for(var i=0;i<dlg.length;i++){try{dlg[i].close()}catch(e){}}
    showGate("超過 5 分鐘沒有使用，已自動登出，請重新登入");
  }
  function idleCheck(){if(unlocked()&&Date.now()-lastAct()>IDLE_MS)idleLock()}
  var lastTouch=0;
  function onAct(){if(!unlocked())return;var n=Date.now();if(n-lastTouch>5000){lastTouch=n;touch()}}
  ["pointerdown","keydown","touchstart","wheel","input"].forEach(function(t){document.addEventListener(t,onAct,{capture:true,passive:true})});
  setInterval(idleCheck,10000);
  document.addEventListener("visibilitychange",function(){if(document.visibilityState==="visible")idleCheck()});
  window.addEventListener("pageshow",idleCheck);

  function logout(){
    var c=cfg();
    jset(CFG_KEY,{who:c.who||"",url:SYNC_URL});
    showGate("");
  }

  // 有網路時在背景確認帳號還有效（被刪除或停用就登出）
  function recheck(){
    if(!loggedIn()||!navigator.onLine)return;
    var c=cfg();
    api({action:"login",who:c.who,code:c.code}).then(function(j){
      if(j&&j.ok){if(j.role&&j.role!==c.role){c.role=j.role;jset(CFG_KEY,c);paintMenu()}return}
      var e=String(j&&j.error||"");
      if(/不正確|停用|請輸入/.test(e)){jset(CFG_KEY,{who:c.who,url:SYNC_URL});showGate(/停用/.test(e)?"這個帳號已被停用，請找管理員":"帳號或密碼已變更，請重新登入")}
    }).catch(function(){});
  }

  // ---------- 齒輪選單：目前使用人、登出、使用者管理 ----------
  var bOut,bUsers;
  function paintMenu(){
    var menu=document.querySelector(".tools-menu");if(!menu)return;
    if(!bOut){
      bOut=document.createElement("button");bOut.className="btn";bOut.type="button";
      bOut.onclick=function(){if(confirm("確定要登出嗎？\n（叫料紀錄會留在這台手機，不會刪除）"))logout()};
      bUsers=document.createElement("button");bUsers.className="btn";bUsers.type="button";
      bUsers.innerHTML='👥<span class="lb"> 使用者管理<small>新增、刪除使用人，改密碼（管理員）</small></span>';
      bUsers.onclick=openUsers;
      menu.appendChild(bUsers);menu.appendChild(bOut);
    }
    var c=cfg();
    bOut.innerHTML='🚪<span class="lb"> 登出<small>目前使用人：'+esc(c.who||"")+'</small></span>';
    bUsers.hidden=!isAdmin();bUsers.style.display=isAdmin()?"":"none";
  }

  // ---------- 使用者管理（管理員） ----------
  function openUsers(){
    var tools=document.getElementById("appTools");if(tools)tools.open=false;
    var d=document.createElement("dialog");d.id="userDlg";
    d.innerHTML='<div class="sh-h"><h2 style="margin:0">👥 使用者管理</h2><small>新增的人用「姓名＋密碼」登入。刪除或停用後，對方下次連上網路時就會被登出。</small></div>'+
      '<div class="sh-b"><p id="uMsg" style="font-size:15px">讀取中…</p><div id="uList"></div>'+
      '<div class="add"><input id="uName" placeholder="姓名" autocomplete="off"><input id="uPw" inputmode="tel" placeholder="密碼" autocomplete="off"><button class="btn btn-p" id="uAdd" type="button">＋ 新增使用人</button></div></div>'+
      '<div class="sh-f"><button class="btn" type="button" id="uClose" style="font-size:16px">關閉</button></div>';
    document.body.appendChild(d);
    var msg=d.querySelector("#uMsg"),list=d.querySelector("#uList");
    d.querySelector("#uClose").onclick=function(){d.close()};
    d.addEventListener("close",function(){d.remove()});
    var me=cfg();
    function call(body){body.who=me.who;body.code=me.code;return api(body)}
    function oldBackend(){msg.innerHTML='<b style="color:#dc2626">後台程式還沒更新。</b>請擁有 Google 試算表的人把新版 Code.gs 貼上並重新部署，這裡才能管理使用者。在那之前，可以直接到試算表的「使用者」分頁新增或刪除。'}
    function load(){
      if(!navigator.onLine){msg.textContent="需要網路才能管理使用者";return}
      msg.textContent="讀取中…";
      call({action:"users"}).then(function(j){
        if(!j||!j.ok){msg.textContent=(j&&j.error)||"讀取失敗";return}
        if(!Array.isArray(j.users)){oldBackend();return}
        msg.textContent="共 "+j.users.length+" 位使用人";
        list.innerHTML=j.users.map(function(u,i){
          return '<div class="ur"><div><b>'+esc(u.name)+'</b>'+(u.admin?'<span class="tag">管理員</span>':'')+(u.disabled?'<span class="tag off">停用</span>':'')+'</div>'+
            '<span>密碼：'+esc(u.pw)+'</span>'+
            '<div class="ur-act"><button class="btn" data-a="pw" data-i="'+i+'" type="button">改密碼</button>'+
            (u.name===me.who?'':'<button class="btn" data-a="off" data-i="'+i+'" type="button">'+(u.disabled?'啟用':'停用')+'</button><button class="btn" data-a="del" data-i="'+i+'" type="button" style="color:#dc2626">刪除</button>')+
            '</div></div>';
        }).join("");
        list.onclick=function(ev){
          var b=ev.target.closest("button[data-a]");if(!b)return;
          var u=j.users[+b.dataset.i],a=b.dataset.a;
          if(a==="pw"){var np=prompt("請輸入「"+u.name+"」的新密碼",u.pw);if(np==null||!np.trim())return;save({name:u.name,pw:np.trim(),disabled:u.disabled})}
          else if(a==="off")save({name:u.name,pw:u.pw,disabled:!u.disabled});
          else if(a==="del"&&confirm("確定刪除「"+u.name+"」？對方將無法再登入。"))act({action:"deleteUser",name:u.name});
        };
      }).catch(function(){msg.textContent="連不上網路，請稍後再試"});
    }
    function act(body,after){
      msg.textContent="儲存中…";
      call(body).then(function(j){
        if(!j||!j.ok){msg.textContent=(j&&j.error)||"儲存失敗";return}
        if(!j.done){oldBackend();return}
        if(after)after();
        load();
      }).catch(function(){msg.textContent="連不上網路，請稍後再試"});
    }
    function save(u){act({action:"saveUser",user:u},function(){if(u.name===me.who){me.code=u.pw;jset(CFG_KEY,me)}})}
    d.querySelector("#uAdd").onclick=function(){
      var n=d.querySelector("#uName").value.trim(),p=d.querySelector("#uPw").value.trim();
      if(!n||!p){msg.textContent="請填寫姓名和密碼";return}
      save({name:n,pw:p,disabled:false});
      d.querySelector("#uName").value="";d.querySelector("#uPw").value="";
    };
    d.showModal();load();
  }

  function start(){
    if(unlocked()&&Date.now()-lastAct()>IDLE_MS){idleLock();return}
    if(unlocked()){touch();paintMenu();recheck()}
    else{showGate(loggedIn()?"已自動登出，請重新登入":"");if(loggedIn())recheck()}
  }
  window.addEventListener("online",recheck);
  if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",start);else start();
})();
