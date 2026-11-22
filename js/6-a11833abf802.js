// 雲端同步：叫料單存檔 / 匯出時，自動把訂單送到公司的 Google 試算表做統計。
// 不改動原本的儲存方式；沒網路時先排隊，連上網路後自動補送。
(function(){
  "use strict";
  var KEY="jiaoliao_v1";
  var CFG_KEY="jiaoliao_sync_cfg";
  var SEEN_KEY="jiaoliao_sync_seen";
  var QUEUE_KEY="jiaoliao_sync_queue";
  var DEL_KEY="jiaoliao_sync_delq";
  var SITEQ_KEY="jiaoliao_sync_siteq";
  // 部署 Google Apps Script 後，把網頁應用程式網址貼在這裡
  var SYNC_URL="https://script.google.com/macros/s/AKfycbwf1DwO6ARo_ePfa3TT3ZPecLmNnqjksK3sPxcIN-5_Obn3mHSQerYo1rRzQRJyNtM/exec";

  function jget(k,d){try{var s=localStorage.getItem(k);return s?JSON.parse(s):d}catch(e){return d}}
  var rawSet=Storage.prototype.setItem;
  function jset(k,v){try{rawSet.call(localStorage,k,JSON.stringify(v))}catch(e){}}

  function cfg(){var c=jget(CFG_KEY,{});return{who:c.who||"",code:c.code||"",url:c.url||SYNC_URL}}
  function ready(){var c=cfg();return !!(c.url&&c.code&&c.who)}

  var seen=jget(SEEN_KEY,{});
  var queue=jget(QUEUE_KEY,{});
  var delq=jget(DEL_KEY,{});
  var siteq=jget(SITEQ_KEY,[]);
  var sending=false,timer=null,lastMsg="";

  function pack(p,o){
    return{
      id:o.id,no:o.no,site:o.project||p.name,company:o.company||"",
      orderDate:o.orderDate||"",dueDate:o.dueDate||"",contact:o.contact||"",contactTel:o.contactTel||"",
      exported:!!o.exported,saved:o.saved||0,note:o.note||"",adminRev:o.adminRev||0,
      lines:(o.lines||[]).map(function(l){return{id:l.id||"",vk:l.vk||"",ik:l.ik||"",len:l.len||"",cat:l.cat||"",name:l.name||"",size:l.size==null?"":String(l.size),spec:l.spec||"",unit:l.unit||"",qty:+l.qty||0,remark:l.remark||""}})
    };
  }
  function sig(o){return (o.saved||0)+"|"+(o.exported?1:0)+"|"+(o.lines||[]).length+"|"+(o.adminRev||0)}

  function scan(raw){
    var S;try{S=JSON.parse(raw)}catch(e){return}
    if(!S||!S.projects)return;
    var changed=false;
    S.projects.forEach(function(p){(p.orders||[]).forEach(function(o){
      if(!o||!o.no||!o.id)return;
      var s=sig(o);
      if(seen[o.id]===s)return;
      queue[o.id]={sig:s,data:pack(p,o)};changed=true;
    })});
    if(changed){jset(QUEUE_KEY,queue);flushSoon()}
    paint();
  }

  function flushSoon(){if(timer)clearTimeout(timer);timer=setTimeout(flush,1200)}

  // 叫料人在「紀錄」按 ✕ 刪除自己的單：同步刪掉試算表那筆（沒網路時先排隊）
  window.jiaoliaoSyncDelete=function(id){if(!id)return;delq[id]=1;delete queue[id];delete seen[id];jset(DEL_KEY,delq);jset(QUEUE_KEY,queue);jset(SEEN_KEY,seen);flushSoon()};
  function flushDeletes(){
    var ids=Object.keys(delq);
    if(!ids.length)return Promise.resolve();
    var c=cfg();
    return fetch(c.url,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify({action:"deleteOwn",code:c.code,who:c.who,ids:ids})})
      .then(function(r){return r.json()})
      .then(function(j){if(j&&j.ok&&j.done){ids.forEach(function(id){delete delq[id]});jset(DEL_KEY,delq)}})
      .catch(function(){});
  }

  function flush(){
    if(sending||!ready()||!navigator.onLine)return paint();
    if(Object.keys(delq).length&&!flush.del){flush.del=true;flushDeletes().then(function(){flush.del=false;flush()});return}
    var ids=Object.keys(queue);
    if(!ids.length){lastMsg="";return paint()}
    sending=true;lastMsg="同步中…";paint();
    var c=cfg();
    var batch=ids.slice(0,10).map(function(id){return queue[id]});
    fetch(c.url,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify({code:c.code,who:c.who,orders:batch.map(function(b){return b.data})})})
      .then(function(r){return r.json()})
      .then(function(j){
        if(!j||!j.ok)throw new Error(j&&j.error||"失敗");
        batch.forEach(function(b){seen[b.data.id]=b.sig;if(queue[b.data.id]&&queue[b.data.id].sig===b.sig)delete queue[b.data.id]});
        jset(SEEN_KEY,seen);jset(QUEUE_KEY,queue);lastMsg="已同步";
      })
      .catch(function(e){lastMsg=/姓名|密碼|停用/.test(String(e.message))?"帳號或密碼有誤，請重新登入":"同步失敗，稍後自動重試"})
      .then(function(){sending=false;paint();if(Object.keys(queue).length&&lastMsg==="已同步")flushSoon();else if(Object.keys(queue).length&&lastMsg.indexOf("重新登入")<0)setTimeout(flush,60000)});
  }

  // 攔截原本的存檔：資料寫進 localStorage 時順便檢查有沒有新訂單
  Storage.prototype.setItem=function(k,v){
    rawSet.apply(this,arguments);
    if(k===KEY&&this===window.localStorage){try{scan(v)}catch(e){}}
  };

  // ---------- 介面 ----------
  var chip;
  function paint(){
    if(!chip)return;
    var n=Object.keys(queue).length;
    var t;
    if(!ready())t="☁ 雲端同步：尚未登入";
    else if(lastMsg&&lastMsg!=="已同步")t="☁ "+lastMsg+(n?"（待送 "+n+" 筆）":"");
    else t=n?"☁ 待同步 "+n+" 筆":"☁ 雲端同步：已開啟";
    chip.textContent=t;
  }
  function openDlg(){
    var c=cfg();
    var d=document.createElement("dialog");
    d.id="syncDlg";
    d.innerHTML='<div class="sh-h"><h2>登入雲端同步</h2><small>叫料單存檔或匯出時，自動送到公司 Google 試算表</small></div>'+
      '<form method="dialog" class="sh-b" id="syncForm">'+
      '<label class="f">帳號（你的姓名）<input name="who" value="'+(c.who||"").replace(/"/g,"&quot;")+'" placeholder="例如：王小明" autocomplete="off"></label>'+
      '<label class="f">密碼（預設為你的電話號碼）<input name="code" type="password" inputmode="tel" value="'+(c.code||"").replace(/"/g,"&quot;")+'" placeholder="例如：0912345678" autocomplete="off"></label>'+
      (SYNC_URL?'':'<label class="f">同步網址<input name="url" value="'+(c.url||"").replace(/"/g,"&quot;")+'" placeholder="https://script.google.com/macros/s/…/exec"></label>')+
      '<p style="font-size:12px;color:#6b7280;margin:8px 0 0">只會上傳叫料單內容（工地、日期、品項、數量），不含價格。</p>'+
      '<p id="syncStat" style="font-size:12px;margin:6px 0 0"></p></form>'+
      '<div class="sh-f"><button class="btn" id="syncClose" type="button">關閉</button><button class="btn btn-p" id="syncSave" type="button">儲存並同步</button></div>';
    document.body.appendChild(d);
    d.querySelector("#syncClose").onclick=function(){d.close()};
    d.addEventListener("close",function(){d.remove()});
    d.querySelector("#syncSave").onclick=function(){
      var f=new FormData(d.querySelector("#syncForm"));
      var next={who:String(f.get("who")||"").trim(),code:String(f.get("code")||"").trim(),url:SYNC_URL||String(f.get("url")||"").trim()};
      var stat=d.querySelector("#syncStat");
      if(!next.who||!next.code){stat.textContent="請填寫姓名與密碼";return}
      if(!next.url){stat.textContent="尚未登入同步網址";return}
      var btn=d.querySelector("#syncSave");btn.disabled=true;stat.textContent="驗證中…";
      fetch(next.url,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify({action:"login",who:next.who,code:next.code})})
        .then(function(r){return r.json()})
        .then(function(j){
          if(!j||!j.ok){stat.textContent=(j&&j.error)||"登入失敗";btn.disabled=false;return}
          jset(CFG_KEY,next);
          // 登入成功：把這支手機上已存的叫料單全部補送一次
          var raw=localStorage.getItem(KEY);if(raw)scan(raw);
          lastMsg="";flush();paint();d.close();
        })
        .catch(function(){stat.textContent="連不上網路，請稍後再試";btn.disabled=false});
    };
    d.showModal();
  }
  function mount(){
    var foot=document.querySelector(".foot");if(!foot)return;
    chip=document.createElement("button");
    chip.type="button";
    chip.style.cssText="display:block;margin:8px auto 0;padding:4px 12px;border:1px solid #d1d5db;border-radius:999px;background:#fff;color:#4b5563;font-size:12px";
    chip.onclick=openDlg;
    foot.appendChild(chip);paint();
    pull();
    var raw=localStorage.getItem(KEY);if(raw&&ready())scan(raw);
    flushSoon();
  }
  // ---------- 從後台拿回管理員的修改／刪除，更新這支手機的「紀錄」 ----------
  var pulling=false;
  function post(body){var c=cfg();body.who=c.who;body.code=c.code;return fetch(c.url,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify(body)}).then(function(r){return r.json()})}
  // 手機上改了工地／收貨人：先排隊，有網路就送到試算表「工地資料」「收貨人」
  window.jiaoliaoSiteSave=function(c){if(!c||!c.name)return;siteq.push(c);jset(SITEQ_KEY,siteq);setTimeout(pull,300)};
  function flushSites(){
    if(!siteq.length)return Promise.resolve(true);
    var batch=siteq.slice();
    return post({action:"saveSites",sites:batch}).then(function(j){
      if(j&&j.ok&&j.sites){siteq=siteq.slice(batch.length);jset(SITEQ_KEY,siteq);return true}
      return false}).catch(function(){return false});
  }
  function pull(){
    if(pulling||!ready()||!navigator.onLine)return;
    pulling=true;
    flushSites().then(function(){return post({action:"pull"})})
      .then(function(j){if(j&&j.ok&&Array.isArray(j.orders))applyPull(j)})
      .catch(function(){})
      .then(function(){pulling=false});
  }
  function applyPull(j){
    var raw=localStorage.getItem(KEY);if(!raw)return;
    var S;try{S=JSON.parse(raw)}catch(e){return}
    if(!S||!S.projects)return;
    var dead={};(j.deleted||[]).forEach(function(d){dead[d.id]=1});
    var srv={};j.orders.forEach(function(x){if(x&&x.rev&&x.order)srv[x.id]=x});
    var nEdit=0,nDel=0,touched=[];
    var siteCh=Array.isArray(j.sites)&&j.sites.length?applySites(S,j.sites):false;
    // 管理員在後台「清除所有叫料單」：這支手機在那之前的叫料紀錄和未完成的叫料單也一起清掉
    var nReset=0;
    if(j.resetAt&&j.resetAt>(S.resetDone||0)){
      S.projects.forEach(function(p){
        p.orders=(p.orders||[]).filter(function(o){var t=o.saved||o.created||0;if(t&&t<j.resetAt){nReset++;delete queue[o.id];delete seen[o.id];return false}return true});
        if(p.draft&&(p.draft.created||0)<j.resetAt){p.draft=null;nReset++}
      });
      S.resetDone=j.resetAt;siteCh=true;
    }
    S.projects.forEach(function(p){
      var keep=[];
      (p.orders||[]).forEach(function(o){
        if(dead[o.id]){nDel++;delete queue[o.id];delete seen[o.id];return}
        var x=srv[o.id];
        if(x&&x.rev>(o.adminRev||0)){
          var d=x.order;
          ["orderDate","dueDate","contact","contactTel","company","note"].forEach(function(k){if(d[k]!=null)o[k]=d[k]});
          o.lines=(d.lines||[]).map(function(sl){
            var ll=null;for(var i=0;i<(o.lines||[]).length;i++)if(sl.id&&o.lines[i].id===sl.id){ll=o.lines[i];break}
            var n=Object.assign({},ll||{list:null,priceNote:"",price:null},sl);
            if(!n.id)n.id=Date.now().toString(36)+Math.random().toString(36).slice(2,7);
            if(!n.vk)n.vk="c|"+n.id;if(!n.ik)n.ik="custom";if(!n.len)delete n.len;
            return n});
          o.adminRev=x.rev;o.adminBy=d.adminBy||"";
          nEdit++;touched.push(o);
          // 這張單如果正開在「叫料單」頁，也換成新內容
          if(p.draft&&p.draft.id===o.id)p.draft=JSON.parse(JSON.stringify(o));
        }
        keep.push(o);
      });
      p.orders=keep;
      // 被刪除的單如果還開在「叫料單」頁，改成新的一張，免得再存回去
      if(p.draft&&dead[p.draft.id]){p.draft.id=Date.now().toString(36)+Math.random().toString(36).slice(2,7);p.draft.no="";delete p.draft.adminRev}
    });
    if(!nEdit&&!nDel&&!siteCh)return;
    rawSet.call(localStorage,KEY,JSON.stringify(S));
    touched.forEach(function(o){seen[o.id]=sig(o);delete queue[o.id]});
    jset(SEEN_KEY,seen);jset(QUEUE_KEY,queue);
    if(window.jiaoliaoAppReload)window.jiaoliaoAppReload();
    if(nReset){jset(SEEN_KEY,seen);jset(QUEUE_KEY,queue);lastMsg="管理員已清除舊的叫料紀錄";paint()}
    if(!nEdit&&!nDel){paint();return}
    lastMsg="管理員"+(nEdit?"修改了 "+nEdit+" 張":"")+(nEdit&&nDel?"、":"")+(nDel?"刪除了 "+nDel+" 張":"")+"叫料單，紀錄已更新";
    paint();
  }
  // 試算表「工地資料」「收貨人」→ 這支手機的工地（試算表為準）
  function nid(){return Date.now().toString(36)+Math.random().toString(36).slice(2,7)}
  function applySites(S,sites){
    var before=JSON.stringify([S.projects.map(function(p){return[p.id,p.name,p.sid,p.co,p.addr,p.ppl,p.pplGone,p.off]}),S.cur,S.siteSynced]);
    var pend={};siteq.forEach(function(c){if(c.id)pend[c.id]=1;pend["n:"+c.name]=1;if(c.oldName)pend["n:"+c.oldName]=1});
    var byName={};sites.forEach(function(s){byName[s.name]=s});
    // 第一次接上試算表：這支手機自己加過、試算表沒有的工地／收貨人，補送上去（不會刪試算表的資料）
    if(!S.siteSynced){
      S.projects.forEach(function(p){
        var s=byName[p.name],mine=p.ppl||[];
        if(!s){if(p.orders&&p.orders.length||mine.length||p.addr)siteq.push({name:p.name,co:p.co||"",addr:p.addr||"",add:mine,undel:true});return}
        var add=mine.filter(function(q){return q[0]&&!s.ppl.some(function(x){return x[0]===q[0]})&&!s.gone.some(function(x){return x[0]===q[0]})});
        if(add.length)siteq.push({id:s.id,name:s.name,add:add});
      });
      jset(SITEQ_KEY,siteq);S.siteSynced=true;
      if(siteq.length)setTimeout(pull,1000);
      siteq.forEach(function(c){if(c.id)pend[c.id]=1;pend["n:"+c.name]=1});
    }
    var order={};
    sites.forEach(function(s,i){
      order[s.id]=i;
      var p=null,k;
      for(k=0;k<S.projects.length&&!p;k++)if(S.projects[k].sid===s.id)p=S.projects[k];
      for(k=0;k<S.projects.length&&!p;k++)if(!S.projects[k].sid&&S.projects[k].name===s.name)p=S.projects[k];
      if(pend[s.id]||pend["n:"+s.name]||(p&&pend["n:"+p.name]))return;
      if(s.off){
        if(!p)return;
        var busy=(p.orders&&p.orders.length)||(p.draft&&p.draft.lines&&p.draft.lines.length);
        if(!busy&&S.projects.length>1){S.projects=S.projects.filter(function(z){return z!==p});if(S.cur===p.id)S.cur=S.projects[0].id}
        else{p.sid=s.id;p.off=true}
        return;
      }
      if(!p){p={id:nid(),name:s.name,addr:s.addr,orders:[],draft:null};S.projects.push(p)}
      if(p.name!==s.name){if(p.draft&&p.draft.project===p.name)p.draft.project=s.name;p.name=s.name}
      p.sid=s.id;p.co=s.co;p.addr=s.addr;p.ppl=s.ppl;p.pplGone=s.gone;delete p.off;
    });
    S.projects.sort(function(a,b){var x=a.sid in order?order[a.sid]:1e6,y=b.sid in order?order[b.sid]:1e6;return x-y});
    var after=JSON.stringify([S.projects.map(function(p){return[p.id,p.name,p.sid,p.co,p.addr,p.ppl,p.pplGone,p.off]}),S.cur,S.siteSynced]);
    return before!==after;
  }
  window.jiaoliaoSyncPull=pull;
  setInterval(function(){if(document.visibilityState==="visible")pull()},180000);
  document.addEventListener("visibilitychange",function(){if(document.visibilityState==="visible")pull()});

  // 登入畫面（js/7-login.js）登入成功後呼叫：把這支手機已存的叫料單補送一次
  window.jiaoliaoSyncRescan=function(){pull();var raw=localStorage.getItem(KEY);if(raw)scan(raw);lastMsg="";flushSoon();paint()};
  window.addEventListener("online",function(){pull();flushSoon()});
  if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",mount);else mount();
})();
