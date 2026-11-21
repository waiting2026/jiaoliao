// 管理員：查看、修改、刪除所有人的叫料單。
// 存檔後寫進 Google 試算表，叫料人的手機下次連上網路時，「紀錄」會自動換成修改後的內容（刪除的會消失）。
(function(){
  "use strict";
  var CFG_KEY="jiaoliao_sync_cfg";
  var ADMINS=["李羿萱"];
  function cfg(){try{return JSON.parse(localStorage.getItem(CFG_KEY)||"{}")}catch(e){return{}}}
  function isAdmin(){var c=cfg();return !!(c.who&&c.code&&!c.locked)&&(c.role==="admin"||ADMINS.indexOf(c.who)>=0)}
  function esc(s){return String(s==null?"":s).replace(/[&<>"]/g,function(ch){return{"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[ch]})}
  function uid(){return Date.now().toString(36)+Math.random().toString(36).slice(2,7)}
  function api(body){
    var c=cfg();body.who=c.who;body.code=c.code;
    return fetch(c.url,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify(body)}).then(function(r){return r.json()});
  }

  var css=document.createElement("style");
  css.textContent=
    "#admDlg{max-width:720px}#admDlg h2{margin:0;font-size:20px}"+
    "#admDlg .ad-f{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:8px}#admDlg .ad-f select,#admDlg .ad-f input{font-size:16px;padding:8px;border:1px solid #d1d5db;border-radius:10px;flex:1;min-width:120px;background:#fff}"+
    "#admDlg .ao{border:1px solid #e5e7eb;border-radius:12px;padding:10px 12px;margin:8px 0;display:flex;gap:8px;align-items:center}"+
    "#admDlg .ao div{flex:1;min-width:0}#admDlg .ao b{font-size:16px}#admDlg .ao small{display:block;color:#6b7280;font-size:14px;margin-top:2px}"+
    "#admDlg .tag{font-size:12px;border-radius:999px;padding:1px 8px;margin-left:4px;background:#fef3c7;color:#92400e}"+
    "#admDlg .btn{font-size:15px;padding:8px 12px}"+
    "#admDlg .ae label{display:block;font-size:15px;color:#4b5563;margin:8px 0 4px}#admDlg .ae input,#admDlg .ae textarea{width:100%;font-size:17px;padding:9px;border:1px solid #d1d5db;border-radius:10px;background:#fff}"+
    "#admDlg .al{border-bottom:1px solid #e5e7eb;padding:8px 0}#admDlg .al .nm{font-size:16px;font-weight:600}#admDlg .al .nm small{font-weight:400;color:#6b7280}"+
    "#admDlg .al .row{display:grid;grid-template-columns:90px 1fr auto;gap:6px;align-items:center;margin-top:6px}"+
    "#admDlg .al .row input{font-size:17px;padding:8px;border:1px solid #d1d5db;border-radius:10px;width:100%}"+
    "#admDlg .add{display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-top:10px;background:#f9fafb;border-radius:12px;padding:10px}"+
    "#admDlg .add input{font-size:16px;padding:8px;border:1px solid #d1d5db;border-radius:10px;width:100%}#admDlg .add .btn{grid-column:1/-1}"+
    ".adm-card{background:#fffbeb;border:1px solid #fcd34d;border-radius:14px;padding:12px;margin:10px 0;display:flex;align-items:center;gap:10px}"+
    ".adm-card p{margin:0;flex:1;font-size:15px}";
  document.head.appendChild(css);

  var all=[];
  function open(){
    var tools=document.getElementById("appTools");if(tools)tools.open=false;
    var d=document.createElement("dialog");d.id="admDlg";
    document.body.appendChild(d);
    d.addEventListener("close",function(){d.remove()});
    d.addEventListener("cancel",function(e){if(d.dataset.view==="edit"){e.preventDefault();list()}});
    function frame(title,sub,body,foot){
      d.innerHTML='<div class="sh-h"><h2>'+title+'</h2>'+(sub?'<small>'+sub+'</small>':'')+'</div><div class="sh-b">'+body+'</div><div class="sh-f">'+foot+'</div>';
    }
    function load(){
      d.dataset.view="list";
      frame("📋 全部叫料單（管理員）","",'<p id="aMsg" style="font-size:16px">讀取中…</p>','<button class="btn" type="button" id="aClose">關閉</button>');
      d.querySelector("#aClose").onclick=function(){d.close()};
      if(!navigator.onLine){d.querySelector("#aMsg").textContent="需要網路才能管理叫料單";return}
      api({action:"pull",all:true}).then(function(j){
        if(!j||!j.ok){d.querySelector("#aMsg").textContent=(j&&j.error)||"讀取失敗";return}
        if(!Array.isArray(j.orders)){d.querySelector("#aMsg").innerHTML='<b style="color:#dc2626">後台程式還沒更新。</b>請擁有 Google 試算表的人貼上新版 Code.gs 並重新部署，才能在這裡管理叫料單。';return}
        all=j.orders.slice().sort(function(a,b){return String(b.order.orderDate||"").localeCompare(String(a.order.orderDate||""))||String(b.order.no||"").localeCompare(String(a.order.no||""))});
        list();
      }).catch(function(){d.querySelector("#aMsg").textContent="連不上網路，請稍後再試"});
    }
    var fSite="",fQ="";
    function list(){
      d.dataset.view="list";
      var sites=[];all.forEach(function(x){var s=x.order.site||"";if(sites.indexOf(s)<0)sites.push(s)});
      frame("📋 全部叫料單（管理員）","修改或刪除後，叫料人的手機連上網路時「紀錄」會自動更新。",
        '<div class="ad-f"><select id="aSite"><option value="">全部工地</option>'+sites.map(function(s){return'<option'+(s===fSite?' selected':'')+'>'+esc(s)+'</option>'}).join("")+'</select><input id="aQ" placeholder="搜尋單號、叫料人、品名" value="'+esc(fQ)+'"></div><div id="aList"></div>',
        '<button class="btn" type="button" id="aReload">重新整理</button><button class="btn" type="button" id="aClose">關閉</button>');
      d.querySelector("#aClose").onclick=function(){d.close()};
      d.querySelector("#aReload").onclick=load;
      var box=d.querySelector("#aList");
      function paint(){
        var q=fQ.trim().toLowerCase();
        var rows=all.filter(function(x){var o=x.order;if(fSite&&(o.site||"")!==fSite)return false;if(!q)return true;
          return [o.no,x.who,o.site].concat((o.lines||[]).map(function(l){return l.name})).join(" ").toLowerCase().indexOf(q)>=0});
        box.innerHTML=rows.length?rows.map(function(x){var o=x.order;var i=all.indexOf(x);
          return '<div class="ao"><div><b>'+esc(o.no||"（未編號）")+'</b>'+(x.rev?'<span class="tag">已修改</span>':'')+'<small>'+esc(o.site)+'・叫料人 '+esc(x.who||"")+'</small><small>訂貨 '+esc(o.orderDate)+'・'+(o.lines||[]).length+' 項</small></div>'+
            '<button class="btn" type="button" data-e="'+i+'">修改</button><button class="btn" type="button" data-x="'+i+'" style="color:#dc2626">刪除</button></div>'}).join(""):'<p style="font-size:16px;color:#6b7280">沒有叫料單</p>';
      }
      d.querySelector("#aSite").onchange=function(){fSite=this.value;paint()};
      d.querySelector("#aQ").oninput=function(){fQ=this.value;paint()};
      box.onclick=function(e){
        var b=e.target.closest("button");if(!b)return;
        if(b.dataset.e)edit(all[+b.dataset.e]);
        if(b.dataset.x){var x=all[+b.dataset.x];
          if(!confirm("確定刪除叫料單 "+(x.order.no||"")+"（"+(x.order.site||"")+"，"+(x.who||"")+"）？\n叫料人手機裡的這張紀錄也會一起刪除。"))return;
          b.disabled=true;
          api({action:"adminDeleteOrder",id:x.id,owner:x.who}).then(function(j){
            if(!j||!j.ok||!j.done){alert((j&&j.error)||"刪除失敗（後台程式可能還沒更新）");b.disabled=false;return}
            all.splice(all.indexOf(x),1);paint();
            if(window.jiaoliaoSyncPull)window.jiaoliaoSyncPull();
          }).catch(function(){alert("連不上網路，請稍後再試");b.disabled=false});
        }
      };
      paint();
    }
    function edit(x){
      d.dataset.view="edit";
      var o=JSON.parse(JSON.stringify(x.order));o.lines=o.lines||[];
      function body(){
        return '<div class="ae"><label>訂貨日</label><input type="date" id="eOd" value="'+esc(o.orderDate)+'">'+
          '<label>交貨日</label><input type="date" id="eDd" value="'+esc(o.dueDate)+'">'+
          '<label>備註</label><textarea id="eNote" rows="2">'+esc(o.note)+'</textarea></div>'+
          '<h3 style="margin:16px 0 4px;font-size:17px">品項（'+o.lines.length+'）</h3><div id="eLines">'+
          o.lines.map(function(l,i){return '<div class="al"><div class="nm">'+esc(l.name)+' <small>'+esc([l.size,l.spec].filter(Boolean).join(" "))+'</small></div>'+
            '<div class="row"><input type="number" inputmode="decimal" min="0" step="any" data-q="'+i+'" value="'+esc(l.qty)+'"><span style="font-size:16px">'+esc(l.unit)+'</span><button class="btn" type="button" data-r="'+i+'" style="color:#dc2626">✕ 移除</button></div>'+
            '<div class="row" style="grid-template-columns:1fr"><input data-m="'+i+'" placeholder="備註" value="'+esc(l.remark)+'"></div></div>'}).join("")+
          '</div><div class="add"><input id="nName" placeholder="品名"><input id="nSize" placeholder="尺寸"><input id="nUnit" placeholder="單位（個、支…）"><input id="nQty" type="number" inputmode="decimal" placeholder="數量"><button class="btn btn-p" type="button" id="nAdd">＋ 新增品項</button></div>';
      }
      function grab(){
        o.orderDate=d.querySelector("#eOd").value;o.dueDate=d.querySelector("#eDd").value;o.note=d.querySelector("#eNote").value;
        d.querySelectorAll("[data-q]").forEach(function(inp){o.lines[+inp.dataset.q].qty=+inp.value||0});
        d.querySelectorAll("[data-m]").forEach(function(inp){o.lines[+inp.dataset.m].remark=inp.value});
      }
      function render(){
        frame("修改 "+esc(o.no||""),esc(o.site)+"・叫料人 "+esc(x.who||""),body(),'<button class="btn" type="button" id="eBack">返回</button><button class="btn btn-p" type="button" id="eSave">儲存修改</button>');
        d.querySelector("#eBack").onclick=list;
        d.querySelector("#eLines").onclick=function(e){var b=e.target.closest("[data-r]");if(!b)return;grab();o.lines.splice(+b.dataset.r,1);render()};
        d.querySelector("#nAdd").onclick=function(){
          var n=d.querySelector("#nName").value.trim(),q=+d.querySelector("#nQty").value||0;
          if(!n||q<=0){alert("請填寫品名和數量");return}
          grab();var id=uid();
          o.lines.push({id:id,vk:"c|"+id,ik:"custom",cat:"自訂",name:n,size:d.querySelector("#nSize").value.trim(),spec:"",unit:d.querySelector("#nUnit").value.trim()||"個",qty:q,remark:"管理員新增"});
          render();
        };
        d.querySelector("#eSave").onclick=function(){
          grab();
          o.lines=o.lines.filter(function(l){return (+l.qty||0)>0});
          if(!o.lines.length&&!confirm("這張單已經沒有品項，確定要存嗎？（要整張刪除請用「刪除」）"))return;
          var btn=this;btn.disabled=true;btn.textContent="儲存中…";
          api({action:"adminSaveOrder",order:o,owner:x.who}).then(function(j){
            if(!j||!j.ok||!j.done){alert((j&&j.error)||"儲存失敗（後台程式可能還沒更新）");btn.disabled=false;btn.textContent="儲存修改";return}
            o.adminRev=j.rev;x.order=o;x.rev=j.rev;
            if(window.jiaoliaoSyncPull)window.jiaoliaoSyncPull();
            list();
          }).catch(function(){alert("連不上網路，請稍後再試");btn.disabled=false;btn.textContent="儲存修改"});
        };
      }
      render();
    }
    d.showModal();load();
  }

  // 齒輪選單按鈕
  var bAll;
  function paintMenu(){
    var menu=document.querySelector(".tools-menu");if(!menu)return;
    if(!bAll){bAll=document.createElement("button");bAll.className="btn";bAll.type="button";
      bAll.innerHTML='📋<span class="lb"> 全部叫料單<small>修改、刪除所有人的叫料單（管理員）</small></span>';
      bAll.onclick=open;menu.insertBefore(bAll,menu.firstChild)}
    bAll.hidden=!isAdmin();bAll.style.display=isAdmin()?"":"none";
  }
  // 「紀錄」頁最上面也放一個入口
  function paintHist(){
    var v=document.getElementById("vHist");if(!v)return;
    var card=v.querySelector(".adm-card");
    if(!isAdmin()){if(card)card.remove();return}
    if(card)return;
    card=document.createElement("div");card.className="adm-card";
    card.innerHTML='<p><b>管理員</b>：查看、修改、刪除所有人的叫料單</p><button class="btn btn-p" type="button">📋 全部叫料單</button>';
    card.querySelector("button").onclick=open;
    v.insertBefore(card,v.firstChild);
  }
  function start(){
    paintMenu();paintHist();
    var v=document.getElementById("vHist");
    if(v)new MutationObserver(paintHist).observe(v,{childList:true});
    // 登入／登出後更新
    window.addEventListener("storage",function(){paintMenu();paintHist()});
    setInterval(function(){paintMenu();paintHist()},2000);
  }
  if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",start);else start();
})();
