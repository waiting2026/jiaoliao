/**
 * 工程叫料小幫手 — Google 試算表後台
 *
 * 用法：在要當後台的 Google 試算表選「擴充功能 → Apps Script」，貼上這份程式碼，
 * 1) 執行一次 setup（建立「訂單」「明細」「統計」「使用者」分頁），
 * 2) 到「使用者」分頁填入帳號：A 欄＝姓名（登入帳號）、B 欄＝密碼（可以是電話號碼或任何文字）、C 欄填「停用」可暫停某人、D 欄填「管理員」可管理使用人，
 *    （管理員也可以直接在網站齒輪 →「使用者管理」新增、刪除、改密碼）
 * 3) 「部署 → 新增部署作業 → 網頁應用程式」：執行身分選「我」，誰可以存取選「所有人」，
 * 4) 把部署後的網址貼到 js/6-sync.js 的 SYNC_URL。
 */
// 初次建立「使用者」分頁時帶入的名單，格式：[["姓名","電話"], ...]；之後直接在試算表裡改即可。
var SEED_USERS = [];
// 管理員名單（可在網站上新增、刪除使用人）；試算表 D 欄填「管理員」也算。
var ADMINS = ["李羿萱", "蔡鴻基"];
// 若程式是從試算表「擴充功能 → Apps Script」開啟，這行留空即可；
// 若是獨立的 Apps Script 專案，填入試算表網址中 /d/ 與 /edit 之間那一串 ID。
var SHEET_ID = "";

function book_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive();
}

var SH_ORDER = "訂單";
var SH_LINE = "明細";
var SH_STAT = "統計";
var SH_USER = "使用者";
var USER_HEAD = ["姓名（帳號）", "密碼", "停用", "管理員"];
var ORDER_HEAD = ["訂單ID", "單號", "工地", "公司", "訂貨日", "交貨日", "聯絡人", "電話", "品項數", "已匯出", "叫料人", "儲存時間", "備註", "月份", "管理員修改", "資料（程式用，請勿修改）", "刪除（打勾）"];
var SH_DEL = "已刪除";
var DEL_HEAD = ["訂單ID", "叫料人", "刪除時間", "刪除人", "單號", "工地"];
var COL_REV = 15, COL_DATA = 16;
var LINE_HEAD = ["訂單ID", "單號", "工地", "訂貨日", "分類", "品名", "尺寸", "規格說明", "單位", "數量", "備註", "叫料人", "月份", "刪除（打勾）", "明細ID（程式用）", "到貨狀況", "交貨廠商"];
var COL_ODEL = 17, LCOL_DEL = 14, LCOL_ID = 15;

function setup() {
  var ss = book_();
  ensureSheet_(ss, SH_ORDER, ORDER_HEAD);
  lineSheet_(ss);
  var us = ensureSheet_(ss, SH_USER, USER_HEAD);
  if (us.getLastRow() < 2 && SEED_USERS.length) {
    us.getRange(2, 1, SEED_USERS.length, 2).setValues(SEED_USERS);
  }
  us.setColumnWidths(1, 2, 150);
  var st = ss.getSheetByName(SH_STAT) || ss.insertSheet(SH_STAT);
  st.clear();
  st.getRange("A1").setValue("工程叫料統計（自動更新，請勿手動修改）").setFontWeight("bold").setFontSize(14);
  st.getRange("A3").setValue("各工地").setFontWeight("bold");
  st.getRange("A4").setFormula('=IFERROR(QUERY(' + SH_ORDER + '!A2:N,"select C, count(A), sum(I) where A is not null group by C order by count(A) desc label C \'工地\', count(A) \'叫料單數\', sum(I) \'品項數\'",0),"")');
  st.getRange("E3").setValue("每月叫料單數").setFontWeight("bold");
  st.getRange("E4").setFormula('=IFERROR(QUERY(' + SH_ORDER + '!A2:N,"select N, count(A) where A is not null group by N order by N desc label N \'月份\', count(A) \'叫料單數\'",0),"")');
  st.getRange("H3").setValue("各材料累計數量").setFontWeight("bold");
  st.getRange("H4").setFormula('=IFERROR(QUERY(' + SH_LINE + '!A2:M,"select E, F, G, I, sum(J) where A is not null group by E, F, G, I order by sum(J) desc label E \'分類\', F \'品名\', G \'尺寸\', I \'單位\', sum(J) \'總數量\'",0),"")');
  st.getRange("M3").setValue("各叫料人").setFontWeight("bold");
  st.getRange("M4").setFormula('=IFERROR(QUERY(' + SH_ORDER + '!A2:N,"select K, count(A) where A is not null group by K order by count(A) desc label K \'叫料人\', count(A) \'叫料單數\'",0),"")');
  st.setColumnWidths(1, 1, 120);
  st.setColumnWidths(8, 1, 100);
  st.setColumnWidths(9, 1, 220);
  var first = ss.getSheetByName("工作表1") || ss.getSheetByName("Sheet1");
  if (first && ss.getSheets().length > 1) ss.deleteSheet(first);
  syncSiteSheets_(ss);
  ss.setActiveSheet(st);
}

function ensureSheet_(ss, name, head) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold").setBackground("#e8f0e4");
    sh.setFrozenRows(1);
    // 日期、電話、尺寸（如 1/2"）這幾欄設成純文字，避免被轉成日期或數字
    var textCols = name === SH_ORDER ? [5, 6, 8] : name === SH_USER ? [2] : [4, 7];
    textCols.forEach(function (c) { sh.getRange(2, c, sh.getMaxRows() - 1, 1).setNumberFormat("@"); });
  }
  return sh;
}

// ---------- 工地資料（工地、公司抬頭、地址、收貨人）----------
// 「工地資料」一列一個工地；「收貨人」一列一個人，同一工地可以很多人（手機上會變成下拉選單）。
// 不要整列刪除：在「停用／刪除」欄打勾就等於刪除，取消打勾就恢復。
var SH_SITE = "工地資料";
var SITE_HEAD = ["工地名稱", "公司抬頭", "交貨地址", "停用（打勾＝刪除，取消＝恢復）", "工地ID（程式用，請勿修改）", "最後修改"];
var SH_PPL = "收貨人";
var PPL_HEAD = ["工地", "收貨人", "電話", "刪除（打勾＝刪除，取消＝恢復）", "最後修改"];
var COMPANIES = ["聖和工程有限公司", "紘誠水電工程行", "澧育有限公司", "白鐵工程行"];
var SEED_SITES = [["公司", "", "新北市新莊區民安路398號", []], ["五股倉庫", "", "新北市五股區民義路一段312號-16B", [["陳清榮", "0383-969-599"]]], ["宜蘭山河", "澧育有限公司", "宜蘭縣五結鄉親河路一段228巷26號旁工地", [["林顯呈", "0937-162-938"], ["陳豐乙", "0961-151-005"], ["曾建智", "0930-891-963"], ["王耀德", "0979-890-000"]]], ["新店大廣", "紘誠水電工程行", "新北市新店區三民路94巷6號對面", [["楊志華", "0960-119-729"]]], ["敦北飯店", "紘誠水電工程行", "台北市敦化北路242號", [["高清華", "0933-066-337"]]], ["南山A21", "紘誠水電工程行", "台北市信義區松壽路6號（大門開在市府路上5號門-互助營造）", [["林志雄", "0981-575-553"]]], ["國賓飯店", "紘誠水電工程行", "台北市中山區中山北路二段65巷9號", [["李弘溢", "0926-283-762"], ["陳致豪", "0983-822-193"]]], ["南港冠德", "澧育有限公司", "台北市南港區合成里忠孝東路六段39號", [["陳致豪", "0983-822-193"], ["李弘溢", "0926-283-762"]]], ["南海段", "聖和工程有限公司", "台北市中正區泉州街9巷12號旁工地", [["溫祥淵", "0987-419-519"]]], ["內湖緯創", "聖和工程有限公司", "台北市內湖區新湖一路191號", [["李羿萱", "0960-905-782"]]], ["竹北緯創", "聖和工程有限公司", "竹北市智慧路1號", [["游伸鋒", "0955-895-658"]]], ["竹北聯發廠", "聖和工程有限公司", "竹北市十興路291號", [["毛亮友", "0903-579-317"]]], ["湖口緯創", "聖和工程有限公司", "新竹縣湖口鄉光復北路50號", [["蔡永元", "0925-763-735"]]], ["湖口緯創 N1F", "紘誠水電工程行", "新竹縣湖口鄉光復北路50號", [["蔡永元", "0925-763-735"]]], ["湖口緯創 N1J7冰機機房", "澧育有限公司", "新竹縣湖口鄉光復北路50號", [["蔡永元", "0925-763-735"]]], ["高雄緯創", "聖和工程有限公司", "高雄市前鎮區東七街2號（高雄前鎮科技產業園區）", [["游伸鋒", "0955-895-658"], ["蔡永元", "0925-763-735"], ["覃輝豐", "0936-757-907"]]], ["龜山微星", "聖和工程有限公司", "龜山區中興路2號", [["李重陽", "0978-723-297"], ["陳柏竹", "0970-522-793"]]], ["觀音聯華", "白鐵工程行", "桃園市觀音區工業一路16之1號", [["蔡永元", "0925-763-735"]]], ["三重捷六", "聖和工程有限公司", "三重區捷運路59號對面工地", []], ["五谷王(峰左岸)", "聖和工程有限公司", "新北市三重區(疏洪西路與興德路口)", [["老闆", "0913-515-179"]]], ["青埔B基地", "紘誠水電工程行", "中壢區領航北路三段青昇路交叉口", [["覃輝豐", "0936-757-907"]]]];

function siteSheets_(ss) {
  var sh = ss.getSheetByName(SH_SITE), ph = ss.getSheetByName(SH_PPL);
  if (sh && ph) return { site: sh, ppl: ph };
  sh = ensureSheet_(ss, SH_SITE, SITE_HEAD);
  ph = ensureSheet_(ss, SH_PPL, PPL_HEAD);
  sh.getRange(2, 3, sh.getMaxRows() - 1, 1).setNumberFormat("@");
  ph.getRange(2, 3, ph.getMaxRows() - 1, 1).setNumberFormat("@");
  if (sh.getLastRow() < 2) {
    var now = stamp_();
    var rows = SEED_SITES.map(function (x) { return [x[0], x[1], x[2], false, newId_(), now]; });
    sh.getRange(2, 1, rows.length, SITE_HEAD.length).setValues(rows);
    var pr = [];
    SEED_SITES.forEach(function (x) { x[3].forEach(function (q) { pr.push([x[0], q[0], q[1], false, now]); }); });
    if (pr.length) ph.getRange(2, 1, pr.length, PPL_HEAD.length).setValues(pr);
  }
  siteRules_(ss);
  sh.setColumnWidth(1, 170); sh.setColumnWidth(2, 150); sh.setColumnWidth(3, 320); sh.setColumnWidth(4, 120); sh.setColumnWidth(5, 110);
  ph.setColumnWidth(1, 170); ph.setColumnWidth(2, 110); ph.setColumnWidth(3, 130); ph.setColumnWidth(4, 120);
  return { site: sh, ppl: ph };
}

// 下拉選單與打勾框
function siteRules_(ss) {
  var sh = ss.getSheetByName(SH_SITE), ph = ss.getSheetByName(SH_PPL);
  var n1 = sh.getMaxRows() - 1, n2 = ph.getMaxRows() - 1;
  sh.getRange(2, 2, n1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(COMPANIES, true).setAllowInvalid(true).build());
  if (sh.getLastRow() > 1) sh.getRange(2, 4, sh.getLastRow() - 1, 1).insertCheckboxes();
  ph.getRange(2, 1, n2, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(sh.getRange("A2:A"), true).setAllowInvalid(false).setHelpText("請從工地清單選").build());
  if (ph.getLastRow() > 1) ph.getRange(2, 4, ph.getLastRow() - 1, 1).insertCheckboxes();
}

function stamp_() { return Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm"); }
function newId_() { return "s" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }

// 給手機：所有工地（含停用的），每個工地的收貨人與已刪除的收貨人
function getSites_(ss) {
  var t = siteSheets_(ss);
  var srows = t.site.getLastRow() > 1 ? t.site.getRange(2, 1, t.site.getLastRow() - 1, SITE_HEAD.length).getValues() : [];
  var prows = t.ppl.getLastRow() > 1 ? t.ppl.getRange(2, 1, t.ppl.getLastRow() - 1, PPL_HEAD.length).getValues() : [];
  var fix = [];
  var out = [], byName = {};
  srows.forEach(function (r, i) {
    var name = String(r[0] || "").trim();
    if (!name) return;
    var id = String(r[4] || "");
    if (!id) { id = newId_(); fix.push([i + 2, id]); }
    var s = { id: id, name: name, co: String(r[1] || "").trim(), addr: String(r[2] || "").trim(), off: r[3] === true, ppl: [], gone: [] };
    out.push(s); byName[name] = s;
  });
  fix.forEach(function (f) { t.site.getRange(f[0], 5).setValue(f[1]); });
  prows.forEach(function (r) {
    var s = byName[String(r[0] || "").trim()], nm = String(r[1] || "").trim();
    if (!s || !nm) return;
    var q = [nm, String(r[2] || "").trim()];
    (r[3] === true ? s.gone : s.ppl).push(q);
  });
  out.forEach(function (s) { s.gone = s.gone.filter(function (g) { return !s.ppl.some(function (q) { return q[0] === g[0]; }); }); });
  return out;
}

// 手機上改工地：{id, oldName, name, co, addr, add:[[姓名,電話]], remove:[姓名], del:true/false, undel:true}
// 只改有變動的地方，收貨人刪除＝打勾（可恢復），不會整列刪掉。
function saveSites_(ss, list, who) {
  var t = siteSheets_(ss), sh = t.site, ph = t.ppl, now = stamp_();
  (list || []).forEach(function (c) {
    var name = String(c.name || "").trim();
    if (!name) return;
    var n = sh.getLastRow() - 1;
    var rows = n > 0 ? sh.getRange(2, 1, n, SITE_HEAD.length).getValues() : [];
    var at = -1;
    for (var i = 0; i < rows.length && at < 0; i++) if (c.id && String(rows[i][4]) === c.id) at = i;
    for (i = 0; i < rows.length && at < 0; i++) if (String(rows[i][0]).trim() === String(c.oldName || name).trim()) at = i;
    for (i = 0; i < rows.length && at < 0; i++) if (String(rows[i][0]).trim() === name) at = i;
    var old = at >= 0 ? String(rows[at][0]).trim() : name;
    if (at < 0) {
      sh.appendRow([name, c.co || "", c.addr || "", !!c.del, c.id || newId_(), now]);
      sh.getRange(sh.getLastRow(), 4).insertCheckboxes();
    } else {
      var r = rows[at];
      var nv = [name, c.co != null ? c.co : r[1], c.addr != null ? c.addr : r[2], c.del ? true : (c.undel ? false : r[3]), r[4] || c.id || newId_(), now];
      sh.getRange(at + 2, 1, 1, SITE_HEAD.length).setValues([nv]);
      if (old !== name) renamePpl_(ph, old, name);
    }
    var m = ph.getLastRow() - 1;
    var pr = m > 0 ? ph.getRange(2, 1, m, PPL_HEAD.length).getValues() : [];
    (c.remove || []).forEach(function (nm) {
      for (var k = 0; k < pr.length; k++) if (String(pr[k][0]).trim() === name && String(pr[k][1]).trim() === nm && pr[k][3] !== true) {
        ph.getRange(k + 2, 4, 1, 2).setValues([[true, now + " " + who]]); pr[k][3] = true;
      }
    });
    (c.add || []).forEach(function (q) {
      var nm = String(q[0] || "").trim(), tel = String(q[1] || "").trim();
      if (!nm) return;
      var hit = -1;
      for (var k = 0; k < pr.length; k++) if (String(pr[k][0]).trim() === name && String(pr[k][1]).trim() === nm) { if (hit < 0 || pr[k][3] !== true) hit = k; }
      if (hit >= 0) {
        ph.getRange(hit + 2, 3, 1, 3).setValues([[tel || pr[hit][2], false, now + " " + who]]); pr[hit][3] = false;
      } else {
        ph.appendRow([name, nm, tel, false, now + " " + who]);
        ph.getRange(ph.getLastRow(), 4).insertCheckboxes();
        pr.push([name, nm, tel, false, now]);
      }
    });
  });
  return getSites_(ss);
}

function renamePpl_(ph, old, name) {
  var m = ph.getLastRow() - 1;
  if (m < 1) return;
  var a = ph.getRange(2, 1, m, 1).getValues(), ch = false;
  a.forEach(function (r) { if (String(r[0]).trim() === old) { r[0] = name; ch = true; } });
  if (ch) ph.getRange(2, 1, m, 1).setValues(a);
}

function doGet() {
  return json_({ ok: true, name: "工程叫料小幫手後台" });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var body = JSON.parse(e.postData.contents);
    var who = String((body && body.who) || "").trim();
    var auth = checkUser_(who, (body && body.code) || "");
    if (!auth.ok) return json_({ ok: false, error: auth.error });
    if (body.action === "login") return json_({ ok: true, name: who, role: auth.admin ? "admin" : "user" });
    if (body.action === "users" || body.action === "saveUser" || body.action === "deleteUser") {
      if (!auth.admin) return json_({ ok: false, error: "只有管理員可以管理使用人" });
      lock.waitLock(25000);
      return json_(manageUsers_(body, who));
    }
    if (body.action === "sites") return json_({ ok: true, sites: getSites_(book_()) });
    // 電腦版：廠商名單、案場採購紀錄（只有管理員）
    if (body.action === "table" || body.action === "saveRow") {
      if (!auth.admin) return json_({ ok: false, error: "只有管理員可以看廠商資料" });
      if (body.action === "saveRow") lock.waitLock(25000);
      return json_(body.action === "table" ? getTable_(body.name) : saveRow_(body, who));
    }
    // 電腦版：出貨核對（照片存雲端硬碟、照片判讀；只有管理員）
    // 電腦版：價目表（只有管理員）
    if (/^price/.test(String(body.action))) {
      if (!auth.admin) return json_({ ok: false, error: "只有管理員可以看價目表" });
      if (body.action !== "priceList" && body.action !== "priceGetFile" && body.action !== "priceFile") lock.waitLock(25000);
      return json_(price_(body, who));
    }
    if (/^ship/.test(String(body.action))) {
      if (!auth.admin) return json_({ ok: false, error: "只有管理員可以做出貨核對" });
      if (body.action !== "shipList" && body.action !== "shipImg") lock.waitLock(25000);
      return json_(ship_(body, who));
    }
    if (body.action === "saveSites") {
      lock.waitLock(25000);
      return json_({ ok: true, sites: saveSites_(book_(), body.sites || [], who) });
    }
    if (body.action === "pull") return json_(pull_(who, auth.admin, body.all));
    if (body.action === "deleteOwn") {
      lock.waitLock(25000);
      return json_(deleteOwn_(body.ids || [], who, auth.admin));
    }
    if (body.action === "adminSaveOrder" || body.action === "adminDeleteOrder") {
      if (!auth.admin) return json_({ ok: false, error: "只有管理員可以修改別人的叫料單" });
      lock.waitLock(25000);
      return json_(adminOrder_(body, who));
    }
    lock.waitLock(25000);
    var ss = book_();
    var so = orderSheet_(ss);
    var sl = lineSheet_(ss);
    var dead = deletedIds_(ss);
    var orders = body.orders || [];
    var skipped = 0;
    for (var i = 0; i < orders.length; i++) {
      var o = orders[i];
      if (!o || !o.id || dead[o.id]) { skipped++; continue; }
      var cur = findOrder_(so, o.id);
      // 管理員改過的單：手機上的舊版本不能蓋掉管理員的修改
      if (cur && cur.rev && Number(o.adminRev || 0) < cur.rev) { skipped++; continue; }
      upsert_(so, sl, o, cur ? cur.who || who : who, cur ? cur.rev : 0);
    }
    return json_({ ok: true, n: orders.length - skipped, skipped: skipped });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function digits_(v) {
  return String(v == null ? "" : v).replace(/\D/g, "");
}

// 密碼相同才算對；若試算表裡的密碼是電話號碼，有沒有打破折號都可以
function samePw_(stored, pw) {
  var a = String(stored == null ? "" : stored).trim(), b = String(pw == null ? "" : pw).trim();
  if (!a || !b) return false;
  if (a === b) return true;
  return /^[\d\s\-()+]+$/.test(a) && /^[\d\s\-()+]+$/.test(b) && digits_(a) === digits_(b);
}

function isAdmin_(name, flag) {
  return ADMINS.indexOf(name) >= 0 || String(flag || "").trim() !== "";
}

function userSheet_() {
  var sh = ensureSheet_(book_(), SH_USER, USER_HEAD);
  if (!String(sh.getRange(1, 4).getValue()).trim()) sh.getRange(1, 4).setValue(USER_HEAD[3]).setFontWeight("bold").setBackground("#e8f0e4");
  return sh;
}

// 管理員：列出、新增／修改、刪除使用人
function manageUsers_(body, me) {
  var sh = userSheet_();
  var n = sh.getLastRow() - 1;
  var rows = n > 0 ? sh.getRange(2, 1, n, 4).getValues() : [];
  function find(name) { for (var i = 0; i < rows.length; i++) if (String(rows[i][0]).trim() === name) return i; return -1; }
  if (body.action === "saveUser") {
    var u = body.user || {};
    var name = String(u.name || "").trim(), pw = String(u.pw || "").trim();
    if (!name || !pw) return { ok: false, error: "請填寫姓名和密碼" };
    var i = find(name);
    var disabled = u.disabled && name !== me ? "停用" : "";
    if (i < 0) {
      sh.getRange(sh.getLastRow() + 1, 1, 1, 3).setNumberFormat("@").setValues([[name, pw, disabled]]);
      i = sh.getLastRow() - 2;
    } else {
      sh.getRange(i + 2, 2, 1, 2).setNumberFormat("@").setValues([[pw, disabled]]);
    }
    // 電腦版會送 admin：true＝管理員、false＝助理（自己不能取消自己的管理員）
    if (typeof u.admin === "boolean" && !(name === me && !u.admin)) sh.getRange(i + 2, 4).setValue(u.admin ? "管理員" : "");
    return { ok: true, done: true };
  }
  if (body.action === "deleteUser") {
    var dn = String(body.name || "").trim();
    if (dn === me) return { ok: false, error: "不能刪除自己" };
    var di = find(dn);
    if (di >= 0) sh.deleteRow(di + 2);
    return { ok: true, done: true };
  }
  return {
    ok: true,
    users: rows.filter(function (r) { return String(r[0]).trim(); }).map(function (r) {
      var nm = String(r[0]).trim();
      return { name: nm, pw: String(r[1]), disabled: !!String(r[2]).trim(), admin: isAdmin_(nm, r[3]) };
    })
  };
}

// 帳號＝姓名，密碼＝試算表 B 欄（使用者 2026-11-23 要求：取消「密碼錯太多次鎖 10 分鐘」）
function checkUser_(name, pw) {
  if (!name || !String(pw || "").trim()) return { ok: false, error: "請輸入姓名與密碼" };
  var sh = book_().getSheetByName(SH_USER);
  if (sh && sh.getLastRow() > 1) {
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(3, Math.min(4, sh.getLastColumn()))).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === name && samePw_(rows[i][1], pw)) {
        if (String(rows[i][2]).trim()) return { ok: false, error: "此帳號已停用" };
        return { ok: true, admin: isAdmin_(name, rows[i][3]) };
      }
    }
  }
  return { ok: false, error: "姓名或密碼不正確" };
}

function upsert_(so, sl, o, who, rev) {
  if (!o || !o.id) return;
  var month = String(o.orderDate || "").slice(0, 7);
  removeRows_(so, o.id);
  removeRows_(sl, o.id);
  var lines = o.lines || [];
  var saved = o.saved ? Utilities.formatDate(new Date(o.saved), "Asia/Taipei", "yyyy-MM-dd HH:mm") : "";
  // 先把日期、電話欄設成純文字再寫入，避免 0912… 變成 912…、日期被轉換
  var r0 = so.getLastRow() + 1;
  so.getRange(r0, 5, 1, 2).setNumberFormat("@");
  so.getRange(r0, 8, 1, 1).setNumberFormat("@");
  so.getRange(r0, COL_REV, 1, 2).setNumberFormat("@");
  var full = JSON.stringify(o);
  if (full.length > 45000) full = "";
  so.getRange(r0, 1, 1, ORDER_HEAD.length).setValues([[o.id, o.no, o.site, o.company, text_(o.orderDate), text_(o.dueDate), o.contact, text_(o.contactTel), lines.length, o.exported ? "是" : "否", who, saved, o.note || "", month, rev ? String(rev) : "", full, false]]);
  so.getRange(r0, COL_ODEL).insertCheckboxes();
  if (lines.length) {
    var rows = lines.map(function (l) {
      return [o.id, o.no, o.site, text_(o.orderDate), l.cat, l.name, text_(l.size), l.spec, l.unit, Number(l.qty) || 0, l.remark, who, month, false, l.id || "", text_(l.arrTime), text_(l.supplier)];
    });
    var r1 = sl.getLastRow() + 1;
    sl.getRange(r1, 4, rows.length, 1).setNumberFormat("@");
    sl.getRange(r1, 7, rows.length, 1).setNumberFormat("@");
    sl.getRange(r1, 1, rows.length, LINE_HEAD.length).setValues(rows);
    sl.getRange(r1, LCOL_DEL, rows.length, 1).insertCheckboxes();
  }
  ensureSiteSheet_(so.getParent(), o.site);
}

function text_(v) {
  return v == null ? "" : String(v);
}

function removeRows_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  // 由下往上刪，列號才不會跑掉
  for (var r = ids.length - 1; r >= 0; r--) {
    if (ids[r][0] === id) sh.deleteRow(r + 2);
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- 依工地（案子）分頁 ----------
// 每個工地一個分頁，內容用公式從「明細」自動帶出，只列該工地的叫料品項（新到舊）。
function siteSheetName_(site) {
  var n = String(site || "未指定工地").replace(/[\[\]*?:\/\\'"]/g, "-").trim().slice(0, 90) || "未指定工地";
  if ([SH_ORDER, SH_LINE, SH_STAT, SH_USER, SH_DEL].indexOf(n) >= 0) n += "（案）";
  return n;
}

function ensureSiteSheet_(ss, site) {
  var name = siteSheetName_(site);
  if (ss.getSheetByName(name)) return;
  var q = String(site || "未指定工地").replace(/["']/g, "");
  var sh = ss.insertSheet(name);
  sh.getRange("A1").setFormula('=IFERROR(QUERY(' + SH_LINE + '!A1:M,"select B, D, E, F, G, H, I, J, K, L where C = \'' + q + '\' order by D desc, B desc",1),"尚無叫料單")');
  sh.setFrozenRows(1);
  sh.getRange("1:1").setFontWeight("bold").setBackground("#e8f0e4");
  sh.setColumnWidths(1, 1, 130);
  sh.setColumnWidths(4, 1, 220);
}

// 補建所有已出現過的工地分頁（setup 會順便執行）
function syncSiteSheets_(ss) {
  var so = ss.getSheetByName(SH_ORDER);
  if (!so || so.getLastRow() < 2) return;
  var seen = {};
  so.getRange(2, 3, so.getLastRow() - 1, 1).getValues().forEach(function (r) {
    var s = String(r[0]).trim();
    if (s && !seen[s]) { seen[s] = true; ensureSiteSheet_(ss, s); }
  });
}

// ---------- 管理員修改／刪除叫料單，並同步回每支手機的「紀錄」 ----------
function lineSheet_(ss) {
  var sl = ensureSheet_(ss, SH_LINE, LINE_HEAD);
  if (sl.getLastColumn() < LINE_HEAD.length || String(sl.getRange(1, 16).getValue()) !== LINE_HEAD[15]) {
    sl.getRange(1, 1, 1, LINE_HEAD.length).setValues([LINE_HEAD]).setFontWeight("bold").setBackground("#e8f0e4");
  }
  return sl;
}

function orderSheet_(ss) {
  var so = ensureSheet_(ss, SH_ORDER, ORDER_HEAD);
  if (so.getLastColumn() < ORDER_HEAD.length) {
    so.getRange(1, 1, 1, ORDER_HEAD.length).setValues([ORDER_HEAD]).setFontWeight("bold").setBackground("#e8f0e4");
  }
  return so;
}

function deletedIds_(ss) {
  var sh = ss.getSheetByName(SH_DEL), m = {};
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function (r) { if (r[0]) m[r[0]] = { who: String(r[1]), at: Number(r[2]) || 0 }; });
  return m;
}

function findOrder_(so, id) {
  var last = so.getLastRow();
  if (last < 2) return null;
  var ids = so.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) {
      var row = so.getRange(i + 2, 1, 1, ORDER_HEAD.length).getValues()[0];
      return { row: i + 2, who: String(row[10] || ""), rev: Number(row[COL_REV - 1]) || 0, data: row[COL_DATA - 1] };
    }
  }
  return null;
}

// 舊資料沒有「資料」欄時，用訂單＋明細組回來
function rowsToOrders_(ss) {
  var so = orderSheet_(ss), sl = ss.getSheetByName(SH_LINE);
  var out = [], byId = {};
  if (so.getLastRow() < 2) return out;
  so.getRange(2, 1, so.getLastRow() - 1, ORDER_HEAD.length).getValues().forEach(function (r) {
    if (!r[0]) return;
    var d = null;
    try { d = r[COL_DATA - 1] ? JSON.parse(r[COL_DATA - 1]) : null; } catch (e) { d = null; }
    var o = d || { id: r[0], no: String(r[1]), site: String(r[2]), company: String(r[3]), orderDate: String(r[4]), dueDate: String(r[5]), contact: String(r[6]), contactTel: String(r[7]), exported: r[9] === "是", note: String(r[12] || ""), lines: null };
    var item = { id: r[0], who: String(r[10] || ""), rev: Number(r[COL_REV - 1]) || 0, order: o };
    out.push(item);
    if (!o.lines) { o.lines = []; byId[r[0]] = o; }
  });
  if (sl && sl.getLastRow() > 1 && Object.keys(byId).length) {
    sl.getRange(2, 1, sl.getLastRow() - 1, LINE_HEAD.length).getValues().forEach(function (r) {
      var o = byId[r[0]];
      if (o) o.lines.push({ cat: String(r[4]), name: String(r[5]), size: String(r[6]), spec: String(r[7]), unit: String(r[8]), qty: Number(r[9]) || 0, remark: String(r[10] || ""), arrTime: String(r[15] || ""), supplier: String(r[16] || "") });
    });
  }
  return out;
}

// 手機來拿：管理員改過的單、被刪除的單（管理員 all=true 時拿全部）
function pull_(who, admin, all) {
  var ss = book_();
  var list = rowsToOrders_(ss);
  var dead = deletedIds_(ss), gone = [];
  Object.keys(dead).forEach(function (id) { if ((admin && all) || dead[id].who === who) gone.push({ id: id, at: dead[id].at }); });
  var sites = getSites_(ss);
  var resetAt = Number(PropertiesService.getScriptProperties().getProperty("RESET_AT") || 0);
  if (admin && all) return { ok: true, orders: list, deleted: gone, sites: sites, resetAt: resetAt, now: Date.now() };
  return { ok: true, orders: list.filter(function (x) { return x.who === who && x.rev; }), deleted: gone, sites: sites, resetAt: resetAt, now: Date.now() };
}

function adminOrder_(body, me) {
  var ss = book_();
  var so = orderSheet_(ss), sl = lineSheet_(ss);
  if (body.action === "adminDeleteOrder") {
    var id = String(body.id || "");
    var cur = findOrder_(so, id);
    var del = ensureSheet_(ss, SH_DEL, DEL_HEAD);
    var info = {};
    try { info = cur && cur.data ? JSON.parse(cur.data) : {}; } catch (e) {}
    del.appendRow([id, cur ? cur.who : String(body.owner || ""), Date.now(), me, info.no || "", info.site || ""]);
    removeRows_(so, id);
    removeRows_(sl, id);
    return { ok: true, done: true };
  }
  var o = body.order;
  if (!o || !o.id) return { ok: false, error: "資料不完整" };
  var c = findOrder_(so, o.id);
  var rev = Date.now();
  o.adminRev = rev;
  o.adminBy = me;
  upsert_(so, sl, o, c ? c.who || me : String(body.owner || me), rev);
  return { ok: true, done: true, rev: rev };
}

// 叫料人在手機「紀錄」刪除自己的單：試算表也刪掉（只能刪自己的；管理員可刪任何人的）
function deleteOwn_(ids, who, admin) {
  var ss = book_();
  var so = orderSheet_(ss), sl = lineSheet_(ss), del = ensureSheet_(ss, SH_DEL, DEL_HEAD);
  var n = 0;
  ids.forEach(function (id) {
    id = String(id || "");
    if (!id) return;
    var cur = findOrder_(so, id);
    if (cur && cur.who && cur.who !== who && !admin) return;
    var info = {};
    try { info = cur && cur.data ? JSON.parse(cur.data) : {}; } catch (e) {}
    del.appendRow([id, cur ? cur.who || who : who, Date.now(), who, info.no || "", info.site || ""]);
    if (cur) { removeRows_(so, id); removeRows_(sl, id); n++; }
  });
  return { ok: true, done: true, n: n };
}

// ---------- 直接在試算表修改／刪除，也同步回手機 ----------
// 「訂單」：改訂貨日、交貨日、聯絡人、電話、備註，或在「刪除（打勾）」打勾刪整張。
// 「明細」：改品名、尺寸、規格、單位、數量、備註，或打勾刪那一項；
//          在最下面新增一列，填「單號」「品名」「數量」就會加進那張單。
// （工地分頁是公式自動產生的，請到「訂單」「明細」改。）
// 第一次要在編輯器執行一次 installEditTrigger 並授權。
function installEditTrigger() {
  var ss = book_();
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === "onSheetEdit") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("onSheetEdit").forSpreadsheet(ss).onEdit().create();
  var so = orderSheet_(ss), sl = lineSheet_(ss);
  if (so.getLastRow() > 1) so.getRange(2, COL_ODEL, so.getLastRow() - 1, 1).insertCheckboxes();
  if (sl.getLastRow() > 1) sl.getRange(2, LCOL_DEL, sl.getLastRow() - 1, 1).insertCheckboxes();
}

var ORDER_EDIT_COLS = { 5: "orderDate", 6: "dueDate", 7: "contact", 8: "contactTel", 13: "note" };
var LINE_EDIT_COLS = { 6: "name", 7: "size", 8: "spec", 9: "unit", 10: "qty", 11: "remark", 16: "arrTime", 17: "supplier" };

function onSheetEdit(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet(), name = sh.getName();
  if (name === SH_SITE) { siteEdit_(e, sh); return; }
  if (name === SH_PPL) { pplEdit_(e, sh); return; }
  if (name !== SH_ORDER && name !== SH_LINE) return;
  var row = e.range.getRow(), col = e.range.getColumn();
  if (row < 2 || e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var ss = sh.getParent();
    var so = orderSheet_(ss), sl = lineSheet_(ss);
    if (name === SH_ORDER) {
      var ov = so.getRange(row, 1, 1, ORDER_HEAD.length).getValues()[0];
      var oid = ov[0];
      if (!oid) return;
      if (col === COL_ODEL) {
        if (ov[COL_ODEL - 1] !== true) return;
        var info = {};
        try { info = ov[COL_DATA - 1] ? JSON.parse(ov[COL_DATA - 1]) : {}; } catch (x) {}
        ensureSheet_(ss, SH_DEL, DEL_HEAD).appendRow([oid, String(ov[10] || ""), Date.now(), "試算表", info.no || ov[1], info.site || ov[2]]);
        removeRows_(so, oid);
        removeRows_(sl, oid);
        return;
      }
      var f = ORDER_EDIT_COLS[col];
      if (!f) return;
      var o = orderData_(ss, oid);
      if (!o) return;
      o[f] = text_(e.range.getDisplayValue());
      saveFromSheet_(so, sl, o, String(ov[10] || ""));
      return;
    }
    var lv = sl.getRange(row, 1, 1, LINE_HEAD.length).getValues()[0];
    var id = lv[0];
    if (!id) { addLineFromSheet_(ss, so, sl, row, lv); return; }
    var lf = LINE_EDIT_COLS[col];
    if (col !== LCOL_DEL && !lf) return;
    if (col === LCOL_DEL && lv[LCOL_DEL - 1] !== true) return;
    var od = orderData_(ss, id);
    if (!od) return;
    // 找出這一列是那張單的第幾項：先用明細ID，沒有就用順序
    var idx = -1, lid = String(lv[LCOL_ID - 1] || "");
    if (lid) for (var i = 0; i < od.lines.length; i++) if (od.lines[i].id === lid) { idx = i; break; }
    if (idx < 0) {
      var ids = sl.getRange(2, 1, row - 1, 1).getValues(), k = 0;
      for (var r = 0; r < ids.length - 1; r++) if (ids[r][0] === id) k++;
      idx = k;
    }
    if (idx < 0 || idx >= od.lines.length) return;
    if (col === LCOL_DEL) od.lines.splice(idx, 1);
    else od.lines[idx][lf] = lf === "qty" ? (Number(e.range.getValue()) || 0) : text_(e.range.getDisplayValue());
    saveFromSheet_(so, sl, od, String(lv[11] || ""));
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

// 在「明細」最下面新增一列：填「單號」「品名」「數量」（可再填尺寸、單位、備註），三個都填好就會加進那張叫料單
function addLineFromSheet_(ss, so, sl, row, lv) {
  var no = String(lv[1] || "").trim(), nm = String(lv[5] || "").trim(), q = Number(lv[9]) || 0;
  if (!no || !nm || q <= 0) return;
  var list = rowsToOrders_(ss), hit = null;
  for (var i = 0; i < list.length; i++) if (String(list[i].order.no || "").trim() === no) { hit = list[i]; break; }
  if (!hit) { sl.getRange(row, 11).setValue("找不到單號 " + no + "，請確認"); return; }
  var o = hit.order; o.lines = o.lines || [];
  var lid = "s" + Date.now().toString(36);
  o.lines.push({ id: lid, vk: "c|" + lid, ik: "custom", cat: String(lv[4] || "自訂"), name: nm, size: text_(lv[6]), spec: text_(lv[7]), unit: String(lv[8] || "個"), qty: q, remark: text_(lv[10]), arrTime: text_(lv[15]), supplier: text_(lv[16]) });
  sl.deleteRow(row);
  saveFromSheet_(so, sl, o, hit.who);
}

function orderData_(ss, id) {
  var list = rowsToOrders_(ss);
  for (var i = 0; i < list.length; i++) if (list[i].id === id) { var o = list[i].order; o.lines = o.lines || []; return o; }
  return null;
}

function saveFromSheet_(so, sl, o, who) {
  var c = findOrder_(so, o.id);
  var rev = Date.now();
  o.adminRev = rev;
  o.adminBy = "試算表";
  upsert_(so, sl, o, c ? c.who || who : who, rev);
}

// 在「工地資料」改工地名稱：收貨人分頁跟著改名；新增一列時自動給工地ID
function siteEdit_(e, sh) {
  var row = e.range.getRow();
  if (row < 2) return;
  var ss = sh.getParent(), ph = ss.getSheetByName(SH_PPL);
  if (e.range.getNumRows() === 1 && e.range.getNumColumns() === 1 && e.range.getColumn() === 1 && ph && e.oldValue) {
    var nv = String(e.value || "").trim();
    if (nv) renamePpl_(ph, String(e.oldValue).trim(), nv);
  }
  for (var r = row; r < row + e.range.getNumRows(); r++) {
    var v = sh.getRange(r, 1, 1, 5).getValues()[0];
    if (String(v[0]).trim() && !v[4]) { sh.getRange(r, 5).setValue(newId_()); if (v[3] !== true && v[3] !== false) sh.getRange(r, 4).insertCheckboxes(); }
    if (String(v[0]).trim()) sh.getRange(r, 6).setValue(stamp_());
  }
}

// 在「收貨人」最下面自己加一列：自動補上打勾框與修改時間
function pplEdit_(e, sh) {
  var row = e.range.getRow();
  if (row < 2) return;
  for (var r = row; r < row + e.range.getNumRows(); r++) {
    var v = sh.getRange(r, 1, 1, 4).getValues()[0];
    if (!String(v[0]).trim() && !String(v[1]).trim()) continue;
    if (v[3] !== true && v[3] !== false) sh.getRange(r, 4).insertCheckboxes();
    sh.getRange(r, 5).setValue(stamp_());
  }
}

// 建立／補齊「工地資料」「收貨人」分頁（可在編輯器直接執行）
function setupSites() {
  siteSheets_(book_());
}

// 清除所有叫料單（重新開始測試用）：清空「訂單」「明細」「已刪除」，刪掉舊的備份分頁，
// 手機下次同步時也會把這個時間以前的叫料紀錄清掉。
function clearAllOrders() {
  var ss = book_(), so = orderSheet_(ss), sl = lineSheet_(ss);
  PropertiesService.getScriptProperties().setProperty("RESET_AT", String(Date.now()));
  [so, sl, ss.getSheetByName(SH_DEL)].forEach(function (sh) {
    if (sh && sh.getLastRow() > 1) clearRows_(sh, sh.getLastRow() - 1);
  });
  ss.getSheets().forEach(function (sh) { if (/^備份(訂單|明細)_/.test(sh.getName())) ss.deleteSheet(sh); });
}
function clearRows_(sh, n) {
  if (sh.getMaxRows() <= n + 1) sh.insertRowsAfter(sh.getMaxRows(), 5);
  sh.deleteRows(2, n);
}

// ---------- 電腦版：廠商名單、案場採購紀錄 ----------
// 把「常用廠商名單_整理版.xlsx」用「檔案 → 匯入 → 插入新工作表」匯進這份試算表，就會有這兩個分頁。
// 第一列是標題；程式會自動在最後補「聯絡人」「電話」「刪除（打勾）」欄。刪除＝打勾，可以恢復。
var TABLES = { "常用廠商名單": ["聯絡人", "電話"], "案場採購紀錄": [] };
var DEL_COL = "刪除（打勾）";
function tableSheet_(name) {
  if (!Object.prototype.hasOwnProperty.call(TABLES, name)) return { error: "沒有這個分頁" };
  var sh = book_().getSheetByName(name);
  if (!sh) return { error: "試算表裡還沒有「" + name + "」分頁，請先把常用廠商名單 Excel 匯入" };
  var head = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(function (h) { return String(h).trim(); });
  TABLES[name].concat([DEL_COL]).forEach(function (h) {
    if (head.indexOf(h) < 0) { head.push(h); sh.getRange(1, head.length).setValue(h).setFontWeight("bold"); }
  });
  return { sh: sh, head: head };
}
function getTable_(name) {
  var t = tableSheet_(name);
  if (t.error) return { ok: false, error: t.error, missing: true };
  var n = t.sh.getLastRow() - 1;
  var vals = n > 0 ? t.sh.getRange(2, 1, n, t.head.length).getDisplayValues() : [];
  var di = t.head.indexOf(DEL_COL);
  var rows = [];
  vals.forEach(function (r, i) {
    if (!r.some(function (v, k) { return k !== di && String(v).trim(); })) return;
    rows.push({ row: i + 2, v: r.map(function (v, k) { return k === di ? (String(v).toUpperCase() === "TRUE" || String(v) === "是") : String(v); }) });
  });
  return { ok: true, head: t.head, rows: rows };
}
// {name, row:列號(新增填0), values:{標題:值}, del:true/false}
function saveRow_(body, who) {
  var t = tableSheet_(String(body.name || ""));
  if (t.error) return { ok: false, error: t.error };
  var sh = t.sh, head = t.head, row = Number(body.row) || 0;
  if (row < 2 || row > sh.getLastRow()) { row = sh.getLastRow() + 1; }
  var cur = sh.getRange(row, 1, 1, head.length).getValues()[0];
  var vals = body.values || {};
  head.forEach(function (h, k) {
    if (h === DEL_COL) { if (typeof body.del === "boolean") cur[k] = body.del; return; }
    if (Object.prototype.hasOwnProperty.call(vals, h)) cur[k] = String(vals[h]);
  });
  sh.getRange(row, 1, 1, head.length).setValues([cur]);
  sh.getRange(row, head.indexOf(DEL_COL) + 1).insertCheckboxes();
  return getTable_(String(body.name));
}

// ---------- 電腦版：出貨核對 ----------
// 試算表「出貨核對」分頁一張叫料單一列；照片存在雲端硬碟「叫料小幫手_出貨照片／工地／年-月」資料夾，只有這個 Google 帳號看得到。
var SH_SHIP = "出貨核對";
var SHIP_HEAD = ["訂單ID", "單號", "工地", "狀態", "核對人", "核對時間", "備註", "照片（檔案ID）", "到貨數量（程式用）", "照片判讀文字", "出貨單明細（程式用）"];
// 每張出貨單上的廠商品名、數量（之後對帳單核對用）；儲存出貨核對時整張叫料單的明細會重寫
var SH_SLIP = "出貨明細";
var SLIP_HEAD = ["叫料單ID", "叫料單號", "工地", "廠商", "出貨單號", "出貨日期", "廠商品名（出貨單上）", "數量", "對應叫料品項", "照片（檔案ID）", "核對人", "核對時間"];
function slipWrite_(id, no, site, slips, who, now) {
  var sh = ensureSheet_(book_(), SH_SLIP, SLIP_HEAD);
  if (sh.getFrozenRows() < 1) sh.setFrozenRows(1);
  var n = sh.getLastRow() - 1;
  if (n > 0) {
    var ids = sh.getRange(2, 1, n, 1).getValues();
    for (var i = n - 1; i >= 0; i--) if (String(ids[i][0]) === String(id)) sh.deleteRow(i + 2);
  }
  var rows = [];
  (slips || []).forEach(function (s) {
    (s.items || []).forEach(function (it) {
      if (!String(it.name || "").trim()) return;
      rows.push([id, no, site, s.vendor || "", s.no || "", s.date || "", it.name, it.qty === "" || it.qty == null ? "" : it.qty, it.lkName || "", s.fileId || "", who, now].map(String));
    });
  });
  if (rows.length) {
    var r0 = sh.getLastRow() + 1;
    sh.getRange(r0, 1, rows.length, SLIP_HEAD.length).setNumberFormat("@").setValues(rows);
  }
}
var SHIP_ROOT = "叫料小幫手_出貨照片";
function shipSheet_() {
  var sh = ensureSheet_(book_(), SH_SHIP, SHIP_HEAD);
  if (sh.getLastColumn() < SHIP_HEAD.length) sh.getRange(1, 1, 1, SHIP_HEAD.length).setValues([SHIP_HEAD]).setFontWeight("bold").setBackground("#e8f0e4");
  return sh;
}
function shipRows_(sh) {
  var n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, SHIP_HEAD.length).getValues() : [];
}
function shipRec_(r) {
  var recv = {};
  try { recv = JSON.parse(r[8] || "{}"); } catch (e) {}
  var slips = [];
  try { slips = JSON.parse(r[10] || "[]"); } catch (e) {}
  return { id: String(r[0]), no: String(r[1]), site: String(r[2]), status: String(r[3]), who: String(r[4]),
    at: r[5] instanceof Date ? Utilities.formatDate(r[5], "Asia/Taipei", "yyyy-MM-dd HH:mm") : String(r[5]),
    note: String(r[6]), photos: String(r[7]).split(",").filter(String), recv: recv, text: String(r[9] || ""), slips: slips };
}
function shipFind_(sh, id) {
  var rows = shipRows_(sh);
  for (var i = 0; i < rows.length; i++) if (String(rows[i][0]) === String(id)) return { row: i + 2, r: rows[i] };
  return null;
}
function shipPut_(sh, body, f, patch) {
  var r = f ? f.r : [String(body.id), "", "", "", "", "", "", "", "{}", "", "[]"];
  r[1] = String(body.no || r[1]); r[2] = String(body.site || r[2]);
  patch(r);
  var row = f ? f.row : sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, SHIP_HEAD.length).setNumberFormat("@").setValues([r.map(String)]);
  return shipRec_(r);
}
function shipFolder_(site, month) {
  var get = function (parent, name) {
    var it = parent ? parent.getFoldersByName(name) : DriveApp.getFoldersByName(name);
    return it.hasNext() ? it.next() : (parent ? parent.createFolder(name) : DriveApp.createFolder(name));
  };
  return get(get(get(null, SHIP_ROOT), site || "未填工地"), month);
}
// 用 Google 雲端硬碟內建的文字辨識讀照片上的字（印刷字比較準，手寫的要人工確認）
function shipOcr_(blob) {
  var tok = ScriptApp.getOAuthToken(), bd = "jlbd" + Date.now();
  var meta = JSON.stringify({ name: "ocr_tmp_" + Date.now(), mimeType: "application/vnd.google-apps.document" });
  var payload = Utilities.newBlob("--" + bd + "\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n" + meta + "\r\n--" + bd + "\r\nContent-Type: " + blob.getContentType() + "\r\n\r\n").getBytes()
    .concat(blob.getBytes()).concat(Utilities.newBlob("\r\n--" + bd + "--").getBytes());
  var res = UrlFetchApp.fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart", {
    method: "post", contentType: "multipart/related; boundary=" + bd, payload: payload, headers: { Authorization: "Bearer " + tok }, muteHttpExceptions: true });
  var doc = JSON.parse(res.getContentText());
  if (!doc.id) return "";
  var txt = UrlFetchApp.fetch("https://www.googleapis.com/drive/v3/files/" + doc.id + "/export?mimeType=text/plain", { headers: { Authorization: "Bearer " + tok }, muteHttpExceptions: true }).getContentText("UTF-8");
  try { DriveApp.getFileById(doc.id).setTrashed(true); } catch (e) {}
  return String(txt || "").trim();
}
function ship_(body, who) {
  var sh = shipSheet_(), now = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm");
  if (body.action === "shipList") {
    var all = {};
    shipRows_(sh).forEach(function (r) { if (r[0]) { var x = shipRec_(r); delete x.text; all[x.id] = x; } });
    return { ok: true, ships: all };
  }
  var id = String(body.id || "");
  if (!id) return { ok: false, error: "缺少叫料單" };
  var f = shipFind_(sh, id);
  if (body.action === "shipGet") return { ok: true, ship: f ? shipRec_(f.r) : null };
  if (body.action === "shipSave") {
    var rec = shipPut_(sh, body, f, function (r) {
      r[3] = String(body.status || ""); r[4] = who; r[5] = now; r[6] = String(body.note || "");
      r[8] = JSON.stringify(body.recv || {});
      if (body.slips) r[10] = JSON.stringify(body.slips).slice(0, 45000);
    });
    if (body.slips) slipWrite_(id, String(body.no || rec.no), String(body.site || rec.site), body.slips, who, now);
    return { ok: true, ship: rec };
  }
  if (body.action === "shipPhoto") {
    var data = String(body.data || "");
    if (!data || data.length > 8000000) return { ok: false, error: "照片太大或是空的" };
    var mime = /^image\/(jpeg|png|webp|heic)$/.test(body.mime) ? body.mime : "image/jpeg";
    var month = now.slice(0, 7), site = String(body.site || (f && f.r[2]) || "");
    var blob = Utilities.newBlob(Utilities.base64Decode(data), mime, (body.no || id) + "_" + now.replace(/[ :]/g, "") + ".jpg");
    var file = shipFolder_(site, month).createFile(blob);
    var text = "";
    try { text = shipOcr_(blob); } catch (e) { text = ""; }
    var rec = shipPut_(sh, body, f, function (r) {
      r[7] = String(r[7]).split(",").filter(String).concat([file.getId()]).join(",");
      if (text) r[9] = text.slice(0, 40000);
    });
    return { ok: true, ship: rec, fileId: file.getId(), text: text };
  }
  var fid = String(body.fileId || "");
  if (!f || String(f.r[7]).split(",").indexOf(fid) < 0) return { ok: false, error: "找不到這張照片" };
  if (body.action === "shipImg") {
    var b = DriveApp.getFileById(fid).getBlob();
    return { ok: true, mime: b.getContentType(), data: Utilities.base64Encode(b.getBytes()) };
  }
  if (body.action === "shipDelPhoto") {
    try { DriveApp.getFileById(fid).setTrashed(true); } catch (e) {}
    return { ok: true, ship: shipPut_(sh, body, f, function (r) {
      r[7] = String(r[7]).split(",").filter(function (x) { return x && x !== fid; }).join(",");
    }) };
  }
  return { ok: false, error: "未知的動作" };
}
// 第一次用出貨核對前，在編輯器選這個函式按「執行」，讓 Google 允許程式存取雲端硬碟。
function authShip() {
  DriveApp.getRootFolder();
  UrlFetchApp.fetch("https://www.googleapis.com/drive/v3/about?fields=user", { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  shipSheet_();
}

// ---------- 電腦版：價目表 ----------
// 試算表「價目表」分頁一列一個「廠商 × 品項」的價格；原始報價單檔案存在雲端硬碟「叫料小幫手_報價單／廠商／民國年」。只有管理員能用。
var SH_PRICE = "價目表";
var PRICE_HEAD = ["編號", "廠商", "日期", "品名規格", "材質", "尺寸", "規格標準", "單位", "單價", "牌價", "折數", "工地／報價單號", "來源檔", "雲端檔案ID", "資料狀態／備註", "上傳人", "上傳時間"];
var PRICE_ROOT = "叫料小幫手_報價單";
function priceSheet_() {
  var sh = ensureSheet_(book_(), SH_PRICE, PRICE_HEAD);
  if (sh.getFrozenRows() < 1) sh.setFrozenRows(1);
  return sh;
}
function priceRows_(sh) {
  var n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, PRICE_HEAD.length).getValues() : [];
}
function priceWrite_(sh, row, rows) {
  if (!rows.length) return;
  var rg = sh.getRange(row, 1, rows.length, PRICE_HEAD.length);
  rg.setNumberFormat("@");
  sh.getRange(row, 9, rows.length, 3).setNumberFormat("#,##0.####");
  rg.setValues(rows.map(function (r) {
    return PRICE_HEAD.map(function (h, i) {
      var v = r[i] == null ? "" : r[i];
      if (v instanceof Date) v = Utilities.formatDate(v, "Asia/Taipei", "yyyy-MM-dd HH:mm");
      if (i >= 8 && i <= 10) { var n = Number(String(v).replace(/[,\s]/g, "")); return String(v).trim() !== "" && isFinite(n) ? n : String(v); }
      return String(v);
    });
  }));
}
function priceFolder_(vendor, date) {
  var get = function (parent, name) {
    var it = parent ? parent.getFoldersByName(name) : DriveApp.getFoldersByName(name);
    return it.hasNext() ? it.next() : (parent ? parent.createFolder(name) : DriveApp.createFolder(name));
  };
  var y = (String(date || "").match(/^\d{2,3}/) || [String(new Date().getFullYear() - 1911)])[0];
  return get(get(get(null, PRICE_ROOT), vendor || "未填廠商"), y + "年");
}
function price_(body, who) {
  var sh = priceSheet_(), now = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm");
  if (body.action === "priceList") {
    return { ok: true, head: PRICE_HEAD, rows: priceRows_(sh).filter(function (r) { return r[0]; }).map(function (r) {
      return r.map(function (v) { return v instanceof Date ? Utilities.formatDate(v, "Asia/Taipei", "yyyy-MM-dd HH:mm") : v; });
    }) };
  }
  if (body.action === "priceFile") {
    var data = String(body.data || "");
    if (!data || data.length > 30000000) return { ok: false, error: "檔案太大或是空的（上限約 20MB）" };
    var mime = String(body.mime || "application/octet-stream");
    var blob = Utilities.newBlob(Utilities.base64Decode(data), mime, String(body.name || ("報價單_" + now.replace(/[ :]/g, ""))));
    var file = priceFolder_(String(body.vendor || ""), String(body.date || "")).createFile(blob);
    var text = "";
    if (body.ocr && /^image\/|^application\/pdf$/.test(mime)) { try { text = shipOcr_(blob); } catch (e) { text = ""; } }
    return { ok: true, fileId: file.getId(), name: file.getName(), text: text };
  }
  if (body.action === "priceGetFile") {
    var fid = String(body.fileId || ""), okf = priceRows_(sh).some(function (r) { return String(r[13]) === fid; });
    if (!fid || !okf) return { ok: false, error: "找不到這個檔案" };
    var b = DriveApp.getFileById(fid).getBlob();
    return { ok: true, mime: b.getContentType(), name: b.getName(), data: Utilities.base64Encode(b.getBytes()) };
  }
  if (body.action === "priceAdd") {
    var add = (body.rows || []).slice(0, 2000), base = "P" + Date.now().toString(36);
    var rows = add.map(function (v, i) {
      var r = PRICE_HEAD.map(function (h, k) { return v[k] == null ? "" : v[k]; });
      r[0] = base + "-" + i; r[15] = who; r[16] = now;
      return r;
    }).filter(function (r) { return String(r[3]).trim(); });
    priceWrite_(sh, sh.getLastRow() + 1, rows);
    return { ok: true, added: rows.length, ids: rows.map(function (r) { return r[0]; }) };
  }
  if (body.action === "priceSave") {
    var all = priceRows_(sh), id = String(body.id || "");
    for (var i = 0; i < all.length; i++) if (String(all[i][0]) === id) {
      var r = all[i].slice(), v = body.v || [];
      for (var k = 1; k <= 14; k++) if (v[k] != null) r[k] = v[k];
      r[15] = who; r[16] = now;
      priceWrite_(sh, i + 2, [r]);
      return { ok: true };
    }
    return { ok: false, error: "找不到這一筆" };
  }
  if (body.action === "priceDel") {
    var ids = {}; (body.ids || []).forEach(function (x) { ids[String(x)] = 1; });
    var all2 = priceRows_(sh), keep = all2.filter(function (r) { return !ids[String(r[0])]; }), n = all2.length - keep.length;
    if (!n) return { ok: true, deleted: 0 };
    if (n <= 10) { for (var j = all2.length - 1; j >= 0; j--) if (ids[String(all2[j][0])]) sh.deleteRow(j + 2); }
    else { sh.getRange(2, 1, all2.length, PRICE_HEAD.length).clearContent(); priceWrite_(sh, 2, keep); }
    return { ok: true, deleted: n };
  }
  return { ok: false, error: "未知的動作" };
}
