/**
 * 工程叫料小幫手 — Google 試算表後台
 *
 * 用法：在要當後台的 Google 試算表選「擴充功能 → Apps Script」，貼上這份程式碼，
 * 1) 執行一次 setup（建立「訂單」「明細」「統計」「使用者」分頁），
 * 2) 到「使用者」分頁填入帳號：A 欄＝姓名（登入帳號）、B 欄＝預設電話號碼（登入密碼）、C 欄填「停用」可暫停某人，
 * 3) 「部署 → 新增部署作業 → 網頁應用程式」：執行身分選「我」，誰可以存取選「所有人」，
 * 4) 把部署後的網址貼到 js/6-sync.js 的 SYNC_URL。
 */
// 初次建立「使用者」分頁時帶入的名單，格式：[["姓名","電話"], ...]；之後直接在試算表裡改即可。
var SEED_USERS = [];
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
var USER_HEAD = ["姓名（帳號）", "預設電話（密碼）", "停用"];
var ORDER_HEAD = ["訂單ID", "單號", "工地", "公司", "訂貨日", "交貨日", "聯絡人", "電話", "品項數", "已匯出", "叫料人", "儲存時間", "備註", "月份"];
var LINE_HEAD = ["訂單ID", "單號", "工地", "訂貨日", "分類", "品名", "尺寸", "規格說明", "單位", "數量", "備註", "叫料人", "月份"];

function setup() {
  var ss = book_();
  ensureSheet_(ss, SH_ORDER, ORDER_HEAD);
  ensureSheet_(ss, SH_LINE, LINE_HEAD);
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
    if (body.action === "login") return json_({ ok: true, name: who });
    lock.waitLock(25000);
    var ss = book_();
    var so = ensureSheet_(ss, SH_ORDER, ORDER_HEAD);
    var sl = ensureSheet_(ss, SH_LINE, LINE_HEAD);
    var orders = body.orders || [];
    for (var i = 0; i < orders.length; i++) upsert_(so, sl, orders[i], who);
    return json_({ ok: true, n: orders.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function digits_(v) {
  return String(v == null ? "" : v).replace(/\D/g, "");
}

// 帳號＝姓名，密碼＝預設電話號碼（不分有無破折號）；連續輸入錯誤太多次會暫時鎖住
function checkUser_(name, pw) {
  if (!name || !digits_(pw)) return { ok: false, error: "請輸入姓名與密碼" };
  var cache = CacheService.getScriptCache();
  var ck = "fail:" + name;
  if (Number(cache.get(ck) || 0) >= 8) return { ok: false, error: "密碼錯誤次數過多，請 10 分鐘後再試" };
  var sh = book_().getSheetByName(SH_USER);
  if (sh && sh.getLastRow() > 1) {
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === name && digits_(rows[i][1]) === digits_(pw)) {
        if (String(rows[i][2]).trim()) return { ok: false, error: "此帳號已停用" };
        return { ok: true };
      }
    }
  }
  cache.put(ck, String(Number(cache.get(ck) || 0) + 1), 600);
  return { ok: false, error: "姓名或密碼不正確" };
}

function upsert_(so, sl, o, who) {
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
  so.getRange(r0, 1, 1, ORDER_HEAD.length).setValues([[o.id, o.no, o.site, o.company, text_(o.orderDate), text_(o.dueDate), o.contact, text_(o.contactTel), lines.length, o.exported ? "是" : "否", who, saved, o.note || "", month]]);
  if (lines.length) {
    var rows = lines.map(function (l) {
      return [o.id, o.no, o.site, text_(o.orderDate), l.cat, l.name, text_(l.size), l.spec, l.unit, Number(l.qty) || 0, l.remark, who, month];
    });
    var r1 = sl.getLastRow() + 1;
    sl.getRange(r1, 4, rows.length, 1).setNumberFormat("@");
    sl.getRange(r1, 7, rows.length, 1).setNumberFormat("@");
    sl.getRange(r1, 1, rows.length, LINE_HEAD.length).setValues(rows);
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
  if ([SH_ORDER, SH_LINE, SH_STAT, SH_USER].indexOf(n) >= 0) n += "（案）";
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
