/**
 * מערכת קבלת סחורה – פירות וירקות
 * קובץ שרת (Google Apps Script). מדביקים אותו כ-Code.gs, ואת Index.html כקובץ HTML בשם Index.
 */
const SH = {P:'מוצרים', S:'ספקים', L:'מחירונים', F:'קבצי מחירון', R:'קבלות', O:'הזמנות', G:'הגדרות'};
const O_HEAD = ['מזהה','נוצרה','תאריך אספקה','ספק','סטטוס','פירוט','הערה','נוצר ע"י','פריטים (JSON)'];
const O_OPEN = ['טיוטה','נשלחה','התקבלה חלקית'];
const TYPE_OUT = {weekly:'שבועי', fixed:'קבוע', manual:'בקבלה'};
const TYPE_IN = {'שבועי':'weekly', 'קבוע':'fixed', 'בקבלה':'manual'};
const R_HEAD = ['מזהה','תאריך','שעה','מזהה מוצר','מוצר','יחידה','כמות','מחיר ליחידה','מקור מחיר','סה"כ','ספק','מצב','הערה','תמונות','מקבל','טופל','נמחק','מזהה הזמנה'];
const SEED = [["p000", "עגבניה", "ק\"ג", 3.3], ["p001", "מלפפון", "ק\"ג", 3.7], ["p002", "פלפל ירוק", "ק\"ג", 5.5], ["p003", "פלפל בהיר", "ק\"ג", 6.2], ["p004", "פלפל אדום", "ק\"ג", 5.8], ["p005", "פלפל צהוב", "ק\"ג", 5.5], ["p006", "פלפל כתום", "ק\"ג", 8.1], ["p007", "פלפל חריף", "ק\"ג", 5.7], ["p008", "קישוא", "ק\"ג", 3.5], ["p009", "חציל", "ק\"ג", 2.6], ["p010", "גזר ארוז", "ק\"ג", 3], ["p011", "גזר תפזורת", "ק\"ג", 2.3], ["p012", "סלק", "ק\"ג", 2.8], ["p013", "קולרבי", "ק\"ג", 5.5], ["p014", "בטטה", "ק\"ג", 3.7], ["p015", "דלורית", "ק\"ג", 3.7], ["p016", "דלעת", "ק\"ג", 3.4], ["p017", "שרי (יח')", "יחידה", 5.3], ["p018", "בצל", "ק\"ג", 2.8], ["p019", "בצל אדום", "ק\"ג", 2.9], ["p020", "שום רביעיה", "יחידה", 2.7], ["p021", "תפו\"א לבן", "ק\"ג", 3], ["p022", "תפו\"א אדום", "ק\"ג", 3], ["p023", "תפו\"ע סמיט 6.5", "ק\"ג", 8], ["p024", "תפו\"ע זהוב 6.5", "ק\"ג", 8], ["p025", "תפו\"ע חרמון 6.5", "ק\"ג", 8], ["p026", "אגס", "ק\"ג", 9.5], ["p027", "בננה", "ק\"ג", 8.5], ["p028", "לימון", "ק\"ג", 7.7], ["p029", "מלון", "ק\"ג", 3.1], ["p030", "אבטיח", "ק\"ג", 2], ["p031", "אפרסק", "ק\"ג", 9.5], ["p032", "סלק וואקום 2.8 ק\"ג", "יחידה", 19.3], ["p033", "סלק וואקום יח'", "יחידה", 4], ["p034", "סלק וואקום 500 גרם", "יחידה", 4.2], ["p035", "אננס (מחיר ליחידה)", "יחידה", 5]];
const SEED_WEEK = '2026-07-12';

function doGet() {
  setup_(); ensure_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('קבלת סחורה')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** להרצה ידנית פעם אחת מהעורך: יוצר את הלשוניות ומבקש הרשאות. */
function התקנה() {
  setup_(); ensure_();
  folder_('מחירונים');
  const s = settings_();
  Logger.log('מוכן. קוד עובד: ' + s['קוד עובד'] + ' | קוד מנהל מטבח: ' + s['קוד מנהל מטבח'] + ' | קוד מנהל: ' + s['קוד מנהל']);
}

/* ---------- setup ---------- */
function setup_() {
  const ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName(SH.R)) return;
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    if (ss.getSheetByName(SH.R)) return;
    ss.setSpreadsheetTimeZone('Asia/Jerusalem');
    const mk = (name, head, textCols) => {
      const sh = ss.getSheetByName(name) || ss.insertSheet(name);
      sh.clear();
      (textCols || []).forEach(c => sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'));
      sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#e3f1e7');
      sh.setFrozenRows(1); sh.setRightToLeft(true);
      return sh;
    };
    const p = mk(SH.P, ['מזהה','שם','יחידה','סוג מחיר','מחיר קבוע','פעיל'], [1]);
    const rows = SEED.map(r => [r[0], r[1], r[2], 'שבועי', '', 'כן']);
    rows.push(['p900', 'פטריות שמפיניון', 'ק"ג', 'בקבלה', '', 'כן']);
    p.getRange(2, 1, rows.length, 6).setValues(rows);
    const s = mk(SH.S, ['ספק','טלפון'], [2]); s.getRange(2, 1).setValue('ביכורי השדה');
    const l = mk(SH.L, ['שבוע (ראשון)','מזהה מוצר','מוצר','מחיר','עודכן'], [1, 2, 5]);
    const now = stamp_();
    l.getRange(2, 1, SEED.length, 5).setValues(SEED.map(r => [SEED_WEEK, r[0], r[1], r[3], now]));
    mk(SH.F, ['הועלה','שבוע','שם קובץ','קישור','מזהה קובץ'], [1, 2]);
    const r = mk(SH.R, R_HEAD, [1, 2, 3, 4]);
    r.getRange('H:H').setNumberFormat('₪#,##0.00'); r.getRange('J:J').setNumberFormat('₪#,##0.00');
    const g = mk(SH.G, ['הגדרה','ערך'], [2]);
    g.getRange(2, 1, 3, 2).setValues([['קוד עובד', '1234'], ['קוד מנהל', String(Math.floor(100000 + Math.random() * 900000))], ['תיקיית קבצים', '']]);
    ss.getSheets().forEach(x => { if (Object.values(SH).indexOf(x.getName()) < 0 && ss.getSheets().length > 1) ss.deleteSheet(x); });
  } finally { lock.releaseLock(); }
}

/** תוספות לגרסה עם הזמנות – רץ גם על התקנה קיימת */
function ensure_() {
  const ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName(SH.O)) return;
  withLock_(() => {
    if (ss.getSheetByName(SH.O)) return;
    const o = ss.insertSheet(SH.O);
    [1, 2, 3].forEach(c => o.getRange(1, c, o.getMaxRows(), 1).setNumberFormat('@'));
    o.getRange(1, 1, 1, O_HEAD.length).setValues([O_HEAD]).setFontWeight('bold').setBackground('#e3f1e7');
    o.setFrozenRows(1); o.setRightToLeft(true);
    sh_(SH.R).getRange(1, 18).setValue('מזהה הזמנה').setFontWeight('bold').setBackground('#e3f1e7');
    sh_(SH.S).getRange(1, 2).setValue('טלפון').setFontWeight('bold').setBackground('#e3f1e7');
    const s = settings_();
    if (!s['קוד מנהל מטבח']) sh_(SH.G).appendRow(['קוד מנהל מטבח', String(Math.floor(10000 + Math.random() * 90000))]);
    if (!('שם העסק' in s)) sh_(SH.G).appendRow(['שם העסק', '']);
    if (!('מפתח Gemini' in s)) sh_(SH.G).appendRow(['מפתח Gemini', '']);
  });
}

/* ---------- helpers ---------- */
function tz_() { return SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'Asia/Jerusalem'; }
function stamp_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm'); }
function dstr_(v) { return v instanceof Date ? Utilities.formatDate(v, tz_(), 'yyyy-MM-dd') : String(v || ''); }
function tstr_(v, f) { return v instanceof Date ? Utilities.formatDate(v, tz_(), f || 'HH:mm') : String(v || ''); }
/** טקסט שגוגל לא יהפוך למספר/תאריך */
function txt_(v) { return v === '' || v == null ? '' : "'" + String(v); }
function sh_(n) { return SpreadsheetApp.getActive().getSheetByName(n); }
function rows_(n) { const v = sh_(n).getDataRange().getValues(); v.shift(); return v; }
/* ---------- מטמון בזיכרון של גוגל, כדי לא לקרוא את הגיליון בכל בקשה ---------- */
function cget_(k) { try { const v = CacheService.getScriptCache().get(k); return v ? JSON.parse(v) : null; } catch (e) { return null; } }
function cput_(k, v, sec) { try { const s = JSON.stringify(v); if (s.length < 90000) CacheService.getScriptCache().put(k, s, sec || 300); } catch (e) {} return v; }
function cdel_() { try { CacheService.getScriptCache().removeAll(Array.prototype.slice.call(arguments)); } catch (e) {} }
function settings_() {
  const c = cget_('settings'); if (c) return c;
  const o = {}; rows_(SH.G).forEach(r => o[String(r[0])] = String(r[1]).trim()); return cput_('settings', o, 120);
}
function setSetting_(k, v) {
  const sh = sh_(SH.G), c = sh.getRange('A:A').createTextFinder(k).matchEntireCell(true).findNext();
  if (c) sh.getRange(c.getRow(), 2).setValue(v); else sh.appendRow([k, v]);
  cdel_('settings');
}
function role_(pin) {
  const s = settings_(); pin = String(pin || '').trim();
  if (pin && pin === s['קוד מנהל']) return 'manager';
  if (pin && pin === s['קוד מנהל מטבח']) return 'chef';
  if (pin && pin === s['קוד עובד']) return 'worker';
  throw new Error('קוד כניסה שגוי');
}
function mgr_(pin) { if (role_(pin) !== 'manager') throw new Error('הפעולה מותרת רק למנהל'); }
function chef_(pin) { const r = role_(pin); if (r !== 'manager' && r !== 'chef') throw new Error('הפעולה מותרת רק למנהל המטבח'); }
function withLock_(fn) { const l = LockService.getScriptLock(); l.waitLock(30000); try { return fn(); } finally { l.releaseLock(); } }
function weekOf_(ds) { const d = new Date(ds + 'T12:00:00'); d.setDate(d.getDate() - d.getDay()); return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd'); }
function addDays_(ds, n) { const d = new Date(ds + 'T12:00:00'); d.setDate(d.getDate() + n); return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd'); }
function folder_(sub) {
  const s = settings_(); let root = null;
  try { if (s['תיקיית קבצים']) root = DriveApp.getFolderById(s['תיקיית קבצים']); } catch (e) { root = null; }
  if (!root) {
    const parents = DriveApp.getFileById(SpreadsheetApp.getActive().getId()).getParents();
    root = (parents.hasNext() ? parents.next() : DriveApp.getRootFolder()).createFolder('קבלת סחורה – קבצים');
    setSetting_('תיקיית קבצים', root.getId());
  }
  const it = root.getFoldersByName(sub); return it.hasNext() ? it.next() : root.createFolder(sub);
}

function products_() {
  const c = cget_('products'); if (c) return c;
  return cput_('products', rows_(SH.P).filter(r => r[0] !== '').map(r => ({
    id: String(r[0]), name: String(r[1]), unit: String(r[2]), type: TYPE_IN[r[3]] || 'weekly',
    fixed: r[4] === '' || r[4] === null ? null : Number(r[4]), active: String(r[5]) !== 'לא'})), 600);
}
function lists_(fromWeek) {
  let all = cget_('lists');
  if (!all) { all = {}; rows_(SH.L).forEach(r => { const w = dstr_(r[0]); (all[w] = all[w] || {})[String(r[1])] = Number(r[3]); }); cput_('lists', all, 600); }
  const o = {}; Object.keys(all).forEach(w => { if (!fromWeek || w >= fromWeek) o[w] = all[w]; }); return o;
}
function rec_(r) {
  return {id: String(r[0]), ds: dstr_(r[1]), t: tstr_(r[2]), pid: String(r[3]), name: String(r[4]), unit: String(r[5]), qty: Number(r[6]),
    price: r[7] === '' ? null : Number(r[7]), src: String(r[8]), total: r[9] === '' ? null : Number(r[9]), sup: String(r[10]), status: String(r[11]),
    note: String(r[12]), photos: String(r[13] || '').split(/\s+/).filter(String), by: String(r[14]), handled: r[15] === 'כן', del: r[16] === 'כן', order: String(r[17] || '')};
}
function row_(o) {
  return [txt_(o.id), txt_(o.ds), txt_(o.t), txt_(o.pid), o.name, o.unit, o.qty, o.price == null ? '' : o.price, o.src, o.total == null ? '' : o.total, o.sup, o.status,
    o.note || '', (o.photos || []).join('\n'), o.by || '', o.handled ? 'כן' : '', o.del ? 'כן' : '', txt_(o.order || '')];
}
function findRow_(id) { const c = sh_(SH.R).getRange('A:A').createTextFinder(id).matchEntireCell(true).findNext(); return c ? c.getRow() : 0; }
function month_(m) { return rows_(SH.R).map(rec_).filter(r => r.ds.slice(0, 7) === m && !r.del); }

/* ---------- API (called from the page) ---------- */
function boot(pin) {
  const role = role_(pin);
  const today = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
  const sup = cget_('sups') || cput_('sups', rows_(SH.S).filter(r => String(r[0])).map(r => [String(r[0]), String(r[1] || '')]), 600);
  return JSON.stringify({role, today, products: products_(), suppliers: sup.map(r => String(r[0])), phones: Object.fromEntries(sup.map(r => [String(r[0]), String(r[1] || '')])),
    lists: lists_(addDays_(weekOf_(today), -7 * 10)), receipts: month_(today.slice(0, 7)), orders: orders_(addDays_(today, -21)),
    biz: settings_()['שם העסק'] || '', sheetUrl: role === 'manager' ? SpreadsheetApp.getActive().getUrl() : '', hasAI: role === 'manager' && !!settings_()['מפתח Gemini']});
}
function getMonth(pin, m) { role_(pin); return JSON.stringify(month_(m)); }
function getWeek(pin, w) { role_(pin); return JSON.stringify(lists_(w)[w] || {}); }

function saveReceipt(pin, json) {
  role_(pin); const o = JSON.parse(json);
  if (!o.pid || !(o.qty > 0)) throw new Error('חסרים נתונים');
  return withLock_(() => {
    const sh = sh_(SH.R);
    const r = o.id ? findRow_(o.id) : 0;
    if (r) {
      const old = rec_(sh.getRange(r, 1, 1, R_HEAD.length).getValues()[0]);
      o.handled = old.handled; o.t = old.t; o.ds = old.ds; if (old.del) o.del = true;
      sh.getRange(r, 1, 1, R_HEAD.length).setValues([row_(o)]);
    } else {
      o.id = o.id || ('r' + Utilities.getUuid().slice(0, 8));
      if (!/^\d{4}-\d{2}-\d{2}$/.test(o.ds || '')) o.ds = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
      if (!/^\d{1,2}:\d{2}$/.test(o.t || '')) o.t = Utilities.formatDate(new Date(), tz_(), 'HH:mm');
      sh.appendRow(row_(o));
    }
    if (o.order) orderStatus_(o.order);
    return JSON.stringify(o);
  });
}
function deleteReceipt(pin, id) {
  role_(pin);
  return withLock_(() => { const r = findRow_(id); if (!r) return true; sh_(SH.R).getRange(r, 17).setValue('כן');
    const oid = String(sh_(SH.R).getRange(r, 18).getValue() || ''); if (oid) orderStatus_(oid); return true; });
}
function setHandled(pin, id, val) {
  mgr_(pin);
  return withLock_(() => { const r = findRow_(id); if (r) sh_(SH.R).getRange(r, 16).setValue(val ? 'כן' : ''); return true; });
}
function uploadPhoto(pin, b64, mime, name) {
  role_(pin);
  const f = folder_('תמונות קבלות').createFile(Utilities.newBlob(Utilities.base64Decode(b64), mime || 'image/jpeg', name || ('photo-' + stamp_() + '.jpg')));
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return f.getUrl();
}

/* ---------- products & suppliers ---------- */
function saveProduct(pin, json) {
  mgr_(pin); const p = JSON.parse(json);
  return withLock_(() => {
    const sh = sh_(SH.P);
    const vals = [txt_(p.id || ('p' + Utilities.getUuid().slice(0, 6))), p.name, p.unit || 'ק"ג', TYPE_OUT[p.type] || 'שבועי', p.fixed == null ? '' : p.fixed, p.active === false ? 'לא' : 'כן'];
    const c = p.id ? sh.getRange('A:A').createTextFinder(p.id).matchEntireCell(true).findNext() : null;
    if (c) sh.getRange(c.getRow(), 1, 1, 6).setValues([vals]); else sh.appendRow(vals);
    cdel_('products');
    return JSON.stringify(products_());
  });
}
function saveSuppliers(pin, json) {
  mgr_(pin); const list = JSON.parse(json);
  return withLock_(() => {
    const sh = sh_(SH.S); if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
    if (list.length) sh.getRange(2, 1, list.length, 2).setValues(list.map(x => typeof x === 'string' ? [x, ''] : [x.name, x.phone || '']));
    cdel_('sups');
    return true;
  });
}

/* ---------- price lists ---------- */
function priceFiles(pin) {
  mgr_(pin);
  return JSON.stringify(rows_(SH.F).map(r => ({at: tstr_(r[0], 'dd/MM/yyyy HH:mm'), week: dstr_(r[1]), name: String(r[2]), url: String(r[3]), id: String(r[4])})).reverse().slice(0, 60));
}
function uploadPriceFile(pin, b64, mime, name) {
  mgr_(pin);
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), mime, name);
  const file = folder_('מחירונים').createFile(blob);
  withLock_(() => sh_(SH.F).appendRow([txt_(stamp_()), '', name, file.getUrl(), txt_(file.getId())]));
  const products = products_();
  // 1) קריאה חכמה (Gemini) – אם הוגדר מפתח
  let ai = null, aiErr = '';
  if (settings_()['מפתח Gemini']) { try { ai = aiRead_(b64, mime, products); } catch (e) { aiErr = String(e.message || e).slice(0, 300); } }
  // 2) זיהוי טקסט רגיל – תמיד, לבדיקה הצלבה
  let ocr = {date: null, prices: {}, extra: [], how: 'none'}, ocrErr = '';
  try { const text = ocr_(blob, mime); if (text) ocr = parse_(text, products); } catch (e) { ocrErr = String(e.message || e).slice(0, 300); }
  // 3) איחוד: Gemini קובע, הזיהוי הרגיל מאמת
  const prices = {}, flags = {}, conflicts = [];
  const src = ai ? ai.prices : {};
  Object.keys(src).forEach(id => { prices[id] = src[id]; flags[id] = ocr.prices[id] == null ? 'ai' : (ocr.prices[id] === src[id] ? 'ok' : 'conflict');
    if (flags[id] === 'conflict') conflicts.push({id, ai: src[id], ocr: ocr.prices[id]}); });
  Object.keys(ocr.prices).forEach(id => { if (prices[id] == null) { prices[id] = ocr.prices[id]; flags[id] = 'ocr'; } });
  const extra = ai ? ai.extra : ocr.extra;
  const date = (ai && ai.date) || ocr.date;
  const agreed = Object.keys(flags).filter(id => flags[id] === 'ok').length;
  const weeklyCount = products.filter(p => p.active && p.type === 'weekly').length;
  // שמירה אוטומטית רק כששתי השיטות מסכימות על הכל, אין מוצרים חדשים, ויש תאריך
  const auto = !!ai && conflicts.length === 0 && extra.length === 0 && !!date && agreed >= Math.max(5, Math.round(Object.keys(prices).length * 0.8));
  let saved = false, fixed = 0, week = '';
  if (auto) { week = weekOf_(date); fixed = savePriceList(pin, week, JSON.stringify(prices), file.getId()); saved = true; }
  return JSON.stringify({fileId: file.getId(), url: file.getUrl(), date, prices, extra, flags, conflicts, agreed, total: weeklyCount,
    how: Object.keys(prices).length ? 'names' : 'none', source: ai ? (ocr.how === 'names' ? 'ai+ocr' : 'ai') : 'ocr',
    aiErr, ocrErr, saved, week, fixed});
}

/* ---------- קריאה חכמה עם Gemini ---------- */
const AI_MODELS = ['gemini-flash-latest', 'gemini-3.6-flash', 'gemini-2.5-flash'];
function aiCall_(parts, schema) {
  const key = settings_()['מפתח Gemini']; if (!key) throw new Error('לא הוגדר מפתח Gemini');
  let last = '';
  for (const m of AI_MODELS) {
    const body = {contents: [{role: 'user', parts}], generationConfig: {responseMimeType: 'application/json'}};
    if (schema) body.generationConfig.responseSchema = schema;
    const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + m + ':generateContent', {
      method: 'post', contentType: 'application/json', headers: {'x-goog-api-key': key}, payload: JSON.stringify(body), muteHttpExceptions: true});
    const code = res.getResponseCode();
    if (code === 404) { last = m + ': 404'; continue; }
    if (code >= 300) throw new Error('Gemini ' + code + ' ' + res.getContentText().replace(/\s+/g, ' ').slice(0, 250));
    const j = JSON.parse(res.getContentText());
    const txt = ((((j.candidates || [])[0] || {}).content || {}).parts || []).filter(p => p.text && !p.thought).map(p => p.text).join('');
    if (!txt) throw new Error('Gemini לא החזיר תשובה');
    return JSON.parse(txt.replace(/^```(json)?|```$/g, '').trim());
  }
  throw new Error('לא נמצא מודל Gemini זמין (' + last + ')');
}
function aiRead_(b64, mime, products) {
  const weekly = products.filter(p => p.active && p.type === 'weekly');
  const names = weekly.map(p => p.name);
  const prompt = [
    'This image/PDF is a weekly wholesale fruit & vegetable price list from an Israeli supplier, in Hebrew.',
    'Read ONLY the main product table (the table with columns "מוצר" and "מחיר"). Ignore any other sections such as',
    '"ירקות עלים", "ירקות קלופים", oil/sugar boxes, phone numbers, notes and headers.',
    'For every row of that main table return: "name" exactly as printed, "price" as a number (as printed, e.g. 6.8),',
    'and "match": the name from OUR PRODUCT LIST below that is the very same product (same variety and packaging), or "" if none.',
    'Never guess a price. Never invent rows. Also return "date" of the price list as YYYY-MM-DD (a date like 16.11.26 is day.month.year), or "".',
    'OUR PRODUCT LIST:', names.join('\n')
  ].join('\n');
  const schema = {type: 'OBJECT', properties: {date: {type: 'STRING'}, items: {type: 'ARRAY', items: {type: 'OBJECT',
    properties: {name: {type: 'STRING'}, price: {type: 'NUMBER'}, match: {type: 'STRING'}}, required: ['name', 'price', 'match']}}}, required: ['date', 'items']};
  const r = aiCall_([{inline_data: {mime_type: mime, data: b64}}, {text: prompt}], schema);
  const byName = {}; weekly.forEach(p => byName[norm_(p.name)] = p.id);
  const prices = {}, extra = [];
  (r.items || []).forEach(it => {
    const pr = Number(it.price); if (!(pr > 0.3 && pr < 1000)) return;
    const id = it.match ? byName[norm_(it.match)] : null;
    if (id) { if (prices[id] == null) prices[id] = Math.round(pr * 100) / 100; }
    else if (it.name) extra.push({name: norm_(it.name), price: Math.round(pr * 100) / 100});
  });
  const date = /^\d{4}-\d{2}-\d{2}$/.test(r.date || '') ? r.date : null;
  return {date, prices, extra: extra.slice(0, 30)};
}
/** שמירת מפתח Gemini ובדיקה שהוא עובד */
function setAiKey(pin, key) {
  mgr_(pin); key = String(key || '').trim();
  withLock_(() => setSetting_('מפתח Gemini', key));
  if (!key) return JSON.stringify({ok: true, off: true});
  try { aiCall_([{text: 'Reply with the JSON {"ok": true}'}]); return JSON.stringify({ok: true}); }
  catch (e) { return JSON.stringify({ok: false, error: String(e.message || e).slice(0, 300)}); }
}
/** Google Drive OCR: converts the PDF/image to a temporary Google Doc and reads its text. */
function ocr_(blob, mime) {
  // שתי דרכים להמרה עם זיהוי טקסט; אם הראשונה נכשלת מנסים את השנייה
  const tries = [
    {url: 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id', meta: {name: 'ocr-temp-' + Date.now(), mimeType: 'application/vnd.google-apps.document'}},
    {url: 'https://www.googleapis.com/upload/drive/v2/files?uploadType=multipart&convert=true&ocr=true', meta: {title: 'ocr-temp-' + Date.now(), mimeType: mime}}
  ];
  let lastErr = '';
  for (const t of tries) {
    const boundary = 'b' + Date.now();
    const head = Utilities.newBlob('--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + JSON.stringify(t.meta) + '\r\n--' + boundary + '\r\nContent-Type: ' + mime + '\r\n\r\n').getBytes();
    const tail = Utilities.newBlob('\r\n--' + boundary + '--').getBytes();
    const res = UrlFetchApp.fetch(t.url, {method: 'post', contentType: 'multipart/related; boundary=' + boundary, payload: head.concat(blob.getBytes()).concat(tail),
      headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}, muteHttpExceptions: true});
    if (res.getResponseCode() >= 300) { lastErr = 'OCR ' + res.getResponseCode() + ' ' + res.getContentText().replace(/\s+/g, ' ').slice(0, 300); continue; }
    const id = JSON.parse(res.getContentText()).id;
    try { return DocumentApp.openById(id).getBody().getText(); }
    catch (e) { lastErr = String(e.message || e); }
    finally { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} }
  }
  throw new Error(lastErr || 'OCR נכשל');
}
function norm_(s) {
  return String(s).replace(/[״“”]/g, '"').replace(/[׳’`]/g, "'").replace(/[‎‏‪-‮﻿]/g, '').replace(/\s+/g, ' ').trim();
}
/** מילים "נקיות" להשוואה: בלי גרשיים, סוגריים וסימנים */
function words_(s) { return norm_(s).replace(/["']/g, '').replace(/(\d)[.,](\d)/g, '$1#$2').replace(/[()\[\]{}.,:;\-–]/g, ' ').replace(/#/g, '.').split(' ').filter(String); }
function lev_(a, b) {
  const m = a.length, n = b.length, d = []; for (let i = 0; i <= m; i++) { d[i] = [i]; }
  for (let j = 1; j <= n; j++) d[0][j] = j;
  for (let i = 1; i <= m; i++) for (let j = 1; j <= n; j++) d[i][j] = Math.min(d[i-1][j] + 1, d[i][j-1] + 1, d[i-1][j-1] + (a[i-1] === b[j-1] ? 0 : 1));
  return d[m][n];
}
/** מפתח השוואה: אותיות סופיות לרגילות, בלי ו' ו-י' (ה-OCR מתבלבל בהן ובגרשיים) */
function key_(w) { return w.replace(/ן/g, 'נ').replace(/ם/g, 'מ').replace(/ץ/g, 'צ').replace(/ך/g, 'כ').replace(/ף/g, 'פ').replace(/[וי]/g, ''); }
function wordEq_(a, b, multi) {
  if (a === b) return true;
  if (/\d/.test(a) || /\d/.test(b)) return a.replace(/\.$/, '') === b.replace(/\.$/, '');
  const x = key_(a), y = key_(b), L = Math.max(x.length, y.length);
  if (x === y) return L > 0;
  // מוצר של מילה אחת: קפדני יותר, כדי ש"פומלו" לא ייקרא "מלון"
  const tol = multi ? (L >= 6 ? 2 : L >= 3 ? 1 : 0) : (L >= 6 ? 1 : 0);
  return lev_(x, y) <= tol;
}
const NUM_LINE = /^₪?\s*(\d{1,3}(?:[.,]\d{1,2})?)\s*₪?$/;
const HEADER_LINE = /^(מוצר|מוצרים|שם מוצר|פריט|מחיר|מחיר ליחידה)$/;
const DATE_IN = /\d{1,2}[.\/]\d{1,2}[.\/]\d{2,4}/;
/**
 * קורא את טבלת המוצרים הראשית במחירון.
 * ה-OCR של גוגל מחזיר את הטבלה בקטעים: רצף של שמות ואחריו רצף של מחירים באותו סדר,
 * ולפעמים שורה של "שם + מחיר". מצמידים כל קטע שמות לקטע המחירים שאחריו, רק כשהמספרים תואמים.
 * קטעים בלי אף מוצר מוכר (עלים, קלופים, שמן) נדחים, אלא אם הם צמודים לטבלה הראשית (מוצרים חדשים).
 */
function parse_(text, products) {
  const lines = text.split(/\n/).map(norm_).filter(String);
  let date = null;
  const dm = norm_(text).match(/(\d{1,2})[.\/](\d{1,2})[.\/](\d{2,4})/);
  if (dm) { const y = dm[3].length === 2 ? '20' + dm[3] : dm[3]; date = y + '-' + ('0' + dm[2]).slice(-2) + '-' + ('0' + dm[1]).slice(-2); }
  const weekly = products.filter(p => p.active && p.type === 'weekly');
  const cands = weekly.map(p => ({p, w: words_(p.name)}));
  const isNum = w => /^\d{1,3}(?:\.\d{1,2})?$/.test(w);
  const segment = line => {
    const w = words_(line).filter(x => x !== '₪'); const out = []; let k = 0, unk = [];
    const flushU = () => { if (unk.length) { out.push({name: unk.join(' ')}); unk = []; } };
    while (k < w.length) {
      let hit = null, hd = 1e9;
      for (const c of cands) {
        if (c.w.length > w.length - k || (hit && c.w.length < hit.w.length)) continue;
        if (!c.w.every((x, t) => wordEq_(x, w[k + t], c.w.length > 1))) continue;
        const d = c.w.reduce((a, x, t) => a + lev_(key_(x), key_(w[k + t])), 0);
        if (!hit || c.w.length > hit.w.length || d < hd) { hit = c; hd = d; }
      }
      if (hit) { flushU(); out.push({p: hit.p}); k += hit.w.length; }
      else if (isNum(w[k])) { flushU(); out.push({num: Number(w[k])}); k++; }
      else { unk.push(w[k]); k++; }
    }
    flushU(); return out;
  };
  const prices = {}, extra = [];
  let buf = [], nums = [], adj = false, segs = 0, pend = null;
  const kn = arr => arr.filter(i => i.p).length;
  const okMix = arr => arr.length > 0 && kn(arr) > 0 && arr.length - kn(arr) <= 3;   // בעיקר מוצרים מוכרים, עד 3 חדשים
  const commit = (items, vals) => {
    items.forEach((it, i) => { if (it.p) prices[it.p.id] = vals[i]; else if (it.name) extra.push({name: it.name, price: vals[i]}); });
    adj = true; segs++;
  };
  const dropLead = arr => { const x = arr.slice(); while (x.length && !x[0].p && x.some(i => i.p)) x.shift(); return x; };
  // הטבלה יכולה להגיע בכמה חלקים: שמות, מחירים, עוד שמות, עוד מחירים. אוספים עד שהמספרים מתאזנים.
  const endPend = () => {
    if (!pend) return;
    if (buf.length && okMix(buf)) { pend.items = pend.items.concat(buf); buf = []; }
    if (pend.items.length === pend.vals.length) commit(pend.items, pend.vals); else adj = false;
    pend = null;
  };
  const pair = () => {
    if (pend) {
      if (!buf.length || okMix(buf)) { pend.items = pend.items.concat(buf); pend.vals = pend.vals.concat(nums); buf = []; nums = [];
        if (pend.items.length === pend.vals.length) { commit(pend.items, pend.vals); pend = null; }
        else if (Math.abs(pend.items.length - pend.vals.length) > 40) { pend = null; adj = false; }
        return; }
      endPend();
    }
    const items = dropLead(buf);
    if (okMix(items)) {
      if (items.length === nums.length) commit(items, nums); else pend = {items, vals: nums};
    } else if (items.length && kn(items) === 0 && adj) {
      let it = items;   // מוצרים חדשים צמודים לטבלה; אם חסרים פריטים, אולי שתי מילים הן שני מוצרים
      if (it.length !== nums.length) { const split = []; it.forEach(x => x.name.split(' ').forEach(w => split.push({name: w}))); if (split.length === nums.length) it = split; }
      if (it.length === nums.length) commit(it, nums); else adj = false;
    } else adj = false;
    buf = []; nums = [];
  };
  for (const L of lines) {
    if (NUM_LINE.test(L)) { nums.push(Number(L.match(NUM_LINE)[1].replace(',', '.'))); continue; }
    if (!L.replace(/[₪\s.,:;|_\-–]/g, '')) continue;
    if (nums.length) pair();
    if (HEADER_LINE.test(L)) continue;
    if (DATE_IN.test(L) || !/[֐-׿]/.test(L)) { endPend(); buf = []; adj = false; continue; }
    const items = segment(L), last = items[items.length - 1];
    if (last && last.num != null && items.length >= 2 && items.slice(0, -1).every(i => i.num == null)) {
      endPend();
      const it = items[items.length - 2]; buf = [];
      if (it.p) { prices[it.p.id] = last.num; adj = true; segs++; }
      else if (adj && it.name && items.length === 2) extra.push({name: it.name, price: last.num});
      continue;
    }
    if (items.some(i => i.num != null)) { endPend(); buf = []; adj = false; continue; }
    buf.push.apply(buf, items);
  }
  if (nums.length) pair();
  endPend();
  const seen = {}; const ex = extra.filter(x => !seen[x.name] && (seen[x.name] = 1)).slice(0, 30);
  return {date, prices, extra: ex, how: segs ? 'names' : 'none'};
}
function savePriceList(pin, week, json, fileId) {
  mgr_(pin); const prices = JSON.parse(json); const pmap = {}; products_().forEach(p => pmap[p.id] = p);
  return withLock_(() => {
    const sh = sh_(SH.L), now = stamp_();
    const keep = rows_(SH.L).filter(r => dstr_(r[0]) !== week).map(r => [txt_(dstr_(r[0])), txt_(String(r[1])), r[2], r[3], txt_(tstr_(r[4], 'yyyy-MM-dd HH:mm'))]);
    const add = Object.keys(prices).filter(id => pmap[id]).map(id => [txt_(week), txt_(id), pmap[id].name, Number(prices[id]), txt_(now)]);
    const all = keep.concat(add);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 5).clearContent();
    if (all.length) sh.getRange(2, 1, all.length, 5).setValues(all);
    cdel_('lists');
    if (fileId) { const c = sh_(SH.F).getRange('E:E').createTextFinder(fileId).matchEntireCell(true).findNext(); if (c) sh_(SH.F).getRange(c.getRow(), 2).setValue(txt_(week)); }
    const rs = sh_(SH.R), data = rs.getDataRange().getValues(), end = addDays_(week, 6); let fixed = 0;
    for (let i = 1; i < data.length; i++) {
      const r = data[i], ds = dstr_(r[1]);
      if (ds >= week && ds <= end && r[7] === '' && r[16] !== 'כן' && prices[String(r[3])] != null) {
        const pr = Number(prices[String(r[3])]);
        rs.getRange(i + 1, 8, 1, 3).setValues([[pr, 'מחירון', Math.round(Number(r[6]) * pr * 100) / 100]]); fixed++;
      }
    }
    return fixed;
  });
}


/* ---------- API לאפליקציה (מתארחת מחוץ ל-Google) ---------- */
const API = {boot, getMonth, getWeek, saveReceipt, deleteReceipt, setHandled, uploadPhoto, saveProduct, saveSuppliers,
  priceFiles, uploadPriceFile, savePriceList, getOrders, saveOrder, setOrderStatus, setAiKey};
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const f = Object.prototype.hasOwnProperty.call(API, req.fn) ? API[req.fn] : null;
    if (!f) throw new Error('פעולה לא מוכרת');
    if (!cget_('ready')) { setup_(); ensure_(); cput_('ready', 1, 21600); }
    out = {ok: true, data: f.apply(null, req.args || [])};
  } catch (err) { out = {ok: false, error: String(err && err.message || err)}; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- orders ---------- */
function ordRec_(r) {
  let items = []; try { items = JSON.parse(r[8] || '[]'); } catch (e) {}
  return {id: String(r[0]), at: tstr_(r[1], 'yyyy-MM-dd HH:mm'), ds: dstr_(r[2]), sup: String(r[3]), status: String(r[4]), note: String(r[6] || ''), by: String(r[7] || ''), items};
}
function ordRow_(o) {
  return [txt_(o.id), txt_(o.at), txt_(o.ds), o.sup, o.status, o.items.map(i => i.name + ' ' + i.qty + ' ' + (i.unit || '')).join(', '), o.note || '', o.by || '', JSON.stringify(o.items)];
}
function received_() {
  const got = {};
  rows_(SH.R).forEach(r => { const oid = String(r[17] || ''); if (!oid || r[16] === 'כן') return; const g = got[oid] = got[oid] || {}; g[String(r[3])] = (g[String(r[3])] || 0) + Number(r[6] || 0); });
  return got;
}
function orders_(from, month) {
  const got = received_();
  return rows_(SH.O).filter(r => r[0] !== '').map(ordRec_)
    .filter(o => month ? o.ds.slice(0, 7) === month : (o.ds >= from || O_OPEN.indexOf(o.status) >= 0))
    .map(o => (o.got = got[o.id] || {}, o)).sort((a, b) => a.ds < b.ds ? 1 : -1);
}
function getOrders(pin, month) { role_(pin); return JSON.stringify(orders_('', month)); }
function saveOrder(pin, json) {
  chef_(pin); const o = JSON.parse(json);
  o.items = (o.items || []).filter(i => i.pid && Number(i.qty) > 0).map(i => ({pid: String(i.pid), name: String(i.name), unit: String(i.unit || ''), qty: Number(i.qty)}));
  if (!o.sup) throw new Error('נא לבחור ספק');
  if (!o.items.length) throw new Error('ההזמנה ריקה');
  return withLock_(() => {
    const sh = sh_(SH.O);
    const c = o.id ? sh.getRange('A:A').createTextFinder(o.id).matchEntireCell(true).findNext() : null;
    if (c) { const old = ordRec_(sh.getRange(c.getRow(), 1, 1, O_HEAD.length).getValues()[0]); o.at = old.at; o.by = old.by || o.by;
      if (O_OPEN.indexOf(o.status) < 0 && o.status !== 'בוטלה') o.status = old.status;
      sh.getRange(c.getRow(), 1, 1, O_HEAD.length).setValues([ordRow_(o)]); }
    else { o.id = o.id || ('o' + Utilities.getUuid().slice(0, 7)); o.at = stamp_(); o.status = o.status || 'טיוטה'; sh.appendRow(ordRow_(o)); }
    orderStatus_(o.id);
    return JSON.stringify(orders_('', '').find(x => x.id === o.id) || o);
  });
}
function setOrderStatus(pin, id, st) {
  chef_(pin);
  return withLock_(() => { const c = sh_(SH.O).getRange('A:A').createTextFinder(id).matchEntireCell(true).findNext(); if (c) sh_(SH.O).getRange(c.getRow(), 5).setValue(st); return true; });
}
/** מעדכן סטטוס הזמנה לפי מה שהתקבל בפועל */
function orderStatus_(oid) {
  const sh = sh_(SH.O), c = sh.getRange('A:A').createTextFinder(oid).matchEntireCell(true).findNext(); if (!c) return;
  const o = ordRec_(sh.getRange(c.getRow(), 1, 1, O_HEAD.length).getValues()[0]); if (o.status === 'בוטלה') return;
  const g = received_()[oid] || {};
  const any = Object.keys(g).length > 0, all = o.items.every(i => (g[i.pid] || 0) >= i.qty - 1e-9);
  const st = any ? (all ? 'התקבלה' : 'התקבלה חלקית') : (o.status === 'טיוטה' ? 'טיוטה' : 'נשלחה');
  if (st !== o.status) sh.getRange(c.getRow(), 5).setValue(st);
}

/* ---------- שמירה על השרת "ער" בשעות העבודה ---------- */
function חימום() {
  const h = Number(Utilities.formatDate(new Date(), tz_(), 'H'));
  if (h < 5 || h > 22) return;
  settings_(); products_();
}
/** להרצה ידנית פעם אחת מהעורך: מפעיל חימום אוטומטי כל 5 דקות */
function הפעלת_מהירות() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'חימום').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('חימום').timeBased().everyMinutes(5).create();
  cdel_('settings', 'products', 'lists', 'sups', 'ready');
  Logger.log('הופעל: השרת יישאר ער בין 05:00 ל-23:00');
}
