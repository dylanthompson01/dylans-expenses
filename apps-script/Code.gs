/**
 * Dylan's Expenses — Google Apps Script backend (v2)
 * ---------------------------------------------------
 * Paste this whole file into the Apps Script editor bound to the expenses
 * spreadsheet, then Deploy > Manage deployments > Edit (pencil) >
 * Version: New version > Deploy. Editing the existing deployment keeps the
 * same /exec URL, so index.html doesn't need to change.
 *
 * The web app must be deployed as "Execute as: Me" and
 * "Who has access: Anyone" so the browser can reach it without a Google login.
 * That is why every request now has to carry the passcode: the URL is public.
 *
 * ONE-TIME SETUP — Project Settings (gear) > Script Properties:
 *   GEMINI_KEY  = your Gemini API key          (already set)
 *   PASSCODE    = what you'll type in the app  (6+ characters)
 *   SHARE_KEY   = created automatically the first time you share the
 *                 reimbursement page. Delete it to kill every old link.
 *
 * Column order is defined once, in HEADERS. index.html reads and writes
 * objects, never column positions, so only this file knows the order.
 */

var API_VERSION = 2;

var SHEET_ID    = '193exXTS041Quq_-VOS_0uXMMAv6RM6eFuKHMgkUkIXo';
var FOLDER_ID   = '1VtL_8OtIyHmute0nbAcr0LgCnoE6QLbV';
var TAB_NAME    = 'Expenses';
var BUDGET_TAB  = 'Budgets';
var DELETED_TAB = 'Deleted';

var HEADERS = [
  'Timestamp',
  'Place / Vendor',
  'Expense Category',
  'Date',
  'Amount',
  'Semester',
  'Receipt Type',
  'Receipt Drive Link',
  'Do You Need Reimbursement',
  'Have You Been Reimbursed',
  'Additional Info',
  'Is This A 529 Expense',
  'ID',                       // stable row id, so edits and deletes hit the right row
  'Bank Txn ID',              // Plaid transaction_id — proof the charge happened, and dedupe key
  'Account',                  // e.g. "Chase •4321"
  'Bank Category',            // Plaid's own label, e.g. FOOD_AND_DRINK_GROCERIES
  'Sorted By',                // why it's 529 or not: You / Your history / Bank category / Store name / AI / Receipt
  'Needs Review'              // "Yes" when the app couldn't decide 529 on its own
];
var ID_COL = 13;              // 1-based column of 'ID'
var COL = { txn: 14, account: 15, bankCat: 16, sortedBy: 17, review: 18 };   // 1-based

// Allowed values, used to build the in-sheet dropdowns and to validate edits
var CATEGORIES    = ['Food & Dining','Tuition & Fees','Books & Supplies','Housing',
                     'Transportation','Technology','Healthcare','Entertainment',
                     'Clothing','Personal Care','School Supplies','Other'];
var RECEIPT_TYPES = ['Receipt','Bank Statement','Bank Transaction'];
var YES_NO        = ['Yes','No'];
var YES_NO_MAYBE  = ['Yes','No','Not Sure'];

/** Spring 2025 through two years from now, so the list never runs out. */
function semesterList() {
  var out = [], endYear = new Date().getFullYear() + 2;
  for (var y = 2025; y <= endYear; y++) out.push('Spring ' + y, 'Summer ' + y, 'Fall ' + y);
  return out;
}

function newId() {
  return 'e' + Utilities.getUuid().replace(/-/g, '').slice(0, 12);
}
/** Accept an id from the browser only if it looks like one of ours. */
function cleanId(v) {
  v = String(v || '');
  return /^[A-Za-z0-9_-]{6,40}$/.test(v) ? v : '';
}

/** Object -> row, in HEADERS order. */
function toRow(e) {
  return [
    e.timestamp || new Date().toISOString(),
    e.place || 'Unknown',
    e.category || 'Other',
    e.date || '',
    Number(e.amount) || 0,
    e.semester || '',
    e.receiptType || 'Receipt',
    e.receiptUrl || '',
    e.needsReimb || 'No',
    e.reimbursed || 'No',
    e.notes || '',
    e.is529 || 'Not Sure',
    e.id || newId(),
    e.bankTxnId || '',
    e.account || '',
    e.bankCategory || '',
    e.sortedBy || '',
    e.review || ''
  ];
}

/** Row -> object, in HEADERS order. */
function toObject(row, index) {
  return {
    id:          row[12] ? String(row[12]) : 'sheet-' + index,
    timestamp:   row[0] ? String(row[0]) : '',
    place:       row[1] ? String(row[1]) : 'Unknown',
    category:    row[2] ? String(row[2]) : 'Other',
    date:        formatDate(row[3]),
    amount:      Number(row[4]) || 0,
    semester:    row[5] ? String(row[5]) : '',
    receiptType: row[6] ? String(row[6]) : 'Receipt',
    receiptUrl:  row[7] ? String(row[7]) : '',
    needsReimb:  row[8] ? String(row[8]) : 'No',
    reimbursed:  row[9] ? String(row[9]) : 'No',
    notes:       row[10] ? String(row[10]) : '',
    is529:       row[11] ? String(row[11]) : 'Not Sure',
    bankTxnId:   row[13] ? String(row[13]) : '',
    account:     row[14] ? String(row[14]) : '',
    bankCategory: row[15] ? String(row[15]) : '',
    sortedBy:    row[16] ? String(row[16]) : '',
    review:      row[17] ? String(row[17]) : '',
    savedToGoogle: true
  };
}

/** Sheets may hand back a Date object or a string; the app wants YYYY-MM-DD. */
function formatDate(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, 'America/New_York', 'yyyy-MM-dd');
  }
  var s = String(v).trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);   // M/D/YYYY
  if (m) {
    return m[3] + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
  }
  return s.slice(0, 10);
}

// Opening the spreadsheet is one of the slowest calls in Apps Script, so
// each request opens it once and reuses it.
var _book = null;
function book() { return _book || (_book = SpreadsheetApp.openById(SHEET_ID)); }

function getSheet() {
  var ss = book();
  var sh = ss.getSheetByName(TAB_NAME) || ss.getSheets()[0];
  if (sh.getMaxColumns() < HEADERS.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), HEADERS.length - sh.getMaxColumns());
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    applyFormatting(sh);
  }
  return sh;
}

/** One writer at a time, so two saves can't shift each other's rows.
    Every write goes through here, so it's also where the cached copy of
    the data is invalidated. */
function withLock(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { var out = fn(); bumpData(); return out; } finally { lock.releaseLock(); }
}

/* ===================================================================
   LOAD CACHE

   Reading 500+ rows from the Sheet takes about a second. The result is
   kept in CacheService under the current data version; any write through
   the app bumps the version, so a stale copy is never served. Edits typed
   straight into the Sheet show up within 5 minutes, or at once with the
   app's Refresh button (which asks for fresh data).
   =================================================================== */

function dataVersion() { return PropertiesService.getScriptProperties().getProperty('DATA_VER') || '0'; }
function bumpData() {
  PropertiesService.getScriptProperties().setProperty('DATA_VER', Date.now().toString(36) + Math.random().toString(36).slice(2, 6));
}
var CACHE_CHUNK = 30000;          // characters; stays under the 100 KB per-key limit
function cachePutLoad(ver, data) {
  try {
    var s = JSON.stringify(data), parts = {}, n = Math.ceil(s.length / CACHE_CHUNK) || 1;
    for (var i = 0; i < n; i++) parts['L:' + ver + ':' + i] = s.slice(i * CACHE_CHUNK, (i + 1) * CACHE_CHUNK);
    parts['L:' + ver] = String(n);
    CacheService.getScriptCache().putAll(parts, 300);
  } catch (e) { /* too big or cache unavailable: just don't cache */ }
}
function cacheGetLoad(ver) {
  try {
    var c = CacheService.getScriptCache(), n = Number(c.get('L:' + ver) || 0);
    if (!n) return null;
    var keys = [];
    for (var i = 0; i < n; i++) keys.push('L:' + ver + ':' + i);
    var got = c.getAll(keys), s = '';
    for (var k = 0; k < keys.length; k++) { if (got[keys[k]] == null) return null; s += got[keys[k]]; }
    return JSON.parse(s);
  } catch (e) { return null; }
}

// Compact load format: field names once, then one array per expense.
// About half the size of sending an object per row.
var FIELDS = ['id', 'timestamp', 'place', 'category', 'date', 'amount', 'semester', 'receiptType',
              'receiptUrl', 'needsReimb', 'reimbursed', 'notes', 'is529', 'bankTxnId', 'account',
              'bankCategory', 'sortedBy', 'review'];

/* ===================================================================
   PASSCODE

   The /exec URL is in the page source, and the page is public. Without a
   check here, anyone could download every expense and receipt link, or run
   scans billed to the Gemini key. Every request except the reimbursement
   share link must carry the passcode.

   After 20 wrong guesses in 15 minutes, everything is refused until the
   window passes, which makes guessing a short passcode impractical.
   =================================================================== */

function checkKey(key) {
  var pass = PropertiesService.getScriptProperties().getProperty('PASSCODE');
  if (!pass) return 'PASSCODE is not set. Apps Script > Project Settings > Script Properties > add PASSCODE.';
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('auth_fails') || 0);
  if (fails >= 20) return 'Too many wrong passcodes. Try again in 15 minutes.';
  if (String(key || '') === String(pass)) return '';
  cache.put('auth_fails', String(fails + 1), 900);
  return key ? 'Wrong passcode' : 'Passcode required';
}

function lockedResponse(msg) {
  return json({ success: false, locked: true, error: msg, apiVersion: API_VERSION });
}

/* ===================================================================
   READ
   =================================================================== */

/** GET ?action=load&key=…  — every row as JSON, plus budgets.
    GET ?action=reimb&share=… — the read-only reimbursement list. */
function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.action === 'reimb') return json(shareView(p.share));
    var err = checkKey(p.key);
    if (err) return lockedResponse(err);
    return json(loadAll(p.compact === '1', p.fresh === '1'));
  } catch (err2) {
    return json({ success: false, error: String(err2), apiVersion: API_VERSION });
  }
}

/**
 * Everything the app needs on open: expenses, budgets and bank status,
 * in one request.
 * @param compact  send FIELDS + rows instead of one object per expense
 * @param fresh    skip the cache (the app's Refresh button)
 */
function loadAll(compact, fresh) {
  var ver = dataVersion();
  var data = fresh ? null : cacheGetLoad(ver);
  var cached = !!data;
  if (!data) {
    var sh = getSheet();
    var values = readValues(sh);                      // one read, not three
    if (valuesMissingIds(values)) { ensureIds(sh); values = readValues(sh); }
    ver = dataVersion();                              // ensureIds may have written
    var objs = valuesToObjects(values);
    data = { rows: objs.map(function (o) { return FIELDS.map(function (f) { return o[f]; }); }),
             budgets: readBudgets() };
    cachePutLoad(ver, data);
  }
  var out = { success: true, apiVersion: API_VERSION, count: data.rows.length, budgets: data.budgets,
              cached: cached, bank: bankStatus() };
  if (compact) { out.fields = FIELDS; out.rows = data.rows; }
  else {
    out.expenses = data.rows.map(function (r) {
      var o = { savedToGoogle: true };
      for (var i = 0; i < FIELDS.length; i++) o[FIELDS[i]] = r[i];
      return o;
    });
  }
  return out;
}

function readValues(sh) {
  var last = sh.getLastRow();
  return last > 1 ? sh.getRange(2, 1, last - 1, HEADERS.length).getValues() : [];
}
function valuesToObjects(values) {
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (!row[1] && !row[4]) continue;            // skip blank rows
    out.push(toObject(row, i));
  }
  return out;
}
function valuesMissingIds(values) {
  for (var i = 0; i < values.length; i++) {
    if (!values[i][ID_COL - 1] && (values[i][1] || values[i][4])) return true;
  }
  return false;
}
function readExpenses(sh) { return valuesToObjects(readValues(sh)); }

/** Give every row an ID the first time it's seen — including rows typed
    straight into the sheet. Runs under the lock and re-reads inside it, so
    two devices loading at once can't hand out different IDs for one row. */
function ensureIds(sh) {
  if (sh.getLastRow() < 2) return;
  if (!missingIds(sh)) return;
  withLock(function () {
    var last = sh.getLastRow();
    var head = sh.getRange(1, ID_COL);
    if (head.getValue() !== 'ID') {
      head.setValue('ID').setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#3A6644');
    }
    var idRange = sh.getRange(2, ID_COL, last - 1, 1);
    var ids  = idRange.getValues();
    var data = sh.getRange(2, 1, last - 1, 5).getValues();
    var changed = false;
    for (var i = 0; i < ids.length; i++) {
      if (!ids[i][0] && (data[i][1] || data[i][4])) { ids[i][0] = newId(); changed = true; }
    }
    if (changed) idRange.setValues(ids);
  });
}

function missingIds(sh) {
  var last = sh.getLastRow();
  var ids  = sh.getRange(2, ID_COL, last - 1, 1).getValues();
  var data = sh.getRange(2, 1, last - 1, 5).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (!ids[i][0] && (data[i][1] || data[i][4])) return true;
  }
  return false;
}

/** Row number (1-based) for an ID, or -1. */
function findRow(sh, id) {
  var last = sh.getLastRow();
  if (last < 2 || !id) return -1;
  var ids = sh.getRange(2, ID_COL, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

/* ===================================================================
   GEMINI PROXY

   The API key lives in Script Properties, never in the page. A key shipped
   to the browser is public by definition — Google's scanner disabled the
   previous one within hours of the repo going public. Here the browser
   sends an image and gets back parsed JSON; it never sees the key.
   =================================================================== */

var GEMINI_PRIMARY  = 'gemini-2.5-flash-lite';
var GEMINI_FALLBACK = 'gemini-2.0-flash-lite';

var SCAN_PROMPT =
'You are an expense parser for a college student\'s 529 expense tracker. The image is EITHER a photo of a purchase receipt OR a screenshot of a bank/credit-card statement. Return ONLY a raw JSON object - no markdown, no backticks, no extra text:\n' +
'{"receiptType":"Receipt","place":"store or vendor name","amount":12.34,"date":"YYYY-MM-DD","category":"Food & Dining","is529":"Yes","notes":"brief note max 60 chars"}\n\n' +
'STEP 1 - Identify the document, set "receiptType" to exactly one of:\n' +
'- "Receipt" - an itemised store receipt, invoice, or order confirmation. Has line items, subtotal/tax/total, a store header.\n' +
'- "Bank Statement" - a screenshot of a banking or card app / statement. Shows a list of transactions with dates and amounts, a running balance, or a bank interface.\n\n' +
'STEP 2 - Extract the purchase.\n' +
'- Receipt: use the vendor name from the header and the FINAL total (after tax and discounts). Never the subtotal.\n' +
'- Bank Statement: if one transaction is highlighted or clearly the subject, use it. Otherwise use the single largest non-payment, non-transfer purchase. Ignore payments, transfers, deposits, refunds, interest and fees. Clean the merchant name - "SQ *JOES COFFEE 04412" becomes "Joe\'s Coffee", "AMZN Mktp US*2B4XY" becomes "Amazon".\n\n' +
'STEP 3 - Categorise. Pick the single most specific match:\n' +
'Food & Dining, Tuition & Fees, Books & Supplies, Housing, Transportation, Technology, Healthcare, Entertainment, Clothing, Personal Care, School Supplies, Other.\n' +
'- Groceries AND restaurants both go under "Food & Dining" - the app splits them later using the vendor name, so the vendor name must be accurate.\n' +
'- Rent, utilities, electricity, water go under "Housing".\n' +
'- Laptops, software, phone/internet service go under "Technology".\n' +
'- Gas, Uber, parking, flights, car repair go under "Transportation".\n\n' +
'STEP 4 - Set "is529" using IRS Qualified Higher Education Expense rules:\n' +
'- "Yes": tuition and mandatory fees; required books, supplies and equipment; technology; room and board (rent, utilities, groceries, meal plans); special-needs equipment; registered apprenticeship costs; student loan repayments; professional certification and licensing costs.\n' +
'- "No": restaurants and eating out; transportation, travel and parking; health insurance and medical fees; clothing; personal care; sports or club fees not required for coursework.\n' +
'- "Not Sure": genuinely ambiguous only.\n\n' +
'Notes: what was actually bought, max 60 chars. Unreadable date - use today. Unreadable amount - use 0.';

function geminiKey() {
  var k = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  if (!k) throw new Error('GEMINI_KEY is not set. Project Settings > Script Properties > add GEMINI_KEY.');
  return k;
}

/** Confirms the key is present and working, without printing it. */
function testGeminiKey() {
  var k = geminiKey();
  var resp = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models?key=' + encodeURIComponent(k),
    { muteHttpExceptions: true });
  var out = { code: resp.getResponseCode(), keyLength: k.length,
              ok: resp.getResponseCode() === 200 };
  if (!out.ok) out.body = resp.getContentText().slice(0, 300);
  Logger.log(JSON.stringify(out, null, 2));
  return out;
}

/** Images arrive as JPEG; a PDF (an emailed receipt) is passed through as-is. */
function scanMime(mime) {
  return mime === 'application/pdf' ? 'application/pdf' : 'image/jpeg';
}

function callGemini(model, b64, mime) {
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model +
            ':generateContent?key=' + encodeURIComponent(geminiKey());
  var body = {
    contents: [{ parts: [ { inline_data: { mime_type: scanMime(mime), data: b64 } },
                          { text: SCAN_PROMPT } ] }],
    generationConfig: { temperature: 0.1, maxOutputTokens: 400 }
  };
  var resp = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(body), muteHttpExceptions: true
  });
  var code = resp.getResponseCode(), text = resp.getContentText();

  if (code !== 200) {
    var retryable = code === 429 || code === 503 ||
      /overload|high demand|quota|rate|resource_exhausted/i.test(text);
    return { ok: false, retryable: retryable, error: 'HTTP ' + code + ' ' + text.slice(0, 180) };
  }
  var data = JSON.parse(text);
  var parts = (((data.candidates || [])[0] || {}).content || {}).parts || [];
  var out = '';
  for (var i = 0; i < parts.length; i++) {
    if (!parts[i].thought && parts[i].text) out += parts[i].text;
  }
  var m = out.match(/\{[\s\S]*?\}/);
  if (!m) return { ok: false, retryable: false, error: 'No JSON in model output: ' + out.slice(0, 140) };
  try { return { ok: true, data: JSON.parse(m[0]) }; }
  catch (parseErr) { return { ok: false, retryable: false, error: 'Bad JSON: ' + m[0].slice(0, 140) }; }
}

/** Primary model with retries, then the fallback model. */
function scanReceipt(b64, mime) {
  if (!b64) return { success: false, error: 'No image supplied' };
  var models = [GEMINI_PRIMARY, GEMINI_FALLBACK], lastErr = 'unknown';
  for (var m = 0; m < models.length; m++) {
    for (var attempt = 1; attempt <= 3; attempt++) {
      var res = callGemini(models[m], b64, mime);
      if (res.ok) return { success: true, data: res.data, model: models[m] };
      lastErr = res.error;
      if (!res.retryable) break;          // auth/bad-request: next model, don't hammer
      Utilities.sleep(attempt * 2000);
    }
  }
  return { success: false, error: lastErr };
}

/* ===================================================================
   WRITE
   =================================================================== */

/** POST — every action. The body is a form field `payload` holding JSON. */
function doPost(e) {
  var payload;
  try {
    payload = JSON.parse((e && e.parameter && e.parameter.payload) || '{}');
  } catch (badJson) {
    return json({ success: false, error: 'Bad request' });
  }
  try {
    var err = checkKey(payload.key);
    if (err) return lockedResponse(err);
    delete payload.key;

    var action = payload.action || 'add';
    if (action === 'load')        return json(loadAll());
    if (action === 'scan')        return json(scanReceipt(payload.imageBase64, payload.mimeType));
    if (action === 'add')         return json(addExpense(payload));
    if (action === 'update')      return json(updateExpenses(payload.ids || [payload.id], payload.changes || {}));
    if (action === 'delete')      return json(deleteExpense(payload.id));
    if (action === 'saveBudgets') return json(saveBudgets(payload.budgets || {}));
    if (action === 'shareLink')   return json({ success: true, share: getShareKey(!!payload.reset) });
    if (action === 'migrate')     return json(migrateSheet());
    // Automatic bank import (Plaid)
    if (action === 'bankStatus')    return json(bankStatus());
    if (action === 'bankLinkToken') return json(bankLinkToken(payload.itemId));
    if (action === 'bankExchange')  return json(bankExchange(payload.publicToken, payload.institution));
    if (action === 'bankSync')      return json(syncAllBanks());
    if (action === 'bankRemove')    return json(bankRemove(payload.itemId));
    return json({ success: false, error: 'Unknown action: ' + action });
  } catch (err2) {
    return json({ success: false, error: String(err2) });
  }
}

/** Store the receipt in Drive and return its link. A Drive failure must not
    lose the expense itself, so it comes back as a marked string instead. */
function saveReceiptFile(p) {
  try {
    var folder = DriveApp.getFolderById(FOLDER_ID);
    var stamp  = Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd_HHmmss');
    var safe   = String(p.place || 'receipt').replace(/[^\w\s-]/g, '').trim().slice(0, 40);
    var label  = (p.receiptType === 'Bank Statement') ? 'statement' : 'receipt';
    var isPdf  = p.imageMime === 'application/pdf';
    var blob   = Utilities.newBlob(
      Utilities.base64Decode(p.imageBase64),
      isPdf ? 'application/pdf' : 'image/jpeg',
      stamp + '_' + label + '_' + safe + (isPdf ? '.pdf' : '.jpg')
    );
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (imgErr) {
    return 'UPLOAD_FAILED: ' + String(imgErr);
  }
}

/**
 * One expense, with an optional base64 receipt.
 *
 * The browser picks the ID. If a save times out and the browser retries,
 * the second attempt finds the ID already in the sheet and returns that row
 * instead of adding a duplicate (or uploading the receipt twice).
 */
function addExpense(p) {
  var id = cleanId(p.id) || newId();
  var sh = getSheet();

  var existing = findRow(sh, id);
  if (existing > 0) {
    return { success: true, duplicate: true, id: id,
             expense: toObject(sh.getRange(existing, 1, 1, HEADERS.length).getValues()[0], existing - 2) };
  }

  // Drive upload happens outside the lock — it's the slow part
  var receiptUrl = p.receiptUrl || '';
  if (p.imageBase64) receiptUrl = saveReceiptFile(p);

  return withLock(function () {
    var again = findRow(sh, id);
    if (again > 0) {
      return { success: true, duplicate: true, id: id,
               expense: toObject(sh.getRange(again, 1, 1, HEADERS.length).getValues()[0], again - 2) };
    }
    // Receipt for a charge the bank already imported: attach it there. The
    // receipt shows what was bought, so its 529 answer replaces the guess.
    if (p.receiptType !== 'Bank Transaction') {
      var bankRow = matchBankRow(sh, p.place, Number(p.amount) || 0, String(p.date || ''));
      if (bankRow > 0) {
        if (receiptUrl) sh.getRange(bankRow, 8).setValue(receiptUrl);
        var fields = { place: p.place, category: p.category, is529: p.is529, notes: p.notes,
                       needsReimb: p.needsReimb };
        for (var f in fields) {
          if (fields[f] == null || fields[f] === '') continue;
          sh.getRange(bankRow, EDITABLE[f] + 1).setValue(cleanValue(f, fields[f]));
        }
        sh.getRange(bankRow, COL.sortedBy).setValue('Receipt');
        sh.getRange(bankRow, COL.review).setValue('');
        var merged = toObject(sh.getRange(bankRow, 1, 1, HEADERS.length).getValues()[0], bankRow - 2);
        return { success: true, merged: true, id: merged.id, receiptUrl: merged.receiptUrl, expense: merged };
      }
    }
    p.id = id;
    p.receiptUrl = receiptUrl;
    if (!p.sortedBy) p.sortedBy = 'Receipt';
    var row = toRow(p);
    sh.appendRow(row);
    return { success: true, id: id, receiptUrl: receiptUrl, expense: toObject(row, 0) };
  });
}

// Fields the app may change after saving, and their 0-based column
var EDITABLE = { place: 1, category: 2, date: 3, amount: 4, semester: 5, receiptType: 6,
                 needsReimb: 8, reimbursed: 9, notes: 10, is529: 11 };

function cleanValue(field, v) {
  if (field === 'amount') {
    // Negative is allowed: bank refunds come in as negative amounts, and
    // a refunded 529 purchase has to reduce the 529 total.
    var n = Math.round((Number(v) || 0) * 100) / 100;
    if (!isFinite(n)) throw new Error('Amount must be a number');
    return n;
  }
  if (field === 'date') {
    v = String(v || '').trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(v)) throw new Error('Date must be YYYY-MM-DD');
    return v;
  }
  if (field === 'needsReimb' || field === 'reimbursed') {
    if (YES_NO.indexOf(v) === -1) throw new Error(field + ' must be Yes or No');
    return v;
  }
  if (field === 'is529') {
    if (YES_NO_MAYBE.indexOf(v) === -1) throw new Error('is529 must be Yes, No or Not Sure');
    return v;
  }
  if (field === 'receiptType') {
    if (RECEIPT_TYPES.indexOf(v) === -1) throw new Error('Unknown receipt type');
    return v;
  }
  return String(v == null ? '' : v).trim().slice(0, 500);
}

/** Apply the same changes to one or more expenses, by ID. */
function updateExpenses(ids, changes) {
  ids = (ids || []).map(cleanId).filter(String);
  if (!ids.length) return { success: false, error: 'No expense id given' };

  // Validate everything before touching the sheet
  var clean = {};
  for (var field in changes) {
    if (!changes.hasOwnProperty(field) || !EDITABLE.hasOwnProperty(field)) continue;
    clean[field] = cleanValue(field, changes[field]);
  }
  if (!Object.keys(clean).length) return { success: false, error: 'Nothing to change' };

  // Choosing 529 or a category yourself is the strongest signal there is:
  // the row is marked "You", leaves the review list, and teaches the sorter.
  var taught = clean.hasOwnProperty('is529') || clean.hasOwnProperty('category');

  return withLock(function () {
    var sh = getSheet();
    var updated = [], missing = [], merchants = {};
    for (var i = 0; i < ids.length; i++) {
      var row = findRow(sh, ids[i]);
      if (row < 0) { missing.push(ids[i]); continue; }
      for (var f in clean) sh.getRange(row, EDITABLE[f] + 1).setValue(clean[f]);
      if (taught) {
        sh.getRange(row, COL.sortedBy).setValue('You');
        sh.getRange(row, COL.review).setValue('');
      }
      var obj = toObject(sh.getRange(row, 1, 1, HEADERS.length).getValues()[0], row - 2);
      if (taught) merchants[merchantKey(obj.place)] = true;
      updated.push(obj);
    }
    // Other charges from the same store still waiting on you get re-sorted now
    var resorted = taught ? resortReviewRows(sh, merchants) : [];
    return { success: missing.length === 0, updated: updated.concat(resorted), missing: missing,
             resorted: resorted.length,
             error: missing.length ? missing.length + ' expense(s) not found — refresh and try again' : undefined };
  });
}

/** Moves the row to a "Deleted" tab rather than destroying it. */
function deleteExpense(id) {
  id = cleanId(id);
  if (!id) return { success: false, error: 'No expense id given' };
  return withLock(function () {
    var sh = getSheet();
    var row = findRow(sh, id);
    if (row < 0) return { success: true, missing: true };
    var values = sh.getRange(row, 1, 1, HEADERS.length).getValues()[0];

    var ss = book();
    var bin = ss.getSheetByName(DELETED_TAB);
    if (!bin) {
      bin = ss.insertSheet(DELETED_TAB);
      bin.appendRow(HEADERS.concat(['Deleted At']));
      bin.getRange(1, 1, 1, HEADERS.length + 1)
         .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#8A2E26');
      bin.setFrozenRows(1);
    }
    bin.appendRow(values.concat([new Date()]));
    sh.deleteRow(row);
    return { success: true };
  });
}

/* ===================================================================
   BUDGETS — a "Budgets" tab: Category | Monthly Budget
   Categories are the app's dashboard categories (Groceries and Eating Out
   are separate there), plus "Total" for an overall monthly limit.
   =================================================================== */

function budgetSheet(create) {
  var ss = book();
  var sh = ss.getSheetByName(BUDGET_TAB);
  if (!sh && create) {
    sh = ss.insertSheet(BUDGET_TAB);
    sh.appendRow(['Category', 'Monthly Budget']);
    sh.getRange(1, 1, 1, 2).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#3A6644');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 200);
    sh.setColumnWidth(2, 140);
  }
  return sh;
}

function readBudgets() {
  var sh = budgetSheet(false);
  if (!sh || sh.getLastRow() < 2) return {};
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  var out = {};
  for (var i = 0; i < values.length; i++) {
    var cat = String(values[i][0] || '').trim();
    var amt = Number(values[i][1]) || 0;
    if (cat && amt > 0) out[cat] = Math.round(amt * 100) / 100;
  }
  return out;
}

function saveBudgets(budgets) {
  var rows = [];
  for (var cat in budgets) {
    if (!budgets.hasOwnProperty(cat)) continue;
    var amt = Math.round((Number(budgets[cat]) || 0) * 100) / 100;
    if (amt > 0) rows.push([String(cat).trim().slice(0, 60), amt]);
  }
  return withLock(function () {
    var sh = budgetSheet(true);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
    if (rows.length) {
      sh.getRange(2, 1, rows.length, 2).setValues(rows);
      sh.getRange(2, 2, rows.length, 1).setNumberFormat('$#,##0.00');
    }
    return { success: true, budgets: readBudgets() };
  });
}

/* ===================================================================
   REIMBURSEMENT SHARE LINK

   A separate long random key that unlocks ONE read-only view: expenses
   marked "needs reimbursement" that aren't reimbursed yet. It can't load
   anything else or change anything. Resetting it (or deleting SHARE_KEY
   in Script Properties) kills every link handed out before.
   =================================================================== */

function getShareKey(reset) {
  var props = PropertiesService.getScriptProperties();
  var k = props.getProperty('SHARE_KEY');
  if (!k || reset) {
    k = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
    props.setProperty('SHARE_KEY', k);
  }
  return k;
}

function shareView(share) {
  var k = PropertiesService.getScriptProperties().getProperty('SHARE_KEY');
  if (!k || !share || String(share) !== k) {
    return { success: false, error: 'This link has expired. Ask Dylan for a new one.' };
  }
  var items = [];
  var all = readExpenses(getSheet());
  for (var i = 0; i < all.length; i++) {
    var e = all[i];
    if (e.needsReimb !== 'Yes' || e.reimbursed === 'Yes') continue;
    items.push({
      date: e.date, place: e.place, amount: e.amount, category: e.category, notes: e.notes,
      receiptUrl: /^https:\/\//.test(e.receiptUrl) ? e.receiptUrl : ''
    });
  }
  return { success: true, items: items, generated: new Date().toISOString() };
}

/* ===================================================================
   529 SORTER — for charges that arrive from the bank

   Deterministic first, AI last, and never a guess on a mixed store:
     1. Payment apps (Venmo, Zelle…) can't be sorted from the charge alone.
     2. Your history: how this store has been labeled before. A store you
        corrected yourself counts triple. It takes 80% agreement, so a store
        you've labeled both ways (Walmart groceries vs. Walmart clothes) is
        never guessed from history.
     3. Words that always decide it: student loans, certification exams,
        tuition/bursar, campus dining, parking.
     4. Plaid's own category (groceries vs. restaurants, rent, bookstores…)
        when Plaid is confident — with the store-name food rules on top.
     5. The same store-name food rules the app uses for receipts.
     6. Gemini, given only the store, Plaid's label and the amount.
   Whatever is still unclear is "Not Sure" and goes to Needs a look.
   Every row records which step decided it, in "Sorted By".
   =================================================================== */

// Plaid detailed category -> [app category, 529]
var PFC_MAP = {
  FOOD_AND_DRINK_GROCERIES:                       ['Food & Dining', 'Yes'],   // room & board
  FOOD_AND_DRINK_RESTAURANT:                      ['Food & Dining', 'No'],
  FOOD_AND_DRINK_FAST_FOOD:                       ['Food & Dining', 'No'],
  FOOD_AND_DRINK_COFFEE:                          ['Food & Dining', 'No'],
  FOOD_AND_DRINK_BEER_WINE_AND_LIQUOR:            ['Food & Dining', 'No'],
  FOOD_AND_DRINK_VENDING_MACHINES:                ['Food & Dining', 'No'],
  FOOD_AND_DRINK_OTHER_FOOD_AND_DRINK:            ['Food & Dining', 'Not Sure'],
  RENT_AND_UTILITIES_RENT:                        ['Housing', 'Yes'],
  RENT_AND_UTILITIES_GAS_AND_ELECTRICITY:         ['Housing', 'Yes'],
  RENT_AND_UTILITIES_WATER:                       ['Housing', 'Yes'],
  RENT_AND_UTILITIES_SEWAGE_AND_WASTE_MANAGEMENT: ['Housing', 'Yes'],
  RENT_AND_UTILITIES_OTHER_UTILITIES:             ['Housing', 'Yes'],
  RENT_AND_UTILITIES_INTERNET_AND_CABLE:          ['Technology', 'Yes'],      // internet access qualifies
  RENT_AND_UTILITIES_TELEPHONE:                   ['Technology', 'No'],       // phone plans generally don't
  GENERAL_MERCHANDISE_BOOKSTORES_AND_NEWSSTANDS:  ['Books & Supplies', 'Yes'],
  GENERAL_MERCHANDISE_OFFICE_SUPPLIES:            ['School Supplies', 'Yes'],
  GENERAL_MERCHANDISE_ELECTRONICS:                ['Technology', 'Yes'],
  GENERAL_MERCHANDISE_CLOTHING_AND_ACCESSORIES:   ['Clothing', 'No'],
  GENERAL_MERCHANDISE_SPORTING_GOODS:             ['Entertainment', 'No'],
  GENERAL_MERCHANDISE_GIFTS_AND_NOVELTIES:        ['Other', 'No'],
  GENERAL_MERCHANDISE_PET_SUPPLIES:               ['Other', 'No'],
  GENERAL_MERCHANDISE_TOBACCO_AND_VAPE:           ['Other', 'No'],
  // Stores that sell everything: the charge alone can't say what you bought
  GENERAL_MERCHANDISE_SUPERSTORES:                ['Other', 'Not Sure'],
  GENERAL_MERCHANDISE_ONLINE_MARKETPLACES:        ['Other', 'Not Sure'],
  GENERAL_MERCHANDISE_DEPARTMENT_STORES:          ['Other', 'Not Sure'],
  GENERAL_MERCHANDISE_DISCOUNT_STORES:            ['Other', 'Not Sure'],
  GENERAL_MERCHANDISE_CONVENIENCE_STORES:         ['Food & Dining', 'Not Sure'],
  GENERAL_MERCHANDISE_OTHER_GENERAL_MERCHANDISE:  ['Other', 'Not Sure'],
  GENERAL_SERVICES_EDUCATION:                     ['Tuition & Fees', 'Yes'],
  HOME_IMPROVEMENT_FURNITURE:                     ['Housing', 'Not Sure'],
  LOAN_PAYMENTS_STUDENT_LOAN_PAYMENT:             ['Other', 'Yes'],           // $10,000 lifetime cap
  MEDICAL_PHARMACIES_AND_SUPPLEMENTS:             ['Healthcare', 'No']
};
// Fallback by Plaid primary category
var PFC_PRIMARY_MAP = {
  ENTERTAINMENT:            ['Entertainment', 'No'],
  PERSONAL_CARE:            ['Personal Care', 'No'],
  MEDICAL:                  ['Healthcare', 'No'],
  TRANSPORTATION:           ['Transportation', 'No'],
  TRAVEL:                   ['Transportation', 'No'],
  BANK_FEES:                ['Other', 'No'],
  GOVERNMENT_AND_NON_PROFIT:['Other', 'No'],
  GENERAL_SERVICES:         ['Other', 'No'],
  HOME_IMPROVEMENT:         ['Other', 'No'],
  FOOD_AND_DRINK:           ['Food & Dining', 'Not Sure'],
  RENT_AND_UTILITIES:       ['Housing', 'Yes']
};

// Same lists as index.html's receipt rules — keep the two in step.
var CAMPUS_FOOD = ['meal plan','dining hall','campus dining','dining dollars','knights cash',
  'meal swipe','dining services','ucf dining','student union food'];
var RESTAURANT_SIGNALS = ['pizza','grill','cafe','café','kitchen','burger','sandwich','taco',
  'sushi','diner','bistro','wings','bbq','barbecue','steakhouse','noodle','ramen','pho',
  'buffet','bakery','coffee','espresso','juice','smoothie','creamery','ice cream','donut',
  'bagel','pub','tavern','cantina','grille','eatery','restaurant','chicken','wok','curry'];
var RESTAURANT_CHAINS = ['chipotle','panda express','chick-fil-a','chick fil a','chik fil a',
  'chickfila','zaxby','mcdonald','burger king','wendy','taco bell','subway','panera','starbucks',
  'dunkin','popeye','raising cane','five guys','jimmy john','jersey mike','moe\'s','qdoba',
  'halal guys','domino','papa john','kfc','arby','sonic drive','whataburger','in-n-out',
  'shake shack','olive garden','chili\'s','applebee','ihop','denny','waffle house','huey magoo',
  'dippin','insomnia cookie','crumbl','jamba','tropical smoothie','acai','açaí','poke',
  'firehouse','potbelly','sweetgreen','cava','nando','wingstop','checkers','culver',
  'steak n shake','steak \'n shake','little caesar','marco\'s','hungry howie','bojangles',
  'la granja','pollo tropical','miller\'s ale','buffalo wild','hooters','twin peaks',
  'doordash','uber eats','ubereats','grubhub'];
var GROCERY_CHAINS = ['walmart','publix','costco','kroger','aldi','trader joe','whole foods',
  'winn dixie','winn-dixie','safeway','sprouts','lidl','food lion','h-e-b','heb','ingles',
  'piggly wiggly','price chopper','shoprite','stop & shop','meijer','harris teeter','wegmans',
  'sam\'s club','bj\'s wholesale','fresh market','fresh thyme','earth fare','sedano','presidente',
  'bravo supermarket','patel brothers','h mart','99 ranch','ifresh','natural grocers','winco',
  'hy-vee','albertsons','vons','ralphs','fred meyer','king soopers','jewel-osco','hannaford',
  'food 4 less','save a lot','giant eagle','giant food','stater bros','raley','wawa food',
  'target','walgreens','cvs','instacart'];
var GROCERY_TOKENS = ['grocery','groceries','grocer','supermarket','supermercado',
  'market','mart','bodega','produce','butcher'];

function foodIs529(text) {
  var t = ' ' + String(text || '').toLowerCase() + ' ';
  function wordish(w) {
    return new RegExp('(^|[^a-z])' + w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(s|es)?([^a-z]|$)').test(t);
  }
  function has(list) { for (var i = 0; i < list.length; i++) if (t.indexOf(list[i]) !== -1) return true; return false; }
  function hasWord(list) { for (var i = 0; i < list.length; i++) if (wordish(list[i])) return true; return false; }
  if (has(CAMPUS_FOOD)) return 'Yes';
  if (has(RESTAURANT_CHAINS)) return 'No';
  if (hasWord(RESTAURANT_SIGNALS)) return 'No';
  if (has(GROCERY_CHAINS)) return 'Yes';
  if (hasWord(GROCERY_TOKENS)) return 'Yes';
  return 'Not Sure';
}

var P2P_RE = /venmo|zelle|cash ?app|square cash|paypal|apple cash/i;

/** Store name for a bank charge: Plaid's cleaned merchant, else a tidied raw line. */
function bankPlace(t) {
  if (t.merchant_name) return String(t.merchant_name).slice(0, 60);
  var cp = (t.counterparties || [])[0];
  if (cp && cp.name) return String(cp.name).slice(0, 60);
  var s = String(t.name || '')
    .replace(/^(sq|tst|pp|sp|py)\s*\*\s*/i, '')
    .replace(/\b(pos|debit|purchase|card|recurring|checkcard|ach)\b/gi, ' ')
    .replace(/#?\d{3,}/g, ' ')
    .replace(/\s{2,}/g, ' ').trim();
  s = s.toLowerCase().replace(/\b[a-z]/g, function (c) { return c.toUpperCase(); });
  return s.slice(0, 60) || 'Unknown';
}

/** "Publix Super Markets #1234" and "PUBLIX" both become "publix". */
function merchantKey(name) {
  var s = String(name || '').toLowerCase()
    .replace(/\.com\b/g, '')
    .replace(/['’`.\-]/g, '')
    .replace(/[^a-z0-9 ]+/g, ' ')
    .replace(/\b(the|inc|llc|co|store|stores|supercenter)\b/g, ' ');
  var t = s.split(/\s+/).filter(function (w) { return w && !/^\d+$/.test(w); });
  if (!t.length) return '';
  return t[0].length >= 3 ? t[0] : t.slice(0, 2).join(' ');
}

function isP2P(t) {
  if (t.merchant_name) return P2P_RE.test(t.merchant_name);
  var cp = (t.counterparties || [])[0];
  return P2P_RE.test(String(t.name || '')) || !!(cp && P2P_RE.test(String(cp.name || '')));
}

/** Money that isn't spending: income, transfers between your own accounts,
    card payments (the purchases themselves are already counted), ATM cash. */
function isNotSpending(t) {
  var pfc = t.personal_finance_category || {};
  var pr = pfc.primary || '', det = pfc.detailed || '';
  if (isP2P(t)) return t.amount < 0;                   // money sent to you
  if (pr === 'INCOME' || pr === 'TRANSFER_IN') return true;
  if (pr === 'TRANSFER_OUT') return true;
  if (pr === 'LOAN_PAYMENTS' && det !== 'LOAN_PAYMENTS_STUDENT_LOAN_PAYMENT') return true;
  return false;
}

/** How you've labeled each store so far: key -> { yes, no, cats }. */
function buildHistory(rows) {
  var h = {};
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!r[1]) continue;
    var is = String(r[11]);
    if (is !== 'Yes' && is !== 'No') continue;
    var by = String(r[16] || '');
    var fromBank = String(r[6]) === 'Bank Transaction';
    var w = by === 'You' ? 3 : (!fromBank && (by === '' || by === 'Receipt') ? 1 : 0);
    if (!w) continue;                                  // auto-sorted rows never vote
    var key = merchantKey(r[1]);
    if (!key) continue;
    var e = h[key] || (h[key] = { yes: 0, no: 0, cats: {} });
    if (is === 'Yes') e.yes += w; else e.no += w;
    var cat = String(r[2] || 'Other');
    e.cats[cat] = (e.cats[cat] || 0) + w;
  }
  return h;
}

function historyVote(h, key) {
  var e = h[key];
  if (!e) return null;
  var total = e.yes + e.no;
  if (total < 2) return null;
  if (Math.max(e.yes, e.no) / total < 0.8) return null;
  var best = 'Other', bestW = -1;
  for (var c in e.cats) if (e.cats[c] > bestW) { best = c; bestW = e.cats[c]; }
  return { is529: e.yes >= e.no ? 'Yes' : 'No', category: best };
}

function keywordVerdict(text) {
  var t = String(text || '').toLowerCase();
  if (/student loan|loan payment|sallie mae|nelnet|mohela|navient|aidvantage/.test(t)) return ['Other', 'Yes'];
  if (/apprentice/.test(t)) return ['Other', 'Yes'];
  if (/certification|licensure|license exam|pearson vue|prometric|comptia|ncees/.test(t)) return ['Tuition & Fees', 'Yes'];
  if (/tuition|bursar|registrar|student account/.test(t)) return ['Tuition & Fees', 'Yes'];
  if (/parking|park mobile|parkmobile/.test(t)) return ['Transportation', 'No'];
  for (var i = 0; i < CAMPUS_FOOD.length; i++) if (t.indexOf(CAMPUS_FOOD[i]) !== -1) return ['Food & Dining', 'Yes'];
  return null;
}

/**
 * Decide category and 529 for one bank charge.
 * @param ai  budget object { left: n } so one sync can't make hundreds of AI calls
 */
function classifyCharge(t, history, ai) {
  var place = bankPlace(t);
  var pfc = t.personal_finance_category || {};
  var det = pfc.detailed || '', pr = pfc.primary || '';
  var confident = ['LOW', 'UNKNOWN'].indexOf(String(pfc.confidence_level || '').toUpperCase()) === -1;
  var text = place + ' ' + String(t.name || '');

  if (isP2P(t)) return { category: 'Other', is529: 'Not Sure', sortedBy: 'Needs you',
                         notes: 'Payment app — what was it for?' };

  var h = historyVote(history, merchantKey(place));
  if (h) return { category: h.category, is529: h.is529, sortedBy: 'Your history' };

  var kw = keywordVerdict(text);
  if (kw) return { category: kw[0], is529: kw[1], sortedBy: 'Store name' };

  var m = PFC_MAP[det] || PFC_PRIMARY_MAP[pr];
  if (m && confident) {
    if (m[0] === 'Food & Dining') {
      var food = foodIs529(text);
      if (food !== 'Not Sure') return { category: 'Food & Dining', is529: food, sortedBy: 'Store name' };
    }
    if (m[1] !== 'Not Sure') return { category: m[0], is529: m[1], sortedBy: 'Bank category' };
  }

  // Store-name food rules, but not for "sells everything" stores: on a
  // receipt "Walmart" usually means groceries, on a card charge it could be
  // a TV. Those wait for your history or for you.
  var food2 = foodIs529(text);
  if (food2 !== 'Not Sure' && (!m || m[0] === 'Food & Dining')) {
    return { category: 'Food & Dining', is529: food2, sortedBy: 'Store name' };
  }

  if (ai && ai.left > 0) {
    ai.left--;
    var g = aiClassify(place, det, t.amount);
    if (g && g.is529 !== 'Not Sure') return { category: g.category, is529: g.is529, sortedBy: 'AI' };
    if (g) return { category: g.category, is529: 'Not Sure', sortedBy: 'AI' };
  }
  return { category: m ? m[0] : 'Other', is529: 'Not Sure', sortedBy: m ? 'Bank category' : '' };
}

var CLASSIFY_PROMPT =
'You sort one bank card charge for a US college student\'s 529 plan records. Return ONLY raw JSON, no markdown:\n' +
'{"category":"Food & Dining","is529":"Yes","confident":true}\n' +
'category is exactly one of: Food & Dining, Tuition & Fees, Books & Supplies, Housing, Transportation, Technology, Healthcare, Entertainment, Clothing, Personal Care, School Supplies, Other.\n' +
'is529 follows IRS Qualified Higher Education Expense rules: "Yes" for tuition and required fees, required books/supplies/equipment, computers/software/internet, room and board (rent, utilities, groceries, meal plans). "No" for restaurants, travel, transportation, parking, medical, insurance, clothing, personal care, entertainment, phone plans.\n' +
'If the store sells many kinds of things (Walmart, Target, Amazon, Costco, a department store) and the name does not say what was bought, use "Not Sure". Set confident to false whenever you are guessing.';

function aiClassify(place, det, amount) {
  var prompt = CLASSIFY_PROMPT + '\n\nStore: ' + place + '\nBank label: ' + (det || 'none') +
               '\nAmount: $' + Math.abs(Number(amount) || 0).toFixed(2);
  var models = [GEMINI_PRIMARY, GEMINI_FALLBACK];
  for (var i = 0; i < models.length; i++) {
    try {
      var resp = UrlFetchApp.fetch(
        'https://generativelanguage.googleapis.com/v1beta/models/' + models[i] +
        ':generateContent?key=' + encodeURIComponent(geminiKey()),
        { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
          payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }],
                                    generationConfig: { temperature: 0, maxOutputTokens: 80 } }) });
      if (resp.getResponseCode() !== 200) continue;
      var parts = ((((JSON.parse(resp.getContentText()).candidates || [])[0] || {}).content || {}).parts) || [];
      var out = '';
      for (var p = 0; p < parts.length; p++) if (!parts[p].thought && parts[p].text) out += parts[p].text;
      var m = out.match(/\{[\s\S]*?\}/);
      if (!m) continue;
      var d = JSON.parse(m[0]);
      var cat = CATEGORIES.indexOf(d.category) !== -1 ? d.category : 'Other';
      var is = YES_NO_MAYBE.indexOf(d.is529) !== -1 ? d.is529 : 'Not Sure';
      if (d.confident === false) is = 'Not Sure';
      return { category: cat, is529: is };
    } catch (e) { /* try the next model */ }
  }
  return null;
}

/** After you answer one charge, re-sort the others from that store still waiting. */
function resortReviewRows(sh, merchants) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  var history = buildHistory(rows);
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (String(r[17]) !== 'Yes' || String(r[16]) === 'You') continue;
    var key = merchantKey(r[1]);
    if (!merchants[key]) continue;
    var v = historyVote(history, key);
    if (!v) continue;
    var rowNum = i + 2;
    sh.getRange(rowNum, 3).setValue(v.category);
    sh.getRange(rowNum, 12).setValue(v.is529);
    if (String(r[9]) !== 'Yes') sh.getRange(rowNum, 9).setValue(v.is529 === 'Yes' ? 'Yes' : 'No');
    sh.getRange(rowNum, COL.sortedBy).setValue('Your history');
    sh.getRange(rowNum, COL.review).setValue('');
    out.push(toObject(sh.getRange(rowNum, 1, 1, HEADERS.length).getValues()[0], i));
  }
  return out;
}

function semesterOf(dateStr) {
  var m = String(dateStr || '').match(/^(\d{4})-(\d{2})/);
  if (!m) return '';
  var mo = Number(m[2]);
  return (mo <= 5 ? 'Spring ' : mo <= 8 ? 'Summer ' : 'Fall ') + m[1];
}

/* ===================================================================
   AUTOMATIC BANK IMPORT (Plaid)

   ONE-TIME SETUP — Script Properties:
     PLAID_CLIENT_ID, PLAID_SECRET   from dashboard.plaid.com > Developers > Keys
     PLAID_ENV                       "sandbox" to test with fake banks; leave
                                     unset (production) for your real accounts
   Then run setupBankSync() once from the editor (Run button) so Google asks
   for permission to schedule the sync. After that it runs every 6 hours.

   Each connected bank is kept in a Script Property named BANKITEM_<id>
   (its Plaid access token and sync position). Nothing bank-related is ever
   sent to the browser except names, last-4 digits and sync status.
   =================================================================== */

function plaidEnv() {
  var e = String(PropertiesService.getScriptProperties().getProperty('PLAID_ENV') || '').toLowerCase();
  return e === 'sandbox' ? 'sandbox' : 'production';
}

function plaidCall(path, body) {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('PLAID_CLIENT_ID'), secret = props.getProperty('PLAID_SECRET');
  if (!id || !secret) throw new Error('Plaid keys are not set. Add PLAID_CLIENT_ID and PLAID_SECRET in Script Properties.');
  body.client_id = id;
  body.secret = secret;
  var resp = UrlFetchApp.fetch('https://' + plaidEnv() + '.plaid.com' + path, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true
  });
  var data = {};
  try { data = JSON.parse(resp.getContentText() || '{}'); } catch (e) { /* non-JSON error page */ }
  if (resp.getResponseCode() !== 200) {
    var err = new Error(data.error_message || data.display_message || ('Plaid returned HTTP ' + resp.getResponseCode()));
    err.plaidCode = data.error_code || '';
    throw err;
  }
  return data;
}

var ITEM_PREFIX = 'BANKITEM_';
function getItem(id) {
  var v = PropertiesService.getScriptProperties().getProperty(ITEM_PREFIX + id);
  return v ? JSON.parse(v) : null;
}
function saveItem(it) {
  PropertiesService.getScriptProperties().setProperty(ITEM_PREFIX + it.itemId, JSON.stringify(it));
}
function listItems() {
  var all = PropertiesService.getScriptProperties().getProperties(), out = [];
  for (var k in all) if (k.indexOf(ITEM_PREFIX) === 0) { try { out.push(JSON.parse(all[k])); } catch (e) {} }
  out.sort(function (a, b) { return String(a.institution).localeCompare(String(b.institution)); });
  return out;
}
function publicItem(it) {
  var accts = [];
  for (var a in (it.accounts || {})) accts.push(it.accounts[a]);
  return { itemId: it.itemId, institution: it.institution, accounts: accts,
           lastSync: it.lastSync || '', error: it.error || '',
           needsRelink: /ITEM_LOGIN_REQUIRED|PENDING_EXPIRATION|INVALID_CREDENTIALS|ACCESS_NOT_GRANTED/.test(it.error || '') };
}

function shortInstitution(name) {
  var n = String(name || 'Bank');
  if (/american express/i.test(n)) return 'Amex';
  if (/^usaa/i.test(n)) return 'USAA';
  if (/chase/i.test(n)) return 'Chase';
  return n.replace(/\b(bank|financial|credit union|n\.?a\.?)\b/gi, '').replace(/\s{2,}/g, ' ').trim().slice(0, 20) || n.slice(0, 20);
}

function accountLabels(accessToken, institution) {
  var res = plaidCall('/accounts/get', { access_token: accessToken });
  var out = {}, inst = shortInstitution(institution);
  (res.accounts || []).forEach(function (a) {
    out[a.account_id] = inst + (a.mask ? ' •' + a.mask : ' ' + (a.name || ''));
  });
  return out;
}

function bankStatus() {
  var props = PropertiesService.getScriptProperties();
  var scheduled = null;
  try {
    scheduled = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'syncAllBanks'; });
  } catch (e) { scheduled = null; }                  // permission not granted yet
  return {
    success: true,
    configured: !!(props.getProperty('PLAID_CLIENT_ID') && props.getProperty('PLAID_SECRET')),
    env: plaidEnv(),
    scheduled: scheduled,
    lastRun: props.getProperty('BANK_LAST_RUN') || '',
    startDate: props.getProperty('BANK_START') || '',
    banks: listItems().map(publicItem)
  };
}

/** A short-lived token that opens Plaid's bank sign-in window. With an
    itemId it opens "update mode" to fix an expired bank login. */
function bankLinkToken(itemId) {
  var body = { client_name: 'Dylan\'s Expenses', language: 'en', country_codes: ['US'],
               user: { client_user_id: 'dylan' } };
  if (itemId) {
    var it = getItem(itemId);
    if (!it) return { success: false, error: 'That bank isn\'t connected anymore' };
    body.access_token = it.accessToken;
  } else {
    body.products = ['transactions'];
    body.transactions = { days_requested: 730 };
  }
  return { success: true, linkToken: plaidCall('/link/token/create', body).link_token, env: plaidEnv() };
}

function bankExchange(publicToken, institution) {
  if (!publicToken) return { success: false, error: 'Missing token from Plaid' };
  var ex = plaidCall('/item/public_token/exchange', { public_token: String(publicToken) });
  var it = getItem(ex.item_id) || { itemId: ex.item_id, cursor: '' };
  it.accessToken = ex.access_token;
  it.institution = String(institution || it.institution || 'Bank').slice(0, 40);
  it.accounts = accountLabels(ex.access_token, it.institution);
  it.error = '';
  saveItem(it);

  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('BANK_START')) {
    // Import from the start of this month. Earlier spending is already in the
    // Sheet from receipts, and this month's receipts get matched, not doubled.
    var d = new Date();
    props.setProperty('BANK_START', Utilities.formatDate(new Date(d.getFullYear(), d.getMonth(), 1), 'America/New_York', 'yyyy-MM-dd'));
  }

  var warning = '';
  try { ensureSyncTriggers(true); }
  catch (e) { warning = 'Automatic syncing needs one more step: run setupBankSync once in the Apps Script editor.'; }

  var sync = syncItem(it);
  return { success: true, bank: publicItem(getItem(it.itemId)), sync: sync, warning: warning };
}

function bankRemove(itemId) {
  var it = getItem(itemId);
  if (!it) return bankStatus();
  try { plaidCall('/item/remove', { access_token: it.accessToken }); } catch (e) { /* already gone at Plaid */ }
  PropertiesService.getScriptProperties().deleteProperty(ITEM_PREFIX + itemId);
  return bankStatus();                                // rows already imported stay in the Sheet
}

/** Pull everything new for one bank since its last sync. */
function syncItem(it) {
  var added = [], modified = [], removed = [];
  var start = it.cursor || '', cursor = start, hasMore = true, guard = 0, restarts = 0;
  try {
    while (hasMore && guard++ < 60) {
      var body = { access_token: it.accessToken, count: 500 };
      if (cursor) body.cursor = cursor;
      var r;
      try {
        r = plaidCall('/transactions/sync', body);
      } catch (e) {
        if (e.plaidCode === 'TRANSACTIONS_SYNC_MUTATION_DURING_PAGINATION' && restarts++ < 3) {
          cursor = start; added = []; modified = []; removed = []; hasMore = true;
          continue;
        }
        throw e;
      }
      added = added.concat(r.added || []);
      modified = modified.concat(r.modified || []);
      removed = removed.concat(r.removed || []);
      hasMore = !!r.has_more;
      cursor = r.next_cursor;
    }
  } catch (err) {
    it.error = err.plaidCode || String(err.message || err);
    saveItem(it);
    return { error: it.error, added: 0, linked: 0, updated: 0, removed: 0, review: 0 };
  }
  var res = applyBankChanges(it, added, modified, removed);
  it.cursor = cursor;                                 // only after the rows are written
  it.lastSync = new Date().toISOString();
  it.error = '';
  saveItem(it);
  return res;
}

function syncAllBanks() {
  var items = listItems();
  var total = { success: true, banks: items.length, added: 0, linked: 0, updated: 0, removed: 0, review: 0, errors: [] };
  for (var i = 0; i < items.length; i++) {
    var r = syncItem(items[i]);
    if (r.error) total.errors.push(items[i].institution + ': ' + r.error);
    total.added += r.added; total.linked += r.linked; total.updated += r.updated;
    total.removed += r.removed; total.review += r.review;
  }
  PropertiesService.getScriptProperties().setProperty('BANK_LAST_RUN', new Date().toISOString());
  return total;
}

/** One-off run ten minutes after connecting a bank — Plaid needs a few
    minutes to gather history the first time. Deletes its own trigger. */
function syncBanksSoon() {
  try { syncAllBanks(); } finally { deleteTriggers('syncBanksSoon'); }
}

function deleteTriggers(handler) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === handler) ScriptApp.deleteTrigger(t);
  });
}

function ensureSyncTriggers(soon) {
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'syncAllBanks'; });
  if (!has) ScriptApp.newTrigger('syncAllBanks').timeBased().everyHours(6).create();
  if (soon) {
    deleteTriggers('syncBanksSoon');
    ScriptApp.newTrigger('syncBanksSoon').timeBased().after(10 * 60 * 1000).create();
  }
}

/** Run this once from the editor: it asks Google for permission to schedule
    the sync, then turns on the every-6-hours run. */
function setupBankSync() {
  ensureSyncTriggers(false);
  var s = bankStatus();
  Logger.log('Bank sync scheduled: ' + s.scheduled + ' · banks connected: ' + s.banks.length + ' · Plaid: ' + s.env);
  return s;
}

/** Write one sync's worth of changes into the Sheet. */
function applyBankChanges(it, added, modified, removed) {
  return withLock(function () {
    var sh = getSheet();
    var last = sh.getLastRow();
    var rows = last > 1 ? sh.getRange(2, 1, last - 1, HEADERS.length).getValues() : [];
    var startDate = PropertiesService.getScriptProperties().getProperty('BANK_START') || '0000-00-00';

    var byTxn = {}, candidates = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (r[13]) { byTxn[String(r[13])] = i; continue; }
      if (String(r[6]) === 'Bank Transaction' || !r[1]) continue;
      candidates.push({ i: i, amount: Number(r[4]) || 0, date: formatDate(r[3]), key: merchantKey(r[1]) });
    }
    var history = buildHistory(rows);
    var ai = { left: 40 };
    var res = { added: 0, linked: 0, updated: 0, removed: 0, review: 0 };
    var newRows = [];

    for (var a = 0; a < added.length; a++) {
      var t = added[a];
      if (t.pending || byTxn.hasOwnProperty(t.transaction_id)) continue;
      var date = String(t.authorized_date || t.date || '').slice(0, 10);
      if (!date || date < startDate) continue;
      if (isNotSpending(t)) continue;

      var place = bankPlace(t);
      var amount = Math.round(Number(t.amount) * 100) / 100;
      var account = (it.accounts || {})[t.account_id] || shortInstitution(it.institution);
      var det = (t.personal_finance_category || {}).detailed || '';

      // A receipt you already uploaded for this charge: attach the bank proof
      // to it instead of adding the purchase twice.
      var m = matchReceipt(candidates, place, amount, date);
      if (m) {
        var rowNum = m.i + 2;
        sh.getRange(rowNum, COL.txn).setValue(t.transaction_id);
        sh.getRange(rowNum, COL.account).setValue(account);
        sh.getRange(rowNum, COL.bankCat).setValue(det);
        candidates.splice(candidates.indexOf(m), 1);
        byTxn[t.transaction_id] = m.i;
        res.linked++;
        continue;
      }

      var c = classifyCharge(t, history, ai);
      var review = c.is529 === 'Not Sure' ? 'Yes' : '';
      newRows.push(toRow({
        id: 'b' + String(t.transaction_id).replace(/[^A-Za-z0-9_-]/g, '').slice(0, 38),
        timestamp: new Date().toISOString(), place: place, category: c.category, date: date,
        amount: amount, semester: semesterOf(date), receiptType: 'Bank Transaction', receiptUrl: '',
        needsReimb: c.is529 === 'Yes' ? 'Yes' : 'No', reimbursed: 'No',
        notes: c.notes || (amount < 0 ? 'Refund' : ''), is529: c.is529,
        bankTxnId: t.transaction_id, account: account, bankCategory: det,
        sortedBy: c.sortedBy, review: review
      }));
      byTxn[t.transaction_id] = -1;
      if (review) res.review++;
    }
    if (newRows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, newRows.length, HEADERS.length).setValues(newRows);
      res.added = newRows.length;
    }

    // Amount or date changed at the bank (e.g. a tip posted): follow it,
    // unless you've edited that row yourself.
    if (modified.length) {
      var idx = txnIndex(sh);
      for (var mo = 0; mo < modified.length; mo++) {
        var tm = modified[mo], rn = idx[tm.transaction_id];
        if (!rn) continue;
        if (String(sh.getRange(rn, COL.sortedBy).getValue()) === 'You') continue;
        if (String(sh.getRange(rn, 7).getValue()) !== 'Bank Transaction') continue;
        sh.getRange(rn, 5).setValue(Math.round(Number(tm.amount) * 100) / 100);
        var md = String(tm.authorized_date || tm.date || '').slice(0, 10);
        if (md) { sh.getRange(rn, 4).setValue(md); sh.getRange(rn, 6).setValue(semesterOf(md)); }
        res.updated++;
      }
    }

    // Charges the bank dropped (usually a reversed hold): move to Deleted
    if (removed.length) {
      var idx2 = txnIndex(sh), gone = [];
      for (var rm = 0; rm < removed.length; rm++) {
        var rn2 = idx2[removed[rm].transaction_id];
        if (rn2 && String(sh.getRange(rn2, 7).getValue()) === 'Bank Transaction') gone.push(rn2);
      }
      gone.sort(function (x, y) { return y - x; });
      if (gone.length) {
        var ss = book();
        var bin = ss.getSheetByName(DELETED_TAB) || ss.insertSheet(DELETED_TAB);
        if (bin.getLastRow() === 0) bin.appendRow(HEADERS.concat(['Deleted At']));
        for (var g = 0; g < gone.length; g++) {
          bin.appendRow(sh.getRange(gone[g], 1, 1, HEADERS.length).getValues()[0].concat([new Date()]));
          sh.deleteRow(gone[g]);
          res.removed++;
        }
      }
    }
    return res;
  });
}

/** A bank-imported charge with no receipt yet that this receipt is for. */
function matchBankRow(sh, place, amount, date) {
  var last = sh.getLastRow();
  if (last < 2 || !/^\d{4}-\d{2}-\d{2}$/.test(date)) return -1;
  var rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  var cands = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (String(r[6]) !== 'Bank Transaction' || r[7]) continue;
    cands.push({ i: i, amount: Number(r[4]) || 0, date: formatDate(r[3]), key: merchantKey(r[1]) });
  }
  // Same scoring as matchReceipt, seen from the receipt's side: the bank
  // amount may be higher than the receipt (tip), never lower.
  var key = merchantKey(place), best = null, bestScore = 0;
  var d0 = new Date(date + 'T12:00:00').getTime();
  for (var c = 0; c < cands.length; c++) {
    var k = cands[c];
    var days = Math.abs(new Date(k.date + 'T12:00:00').getTime() - d0) / 86400000;
    var same = key && k.key === key, score = 0;
    if (Math.abs(k.amount - amount) < 0.01 && days <= 4) score = same ? 3 : 2;
    else if (same && days <= 3 && k.amount >= amount && k.amount <= amount * 1.35) score = 1;
    if (score > bestScore) { best = k; bestScore = score; }
  }
  return best ? best.i + 2 : -1;
}

function txnIndex(sh) {
  var last = sh.getLastRow(), out = {};
  if (last < 2) return out;
  var v = sh.getRange(2, COL.txn, last - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) if (v[i][0]) out[String(v[i][0])] = i + 2;
  return out;
}

/** An uploaded receipt that is this bank charge: same amount within 4 days,
    or the same store within 3 days with the bank amount up to 35% higher
    (a restaurant receipt photographed before the tip). */
function matchReceipt(candidates, place, amount, date) {
  var key = merchantKey(place), best = null, bestScore = 0;
  var d0 = new Date(date + 'T12:00:00').getTime();
  for (var i = 0; i < candidates.length; i++) {
    var c = candidates[i];
    if (!c.date) continue;
    var days = Math.abs(new Date(c.date + 'T12:00:00').getTime() - d0) / 86400000;
    var sameStore = key && c.key === key;
    var score = 0;
    if (Math.abs(c.amount - amount) < 0.01 && days <= 4) score = sameStore ? 3 : 2;
    else if (sameStore && days <= 3 && amount >= c.amount && amount <= c.amount * 1.35) score = 1;
    if (score > bestScore) { best = c; bestScore = score; }
  }
  return best;
}

/** Pad a row out to the full HEADERS width (the recovery tools build 12). */
function padRow(r) {
  while (r.length < HEADERS.length) r.push('');
  return r;
}

/**
 * One-time migration to the HEADERS schema.
 *
 * Maps the OLD columns onto the new ones by reading the existing header row,
 * so it works regardless of what order the sheet is currently in. Duplicates
 * the whole spreadsheet to Drive first — run it once, check the result, and
 * only then keep using the app.
 *
 * Run it from the Apps Script editor: select migrateSheet, press Run.
 */
function migrateSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(TAB_NAME) || ss.getSheets()[0];

  // 1. Back up before touching anything
  var backupName = 'Expenses BACKUP ' +
    Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd HH:mm:ss');
  DriveApp.getFileById(SHEET_ID).makeCopy(backupName);

  var last = sh.getLastRow();
  if (last < 1) { sh.appendRow(HEADERS); return { success: true, migrated: 0, note: 'empty sheet' }; }

  var width   = Math.max(sh.getLastColumn(), HEADERS.length);
  var all     = sh.getRange(1, 1, last, width).getValues();
  var oldHead = all[0].map(function (h) { return String(h).toLowerCase().trim(); });

  // Find a column by trying several spellings, including the typo'd ones
  function col() {
    for (var a = 0; a < arguments.length; a++) {
      var want = String(arguments[a]).toLowerCase();
      for (var i = 0; i < oldHead.length; i++) {
        if (oldHead[i].indexOf(want) !== -1) return i;
      }
    }
    return -1;
  }

  var map = {
    timestamp:  col('timestamp', 'time stamp'),
    place:      col('place', 'vendor', 'store', 'merchant'),
    category:   col('expense cartegory', 'expense category', 'cartegory', 'category'),
    date:       col('date'),
    amount:     col('amount', 'total', 'cost'),
    semester:   col('smester', 'semester'),
    rtype:      col('reciept type', 'receipt type'),
    link:       col('recripet google drive', 'receipt drive', 'drive link', 'google drive', 'receipt link'),
    needs:      col('do you need reimbursement', 'need reimbursement', 'needs reimb'),
    got:        col('have you been reimbursed', 'been reimbursed', 'reimbursed'),
    info:       col('additional info', 'notes', 'note'),
    is529:      col('is this a 529', '529')
  };

  function cell(row, idx) { return idx >= 0 && idx < row.length ? row[idx] : ''; }

  var rebuilt = [HEADERS];
  var count = 0;
  for (var r = 1; r < all.length; r++) {
    var row = all[r];
    var place  = String(cell(row, map.place) || '').trim();
    var amount = Number(cell(row, map.amount)) || 0;
    if (!place && !amount) continue;               // genuinely blank row

    // Old sheets stored 'Not yet' / 'Completed' / 'Not needed' in one column
    var gotRaw = String(cell(row, map.got) || '').toLowerCase();
    var reimbursed = (gotRaw.indexOf('yes') !== -1 || gotRaw.indexOf('complete') !== -1) ? 'Yes' : 'No';
    var needsRaw = String(cell(row, map.needs) || '').toLowerCase();
    var needs = needsRaw ? (needsRaw.indexOf('yes') !== -1 ? 'Yes' : 'No') : 'No';

    rebuilt.push(padRow([
      cell(row, map.timestamp) || cell(row, map.date) || '',
      place || 'Unknown',
      String(cell(row, map.category) || 'Other'),
      formatDate(cell(row, map.date)),
      amount,
      String(cell(row, map.semester) || ''),
      String(cell(row, map.rtype) || 'Receipt'),
      String(cell(row, map.link) || ''),
      needs,
      reimbursed,
      String(cell(row, map.info) || ''),
      String(cell(row, map.is529) || 'Not Sure')
    ]));
    count++;
  }

  // Old dropdown rules (e.g. 'Not needed / Not yet / Completed' on the
  // reimbursement column) reject the new values on write, so they have to go
  // BEFORE setValues. sheet.clear() does not remove data validation.
  var whole = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  whole.clearDataValidations();
  whole.clearFormat();
  whole.clearContent();
  sh.clearConditionalFormatRules();

  sh.getRange(1, 1, rebuilt.length, HEADERS.length).setValues(rebuilt);
  applyFormatting(sh);

  return { success: true, migrated: count, backup: backupName, columnsFound: map };
}

/**
 * The sheet's own presentation and validation: dropdowns, currency, dates,
 * banding and colour-coded 529 status. Safe to re-run at any time.
 */
function applyFormatting(sh) {
  var lastRow = Math.max(sh.getLastRow(), 2);
  var maxRows = sh.getMaxRows();

  // Header
  var head = sh.getRange(1, 1, 1, HEADERS.length);
  head.setValues([HEADERS])
      .setFontWeight('bold')
      .setFontColor('#FFFFFF')
      .setBackground('#3A6644')
      .setVerticalAlignment('middle')
      .setWrap(true);
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 42);

  // Column widths, in HEADERS order
  var widths = [150, 190, 150, 100, 95, 115, 130, 210, 150, 150, 240, 130, 120, 150, 120, 200, 110, 100];
  for (var i = 0; i < widths.length; i++) sh.setColumnWidth(i + 1, widths[i]);

  // Number and date formats
  sh.getRange(2, 4, maxRows - 1, 1).setNumberFormat('yyyy-mm-dd');    // Date
  sh.getRange(2, 5, maxRows - 1, 1).setNumberFormat('$#,##0.00');     // Amount
  sh.getRange(2, 5, maxRows - 1, 1).setHorizontalAlignment('right');

  // Dropdowns — rebuilt from scratch so stale rules can't linger
  function dropdown(colIndex, values) {
    var range = sh.getRange(2, colIndex, maxRows - 1, 1);
    range.clearDataValidations();
    range.setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(values, true)
        .setAllowInvalid(true)   // tolerant: a stray value warns, never blocks a write
        .build()
    );
  }
  dropdown(3,  CATEGORIES);
  dropdown(6,  semesterList());
  dropdown(7,  RECEIPT_TYPES);
  dropdown(9,  YES_NO);
  dropdown(10, YES_NO);
  dropdown(12, YES_NO_MAYBE);

  // Colour-code the 529 column and flag pending reimbursements
  var rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Yes')
    .setBackground('#DCEBE0').setFontColor('#1E4A2C')
    .setRanges([sh.getRange(2, 12, maxRows - 1, 1)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('No')
    .setBackground('#F6DEDC').setFontColor('#8A2E26')
    .setRanges([sh.getRange(2, 12, maxRows - 1, 1)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Not Sure')
    .setBackground('#FAECD9').setFontColor('#A8621E')
    .setRanges([sh.getRange(2, 12, maxRows - 1, 1)]).build());
  // Needs reimbursement but not yet reimbursed
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($I2="Yes",$J2="No")')
    .setBackground('#FBF0DF')
    .setRanges([sh.getRange(2, 1, maxRows - 1, HEADERS.length)]).build());
  sh.setConditionalFormatRules(rules);

  // Alternating row shading
  try {
    var existing = sh.getBandings();
    for (var b = 0; b < existing.length; b++) existing[b].remove();
    sh.getRange(1, 1, lastRow, HEADERS.length)
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  } catch (bandErr) { /* banding is cosmetic — never fail the migration over it */ }

  sh.getRange(2, 1, maxRows - 1, HEADERS.length).setVerticalAlignment('middle');
}

/** Re-apply presentation without touching data. Safe to run any time. */
function formatSheetOnly() {
  applyFormatting(getSheet());
  return { success: true };
}

/* ===================================================================
   RECOVERY
   =================================================================== */

/**
 * DIAGNOSTIC — read-only. Changes nothing.
 *
 * Prints every tab in the live spreadsheet and in every backup, with row
 * counts and the first few rows, so we can see where the expense data
 * actually lives before touching anything.
 */
function inspectAll() {
  var lines = [];

  function describe(label, ss) {
    lines.push('');
    lines.push('=========== ' + label + ' ===========');
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var sh = sheets[i];
      var lr = sh.getLastRow(), lc = sh.getLastColumn();
      lines.push('  TAB "' + sh.getName() + '"  rows=' + lr + '  cols=' + lc);
      var n = Math.min(lr, 3);
      if (n > 0 && lc > 0) {
        var vals = sh.getRange(1, 1, n, Math.min(lc, 14)).getValues();
        for (var r = 0; r < vals.length; r++) {
          lines.push('     row' + (r + 1) + ': ' + JSON.stringify(vals[r]));
        }
      }
    }
  }

  describe('LIVE SPREADSHEET', SpreadsheetApp.openById(SHEET_ID));

  var files = DriveApp.searchFiles('title contains "Expenses BACKUP"');
  while (files.hasNext()) {
    var f = files.next();
    try {
      describe('BACKUP: ' + f.getName() + '   [id ' + f.getId() + ']',
               SpreadsheetApp.openById(f.getId()));
    } catch (err) {
      lines.push('=========== BACKUP ' + f.getName() + ' UNREADABLE: ' + err + ' ===========');
    }
  }

  var text = lines.join('\n');
  Logger.log(text);
  return text;
}

/* -------------------------------------------------------------------
   RESTORE — verified against the real backup layout.

   Backup "Expenses BACKUP 2026-08-16 14:07:38", tab "Expenses", 408 rows.
   Its header row is:
     0  "Nop "                    <- actually the DATE (e.g. 2024-12-18T05:00:00Z)
     1  "Place / Vendor"
     2  "Amount ($)"
     3  "Category"
     4  "529 Eligible"
     5  "Semester"
     6  "Needs Reimbursement"
     7  "Reimbursement Status"    <- 'Not needed' / 'Not yet' / 'Completed'
     8  "Notes"
     9  "Receipt (Google Drive)"
     10 "Submitted At"
   Column 0's header is garbage, which is why header matching failed. These
   indices are hardcoded on purpose — we know exactly what this file is.
   ------------------------------------------------------------------- */
var GOOD_BACKUP_ID = '1_cAhBf_hDR3799dIThwhPhpSPT6zultVlzV-BJu2Njs';
var RESTORE_TAB    = 'Expenses RESTORED';

/**
 * STEP 1 — Non-destructive. Builds a new tab from the backup.
 * Nothing existing is cleared, renamed, or deleted. If the result looks
 * wrong, just delete the new tab and nothing has been lost.
 */
function restoreToNewTab() {
  var src = SpreadsheetApp.openById(GOOD_BACKUP_ID);
  var srcSheet = src.getSheetByName('Expenses');
  if (!srcSheet) throw new Error('Backup has no tab named "Expenses".');

  var last = srcSheet.getLastRow();
  if (last < 2) throw new Error('Backup tab "Expenses" has no data rows.');

  var all = srcSheet.getRange(1, 1, last, 11).getValues();

  var D = 0, PLACE = 1, AMT = 2, CAT = 3, IS529 = 4, SEM = 5,
      NEEDS = 6, STATUS = 7, NOTES = 8, LINK = 9, SUBMITTED = 10;

  var rebuilt = [HEADERS];
  for (var r = 1; r < all.length; r++) {
    var row = all[r];
    var place  = String(row[PLACE] || '').trim();
    var amount = Number(row[AMT]) || 0;
    var date   = formatDate(row[D]);
    if (!place && !amount && !date) continue;

    // Old three-state status collapses into a plain Yes/No
    var status = String(row[STATUS] || '').toLowerCase();
    var reimbursed = (status.indexOf('complete') !== -1) ? 'Yes' : 'No';
    var needs = String(row[NEEDS] || '').toLowerCase().indexOf('yes') !== -1 ? 'Yes' : 'No';

    var is529 = String(row[IS529] || '').trim();
    if (is529 !== 'Yes' && is529 !== 'No') is529 = 'Not Sure';

    rebuilt.push(padRow([
      row[SUBMITTED] ? String(row[SUBMITTED]) : date,
      place || 'Unknown',
      String(row[CAT] || 'Other'),
      date,
      amount,
      String(row[SEM] || ''),
      'Receipt',                       // old schema had no receipt type
      String(row[LINK] || ''),
      needs,
      reimbursed,
      String(row[NOTES] || ''),
      is529
    ]));
  }

  if (rebuilt.length < 2) throw new Error('Built 0 rows — aborting without changes.');

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var old = ss.getSheetByName(RESTORE_TAB);
  if (old) ss.deleteSheet(old);                    // only ever our own scratch tab
  var dest = ss.insertSheet(RESTORE_TAB);

  dest.getRange(1, 1, rebuilt.length, HEADERS.length).setValues(rebuilt);
  applyFormatting(dest);

  var total = 0;
  for (var i = 1; i < rebuilt.length; i++) total += Number(rebuilt[i][4]) || 0;

  Logger.log('Built tab "' + RESTORE_TAB + '" with ' + (rebuilt.length - 1) +
             ' rows, total $' + total.toFixed(2));
  return { success: true, rows: rebuilt.length - 1, total: total, tab: RESTORE_TAB };
}

/**
 * STEP 2 — Run only after eyeballing the restored tab.
 * Renames the tabs so the restored data becomes the live "Expenses".
 * Renames rather than deletes, so the wiped tab is still there if needed.
 */
function swapInRestoredTab() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var restored = ss.getSheetByName(RESTORE_TAB);
  if (!restored) throw new Error('No "' + RESTORE_TAB + '" tab — run restoreToNewTab first.');
  if (restored.getLastRow() < 2) throw new Error('Restored tab is empty — refusing to swap.');

  var live = ss.getSheetByName(TAB_NAME);
  if (live) live.setName('Expenses WIPED ' +
    Utilities.formatDate(new Date(), 'America/New_York', 'MMdd-HHmm'));

  restored.setName(TAB_NAME);
  ss.setActiveSheet(restored);
  Logger.log('Swapped. "' + TAB_NAME + '" now holds ' + (restored.getLastRow() - 1) + ' rows.');
  return { success: true, rows: restored.getLastRow() - 1 };
}

/**
 * List every backup copy this script has made, newest first.
 * Run this, then read the execution log for the file IDs.
 */
function listBackups() {
  var out = [];
  var files = DriveApp.searchFiles('title contains "Expenses BACKUP"');
  while (files.hasNext()) {
    var f = files.next();
    var rows = 0;
    try {
      rows = SpreadsheetApp.openById(f.getId()).getSheets()[0].getLastRow();
    } catch (err) { rows = -1; }
    out.push({ name: f.getName(), id: f.getId(), created: f.getDateCreated(), rows: rows });
  }
  out.sort(function (a, b) { return b.created - a.created; });
  Logger.log(JSON.stringify(out, null, 2));
  return out;
}

/**
 * Rebuild the live sheet from a backup copy.
 *
 * Pass the backup's file ID, or leave it blank to use whichever backup has
 * the MOST rows (which is what you want after a partial wipe).
 *
 * Unlike the earlier migration, this never clears the destination until the
 * replacement rows are built and verified non-empty.
 */
function restoreFromBackup(backupFileId) {
  // 1. Pick the source
  var src = null;
  if (backupFileId) {
    src = SpreadsheetApp.openById(backupFileId);
  } else {
    var best = null;
    var files = DriveApp.searchFiles('title contains "Expenses BACKUP"');
    while (files.hasNext()) {
      var f = files.next();
      try {
        var n = SpreadsheetApp.openById(f.getId()).getSheets()[0].getLastRow();
        if (!best || n > best.rows) best = { id: f.getId(), rows: n, name: f.getName() };
      } catch (err) { /* skip unreadable copies */ }
    }
    if (!best) throw new Error('No backup found in Drive.');
    Logger.log('Using backup: ' + best.name + ' (' + best.rows + ' rows)');
    src = SpreadsheetApp.openById(best.id);
  }

  var srcSheet = src.getSheets()[0];
  var last = srcSheet.getLastRow();
  if (last < 2) throw new Error('That backup has no data rows.');

  var all     = srcSheet.getRange(1, 1, last, srcSheet.getLastColumn()).getValues();
  var oldHead = all[0].map(function (h) { return String(h).toLowerCase().trim(); });

  function col() {
    for (var a = 0; a < arguments.length; a++) {
      var want = String(arguments[a]).toLowerCase();
      for (var i = 0; i < oldHead.length; i++) {
        if (oldHead[i].indexOf(want) !== -1) return i;
      }
    }
    return -1;
  }
  var map = {
    timestamp: col('timestamp', 'time stamp'),
    place:     col('place', 'vendor', 'store', 'merchant', 'description', 'item'),
    category:  col('expense cartegory', 'expense category', 'cartegory', 'category'),
    date:      col('date'),
    amount:    col('amount', 'total', 'cost', 'price'),
    semester:  col('smester', 'semester'),
    rtype:     col('reciept type', 'receipt type'),
    link:      col('recripet google drive', 'receipt drive', 'drive link', 'google drive', 'receipt link', 'link'),
    needs:     col('do you need reimbursement', 'need reimbursement', 'needs reimb'),
    got:       col('have you been reimbursed', 'been reimbursed', 'reimbursed', 'reimb status'),
    info:      col('additional info', 'notes', 'note'),
    is529:     col('is this a 529', '529')
  };
  Logger.log('Column mapping from backup: ' + JSON.stringify(map));
  Logger.log('Backup headers: ' + JSON.stringify(all[0]));

  function cell(row, idx) { return idx >= 0 && idx < row.length ? row[idx] : ''; }

  var rebuilt = [HEADERS];
  for (var r = 1; r < all.length; r++) {
    var row = all[r];
    var place  = String(cell(row, map.place) || '').trim();
    var amount = Number(cell(row, map.amount)) || 0;
    var date   = formatDate(cell(row, map.date));
    // Keep the row if ANY identifying field survives — never require place
    if (!place && !amount && !date) continue;

    var gotRaw = String(cell(row, map.got) || '').toLowerCase();
    var reimbursed = (gotRaw.indexOf('yes') !== -1 || gotRaw.indexOf('complete') !== -1) ? 'Yes' : 'No';
    var needsRaw = String(cell(row, map.needs) || '').toLowerCase();
    var needs = needsRaw ? (needsRaw.indexOf('yes') !== -1 ? 'Yes' : 'No') : 'No';

    rebuilt.push(padRow([
      cell(row, map.timestamp) || date || '',
      place || 'Unknown',
      String(cell(row, map.category) || 'Other'),
      date,
      amount,
      String(cell(row, map.semester) || ''),
      String(cell(row, map.rtype) || 'Receipt'),
      String(cell(row, map.link) || ''),
      needs,
      reimbursed,
      String(cell(row, map.info) || ''),
      String(cell(row, map.is529) || 'Not Sure')
    ]));
  }

  // 2. Refuse to wipe the live sheet unless we actually have rows to put back
  if (rebuilt.length < 2) {
    throw new Error('Built 0 rows from that backup — refusing to clear the live sheet. ' +
                    'Headers seen: ' + JSON.stringify(all[0]));
  }

  // 3. Only now touch the destination
  var dest = getSheet();
  var whole = dest.getRange(1, 1, dest.getMaxRows(), dest.getMaxColumns());
  whole.clearDataValidations();
  whole.clearFormat();
  whole.clearContent();
  dest.clearConditionalFormatRules();

  dest.getRange(1, 1, rebuilt.length, HEADERS.length).setValues(rebuilt);
  applyFormatting(dest);

  Logger.log('RESTORED ' + (rebuilt.length - 1) + ' rows.');
  return { success: true, restored: rebuilt.length - 1 };
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
