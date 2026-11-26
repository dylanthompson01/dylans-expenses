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
  'ID'                        // stable row id, so edits and deletes hit the right row
];
var ID_COL = 13;              // 1-based column of 'ID'

// Allowed values, used to build the in-sheet dropdowns and to validate edits
var CATEGORIES    = ['Food & Dining','Tuition & Fees','Books & Supplies','Housing',
                     'Transportation','Technology','Healthcare','Entertainment',
                     'Clothing','Personal Care','School Supplies','Other'];
var RECEIPT_TYPES = ['Receipt','Bank Statement'];
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
    e.id || newId()
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

function getSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
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

/** One writer at a time, so two saves can't shift each other's rows. */
function withLock(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

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
    return json(loadAll());
  } catch (err2) {
    return json({ success: false, error: String(err2), apiVersion: API_VERSION });
  }
}

function loadAll() {
  var sh = getSheet();
  ensureIds(sh);
  var out = readExpenses(sh);
  return { success: true, apiVersion: API_VERSION, expenses: out, count: out.length,
           budgets: readBudgets() };
}

function readExpenses(sh) {
  var last = sh.getLastRow();
  var out = [];
  if (last > 1) {
    var values = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      if (!row[1] && !row[4]) continue;          // skip blank rows
      out.push(toObject(row, i));
    }
  }
  return out;
}

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
    p.id = id;
    p.receiptUrl = receiptUrl;
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
    var n = Math.round((Number(v) || 0) * 100) / 100;
    if (n < 0) throw new Error('Amount cannot be negative');
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

  return withLock(function () {
    var sh = getSheet();
    var updated = [], missing = [];
    for (var i = 0; i < ids.length; i++) {
      var row = findRow(sh, ids[i]);
      if (row < 0) { missing.push(ids[i]); continue; }
      for (var f in clean) sh.getRange(row, EDITABLE[f] + 1).setValue(clean[f]);
      updated.push(toObject(sh.getRange(row, 1, 1, HEADERS.length).getValues()[0], row - 2));
    }
    return { success: missing.length === 0, updated: updated, missing: missing,
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

    var ss = SpreadsheetApp.openById(SHEET_ID);
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
  var ss = SpreadsheetApp.openById(SHEET_ID);
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
  var widths = [150, 190, 150, 100, 95, 115, 130, 210, 150, 150, 240, 130, 120];
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
