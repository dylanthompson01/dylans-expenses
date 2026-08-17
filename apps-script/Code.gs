/**
 * Dylan's Expenses — Google Apps Script backend
 * ---------------------------------------------
 * Paste this into the Apps Script editor bound to the expenses spreadsheet,
 * then Deploy > Manage deployments > Edit > Version: New version > Deploy.
 * The web app must be deployed as "Execute as: Me" and
 * "Who has access: Anyone" so the browser can reach it without a login.
 *
 * Column order is defined once, in HEADERS. index.html writes and reads the
 * same order, so the two must be changed together.
 */

var SHEET_ID  = '193exXTS041Quq_-VOS_0uXMMAv6RM6eFuKHMgkUkIXo';
var FOLDER_ID = '1VtL_8OtIyHmute0nbAcr0LgCnoE6QLbV';
var TAB_NAME  = 'Expenses';

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
  'Is This A 529 Expense'
];

// Allowed values, used to build the in-sheet dropdowns
var CATEGORIES    = ['Food & Dining','Tuition & Fees','Books & Supplies','Housing',
                     'Transportation','Technology','Healthcare','Entertainment',
                     'Clothing','Personal Care','School Supplies','Other'];
var SEMESTER_LIST = ['Spring 2025','Summer 2025','Fall 2025','Spring 2026',
                     'Summer 2026','Fall 2026','Spring 2027'];
var RECEIPT_TYPES = ['Receipt','Bank Statement'];
var YES_NO        = ['Yes','No'];
var YES_NO_MAYBE  = ['Yes','No','Not Sure'];

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
    e.is529 || 'Not Sure'
  ];
}

/** Row -> object, in HEADERS order. */
function toObject(row, index) {
  return {
    id:          'sheet-' + index,
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
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    applyFormatting(sh);
  }
  return sh;
}

/** GET ?action=load — every row as JSON. */
function doGet(e) {
  try {
    var sh = getSheet();
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
    return json({ success: true, expenses: out, count: out.length });
  } catch (err) {
    return json({ success: false, error: String(err) });
  }
}

/** POST — one expense, with an optional base64 receipt image. */
function doPost(e) {
  try {
    var payload = JSON.parse(e.parameter.payload);

    if (payload.action === 'migrate') return json(migrateSheet());

    // Store the image first so the row carries its link
    if (payload.imageBase64) {
      try {
        var folder = DriveApp.getFolderById(FOLDER_ID);
        var stamp  = Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd_HHmmss');
        var safe   = String(payload.place || 'receipt').replace(/[^\w\s-]/g, '').trim().slice(0, 40);
        var label  = (payload.receiptType === 'Bank Statement') ? 'statement' : 'receipt';
        var blob   = Utilities.newBlob(
          Utilities.base64Decode(payload.imageBase64),
          'image/jpeg',
          stamp + '_' + label + '_' + safe + '.jpg'
        );
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        payload.receiptUrl = file.getUrl();
      } catch (imgErr) {
        // A Drive failure must not lose the expense itself
        payload.receiptUrl = 'UPLOAD_FAILED: ' + String(imgErr);
      }
    }

    getSheet().appendRow(toRow(payload));
    return json({ success: true, receiptUrl: payload.receiptUrl || '' });
  } catch (err) {
    return json({ success: false, error: String(err) });
  }
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

    rebuilt.push([
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
    ]);
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
  var widths = [150, 190, 150, 100, 95, 115, 130, 210, 150, 150, 240, 130];
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
  dropdown(6,  SEMESTER_LIST);
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

    rebuilt.push([
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
    ]);
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

    rebuilt.push([
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
    ]);
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
