// ================================================================
// T4i New Hire Data Sync Script — v10
// Based on v8_Optimized with all fixes (F1–F3) and optimizations (O1–O6)
// Changes marked with: // [v10:XX]
// ================================================================

// ============ Configuration ============
var SOURCE_SPREADSHEET_ID = '1cAlDPmgQ3mkmC5l5Jw-h6HlDre1RqIenAHVpoulZNwU';
var DEST_SPREADSHEET_ID   = '1mt2tOVCtef0y2vP5_Y53E-wpU2MZN_tKbbfDk7bWxbw';

var EARLIEST_YEAR = 2021;
var BATCH_SIZE = 150;
var LOOKBACK_WEEKS = 6;
var SKIP_LOG_MAX = 10;

// Columns to extract (custom display order)
var SOURCE_COLUMNS = [
  'Start Date',
  'Ship Date (t4i use only)',
  'First Name',
  'Last Name',
  'Full name (do not fill)',
  'Username',
  'SNOW Ticket',
  'Hire Type',
  'Intuit Email',
  'Hardware Type: Coder Mac/PC, Pro Mac/PC',
  'TA Notes',
  'Tech Notes',
  'Who is working it',
  'Provisioned by',
  'Shipped',
  'Asset Tag',
  'Tracking number (laptop/full kit)',
  'On Site Location'
];

var OUTPUT_HEADERS = SOURCE_COLUMNS.concat(['Source Tab']);

var KEY_FIELDS = [
  'First Name',
  'Last Name',
  'Full name (do not fill)',
  'Username',
  'Start Date',
  'Ship Date (t4i use only)'
];

// Column indices (0-based)
var START_DATE_IDX = 0;
var SHIP_DATE_IDX = 1;
var FIRST_NAME_IDX = 2;
var LAST_NAME_IDX = 3;
var USERNAME_IDX = 5;
var SNOW_TICKET_IDX = 6;   // [v10:F3] Added for dedup key
var HIRE_TYPE_IDX = 7;
var SOURCE_TAB_IDX = OUTPUT_HEADERS.length - 1;

// ============ Menu ============
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('T4i Sync')
      .addItem('Full Sync (initial)', 'fullSync')
      .addItem('Incremental Sync (daily)', 'syncNewHireData')
      .addItem('Sort & Format', 'runSortAndFormat')
      .addSeparator()
      .addItem('Cleanup Duplicates', 'cleanupDuplicates')
      .addItem('Cleanup Pre-2021 Data', 'cleanupOldData')
      .addItem('Cleanup Irregular Dates', 'cleanupIrregularDates')
      .addItem('Discover Tabs', 'discoverTabs')
      .addSeparator()
      .addItem('Recall (Undo last operation)', 'recall')
      .addSeparator()
      .addItem('Setup Daily Triggers', 'setupDailyTriggers')
      .addItem('Remove Daily Triggers', 'removeDailyTriggers')
      .addItem('Reset Full Sync State', 'resetFullSync')
      .addToUi();
  } catch (e) {}
}

// ================================================================
//  SECTION 1: TAB NAME PARSERS
// ================================================================

function getTabYear_(tabName) {
  var match4 = tabName.match(/\b(20\d{2})\b/);
  if (match4) return parseInt(match4[1]);
  var match2 = tabName.match(/\b(\d{2})\s*$/);
  if (match2) {
    var yr = parseInt(match2[1]);
    if (yr >= 20 && yr <= 30) return 2000 + yr;
  }
  var match2b = tabName.match(/(\d{1,2})\/(\d{1,2})\/(\d{2})\b/);
  if (match2b) return 2000 + parseInt(match2b[3]);
  return null;
}

// Month-only precision — used by isRecentTab_ (month accuracy is sufficient for lookback).
function getTabApproxDate_(tabName) {
  var year = getTabYear_(tabName);
  if (!year) return null;
  var monthMap = [
    'january', 'february', 'march', 'april', 'may', 'june',
    'july', 'august', 'september', 'october', 'november', 'december',
    'jan', 'feb', 'mar', 'apr', 'may', 'jun',
    'jul', 'aug', 'sep', 'oct', 'nov', 'dec'
  ];
  var monthNum = [0,1,2,3,4,5,6,7,8,9,10,11, 0,1,2,3,4,5,6,7,8,9,10,11];
  var lower = tabName.toLowerCase();
  for (var i = 0; i < monthMap.length; i++) {
    if (lower.indexOf(monthMap[i]) !== -1) {
      return new Date(year, monthNum[i], 1);
    }
  }
  return new Date(year, 0, 1);
}

// [v10:F2] Full-precision date extractor — includes day number.
// "February 16 2026" → Date(2026,1,16), "January 6 2025" → Date(2025,0,6)
// "Cohorts February 16 2026" → Date(2026,1,16), "2/16/2026" → Date(2026,1,16)
function getTabFullDate_(tabName) {
  var year = getTabYear_(tabName);
  if (!year) return null;

  // Try M/D/YYYY or M/D/YY first (most precise)
  var slashMatch = tabName.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (slashMatch) {
    return new Date(year, parseInt(slashMatch[1]) - 1, parseInt(slashMatch[2]));
  }

  var monthNames = [
    'january', 'february', 'march', 'april', 'may', 'june',
    'july', 'august', 'september', 'october', 'november', 'december',
    'jan', 'feb', 'mar', 'apr', 'may', 'jun',
    'jul', 'aug', 'sep', 'oct', 'nov', 'dec'
  ];
  var monthNums = [0,1,2,3,4,5,6,7,8,9,10,11, 0,1,2,3,4,5,6,7,8,9,10,11];
  var lower = tabName.toLowerCase();

  for (var i = 0; i < monthNames.length; i++) {
    var mIdx = lower.indexOf(monthNames[i]);
    if (mIdx !== -1) {
      var month = monthNums[i];
      var afterMonth = lower.substring(mIdx + monthNames[i].length);
      var dayMatch = afterMonth.match(/\s+(\d{1,2})\b/);
      var day = 1;
      if (dayMatch) {
        var d = parseInt(dayMatch[1]);
        if (d >= 1 && d <= 31) day = d;
      }
      return new Date(year, month, day);
    }
  }

  return new Date(year, 0, 1);
}

function isRecentTab_(tabName, cutoffDate) {
  var tabDate = getTabApproxDate_(tabName);
  if (!tabDate) return true;
  return tabDate >= cutoffDate;
}

// ================================================================
//  SECTION 2: CORE SHEET PROCESSOR
// ================================================================

function processSheet_(sheet, sheetName, today) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 3 || lastCol === 0) return [];

  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = [];
  for (var h = 0; h < data[1].length; h++) {
    headers.push(String(data[1][h]).trim());
  }

  var colMap = {};
  var matchCount = 0;
  for (var c = 0; c < SOURCE_COLUMNS.length; c++) {
    var idx = headers.indexOf(SOURCE_COLUMNS[c]);
    if (idx !== -1) {
      colMap[SOURCE_COLUMNS[c]] = idx;
      matchCount++;
    }
  }
  if (matchCount === 0) return [];

  var earliestDate = EARLIEST_YEAR ? new Date(EARLIEST_YEAR, 0, 1) : null;

  var rows = [];
  for (var i = 2; i < data.length; i++) {
    var row = data[i];

    var hasKey = false;
    for (var k = 0; k < KEY_FIELDS.length; k++) {
      var kf = KEY_FIELDS[k];
      if (colMap[kf] !== undefined) {
        var val = row[colMap[kf]];
        if (val !== '' && val !== null && val !== undefined) {
          hasKey = true;
          break;
        }
      }
    }
    if (!hasKey) continue;

    var startRaw = colMap['Start Date'] !== undefined ? row[colMap['Start Date']] : null;
    var shipRaw = colMap['Ship Date (t4i use only)'] !== undefined ? row[colMap['Ship Date (t4i use only)']] : null;

    if (hasIrregularDate_(startRaw) || hasIrregularDate_(shipRaw)) continue;

    var startDate = parseDate_(startRaw);
    var shipDate = parseDate_(shipRaw);

    if (earliestDate && (startDate || shipDate)) {
      var allTooOld = true;
      if (startDate && startDate >= earliestDate) allTooOld = false;
      if (shipDate && shipDate >= earliestDate) allTooOld = false;
      if (allTooOld) continue;
    }

    if (startDate || shipDate) {
      var allFuture = true;
      if (startDate && startDate <= today) allFuture = false;
      if (shipDate && shipDate <= today) allFuture = false;
      if (allFuture) continue;
    }

    var outRow = [];
    for (var d = 0; d < SOURCE_COLUMNS.length; d++) {
      var colName = SOURCE_COLUMNS[d];
      if (colMap[colName] !== undefined) {
        var cellVal = row[colMap[colName]];
        if ((colName === 'Start Date' || colName === 'Ship Date (t4i use only)') && cellVal instanceof Date && !isNaN(cellVal.getTime())) {
          cellVal = new Date(cellVal.getFullYear(), cellVal.getMonth(), cellVal.getDate());
        }
        outRow.push(cellVal);
      } else {
        outRow.push('');
      }
    }
    outRow.push(sheetName);
    rows.push(outRow);
  }

  return rows;
}

// ================================================================
//  SECTION 2b: UNDO / RECALL
// ================================================================

// [v10:O4] Returns { ss, sheet } so callers can reuse the spreadsheet handle.
function saveUndoSnapshot_() {
  var destSS = SpreadsheetApp.openById(DEST_SPREADSHEET_ID);
  var mainSheet = destSS.getSheets()[0];
  var lastRow = mainSheet.getLastRow();
  var lastCol = mainSheet.getLastColumn();

  if (lastRow >= 1 && lastCol >= 1) {
    var undoSheet = destSS.getSheetByName('_Undo');
    if (!undoSheet) {
      undoSheet = destSS.insertSheet('_Undo', destSS.getSheets().length);
    }
    undoSheet.clearContents();
    undoSheet.clearFormats();
    mainSheet.getRange(1, 1, lastRow, lastCol).copyTo(undoSheet.getRange(1, 1));
    undoSheet.hideSheet();
  }

  return { ss: destSS, sheet: mainSheet };
}

// [v10:O6] Restores content AND column formats (M/d/yyyy, plain text).
function recall() {
  var destSS = SpreadsheetApp.openById(DEST_SPREADSHEET_ID);
  var mainSheet = destSS.getSheets()[0];
  var undoSheet = destSS.getSheetByName('_Undo');
  if (!undoSheet) {
    Logger.log('Recall: No undo snapshot found. Run a sync or cleanup first.');
    return;
  }
  var lastRow = undoSheet.getLastRow();
  var lastCol = undoSheet.getLastColumn();
  if (lastRow < 1) {
    Logger.log('Recall: Undo snapshot is empty.');
    return;
  }
  mainSheet.clearContents();
  mainSheet.clearFormats();
  undoSheet.getRange(1, 1, lastRow, lastCol).copyTo(mainSheet.getRange(1, 1));

  if (lastRow >= 1) {
    mainSheet.getRange(1, 1, 1, lastCol).setFontWeight('bold');
    mainSheet.setFrozenRows(1);
  }

  // [v10:O6] Re-enforce date and Source Tab formats after restore
  var dataRows = Math.max(lastRow - 1, 1);
  mainSheet.getRange(2, START_DATE_IDX + 1, dataRows, 1).setNumberFormat('M/d/yyyy');
  mainSheet.getRange(2, SHIP_DATE_IDX + 1, dataRows, 1).setNumberFormat('M/d/yyyy');
  mainSheet.getRange(2, SOURCE_TAB_IDX + 1, dataRows, 1).setNumberFormat('@');

  SpreadsheetApp.flush();
  Logger.log('Recall: Restored ' + (lastRow - 1) + ' data rows from undo snapshot.');
}

// ================================================================
//  SECTION 3: SHARED WRITE HELPER
// ================================================================

function writeDestData_(destSheet, dataRows) {
  var totalCols = OUTPUT_HEADERS.length;

  destSheet.clearContents();
  destSheet.clearFormats();

  destSheet.getRange(1, 1, 1, totalCols).setValues([OUTPUT_HEADERS]);
  destSheet.getRange(1, 1, 1, totalCols).setFontWeight('bold');
  destSheet.setFrozenRows(1);

  if (dataRows.length > 0) {
    destSheet.getRange(2, 1, dataRows.length, totalCols).setValues(dataRows);
  }

  var formatRows = Math.max(dataRows.length, 1);
  destSheet.getRange(2, START_DATE_IDX + 1, formatRows, 1).setNumberFormat('M/d/yyyy');
  destSheet.getRange(2, SHIP_DATE_IDX + 1, formatRows, 1).setNumberFormat('M/d/yyyy');
  destSheet.getRange(2, SOURCE_TAB_IDX + 1, formatRows, 1).setNumberFormat('@');

  SpreadsheetApp.flush();
}

// ================================================================
//  SECTION 4: SORT & FORMAT
// ================================================================

function sortAndFormatSheet_(destSheet) {
  var lastRow = destSheet.getLastRow();
  if (lastRow <= 1) return;

  var totalCols = OUTPUT_HEADERS.length;
  var startCol = START_DATE_IDX + 1;
  var shipCol = SHIP_DATE_IDX + 1;
  var dataRows = lastRow - 1;

  var dataRange = destSheet.getRange(2, 1, dataRows, totalCols);
  dataRange.sort([
    {column: startCol, ascending: true},
    {column: shipCol, ascending: true}
  ]);

  destSheet.getRange(2, startCol, dataRows, 1).setNumberFormat('M/d/yyyy');
  destSheet.getRange(2, shipCol, dataRows, 1).setNumberFormat('M/d/yyyy');
  destSheet.getRange(2, SOURCE_TAB_IDX + 1, dataRows, 1).setNumberFormat('@');

  SpreadsheetApp.flush();
  Logger.log('Sort and format complete. Rows: ' + dataRows);
}

function runSortAndFormat() {
  var dest = saveUndoSnapshot_();  // [v10:O4]
  Logger.log('Running sort and format...');
  sortAndFormatSheet_(dest.sheet);
  Logger.log('Done!');
}

// ================================================================
//  SECTION 5: CLEANUP FUNCTIONS
// ================================================================

function cleanupOldData() {
  var dest = saveUndoSnapshot_();  // [v10:O4]
  var destSheet = dest.sheet;
  var lastRow = destSheet.getLastRow();
  if (lastRow <= 1) return;

  var totalCols = OUTPUT_HEADERS.length;
  var earliestDate = new Date(EARLIEST_YEAR, 0, 1);

  var data = destSheet.getRange(2, 1, lastRow - 1, totalCols).getValues();
  var kept = [];
  var removed = 0;

  for (var i = 0; i < data.length; i++) {
    var sd = parseDate_(data[i][START_DATE_IDX]);
    var shd = parseDate_(data[i][SHIP_DATE_IDX]);
    if (sd || shd) {
      var allOld = true;
      if (sd && sd >= earliestDate) allOld = false;
      if (shd && shd >= earliestDate) allOld = false;
      if (allOld) { removed++; continue; }
    }
    kept.push(data[i]);
  }

  writeDestData_(destSheet, kept);
  Logger.log('Old data cleanup: Removed ' + removed + ', Kept ' + kept.length);
}

// [v10:O2] isEmpty defined once outside loop.
function cleanupIrregularDates() {
  var dest = saveUndoSnapshot_();  // [v10:O4]
  var destSheet = dest.sheet;
  var lastRow = destSheet.getLastRow();
  if (lastRow <= 1) return;

  var totalCols = OUTPUT_HEADERS.length;
  var data = destSheet.getRange(2, 1, lastRow - 1, totalCols).getValues();
  var kept = [];
  var removed = 0;
  var examples = [];

  var isEmpty = function(v) { return v === null || v === undefined || (typeof v === 'string' && v.trim() === ''); };

  for (var i = 0; i < data.length; i++) {
    var sd = data[i][START_DATE_IDX];
    var shd = data[i][SHIP_DATE_IDX];
    if (isEmpty(sd) && isEmpty(shd)) {
      removed++;
      if (examples.length < 10) examples.push('Start="' + sd + '", Ship="' + shd + '" (both empty)');
      continue;
    }
    if (hasIrregularDate_(sd) || hasIrregularDate_(shd)) {
      removed++;
      if (examples.length < 10) {
        examples.push('Start="' + sd + '", Ship="' + shd + '"');
      }
      continue;
    }
    kept.push(data[i]);
  }

  writeDestData_(destSheet, kept);
  Logger.log('Irregular date cleanup: Removed ' + removed + ', Kept ' + kept.length);
  if (examples.length > 0) {
    Logger.log('Examples removed:\n' + examples.join('\n'));
  }
}

// [v10:F3] SNOW Ticket added to dedup key — prevents collision when names are empty.
// [v10:O1] All text fields lowercased for case-insensitive matching.
// [v10:O3] In-memory sort replaces writeDestData_ + sortAndFormatSheet_ double I/O.
function cleanupDuplicates(skipUndoSave) {
  var destSS, destSheet;
  if (!skipUndoSave) {
    var dest = saveUndoSnapshot_();  // [v10:O4]
    destSS = dest.ss;
    destSheet = dest.sheet;
  } else {
    destSS = SpreadsheetApp.openById(DEST_SPREADSHEET_ID);
    destSheet = destSS.getSheets()[0];
  }

  var lastRow = destSheet.getLastRow();
  if (lastRow <= 1) return;

  var totalCols = OUTPUT_HEADERS.length;
  var data = destSheet.getRange(2, 1, lastRow - 1, totalCols).getValues();

  var seen = {};
  var kept = [];
  var removed = 0;
  var REMOVED_LOG_MAX = 30;
  var removedLog = [];

  for (var i = 0; i < data.length; i++) {
    var firstName  = String(data[i][FIRST_NAME_IDX] || '').trim().toLowerCase();   // [v10:O1]
    var lastName   = String(data[i][LAST_NAME_IDX] || '').trim().toLowerCase();    // [v10:O1]
    var username   = String(data[i][USERNAME_IDX] || '').trim().toLowerCase();     // [v10:O1]
    var snowTicket = String(data[i][SNOW_TICKET_IDX] || '').trim().toLowerCase();  // [v10:F3]
    var startDate  = parseDate_(data[i][START_DATE_IDX]);
    var startStr   = startDate ? startDate.toDateString() : '';
    var shipDate   = parseDate_(data[i][SHIP_DATE_IDX]);
    var shipStr    = shipDate ? shipDate.toDateString() : '';
    var hireType   = String(data[i][HIRE_TYPE_IDX] || '').trim().toLowerCase();    // [v10:O1]
    var sourceTab  = canonicalSourceTabForDedup_(data[i][SOURCE_TAB_IDX]);

    var key = firstName + '|' + lastName + '|' + username + '|' + snowTicket + '|' + startStr + '|' + shipStr + '|' + hireType + '|' + sourceTab;

    if (seen[key]) {
      removed++;
      if (removedLog.length < REMOVED_LOG_MAX) {
        removedLog.push('[' + removed + '] ' + firstName + ' ' + lastName + ' | ' + username + ' | SNOW: ' + snowTicket + ' | Start: ' + startStr + ' | Ship: ' + shipStr + ' | ' + sourceTab);
      }
      continue;
    }
    seen[key] = true;
    kept.push(data[i]);
  }

  // [v10:O3] Sort in memory — eliminates the second read/write cycle from sortAndFormatSheet_.
  kept.sort(function(a, b) {
    var aStart = parseDate_(a[START_DATE_IDX]);
    var bStart = parseDate_(b[START_DATE_IDX]);
    var aTime = aStart ? aStart.getTime() : 0;
    var bTime = bStart ? bStart.getTime() : 0;
    if (aTime !== bTime) return aTime - bTime;
    var aShip = parseDate_(a[SHIP_DATE_IDX]);
    var bShip = parseDate_(b[SHIP_DATE_IDX]);
    return (aShip ? aShip.getTime() : 0) - (bShip ? bShip.getTime() : 0);
  });

  writeDestData_(destSheet, kept);

  Logger.log('Duplicate cleanup: Removed ' + removed + ', Kept ' + kept.length);
  if (removedLog.length > 0) {
    Logger.log('Removed rows (up to ' + REMOVED_LOG_MAX + '):\n' + removedLog.join('\n'));
    if (removed > REMOVED_LOG_MAX) {
      Logger.log('... and ' + (removed - REMOVED_LOG_MAX) + ' more');
    }
  }
}

// ================================================================
//  SECTION 6: FULL SYNC (Batch)
// ================================================================

function fullSync() {
  var props = PropertiesService.getScriptProperties();
  var sheetStart = parseInt(props.getProperty('FULL_SYNC_INDEX') || '0');

  if (props.getProperty('FULL_SYNC_PHASE') === 'complete') {
    Logger.log('Full sync already completed. Run resetFullSync() to restart.');
    return;
  }

  var sourceSS = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  var destSS, destSheet;

  if (sheetStart === 0) {
    var dest = saveUndoSnapshot_();  // [v10:O4]
    destSS = dest.ss;
    destSheet = dest.sheet;
    destSheet.clearContents();
    destSheet.clearFormats();
    destSheet.getRange(1, 1, 1, OUTPUT_HEADERS.length).setValues([OUTPUT_HEADERS]);
    destSheet.getRange(1, 1, 1, OUTPUT_HEADERS.length).setFontWeight('bold');
    destSheet.setFrozenRows(1);
    destSheet.getRange(2, SOURCE_TAB_IDX + 1, 1, 1).setNumberFormat('@');
    Logger.log('=== Full Sync Started ===');
    Logger.log('Total source tabs: ' + sourceSS.getSheets().length);
  } else {
    destSS = SpreadsheetApp.openById(DEST_SPREADSHEET_ID);
    destSheet = destSS.getSheets()[0];
    Logger.log('=== Full Sync Continuing from sheet index ' + sheetStart + ' ===');
  }

  var sheets = sourceSS.getSheets();
  var today = new Date();
  today.setHours(23, 59, 59, 999);

  var processed = 0;
  var batchRows = 0;
  var errorTabs = [];  // [v10:O5]
  var s = sheetStart;

  for (; s < sheets.length; s++) {
    var sheet = sheets[s];
    var sheetName = sheet.getName();

    if (EARLIEST_YEAR) {
      var tabYear = getTabYear_(sheetName);
      if (tabYear !== null && tabYear < EARLIEST_YEAR) continue;
    }

    // [v10:O5] try-catch: single tab error does not abort entire sync
    var rows;
    try {
      rows = processSheet_(sheet, sheetName, today);
    } catch (e) {
      errorTabs.push(sheetName + ': ' + e.message);
      Logger.log('  ERROR processing "' + sheetName + '": ' + e.message);
      rows = [];
    }

    if (rows.length > 0) {
      var lastRow = destSheet.getLastRow();
      destSheet.getRange(lastRow + 1, 1, rows.length, OUTPUT_HEADERS.length).setValues(rows);
      destSheet.getRange(lastRow + 1, SOURCE_TAB_IDX + 1, rows.length, 1).setNumberFormat('@');
      destSheet.getRange(lastRow + 1, START_DATE_IDX + 1, rows.length, 1).setNumberFormat('M/d/yyyy');
      destSheet.getRange(lastRow + 1, SHIP_DATE_IDX + 1, rows.length, 1).setNumberFormat('M/d/yyyy');
      batchRows += rows.length;
    }

    processed++;
    Logger.log('  [' + processed + '/' + BATCH_SIZE + '] ' + sheetName + ' → ' + rows.length + ' rows');

    if (processed >= BATCH_SIZE) {
      s++;
      break;
    }
  }

  SpreadsheetApp.flush();
  Logger.log('Batch result: ' + processed + ' tabs, ' + batchRows + ' rows added.');
  if (errorTabs.length > 0) {
    Logger.log('Tabs with errors (' + errorTabs.length + '):\n' + errorTabs.join('\n'));
  }

  if (s >= sheets.length) {
    props.setProperty('FULL_SYNC_PHASE', 'complete');
    props.deleteProperty('FULL_SYNC_INDEX');
    removeContinuationTriggers_();

    Logger.log('Sorting and formatting...');
    sortAndFormatSheet_(destSheet);

    Logger.log('Cleaning up duplicates...');
    cleanupDuplicates(true);

    var totalRows = destSheet.getLastRow() - 1;
    Logger.log('=== Full Sync Complete! Total rows: ' + totalRows + ' ===');
  } else {
    props.setProperty('FULL_SYNC_INDEX', String(s));
    ScriptApp.newTrigger('fullSync')
      .timeBased()
      .after(60 * 1000)
      .create();
    Logger.log('More tabs remain. Auto-continuing in ~1 minute...');
  }
}

function resetFullSync() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('FULL_SYNC_INDEX');
  props.deleteProperty('FULL_SYNC_PHASE');
  removeContinuationTriggers_();
  Logger.log('Full sync state reset.');
}

function removeContinuationTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'fullSync') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// ================================================================
//  SECTION 7: INCREMENTAL DAILY SYNC
// ================================================================

// [v10:F1] recentTabSet replaced by canonical comparison — eliminates mismatch that caused 709 duplicates per sync.
function syncNewHireData() {
  var dest = saveUndoSnapshot_();  // [v10:O4]
  var destSheet = dest.sheet;

  var startTime = new Date();
  Logger.log('=== Incremental Sync Started: ' + startTime.toLocaleString() + ' ===');
  Logger.log('Lookback: ' + LOOKBACK_WEEKS + ' weeks');

  var sourceSS = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  var sheets = sourceSS.getSheets();
  var today = new Date();
  today.setHours(23, 59, 59, 999);

  var cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - (LOOKBACK_WEEKS * 7));

  // Step 1: Find recent tabs
  var recentTabNames = [];
  var recentSheets = [];
  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var name = sheet.getName();
    if (EARLIEST_YEAR) {
      var tabYear = getTabYear_(name);
      if (tabYear !== null && tabYear < EARLIEST_YEAR) continue;
    }
    if (isRecentTab_(name, cutoffDate)) {
      recentTabNames.push(name);
      recentSheets.push(sheet);
    }
  }
  Logger.log('Recent tabs to refresh: ' + recentTabNames.length);

  // [v10:F1] Build canonical lookup set — both sides use the same canonicalization.
  var recentCanonicalSet = {};
  for (var t = 0; t < recentTabNames.length; t++) {
    recentCanonicalSet[canonicalSourceTabForDedup_(recentTabNames[t])] = true;
  }

  // Step 2: Preserve historical rows using canonical Source Tab matching
  var historicalRows = [];
  var lastRow = destSheet.getLastRow();
  if (lastRow > 1) {
    var existingData = destSheet.getRange(2, 1, lastRow - 1, OUTPUT_HEADERS.length).getValues();
    for (var i = 0; i < existingData.length; i++) {
      var rowCanonical = canonicalSourceTabForDedup_(existingData[i][SOURCE_TAB_IDX]);  // [v10:F1]
      if (!recentCanonicalSet[rowCanonical]) {
        historicalRows.push(existingData[i]);
      }
    }
    var removedCount = existingData.length - historicalRows.length;
    Logger.log('Historical rows preserved: ' + historicalRows.length + ' (removed ' + removedCount + ' for refresh)');
  }

  // Step 3: Process recent tabs
  var freshRows = [];
  var errorTabs = [];  // [v10:O5]
  for (var r = 0; r < recentSheets.length; r++) {
    // [v10:O5] try-catch per tab
    var rows;
    try {
      rows = processSheet_(recentSheets[r], recentTabNames[r], today);
    } catch (e) {
      errorTabs.push(recentTabNames[r] + ': ' + e.message);
      Logger.log('  ERROR processing "' + recentTabNames[r] + '": ' + e.message);
      rows = [];
    }
    freshRows = freshRows.concat(rows);
    if (rows.length > 0) {
      Logger.log('  Refreshed: ' + recentTabNames[r] + ' → ' + rows.length + ' rows');
    }
  }
  Logger.log('Fresh rows: ' + freshRows.length);
  if (errorTabs.length > 0) {
    Logger.log('Tabs with errors (' + errorTabs.length + '):\n' + errorTabs.join('\n'));
  }

  // Step 4: Combine and write
  var allRows = historicalRows.concat(freshRows);
  writeDestData_(destSheet, allRows);

  // Step 5: Sort and format
  sortAndFormatSheet_(destSheet);

  // Step 6: Dedup
  cleanupDuplicates(true);

  var elapsed = ((new Date() - startTime) / 1000).toFixed(1);
  Logger.log('=== Incremental Sync Complete: ' + (destSheet.getLastRow() - 1) + ' rows, ' + elapsed + 's ===');
}

// ================================================================
//  SECTION 8: DISCOVERY TOOL
// ================================================================

function discoverTabs() {
  var ss = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  var sheets = ss.getSheets();
  var matchCount = 0;
  var skipCount = 0;
  var yearSkipCount = 0;

  Logger.log('=== Tab Discovery Report ===');
  Logger.log('EARLIEST_YEAR: ' + (EARLIEST_YEAR || 'ALL'));
  Logger.log('Total tabs: ' + sheets.length);

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var name = sheet.getName();
    if (EARLIEST_YEAR) {
      var tabYear = getTabYear_(name);
      if (tabYear !== null && tabYear < EARLIEST_YEAR) { yearSkipCount++; continue; }
    }
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol === 0) { skipCount++; continue; }

    var headerRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
    var headers = [];
    for (var i = 0; i < headerRow.length; i++) headers.push(String(headerRow[i]).trim());

    var matched = [];
    var missing = [];
    for (var j = 0; j < SOURCE_COLUMNS.length; j++) {
      if (headers.indexOf(SOURCE_COLUMNS[j]) !== -1) matched.push(SOURCE_COLUMNS[j]);
      else missing.push(SOURCE_COLUMNS[j]);
    }

    if (matched.length > 0) {
      matchCount++;
      var msg = '[MATCH] ' + name + ' — ' + matched.length + '/' + SOURCE_COLUMNS.length;
      if (missing.length > 0 && missing.length <= 5) msg += ' | Missing: ' + missing.join(', ');
      Logger.log(msg);
    } else {
      skipCount++;
    }
  }

  Logger.log('=== Summary: ' + matchCount + ' match, ' + skipCount + ' skip, ' + yearSkipCount + ' pre-' + EARLIEST_YEAR + ' ===');
}

// ================================================================
//  SECTION 9: HELPERS
// ================================================================

function parseDate_(val) {
  if (val === null || val === undefined || val === '') return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  var d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

function hasIrregularDate_(val) {
  if (val === null || val === undefined || val === '') return false;
  if (val instanceof Date) return isNaN(val.getTime());
  if (typeof val === 'number') {
    if (val < 36526 || val > 47848) return true;
    return false;
  }
  var str = String(val).trim();
  if (str !== '' && str.indexOf('//') !== -1) return true;
  if (str !== '' && /\d[\/\-]\d{1,2}\.\d/.test(str)) return true;
  var d = new Date(val);
  if (!isNaN(d.getTime())) {
    var yr = d.getFullYear();
    if (yr < 2000 || yr > 2030) return true;
    return false;
  }
  return true;
}

function normalizeSourceTab_(val) {
  if (val === null || val === undefined || val === '') return '';
  if (val instanceof Date && !isNaN(val.getTime())) {
    return (val.getMonth() + 1) + '/' + val.getDate() + '/' + val.getFullYear();
  }
  return String(val).trim();
}

// [v10:F2] Uses getTabFullDate_ (day-level precision) instead of getTabApproxDate_ (month-only).
// "February 16 2026" → "2/16/2026", "February 23 2026" → "2/23/2026" (no longer collapsed).
function canonicalSourceTabForDedup_(val) {
  if (val === null || val === undefined || val === '') return '';
  var str = '';
  if (val instanceof Date && !isNaN(val.getTime())) {
    str = (val.getMonth() + 1) + '/' + val.getDate() + '/' + val.getFullYear();
  } else {
    str = String(val).trim();
  }
  if (str === '') return '';
  var fullDate = getTabFullDate_(str);  // [v10:F2]
  if (fullDate) {
    return (fullDate.getMonth() + 1) + '/' + fullDate.getDate() + '/' + fullDate.getFullYear();
  }
  return str;
}

// ================================================================
//  SECTION 10: TRIGGER MANAGEMENT
// ================================================================

function setupDailyTriggers() {
  removeDailyTriggers();
  ScriptApp.newTrigger('syncNewHireData')
    .timeBased().atHour(6).everyDays(1).inTimezone('America/New_York').create();
  ScriptApp.newTrigger('syncNewHireData')
    .timeBased().atHour(20).everyDays(1).inTimezone('America/New_York').create();
  Logger.log('Daily triggers created: 6am ET & 8pm ET');
}

function removeDailyTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var count = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'syncNewHireData') {
      ScriptApp.deleteTrigger(triggers[i]);
      count++;
    }
  }
  Logger.log('Removed ' + count + ' daily trigger(s).');
}

function removeTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  Logger.log('All triggers removed.');
}
