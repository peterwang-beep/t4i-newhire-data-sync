// ================================================================
// T4i New Hire Data Sync Script — v7 (Irregular Date Filter)
// Syncs data from 'T4i New Hire List' to 'newHireAutoSyncData'
//
// v7 Changes:
//   - NEW: Skip rows where Start Date or Ship Date contains non-date text
//     (e.g., "TBD", "N/A", "pending", "on hold", etc.)
//   - These rows are excluded until the value is updated to a real date
//   - Added hasIrregularDate_() helper function
// ================================================================

// ============ Configuration ============
var SOURCE_SPREADSHEET_ID = '1cAlDPmgQ3mkmC5l5Jw-h6HlDre1RqIenAHVpoulZNwU';
var DEST_SPREADSHEET_ID   = '1mt2tOVCtef0y2vP5_Y53E-wpU2MZN_tKbbfDk7bWxbw';

var EARLIEST_YEAR = 2021;
var BATCH_SIZE = 150;
var LOOKBACK_WEEKS = 6;

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

// ============ Menu ============
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('T4i Sync')
      .addItem('Full Sync (initial)', 'fullSync')
      .addItem('Incremental Sync (daily)', 'syncNewHireData')
      .addItem('Sort & Format', 'runSortAndFormat')
      .addItem('Cleanup Pre-2021 Data', 'cleanupOldData')
      .addItem('Cleanup Irregular Dates', 'cleanupIrregularDates')
      .addItem('Discover Tabs', 'discoverTabs')
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

    // Key field check
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

    // Get raw date values
    var startRaw = colMap['Start Date'] !== undefined ? row[colMap['Start Date']] : null;
    var shipRaw = colMap['Ship Date (t4i use only)'] !== undefined ? row[colMap['Ship Date (t4i use only)']] : null;

    // ---- NEW in v7: Irregular date check ----
    // If a date field has a non-empty value that is NOT a valid date
    // (e.g., "TBD", "N/A", "pending", "on hold"), skip the entire row.
    // The row will be picked up in a future sync once the date is corrected.
    if (hasIrregularDate_(startRaw) || hasIrregularDate_(shipRaw)) continue;

    // Date parsing (only reached if dates are empty or valid)
    var startDate = parseDate_(startRaw);
    var shipDate = parseDate_(shipRaw);

    // Data-level year filter
    if (earliestDate && (startDate || shipDate)) {
      var allTooOld = true;
      if (startDate && startDate >= earliestDate) allTooOld = false;
      if (shipDate && shipDate >= earliestDate) allTooOld = false;
      if (allTooOld) continue;
    }

    // Future date filter
    if (startDate || shipDate) {
      var allFuture = true;
      if (startDate && startDate <= today) allFuture = false;
      if (shipDate && shipDate <= today) allFuture = false;
      if (allFuture) continue;
    }

    // Build output row
    var outRow = [];
    for (var d = 0; d < SOURCE_COLUMNS.length; d++) {
      var colName = SOURCE_COLUMNS[d];
      outRow.push(colMap[colName] !== undefined ? row[colMap[colName]] : '');
    }
    outRow.push(sheetName);
    rows.push(outRow);
  }

  return rows;
}

// ================================================================
//  SECTION 3: SORT & FORMAT (Optimized for large datasets)
// ================================================================

function sortAndFormatSheet_(destSheet) {
  var lastRow = destSheet.getLastRow();
  if (lastRow <= 1) return;

  var totalCols = OUTPUT_HEADERS.length;
  var startCol = SOURCE_COLUMNS.indexOf('Start Date') + 1;
  var shipCol = SOURCE_COLUMNS.indexOf('Ship Date (t4i use only)') + 1;

  var dataRange = destSheet.getRange(2, 1, lastRow - 1, totalCols);
  dataRange.sort([
    {column: startCol, ascending: true},
    {column: shipCol, ascending: true}
  ]);

  var dateFormat = 'M/d/yyyy';
  destSheet.getRange(2, startCol, lastRow - 1, 1).setNumberFormat(dateFormat);
  destSheet.getRange(2, shipCol, lastRow - 1, 1).setNumberFormat(dateFormat);

  SpreadsheetApp.flush();
  Logger.log('Sort and format complete. Rows: ' + (lastRow - 1));
}

function runSortAndFormat() {
  var destSS = SpreadsheetApp.openById(DEST_SPREADSHEET_ID);
  var destSheet = destSS.getSheets()[0];
  Logger.log('Running sort and format...');
  sortAndFormatSheet_(destSheet);
  Logger.log('Done!');
}

// Remove rows with dates before EARLIEST_YEAR
function cleanupOldData() {
  var destSS = SpreadsheetApp.openById(DEST_SPREADSHEET_ID);
  var destSheet = destSS.getSheets()[0];
  var lastRow = destSheet.getLastRow();
  if (lastRow <= 1) return;

  var totalCols = OUTPUT_HEADERS.length;
  var startCol = SOURCE_COLUMNS.indexOf('Start Date');
  var shipCol = SOURCE_COLUMNS.indexOf('Ship Date (t4i use only)');
  var earliestDate = new Date(EARLIEST_YEAR, 0, 1);

  var data = destSheet.getRange(2, 1, lastRow - 1, totalCols).getValues();
  var kept = [];
  var removed = 0;

  for (var i = 0; i < data.length; i++) {
    var sd = parseDate_(data[i][startCol]);
    var shd = parseDate_(data[i][shipCol]);
    if (sd || shd) {
      var allOld = true;
      if (sd && sd >= earliestDate) allOld = false;
      if (shd && shd >= earliestDate) allOld = false;
      if (allOld) { removed++; continue; }
    }
    kept.push(data[i]);
  }

  destSheet.getRange(2, 1, lastRow - 1, totalCols).clearContent();
  if (kept.length > 0) {
    destSheet.getRange(2, 1, kept.length, totalCols).setValues(kept);
  }

  SpreadsheetApp.flush();
  Logger.log('Cleanup complete. Removed: ' + removed + ', Kept: ' + kept.length);
}

// NEW in v7: Remove rows with irregular (non-date) values in date fields
function cleanupIrregularDates() {
  var destSS = SpreadsheetApp.openById(DEST_SPREADSHEET_ID);
  var destSheet = destSS.getSheets()[0];
  var lastRow = destSheet.getLastRow();
  if (lastRow <= 1) return;

  var totalCols = OUTPUT_HEADERS.length;
  var startCol = SOURCE_COLUMNS.indexOf('Start Date');
  var shipCol = SOURCE_COLUMNS.indexOf('Ship Date (t4i use only)');

  var data = destSheet.getRange(2, 1, lastRow - 1, totalCols).getValues();
  var kept = [];
  var removed = 0;
  var examples = [];

  for (var i = 0; i < data.length; i++) {
    if (hasIrregularDate_(data[i][startCol]) || hasIrregularDate_(data[i][shipCol])) {
      removed++;
      if (examples.length < 10) {
        examples.push('Row ' + (i + 2) + ': Start="' + data[i][startCol] + '", Ship="' + data[i][shipCol] + '"');
      }
      continue;
    }
    kept.push(data[i]);
  }

  destSheet.getRange(2, 1, lastRow - 1, totalCols).clearContent();
  if (kept.length > 0) {
    destSheet.getRange(2, 1, kept.length, totalCols).setValues(kept);
  }

  SpreadsheetApp.flush();
  Logger.log('Irregular date cleanup complete. Removed: ' + removed + ', Kept: ' + kept.length);
  if (examples.length > 0) {
    Logger.log('Examples of removed rows:\n' + examples.join('\n'));
  }
}

// ================================================================
//  SECTION 4: FULL SYNC (Batch)
// ================================================================

function fullSync() {
  var props = PropertiesService.getScriptProperties();
  var sheetStart = parseInt(props.getProperty('FULL_SYNC_INDEX') || '0');

  if (props.getProperty('FULL_SYNC_PHASE') === 'complete') {
    Logger.log('Full sync already completed. Run resetFullSync() to restart.');
    return;
  }

  var sourceSS = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  var destSS = SpreadsheetApp.openById(DEST_SPREADSHEET_ID);
  var destSheet = destSS.getSheets()[0];
  var sheets = sourceSS.getSheets();
  var today = new Date();
  today.setHours(23, 59, 59, 999);

  if (sheetStart === 0) {
    destSheet.clearContents();
    destSheet.clearFormats();
    destSheet.getRange(1, 1, 1, OUTPUT_HEADERS.length).setValues([OUTPUT_HEADERS]);
    destSheet.getRange(1, 1, 1, OUTPUT_HEADERS.length).setFontWeight('bold');
    destSheet.setFrozenRows(1);
    Logger.log('=== Full Sync Started ===');
    Logger.log('Total source tabs: ' + sheets.length);
  } else {
    Logger.log('=== Full Sync Continuing from sheet index ' + sheetStart + ' ===');
  }

  var processed = 0;
  var batchRows = 0;
  var s = sheetStart;

  for (; s < sheets.length; s++) {
    var sheet = sheets[s];
    var sheetName = sheet.getName();

    if (EARLIEST_YEAR) {
      var tabYear = getTabYear_(sheetName);
      if (tabYear !== null && tabYear < EARLIEST_YEAR) continue;
    }

    var rows = processSheet_(sheet, sheetName, today);
    if (rows.length > 0) {
      var lastRow = destSheet.getLastRow();
      destSheet.getRange(lastRow + 1, 1, rows.length, OUTPUT_HEADERS.length).setValues(rows);
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

  if (s >= sheets.length) {
    props.setProperty('FULL_SYNC_PHASE', 'complete');
    props.deleteProperty('FULL_SYNC_INDEX');
    removeContinuationTriggers_();

    Logger.log('Sorting and formatting...');
    sortAndFormatSheet_(destSheet);

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
//  SECTION 5: INCREMENTAL DAILY SYNC
// ================================================================

function syncNewHireData() {
  var startTime = new Date();
  Logger.log('=== Incremental Sync Started: ' + startTime.toLocaleString() + ' ===');
  Logger.log('Lookback: ' + LOOKBACK_WEEKS + ' weeks');

  var sourceSS = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  var destSS = SpreadsheetApp.openById(DEST_SPREADSHEET_ID);
  var destSheet = destSS.getSheets()[0];
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

  // Step 2: Preserve historical rows
  var historicalRows = [];
  var lastRow = destSheet.getLastRow();
  if (lastRow > 1) {
    var existingData = destSheet.getRange(2, 1, lastRow - 1, OUTPUT_HEADERS.length).getValues();
    var sourceTabColIdx = OUTPUT_HEADERS.length - 1;
    for (var i = 0; i < existingData.length; i++) {
      var rowSourceTab = String(existingData[i][sourceTabColIdx]);
      if (recentTabNames.indexOf(rowSourceTab) === -1) {
        historicalRows.push(existingData[i]);
      }
    }
    Logger.log('Historical rows preserved: ' + historicalRows.length);
  }

  // Step 3: Process recent tabs
  var freshRows = [];
  for (var r = 0; r < recentSheets.length; r++) {
    var rows = processSheet_(recentSheets[r], recentTabNames[r], today);
    freshRows = freshRows.concat(rows);
    if (rows.length > 0) {
      Logger.log('  Refreshed: ' + recentTabNames[r] + ' → ' + rows.length + ' rows');
    }
  }
  Logger.log('Fresh rows: ' + freshRows.length);

  // Step 4: Combine and write
  var allRows = historicalRows.concat(freshRows);
  destSheet.clearContents();
  destSheet.clearFormats();
  destSheet.getRange(1, 1, 1, OUTPUT_HEADERS.length).setValues([OUTPUT_HEADERS]);
  destSheet.getRange(1, 1, 1, OUTPUT_HEADERS.length).setFontWeight('bold');
  destSheet.setFrozenRows(1);

  if (allRows.length > 0) {
    destSheet.getRange(2, 1, allRows.length, OUTPUT_HEADERS.length).setValues(allRows);
  }

  // Step 5: Sort and format
  sortAndFormatSheet_(destSheet);

  var elapsed = ((new Date() - startTime) / 1000).toFixed(1);
  Logger.log('=== Incremental Sync Complete: ' + allRows.length + ' rows, ' + elapsed + 's ===');
}

// ================================================================
//  SECTION 6: DISCOVERY TOOL
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
//  SECTION 7: HELPERS
// ================================================================

function parseDate_(val) {
  if (val === null || val === undefined || val === '') return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  var d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

// NEW in v7: Checks if a date field contains a non-empty, non-date value
// Returns true for values like "TBD", "N/A", "pending", "on hold", etc.
// Returns false for empty values (OK) and valid dates (OK)
function hasIrregularDate_(val) {
  // Empty/null/undefined → not irregular, just empty
  if (val === null || val === undefined || val === '') return false;

  // Already a valid Date object → not irregular
  if (val instanceof Date) return isNaN(val.getTime());

  // Number → Google Sheets stores dates as numbers internally → not irregular
  if (typeof val === 'number') return false;

  // String: try to parse as date
  var d = new Date(val);
  if (!isNaN(d.getTime())) return false; // Valid date string → not irregular

  // Non-empty string that is NOT a valid date → irregular (e.g., "TBD", "N/A")
  return true;
}

// ================================================================
//  SECTION 8: TRIGGER MANAGEMENT
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
