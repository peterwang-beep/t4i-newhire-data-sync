// ================================================================
// T4i New Hire Data Sync Script — v8 (Data Integrity Fix)
// Syncs data from 'T4i New Hire List' to 'newHireAutoSyncData'
//
// v8 Changes (6 fixes from comprehensive audit):
//   BUG-1 [CRITICAL] Fixed Source Tab auto-formatting that caused infinite
//         duplication every sync cycle. Now sets column format to plain text
//         and uses robust comparison with normalizeSourceTab_().
//   BUG-2 [MODERATE] Date columns now normalized to midnight — no more
//         hidden time components causing sort inconsistencies.
//   BUG-3 [MODERATE] All write operations now use shared writeDestData_()
//         which enforces correct column formats every time.
//   BUG-4 [MINOR]   hasIrregularDate_() now rejects dates outside 2000-2030.
//   BUG-5 [MINOR]   sortAndFormatSheet_ now also sets Source Tab to plain text.
//   NEW:  cleanupDuplicates() removes accumulated duplicate rows.
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

// Column indices (0-based for arrays, 1-based noted where needed)
var START_DATE_IDX = 0;
var SHIP_DATE_IDX = 1;
var FIRST_NAME_IDX = 2;
var LAST_NAME_IDX = 3;
var HIRE_TYPE_IDX = 7;
var SOURCE_TAB_IDX = OUTPUT_HEADERS.length - 1; // 18 (0-based), 19 (1-based)

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

    // Irregular date check (v7+): skip rows with non-date text like "TBD"
    if (hasIrregularDate_(startRaw) || hasIrregularDate_(shipRaw)) continue;

    // Date parsing
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

    // Build output row with date normalization (BUG-2 fix)
    var outRow = [];
    for (var d = 0; d < SOURCE_COLUMNS.length; d++) {
      var colName = SOURCE_COLUMNS[d];
      if (colMap[colName] !== undefined) {
        var cellVal = row[colMap[colName]];
        // Normalize dates to midnight to strip time components
        if ((colName === 'Start Date' || colName === 'Ship Date (t4i use only)') && cellVal instanceof Date && !isNaN(cellVal.getTime())) {
          cellVal = new Date(cellVal.getFullYear(), cellVal.getMonth(), cellVal.getDate());
        }
        outRow.push(cellVal);
      } else {
        outRow.push('');
      }
    }
    outRow.push(sheetName); // Source Tab as plain string
    rows.push(outRow);
  }

  return rows;
}

// ================================================================
//  SECTION 3: SHARED WRITE HELPER (BUG-1/BUG-3 fix)
// ================================================================

// Centralized function for all destination sheet writes.
// Ensures correct column formats every time: dates as M/d/yyyy, Source Tab as plain text.
function writeDestData_(destSheet, dataRows) {
  var totalCols = OUTPUT_HEADERS.length;

  destSheet.clearContents();
  destSheet.clearFormats();

  // Headers
  destSheet.getRange(1, 1, 1, totalCols).setValues([OUTPUT_HEADERS]);
  destSheet.getRange(1, 1, 1, totalCols).setFontWeight('bold');
  destSheet.setFrozenRows(1);

  if (dataRows.length > 0) {
    destSheet.getRange(2, 1, dataRows.length, totalCols).setValues(dataRows);
  }

  // Enforce column formats (critical for preventing auto-formatting bugs)
  var formatRows = Math.max(dataRows.length, 1);
  destSheet.getRange(2, START_DATE_IDX + 1, formatRows, 1).setNumberFormat('M/d/yyyy');
  destSheet.getRange(2, SHIP_DATE_IDX + 1, formatRows, 1).setNumberFormat('M/d/yyyy');
  destSheet.getRange(2, SOURCE_TAB_IDX + 1, formatRows, 1).setNumberFormat('@'); // Plain text

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

  // Built-in sort
  var dataRange = destSheet.getRange(2, 1, lastRow - 1, totalCols);
  dataRange.sort([
    {column: startCol, ascending: true},
    {column: shipCol, ascending: true}
  ]);

  // Re-enforce formats after sort (BUG-5 fix: includes Source Tab)
  destSheet.getRange(2, startCol, lastRow - 1, 1).setNumberFormat('M/d/yyyy');
  destSheet.getRange(2, shipCol, lastRow - 1, 1).setNumberFormat('M/d/yyyy');
  destSheet.getRange(2, SOURCE_TAB_IDX + 1, lastRow - 1, 1).setNumberFormat('@');

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

// ================================================================
//  SECTION 5: CLEANUP FUNCTIONS
// ================================================================

// Remove rows with dates before EARLIEST_YEAR
function cleanupOldData() {
  var destSS = SpreadsheetApp.openById(DEST_SPREADSHEET_ID);
  var destSheet = destSS.getSheets()[0];
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

// Remove rows with irregular (non-date) values in date fields
function cleanupIrregularDates() {
  var destSS = SpreadsheetApp.openById(DEST_SPREADSHEET_ID);
  var destSheet = destSS.getSheets()[0];
  var lastRow = destSheet.getLastRow();
  if (lastRow <= 1) return;

  var totalCols = OUTPUT_HEADERS.length;
  var data = destSheet.getRange(2, 1, lastRow - 1, totalCols).getValues();
  var kept = [];
  var removed = 0;
  var examples = [];

  for (var i = 0; i < data.length; i++) {
    if (hasIrregularDate_(data[i][START_DATE_IDX]) || hasIrregularDate_(data[i][SHIP_DATE_IDX])) {
      removed++;
      if (examples.length < 10) {
        examples.push('Start="' + data[i][START_DATE_IDX] + '", Ship="' + data[i][SHIP_DATE_IDX] + '"');
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

// NEW in v8: Remove duplicate rows accumulated from the Source Tab formatting bug.
// Dedup key: First Name + Last Name + Start Date + Hire Type + Source Tab
// Keeps only the first occurrence of each unique combination.
function cleanupDuplicates() {
  var destSS = SpreadsheetApp.openById(DEST_SPREADSHEET_ID);
  var destSheet = destSS.getSheets()[0];
  var lastRow = destSheet.getLastRow();
  if (lastRow <= 1) return;

  var totalCols = OUTPUT_HEADERS.length;
  var data = destSheet.getRange(2, 1, lastRow - 1, totalCols).getValues();
  var seen = {};
  var kept = [];
  var removed = 0;

  for (var i = 0; i < data.length; i++) {
    var firstName = String(data[i][FIRST_NAME_IDX] || '').trim();
    var lastName = String(data[i][LAST_NAME_IDX] || '').trim();
    var startDate = parseDate_(data[i][START_DATE_IDX]);
    var startStr = startDate ? startDate.toDateString() : '';
    var hireType = String(data[i][HIRE_TYPE_IDX] || '').trim();
    var sourceTab = normalizeSourceTab_(data[i][SOURCE_TAB_IDX]);

    var key = firstName + '|' + lastName + '|' + startStr + '|' + hireType + '|' + sourceTab;

    if (seen[key]) {
      removed++;
      continue;
    }
    seen[key] = true;
    kept.push(data[i]);
  }

  writeDestData_(destSheet, kept);
  sortAndFormatSheet_(destSheet);
  Logger.log('Duplicate cleanup: Removed ' + removed + ', Kept ' + kept.length);
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
    // Set Source Tab column to plain text from the start (BUG-1 prevention)
    destSheet.getRange(2, SOURCE_TAB_IDX + 1, 1, 1).setNumberFormat('@');
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
      // BUG-1 fix: set Source Tab format to plain text for newly written rows
      destSheet.getRange(lastRow + 1, SOURCE_TAB_IDX + 1, rows.length, 1).setNumberFormat('@');
      // Also enforce date formats for the batch
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
//  SECTION 7: INCREMENTAL DAILY SYNC
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

  // BUG-1 fix: Build a lookup set that matches both tab name AND its date equivalent.
  // This handles existing data where Google Sheets auto-formatted "February 16 2026" → Date(2/16/2026).
  var recentTabSet = {};
  for (var t = 0; t < recentTabNames.length; t++) {
    recentTabSet[recentTabNames[t]] = true;
    // Also add the date-formatted version so we can match auto-formatted Source Tab values
    var approxDate = getTabApproxDate_(recentTabNames[t]);
    if (approxDate) {
      recentTabSet[(approxDate.getMonth() + 1) + '/' + approxDate.getDate() + '/' + approxDate.getFullYear()] = true;
    }
  }

  // Step 2: Preserve historical rows (using robust Source Tab matching)
  var historicalRows = [];
  var lastRow = destSheet.getLastRow();
  if (lastRow > 1) {
    var existingData = destSheet.getRange(2, 1, lastRow - 1, OUTPUT_HEADERS.length).getValues();
    for (var i = 0; i < existingData.length; i++) {
      var rowSourceTab = normalizeSourceTab_(existingData[i][SOURCE_TAB_IDX]);
      if (!recentTabSet[rowSourceTab]) {
        historicalRows.push(existingData[i]);
      }
    }
    var removedCount = existingData.length - historicalRows.length;
    Logger.log('Historical rows preserved: ' + historicalRows.length + ' (removed ' + removedCount + ' for refresh)');
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

  // Step 4: Combine and write (using shared helper with format enforcement)
  var allRows = historicalRows.concat(freshRows);
  writeDestData_(destSheet, allRows);

  // Step 5: Sort and format
  sortAndFormatSheet_(destSheet);

  var elapsed = ((new Date() - startTime) / 1000).toFixed(1);
  Logger.log('=== Incremental Sync Complete: ' + allRows.length + ' rows, ' + elapsed + 's ===');
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

// v7+: Checks if a date field contains non-date text (e.g., "TBD", "N/A")
// v8 improvement: also rejects dates outside 2000-2030 range (BUG-4 fix)
function hasIrregularDate_(val) {
  if (val === null || val === undefined || val === '') return false;
  if (val instanceof Date) return isNaN(val.getTime());
  if (typeof val === 'number') {
    // Validate number-based dates: Sheets serial 36526 = Jan 1 2000, 47848 = Dec 31 2030
    if (val < 36526 || val > 47848) return true;
    return false;
  }
  var d = new Date(val);
  if (!isNaN(d.getTime())) {
    var yr = d.getFullYear();
    if (yr < 2000 || yr > 2030) return true; // Suspicious date range
    return false;
  }
  return true; // Non-date string
}

// v8 NEW: Normalizes a Source Tab cell value for comparison.
// Handles the case where Google Sheets auto-formatted a tab name as a Date object.
function normalizeSourceTab_(val) {
  if (val === null || val === undefined || val === '') return '';
  if (val instanceof Date && !isNaN(val.getTime())) {
    // Convert auto-formatted date back to M/d/yyyy string for matching
    return (val.getMonth() + 1) + '/' + val.getDate() + '/' + val.getFullYear();
  }
  return String(val).trim();
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
