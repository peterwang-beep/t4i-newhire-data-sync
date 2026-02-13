// ================================================================
// T4i New Hire Data Sync Script — v5 (Incremental Sync)
// Syncs data from 'T4i New Hire List' to 'newHireAutoSyncData'
//
// v5 Changes:
//   - NEW: fullSync() — batch processes all historical tabs automatically
//   - NEW: syncNewHireData() — incremental daily sync (recent tabs only)
//   - NEW: Tab date parsing for smart lookback window
//   - Historical data preserved between daily syncs
//   - Batch auto-continuation with 1-minute triggers
// ================================================================

// ============ Configuration ============
var SOURCE_SPREADSHEET_ID = '1cAlDPmgQ3mkmC5l5Jw-h6HlDre1RqIenAHVpoulZNwU';
var DEST_SPREADSHEET_ID   = '1mt2tOVCtef0y2vP5_Y53E-wpU2MZN_tKbbfDk7bWxbw';

// Year filter — Only process tabs from this year onward
var EARLIEST_YEAR = 2021;

// Batch processing — tabs per execution in fullSync
var BATCH_SIZE = 150;

// Incremental sync — how many weeks back to look for daily sync
// Tabs older than this are considered "frozen" and not re-processed
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

// Output = source columns + Source Tab
var OUTPUT_HEADERS = SOURCE_COLUMNS.concat(['Source Tab']);

// Key fields — at least one must be non-empty per row
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
      .addItem('Discover Tabs', 'discoverTabs')
      .addSeparator()
      .addItem('Setup Daily Triggers', 'setupDailyTriggers')
      .addItem('Remove Daily Triggers', 'removeDailyTriggers')
      .addItem('Reset Full Sync State', 'resetFullSync')
      .addToUi();
  } catch (e) {
    // Menu not available in standalone mode
  }
}

// ================================================================
//  SECTION 1: TAB NAME PARSERS
// ================================================================

// Extract year from tab name (e.g., "Week of Jan 11 21" → 2021)
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

// Extract approximate date from tab name for lookback filtering
function getTabApproxDate_(tabName) {
  var year = getTabYear_(tabName);
  if (!year) return null;

  var monthMap = [
    'january', 'february', 'march', 'april', 'may', 'june',
    'july', 'august', 'september', 'october', 'november', 'december',
    'jan', 'feb', 'mar', 'apr', 'may', 'jun',
    'jul', 'aug', 'sep', 'oct', 'nov', 'dec'
  ];
  var monthNum = [
    0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11,
    0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11
  ];

  var lower = tabName.toLowerCase();
  for (var i = 0; i < monthMap.length; i++) {
    if (lower.indexOf(monthMap[i]) !== -1) {
      return new Date(year, monthNum[i], 1);
    }
  }
  return new Date(year, 0, 1); // Default to January
}

// Check if a tab is within the lookback window
function isRecentTab_(tabName, cutoffDate) {
  var tabDate = getTabApproxDate_(tabName);
  if (!tabDate) return true; // Can't determine date → include to be safe
  return tabDate >= cutoffDate;
}

// ================================================================
//  SECTION 2: CORE SHEET PROCESSOR
// ================================================================

// Process a single sheet and return array of output rows
function processSheet_(sheet, sheetName, today) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 3 || lastCol === 0) return [];

  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  // Row 2 (index 1) = headers
  var headers = [];
  for (var h = 0; h < data[1].length; h++) {
    headers.push(String(data[1][h]).trim());
  }

  // Build column mapping
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

    // Date filter: skip if all dates are future
    var startDate = colMap['Start Date'] !== undefined ? parseDate_(row[colMap['Start Date']]) : null;
    var shipDate = colMap['Ship Date (t4i use only)'] !== undefined ? parseDate_(row[colMap['Ship Date (t4i use only)']]) : null;

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
    outRow.push(sheetName); // Source Tab
    rows.push(outRow);
  }

  return rows;
}

// ================================================================
//  SECTION 3: SORT, FORMAT, AND DUPLICATE DETECTION
// ================================================================

function sortAndFormatSheet_(destSheet) {
  var lastRow = destSheet.getLastRow();
  if (lastRow <= 1) return;

  var totalCols = OUTPUT_HEADERS.length;
  var startIdx = SOURCE_COLUMNS.indexOf('Start Date');
  var shipIdx = SOURCE_COLUMNS.indexOf('Ship Date (t4i use only)');
  var data = destSheet.getRange(2, 1, lastRow - 1, totalCols).getValues();

  // Sort: Start Date (primary) → Ship Date (secondary)
  data.sort(function(a, b) {
    var aStart = parseDate_(a[startIdx]);
    var bStart = parseDate_(b[startIdx]);
    var aShip = parseDate_(a[shipIdx]);
    var bShip = parseDate_(b[shipIdx]);

    if (aStart && bStart) {
      var diff = aStart.getTime() - bStart.getTime();
      if (diff !== 0) return diff;
    } else if (aStart && !bStart) return -1;
    else if (!aStart && bStart) return 1;

    if (aShip && bShip) return aShip.getTime() - bShip.getTime();
    if (aShip && !bShip) return -1;
    if (!aShip && bShip) return 1;

    return 0;
  });

  destSheet.getRange(2, 1, data.length, totalCols).setValues(data);

  // Format date columns
  var dateFormat = 'M/d/yyyy';
  destSheet.getRange(2, startIdx + 1, data.length, 1).setNumberFormat(dateFormat);
  destSheet.getRange(2, shipIdx + 1, data.length, 1).setNumberFormat(dateFormat);

  // Auto-resize
  for (var col = 1; col <= totalCols; col++) {
    destSheet.autoResizeColumn(col);
  }

  SpreadsheetApp.flush();

  // Duplicate detection
  logDuplicates_(data);
}

function logDuplicates_(data) {
  var fullNameIdx = SOURCE_COLUMNS.indexOf('Full name (do not fill)');
  var startIdx = SOURCE_COLUMNS.indexOf('Start Date');
  var hireIdx = SOURCE_COLUMNS.indexOf('Hire Type');
  var dupTracker = {};

  for (var i = 0; i < data.length; i++) {
    var fn = String(data[i][fullNameIdx] || '').trim();
    var sd = parseDate_(data[i][startIdx]);
    var ht = String(data[i][hireIdx] || '').trim();
    var key = fn + '|' + (sd ? sd.toDateString() : '') + '|' + ht;
    if (key !== '||') {
      dupTracker[key] = (dupTracker[key] || 0) + 1;
    }
  }

  var dupCount = 0;
  for (var k in dupTracker) {
    if (dupTracker[k] > 1) dupCount++;
  }
  Logger.log('Duplicate groups detected: ' + dupCount);
}

// ================================================================
//  SECTION 4: FULL SYNC (Batch — Run Once)
// ================================================================

// Run this once to load all historical data. It auto-continues in batches.
function fullSync() {
  var props = PropertiesService.getScriptProperties();
  var sheetStart = parseInt(props.getProperty('FULL_SYNC_INDEX') || '0');

  // Check if already completed
  if (props.getProperty('FULL_SYNC_PHASE') === 'complete') {
    Logger.log('Full sync already completed previously.');
    Logger.log('To re-run, first execute resetFullSync(), then run fullSync() again.');
    return;
  }

  var sourceSS = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  var destSS = SpreadsheetApp.openById(DEST_SPREADSHEET_ID);
  var destSheet = destSS.getSheets()[0];
  var sheets = sourceSS.getSheets();
  var today = new Date();
  today.setHours(23, 59, 59, 999);

  // First batch: clear sheet and write headers
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

    // Year filter
    if (EARLIEST_YEAR) {
      var tabYear = getTabYear_(sheetName);
      if (tabYear !== null && tabYear < EARLIEST_YEAR) continue;
    }

    // Process this sheet
    var rows = processSheet_(sheet, sheetName, today);
    if (rows.length > 0) {
      var lastRow = destSheet.getLastRow();
      destSheet.getRange(lastRow + 1, 1, rows.length, OUTPUT_HEADERS.length).setValues(rows);
      batchRows += rows.length;
    }

    processed++;
    Logger.log('  [' + processed + '/' + BATCH_SIZE + '] ' + sheetName + ' → ' + rows.length + ' rows');

    if (processed >= BATCH_SIZE) {
      s++; // Next start point
      break;
    }
  }

  SpreadsheetApp.flush();
  Logger.log('Batch result: ' + processed + ' tabs processed, ' + batchRows + ' rows added.');

  if (s >= sheets.length) {
    // ---- All batches complete ----
    props.setProperty('FULL_SYNC_PHASE', 'complete');
    props.deleteProperty('FULL_SYNC_INDEX');
    removeContinuationTriggers_();

    Logger.log('Sorting and formatting...');
    sortAndFormatSheet_(destSheet);

    var totalRows = destSheet.getLastRow() - 1;
    Logger.log('=== Full Sync Complete! Total rows: ' + totalRows + ' ===');
  } else {
    // ---- More batches needed ----
    props.setProperty('FULL_SYNC_INDEX', String(s));

    // Auto-continue in ~1 minute
    ScriptApp.newTrigger('fullSync')
      .timeBased()
      .after(60 * 1000)
      .create();

    Logger.log('More tabs remain. Auto-continuing in ~1 minute...');
  }
}

// Reset full sync state so it can be run again from scratch
function resetFullSync() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('FULL_SYNC_INDEX');
  props.deleteProperty('FULL_SYNC_PHASE');
  removeContinuationTriggers_();
  Logger.log('Full sync state reset. You can now run fullSync() again.');
}

// Remove only fullSync continuation triggers (not daily triggers)
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

// Daily sync: only refreshes tabs from the last LOOKBACK_WEEKS weeks.
// Historical data already in the destination sheet is preserved.
function syncNewHireData() {
  var startTime = new Date();
  Logger.log('=== Incremental Sync Started: ' + startTime.toLocaleString() + ' ===');
  Logger.log('Lookback window: ' + LOOKBACK_WEEKS + ' weeks');

  var sourceSS = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  var destSS = SpreadsheetApp.openById(DEST_SPREADSHEET_ID);
  var destSheet = destSS.getSheets()[0];
  var sheets = sourceSS.getSheets();
  var today = new Date();
  today.setHours(23, 59, 59, 999);

  // Calculate cutoff date
  var cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - (LOOKBACK_WEEKS * 7));
  Logger.log('Cutoff date: ' + cutoffDate.toDateString());

  // Step 1: Identify recent tabs to refresh
  var recentTabNames = [];
  var recentSheets = [];

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var name = sheet.getName();

    // Year filter
    if (EARLIEST_YEAR) {
      var tabYear = getTabYear_(name);
      if (tabYear !== null && tabYear < EARLIEST_YEAR) continue;
    }

    // Lookback filter
    if (isRecentTab_(name, cutoffDate)) {
      recentTabNames.push(name);
      recentSheets.push(sheet);
    }
  }

  Logger.log('Recent tabs to refresh: ' + recentTabNames.length);

  // Step 2: Read existing data, preserve rows NOT from recent tabs
  var historicalRows = [];
  var lastRow = destSheet.getLastRow();

  if (lastRow > 1) {
    var existingData = destSheet.getRange(2, 1, lastRow - 1, OUTPUT_HEADERS.length).getValues();
    var sourceTabColIdx = OUTPUT_HEADERS.length - 1; // Last column = Source Tab

    for (var i = 0; i < existingData.length; i++) {
      var rowSourceTab = String(existingData[i][sourceTabColIdx]);
      // Keep row only if its Source Tab is NOT being refreshed
      if (recentTabNames.indexOf(rowSourceTab) === -1) {
        historicalRows.push(existingData[i]);
      }
    }

    var removed = existingData.length - historicalRows.length;
    Logger.log('Historical rows preserved: ' + historicalRows.length + ', rows to refresh: ' + removed);
  }

  // Step 3: Process recent tabs for fresh data
  var freshRows = [];
  for (var r = 0; r < recentSheets.length; r++) {
    var rows = processSheet_(recentSheets[r], recentTabNames[r], today);
    freshRows = freshRows.concat(rows);
    if (rows.length > 0) {
      Logger.log('  Refreshed: ' + recentTabNames[r] + ' → ' + rows.length + ' rows');
    }
  }

  Logger.log('Fresh rows from recent tabs: ' + freshRows.length);

  // Step 4: Combine historical + fresh data
  var allRows = historicalRows.concat(freshRows);

  // Step 5: Write to destination
  destSheet.clearContents();
  destSheet.clearFormats();
  destSheet.getRange(1, 1, 1, OUTPUT_HEADERS.length).setValues([OUTPUT_HEADERS]);
  destSheet.getRange(1, 1, 1, OUTPUT_HEADERS.length).setFontWeight('bold');
  destSheet.setFrozenRows(1);

  if (allRows.length > 0) {
    destSheet.getRange(2, 1, allRows.length, OUTPUT_HEADERS.length).setValues(allRows);
  }

  // Step 6: Sort, format, detect duplicates
  sortAndFormatSheet_(destSheet);

  var elapsed = ((new Date() - startTime) / 1000).toFixed(1);
  Logger.log('=== Incremental Sync Complete ===');
  Logger.log('Total rows: ' + allRows.length + ' | Time: ' + elapsed + 's');
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
  Logger.log('');

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var name = sheet.getName();

    if (EARLIEST_YEAR) {
      var tabYear = getTabYear_(name);
      if (tabYear !== null && tabYear < EARLIEST_YEAR) {
        yearSkipCount++;
        continue;
      }
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol === 0) {
      skipCount++;
      continue;
    }

    var headerRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
    var headers = [];
    for (var i = 0; i < headerRow.length; i++) {
      headers.push(String(headerRow[i]).trim());
    }

    var matched = [];
    var missing = [];
    for (var j = 0; j < SOURCE_COLUMNS.length; j++) {
      if (headers.indexOf(SOURCE_COLUMNS[j]) !== -1) {
        matched.push(SOURCE_COLUMNS[j]);
      } else {
        missing.push(SOURCE_COLUMNS[j]);
      }
    }

    if (matched.length > 0) {
      matchCount++;
      var msg = '[MATCH] ' + name + ' — ' + matched.length + '/' + SOURCE_COLUMNS.length;
      if (missing.length > 0 && missing.length <= 5) {
        msg += ' | Missing: ' + missing.join(', ');
      }
      Logger.log(msg);
    } else {
      skipCount++;
      Logger.log('[SKIP] ' + name);
    }
  }

  Logger.log('');
  Logger.log('=== Summary ===');
  Logger.log('Tabs to sync: ' + matchCount);
  Logger.log('Tabs skipped (no match): ' + skipCount);
  Logger.log('Tabs skipped (before ' + EARLIEST_YEAR + '): ' + yearSkipCount);
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

// ================================================================
//  SECTION 8: TRIGGER MANAGEMENT
// ================================================================

function setupDailyTriggers() {
  removeDailyTriggers();

  ScriptApp.newTrigger('syncNewHireData')
    .timeBased()
    .atHour(6)
    .everyDays(1)
    .inTimezone('America/New_York')
    .create();

  ScriptApp.newTrigger('syncNewHireData')
    .timeBased()
    .atHour(20)
    .everyDays(1)
    .inTimezone('America/New_York')
    .create();

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
