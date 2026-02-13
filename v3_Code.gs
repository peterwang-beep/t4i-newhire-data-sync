// ================================================================
// T4i New Hire Data Sync Script — v3
// Syncs data from 'T4i New Hire List' to 'newHireAutoSyncData'
//
// v3 Changes:
//   - All comments converted to English
//   - Start Date moved to column 1, Ship Date to column 2
//   - Added Source Tab column (last column) to track data origin
//   - Duplicate detection based on Full Name + Start Date + Hire Type
//   - QlikSense-friendly output: clean single-sheet format
// ================================================================

// ============ Configuration ============
var SOURCE_SPREADSHEET_ID = '1cAlDPmgQ3mkmC5l5Jw-h6HlDre1RqIenAHVpoulZNwU';
var DEST_SPREADSHEET_ID   = '1mt2tOVCtef0y2vP5_Y53E-wpU2MZN_tKbbfDk7bWxbw';

// Columns to extract from source (in custom display order)
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

// Output headers = source columns + script-generated columns
var OUTPUT_HEADERS = SOURCE_COLUMNS.concat(['Source Tab']);

// Key fields — at least one must be non-empty, otherwise skip entire row
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
  SpreadsheetApp.getUi()
    .createMenu('T4i Sync')
    .addItem('Sync Now', 'syncNewHireData')
    .addItem('Discover Tabs', 'discoverTabs')
    .addSeparator()
    .addItem('Setup Daily Triggers (6am & 8pm ET)', 'setupDailyTriggers')
    .addItem('Remove All Triggers', 'removeTriggers')
    .addToUi();
}

// ============ Tab Discovery Tool ============
// Scans all tabs in the source spreadsheet and reports which ones contain target columns.
// How to use: Click "Discover Tabs" in the T4i Sync menu, or select this function and click Run.
// When to run: After tabs are added/removed/renamed in the source spreadsheet.
function discoverTabs() {
  var ss = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  var sheets = ss.getSheets();
  var results = [];

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var name = sheet.getName();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    if (lastRow < 2 || lastCol === 0) {
      results.push('[SKIP] ' + name + ' — insufficient rows or columns');
      continue;
    }

    // Read row 2 only (column header row)
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
      var msg = '[MATCH] ' + name + ' — ' + matched.length + '/' + SOURCE_COLUMNS.length + ' columns';
      if (missing.length > 0 && missing.length <= 5) {
        msg += '\n        Missing: ' + missing.join(', ');
      }
      results.push(msg);
    } else {
      results.push('[SKIP] ' + name + ' — 0 matching columns');
    }
  }

  var msg = results.join('\n');
  Logger.log(msg);
  SpreadsheetApp.getUi().alert('Tab Discovery Results', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============ Main Sync Function ============
function syncNewHireData() {
  var startTime = new Date();
  var sourceSS = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  var destSS   = SpreadsheetApp.openById(DEST_SPREADSHEET_ID);
  var destSheet = destSS.getSheets()[0]; // First sheet in destination

  // Today's date (end of day as cutoff)
  var today = new Date();
  today.setHours(23, 59, 59, 999);

  var allRows = [];
  var sheets = sourceSS.getSheets();
  var tabsProcessed = 0;

  // Duplicate tracker: key = "FullName|StartDate|HireType"
  var dupTracker = {};

  // ---------- Iterate through each tab ----------
  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var sheetName = sheet.getName();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    // Need at least header row (row 2) + 1 data row (row 3)
    if (lastRow < 3 || lastCol === 0) continue;

    // Batch read all data
    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

    // Row 2 (index 1) contains column headers
    var headers = [];
    for (var h = 0; h < data[1].length; h++) {
      headers.push(String(data[1][h]).trim());
    }

    // Build source column -> column index mapping
    var colMap = {};
    var matchCount = 0;
    for (var c = 0; c < SOURCE_COLUMNS.length; c++) {
      var col = SOURCE_COLUMNS[c];
      var idx = headers.indexOf(col);
      if (idx !== -1) {
        colMap[col] = idx;
        matchCount++;
      }
    }

    // Skip tabs with no matching columns
    if (matchCount === 0) continue;
    tabsProcessed++;

    // ---------- Iterate data rows (row 3 onward = index 2+) ----------
    for (var i = 2; i < data.length; i++) {
      var row = data[i];

      // ---- Key field check: at least one must be non-empty ----
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

      // ---- Date filter: only sync up to today ----
      var startRaw = colMap['Start Date'] !== undefined ? row[colMap['Start Date']] : null;
      var shipRaw  = colMap['Ship Date (t4i use only)'] !== undefined ? row[colMap['Ship Date (t4i use only)']] : null;

      var startDate = parseDate_(startRaw);
      var shipDate  = parseDate_(shipRaw);

      var hasStart = (startDate !== null);
      var hasShip  = (shipDate !== null);

      // If any date exists, require at least one to be <= today; skip if all are future
      if (hasStart || hasShip) {
        var allFuture = true;
        if (hasStart && startDate <= today) allFuture = false;
        if (hasShip  && shipDate  <= today) allFuture = false;
        if (allFuture) continue;
      }

      // ---- Build output row in custom column order ----
      var outRow = [];
      for (var d = 0; d < SOURCE_COLUMNS.length; d++) {
        var colName = SOURCE_COLUMNS[d];
        if (colMap[colName] !== undefined) {
          outRow.push(row[colMap[colName]]);
        } else {
          outRow.push('');
        }
      }
      // Append Source Tab column
      outRow.push(sheetName);

      // ---- Duplicate detection ----
      var fullNameVal = colMap['Full name (do not fill)'] !== undefined ? String(row[colMap['Full name (do not fill)']]).trim() : '';
      var startDateStr = startDate ? startDate.toDateString() : '';
      var hireTypeVal = colMap['Hire Type'] !== undefined ? String(row[colMap['Hire Type']]).trim() : '';
      var dupKey = fullNameVal + '|' + startDateStr + '|' + hireTypeVal;

      // Only track combinations with actual values
      if (dupKey !== '||') {
        if (!dupTracker[dupKey]) {
          dupTracker[dupKey] = 0;
        }
        dupTracker[dupKey]++;
      }

      allRows.push(outRow);
    }
  }

  // ---------- Sort: Start Date (primary) → Ship Date (secondary), earliest first ----------
  var startIdx = SOURCE_COLUMNS.indexOf('Start Date');
  var shipIdx  = SOURCE_COLUMNS.indexOf('Ship Date (t4i use only)');

  allRows.sort(function(a, b) {
    var aStart = parseDate_(a[startIdx]);
    var bStart = parseDate_(b[startIdx]);
    var aShip  = parseDate_(a[shipIdx]);
    var bShip  = parseDate_(b[shipIdx]);

    // Primary sort: Start Date (rows with dates come first, earliest to latest)
    if (aStart && bStart) {
      var diff = aStart.getTime() - bStart.getTime();
      if (diff !== 0) return diff;
    } else if (aStart && !bStart) {
      return -1;
    } else if (!aStart && bStart) {
      return 1;
    }

    // Secondary sort: Ship Date
    if (aShip && bShip) {
      return aShip.getTime() - bShip.getTime();
    } else if (aShip && !bShip) {
      return -1;
    } else if (!aShip && bShip) {
      return 1;
    }

    return 0;
  });

  // ---------- Write to destination sheet ----------
  destSheet.clearContents();
  destSheet.clearFormats();

  var totalCols = OUTPUT_HEADERS.length; // 19 columns (18 source + Source Tab)

  // Write headers (row 1)
  destSheet.getRange(1, 1, 1, totalCols).setValues([OUTPUT_HEADERS]);
  destSheet.getRange(1, 1, 1, totalCols).setFontWeight('bold');
  destSheet.setFrozenRows(1); // Freeze header row for easier QlikSense loading

  // Write data rows
  if (allRows.length > 0) {
    destSheet.getRange(2, 1, allRows.length, totalCols).setValues(allRows);

    // Format date columns as date type (M/d/yyyy)
    var dateFormat = 'M/d/yyyy';
    destSheet.getRange(2, startIdx + 1, allRows.length, 1).setNumberFormat(dateFormat);
    destSheet.getRange(2, shipIdx + 1, allRows.length, 1).setNumberFormat(dateFormat);
  }

  // ---------- Auto-resize columns ----------
  for (var col = 1; col <= totalCols; col++) {
    destSheet.autoResizeColumn(col);
  }

  SpreadsheetApp.flush();

  // ---------- Count duplicates ----------
  var dupCount = 0;
  var dupDetails = [];
  for (var key in dupTracker) {
    if (dupTracker[key] > 1) {
      dupCount++;
      dupDetails.push(key.replace(/\|/g, ' / ') + ' (x' + dupTracker[key] + ')');
    }
  }

  // ---------- Sync summary log ----------
  var elapsed = ((new Date() - startTime) / 1000).toFixed(1);
  var summary = 'Sync complete!\n' +
    '- Rows synced: ' + allRows.length + '\n' +
    '- Tabs processed: ' + tabsProcessed + '/' + sheets.length + '\n' +
    '- Duplicate groups found: ' + dupCount + '\n' +
    '- Time elapsed: ' + elapsed + 's';

  Logger.log(summary);

  if (dupCount > 0) {
    Logger.log('Duplicate details:\n' + dupDetails.join('\n'));
  }
}

// ============ Date Parser Helper ============
function parseDate_(val) {
  if (val === null || val === undefined || val === '') return null;
  if (val instanceof Date) {
    return isNaN(val.getTime()) ? null : val;
  }
  var d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

// ============ Trigger Management ============
function setupDailyTriggers() {
  // Remove existing triggers first
  removeTriggers();

  // Daily at 6:00 AM Eastern Time
  ScriptApp.newTrigger('syncNewHireData')
    .timeBased()
    .atHour(6)
    .everyDays(1)
    .inTimezone('America/New_York')
    .create();

  // Daily at 8:00 PM Eastern Time (20:00)
  ScriptApp.newTrigger('syncNewHireData')
    .timeBased()
    .atHour(20)
    .everyDays(1)
    .inTimezone('America/New_York')
    .create();

  Logger.log('Triggers created: 6am ET & 8pm ET daily.');
  SpreadsheetApp.getUi().alert(
    'Daily triggers set successfully!\n\n' +
    '- Sync 1: 6:00 AM Eastern Time\n' +
    '- Sync 2: 8:00 PM Eastern Time\n\n' +
    'Note: Google triggers may run within a +/- 15 min window.'
  );
}

function removeTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  Logger.log('All triggers removed.');
}
