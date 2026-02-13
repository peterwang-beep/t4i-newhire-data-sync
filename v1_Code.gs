// ================================================================
// T4i New Hire Data Sync Script
// 从 'T4i New Hire List' 同步数据到 'newHireAutoSyncData'
// ================================================================

// ============ 配置区 ============
var SOURCE_SPREADSHEET_ID = '1cAlDPmgQ3mkmC5l5Jw-h6HlDre1RqIenAHVpoulZNwU';
var DEST_SPREADSHEET_ID   = '1mt2tOVCtef0y2vP5_Y53E-wpU2MZN_tKbbfDk7bWxbw';

// 需要保留的列（按自定义显示顺序排列）
var DESIRED_COLUMNS = [
  'Ship Date (t4i use only)',
  'First Name',
  'Last Name',
  'Full name (do not fill)',
  'Username',
  'SNOW Ticket',
  'Start Date',
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

// 关键字段 —— 至少一项必须非空，否则整行跳过
var KEY_FIELDS = [
  'First Name',
  'Last Name',
  'Full name (do not fill)',
  'Username',
  'Start Date',
  'Ship Date (t4i use only)'
];

// ============ 菜单 ============
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

// ============ Tab 发现工具 ============
// 运行此函数可查看源表中哪些 tab 包含目标列
function discoverTabs() {
  var ss = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  var sheets = ss.getSheets();
  var results = [];

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var name = sheet.getName();
    var lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      results.push(name + ' : SKIP (rows < 2)');
      continue;
    }

    // 只读取 row 2（列名行），列数取 sheet 实际列数
    var lastCol = sheet.getLastColumn();
    if (lastCol === 0) {
      results.push(name + ' : SKIP (no columns)');
      continue;
    }

    var headerRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
    var headers = [];
    for (var i = 0; i < headerRow.length; i++) {
      headers.push(String(headerRow[i]).trim());
    }

    var matched = [];
    for (var j = 0; j < DESIRED_COLUMNS.length; j++) {
      if (headers.indexOf(DESIRED_COLUMNS[j]) !== -1) {
        matched.push(DESIRED_COLUMNS[j]);
      }
    }

    if (matched.length > 0) {
      results.push(name + ' : MATCH (' + matched.length + '/' + DESIRED_COLUMNS.length + ' columns)');
    } else {
      results.push(name + ' : SKIP (0 matching columns)');
    }
  }

  var msg = results.join('\n');
  Logger.log(msg);
  SpreadsheetApp.getUi().alert('Tab Discovery Results', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============ 主同步函数 ============
function syncNewHireData() {
  var sourceSS = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  var destSS   = SpreadsheetApp.openById(DEST_SPREADSHEET_ID);
  var destSheet = destSS.getSheets()[0]; // 目标表第一个 sheet

  // 当天日期（23:59:59 作为截止线）
  var today = new Date();
  today.setHours(23, 59, 59, 999);

  var allRows = [];
  var sheets = sourceSS.getSheets();
  var tabsProcessed = 0;

  // ---------- 遍历每个 tab ----------
  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    // 至少需要 header 行 (row 2) + 1 行数据 (row 3)
    if (lastRow < 3 || lastCol === 0) continue;

    // 批量读取所有数据（从 row 1 开始，含 row 1）
    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

    // Row 2 (index 1) 是列名
    var headers = [];
    for (var h = 0; h < data[1].length; h++) {
      headers.push(String(data[1][h]).trim());
    }

    // 建立 desired column -> 源列索引 的映射
    var colMap = {};
    var matchCount = 0;
    for (var c = 0; c < DESIRED_COLUMNS.length; c++) {
      var col = DESIRED_COLUMNS[c];
      var idx = headers.indexOf(col);
      if (idx !== -1) {
        colMap[col] = idx;
        matchCount++;
      }
    }

    // 无匹配列则跳过此 tab
    if (matchCount === 0) continue;
    tabsProcessed++;

    // ---------- 遍历数据行 (从 row 3 = index 2 开始) ----------
    for (var i = 2; i < data.length; i++) {
      var row = data[i];

      // ---- 关键字段检查：至少一项非空 ----
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

      // ---- 日期过滤：只同步到当天 ----
      var startRaw = colMap['Start Date'] !== undefined ? row[colMap['Start Date']] : null;
      var shipRaw  = colMap['Ship Date (t4i use only)'] !== undefined ? row[colMap['Ship Date (t4i use only)']] : null;

      var startDate = parseDate_(startRaw);
      var shipDate  = parseDate_(shipRaw);

      var hasStart = (startDate !== null);
      var hasShip  = (shipDate !== null);

      // 如果存在日期，则要求至少一个日期 <= 今天；如果全部在未来，跳过
      if (hasStart || hasShip) {
        var allFuture = true;
        if (hasStart && startDate <= today) allFuture = false;
        if (hasShip  && shipDate  <= today) allFuture = false;
        if (allFuture) continue;
      }

      // ---- 按自定义顺序构建输出行 ----
      var outRow = [];
      for (var d = 0; d < DESIRED_COLUMNS.length; d++) {
        var colName = DESIRED_COLUMNS[d];
        if (colMap[colName] !== undefined) {
          outRow.push(row[colMap[colName]]);
        } else {
          outRow.push('');
        }
      }

      allRows.push(outRow);
    }
  }

  // ---------- 排序：Start Date (主) → Ship Date (次)，从早到晚 ----------
  var startIdx = DESIRED_COLUMNS.indexOf('Start Date');
  var shipIdx  = DESIRED_COLUMNS.indexOf('Ship Date (t4i use only)');

  allRows.sort(function(a, b) {
    var aStart = parseDate_(a[startIdx]);
    var bStart = parseDate_(b[startIdx]);
    var aShip  = parseDate_(a[shipIdx]);
    var bShip  = parseDate_(b[shipIdx]);

    // 主排序: Start Date
    if (aStart && bStart) {
      var diff = aStart.getTime() - bStart.getTime();
      if (diff !== 0) return diff;
    } else if (aStart && !bStart) {
      return -1; // 有日期的排前面
    } else if (!aStart && bStart) {
      return 1;
    }

    // 次排序: Ship Date
    if (aShip && bShip) {
      return aShip.getTime() - bShip.getTime();
    } else if (aShip && !bShip) {
      return -1;
    } else if (!aShip && bShip) {
      return 1;
    }

    return 0;
  });

  // ---------- 写入目标表 ----------
  destSheet.clearContents();
  destSheet.clearFormats();

  // 写入列名（第 1 行）
  destSheet.getRange(1, 1, 1, DESIRED_COLUMNS.length).setValues([DESIRED_COLUMNS]);
  destSheet.getRange(1, 1, 1, DESIRED_COLUMNS.length).setFontWeight('bold');

  // 写入数据
  if (allRows.length > 0) {
    destSheet.getRange(2, 1, allRows.length, DESIRED_COLUMNS.length).setValues(allRows);

    // 格式化日期列为 date 类型
    var dateFormat = 'M/d/yyyy';
    // Start Date 列
    destSheet.getRange(2, startIdx + 1, allRows.length, 1).setNumberFormat(dateFormat);
    // Ship Date 列
    destSheet.getRange(2, shipIdx + 1, allRows.length, 1).setNumberFormat(dateFormat);
  }

  SpreadsheetApp.flush();
  Logger.log('Sync complete: ' + allRows.length + ' rows from ' + tabsProcessed + ' tabs.');
}

// ============ 日期解析辅助函数 ============
function parseDate_(val) {
  if (val === null || val === undefined || val === '') return null;
  if (val instanceof Date) {
    return isNaN(val.getTime()) ? null : val;
  }
  var d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

// ============ 定时触发器 ============
function setupDailyTriggers() {
  // 先清除已有触发器
  removeTriggers();

  // 每天 6:00 AM ET
  ScriptApp.newTrigger('syncNewHireData')
    .timeBased()
    .atHour(6)
    .everyDays(1)
    .inTimezone('America/New_York')
    .create();

  // 每天 8:00 PM ET (20:00)
  ScriptApp.newTrigger('syncNewHireData')
    .timeBased()
    .atHour(20)
    .everyDays(1)
    .inTimezone('America/New_York')
    .create();

  Logger.log('Triggers created: 6am ET & 8pm ET daily.');
  SpreadsheetApp.getUi().alert('Daily triggers set successfully!\n\n- 6:00 AM ET\n- 8:00 PM ET');
}

function removeTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  Logger.log('All triggers removed.');
}
