// ================================================================
// T4i New Hire Data Sync Script — v2
// 从 'T4i New Hire List' 同步数据到 'newHireAutoSyncData'
// 
// v2 更新:
//   - Start Date 移至第1列, Ship Date 移至第2列
//   - 新增 Source Tab 列（最后一列）标注数据来源 tab
//   - 新增重复数据检测 (Full Name + Start Date + Hire Type)
//   - QlikSense 加载优化：干净的单表格式输出
// ================================================================

// ============ 配置区 ============
var SOURCE_SPREADSHEET_ID = '1cAlDPmgQ3mkmC5l5Jw-h6HlDre1RqIenAHVpoulZNwU';
var DEST_SPREADSHEET_ID   = '1mt2tOVCtef0y2vP5_Y53E-wpU2MZN_tKbbfDk7bWxbw';

// 需要从源表提取的列（按自定义显示顺序排列）
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

// 输出表头 = 源列 + 脚本生成列
var OUTPUT_HEADERS = SOURCE_COLUMNS.concat(['Source Tab']);

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
// 扫描源表所有 tab，报告哪些 tab 包含目标列
// 使用方法: 在菜单中点击 "Discover Tabs"，或在脚本编辑器中选择此函数点 Run
// 建议在以下情况运行: 源表新增/删除/重命名了 tab 时
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
      results.push('[SKIP] ' + name + ' — rows or columns insufficient');
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

// ============ 主同步函数 ============
function syncNewHireData() {
  var startTime = new Date();
  var sourceSS = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  var destSS   = SpreadsheetApp.openById(DEST_SPREADSHEET_ID);
  var destSheet = destSS.getSheets()[0]; // 目标表第一个 sheet

  // 当天日期（23:59:59 作为截止线）
  var today = new Date();
  today.setHours(23, 59, 59, 999);

  var allRows = [];
  var sheets = sourceSS.getSheets();
  var tabsProcessed = 0;

  // 重复检测追踪器: key = "FullName|StartDate|HireType"
  var dupTracker = {};

  // ---------- 遍历每个 tab ----------
  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var sheetName = sheet.getName();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    // 至少需要 header 行 (row 2) + 1 行数据 (row 3)
    if (lastRow < 3 || lastCol === 0) continue;

    // 批量读取所有数据
    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

    // Row 2 (index 1) 是列名
    var headers = [];
    for (var h = 0; h < data[1].length; h++) {
      headers.push(String(data[1][h]).trim());
    }

    // 建立 source column -> 源列索引 的映射
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

      // 如果存在日期，要求至少一个 <= 今天；全部在未来则跳过
      if (hasStart || hasShip) {
        var allFuture = true;
        if (hasStart && startDate <= today) allFuture = false;
        if (hasShip  && shipDate  <= today) allFuture = false;
        if (allFuture) continue;
      }

      // ---- 按自定义顺序构建输出行 ----
      var outRow = [];
      for (var d = 0; d < SOURCE_COLUMNS.length; d++) {
        var colName = SOURCE_COLUMNS[d];
        if (colMap[colName] !== undefined) {
          outRow.push(row[colMap[colName]]);
        } else {
          outRow.push('');
        }
      }
      // 追加 Source Tab 列
      outRow.push(sheetName);

      // ---- 重复检测 ----
      var fullNameVal = colMap['Full name (do not fill)'] !== undefined ? String(row[colMap['Full name (do not fill)']]).trim() : '';
      var startDateStr = startDate ? startDate.toDateString() : '';
      var hireTypeVal = colMap['Hire Type'] !== undefined ? String(row[colMap['Hire Type']]).trim() : '';
      var dupKey = fullNameVal + '|' + startDateStr + '|' + hireTypeVal;

      if (dupKey !== '||') { // 只追踪有实际值的组合
        if (!dupTracker[dupKey]) {
          dupTracker[dupKey] = 0;
        }
        dupTracker[dupKey]++;
      }

      allRows.push(outRow);
    }
  }

  // ---------- 排序：Start Date (主) → Ship Date (次)，从早到晚 ----------
  var startIdx = SOURCE_COLUMNS.indexOf('Start Date');
  var shipIdx  = SOURCE_COLUMNS.indexOf('Ship Date (t4i use only)');

  allRows.sort(function(a, b) {
    var aStart = parseDate_(a[startIdx]);
    var bStart = parseDate_(b[startIdx]);
    var aShip  = parseDate_(a[shipIdx]);
    var bShip  = parseDate_(b[shipIdx]);

    // 主排序: Start Date (有日期排前面，日期从早到晚)
    if (aStart && bStart) {
      var diff = aStart.getTime() - bStart.getTime();
      if (diff !== 0) return diff;
    } else if (aStart && !bStart) {
      return -1;
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

  var totalCols = OUTPUT_HEADERS.length; // 19 列 (18 源列 + Source Tab)

  // 写入表头（第 1 行）
  destSheet.getRange(1, 1, 1, totalCols).setValues([OUTPUT_HEADERS]);
  destSheet.getRange(1, 1, 1, totalCols).setFontWeight('bold');
  destSheet.setFrozenRows(1); // 冻结表头行，方便 QlikSense 加载时识别

  // 写入数据
  if (allRows.length > 0) {
    destSheet.getRange(2, 1, allRows.length, totalCols).setValues(allRows);

    // 格式化日期列为 date 类型 (M/d/yyyy)
    var dateFormat = 'M/d/yyyy';
    destSheet.getRange(2, startIdx + 1, allRows.length, 1).setNumberFormat(dateFormat);
    destSheet.getRange(2, shipIdx + 1, allRows.length, 1).setNumberFormat(dateFormat);
  }

  // ---------- 自动调整列宽 ----------
  for (var col = 1; col <= totalCols; col++) {
    destSheet.autoResizeColumn(col);
  }

  SpreadsheetApp.flush();

  // ---------- 统计重复 ----------
  var dupCount = 0;
  var dupDetails = [];
  for (var key in dupTracker) {
    if (dupTracker[key] > 1) {
      dupCount++;
      dupDetails.push(key.replace(/\|/g, ' / ') + ' (x' + dupTracker[key] + ')');
    }
  }

  // ---------- 同步摘要日志 ----------
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
  SpreadsheetApp.getUi().alert(
    'Daily triggers set successfully!\n\n' +
    '- Sync 1: 6:00 AM Eastern Time\n' +
    '- Sync 2: 8:00 PM Eastern Time\n\n' +
    'Note: Google triggers may run within a ±15 min window.'
  );
}

function removeTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  Logger.log('All triggers removed.');
}
