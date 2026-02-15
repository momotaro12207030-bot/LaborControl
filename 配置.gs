/**
 * 🛠️ LOGI-MATRIX: サーバーサイドロジック v6.5
 * 最適化内容: ループ削減、I/O 最小化、堅牢なロック解放、書き込み範囲最適化
 */

const CONFIG = {
  COL_OFFSET: 1,
  SHEET_NAMES: {
    ASSIGNMENT: '割り当て',
    STAFF_MASTER: 'スタッフマスタ',
    WORK_MASTER: '作業マスタ',
    COMPANY_MASTER: '会社マスタ',
    PASTE_TARGET: '配置表'
  },
  UI: {
    PANEL_WIDTH: 1200,
    PANEL_HEIGHT: 850,
    DIALOG_WIDTH: 460,
    DIALOG_HEIGHT: 260
  },
  LOCK_TIMEOUT: 10000,
  OCR_FOLDER_ID: 'ここにGoogleドライブのフォルダIDを入力',
  PASTE_SETTINGS: {
    START_ROW: 2,
    NUM_ROWS: 20,
    SOURCE_COL: 9,
    TARGET_START_COL: 10,
    TARGET_END_COL: 105
  }
};

const DEFAULT_CONFIG = [
  { id: 'areaA', name: '4F 入荷荷降', floor: '4F', col: 10 },
  { id: 'areaB', name: '4F ピッキング', floor: '4F', col: 11 },
  { id: 'areaC', name: '4F 梱包出荷', floor: '4F', col: 12 },
  { id: 'areaD', name: '5F 入荷検品', floor: '5F', col: 13 },
  { id: 'areaE', name: '5F ピッキング', floor: '5F', col: 14 },
  { id: 'areaF', name: '5F ラベル貼', floor: '5F', col: 15 },
  { id: 'areaG', name: '事務・受付', floor: 'OFFICE', col: 16 }
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚚 配置システム')
    .addItem('配置反映 (色付きセル)', 'showConfirmDialog')
    .addSeparator()
    .addItem('配置管理パネルを開く', 'showAdminPanel')
    .addToUi();
}

function showAdminPanel() {
  const html = HtmlService.createHtmlOutputFromFile('index')
    .setWidth(CONFIG.UI.PANEL_WIDTH)
    .setHeight(CONFIG.UI.PANEL_HEIGHT)
    .setTitle('LOGI-MATRIX | Synapse Sync');
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function showConfirmDialog() {
  const html = HtmlService.createHtmlOutputFromFile('confirmDialog')
    .setWidth(CONFIG.UI.DIALOG_WIDTH)
    .setHeight(CONFIG.UI.DIALOG_HEIGHT);
  SpreadsheetApp.getUi().showModalDialog(html, '配置反映の確認');
}

function runOcrAndReflect() {
  const folder = DriveApp.getFolderById(CONFIG.OCR_FOLDER_ID);
  const files = folder.getFiles();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.ASSIGNMENT);
  if (!sh) throw new Error(`シートが見つかりません: ${CONFIG.SHEET_NAMES.ASSIGNMENT}`);

  let row = CONFIG.PASTE_SETTINGS.START_ROW;
  let count = 0;
  const maxRow = CONFIG.PASTE_SETTINGS.START_ROW + CONFIG.PASTE_SETTINGS.NUM_ROWS - 1;

  while (files.hasNext() && row <= maxRow) {
    const file = files.next();
    const resource = { title: file.getName(), mimeType: file.getMimeType() };
    const docFile = Drive.Files.insert(resource, file.getBlob(), { ocr: true });
    const doc = DocumentApp.openById(docFile.id);
    const translatedText = LanguageApp.translate(doc.getBody().getText().trim(), '', 'ja');

    sh.getRange(row, CONFIG.PASTE_SETTINGS.SOURCE_COL).setValue(translatedText);
    Drive.Files.remove(docFile.id);
    file.setTrashed(true);

    row++;
    count++;
  }

  return `画像 ${count} 件の処理が完了しました。`;
}

function executePaste() {
  const sh = SpreadsheetApp.getActiveSheet();
  const targetSheetName = CONFIG.SHEET_NAMES.PASTE_TARGET;

  if (sh.getName() !== targetSheetName) {
    SpreadsheetApp.getUi().alert(`❌ 実行エラー\nこの機能は「${targetSheetName}」シートでのみ動作します。\n現在のシート: ${sh.getName()}`);
    return;
  }
  pasteToColoredCells_multiRows(sh);
}

function pasteToColoredCells_multiRows(sh) {
  const sRow = CONFIG.PASTE_SETTINGS.START_ROW;
  const nRows = CONFIG.PASTE_SETTINGS.NUM_ROWS;
  const sCol = CONFIG.PASTE_SETTINGS.SOURCE_COL;
  const tStart = CONFIG.PASTE_SETTINGS.TARGET_START_COL;
  const tEnd = CONFIG.PASTE_SETTINGS.TARGET_END_COL;
  const nCols = tEnd - tStart + 1;

  const actualNumRows = Math.min(nRows, sh.getMaxRows() - sRow + 1);
  if (actualNumRows <= 0) return;

  const sourceValues = sh.getRange(sRow, sCol, actualNumRows, 1).getValues();
  const targetRange = sh.getRange(sRow, tStart, actualNumRows, nCols);
  const targetValues = targetRange.getValues();
  const targetBackgrounds = targetRange.getBackgrounds();

  let totalChanged = 0;
  for (let r = 0; r < actualNumRows; r++) {
    const newValue = sourceValues[r][0] || '';
    for (let c = 0; c < nCols; c++) {
      const isWhite = normalizeColor_(targetBackgrounds[r][c]) === '#ffffff';
      targetValues[r][c] = isWhite ? '' : newValue;
      if (!isWhite) totalChanged++;
    }
  }

  targetRange.setValues(targetValues);
  SpreadsheetApp.getActive().toast(`反映完了: ${totalChanged}箇所`, '完了');
}

function normalizeColor_(color) {
  if (!color || color === 'white' || color === 'transparent') return '#ffffff';
  return String(color).trim().toLowerCase();
}

function getStaffDataFromSheet76() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ASSIGNMENT);
  const companyColors = getCompanyColors();
  const config = getWorkConfig();
  const staffAttributes = getStaffAttributes();

  const result = { assignments: { pool: [] }, config, companyColors };
  config.forEach(item => { result.assignments[item.id] = []; });
  if (!sheet) return result;

  const lastRow = Math.max(sheet.getLastRow(), 3);
  const maxCol = Math.max.apply(null, config.map(i => i.col));
  const currentData = sheet.getRange(3, 1, lastRow - 2, maxCol).getValues();

  const staffToCompanyMap = {};
  const assignedSet = new Set();

  currentData.forEach(row => {
    const company = String(row[0] || '自社').trim();
    const name = String(row[1] || '').trim();
    if (name) staffToCompanyMap[name] = company;
  });

  currentData.forEach(row => {
    for (let i = 0; i < config.length; i++) {
      const item = config[i];
      const name = String(row[item.col - 1] || '').trim();
      if (!name || name === 'undefined') continue;
      result.assignments[item.id].push({
        id: name,
        name,
        company: staffToCompanyMap[name] || '未設定',
        attr: staffAttributes[name] || ''
      });
      assignedSet.add(name);
    }
  });

  currentData.forEach(row => {
    const company = String(row[0] || '自社').trim();
    const name = String(row[1] || '').trim();
    if (name && !assignedSet.has(name)) {
      result.assignments.pool.push({ id: name, name, company, attr: staffAttributes[name] || '' });
    }
  });

  return result;
}


function autoAssignByMainWork(jsonString) {
  try {
    const data = JSON.parse(jsonString);
    const config = getWorkConfig();
    const mainWorkMap = getStaffMainWorkMap();
    const { assignments, movedCount, unmatchedMainWorks } = applyAutoAssignByMainWork_(data, config, mainWorkMap);
    return { success: true, assignments, movedCount, unmatchedMainWorks };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function applyAutoAssignByMainWork_(data, config, mainWorkMap) {
  const areaIds = new Set(['pool']);
  config.forEach(c => areaIds.add(c.id));
  Object.keys(data).forEach(k => areaIds.add(k));

  const next = {};
  areaIds.forEach(id => { next[id] = Array.isArray(data[id]) ? [] : []; });

  const workToAreaMap = buildWorkToAreaMap_(config);
  const unmatched = new Set();
  let movedCount = 0;

  Object.keys(data).forEach(fromArea => {
    const list = Array.isArray(data[fromArea]) ? data[fromArea] : [];
    list.forEach(staff => {
      const name = staff && staff.name ? String(staff.name).trim() : '';
      const mainWork = name ? (mainWorkMap[name] || '') : '';
      const toArea = resolveAreaIdFromMainWork_(mainWork, config, workToAreaMap);

      if (mainWork && !toArea) unmatched.add(mainWork);

      const targetArea = toArea || fromArea;
      if (!next[targetArea]) next[targetArea] = [];
      next[targetArea].push(staff);

      if (toArea && toArea !== fromArea) movedCount++;
    });
  });

  return { assignments: next, movedCount, unmatchedMainWorks: Array.from(unmatched) };
}

function getStaffMainWorkMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.STAFF_MASTER);
  const result = {};
  if (!masterSheet || masterSheet.getLastRow() < 2) return result;

  const lastCol = Math.max(masterSheet.getLastColumn(), 3);
  const headers = masterSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
  const headerIndex = {};
  headers.forEach((h, i) => { if (h) headerIndex[h] = i; });

  const nameIdx =
    headerIndex['氏名'] ??
    headerIndex['名前'] ??
    headerIndex['Name'] ??
    1;
  const mainWorkIdx =
    headerIndex['メイン業務'] ??
    headerIndex['主作業'] ??
    2;

  const rows = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, lastCol).getValues();
  rows.forEach(row => {
    const name = String(row[nameIdx] || '').trim();
    const mainWork = String(row[mainWorkIdx] || '').trim();
    if (name) result[name] = mainWork;
  });

  return result;
}

function buildWorkToAreaMap_(config) {
  const map = {};
  config.forEach(item => {
    const normalizedName = normalizeWorkText_(item.name);
    if (normalizedName) map[normalizedName] = item.id;
  });
  return map;
}

function resolveAreaIdFromMainWork_(mainWork, config, workToAreaMap) {
  if (!mainWork) return null;
  const normalized = normalizeWorkText_(mainWork);
  if (!normalized) return null;

  if (workToAreaMap[normalized]) return workToAreaMap[normalized];

  for (let i = 0; i < config.length; i++) {
    const targetNormalized = normalizeWorkText_(config[i].name);
    if (!targetNormalized) continue;
    if (targetNormalized.includes(normalized) || normalized.includes(targetNormalized)) {
      return config[i].id;
    }
  }
  return null;
}

function normalizeWorkText_(text) {
  return String(text || '')
    .trim()
    .toLowerCase()
    .replace(/[\s　]+/g, '')
    .replace(/[→＞>]+/g, '->');
}

function saveAssignmentsToSheet76(jsonString, actionType) {
  const mode = actionType || 'CHECK';
  const lock = LockService.getScriptLock();
  let isLocked = false;

  try {
    isLocked = lock.tryLock(CONFIG.LOCK_TIMEOUT);
    if (!isLocked) throw new Error('保存処理が競合しています');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ASSIGNMENT);
    if (!sheet) throw new Error(`シートが見つかりません: ${CONFIG.SHEET_NAMES.ASSIGNMENT}`);

    let data = JSON.parse(jsonString);
    const validStaffSet = getValidStaffSet(sheet);
    const unknownNames = findUnknownStaff(data, validStaffSet);

    if (mode === 'CHECK' && unknownNames.length > 0) {
      return { success: false, confirmNeeded: true, unknownNames };
    }
    if (mode === 'DELETE') {
      data = removeUnknownStaff(data, validStaffSet);
    }

    writeAssignmentsToSheet(sheet, data, getWorkConfig());
    return { success: true, message: '保存完了' };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    if (isLocked) lock.releaseLock();
  }
}

function getCompanyColors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.COMPANY_MASTER);
  const colorMap = {};
  if (!mSheet || mSheet.getLastRow() < 2) return colorMap;

  const range = mSheet.getRange(2, 1, mSheet.getLastRow() - 1, 1);
  const names = range.getValues();
  const colors = range.getBackgrounds();
  names.forEach((row, i) => {
    if (row[0]) colorMap[String(row[0]).trim()] = colors[i][0];
  });
  return colorMap;
}

function getWorkConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.WORK_MASTER);
  if (!configSheet || configSheet.getLastRow() < 2) return DEFAULT_CONFIG;

  return configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 4).getValues()
    .filter(r => r[0])
    .map(r => ({
      id: String(r[0]),
      name: String(r[1]),
      floor: String(r[2]),
      col: Number(r[3]) + CONFIG.COL_OFFSET
    }));
}

function getStaffAttributes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.STAFF_MASTER);
  const staffAttributes = {};
  if (!masterSheet || masterSheet.getLastRow() < 2) return staffAttributes;

  masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 3).getValues().forEach(row => {
    if (row[0]) staffAttributes[String(row[0]).trim()] = [row[1], row[2]].filter(Boolean).join(' | ');
  });
  return staffAttributes;
}

function getValidStaffSet(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 3);
  return new Set(
    sheet.getRange(3, 2, lastRow - 2, 1)
      .getValues()
      .flat()
      .map(s => String(s).trim())
      .filter(Boolean)
  );
}

function findUnknownStaff(data, validStaffSet) {
  const unknown = new Set();
  Object.keys(data).forEach(key => {
    if (!Array.isArray(data[key])) return;
    data[key].forEach(staff => {
      const staffName = staff && staff.name ? String(staff.name).trim() : '';
      if (staffName && !validStaffSet.has(staffName)) unknown.add(staffName);
    });
  });
  return Array.from(unknown);
}

function removeUnknownStaff(data, validStaffSet) {
  const cleaned = {};
  Object.keys(data).forEach(key => {
    if (Array.isArray(data[key])) {
      cleaned[key] = data[key].filter(s => validStaffSet.has(String(s.name).trim()));
    } else {
      cleaned[key] = data[key];
    }
  });
  return cleaned;
}

function writeAssignmentsToSheet(sheet, data, config) {
  const maxRows = sheet.getMaxRows();
  if (maxRows < 3) return;

  config.forEach(item => {
    const col = item.col;
    const existingLast = Math.max(sheet.getLastRow(), 3);
    const clearRows = Math.max(existingLast - 2, 1);
    sheet.getRange(3, col, clearRows, 1).clearContent();

    const staffArray = data[item.id] || [];
    if (staffArray.length > 0) {
      sheet.getRange(3, col, staffArray.length, 1).setValues(staffArray.map(s => [s.name]));
    }
  });
}

const OPS_SHEET_NAMES = {
  STAFF_MASTER: 'スタッフマスタ',
  SHIFT: 'シフト表',
  EXTRA: '追加人員',
  ATTENDANCE: '出勤者',
  WORK_MASTER: '作業マスタ',
  ASSIGNMENT: '配置表',
  PROGRESS_INPUT: '進捗入力',
  PRODUCTIVITY: '個人別生産性'
};

function generateAttendance(date) {
  const targetDate = normalizeDateInput_(date);
  const weekday = getWeekdayJa_(targetDate);
  const staffSheet = getRequiredSheet_(OPS_SHEET_NAMES.STAFF_MASTER);
  const shiftSheet = getRequiredSheet_(OPS_SHEET_NAMES.SHIFT);
  const extraSheet = getRequiredSheet_(OPS_SHEET_NAMES.EXTRA);
  const attendanceSheet = getRequiredSheet_(OPS_SHEET_NAMES.ATTENDANCE);

  const staffRows = readSheetObjects_(staffSheet);
  const shiftRows = readSheetObjects_(shiftSheet);
  const extraRows = readSheetObjects_(extraSheet);
  const staffById = {};

  staffRows.forEach(row => {
    const staffId = toText_(row['スタッフID']);
    if (staffId) {
      staffById[staffId] = {
        名前: toText_(row['名前']),
        会社: toText_(row['会社']),
        雇用区分: toText_(row['雇用区分'])
      };
    }
  });

  const mergedMap = new Map();

  shiftRows.forEach(row => {
    const staffId = toText_(row['スタッフID']);
    const shiftValue = toText_(row[weekday]);
    if (!staffId || shiftValue !== '○') return;

    const master = staffById[staffId] || {};
    mergedMap.set(staffId, {
      日付: formatDateKey_(targetDate),
      スタッフID: staffId,
      名前: toText_(master.名前),
      会社: toText_(master.会社),
      雇用区分: toText_(master.雇用区分)
    });
  });

  extraRows.forEach(row => {
    const rowDate = normalizeDateString_(row['日付']);
    const targetDateKey = formatDateKey_(targetDate);
    if (rowDate !== targetDateKey) return;

    const staffId = toText_(row['スタッフID']);
    if (!staffId) return;
    mergedMap.set(staffId, {
      日付: targetDateKey,
      スタッフID: staffId,
      名前: toText_(row['名前']),
      会社: toText_(row['会社']),
      雇用区分: toText_(row['雇用区分'])
    });
  });

  const attendanceRows = Array.from(mergedMap.values())
    .sort((a, b) => a['スタッフID'].localeCompare(b['スタッフID'], 'ja'));

  replaceRowsByDate_(attendanceSheet, '日付', formatDateKey_(targetDate), [
    '日付', 'スタッフID', '名前', '会社', '雇用区分'
  ], attendanceRows);

  return {
    date: formatDateKey_(targetDate),
    count: attendanceRows.length,
    message: `出勤者を ${attendanceRows.length} 名生成しました。`
  };
}

function assignWork(date) {
  const targetDate = normalizeDateInput_(date);
  const targetDateKey = formatDateKey_(targetDate);
  const attendanceSheet = getRequiredSheet_(OPS_SHEET_NAMES.ATTENDANCE);
  const workSheet = getRequiredSheet_(OPS_SHEET_NAMES.WORK_MASTER);
  const assignmentSheet = getRequiredSheet_(OPS_SHEET_NAMES.ASSIGNMENT);

  const attendanceRows = readSheetObjects_(attendanceSheet)
    .filter(row => normalizeDateString_(row['日付']) === targetDateKey);
  const workRows = readSheetObjects_(workSheet)
    .filter(row => toText_(row['作業ID']));

  if (workRows.length === 0) {
    throw new Error('作業マスタに作業IDがありません。');
  }

  const assignments = attendanceRows.map((row, index) => {
    const work = workRows[index % workRows.length];
    return {
      日付: targetDateKey,
      スタッフID: toText_(row['スタッフID']),
      作業ID: toText_(work['作業ID'])
    };
  });

  replaceRowsByDate_(assignmentSheet, '日付', targetDateKey, ['日付', 'スタッフID', '作業ID'], assignments);

  return {
    date: targetDateKey,
    count: assignments.length,
    message: `作業割り振りを ${assignments.length} 件作成しました。`
  };
}

function calculateThroughput(date) {
  const targetDate = normalizeDateInput_(date);
  const targetDateKey = formatDateKey_(targetDate);
  const assignmentSheet = getRequiredSheet_(OPS_SHEET_NAMES.ASSIGNMENT);
  const workSheet = getRequiredSheet_(OPS_SHEET_NAMES.WORK_MASTER);

  const assignmentRows = readSheetObjects_(assignmentSheet)
    .filter(row => normalizeDateString_(row['日付']) === targetDateKey);
  const workRows = readSheetObjects_(workSheet);

  const tpByWork = {};
  workRows.forEach(row => {
    const workId = toText_(row['作業ID']);
    if (!workId) return;
    tpByWork[workId] = Number(row['基準TP（1人あたり）']) || 0;
  });

  let totalThroughput = 0;
  assignmentRows.forEach(row => {
    const workId = toText_(row['作業ID']);
    totalThroughput += tpByWork[workId] || 0;
  });

  return {
    date: targetDateKey,
    assignedCount: assignmentRows.length,
    totalThroughput
  };
}

function calculateProgress(date) {
  const targetDate = normalizeDateInput_(date);
  const targetDateKey = formatDateKey_(targetDate);
  const progressSheet = getRequiredSheet_(OPS_SHEET_NAMES.PROGRESS_INPUT);
  const progressRows = readSheetObjects_(progressSheet)
    .filter(row => normalizeDateString_(row['日付']) === targetDateKey);

  const actual = progressRows.reduce((sum, row) => sum + (Number(row['実績数']) || 0), 0);
  const throughput = calculateThroughput(targetDate);
  const target = Number(throughput.totalThroughput) || 0;

  return {
    date: targetDateKey,
    actual,
    target,
    diff: actual - target
  };
}

function calculateProductivity(date) {
  const targetDate = normalizeDateInput_(date);
  const targetDateKey = formatDateKey_(targetDate);
  const progressSheet = getRequiredSheet_(OPS_SHEET_NAMES.PROGRESS_INPUT);
  const productivitySheet = getRequiredSheet_(OPS_SHEET_NAMES.PRODUCTIVITY);
  const rows = readSheetObjects_(progressSheet)
    .filter(row => normalizeDateString_(row['日付']) === targetDateKey);

  const output = rows.map(row => {
    const actual = Number(row['実績数']) || 0;
    const workHours =
      Number(row['作業時間']) ||
      Number(row['稼働時間']) ||
      Number(row['時間']) ||
      1;

    return {
      スタッフID: toText_(row['スタッフID']),
      日付: targetDateKey,
      作業ID: toText_(row['作業ID']),
      実績数: actual,
      生産性: workHours > 0 ? actual / workHours : 0
    };
  });

  replaceRowsByDate_(productivitySheet, '日付', targetDateKey,
    ['スタッフID', '日付', '作業ID', '実績数', '生産性'], output);

  return {
    date: targetDateKey,
    count: output.length
  };
}

function getRequiredSheet_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`シートが見つかりません: ${sheetName}`);
  }
  return sheet;
}

function readSheetObjects_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(header => toText_(header));
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  return values.map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      if (!header) return;
      obj[header] = row[index];
    });
    return obj;
  });
}

function replaceRowsByDate_(sheet, dateHeader, dateKey, requiredHeaders, newRows) {
  const headers = ensureHeaders_(sheet, requiredHeaders);
  const dateColIndex = headers.indexOf(dateHeader);
  if (dateColIndex === -1) {
    throw new Error(`日付列が見つかりません: ${dateHeader}`);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const valueRange = sheet.getRange(2, 1, lastRow - 1, headers.length);
    const values = valueRange.getValues();
    const remained = values.filter(row => normalizeDateString_(row[dateColIndex]) !== dateKey);

    const merged = remained.concat(newRows.map(rowObj => headers.map(header => rowObj[header] ?? '')));
    valueRange.clearContent();
    if (merged.length > 0) {
      sheet.getRange(2, 1, merged.length, headers.length).setValues(merged);
    }
    if (merged.length < values.length) {
      sheet.getRange(2 + merged.length, 1, values.length - merged.length, headers.length).clearContent();
    }
    return;
  }

  if (newRows.length > 0) {
    const output = newRows.map(rowObj => headers.map(header => rowObj[header] ?? ''));
    sheet.getRange(2, 1, output.length, headers.length).setValues(output);
  }
}

function ensureHeaders_(sheet, requiredHeaders) {
  const lastCol = Math.max(sheet.getLastColumn(), requiredHeaders.length);
  let headers = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(value => toText_(value))
    : [];

  if (headers.length === 0) {
    headers = requiredHeaders.slice();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return headers;
  }

  requiredHeaders.forEach(header => {
    if (!headers.includes(header)) headers.push(header);
  });

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return headers;
}

function normalizeDateInput_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return value;
  }
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) {
    throw new Error(`不正な日付です: ${value}`);
  }
  return parsed;
}

function normalizeDateString_(value) {
  if (!value) return '';
  const dateObj = normalizeDateInput_(value);
  return formatDateKey_(dateObj);
}

function formatDateKey_(dateObj) {
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy/MM/dd');
}

function getWeekdayJa_(dateObj) {
  return ['日', '月', '火', '水', '木', '金', '土'][dateObj.getDay()];
}

function toText_(value) {
  return String(value ?? '').trim();
}
