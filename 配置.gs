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

function buildWorkNameToProcessMap_() {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName("作業マスタ");
  if (!master) throw new Error("作業マスタ シートが見つかりません");

  // B=工程名(キー), F=作業カテゴリ(値)
  // 例: Shipdock -> PICK
  const lastRow = Math.max(master.getLastRow(), 2);
  const keyRange = master.getRange(2, 2, lastRow - 1, 1).getDisplayValues(); // B2:B
  const valRange = master.getRange(2, 6, lastRow - 1, 1).getDisplayValues(); // F2:F

  const map = {};
  for (let i = 0; i < keyRange.length; i++) {
    const k = normalizeKey_(keyRange[i][0]);
    const v = (valRange[i][0] || "").trim();
    if (!k || !v) continue;
    map[k] = v;
  }
  return map;
}

function normalizeKey_(s) {
  if (s === null || s === undefined) return "";
  // 全角スペースを半角に、連続スペースを1つに、前後trim
  return String(s)
    .replace(/\u3000/g, " ")
    .replace(/\s+/g, " ")
    .trim();
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

function pasteToGrayCellsByDB_multiRows(sh) {
  const sRow = CONFIG.PASTE_SETTINGS.START_ROW;
  const nRows = CONFIG.PASTE_SETTINGS.NUM_ROWS;

  // 参照元：配置表 I列（工程名）
  const sourceCol = 9; // I

  // 反映先：J～DA
  const tStart = 10;   // J
  const tEnd   = 105;  // DA
  const nCols  = tEnd - tStart + 1;

  // 判定元：DB～GS（J～DAと同じ幅）
  const dbStart = 106; // DB
  const dbEnd   = dbStart + nCols - 1; // GS想定

  const actualNumRows = Math.min(nRows, sh.getMaxRows() - sRow + 1);
  if (actualNumRows <= 0) return;

  // B→F の辞書（工程名→作業カテゴリ）
  const map = buildWorkNameToProcessMap_();

  // I列（工程名）
  const keys = sh.getRange(sRow, sourceCol, actualNumRows, 1).getDisplayValues();

  // DB側（稼働フラグ/稼働量 0 or 0.25）
  const dbValues = sh.getRange(sRow, dbStart, actualNumRows, nCols).getValues();

  // 反映先
  const targetRange = sh.getRange(sRow, tStart, actualNumRows, nCols);
  const targetValues = targetRange.getValues();

  let totalChanged = 0;

  for (let r = 0; r < actualNumRows; r++) {
    const key = normalizeKey_(keys[r][0]);
    const newValue = (key && map[key]) ? map[key] : ""; // 辞書に無ければ空

    for (let c = 0; c < nCols; c++) {
      const active = Number(dbValues[r][c] || 0); // 0 / 0.25 など
      targetValues[r][c] = active > 0 ? newValue : "";
      totalChanged++;
    }
  }

  targetRange.setValues(targetValues);
  SpreadsheetApp.getActive().toast(`反映完了: ${totalChanged}セル`, "完了");
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
