/**
 * LOGI-MATRIX 割り当て管理ロジック。
 */
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
  areaIds.forEach(id => { next[id] = []; });

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

  const nameIdx = headerIndex['氏名'] ?? headerIndex['名前'] ?? headerIndex['Name'] ?? 1;
  const mainWorkIdx = headerIndex['メイン業務'] ?? headerIndex['主作業'] ?? 2;

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
