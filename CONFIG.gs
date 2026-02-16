/**
 * マスタ設定関連。
 * - 作業マスタの入力規則（ドロップダウン）
 * - 条件付き書式の色連動
 * - 各種マスタ読み込み（作業・会社・スタッフ属性）
 */
function setupAreaDropdownForColumnC() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('作業マスタ');
  if (!sh) throw new Error('シート「作業マスタ」が見つかりません。');

  const lastRow = Math.max(sh.getLastRow(), 2);
  const items = buildDropdownItemsFromColumn_(sh.getRange(2, 7, lastRow - 1, 1));
  if (items.length === 0) throw new Error('作業マスタ!G2以下に候補（空白以外）がありません。');

  const applyRows = Math.max(lastRow - 1, 1);
  const targetRange = sh.getRange(2, 3, applyRows, 1);
  applyDropdownWithColorRules_(sh, targetRange, items, '=$C2=');

  targetRange.setHorizontalAlignment('center');
  targetRange.setFontWeight('bold');
}

function applyWorkDropdown_Q_to_Layout_and_WorkMasterF() {
  const ss = SpreadsheetApp.getActive();
  const srcSh = ss.getSheetByName('作業マスタ');
  if (!srcSh) throw new Error('シート「作業マスタ」が見つかりません。');

  const layoutSh = ss.getSheetByName('配置表');
  if (!layoutSh) throw new Error('シート「配置表」が見つかりません。');

  const srcLastRow = Math.max(srcSh.getLastRow(), 2);
  const items = buildDropdownItemsFromColumn_(srcSh.getRange(2, 17, srcLastRow - 1, 1));
  if (items.length === 0) throw new Error('作業マスタ!Q2以下に候補（空白以外）がありません。');

  const dstStartRow = 2;
  const dstStartCol = 10;
  const dstEndCol = 105;
  const dstLastRow = Math.max(layoutSh.getLastRow(), dstStartRow);
  const numRows = dstLastRow - dstStartRow + 1;
  const numCols = dstEndCol - dstStartCol + 1;
  const layoutTargetRange = layoutSh.getRange(dstStartRow, dstStartCol, numRows, numCols);
  applyDropdownWithColorRules_(layoutSh, layoutTargetRange, items, '=J2=');
  layoutTargetRange.setHorizontalAlignment('center');
  layoutTargetRange.setFontWeight('bold');

  const masterTargetRange = srcSh.getRange('F2:F100');
  applyDropdownWithColorRules_(srcSh, masterTargetRange, items, '=$F2=');
  masterTargetRange.setHorizontalAlignment('center');
  masterTargetRange.setFontWeight('bold');
}

function buildDropdownItemsFromColumn_(srcRange) {
  const values = srcRange.getDisplayValues().flat();
  const bgs = srcRange.getBackgrounds().flat();
  const fgs = srcRange.getFontColors().flat();

  const seen = new Set();
  const items = [];
  for (let i = 0; i < values.length; i++) {
    const v = (values[i] || '').trim();
    if (!v || seen.has(v)) continue;
    seen.add(v);
    items.push({ value: v, bg: bgs[i] || '#ffffff', fg: fgs[i] || '#000000' });
  }
  return items;
}

function applyDropdownWithColorRules_(sheet, targetRange, items, formulaPrefix) {
  const dv = SpreadsheetApp.newDataValidation()
    .requireValueInList(items.map(x => x.value), true)
    .setAllowInvalid(false)
    .build();

  targetRange.setDataValidation(dv);
  removeConditionalRulesIntersecting_(sheet, targetRange);

  const newRules = items.map(it => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`${formulaPrefix}${JSON.stringify(it.value)}`)
    .setRanges([targetRange])
    .setBackground(it.bg)
    .setFontColor(it.fg)
    .build());

  sheet.setConditionalFormatRules(sheet.getConditionalFormatRules().concat(newRules));
}

function removeConditionalRulesIntersecting_(sheet, targetRange) {
  const all = sheet.getConditionalFormatRules();

  const tR1 = targetRange.getRow();
  const tC1 = targetRange.getColumn();
  const tR2 = tR1 + targetRange.getNumRows() - 1;
  const tC2 = tC1 + targetRange.getNumColumns() - 1;

  function intersects(rg) {
    const r1 = rg.getRow();
    const c1 = rg.getColumn();
    const r2 = r1 + rg.getNumRows() - 1;
    const c2 = c1 + rg.getNumColumns() - 1;
    return !(r2 < tR1 || tR2 < r1 || c2 < tC1 || tC2 < c1);
  }

  const keep = all.filter(rule => !(rule.getRanges() || []).some(intersects));
  sheet.setConditionalFormatRules(keep);
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
