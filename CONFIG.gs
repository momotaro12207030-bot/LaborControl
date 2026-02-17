/**
 * マスタ設定関連の処理群。
 * 既存ロジックから分離し、初回のみ適用する書式設定をここで管理する。
 */


/**
 * マスタ設定をまとめて実行するエントリポイント。
 */
function setupWorkMasters() {
  setupAreaDropdownForColumnC();
  applyWorkDropdown_Q_to_Layout_and_WorkMasterF();
}
/**
 * 作業マスタ!G2:G の非空セルを候補にして、作業マスタ!C2:C にプルダウンを一括設定。
 * さらに G列の各セルの書式（背景色/文字色）を元に、C列が選択値に応じて同じ色になるよう
 * 条件付き書式を自動生成する。
 */
function setupAreaDropdownForColumnC() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('作業マスタ');
  if (!sh) throw new Error('シート「作業マスタ」が見つかりません。');

  const items = collectDropdownItemsFromColumn_(sh, 7, '作業マスタ!G2以下に候補（空白以外）がありません。');
  const lastRow = Math.max(sh.getLastRow(), 2);
  const applyRows = Math.max(lastRow - 1, 1);
  const targetRange = sh.getRange(2, 3, applyRows, 1); // C2:C

  applyDropdownAndColorRules_(sh, targetRange, items, '=$C2=');

  targetRange.setHorizontalAlignment('center');
  targetRange.setFontWeight('bold');
}

/**
 * 作業マスタ!Q2:Q の非空セルを候補にして、
 * 1) 配置表!J2:DA（2行目～最終行）にドロップダウン＋色連動
 * 2) 作業マスタ!F2:F100 にも同じドロップダウン＋色連動
 * 3) 配置表!K2:DB1000 の初回用条件付き書式を一度だけ反映
 */
function applyWorkDropdown_Q_to_Layout_and_WorkMasterF() {
  const ss = SpreadsheetApp.getActive();

  const srcSh = ss.getSheetByName('作業マスタ');
  if (!srcSh) throw new Error('シート「作業マスタ」が見つかりません。');

  const layoutSh = ss.getSheetByName('配置表');
  if (!layoutSh) throw new Error('シート「配置表」が見つかりません。');

  const items = collectDropdownItemsFromColumn_(srcSh, 17, '作業マスタ!Q2以下に候補（空白以外）がありません。');

  // ====== 反映先1：配置表 J2:DA（2行目～最終行） ======
  {
    const dstStartRow = 2;
    const dstStartCol = 10; // J
    const dstEndCol = 105; // DA
    const dstLastRow = Math.max(layoutSh.getLastRow(), dstStartRow);
    const numRows = dstLastRow - dstStartRow + 1;
    const numCols = dstEndCol - dstStartCol + 1;
    const targetRange = layoutSh.getRange(dstStartRow, dstStartCol, numRows, numCols);

    applyDropdownAndColorRules_(layoutSh, targetRange, items, '=J2=');

    targetRange.setHorizontalAlignment('center');
    targetRange.setFontWeight('bold');
  }

  // ====== 反映先2：作業マスタ F2:F100 ======
  {
    const targetRange = srcSh.getRange('F2:F100');

    applyDropdownAndColorRules_(srcSh, targetRange, items, '=$F2=');

    targetRange.setHorizontalAlignment('center');
    targetRange.setFontWeight('bold');
  }

  // 初回設定時のみ、固定ルールを反映
  applyInitialLayoutConditionalFormatsOnce_();
}

/**
 * マスタ初回設定時のみ、配置表 K2:DB1000 に固定の条件付き書式を反映。
 * 2回目以降は DocumentProperties のフラグでスキップ。
 */
function applyInitialLayoutConditionalFormatsOnce_() {
  const propertyKey = 'MASTER_INITIAL_LAYOUT_CF_APPLIED_V1';
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty(propertyKey) === 'true') return;

  const ss = SpreadsheetApp.getActive();
  const layoutSh = ss.getSheetByName('配置表');
  if (!layoutSh) throw new Error('シート「配置表」が見つかりません。');

  const targetRange = layoutSh.getRange('K2:DB1000');
  removeConditionalRulesIntersecting_(layoutSh, targetRange);

  const fixedRules = [
    { formula: '=K2="PICK"', bg: '#ff0000', fg: '#000000' },
    { formula: '=K2="PACK"', bg: '#ffffff', fg: '#000000' },
    { formula: '=K2="RECEIVE"', bg: '#ffffff', fg: '#000000' },
    { formula: '=K2="STOW"', bg: '#ffffff', fg: '#000000' },
    { formula: '=K2="SHIPDOCK"', bg: '#ffffff', fg: '#000000' },
    { formula: '=DC2=0.25', bg: '#b7e1cd', fg: '#000000' }
  ];

  const newRules = fixedRules.map(rule =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(rule.formula)
      .setRanges([targetRange])
      .setBackground(rule.bg)
      .setFontColor(rule.fg)
      .build()
  );

  layoutSh.setConditionalFormatRules(layoutSh.getConditionalFormatRules().concat(newRules));
  props.setProperty(propertyKey, 'true');
}

/**
 * 指定シートの列（2行目以降）から、空白除外・重複除外（順序保持）で候補を収集。
 */
function collectDropdownItemsFromColumn_(sheet, columnNumber, emptyErrorMessage) {
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const srcRange = sheet.getRange(2, columnNumber, lastRow - 1, 1);
  const values = srcRange.getDisplayValues().flat();
  const bgs = srcRange.getBackgrounds().flat();
  const fgs = srcRange.getFontColors().flat();

  const seen = new Set();
  const items = [];
  for (let i = 0; i < values.length; i++) {
    const v = (values[i] || '').trim();
    if (!v || seen.has(v)) continue;
    seen.add(v);
    items.push({
      value: v,
      bg: bgs[i] || '#ffffff',
      fg: fgs[i] || '#000000'
    });
  }

  if (items.length === 0) throw new Error(emptyErrorMessage);
  return items;
}

/**
 * ドロップダウンと色連動ルールを対象範囲に適用。
 */
function applyDropdownAndColorRules_(sheet, targetRange, items, formulaPrefix) {
  const dv = SpreadsheetApp.newDataValidation()
    .requireValueInList(items.map(item => item.value), true)
    .setAllowInvalid(false)
    .build();

  targetRange.setDataValidation(dv);
  removeConditionalRulesIntersecting_(sheet, targetRange);

  const newRules = items.map(item =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`${formulaPrefix}${JSON.stringify(item.value)}`)
      .setRanges([targetRange])
      .setBackground(item.bg)
      .setFontColor(item.fg)
      .build()
  );

  sheet.setConditionalFormatRules(sheet.getConditionalFormatRules().concat(newRules));
}

/**
 * 指定rangeに交差する条件付き書式ルールを削除（他の場所のルールは残す）
 */
function removeConditionalRulesIntersecting_(sheet, targetRange) {
  const all = sheet.getConditionalFormatRules();
  const keep = [];

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

  for (const rule of all) {
    const ranges = rule.getRanges() || [];
    const hit = ranges.some(intersects);
    if (!hit) keep.push(rule);
  }
  sheet.setConditionalFormatRules(keep);
}
