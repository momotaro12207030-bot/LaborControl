/**
 * 作業マスタ!G2:G の非空セルを候補にして、作業マスタ!C2:C にプルダウンを一括設定。
 * さらに G列の各セルの書式（背景色/文字色）を元に、C列が選択値に応じて同じ色になるよう
 * 条件付き書式を自動生成する。
 */
function setupAreaDropdownForColumnC() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("作業マスタ");
  if (!sh) throw new Error("シート「作業マスタ」が見つかりません。");

  // --- 候補の元：G2:G（空白除外、重複除外、順序保持） ---
  const lastRow = Math.max(sh.getLastRow(), 2);
  const srcRange = sh.getRange(2, 7, lastRow - 1, 1); // G2:G
  const values = srcRange.getDisplayValues().flat();
  const bgs = srcRange.getBackgrounds().flat();
  const fgs = srcRange.getFontColors().flat();

  const seen = new Set();
  const items = [];
  for (let i = 0; i < values.length; i++) {
    const v = (values[i] || "").trim();
    if (!v) continue;
    if (seen.has(v)) continue;
    seen.add(v);
    items.push({
      value: v,
      bg: bgs[i] || "#ffffff",
      fg: fgs[i] || "#000000",
    });
  }
  if (items.length === 0) throw new Error("作業マスタ!G2以下に候補（空白以外）がありません。");

  // --- 適用先：C2:C（表エリアがどこまでか不明なので最終行まで） ---
  // 必要ならここを固定（例：200行）にしてもOK
  const applyRows = Math.max(lastRow - 1, 1);
  const targetRange = sh.getRange(2, 3, applyRows, 1); // C2:C

  // 1) C列にプルダウン（データ検証）を一括設定
  const dv = SpreadsheetApp.newDataValidation()
    .requireValueInList(items.map(x => x.value), true) // true: ドロップダウン表示
    .setAllowInvalid(false)
    .build();
  targetRange.setDataValidation(dv);

  // 2) C列範囲に当たっている条件付き書式を除去してから、新しく作る
  removeConditionalRulesIntersecting_(sh, targetRange);

  // 3) 選択値に応じて色を変える条件付き書式（C2基準の相対参照）
  //    ※範囲が C2:C なので、数式は =$C2="xxx" でOK（行は自動で相対になります）
  const newRules = [];
  for (const it of items) {
    const formula = `=$C2=${JSON.stringify(it.value)}`;
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setRanges([targetRange])
      .setBackground(it.bg)
      .setFontColor(it.fg)
      .build();
    newRules.push(rule);
  }

  sh.setConditionalFormatRules(sh.getConditionalFormatRules().concat(newRules));

  // ちょい見た目（任意）
  targetRange.setHorizontalAlignment("center");
  targetRange.setFontWeight("bold");
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

/**
 * 作業マスタ!Q2:Q の非空セルを候補にして、
 * 1) 配置表!J2:DA（2行目～最終行）にドロップダウン＋色連動
 * 2) 作業マスタ!F2:F100 にも同じドロップダウン＋色連動
 *
 * ※条件付き書式は対象範囲に交差するルールを一度除去してから付与。
 */
function applyWorkDropdown_Q_to_Layout_and_WorkMasterF() {
  const ss = SpreadsheetApp.getActive();

  const srcSh = ss.getSheetByName("作業マスタ");
  if (!srcSh) throw new Error("シート「作業マスタ」が見つかりません。");

  const layoutSh = ss.getSheetByName("配置表");
  if (!layoutSh) throw new Error("シート「配置表」が見つかりません。");

  // ---- 候補の元：作業マスタ!Q2:Q（空白除外・重複除外・順序保持） ----
  const srcLastRow = Math.max(srcSh.getLastRow(), 2);
  const srcRange = srcSh.getRange(2, 17, srcLastRow - 1, 1); // Q列=17
  const vals = srcRange.getDisplayValues().flat();
  const bgs  = srcRange.getBackgrounds().flat();
  const fgs  = srcRange.getFontColors().flat();

  const seen = new Set();
  const items = [];
  for (let i = 0; i < vals.length; i++) {
    const v = (vals[i] || "").trim();
    if (!v) continue;
    if (seen.has(v)) continue;
    seen.add(v);
    items.push({
      value: v,
      bg: bgs[i] || "#ffffff",
      fg: fgs[i] || "#000000",
    });
  }
  if (items.length === 0) throw new Error("作業マスタ!Q2以下に候補（空白以外）がありません。");

  // ---- データ検証（ドロップダウン）ルール ----
  const dv = SpreadsheetApp.newDataValidation()
    .requireValueInList(items.map(x => x.value), true)
    .setAllowInvalid(false)
    .build();

  // ====== 反映先1：配置表 J2:DA（2行目～最終行） ======
  {
    const dstStartRow = 2;
    const dstStartCol = 10; // J
    const dstEndCol = 105;  // DA
    const dstLastRow = Math.max(layoutSh.getLastRow(), dstStartRow);
    const numRows = dstLastRow - dstStartRow + 1;
    const numCols = dstEndCol - dstStartCol + 1;

    const targetRange = layoutSh.getRange(dstStartRow, dstStartCol, numRows, numCols);

    targetRange.setDataValidation(dv);
    removeConditionalRulesIntersecting_(layoutSh, targetRange);

    const newRules = [];
    for (const it of items) {
      // 先頭セル J2 を相対参照として使用（範囲内で各セルに自動追従）
      const formula = `=J2=${JSON.stringify(it.value)}`;
      newRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(formula)
          .setRanges([targetRange])
          .setBackground(it.bg)
          .setFontColor(it.fg)
          .build()
      );
    }
    layoutSh.setConditionalFormatRules(layoutSh.getConditionalFormatRules().concat(newRules));

    targetRange.setHorizontalAlignment("center");
    targetRange.setFontWeight("bold");
  }

  // ====== 反映先2：作業マスタ F2:F100 ======
  {
    const targetRange = srcSh.getRange("F2:F100");

    targetRange.setDataValidation(dv);
    removeConditionalRulesIntersecting_(srcSh, targetRange);

    const newRules = [];
    for (const it of items) {
      // 先頭セル F2 を相対参照（範囲内で行に追従）
      const formula = `=$F2=${JSON.stringify(it.value)}`;
      newRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(formula)
          .setRanges([targetRange])
          .setBackground(it.bg)
          .setFontColor(it.fg)
          .build()
      );
    }
    srcSh.setConditionalFormatRules(srcSh.getConditionalFormatRules().concat(newRules));

    targetRange.setHorizontalAlignment("center");
    targetRange.setFontWeight("bold");
  }
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
