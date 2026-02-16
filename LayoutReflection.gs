/**
 * 配置表への反映処理（OCR・一括貼り付け）
 */
function buildWorkNameToProcessMap_() {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName('作業マスタ');
  if (!master) throw new Error('作業マスタ シートが見つかりません');

  const lastRow = Math.max(master.getLastRow(), 2);
  const keyRange = master.getRange(2, 2, lastRow - 1, 1).getDisplayValues();
  const valRange = master.getRange(2, 6, lastRow - 1, 1).getDisplayValues();

  const map = {};
  for (let i = 0; i < keyRange.length; i++) {
    const k = normalizeKey_(keyRange[i][0]);
    const v = (valRange[i][0] || '').trim();
    if (!k || !v) continue;
    map[k] = v;
  }
  return map;
}

function normalizeKey_(s) {
  if (s === null || s === undefined) return '';
  return String(s)
    .replace(/\u3000/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
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
  pasteToGrayCellsByDB_multiRows(sh);
}

function pasteToGrayCellsByDB_multiRows(sh) {
  const sRow = CONFIG.PASTE_SETTINGS.START_ROW;
  const nRows = CONFIG.PASTE_SETTINGS.NUM_ROWS;
  const sourceCol = 9;
  const tStart = 10;
  const tEnd = 105;
  const nCols = tEnd - tStart + 1;
  const dbStart = 106;

  const actualNumRows = Math.min(nRows, sh.getMaxRows() - sRow + 1);
  if (actualNumRows <= 0) return;

  const map = buildWorkNameToProcessMap_();
  const keys = sh.getRange(sRow, sourceCol, actualNumRows, 1).getDisplayValues();
  const dbValues = sh.getRange(sRow, dbStart, actualNumRows, nCols).getValues();

  const targetRange = sh.getRange(sRow, tStart, actualNumRows, nCols);
  const targetValues = targetRange.getValues();

  let totalChanged = 0;
  for (let r = 0; r < actualNumRows; r++) {
    const key = normalizeKey_(keys[r][0]);
    const newValue = (key && map[key]) ? map[key] : '';

    for (let c = 0; c < nCols; c++) {
      const active = Number(dbValues[r][c] || 0);
      targetValues[r][c] = active > 0 ? newValue : '';
      totalChanged++;
    }
  }

  targetRange.setValues(targetValues);
  SpreadsheetApp.getActive().toast(`反映完了: ${totalChanged}セル`, '完了');
}
