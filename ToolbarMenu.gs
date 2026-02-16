/**
 * スプレッドシート上部メニュー（ツールバー）関連。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚚 配置システム')
    .addItem('配置反映 (色付きセル)', 'showConfirmDialog')
    .addSeparator()
    .addItem('配置管理パネルを開く', 'showAdminPanel')
    .addToUi();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function showAdminPanel() {
  const html = HtmlService.createTemplateFromFile('index').evaluate()
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
