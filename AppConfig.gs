/**
 * アプリ全体で使用する定数定義。
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
