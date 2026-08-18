/**
 * Config.gs
 * スクリプトプロパティの読み込みと、スプレッドシート/フォルダへのハンドル取得をまとめる。
 * ここに書かれているキーは全て「スクリプトプロパティ」に事前設定しておくこと。
 * （ファイル > プロジェクトの設定 > スクリプト プロパティ）
 */

// スクリプトプロパティのキー名（合意済み一覧）
const PROP_KEYS = {
  MAIN_SPREADSHEET_ID: 'MAIN_SPREADSHEET_ID',       // メインスプレッドシートのID
  SHEET_NAME_CLIENT_DB: 'SHEET_NAME_CLIENT_DB',     // 取引先DBシート名（期をまたいで固定）
  SHEET_NAME_STAFF_DB: 'SHEET_NAME_STAFF_DB',       // 社員DBシート名（期をまたいで固定）
  FOLDER_ID_QUOTE_ROOT: 'FOLDER_ID_QUOTE_ROOT',     // 見積書ルートフォルダID
  FOLDER_ID_INVOICE_ROOT: 'FOLDER_ID_INVOICE_ROOT', // 請求書ルートフォルダID
  FOLDER_ID_DELIVERY_ROOT: 'FOLDER_ID_DELIVERY_ROOT', // 納品書ルートフォルダID
  QUOTE_TEMPLATE_FILE_ID: 'QUOTE_TEMPLATE_FILE_ID', // 見積書テンプレート（スプレッドシート）
  INVOICE_TEMPLATE_FILE_ID: 'INVOICE_TEMPLATE_FILE_ID',
  DELIVERY_TEMPLATE_FILE_ID: 'DELIVERY_TEMPLATE_FILE_ID',
  COMPANY_SEAL_IMAGE_URL: 'COMPANY_SEAL_IMAGE_URL', // 承認後に貼り付ける社印画像（Drive直リンク等）
  CHATWORK_API_TOKEN: 'CHATWORK_API_TOKEN',         // Chatwork APIトークン
};

/**
 * 必須のスクリプトプロパティを取得する。無い場合は例外を投げる（設定漏れに早期に気づくため）。
 */
function getProp_(key) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) {
    throw new AppError_('CONFIG_MISSING', `スクリプトプロパティ「${key}」が設定されていません。`);
  }
  return value;
}

/**
 * メインスプレッドシートのハンドルを取得する。
 *
 * 重要: GAS のコンテナバインドスクリプトでは、サイドバー実行文脈でも
 * SpreadsheetApp.getActiveSpreadsheet() がバインド先スプレッドシートを返すため、
 * openById() との混在によるインスタンス不一致を避けるため、常に
 * SpreadsheetApp.getActiveSpreadsheet() を使用する。
 * MAIN_SPREADSHEET_ID プロパティが設定されていれば openById() も利用可能だが、
 * 同一ファイルへの2つのインスタンスで getActiveSheet() の結果が食い違う問題を
 * 回避するため、getActiveSpreadsheet() を優先する。
 */
function getMainSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/** 取引先DBシートを取得 */
function getClientDbSheet_() {
  return getMainSpreadsheet_().getSheetByName(getProp_(PROP_KEYS.SHEET_NAME_CLIENT_DB));
}

/** 社員DBシートを取得 */
function getStaffDbSheet_() {
  return getMainSpreadsheet_().getSheetByName(getProp_(PROP_KEYS.SHEET_NAME_STAFF_DB));
}

/** 各書類のルートフォルダを取得 */
function getRootFolder_(kind) {
  const map = {
    quote: PROP_KEYS.FOLDER_ID_QUOTE_ROOT,
    invoice: PROP_KEYS.FOLDER_ID_INVOICE_ROOT,
    delivery: PROP_KEYS.FOLDER_ID_DELIVERY_ROOT,
  };
  return DriveApp.getFolderById(getProp_(map[kind]));
}

/** 各書類のテンプレートファイルIDを取得 */
function getTemplateFileId_(kind) {
  const map = {
    quote: PROP_KEYS.QUOTE_TEMPLATE_FILE_ID,
    invoice: PROP_KEYS.INVOICE_TEMPLATE_FILE_ID,
    delivery: PROP_KEYS.DELIVERY_TEMPLATE_FILE_ID,
  };
  return getProp_(map[kind]);
}
