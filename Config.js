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
  // 過去期のアーカイブファイル（{n}期_案件データ）の保存先フォルダID。
  // メインスプレッドシートが肥大化しないよう、期が切り替わったら前期の
  // 表示シート・全案件DBシートをこのフォルダ内の別ファイルへ退避する。
  FOLDER_ID_ARCHIVE_ROOT: 'FOLDER_ID_ARCHIVE_ROOT',
  // 全案件DB・取引先DBの定期バックアップの保存先フォルダID。
  FOLDER_ID_BACKUP_ROOT: 'FOLDER_ID_BACKUP_ROOT',
  QUOTE_TEMPLATE_FILE_ID: 'QUOTE_TEMPLATE_FILE_ID', // 見積書テンプレート（スプレッドシート）
  INVOICE_TEMPLATE_FILE_ID: 'INVOICE_TEMPLATE_FILE_ID',
  DELIVERY_TEMPLATE_FILE_ID: 'DELIVERY_TEMPLATE_FILE_ID',
  // 承認後（納品書は作成時）に貼り付ける社印画像。Googleドライブの通常の共有リンク
  // （https://drive.google.com/file/d/xxxxx/view?usp=sharing 形式）をそのまま設定してよい。
  // スクリプトがファイルIDを抽出してDriveApp経由で画像を取得するため、一般公開設定は不要
  // （実行者がそのファイルを閲覧できる権限を持っていればよい）。
  COMPANY_SEAL_IMAGE_URL: 'COMPANY_SEAL_IMAGE_URL',
  CHATWORK_API_TOKEN: 'CHATWORK_API_TOKEN',         // Chatwork APIトークン
  // 管理用Googleアカウントのメールアドレス。以下2つの役割を持つ。
  //  (1) GASのインストール型トリガー（installTriggers()）を設定するアカウント
  //  (2) 書類ファイルを作成するアカウント（＝生成される全書類のオーナー）
  // Google Workspace を使わない環境では setOwner() によるオーナー移譲が
  // 招待制となり機能しないため、「書類の作成操作は必ずこのアカウントで行う」ことで
  // 全書類のオーナーを管理用アカウントに統一する運用とする。
  // 別のアカウントで書類が作成された場合は、操作ログにエラーとして記録される
  // （TemplateFillService.gs の verifyAdminAccountExecution_ 参照）。
  ADMIN_TRIGGER_ACCOUNT_EMAIL: 'ADMIN_TRIGGER_ACCOUNT_EMAIL',
  // Webアプリとしてデプロイした際のURL（.../exec 形式）。callAsAdmin_（AdminProxyService.gs）が
  // 書き込み系操作を管理用アカウント権限で実行するために参照する。
  // デプロイIDが変わるたびに更新が必要（クイックデプロイ「テストとして実行」のURLではなく、
  // 通常のデプロイで発行される安定したURLを設定すること）。
  WEBAPP_URL: 'WEBAPP_URL',
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

/**
 * IDでDriveフォルダを取得する。アクセスできない場合は、原因（フォルダの共有設定が
 * 足りない可能性が高い）が分かるメッセージで例外を投げる。
 * 一般社員が書類を作成する際も、このフォルダを閲覧・編集できる必要があるため、
 * 初期設定時に対象フォルダを全利用者へ共有しておくこと（初期設定マニュアル参照）。
 */
function getFolderByIdSafe_(id, label) {
  try {
    return DriveApp.getFolderById(id);
  } catch (e) {
    throw AppError_('DRIVE_ACCESS_DENIED',
      `${label}（フォルダID: ${id}）にアクセスできません。実行アカウントがこのフォルダを`
      + `閲覧・編集できる権限を持っているか確認してください。管理用アカウントのみに`
      + `共有されたままで、一般社員へ共有されていない可能性があります。`
      + `詳細: ${e && e.message ? e.message : e}`);
  }
}

/** IDでDriveファイルを取得する。アクセスできない場合の挙動は getFolderByIdSafe_ と同様。 */
function getFileByIdSafe_(id, label) {
  try {
    return DriveApp.getFileById(id);
  } catch (e) {
    throw AppError_('DRIVE_ACCESS_DENIED',
      `${label}（ファイルID: ${id}）にアクセスできません。実行アカウントがこのファイルを`
      + `閲覧できる権限を持っているか確認してください。管理用アカウントのみに`
      + `共有されたままで、一般社員へ共有されていない可能性があります。`
      + `詳細: ${e && e.message ? e.message : e}`);
  }
}

/** 各書類のルートフォルダを取得 */
function getRootFolder_(kind) {
  const map = {
    quote: { key: PROP_KEYS.FOLDER_ID_QUOTE_ROOT, label: '見積書ルートフォルダ' },
    invoice: { key: PROP_KEYS.FOLDER_ID_INVOICE_ROOT, label: '請求書ルートフォルダ' },
    delivery: { key: PROP_KEYS.FOLDER_ID_DELIVERY_ROOT, label: '納品書ルートフォルダ' },
  };
  return getFolderByIdSafe_(getProp_(map[kind].key), map[kind].label);
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

/** 管理用アカウント（GASインストール型トリガー設定アカウント）のメールアドレスを取得 */
function getAdminTriggerAccountEmail_() {
  return getProp_(PROP_KEYS.ADMIN_TRIGGER_ACCOUNT_EMAIL);
}

/** 過去期アーカイブの保存先フォルダを取得 */
function getArchiveFolder_() {
  return getFolderByIdSafe_(getProp_(PROP_KEYS.FOLDER_ID_ARCHIVE_ROOT), 'アーカイブフォルダ');
}

/** バックアップの保存先フォルダを取得 */
function getBackupFolder_() {
  return getFolderByIdSafe_(getProp_(PROP_KEYS.FOLDER_ID_BACKUP_ROOT), 'バックアップフォルダ');
}
