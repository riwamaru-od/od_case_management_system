/**
 * Constants.gs
 * シート列定義・ロール・ステータス・テンプレート転記先セルなど、
 * 「確定仕様」に基づく固定値をまとめる。
 * 列構成やセル番地の変更は、原則このファイルの修正のみで完結するようにしている。
 */

// ------------------------------------------------------------------
// 全案件DB / メイン画面UI 共通の列構成（1始まり）
// ※ 全案件DBとメイン画面UIは同一列構成（確定仕様2章）
// ------------------------------------------------------------------
const CASE_COLS = {
  CASE_NO: 1,            // 案件番号
  CLIENT_NAME: 2,        // 取引先名
  CASE_NAME: 3,          // 案件名
  END_SCHEDULE: 4,       // 終了予定（直近3ヶ月の前半/後半をプルダウン選択）
  BILLING_SCHEDULE: 5,   // 請求予定（同上）
  STAFF_IN_CHARGE: 6,    // 担当
  STATUS: 7,             // ステータス（案件/見積書/請求書の進行ステータス）
  BILLING_STATUS: 8,     // 請求ステータス（未請求 / 請求済み / 最終承認済み）
  QUOTE_LINK: 9,         // 見積書（スプレッドシートへのリンク）
  QUOTE_OUTPUT_LINK: 10, // 見積書(出力)（PDFへのリンク）
  INVOICE_LINK: 11,      // 請求書（スプレッドシートへのリンク）
  INVOICE_OUTPUT_LINK: 12, // 請求書(出力)（PDFへのリンク）
  DELIVERY_LINK: 13,     // 納品書（スプレッドシートへのリンク）
  DELIVERY_OUTPUT_LINK: 14, // 納品書(出力)（PDFへのリンク）
  QUOTE_CREATOR: 15,
  QUOTE_CREATED_AT: 16,
  QUOTE_APPROVER: 17,
  QUOTE_APPROVED_AT: 18,
  QUOTE_OUTPUT_BY: 19,
  QUOTE_OUTPUT_AT: 20,
  INVOICE_CREATOR: 21,
  INVOICE_CREATED_AT: 22,
  INVOICE_APPROVER: 23,
  INVOICE_APPROVED_AT: 24,
  INVOICE_OUTPUT_BY: 25,
  INVOICE_OUTPUT_AT: 26,
  DELIVERY_CREATOR: 27,
  DELIVERY_CREATED_AT: 28,
  FINAL_APPROVER: 29,
  FINAL_APPROVED_AT: 30,
  // 差し戻し日時（差し戻し後は空文字になるまで保持し、再作成ボタンの活性判定に使う内部フラグ）。
  // 列追加時は「xx期_表示」「xx期_全案件DB」「原本_表示」「原本_全案件DB」各シートの
  // ヘッダー行(1行目)に、AE列="見積書差し戻し日時"、AF列="請求書差し戻し日時" を運用者が追記すること。
  QUOTE_REJECTED_AT: 31,
  INVOICE_REJECTED_AT: 32,
};
const CASE_LAST_COL = 32;
const CASE_HEADER_ROW = 1;
const CASE_DATA_START_ROW = 2;

// ------------------------------------------------------------------
// 取引先DB（Googleフォームの新規取引先登録フォームから自動転記される）
// ------------------------------------------------------------------
const CLIENT_COLS = {
  OD_STAFF: 1,             // OD担当スタッフ
  COMPANY_NAME: 2,         // 社名
  CONTACT_NAME: 3,         // 担当者名
  DEPARTMENT: 4,           // 担当部署名
  POSTAL_CODE: 5,          // 郵便番号(半角数字)
  ADDRESS1: 6,             // 住所1
  ADDRESS2: 7,             // 住所2(建物名)
  CONTACT_PHONE: 8,        // 担当者電話番号
  CONTACT_EMAIL: 9,        // 担当者メールアドレス
  INVOICE_DELIVERY_METHOD: 10, // 請求書送付方法
  INVOICE_FORMAT_SPEC: 11, // 請求書書式指定の有無
  QUALIFIED_INVOICE_NO: 12, // 適格請求書発行事業者の登録番号
  INVOICE_CUTOFF_DAY: 13,  // 請求書締め日
  PAYMENT_MONTH: 14,       // 支払月
  PAYMENT_DAY: 15,         // 支払日
  NOTES: 16,               // 留意点
};

// ------------------------------------------------------------------
// 社員DB
// ------------------------------------------------------------------
const STAFF_COLS = {
  STAFF_NO: 1,
  STAFF_NAME: 2,
  EMAIL: 3,
  CHATWORK_ROOM_ID: 4,
  DEPARTMENT: 5,
  ROLE_1: 6,
  ROLE_2: 7,
  ROLE_3: 8,
  ROLE_4: 9,
  ROLE_5: 10,
  ROLE_6: 11,
};
const STAFF_ROLE_COLS = [STAFF_COLS.ROLE_1, STAFF_COLS.ROLE_2, STAFF_COLS.ROLE_3,
  STAFF_COLS.ROLE_4, STAFF_COLS.ROLE_5, STAFF_COLS.ROLE_6];

// ロール定義（社員DBのロール1〜6セルに入る値）
const ROLES = {
  ADMIN: 'Admin',
  STAFF: '一般社員',
  QUOTE_APPROVER: '見積書承認',
  INVOICE_APPROVER: '請求書承認',
  FINAL_APPROVER: '最終承認',
  ADMIN_DEPT: '総務',
};

// 各操作に必要なロール（いずれか1つ以上を保持していればOK）
// Admin・総務は全権限を持つため、すべての判定に含める。
const REQUIRED_ROLES = {
  APPROVE_QUOTE: [ROLES.ADMIN, ROLES.ADMIN_DEPT, ROLES.QUOTE_APPROVER],
  APPROVE_INVOICE: [ROLES.ADMIN, ROLES.ADMIN_DEPT, ROLES.INVOICE_APPROVER],
  FINAL_APPROVE: [ROLES.ADMIN, ROLES.ADMIN_DEPT, ROLES.FINAL_APPROVER],
};

// ------------------------------------------------------------------
// ステータス文言（表記はフローチャート準拠。差し戻し・再作成後の遷移先は
// 見積書／請求書それぞれの書類種別で正しく書き分ける - 確定仕様7章）
// ------------------------------------------------------------------
const STATUS = {
  CASE_REGISTERED: '案件登録済み',
  QUOTE_IN_PROGRESS: '見積書作成中',
  QUOTE_DRAFTED: '見積書承認待ち',
  QUOTE_APPROVED: '見積書承認済み',
  INVOICE_IN_PROGRESS: '請求書作成中',
  INVOICE_DRAFTED: '請求書承認待ち',
  INVOICE_APPROVED: '請求書承認済み',
  FINAL_APPROVED: '最終承認済み',
  // 案件中止（社員が途中の状態からでも中止できる。中止後は表示用シートから削除する）
  CANCELLED: '中止',
};

const BILLING_STATUS = {
  NOT_BILLED: '未請求',
  BILLED: '請求済み',
  FINAL_APPROVED: '最終承認済み',
};

// ------------------------------------------------------------------
// テンプレート転記先セル（各テンプレートはGoogleスプレッドシート）
// ------------------------------------------------------------------
const QUOTE_TEMPLATE_CELLS = {
  POSTAL_CODE: 'B5',
  ADDRESS1: 'B6',
  ADDRESS2: 'B7',
  COMPANY_NAME: 'B8',
  DEPARTMENT: 'B10',
  CONTACT_NAME: 'B11',      // 末尾に「様」を付けて記載
  CASE_NO: 'G1',
  SERIAL_NO: 'G6',          // この期の見積書通し番号（再作成時も新規採番）
  SEAL_IMAGE_RANGE: 'G8:G13', // 承認後、社印画像を挿入
  STAFF_NAME: 'G12',
  STAFF_EMAIL: 'G13',
  SUBJECT: 'C16',
  INVOICE_CUTOFF_DAY: 'O7',
  PAYMENT_MONTH_DAY_TEXT: 'O8', // 支払月と支払日を統合したテキスト
  INVOICE_DELIVERY_METHOD: 'O10',
  INVOICE_FORMAT_SPEC: 'O11',
  CLIENT_CONTACT_EMAIL: 'O13',
  CLIENT_CONTACT_PHONE: 'O14',
  CREATOR_NAME: 'K56',
  CREATED_AT: 'L56',
  REQUEST_COMMENT: 'N56',   // 承認依頼時コメント
  APPROVER_NAME: 'K57',
  APPROVED_AT: 'L57',
  APPROVE_COMMENT: 'N57',   // 承認時コメント
  PDF_OUTPUT_BY: 'K58',
  PDF_OUTPUT_AT: 'L58',
  BODY_COPY_RANGE_FOR_INVOICE: ['B22:E47', 'E49', 'F50', 'F53'], // 請求書へ転記する範囲
};

const INVOICE_TEMPLATE_CELLS = {
  POSTAL_CODE: 'B5',
  ADDRESS1: 'B6',
  ADDRESS2: 'B7',
  COMPANY_NAME: 'B8',
  DEPARTMENT: 'B10',
  CONTACT_NAME: 'B11',
  CASE_NO: 'G1',
  SERIAL_NO: 'G6',          // この期の請求書通し番号（再作成時も新規採番）
  SEAL_IMAGE_RANGE: 'G8:G13',
  STAFF_NAME: 'G12',
  STAFF_EMAIL: 'G13',
  SUBJECT: 'C16',
  QUOTE_SERIAL_REF: 'G53',  // 請求書作成時点の最新の見積書通し番号を「見積書No.xxx」の形式で自動記載
  CREATOR_NAME: 'C56',
  CREATED_AT: 'D56',
  REQUEST_COMMENT: 'F56',
  APPROVER_NAME: 'C57',
  APPROVED_AT: 'D57',
  APPROVE_COMMENT: 'F57',
  PDF_OUTPUT_BY: 'C58',
  PDF_OUTPUT_AT: 'D58',
  BODY_COPY_RANGE_FOR_DELIVERY: ['B22:G47', 'F48:F54', 'E49'], // 納品書へ転記する範囲
};

const DELIVERY_TEMPLATE_CELLS = {
  POSTAL_CODE: 'B5',
  ADDRESS1: 'B6',
  ADDRESS2: 'B7',
  COMPANY_NAME: 'B8',
  DEPARTMENT: 'B10',
  CONTACT_NAME: 'B11',
  CASE_NO: 'G1',
  SERIAL_NO: 'G6',          // この期の納品書通し番号（再作成時も新規採番）
  SEAL_IMAGE_RANGE: 'G8:G13',
  STAFF_NAME: 'G12',
  STAFF_EMAIL: 'G13',
  SUBJECT: 'C16',
  QUOTE_SERIAL_REF: 'G52',   // 納品書作成時点の最新の見積書通し番号を「見積書No.xxx」の形式で自動記載
  INVOICE_SERIAL_REF: 'G53', // 納品書作成時点の最新の請求書通し番号を「請求書No.xxx」の形式で自動記載
  CREATOR_NAME: 'C56',
  CREATED_AT: 'D56',
  PDF_OUTPUT_BY: 'C57',
  PDF_OUTPUT_AT: 'D57',
};

// ------------------------------------------------------------------
// 社印画像のサイズ（px）。insertSealImage_ が挿入後の画像に明示的に設定する。
// 見積書・請求書・納品書で共通のサイズを使う。サイズを変更したい場合はここを直接編集する。
// ------------------------------------------------------------------
const SEAL_IMAGE_WIDTH_PX = 70;
const SEAL_IMAGE_HEIGHT_PX = 70;

// ------------------------------------------------------------------
// PDF出力時のページ範囲（セル指定）。書類種別ごとに「何ページ目に何のセル範囲を
// 印刷するか」をA1形式で定義する。
// 注意: Google スプレッドシートのPDFエクスポートAPIは、1回のリクエストで
// 複数の飛び地レンジを直接指定することができないため、buildPdfExportUrl_ 側で
// これらのページ範囲すべてを包含する最小の矩形（バウンディングボックス）を
// エクスポート範囲として指定し、拡大縮小なし（実寸）・ページ順序「右へ、それから下へ」
// で出力することで、Google側の自動改ページに委ねて実現している。
// そのため、各テンプレートの列幅・行高が、ここで指定した列・行の境界でちょうど
// 1ページ分の印刷可能領域（A4・余白0.5inch）に収まっている必要がある。
// 境界がずれる場合は、各テンプレート側で「ファイル > 印刷設定」から明示的に
// ページ区切りを設定して調整すること。
// ------------------------------------------------------------------
const PDF_PAGE_RANGES = {
  quote: ['A1:H54', 'I1:P54', 'I55:P58'],
  invoice: ['A1:H54', 'A55:H58'],
  delivery: ['A1:H54', 'A55:H58'],
};

// ------------------------------------------------------------------
// 期の起点定義（確定仕様2章: 5/1切り替え。2026/7時点で17期）
// ------------------------------------------------------------------
const PERIOD_BASE_NUMBER = 17;
const PERIOD_BASE_YEAR = 2026; // この年の5/1開始分が17期
const PERIOD_SWITCH_MONTH_INDEX = 4; // 0始まりで4 = 5月
const PERIOD_SWITCH_DAY = 1;
const PERIOD_PREP_MONTH_OFFSET = 1; // 切り替えの1ヶ月前(4/1)に翌期の準備をする

// 原本（マスター）シート名。期の切替時にこれをコピーして新シートを作る。
const MASTER_UI_SHEET_NAME = '原本_表示';
const MASTER_DB_SHEET_NAME = '原本_全案件DB';

// フォルダ配下のサブフォルダ名
const SUBFOLDER = {
  UNBILLED: '未請求案件',   // 見積書・請求書の作成〜未請求の間
  BILLING: '請求中案件',    // 請求書承認済み以降（PDF出力・印刷対象）
};

// ------------------------------------------------------------------
// 操作ログ（確定仕様5.3節）: 期をまたいで固定の1枚のシートに、
// 全操作を実行日時・実行者・案件番号・操作種別・結果の形で追記する。
// ------------------------------------------------------------------
const OPERATION_LOG_SHEET_NAME = '操作ログ';
