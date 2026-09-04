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
  // 再承認待ちフラグ（差し戻し・再作成した時点で日時が入り、再承認されると空になる）。
  // 承認記録（承認者・承認日時）は履歴として残したまま、「今の版はまだ承認されていない」
  // ことを表すために独立して持つ。PDF出力ボタンの活性判定に使う。
  QUOTE_REAPPROVAL_PENDING: 33,
  INVOICE_REAPPROVAL_PENDING: 34,
  // 着手日時（書類ファイルを作成・再作成した時点の日時）。
  // 作成日時（QUOTE_CREATED_AT 等）は「完成」ボタン時に記録されるため、
  // 「作成中のまま放置されている書類」の経過日数はこちらで判定する。
  QUOTE_STARTED_AT: 35,
  INVOICE_STARTED_AT: 36,
};
const CASE_LAST_COL = 36;
const CASE_HEADER_ROW = 1; // 全案件DBシート用（見出しは1行目のみ）
const CASE_DATA_START_ROW = 2; // 全案件DBシート用（見出しは1行目のみ、データは2行目から）

// 表示シートは見出しが1〜2行目の2行にまたがっており、実際の案件データは3行目から始まる
// （全案件DBシートとは見出しの行数が異なる）。表示シートの案件データ範囲を明示的に
// 扱う処理（範囲一括クリア・件数集計など）はこちらを使うこと。
// SCHEDULE_VALIDATION_START_ROW（下記）も同じ理由で3になっている。
const CASE_UI_DATA_START_ROW = 3;

// 表示シートで、列ごとの見出し（CASE_HEADERS）が入るのは2行目。
// 1行目は見出しの1段目（項目のグループ名など）で、システムからは書き換えない。
// メニュー「案件シートの構成を確認・修復する」はこの行を対象に不足分を補う。
const CASE_UI_HEADER_ROW = 2;

// ヘッダー行（1行目）の見出し。CASE_COLS と同じ並び順で定義すること。
// メニュー「案件シートの構成を確認・修復する」（SetupService.gs）が、
// 列数の不足と未設定の見出しを補うために参照する。
const CASE_HEADERS = [
  '案件番号', '取引先名', '案件名', '終了予定', '請求予定', '担当',
  'ステータス', '請求ステータス',
  '見積書', '見積書(出力)', '請求書', '請求書(出力)', '納品書', '納品書(出力)',
  '見積書作成者', '見積書作成日時', '見積書承認者', '見積書承認日時', '見積書出力者', '見積書出力日時',
  '請求書作成者', '請求書作成日時', '請求書承認者', '請求書承認日時', '請求書出力者', '請求書出力日時',
  '納品書作成者', '納品書作成日時', '最終承認者', '最終承認日時',
  '見積書差し戻し日時', '請求書差し戻し日時',
  '見積書再承認待ち', '請求書再承認待ち',
  '見積書着手日時', '請求書着手日時',
];

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
  // 承認後に社印画像を挿入する基準セル。横位置はこのセルの左上から
  // SEAL_IMAGE_OFFSET_X_PX 分ずらし、縦位置は SEAL_IMAGE_BOTTOM_ROW の下端に
  // 画像の下端が揃うよう自動計算する
  SEAL_IMAGE_RANGE: 'G4',
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
  SEAL_IMAGE_RANGE: 'G4', // 社印を挿入する基準セル（見積書と同じ扱い）
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
  SEAL_IMAGE_RANGE: 'G4', // 社印を挿入する基準セル（見積書と同じ扱い）
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

// 見積書は承認するとシート全体を保護（編集不可に）するが、この範囲だけは保護しない。
// 社内の作業用エリアであり、承認後も記入・修正を続ける必要があるため。
// PDF出力範囲（A1:H54）の外側にあるため、取引先へ渡すPDFには影響しない。
const QUOTE_UNPROTECTED_RANGES_AFTER_APPROVAL = ['J21:O54'];

// ------------------------------------------------------------------
// 見積書と請求書の内容差分チェック
//
// 請求書は見積書から下記の範囲を「同一番地へ」コピーして作成される。作成後に
// 請求書側だけが編集されて見積書と食い違うことを検知するため、請求書ファイル内に
// 非表示の比較用シートを作り、作成時点の見積書の値を保持しておく。請求書本体には
// 条件付き書式を設定し、比較用シートと値が異なるセルを自動で黄色にする。
// （条件付き書式はスプレッドシートの機能なので、入力と同時に色が付き、
//   見積書と同じ内容に直せば即座に色が消える＝リアルタイム判定になる。
//   書類ファイルごとにスクリプトのトリガーを設置する方式は、1ユーザーあたりの
//   トリガー数上限に抵触するため採用していない）
// ------------------------------------------------------------------
const QUOTE_INVOICE_DIFF_SHEET_NAME = '_見積比較用';
const QUOTE_INVOICE_DIFF_RANGES = [
  'B5:B8',   // 郵便番号・住所1・住所2・会社名（取引先DB由来）
  'B10:B11', // 部署・担当者名（取引先DB由来）
  'G1',      // 案件番号
  'C16',     // 件名（案件名）
  'G12:G13', // 担当社員の氏名・メールアドレス（社員DB由来）
  'B22:E47', // 明細（見積書から転記）
  'E49', 'F50', 'F53',
];
// 差分があるセルの背景色。請求書独自の項目（G6=請求書通し番号、G53=見積書No.参照）は
// 見積書と一致しなくて当然のため、上記 QUOTE_INVOICE_DIFF_RANGES には含めていない。
const QUOTE_INVOICE_DIFF_COLOR = '#FFF2A8';

// ------------------------------------------------------------------
// 社印画像の表示サイズ（px）。insertSealImage_ が挿入後の画像に明示的に設定する。
// 見積書・請求書・納品書で共通。サイズを変更したい場合はこの2つの値を編集する。
// ------------------------------------------------------------------
const SEAL_IMAGE_WIDTH_PX = 150;
const SEAL_IMAGE_HEIGHT_PX = 150;

// 社印の横位置の微調整（px）。各テンプレートの *_TEMPLATE_CELLS.SEAL_IMAGE_RANGE で
// 指定したセルの左上を基準に、この分だけ右へずらして配置する。
// 「もう少し右へ」なら値を増やす（負の値も指定可）。
const SEAL_IMAGE_OFFSET_X_PX = 60;

// 社印の下端を合わせる行。縦位置はこの行の下端に画像の下端が揃うよう、
// 挿入時に実際の行の高さから自動計算する（calcSealImageOffsetY_ 参照）。
// 「もう1行下げたい」ならこの値を増やす。
const SEAL_IMAGE_BOTTOM_ROW = 13;

// Sheet.insertImage には「2MB以下・100万画素以下」という上限があり、高解像度の
// 社印画像はそのままでは挿入できない。その場合に使うDrive生成の縮小版の最大辺（px）。
// 表示サイズ（上記150px角）に対して十分な解像度を確保しつつ、上限に掛からない値にする。
const SEAL_THUMBNAIL_MAX_PX = 600;

// ------------------------------------------------------------------
// PDF出力時のページ範囲（セル指定）。書類種別ごとに「何ページ目に何のセル範囲を
// 印刷するか」をA1形式で定義する。
//
// 注意: Googleスプレッドシートには（Excelと異なり）手動でページ区切りを挿入する
// 機能が無く、またPDFエクスポートAPIも1回のリクエストで複数の飛び地セル範囲を
// 別々のページとして書き出すことができない。そのため exportRangesAsPagedPdfBlob_
// （DocumentService.js）は、一時スプレッドシートに「1ページ＝1シート」を作り
// （範囲外の行・列は非表示にする）、ファイル全体をPDF化することで実現している
// （スプレッドシート全体をPDF出力すると各シートが必ず新しいページから始まる仕様を利用）。
//
// 全書類とも本文の1ページのみを出力する。56行目以降の操作履歴欄（作成者・承認者・
// 出力者と各日時）と、見積書のI〜P列にある請求書発行情報は社内管理用のため、
// 取引先へ渡すPDFには含めない。
// ------------------------------------------------------------------
const PDF_PAGE_RANGES = {
  quote: ['A1:H54'],
  invoice: ['A1:H54'],
  delivery: ['A1:H54'],
};

// ------------------------------------------------------------------
// 終了予定・請求予定（D列/E列）のプルダウン選択肢。
// 「config」シートに自動生成される（ScheduleOptionsService.gs参照）。
// 「YYYY/MM 前半」「YYYY/MM 後半」の月範囲（今月を含め、何ヶ月前〜何ヶ月後まで）と、
// 固定の特殊選択肢を定義する。
// ------------------------------------------------------------------
const CONFIG_SHEET_NAME = 'config';
// 案件シート（表示・全案件DB）側で、終了予定・請求予定のプルダウン（データ入力規則）を
// 適用する開始行。案件データ自体はCASE_DATA_START_ROW（2行目）から始まるが、
// プルダウンの適用範囲だけは運用上3行目からにする。
const SCHEDULE_VALIDATION_START_ROW = 3;
const SCHEDULE_OPTIONS_MONTHS_BEFORE = 1; // 今月の何ヶ月前から選択肢に含めるか
const SCHEDULE_OPTIONS_MONTHS_AFTER = 6;  // 今月の何ヶ月後まで選択肢に含めるか
const SCHEDULE_OPTIONS_FIXED_CHOICES = ['未定', '営業案件']; // 月に紐づかない固定の選択肢

// 請求予定（E列）が今月分（前半・後半）に一致する行の文字色。今月中に請求すべき案件が
// 一目でわかるよう、条件付き書式で自動的に色付けする（ScheduleOptionsService.gs参照）。
const CURRENT_MONTH_BILLING_TEXT_COLOR = '#1155CC';

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
  CANCELLED: '中止案件',    // 案件中止時、全書類の案件フォルダをここへ移動する
};

// ------------------------------------------------------------------
// 操作ログ（確定仕様5.3節）: 期をまたいで固定の1枚のシートに、
// 全操作を実行日時・実行者・案件番号・操作種別・結果の形で追記する。
// ------------------------------------------------------------------
// シートを直接編集した人を特定できなかった場合に、操作ログの実行者欄へ記録する名称。
// この環境（実Google Workspaceドメインなし）では、インストール型onEditトリガーの
// e.user も Session.getActiveUser() も空になり編集者を特定できないため、
// 自動処理と区別できるようにこの名称を使う。
const SHEET_EDITOR_UNKNOWN_ACTOR_NAME = 'シート直接編集（実行者不明）';

const OPERATION_LOG_SHEET_NAME = '操作ログ';

// ------------------------------------------------------------------
// アーカイブ / バックアップ
// メインスプレッドシートに期ごとのシートを積み重ね続けると動作が重くなるため、
// 期が切り替わったら前期分を別ファイルへ退避し、本体からは削除する（ArchiveService.gs）。
// あわせて、全案件DB・取引先DBの日次バックアップを別ファイルとして保存する（BackupService.gs）。
// ------------------------------------------------------------------
const ARCHIVE_FILE_NAME_PREFIX = '案件データ_';   // 例: 案件データ_17期
const BACKUP_FILE_NAME_PREFIX = 'バックアップ_';  // 例: バックアップ_20260818
const BACKUP_GENERATIONS = 30;                    // 保持する世代数。超過分は古いものから自動削除

/** 期別シート名（{n}期_表示 / {n}期_全案件DB）を判定するパターン */
const PERIOD_SHEET_NAME_PATTERN = /^(\d+)期_(表示|全案件DB)$/;
