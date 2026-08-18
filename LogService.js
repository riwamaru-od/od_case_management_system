/**
 * LogService.gs
 * 全操作を時系列に追える「操作ログ」シート（確定仕様5.3節）への記録。
 * 期をまたいで失われないよう、期別シートとは別にメインスプレッドシート直下に
 * 固定で1枚だけ存在する永続シートとする（無ければ自動作成する）。
 */

/** 操作ログシートを取得する（無ければヘッダー付きで新規作成） */
function getOrCreateOperationLogSheet_() {
  const ss = getMainSpreadsheet_();
  let sheet = ss.getSheetByName(OPERATION_LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(OPERATION_LOG_SHEET_NAME);
    sheet.appendRow(['実行日時', '実行者(氏名)', '実行者(メール)', '案件番号', '操作種別', '結果', '詳細']);
  }
  return sheet;
}

/**
 * 操作ログを1行追記する。呼び出し元の withLock_ による排他制御の中から呼ぶこと。
 * @param {string} caseNo 案件番号（案件に紐付かない操作の場合は空文字）
 * @param {string} actionLabel 操作種別（例:「見積書作成」「最終承認」）
 * @param {string} detail 詳細（コメントやURLなど。無ければ空文字）
 * @param {boolean} isError true の場合は結果欄に「エラー」と記録する
 */
function appendOperationLog_(caseNo, actionLabel, detail, isError) {
  const sheet = getOrCreateOperationLogSheet_();
  let email = '';
  let name = 'システム(自動実行)';
  try {
    email = getActiveUserEmail_();
    const staff = findStaffByEmail_(email);
    name = staff ? staff.name : email;
  } catch (e) {
    email = '';
  }
  sheet.appendRow([
    formatDateTime_(new Date()),
    name,
    email,
    caseNo || '',
    actionLabel,
    isError ? 'エラー' : '成功',
    detail || '',
  ]);
}
