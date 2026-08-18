/**
 * Triggers.gs
 * onOpen / onEdit（インストール型） / 時間主導型トリガーの入り口をまとめる。
 *
 * セットアップ手順:
 *   1) この関数をエディタから1回だけ手動実行する: installTriggers()
 *      → onEdit（インストール型）と、毎日0:30に走る dailyScheduledTasks() が登録される。
 *   2) 単純トリガー onEdit(e) は権限の都合上フル機能を実行できないため、
 *      本プロジェクトでは自動採番などの副作用を伴う処理は全てインストール型トリガー
 *      （onEditInstallable）側で行う。単純トリガーの onEdit(e) は何もしない。
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('案件管理システム')
    .addItem('サイドバーを開く', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar').evaluate().setTitle('案件管理');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** 単純トリガー（権限制限があるため副作用のある処理はここでは行わない） */
function onEdit(e) {
  // 意図的に空実装。インストール型トリガー onEditInstallable が実処理を担当する。
}

/**
 * インストール型 onEdit トリガー本体。
 * メイン画面UIシート（当期分）の、採番対象列が編集された場合に自動採番を行う。
 */
function onEditInstallable(e) {
  try {
    const sheet = e.range.getSheet();
    const activeUiSheet = getActiveUiSheet_();
    if (sheet.getSheetId() !== activeUiSheet.getSheetId()) return;
    if (e.range.getRow() < CASE_DATA_START_ROW) return;

    const editedCol = e.range.getColumn();
    const editedColEnd = e.range.getLastColumn();
    const touchesNumberingCols = getCaseNumberingTriggerCols_().some(
      col => col >= editedCol && col <= editedColEnd
    );
    if (!touchesNumberingCols) return;

    // 複数行に一括貼り付けされた場合も考慮し、範囲内の全行をチェックする
    for (let row = e.range.getRow(); row <= e.range.getLastRow(); row++) {
      handleCaseRowEdit_(sheet, row);
    }
  } catch (err) {
    console.error(`onEditInstallable error: ${err && err.stack ? err.stack : err}`);
  }
}

/** 時間主導型トリガーから毎日呼ばれる想定のエントリーポイント */
function dailyScheduledTasks() {
  ensureNextPeriodResourcesIfNeeded();
  sendBillingSummaryReportIfNeeded_();
}

/** 初回セットアップ用。既存トリガーを削除してから登録し直す（重複登録防止）。 */
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('onEditInstallable')
    .forSpreadsheet(getMainSpreadsheet_())
    .onEdit()
    .create();

  ScriptApp.newTrigger('dailyScheduledTasks')
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .nearMinute(30)
    .create();
}
