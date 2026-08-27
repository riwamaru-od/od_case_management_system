/**
 * Triggers.gs
 * onOpen / onEdit（インストール型） / 時間主導型トリガーの入り口をまとめる。
 *
 * インストール型トリガーは「登録したユーザー個人」に紐づき、
 * ScriptApp.getProjectTriggers() も実行ユーザー自身のトリガーしか返さない。
 * この性質を踏まえ、トリガー登録は用途ごとに2つの関数へ分離している。
 *
 * セットアップ手順:
 *   1) 【管理者のみ1回】管理用アカウントでエディタから installTriggers() を実行する。
 *      → 自動採番（onEditInstallable）・日次処理（dailyScheduledTasks）に加えて、
 *        実行者本人のサイドバー自動表示トリガーが登録される。
 *      重要: onEdit と日次処理は「誰か1人だけ」が登録すること。複数人が登録すると
 *      1回の編集で人数分の採番処理が走り、請求予定レポートも人数分重複送信される。
 *   2) 【利用者が各自1回】メニュー「案件管理システム > サイドバー自動表示を有効にする」
 *      をクリックする（スクリプトエディタを開く必要はない）。
 *      → その人専用のサイドバー自動表示トリガー（onOpenInstallable）だけが登録される。
 *        初回はGoogleの認可ダイアログが表示されるので承認すること。
 *   3) 単純トリガー onEdit(e) は権限の都合上フル機能を実行できないため、
 *      副作用を伴う処理は全てインストール型トリガー側で行う（onEdit(e) は何もしない）。
 *      自動表示を設定していない利用者も、メニュー「サイドバーを開く」からいつでも開ける。
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('案件管理システム')
    .addItem('サイドバーを開く', 'showSidebar')
    .addSeparator()
    .addItem('サイドバー自動表示を有効にする', 'enableSidebarAutoOpen')
    .addItem('サイドバー自動表示を解除する', 'disableSidebarAutoOpen')
    .addSeparator()
    .addItem('社印設定を確認する', 'checkSealImageSetting')
    .addItem('案件シートの構成を確認・修復する', 'checkAndRepairCaseSheets')
    .addItem('終了予定・請求予定の選択肢を更新する', 'refreshScheduleOptionsManually')
    .addSeparator()
    .addItem('今すぐバックアップを作成する', 'createDatabaseBackupManually')
    .addItem('過去期のデータをアーカイブする', 'archiveOldPeriodSheetsManually')
    .addToUi();
  // 単純トリガーの文脈では失敗しうるため、確実な自動オープンは
  // インストール型トリガー（onOpenInstallable）側に任せる。
  try {
    showSidebar();
  } catch (e) {
    console.warn(`サイドバーの自動オープンに失敗しました: ${e}`);
  }
}

/**
 * インストール型 onOpen トリガー本体。スプレッドシートを開いた際にサイドバーを開く。
 * 単純トリガーと違い認可済みの権限で実行されるため、確実にサイドバーを表示できる。
 */
function onOpenInstallable(e) {
  try {
    showSidebar();
  } catch (err) {
    console.error(`onOpenInstallable error: ${err && err.stack ? err.stack : err}`);
  }
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

/**
 * 時間主導型トリガーから毎日呼ばれる想定のエントリーポイント。
 * 1つの処理が失敗しても後続が止まらないよう、それぞれ個別に例外を捕捉する。
 */
function dailyScheduledTasks() {
  runDailyTask_('翌期リソースの準備', ensureNextPeriodResourcesIfNeeded);
  runDailyTask_('終了予定・請求予定の選択肢更新', refreshScheduleOptionsIfNeeded_);
  runDailyTask_('過去期データのアーカイブ', archiveOldPeriodSheetsIfNeeded_);
  runDailyTask_('データベースのバックアップ', backupDatabasesIfNeeded_);
  runDailyTask_('請求予定レポートの送信', sendBillingSummaryReportIfNeeded_);
}

/** 日次処理の1項目を実行する。失敗しても後続の処理は続行し、操作ログへ記録する。 */
function runDailyTask_(label, fn) {
  try {
    fn();
  } catch (e) {
    console.error(`[日次処理] ${label} に失敗しました: ${e && e.stack ? e.stack : e}`);
    try {
      appendOperationLog_('', `日次処理: ${label}`, e && e.message ? String(e.message) : String(e), true);
    } catch (logError) {
      console.error(`操作ログの記録にも失敗しました: ${logError}`);
    }
  }
}

/**
 * 実行ユーザー自身が持つ、指定ハンドラ関数のトリガーだけを削除する。
 * （getProjectTriggers() は実行ユーザーのトリガーしか返さないため、
 *  他の担当者が登録したトリガーには影響しない）
 */
function deleteMyTriggersByHandler_(handlerName) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === handlerName)
    .forEach(t => ScriptApp.deleteTrigger(t));
}

/**
 * 【利用者が各自1回実行】サイドバー自動表示トリガーを、実行ユーザー個人に対して登録する。
 * 複数人がそれぞれ実行してよい（各自のトリガーは独立しており、他の人には影響しない）。
 * 重複登録を避けるため、自分の既存トリガーがあれば削除してから作り直す。
 */
function installSidebarAutoOpenTrigger() {
  deleteMyTriggersByHandler_('onOpenInstallable');
  ScriptApp.newTrigger('onOpenInstallable')
    .forSpreadsheet(getMainSpreadsheet_())
    .onOpen()
    .create();
}

/** メニュー「サイドバー自動表示を有効にする」から呼ばれる */
function enableSidebarAutoOpen() {
  const ui = SpreadsheetApp.getUi();
  try {
    installSidebarAutoOpenTrigger();
    ui.alert('サイドバー自動表示を有効にしました。\n次回このスプレッドシートを開いたときから、自動的にサイドバーが表示されます。');
  } catch (e) {
    console.error(`enableSidebarAutoOpen error: ${e && e.stack ? e.stack : e}`);
    ui.alert(`サイドバー自動表示の設定に失敗しました。\n${e && e.message ? e.message : e}`);
  }
}

/** メニュー「サイドバー自動表示を解除する」から呼ばれる */
function disableSidebarAutoOpen() {
  const ui = SpreadsheetApp.getUi();
  try {
    deleteMyTriggersByHandler_('onOpenInstallable');
    ui.alert('サイドバー自動表示を解除しました。\nメニュー「サイドバーを開く」からはこれまで通り手動で開けます。');
  } catch (e) {
    console.error(`disableSidebarAutoOpen error: ${e && e.stack ? e.stack : e}`);
    ui.alert(`サイドバー自動表示の解除に失敗しました。\n${e && e.message ? e.message : e}`);
  }
}

/**
 * 【管理者のみ1回実行】システム全体で1組だけ必要なトリガーを登録する。
 * 自動採番（onEditInstallable）と日次処理（dailyScheduledTasks）は、
 * 複数人が登録すると1回の編集で人数分の処理が走り、請求予定レポートも
 * 人数分重複送信されてしまうため、必ず管理用アカウント1人だけが登録すること。
 * あわせて、実行者本人のサイドバー自動表示トリガーも登録する。
 */
function installTriggers() {
  deleteMyTriggersByHandler_('onEditInstallable');
  deleteMyTriggersByHandler_('dailyScheduledTasks');

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

  installSidebarAutoOpenTrigger();
}
