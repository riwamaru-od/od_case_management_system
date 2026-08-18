/**
 * BackupService.gs
 * 全案件DB（当期分）と取引先DBの日次バックアップ。
 *
 * バックアップフォルダ内に「バックアップ_{yyyyMMdd}」という名前のスプレッドシートを作り、
 * 対象シートをコピーして保存する。1日1ファイルとし、同日に2回目以降が実行された場合は
 * 何もしない（日次トリガーの再実行や手動実行で重複しないようにするため）。
 *
 * 保持世代数は Constants.gs の BACKUP_GENERATIONS（既定30世代）。
 * 超過した古いバックアップはゴミ箱へ移動する（完全削除はしない）。
 */

/** バックアップ対象のシートを取得する。取得できないものはスキップして警告に留める。 */
function getBackupTargetSheets_() {
  const targets = [];

  try {
    targets.push({ sheet: getActiveDbSheet_(), name: getPeriodSheetNames_(getCurrentPeriodNumber_()).db });
  } catch (e) {
    console.warn(`全案件DBシートを取得できないためバックアップ対象から除外します: ${e}`);
  }

  try {
    const clientSheet = getClientDbSheet_();
    if (clientSheet) targets.push({ sheet: clientSheet, name: clientSheet.getName() });
  } catch (e) {
    console.warn(`取引先DBシートを取得できないためバックアップ対象から除外します: ${e}`);
  }

  return targets;
}

/** 本日分のバックアップファイル名 */
function buildBackupFileName_(date) {
  return `${BACKUP_FILE_NAME_PREFIX}${Utilities.formatDate(date, 'Asia/Tokyo', 'yyyyMMdd')}`;
}

/**
 * バックアップを1件作成する。
 * @return {string} 実施内容の説明
 */
function createDatabaseBackup_() {
  const folder = getBackupFolder_();
  const fileName = buildBackupFileName_(new Date());

  if (folder.getFilesByName(fileName).hasNext()) {
    return `本日分のバックアップ「${fileName}」は作成済みのためスキップしました`;
  }

  const targets = getBackupTargetSheets_();
  if (targets.length === 0) {
    throw AppError_('BACKUP_NO_TARGET', 'バックアップ対象のシートを1つも取得できませんでした。');
  }

  const backupSs = SpreadsheetApp.create(fileName);
  moveFileOrFolder_(DriveApp.getFileById(backupSs.getId()), folder);

  const copied = targets.map(target => {
    const copiedSheet = target.sheet.copyTo(backupSs).setName(target.name);
    // 元ファイルの範囲を参照する入力規則はバックアップ側では無効になるため取り除く
    copiedSheet.getDataRange().clearDataValidations();
    return `${target.name}(${copiedSheet.getLastRow()}行)`;
  });

  removeDefaultSheetIfUnused_(backupSs);
  SpreadsheetApp.flush();

  return `${fileName}: ${copied.join('、')}`;
}

/**
 * 保持世代数を超えた古いバックアップをゴミ箱へ移動する。
 * @return {string} 実施内容の説明
 */
function pruneOldBackups_() {
  const folder = getBackupFolder_();
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const file = it.next();
    if (file.getName().indexOf(BACKUP_FILE_NAME_PREFIX) === 0) {
      files.push({ file: file, createdAt: file.getDateCreated() });
    }
  }
  if (files.length <= BACKUP_GENERATIONS) {
    return `保持世代数以内（${files.length}/${BACKUP_GENERATIONS}）のため削除なし`;
  }

  files.sort((a, b) => b.createdAt - a.createdAt); // 新しい順
  const removed = files.slice(BACKUP_GENERATIONS).map(entry => {
    const name = entry.file.getName();
    entry.file.setTrashed(true);
    return name;
  });
  return `古いバックアップ${removed.length}件をゴミ箱へ移動: ${removed.join('、')}`;
}

/** 日次トリガーから呼ばれる。バックアップ作成と、世代超過分の整理を行う。 */
function backupDatabasesIfNeeded_() {
  withLock_('データベースのバックアップ', () => {
    const created = createDatabaseBackup_();
    const pruned = pruneOldBackups_();
    appendOperationLog_('', 'データベースのバックアップ', `${created} / ${pruned}`, false);
  });
}

/** メニュー「今すぐバックアップを作成する」から呼ばれる */
function createDatabaseBackupManually() {
  const ui = SpreadsheetApp.getUi();
  try {
    backupDatabasesIfNeeded_();
    ui.alert('バックアップ\n\n完了しました。\n'
      + `保存先: バックアップフォルダ内の「${buildBackupFileName_(new Date())}」\n`
      + `（同日に既にバックアップがある場合は作成をスキップします。保持世代数: ${BACKUP_GENERATIONS}）`);
  } catch (e) {
    ui.alert(`バックアップに失敗しました。\n\n${e && e.message ? e.message : e}`);
  }
}
