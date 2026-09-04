/**
 * ArchiveService.gs
 * 過去期のデータを別ファイルへ退避し、メインスプレッドシートの肥大化を防ぐ。
 *
 * 方式（アーカイブ方式）:
 *   メインスプレッドシートには「当期分の表示シート・全案件DBシート」と、期をまたいで
 *   使う共有シート（取引先DB・社員DB・操作ログ・原本）だけを置く。
 *   期が切り替わったら、前期以前の「{n}期_表示」「{n}期_全案件DB」を
 *   アーカイブフォルダ内の別ファイル「案件データ_{n}期」へ移し、本体からは削除する。
 *   これにより、作業ファイルは毎年ほぼ同じ大きさに保たれる。
 *
 * 注意:
 *   アーカイブ済みの期の案件は、サイドバーからは参照できなくなる
 *   （サイドバーは当期の表示シート・全案件DBのみを対象とするため）。
 *   過去案件を確認する場合は、アーカイブファイルを直接開くこと。
 */

/**
 * メインスプレッドシート内にある、当期より前の期を列挙する。
 * 表示シートに残っている行数（＝進行中の案件数）も併せて返す。
 * 期の切り替わり時点で未完了の案件が残っている場合、そのままアーカイブすると
 * 進行中の案件が作業ファイルから消えてしまうため、呼び出し元で判断できるようにする。
 * @return {Array<{period: number, activeRows: number}>} 期番号の昇順
 */
function findArchivablePeriods_() {
  const currentPeriod = getCurrentPeriodNumber_();
  const ss = getMainSpreadsheet_();
  const periods = {};

  ss.getSheets().forEach(sheet => {
    const matched = sheet.getName().match(PERIOD_SHEET_NAME_PATTERN);
    if (!matched) return;
    const periodNumber = Number(matched[1]);
    if (periodNumber < currentPeriod) periods[periodNumber] = true;
  });

  return Object.keys(periods).map(Number).sort((a, b) => a - b).map(periodNumber => {
    const uiSheet = ss.getSheetByName(getPeriodSheetNames_(periodNumber).ui);
    const lastRow = uiSheet ? uiSheet.getLastRow() : 0;
    return {
      period: periodNumber,
      activeRows: Math.max(0, lastRow - CASE_UI_DATA_START_ROW + 1),
    };
  });
}

/**
 * 指定した期のシートをアーカイブファイルへ退避し、メインスプレッドシートから削除する。
 * コピーが正しく行われたことを確認してから削除する（データ消失を防ぐため）。
 * @return {string} 実施内容の説明
 */
function archivePeriodSheets_(periodNumber) {
  const ss = getMainSpreadsheet_();
  const names = getPeriodSheetNames_(periodNumber);
  const sourceSheets = [names.ui, names.db]
    .map(name => ss.getSheetByName(name))
    .filter(sheet => sheet !== null);

  if (sourceSheets.length === 0) return `${periodNumber}期: 対象シートなし`;

  // アーカイブ先ファイルを作成（既存があればそれを使う）
  const folder = getArchiveFolder_();
  const archiveName = `${ARCHIVE_FILE_NAME_PREFIX}${periodNumber}期`;
  const archiveSs = openOrCreateArchiveSpreadsheet_(archiveName, folder);

  const copied = [];
  sourceSheets.forEach(sourceSheet => {
    const sheetName = sourceSheet.getName();
    // 同名シートが既にある場合は作り直す（再実行時の重複を避ける）
    const existing = archiveSs.getSheetByName(sheetName);
    if (existing) archiveSs.deleteSheet(existing);

    const copiedSheet = sourceSheet.copyTo(archiveSs).setName(sheetName);
    // 元ファイルの範囲を参照している入力規則はアーカイブ側では無効になるため取り除く
    copiedSheet.getDataRange().clearDataValidations();

    // 削除前の整合性チェック: 行数・列数が一致していることを確認する
    if (copiedSheet.getLastRow() !== sourceSheet.getLastRow()
      || copiedSheet.getLastColumn() !== sourceSheet.getLastColumn()) {
      throw AppError_('ARCHIVE_VERIFY_FAILED',
        `アーカイブしたシート「${sheetName}」の内容が元シートと一致しません。`
        + '安全のため元シートは削除しませんでした。アーカイブファイルの内容を確認してください。');
    }
    copied.push({ sourceSheet, sheetName, rows: copiedSheet.getLastRow() });
  });

  SpreadsheetApp.flush();

  // コピーの検証が全て通ってから、本体側のシートを削除する
  const removed = copied.map(entry => {
    ss.deleteSheet(entry.sourceSheet);
    return `${entry.sheetName}(${entry.rows}行)`;
  });

  return `${periodNumber}期: ${removed.join('、')} を「${archiveName}」へ退避`;
}

/** アーカイブ用スプレッドシートを取得（無ければ作成してアーカイブフォルダへ移動） */
function openOrCreateArchiveSpreadsheet_(name, folder) {
  const existingFiles = folder.getFilesByName(name);
  if (existingFiles.hasNext()) {
    return SpreadsheetApp.openById(existingFiles.next().getId());
  }
  const created = SpreadsheetApp.create(name);
  moveFileOrFolder_(DriveApp.getFileById(created.getId()), folder);
  return created;
}

/** 作成直後のアーカイブファイルに残る既定の空シートを削除する */
function removeDefaultSheetIfUnused_(ss) {
  if (ss.getSheets().length <= 1) return;
  ss.getSheets().forEach(sheet => {
    const isDefault = /^(シート1|Sheet1)$/.test(sheet.getName());
    if (isDefault && sheet.getLastRow() === 0) ss.deleteSheet(sheet);
  });
}

/** 指定した期の一覧をアーカイブする（内部共通処理） */
function archivePeriods_(periodNumbers) {
  return withLock_('過去期データのアーカイブ', () => {
    const results = periodNumbers.map(periodNumber => {
      const result = archivePeriodSheets_(periodNumber);
      const archiveName = `${ARCHIVE_FILE_NAME_PREFIX}${periodNumber}期`;
      removeDefaultSheetIfUnused_(openOrCreateArchiveSpreadsheet_(archiveName, getArchiveFolder_()));
      return result;
    });
    appendOperationLog_('', '過去期データのアーカイブ', results.join(' | '), false);
    return results;
  });
}

/** 「その期の案件がすべて完了した日」を記録するスクリプトプロパティのキー */
function archiveReadySinceKey_(periodNumber) {
  return `ARCHIVE_READY_SINCE_${periodNumber}`;
}

/** 完了状態の記録を消す（案件が戻された場合と、アーカイブ完了後に呼ぶ） */
function clearArchiveReadyMark_(periodNumber) {
  PropertiesService.getScriptProperties().deleteProperty(archiveReadySinceKey_(periodNumber));
}

/**
 * その期の案件がすべて完了した状態になってからの経過日数を返す。
 * 初めて完了状態を確認した日を記録し、以降はその日からの日数を数える。
 * @return {number} 経過日数（初回は0）
 */
function daysSinceArchiveReady_(periodNumber, today) {
  const props = PropertiesService.getScriptProperties();
  const key = archiveReadySinceKey_(periodNumber);
  const todayText = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd');
  const recorded = props.getProperty(key);

  if (!recorded) {
    props.setProperty(key, todayText);
    return 0;
  }

  const from = new Date(`${String(recorded).replace(/-/g, '/')} 00:00:00`);
  if (isNaN(from.getTime())) {
    props.setProperty(key, todayText);
    return 0;
  }
  const to = new Date(`${todayText} 00:00:00`);
  return Math.round((to.getTime() - from.getTime()) / (24 * 60 * 60 * 1000));
}

/**
 * 日次トリガーから呼ばれる。当期より前の期のシートが残っていればアーカイブする。
 *
 * アーカイブするのは「その期の案件がすべて完了（最終承認または中止）し、そこから
 * ARCHIVE_WAIT_DAYS_AFTER_CLOSED 日が経過した期」に限る。完了した当日に作業ファイルから
 * 消してしまうと、最終承認の取り消しなどの戻し操作ができなくなるため、猶予を置いている。
 *
 * 猶予期間中に案件が戻された場合（最終承認の取り消しなど）は完了日の記録を消し、
 * 改めて全件完了した日から数え直す。
 */
function archiveOldPeriodSheetsIfNeeded_() {
  const candidates = findArchivablePeriods_();
  if (candidates.length === 0) return;

  const today = new Date();
  const ready = [];
  const waiting = [];
  const pending = [];

  candidates.forEach(entry => {
    if (entry.activeRows > 0) {
      // まだ案件が残っている（または戻された）ので、完了日の記録をやり直す
      clearArchiveReadyMark_(entry.period);
      pending.push(entry);
      return;
    }
    const elapsedDays = daysSinceArchiveReady_(entry.period, today);
    if (elapsedDays >= ARCHIVE_WAIT_DAYS_AFTER_CLOSED) {
      ready.push(entry);
    } else {
      waiting.push(entry);
    }
  });

  if (pending.length > 0) {
    const detail = pending.map(e => `${e.period}期(${e.activeRows}件)`).join('、');
    console.warn(`進行中の案件が残っているためアーカイブを見送りました: ${detail}`);
    appendOperationLog_('', '過去期データのアーカイブ',
      `進行中の案件が残っている期は自動アーカイブを見送りました: ${detail}`, false);
  }

  if (waiting.length > 0) {
    const detail = waiting.map(e => `${e.period}期`).join('、');
    appendOperationLog_('', '過去期データのアーカイブ',
      `全案件が完了したため、${ARCHIVE_WAIT_DAYS_AFTER_CLOSED}日後にアーカイブします: ${detail}`, false);
  }

  if (ready.length === 0) return;
  archivePeriods_(ready.map(entry => entry.period));
  ready.forEach(entry => clearArchiveReadyMark_(entry.period));
}

/** メニュー「過去期のデータをアーカイブする」から呼ばれる */
function archiveOldPeriodSheetsManually() {
  const ui = SpreadsheetApp.getUi();
  const candidates = findArchivablePeriods_();

  if (candidates.length === 0) {
    ui.alert('過去期データのアーカイブ\n\nアーカイブ対象の期はありません。\n'
      + `（メインスプレッドシートには当期(${getCurrentPeriodNumber_()}期)のシートのみが残っています）`);
    return;
  }

  const pending = candidates.filter(entry => entry.activeRows > 0);
  const lines = candidates.map(e =>
    `・${e.period}期${e.activeRows > 0 ? `（表示シートに ${e.activeRows} 件の案件が残っています）` : '（残案件なし）'}`);

  const warning = pending.length > 0
    ? '\n\n【注意】残案件がある期を含めてアーカイブすると、その案件は作業ファイルから消え、\n'
      + 'サイドバーから操作できなくなります（アーカイブファイル内には残ります）。'
    : '';

  const confirmed = ui.alert(
    '過去期データのアーカイブ',
    '以下の期のシートを別ファイルへ退避し、このスプレッドシートから削除します。\n\n'
    + lines.join('\n')
    + '\n\n退避先: アーカイブフォルダ内の「案件データ_{期}期」'
    + warning
    + `\n\n※自動アーカイブは「その期の案件がすべて完了してから${ARCHIVE_WAIT_DAYS_AFTER_CLOSED}日後」に実行されます。`
    + '\n　この操作は猶予を待たずに今すぐ実行します。'
    + '\n\n実行してよろしいですか？',
    ui.ButtonSet.OK_CANCEL);
  if (confirmed !== ui.Button.OK) return;

  try {
    const results = archivePeriods_(candidates.map(entry => entry.period));
    ui.alert(`過去期データのアーカイブが完了しました。\n\n${results.join('\n')}`);
  } catch (e) {
    ui.alert(`アーカイブに失敗しました。\n\n${e && e.message ? e.message : e}`);
  }
}
