/**
 * TestDataResetService.gs
 *
 * 【管理者専用・使い捨てではない常設ユーティリティ】テスト用に作成したデータを消し、
 * 実運用を開始できる状態（またはテストのやり直しができる状態）に戻すための処理。
 * メニューには登録しない。管理用アカウントがスクリプトエディタから
 * 手動で runResetForUserTest() を実行すること。
 *
 * 対象:
 *  - 当期の メイン画面UI・全案件DB シートのデータ行（ヘッダー・書式・入力規則は残す）
 *  - 当期の案件番号カウンター（SEQ_CASE_<期>）を0にリセット
 *  - 操作ログシートのデータ行
 *  - 見積書・請求書・納品書ルートフォルダ配下、当期フォルダ内のファイル（ゴミ箱へ）
 *
 * 対象外（意図的に触らない）:
 *  - 取引先DB
 *
 * 注意:
 *  - Google Driveの仕様上、管理用アカウントが所有者でないファイル（Webアプリ経由の
 *    代理実行を導入する前に一般社員アカウントが直接作成したファイル等）はゴミ箱へ
 *    移動できない。その場合は実行結果に一覧が表示されるので、所有者本人に削除を
 *    依頼するか、Driveの共有設定で所有権を管理用アカウントへ移してから再実行すること。
 */

/**
 * スクリプトエディタの「実行する関数」プルダウンから選べるように、
 * 末尾が「_」で終わらない名前で用意した実行用の入り口。
 * （Apps Scriptエディタは末尾が「_」の関数をプルダウンに表示しないため）
 * 管理用アカウントでこの関数を選択して実行すること。
 */
function runResetForUserTest() {
  const message = resetForUserTest_();
  try {
    SpreadsheetApp.getUi().alert('テストデータのリセット完了\n\n' + message);
  } catch (e) {
    Logger.log(message);
  }
}

/**
 * 表示シートの1〜2行目の見出しが何らかの理由で崩れた場合に、
 * 原本_表示シート（データ行の書き込みが一切行われない元シート）からコピーして復元する。
 * 末尾が「_」で終わらない名前で用意しているのは、スクリプトエディタの
 * 「実行する関数」プルダウンから選べるようにするため。
 */
function runRestoreUiHeaderRow2() {
  const ss = getMainSpreadsheet_();
  const master = ss.getSheetByName(MASTER_UI_SHEET_NAME);
  if (!master) throw AppError_('SHEET_NOT_FOUND', `シート「${MASTER_UI_SHEET_NAME}」が見つかりません。`);
  const uiSheet = getActiveUiSheet_();

  const headerRange = master.getRange(1, 1, 2, CASE_LAST_COL);
  uiSheet.getRange(1, 1, 2, CASE_LAST_COL).setValues(headerRange.getValues());

  const message = `${uiSheet.getName()} の1〜2行目の見出しを ${MASTER_UI_SHEET_NAME} からコピーして復元しました。`;
  SpreadsheetApp.getUi().alert(message);
  Logger.log(message);
}

function resetForUserTest_() {
  return withLock_('テストデータのリセット', () => {
    const period = getCurrentPeriodNumber_();
    const summary = [];

    const uiSheet = getActiveUiSheet_();
    const uiCleared = clearCaseDataRows_(uiSheet, CASE_UI_DATA_START_ROW);
    summary.push(`メイン画面UI（${uiSheet.getName()}）: ${uiCleared}行クリア`);

    const dbSheet = getActiveDbSheet_();
    const dbCleared = clearCaseDataRows_(dbSheet, CASE_DATA_START_ROW);
    summary.push(`全案件DB（${dbSheet.getName()}）: ${dbCleared}行クリア`);

    const seqKey = seqKey_('CASE', period);
    PropertiesService.getScriptProperties().deleteProperty(seqKey);
    summary.push(`案件番号カウンター（${seqKey}）: リセット`);

    const logCleared = clearOperationLogRows_();
    summary.push(`操作ログ: ${logCleared}行クリア`);

    const skipped = [];
    ['quote', 'invoice', 'delivery'].forEach(kind => {
      const trashedCount = trashAllFilesUnderCurrentPeriodFolder_(kind, period, skipped);
      summary.push(`${kind}ルートフォルダの${period}期フォルダ: ファイル${trashedCount}件をゴミ箱へ`);
    });
    if (skipped.length > 0) {
      summary.push(
        `※ ${skipped.length}件は削除できませんでした（管理用アカウントが所有者でないため。`
        + `プロキシ導入前のテストで一般社員アカウントが直接作成したファイルと考えられます。`
        + `所有者本人が削除するか、Driveの共有設定で管理用アカウントに所有権を移してから再実行してください）:\n`
        + skipped.map(s => `  - ${s}`).join('\n')
      );
    }

    const message = summary.join('\n');
    Logger.log(message);
    return message;
  });
}

/** シートのヘッダー・書式・入力規則を残したまま、データ行（値のみ）をクリアする */
function clearCaseDataRows_(sheet, dataStartRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) return 0;
  const range = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, CASE_LAST_COL);
  range.clearContent();
  return lastRow - dataStartRow + 1;
}

/** 操作ログシートのヘッダー行（1行目）を残し、データ行をクリアする */
function clearOperationLogRows_() {
  const sheet = getOrCreateOperationLogSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  return lastRow - 1;
}

/**
 * 書類種別ルートフォルダ配下の「当期フォルダ」以下にある全ファイル（サブフォルダ含む）をゴミ箱へ移動する。
 * @param {string[]} skipped 所有者が管理用アカウントでないなどの理由で削除できなかったファイル名を集める配列
 */
function trashAllFilesUnderCurrentPeriodFolder_(kind, periodNumber, skipped) {
  const periodFolder = getDocumentFolder_(kind, periodNumber, null);
  return trashFilesRecursively_(periodFolder, skipped);
}

function trashFilesRecursively_(folder, skipped) {
  let count = 0;
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    try {
      file.setTrashed(true);
      count++;
    } catch (e) {
      skipped.push(`${file.getName()}（${folder.getName()}）: ${e && e.message ? e.message : e}`);
    }
  }
  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    count += trashFilesRecursively_(subfolders.next(), skipped);
  }
  return count;
}
