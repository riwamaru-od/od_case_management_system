/**
 * ClientOptionsService.gs
 * 案件シートの取引先名（B列）のプルダウン選択肢を「config」シートに自動生成し、
 * データ入力規則として反映する。
 *
 * 取引先DBには同じ社名で担当者だけが異なる行が複数登録される。社名だけを選ばせると
 * どの行の住所・担当者名を書類へ転記すべきか決められないため、選択肢は
 * 「社名（担当者名）」の形式にする（表示名の組み立ては ClientService.gs の
 * buildClientOptionLabel_、表示名から取引先DBの行を引くのは getClientByName_）。
 *
 * 取引先はGoogleフォームからの登録で随時増えるため、終了予定・請求予定の選択肢と同様に
 * 日次トリガーで再生成する（管理者用メニューから手動でも更新できる）。
 * configシートの取得（ensureConfigSheet_）は ScheduleOptionsService.gs と共用し、
 * 書き出す列だけを分けている（CLIENT_OPTIONS_CONFIG_COL）。
 */

/**
 * configシートへ取引先名の選択肢一覧を書き込む。
 * 既存の内容は広めにクリアしてから書き込み、取引先が減っても古い行が残らないようにする。
 * @return {GoogleAppsScript.Spreadsheet.Range|null} 書き込んだ選択肢のセル範囲（取引先が0件なら null）
 */
function writeClientOptionsToConfigSheet_(sheet, options) {
  const col = CLIENT_OPTIONS_CONFIG_COL;
  sheet.getRange(1, col, CLIENT_OPTIONS_CLEAR_ROWS, 1).clearContent();
  sheet.getRange(1, col).setValue('取引先名の選択肢（自動生成・編集不要）');
  if (!options.length) return null;

  const range = sheet.getRange(2, col, options.length, 1);
  range.setValues(options.map(o => [o]));
  return range;
}

/**
 * 指定シートのB列（取引先名）へ、configシートの範囲を参照するプルダウンを設定する。
 * 適用開始行は終了予定・請求予定と同じ（SCHEDULE_VALIDATION_START_ROW）。
 */
function applyClientDataValidation_(sheet, optionsRange) {
  const maxRows = sheet.getMaxRows();
  if (maxRows < SCHEDULE_VALIDATION_START_ROW) return;
  const numRows = maxRows - SCHEDULE_VALIDATION_START_ROW + 1;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(optionsRange, true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(SCHEDULE_VALIDATION_START_ROW, CASE_COLS.CLIENT_NAME, numRows, 1).setDataValidation(rule);
}

/**
 * configシートの取引先名の選択肢を最新化し、対象シート（当期のUI/DBシートと原本シート）へ
 * プルダウンの参照設定を適用する。
 * 原本シートにも設定しておくことで、期切替でコピーされる将来のシートへ引き継がれる。
 */
function refreshClientOptionsAndValidation_() {
  const configSheet = ensureConfigSheet_();
  clearClientCache_(); // 取引先が追加された直後でも最新の一覧で作り直せるようにする
  const options = listClientOptionLabels_();
  const optionsRange = writeClientOptionsToConfigSheet_(configSheet, options);
  if (!optionsRange) {
    return { options, results: ['× 取引先DBに登録がないため、プルダウンは設定しませんでした'] };
  }

  const ss = getMainSpreadsheet_();
  const results = getCaseSheetNamesForSetup_().map(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return `× ${name}: シートが見つかりません`;
    applyClientDataValidation_(sheet, optionsRange);
    return `○ ${name}: 取引先${options.length}件を設定`;
  });

  return { options, results };
}

/** 日次トリガーから呼ばれる */
function refreshClientOptionsIfNeeded_() {
  withLock_('取引先名の選択肢更新', () => {
    const { options, results } = refreshClientOptionsAndValidation_();
    appendOperationLog_('', '取引先名の選択肢更新',
      `選択肢${options.length}件 | ${results.join(' | ')}`, false);
  });
}

/** メニュー「取引先名の選択肢を更新する」から呼ばれる */
function refreshClientOptionsManually() {
  const ui = SpreadsheetApp.getUi();
  try {
    refreshClientOptionsIfNeeded_();
    const options = listClientOptionLabels_();
    ui.alert([
      '取引先名の選択肢を更新しました',
      '',
      `選択肢: ${options.length}件（「社名（担当者名）」の形式）`,
      `参照先: 「${CONFIG_SHEET_NAME}」シートの${columnIndexToLetter_(CLIENT_OPTIONS_CONFIG_COL)}列`,
      '',
      '取引先DBに新しい取引先が登録されたら、この操作で選択肢へ反映されます',
      '（日次処理でも毎日自動で更新されます）。',
    ].join('\n'));
  } catch (e) {
    ui.alert(`取引先名の選択肢の更新に失敗しました。\n\n${e && e.message ? e.message : e}`);
  }
}
