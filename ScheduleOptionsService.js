/**
 * ScheduleOptionsService.gs
 * 終了予定・請求予定（D列/E列）のプルダウン選択肢を「config」シートに自動生成し、
 * 案件シートのデータ入力規則（プルダウン）として反映する。
 * あわせて、請求予定が「今月」に一致する行を青文字にする条件付き書式を設定する。
 *
 * 選択肢は日付に連動して変わる（今月を基準にローリングする）ため、日次トリガーで
 * 毎日再生成する。生成される選択肢の並びは以下の通り:
 *   1. SCHEDULE_OPTIONS_FIXED_CHOICES（例:「未定」「営業案件」）
 *   2. {SCHEDULE_OPTIONS_MONTHS_BEFORE}ヶ月前 〜 {SCHEDULE_OPTIONS_MONTHS_AFTER}ヶ月後 の
 *      各月について「YYYY/MM 前半」「YYYY/MM 後半」
 */

/** configシートを取得（無ければ作成） */
function ensureConfigSheet_() {
  const ss = getMainSpreadsheet_();
  let sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG_SHEET_NAME);
  return sheet;
}

/**
 * 終了予定・請求予定の選択肢一覧を組み立てる。
 * @return {string[]} 固定選択肢 + 月範囲（前半/後半）のラベル一覧
 */
function buildScheduleOptionList_() {
  const now = new Date();
  const baseYear = Number(Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy'));
  const baseMonthIndex = Number(Utilities.formatDate(now, 'Asia/Tokyo', 'M')) - 1; // 0始まり

  const options = SCHEDULE_OPTIONS_FIXED_CHOICES.slice();
  for (let offset = -SCHEDULE_OPTIONS_MONTHS_BEFORE; offset <= SCHEDULE_OPTIONS_MONTHS_AFTER; offset++) {
    const target = new Date(baseYear, baseMonthIndex + offset, 1);
    const label = Utilities.formatDate(target, 'Asia/Tokyo', 'yyyy/MM');
    options.push(`${label} 前半`);
    options.push(`${label} 後半`);
  }
  return options;
}

/**
 * configシートへ選択肢一覧を書き込む。
 * 既存の内容は範囲を広めにクリアしてから書き込み、月範囲がずれても古い行が
 * 残らないようにする。
 * @return {GoogleAppsScript.Spreadsheet.Range} 書き込んだ選択肢のセル範囲（A2以降）
 */
function writeScheduleOptionsToConfigSheet_(sheet, options) {
  const CLEAR_ROWS = 200; // 選択肢の増減に十分な余裕を持たせてクリアする
  sheet.getRange(1, 1, CLEAR_ROWS, 1).clearContent();

  sheet.getRange(1, 1).setValue('終了予定・請求予定の選択肢（自動生成・編集不要）');
  const range = sheet.getRange(2, 1, options.length, 1);
  range.setValues(options.map(o => [o]));
  return range;
}

/**
 * 指定シートのD列（終了予定）・E列（請求予定）へ、configシートの範囲を参照するプルダウンを設定する。
 * 案件データはCASE_DATA_START_ROW（2行目）から始まるが、プルダウンの適用開始行は
 * それとは別に SCHEDULE_VALIDATION_START_ROW（既定3行目）を使う。
 */
function applyScheduleDataValidation_(sheet, optionsRange) {
  const maxRows = sheet.getMaxRows();
  if (maxRows < SCHEDULE_VALIDATION_START_ROW) return;
  const numRows = maxRows - SCHEDULE_VALIDATION_START_ROW + 1;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(optionsRange, true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(SCHEDULE_VALIDATION_START_ROW, CASE_COLS.END_SCHEDULE, numRows, 1).setDataValidation(rule);
  sheet.getRange(SCHEDULE_VALIDATION_START_ROW, CASE_COLS.BILLING_SCHEDULE, numRows, 1).setDataValidation(rule);
}

/**
 * 指定シートの請求予定（E列）に、「今月分（前半・後半）」を青文字にする条件付き書式を設定する。
 * 実行するたびに、対象範囲が完全一致する既存ルール（＝本関数が過去に設定したもの）を
 * 差し替える形で1本だけ登録し直す。他の目的で設定された条件付き書式には影響しない。
 */
function applyCurrentMonthBillingHighlight_(sheet) {
  const maxRows = sheet.getMaxRows();
  if (maxRows < CASE_DATA_START_ROW) return;
  const numRows = maxRows - CASE_DATA_START_ROW + 1;
  const range = sheet.getRange(CASE_DATA_START_ROW, CASE_COLS.BILLING_SCHEDULE, numRows, 1);
  const targetA1 = range.getA1Notation();

  const otherRules = sheet.getConditionalFormatRules().filter(rule => {
    const ranges = rule.getRanges().map(r => r.getA1Notation());
    return !(ranges.length === 1 && ranges[0] === targetA1);
  });

  // 条件付き書式の数式は範囲の左上セルを相対参照で書く（各行に自動調整される）
  const topLeftA1 = range.getCell(1, 1).getA1Notation();
  const formula = `=OR(${topLeftA1}=TEXT(TODAY(),"yyyy/mm")&" 前半",${topLeftA1}=TEXT(TODAY(),"yyyy/mm")&" 後半")`;

  const newRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setFontColor(CURRENT_MONTH_BILLING_TEXT_COLOR)
    .setRanges([range])
    .build();

  sheet.setConditionalFormatRules([...otherRules, newRule]);
}

/**
 * config シートの選択肢を最新化し、対象シート（当期のUI/DBシートと原本シート）へ
 * プルダウンの参照設定と今月ハイライトを適用する。
 * 原本シートに設定しておくことで、期切替でコピーされる将来のシートにも自動的に引き継がれる。
 */
function refreshScheduleOptionsAndFormatting_() {
  const configSheet = ensureConfigSheet_();
  const options = buildScheduleOptionList_();
  const optionsRange = writeScheduleOptionsToConfigSheet_(configSheet, options);

  const ss = getMainSpreadsheet_();
  const results = getCaseSheetNamesForSetup_().map(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return `× ${name}: シートが見つかりません`;
    applyScheduleDataValidation_(sheet, optionsRange);
    applyCurrentMonthBillingHighlight_(sheet);
    return `○ ${name}: 選択肢${options.length}件・今月ハイライトを設定`;
  });

  return { options, results };
}

/** 日次トリガーから呼ばれる */
function refreshScheduleOptionsIfNeeded_() {
  withLock_('終了予定・請求予定の選択肢更新', () => {
    const { options, results } = refreshScheduleOptionsAndFormatting_();
    appendOperationLog_('', '終了予定・請求予定の選択肢更新',
      `選択肢${options.length}件 | ${results.join(' | ')}`, false);
  });
}

/** メニュー「終了予定・請求予定の選択肢を更新する」から呼ばれる */
function refreshScheduleOptionsManually() {
  const ui = SpreadsheetApp.getUi();
  try {
    refreshScheduleOptionsIfNeeded_();
    const options = buildScheduleOptionList_();
    ui.alert([
      '終了予定・請求予定の選択肢を更新しました',
      '',
      `選択肢（${options.length}件）:`,
      options.join('、'),
      '',
      `参照先: 「${CONFIG_SHEET_NAME}」シート`,
      '請求予定が今月分の行は、自動的に文字色が青になります。',
    ].join('\n'));
  } catch (e) {
    ui.alert(`選択肢の更新に失敗しました。\n\n${e && e.message ? e.message : e}`);
  }
}
