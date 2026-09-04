/**
 * SetupService.gs
 * 案件シート（メイン画面UI・全案件DB・それぞれの原本）の構成を検証し、
 * 不足している列・未設定のヘッダーを補う保守用の処理。
 *
 * 主な用途は、列を追加する仕様変更（例: v2.0 で AE列「見積書差し戻し日時」、
 * AF列「請求書差し戻し日時」を追加）を、既存のスプレッドシートへ反映すること。
 * スクリプト側は CASE_LAST_COL 列ぶんの読み書きを行うため、シートの列数が
 * 足りないと範囲外エラーになる。この処理で事前に揃えておく。
 */

/**
 * 構成を確認・修復する対象シートの一覧（当期分 + 原本）。
 * 表示シートと全案件DBシートでは見出しの行数が異なる（表示=1〜2行目、DB=1行目のみ）ため、
 * シートごとに「列見出しが入る行」を持たせる。
 */
function getCaseSheetTargetsForSetup_() {
  const names = getPeriodSheetNames_(getCurrentPeriodNumber_());
  return [
    { name: names.ui, headerRow: CASE_UI_HEADER_ROW },
    { name: names.db, headerRow: CASE_HEADER_ROW },
    { name: MASTER_UI_SHEET_NAME, headerRow: CASE_UI_HEADER_ROW },
    { name: MASTER_DB_SHEET_NAME, headerRow: CASE_HEADER_ROW },
  ];
}

/** 構成を確認・修復する対象シート名の一覧（当期分 + 原本） */
function getCaseSheetNamesForSetup_() {
  return getCaseSheetTargetsForSetup_().map(target => target.name);
}

/**
 * 1枚の案件シートについて、列数とヘッダーを揃える。
 * 既に入力されているヘッダーは尊重し、空欄の見出しのみを補う
 * （利用者が独自の表記に変更している場合を上書きしないため）。
 * @param {Sheet} sheet 対象シート
 * @param {number} headerRow 列見出しが入る行。表示シートは2行目、全案件DBは1行目。
 * @return {string[]} 実施した変更の説明（変更が無ければ空配列）
 */
function ensureCaseSheetStructure_(sheet, headerRow) {
  const changes = [];

  const maxColumns = sheet.getMaxColumns();
  if (maxColumns < CASE_LAST_COL) {
    sheet.insertColumnsAfter(maxColumns, CASE_LAST_COL - maxColumns);
    changes.push(`列数を ${maxColumns} → ${CASE_LAST_COL} へ拡張`);
  }

  const headerRange = sheet.getRange(headerRow, 1, 1, CASE_LAST_COL);
  const headers = headerRange.getValues()[0];
  const filled = [];
  for (let i = 0; i < CASE_LAST_COL; i++) {
    if (String(headers[i] == null ? '' : headers[i]).trim() === '') {
      headers[i] = CASE_HEADERS[i];
      filled.push(`${columnIndexToLetter_(i + 1)}列「${CASE_HEADERS[i]}」`);
    }
  }
  if (filled.length > 0) {
    headerRange.setValues([headers]);
    changes.push(`${headerRow}行目の未設定のヘッダーを追加: ${filled.join('、')}`);
  }

  return changes;
}

/**
 * メニュー「案件シートの構成を確認・修復する」から呼ばれる。
 * 当期のUI/DBシートと原本シートについて、列数とヘッダーを揃えて結果を表示する。
 */
function checkAndRepairCaseSheets() {
  const ui = SpreadsheetApp.getUi();
  const lines = [];

  try {
    withLock_('案件シート構成の確認・修復', () => {
      const ss = getMainSpreadsheet_();
      getCaseSheetTargetsForSetup_().forEach(target => {
        const sheet = ss.getSheetByName(target.name);
        if (!sheet) {
          lines.push(`× ${target.name}: シートが見つかりません`);
          return;
        }
        const changes = ensureCaseSheetStructure_(sheet, target.headerRow);
        lines.push(changes.length > 0
          ? `● ${target.name}: ${changes.join(' / ')}`
          : `○ ${target.name}: 変更なし（正常）`);
      });
      appendOperationLog_('', '案件シート構成の確認・修復', lines.join(' | '), false);
    });
  } catch (e) {
    ui.alert(`案件シート構成の確認・修復に失敗しました。\n\n${e && e.message ? e.message : e}`);
    return;
  }

  ui.alert([
    '案件シート構成の確認・修復',
    '',
    ...lines,
    '',
    `必要な列数: ${CASE_LAST_COL}列（A〜${columnIndexToLetter_(CASE_LAST_COL)}）`,
    `見出しの行: 表示シート=${CASE_UI_HEADER_ROW}行目 / 全案件DB=${CASE_HEADER_ROW}行目`,
  ].join('\n'));
}

/**
 * WEBAPP_URL スクリプトプロパティを設定する（見出しUIなしで呼べる版）。
 * 新しいデプロイを作成してURLが変わった場合、この関数の再実行が必要
 * （メニューの「WebアプリのURLを設定する」からも呼べる）。
 */
function setWebAppUrl_(url) {
  const trimmed = String(url || '').trim();
  if (!/^https:\/\/script\.google\.com\/macros\/s\/.+\/exec$/.test(trimmed)) {
    throw AppError_('INVALID_WEBAPP_URL', `WebアプリのURLの形式が正しくありません: ${trimmed}`);
  }
  PropertiesService.getScriptProperties().setProperty(PROP_KEYS.WEBAPP_URL, trimmed);
}

/** メニュー「WebアプリのURLを設定する」から呼ばれる */
function setWebAppUrlFromPrompt() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'WebアプリのURLを設定',
    '新しくデプロイした際に発行された、".../exec" で終わるURLを貼り付けてください。',
    ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  try {
    setWebAppUrl_(response.getResponseText());
    ui.alert('設定しました。');
  } catch (e) {
    ui.alert(`設定に失敗しました。\n\n${e && e.message ? e.message : e}`);
  }
}
