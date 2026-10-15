/**
 * CaseImportService.gs
 *
 * 【管理者専用】旧「案件番号一覧」で管理していた当期の案件を、このシステムへ移す処理。
 * メニューには登録しない。管理用アカウントがスクリプトエディタから手動で実行すること。
 *
 * 手順:
 *  1) 本番スプレッドシートに「移行データ」シート（MIGRATION_SHEET_NAME）を作り、
 *     1行目を見出し、2行目以降に移行する案件を貼り付ける（列の並びは MIGRATION_COLS）。
 *  2) runCheckMigrationSheet() で内容を点検する（何も書き込まない）。
 *  3) テストデータが残っていれば runResetForUserTest() で片付ける。
 *  4) runImportCasesFromMigrationSheet() で取り込む。
 *
 * 取り込みの内容:
 *  - 案件番号は旧一覧の番号をそのまま使う（請求書等で取引先に伝わっているため振り直さない）
 *  - 進行中: 表示シートと全案件DBへ「案件登録済み・未請求」で登録する
 *  - 完了  : 全案件DBにだけ「最終承認済み」で記録する（書類はシステム上に無い）
 *  - 中止  : 全案件DBにだけ「中止」で記録する
 *  - 当期の案件番号カウンターを、取り込んだ最大の番号に合わせる（次の新規案件はその続きから）
 *
 * 進行中の案件はこの後システム上で書類を作るため、取引先名・担当・予定がプルダウンの値と
 * 一致していることを確かめる。完了・中止の案件は履歴として残すだけなので、取引先が取引先DBに
 * 未登録でも、担当が複数名でも、予定が空欄でもそのまま記録する。
 *
 * 1件でも不備があれば何も書き込まずに中止する。また、表示シート・全案件DBに
 * 既に案件が入っている場合も中止する（既存の案件を上書きしないため）。
 */

/** スクリプトエディタから実行する: 移行データシートを点検する（書き込みは行わない） */
function runCheckMigrationSheet() {
  const { rows, errors } = readAndValidateMigrationRows_();
  const message = errors.length
    ? [`不備が${errors.length}件あります。修正してから再度点検してください。`, '', ...errors].join('\n')
    : ['問題はありません。', '', ...summarizeMigrationRows_(rows)].join('\n');
  showMigrationMessage_('移行データの点検', message);
}

/** スクリプトエディタから実行する: 移行データシートの案件を取り込む */
function runImportCasesFromMigrationSheet() {
  const message = importCasesFromMigrationSheet_();
  showMigrationMessage_('移行データの取り込み', message);
}

function importCasesFromMigrationSheet_() {
  return withLock_('旧一覧からの案件移行', () => {
    const { rows, errors } = readAndValidateMigrationRows_();
    if (errors.length) {
      return [`不備が${errors.length}件あるため、取り込みを中止しました（何も書き込んでいません）。`, '', ...errors].join('\n');
    }
    if (!rows.length) return '移行データシートに案件がありません。';

    const uiSheet = getActiveUiSheet_();
    const dbSheet = getActiveDbSheet_();
    const occupied = [
      countFilledCaseRows_(uiSheet, CASE_UI_DATA_START_ROW) && uiSheet.getName(),
      countFilledCaseRows_(dbSheet, CASE_DATA_START_ROW) && dbSheet.getName(),
    ].filter(Boolean);
    if (occupied.length) {
      return [
        `${occupied.join('・')} に既に案件が入っているため、取り込みを中止しました（何も書き込んでいません）。`,
        'テストデータであれば runResetForUserTest() で片付けてから、もう一度実行してください。',
      ].join('\n');
    }

    const sorted = rows.slice().sort((a, b) => (a.caseNo < b.caseNo ? -1 : 1));
    const dbValues = sorted.map(buildCaseRowValues_);
    const uiValues = sorted.filter(r => r.kind === MIGRATION_KIND.OPEN).map(buildCaseRowValues_);

    writeCaseRows_(dbSheet, CASE_DATA_START_ROW, dbValues);
    writeCaseRows_(uiSheet, CASE_UI_DATA_START_ROW, uiValues);

    // 次に採番される番号が、取り込んだ番号と重ならないようにする
    const period = getCurrentPeriodNumber_();
    const maxSeq = Math.max(...sorted.map(r => Number(r.caseNo.split('-')[1])));
    PropertiesService.getScriptProperties().setProperty(seqKey_('CASE', period), String(maxSeq));
    SpreadsheetApp.flush();

    const summary = summarizeMigrationRows_(rows);
    summary.push(`案件番号カウンター: ${maxSeq}（次の新規案件は ${period}-${pad3_(maxSeq + 1)}）`);
    appendOperationLog_('', '旧一覧からの案件移行', summary.join(' / '), false);
    return ['取り込みが完了しました。', '', ...summary].join('\n');
  });
}

/**
 * 移行データシートを読み、1行ずつ点検する。
 * @return {{rows: Object[], errors: string[]}} errors が空なら rows はそのまま取り込める
 */
function readAndValidateMigrationRows_() {
  const sheet = getMainSpreadsheet_().getSheetByName(MIGRATION_SHEET_NAME);
  if (!sheet) {
    return { rows: [], errors: [`シート「${MIGRATION_SHEET_NAME}」が見つかりません。`] };
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < MIGRATION_DATA_START_ROW) return { rows: [], errors: [] };

  const period = getCurrentPeriodNumber_();
  const caseNoPattern = new RegExp(`^${period}-\\d{3}$`);
  const clientLabels = toLookup_(listClientOptionLabels_());
  const staffNames = toLookup_(listStaffNames_());
  const scheduleOptions = toLookup_(buildScheduleOptionList_());
  // 完了・中止の案件は今のプルダウンの範囲より前の月になるため、月だけの表記（前半/後半なし）も認める
  const closedSchedulePattern = /^\d{4}\/\d{2}( 前半| 後半)?$/;
  const kinds = Object.keys(MIGRATION_KIND).map(key => MIGRATION_KIND[key]);

  const values = sheet.getRange(MIGRATION_DATA_START_ROW, 1,
    lastRow - MIGRATION_DATA_START_ROW + 1, MIGRATION_LAST_COL).getDisplayValues();
  const rows = [];
  const errors = [];
  const seen = {};

  values.forEach((raw, i) => {
    const get = col => String(raw[col - 1] == null ? '' : raw[col - 1]).trim();
    const row = {
      caseNo: get(MIGRATION_COLS.CASE_NO),
      clientName: get(MIGRATION_COLS.CLIENT_NAME),
      caseName: get(MIGRATION_COLS.CASE_NAME),
      endSchedule: get(MIGRATION_COLS.END_SCHEDULE),
      billingSchedule: get(MIGRATION_COLS.BILLING_SCHEDULE),
      staffInCharge: get(MIGRATION_COLS.STAFF_IN_CHARGE),
      kind: get(MIGRATION_COLS.KIND),
    };
    if (Object.keys(row).every(key => !row[key])) return; // 空行は読み飛ばす

    const problems = [];
    if (!caseNoPattern.test(row.caseNo)) {
      problems.push(`案件番号「${row.caseNo}」は ${period}-001 の形式ではありません`);
    } else if (seen[row.caseNo]) {
      problems.push(`案件番号が${seen[row.caseNo]}行目と重複しています`);
    }
    if (kinds.indexOf(row.kind) === -1) {
      problems.push(`移行区分「${row.kind}」は ${kinds.map(k => `「${k}」`).join('')} のいずれかにしてください`);
    }
    if (!row.caseName) problems.push('案件名が空欄です');
    if (row.kind === MIGRATION_KIND.OPEN) {
      if (!clientLabels[row.clientName]) {
        problems.push(`取引先名「${row.clientName}」が取引先名の選択肢にありません`);
      }
      if (!staffNames[row.staffInCharge]) {
        problems.push(`担当「${row.staffInCharge}」が社員DB（メールアドレス登録済み）にいません`);
      }
      [['終了予定', row.endSchedule], ['請求予定', row.billingSchedule]].forEach(([label, value]) => {
        if (!scheduleOptions[value]) problems.push(`${label}「${value}」が今の選択肢にありません`);
      });
    } else {
      if (!row.clientName) problems.push('取引先名が空欄です');
      [['終了予定', row.endSchedule], ['請求予定', row.billingSchedule]].forEach(([label, value]) => {
        if (value && !scheduleOptions[value] && !closedSchedulePattern.test(value)) {
          problems.push(`${label}「${value}」は「2026/05」または選択肢の形式にしてください`);
        }
      });
    }

    const sheetRow = MIGRATION_DATA_START_ROW + i;
    if (row.caseNo && !seen[row.caseNo]) seen[row.caseNo] = sheetRow;
    if (problems.length) {
      errors.push(`${sheetRow}行目（${row.caseNo || '案件番号なし'}）: ${problems.join(' / ')}`);
    } else {
      rows.push(row);
    }
  });

  return { rows, errors };
}

/** 移行データ1件を、案件シート（表示・全案件DB共通の列構成）の1行分の値にする */
function buildCaseRowValues_(row) {
  const values = new Array(CASE_LAST_COL).fill('');
  const set = (col, value) => { values[col - 1] = value; };
  set(CASE_COLS.CASE_NO, row.caseNo);
  set(CASE_COLS.CLIENT_NAME, row.clientName);
  set(CASE_COLS.CASE_NAME, row.caseName);
  set(CASE_COLS.END_SCHEDULE, row.endSchedule);
  set(CASE_COLS.BILLING_SCHEDULE, row.billingSchedule);
  set(CASE_COLS.STAFF_IN_CHARGE, row.staffInCharge);
  if (row.kind === MIGRATION_KIND.CLOSED) {
    set(CASE_COLS.STATUS, STATUS.FINAL_APPROVED);
    set(CASE_COLS.BILLING_STATUS, BILLING_STATUS.FINAL_APPROVED);
    set(CASE_COLS.FINAL_APPROVER, MIGRATION_FINAL_APPROVER_LABEL);
  } else if (row.kind === MIGRATION_KIND.CANCELLED) {
    set(CASE_COLS.STATUS, STATUS.CANCELLED);
  } else {
    set(CASE_COLS.STATUS, STATUS.CASE_REGISTERED);
    set(CASE_COLS.BILLING_STATUS, BILLING_STATUS.NOT_BILLED);
  }
  return values;
}

/** 案件シートへ行をまとめて書き込む（シートの行数が足りなければ末尾に行を足す） */
function writeCaseRows_(sheet, startRow, values) {
  if (!values.length) return;
  const shortage = startRow + values.length - 1 - sheet.getMaxRows();
  if (shortage > 0) sheet.insertRowsAfter(sheet.getMaxRows(), shortage);
  sheet.getRange(startRow, 1, values.length, CASE_LAST_COL).setValues(values);
}

/** 案件シートのデータ範囲で、A〜F列（案件番号〜担当）のどこかに値がある行数を数える */
function countFilledCaseRows_(sheet, dataStartRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) return 0;
  return sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, CASE_COLS.STAFF_IN_CHARGE)
    .getValues()
    .filter(row => row.some(value => String(value).trim() !== ''))
    .length;
}

function summarizeMigrationRows_(rows) {
  const count = kind => rows.filter(r => r.kind === kind).length;
  return [
    `取り込む案件: ${rows.length}件`,
    `  ${MIGRATION_KIND.OPEN}: ${count(MIGRATION_KIND.OPEN)}件（表示シート・全案件DBへ登録）`,
    `  ${MIGRATION_KIND.CLOSED}: ${count(MIGRATION_KIND.CLOSED)}件（全案件DBにのみ最終承認済みで記録）`,
    `  ${MIGRATION_KIND.CANCELLED}: ${count(MIGRATION_KIND.CANCELLED)}件（全案件DBにのみ中止で記録）`,
  ];
}

function toLookup_(list) {
  const lookup = {};
  list.forEach(item => { lookup[String(item).trim()] = true; });
  return lookup;
}

function showMigrationMessage_(title, message) {
  Logger.log(message);
  try {
    SpreadsheetApp.getUi().alert(`${title}\n\n${message}`);
  } catch (e) {
    // スクリプトエディタから実行した場合はダイアログを出せないため、実行ログで確認する
  }
}
