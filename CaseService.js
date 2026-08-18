/**
 * CaseService.gs
 * 案件行（メイン画面UI / 全案件DB）に関する共通処理。
 *  - 新規案件の自動採番（取引先名・案件名・終了予定・請求予定・担当の5項目が揃った時点）
 *  - メイン画面UI → 全案件DB への同期（都度・片方向）
 *  - 最終承認後のメイン画面UIからの物理削除（全案件DBには残す）
 */

/** 案件番号生成に必要な5項目（列インデックス） */
function getCaseNumberingTriggerCols_() {
  return [
    CASE_COLS.CLIENT_NAME, CASE_COLS.CASE_NAME, CASE_COLS.END_SCHEDULE,
    CASE_COLS.BILLING_SCHEDULE, CASE_COLS.STAFF_IN_CHARGE,
  ];
}

/**
 * onEdit から呼ばれる。編集された行が「採番トリガー列」のいずれかで、
 * 5項目すべてが埋まっており、かつ案件番号が未採番であれば自動採番する。
 * 複数列がほぼ同時に埋まった場合はイベントごとに tryLock されるため、
 * 実際に採番されるのは「5項目が揃った最後の1回」のみになる。
 */
function handleCaseRowEdit_(sheet, row) {
  const rowValues = sheet.getRange(row, 1, 1, CASE_LAST_COL).getValues()[0];
  const already = rowValues[CASE_COLS.CASE_NO - 1];
  if (already) return; // 採番済み

  const allFilled = getCaseNumberingTriggerCols_().every(col => !!rowValues[col - 1]);
  if (!allFilled) return;

  withLock_('案件番号の自動採番', () => {
    // ロック取得後に再度未採番であることを確認（他イベントとの競合対策）
    const current = sheet.getRange(row, CASE_COLS.CASE_NO).getValue();
    if (current) return;

    const period = getCurrentPeriodNumber_();
    const seq = nextSequence_(seqKey_('CASE', period));
    const caseNo = `${period}-${pad3_(seq)}`;

    sheet.getRange(row, CASE_COLS.CASE_NO).setValue(caseNo);
    sheet.getRange(row, CASE_COLS.STATUS).setValue(STATUS.CASE_REGISTERED);
    sheet.getRange(row, CASE_COLS.BILLING_STATUS).setValue(BILLING_STATUS.NOT_BILLED);

    syncCaseToDb_(caseNo);
  });
}

/**
 * UIシートから案件情報を読み取る（UID管理の正式な読み取り元）。
 * getMainSpreadsheet_() 経由で取得したシートを使い、
 * SpreadsheetApp.getActiveSheet() には一切依存しない。
 */
function getCaseInfo_(caseNo) {
  const uiSheet = getActiveUiSheet_();
  const row = findRowByCaseNo_(uiSheet, caseNo);
  if (!row) {
    throw AppError_('CASE_NOT_FOUND', `案件番号「${caseNo}」が見つかりません。`);
  }
  const v = uiSheet.getRange(row, 1, 1, CASE_LAST_COL).getValues()[0];
   const formatCell = val => {
    if (val instanceof Date) return formatDateTime_(val);
    return val == null ? '' : String(val);
  };
  const get = col => formatCell(v[col - 1]);
  return {
    caseNo: get(CASE_COLS.CASE_NO),
    clientName: get(CASE_COLS.CLIENT_NAME),
    caseName: get(CASE_COLS.CASE_NAME),
    endSchedule: get(CASE_COLS.END_SCHEDULE),
    billingSchedule: get(CASE_COLS.BILLING_SCHEDULE),
    staffInCharge: get(CASE_COLS.STAFF_IN_CHARGE),
    status: get(CASE_COLS.STATUS),
    billingStatus: get(CASE_COLS.BILLING_STATUS),
    quoteLink: get(CASE_COLS.QUOTE_LINK),
    quoteOutputLink: get(CASE_COLS.QUOTE_OUTPUT_LINK),
    invoiceLink: get(CASE_COLS.INVOICE_LINK),
    invoiceOutputLink: get(CASE_COLS.INVOICE_OUTPUT_LINK),
    deliveryLink: get(CASE_COLS.DELIVERY_LINK),
    deliveryOutputLink: get(CASE_COLS.DELIVERY_OUTPUT_LINK),
    // 作成者・承認者・出力者などの記録列（承認後の通知先解決・サイドバー履歴表示に使用）
    quoteCreator: get(CASE_COLS.QUOTE_CREATOR),
    quoteCreatedAt: get(CASE_COLS.QUOTE_CREATED_AT),
    quoteApprover: get(CASE_COLS.QUOTE_APPROVER),
    quoteApprovedAt: get(CASE_COLS.QUOTE_APPROVED_AT),
    quoteOutputBy: get(CASE_COLS.QUOTE_OUTPUT_BY),
    quoteOutputAt: get(CASE_COLS.QUOTE_OUTPUT_AT),
    invoiceCreator: get(CASE_COLS.INVOICE_CREATOR),
    invoiceCreatedAt: get(CASE_COLS.INVOICE_CREATED_AT),
    invoiceApprover: get(CASE_COLS.INVOICE_APPROVER),
    invoiceApprovedAt: get(CASE_COLS.INVOICE_APPROVED_AT),
    invoiceOutputBy: get(CASE_COLS.INVOICE_OUTPUT_BY),
    invoiceOutputAt: get(CASE_COLS.INVOICE_OUTPUT_AT),
    deliveryCreator: get(CASE_COLS.DELIVERY_CREATOR),
    deliveryCreatedAt: get(CASE_COLS.DELIVERY_CREATED_AT),
    finalApprover: get(CASE_COLS.FINAL_APPROVER),
    finalApprovedAt: get(CASE_COLS.FINAL_APPROVED_AT),
    _row: row,
  };
}

/** UI行の内容をキー(案件番号)に基づき全案件DBへ反映する（無ければ追記、あれば上書き） */
function syncCaseToDb_(caseNo) {
  const uiSheet = getActiveUiSheet_();
  const dbSheet = getActiveDbSheet_();

  const uiRow = findRowByCaseNo_(uiSheet, caseNo);
  if (!uiRow) return;
  const values = uiSheet.getRange(uiRow, 1, 1, CASE_LAST_COL).getValues();

  const dbRow = findRowByCaseNo_(dbSheet, caseNo);
  if (dbRow) {
    dbSheet.getRange(dbRow, 1, 1, CASE_LAST_COL).setValues(values);
  } else {
    dbSheet.getRange(dbSheet.getLastRow() + 1, 1, 1, CASE_LAST_COL).setValues(values);
  }
}

/** 複数列をまとめて更新する場合（同期は最後に1回だけ実行） */
function setCaseFields_(caseNo, fieldMap) {
  const uiSheet = getActiveUiSheet_();
  const uiRow = findRowByCaseNo_(uiSheet, caseNo);
  if (!uiRow) {
    throw AppError_('CASE_NOT_FOUND', `案件番号「${caseNo}」がメイン画面に見つかりません。`);
  }
  Object.keys(fieldMap).forEach(colIndex => {
    uiSheet.getRange(uiRow, Number(colIndex)).setValue(fieldMap[colIndex]);
  });
  syncCaseToDb_(caseNo);
}

/** 案件番号から、UI・DB双方の該当セルへ同一の値を書き込む（1つの列だけ更新したい場合に使用） */
function setCaseField_(caseNo, colIndex, value) {
  setCaseFields_(caseNo, { [colIndex]: value });
}

/**
 * 最終承認完了後の後処理（確定仕様2章）:
 *   - 全案件DBには残したまま、メイン画面UIから当該行を物理削除する。
 *   - 削除前に必ず syncCaseToDb_ でDBへ反映してからでないと、UI固有の最新情報が失われるため注意。
 */
function removeCaseFromUiAfterFinalApproval_(caseNo) {
  const uiSheet = getActiveUiSheet_();
  syncCaseToDb_(caseNo); // 念のため最新状態をDBへバックアップ
  const row = findRowByCaseNo_(uiSheet, caseNo);
  if (row) {
    uiSheet.deleteRow(row);
  }
}
