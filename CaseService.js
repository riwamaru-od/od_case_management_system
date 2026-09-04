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
 * 採番後に変更されたら、操作ログへ記録して全案件DBへ反映する対象の列。
 * これらは案件の進行や請求のタイミングを左右するため、いつ誰が変えたかを追えるようにする。
 */
function getTrackedChangeCols_() {
  return [CASE_COLS.END_SCHEDULE, CASE_COLS.BILLING_SCHEDULE, CASE_COLS.STAFF_IN_CHARGE];
}

/**
 * onEdit から呼ばれる。採番済みの案件について、終了予定・請求予定・担当が
 * 変更された場合に、変更内容を操作ログへ記録し、全案件DBへ反映する。
 * （未採番の行は「変更」ではなく新規入力の途中なので対象外。採番時の
 *   handleCaseRowEdit_ 側で改めてDBへ同期される）
 *
 * @param {Sheet} sheet メイン画面UIシート
 * @param {number} row 編集された行
 * @param {number[]} changedCols 変更された対象列（getTrackedChangeCols_ のうち編集範囲に含まれるもの）
 * @param {Object} e onEdit のイベントオブジェクト
 */
function handleTrackedFieldEdit_(sheet, row, changedCols, e) {
  const caseNo = String(sheet.getRange(row, CASE_COLS.CASE_NO).getValue() || '').trim();
  if (!caseNo) return;

  // 変更前の値は単一セルの編集時のみ取得できる（複数セルの貼り付け時は e.oldValue が来ない）
  const isSingleCell = e && e.range && e.range.getNumRows() === 1 && e.range.getNumColumns() === 1;
  const formatValue = val => {
    if (val instanceof Date) return formatDateTime_(val);
    const text = String(val == null ? '' : val).trim();
    return text === '' ? '（空欄）' : text;
  };

  const details = changedCols.map(col => {
    const label = CASE_HEADERS[col - 1];
    const after = formatValue(sheet.getRange(row, col).getValue());
    if (!isSingleCell) return `${label}: ${after}`;
    return `${label}: ${formatValue(e.oldValue)} → ${after}`;
  });

  withLock_('案件情報の変更記録', () => {
    syncCaseToDb_(caseNo);
    // 実際に編集した本人（e.user）が取得できる場合はその人を実行者として記録する。
    // ただしGoogle Workspaceドメインが無い環境では e.user も Session.getActiveUser() も
    // 空になり、編集者を特定できない。その場合は自動処理と誤解されないよう、
    // 「システム(自動実行)」ではなく人の操作であることが分かる名称で記録する。
    const editorEmail = getEditorEmailFromEvent_(e);
    ACTIVE_USER_EMAIL_OVERRIDE_ = editorEmail || null;
    try {
      appendOperationLog_(caseNo, '案件情報の変更', details.join(' / '), false, SHEET_EDITOR_UNKNOWN_ACTOR_NAME);
    } finally {
      ACTIVE_USER_EMAIL_OVERRIDE_ = null;
    }
  }, caseNo);
}

/** onEdit イベントから、実際に編集した利用者のメールアドレスを取得する（取れない場合は空文字） */
function getEditorEmailFromEvent_(e) {
  try {
    return (e && e.user && e.user.getEmail()) || '';
  } catch (err) {
    console.warn(`編集者のメールアドレスを取得できませんでした: ${err}`);
    return '';
  }
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

    const caseNo = generateUniqueCaseNo_(getCurrentPeriodNumber_());

    sheet.getRange(row, CASE_COLS.CASE_NO).setValue(caseNo);
    sheet.getRange(row, CASE_COLS.STATUS).setValue(STATUS.CASE_REGISTERED);
    sheet.getRange(row, CASE_COLS.BILLING_STATUS).setValue(BILLING_STATUS.NOT_BILLED);
    // ロックを手放す前に書き込みを確定させる。これを行わないと、書き込みが
    // バッファに残ったままロックが解放され、次に採番へ入ってきた実行が
    // 「まだ未採番」と誤認して同じ行・同じ番号を二重に採番しうる。
    SpreadsheetApp.flush();

    syncCaseToDb_(caseNo);

    appendOperationLog_(caseNo, '案件登録（自動採番）', '', false);
  });
}

/**
 * この期の未使用の案件番号を発行する。
 *
 * 通常はカウンター（スクリプトプロパティ）を1つ進めるだけで足りるが、
 * カウンターと実際のシートの状態がずれると番号が重複しうる
 * （テストデータのリセットでカウンターだけ戻した、複数の担当者が
 *  onEditトリガーを登録していて採番が二重に走った、など）。
 * そのため発行後に既存の番号と突き合わせ、使用済みなら次の番号を取り直す。
 */
function generateUniqueCaseNo_(period) {
  const used = collectUsedCaseNos_();
  for (let attempt = 0; attempt < 1000; attempt++) {
    const caseNo = `${period}-${pad3_(nextSequence_(seqKey_('CASE', period)))}`;
    if (!used[caseNo]) return caseNo;
    console.warn(`案件番号 ${caseNo} は既に使われているため、次の番号を採番します。`);
  }
  throw AppError_('CASE_NO_UNAVAILABLE',
    '案件番号を採番できませんでした。番号が重複している可能性があります。管理者にお問い合わせください。');
}

/** メイン画面UI・全案件DBの両シートで既に使われている案件番号を集める */
function collectUsedCaseNos_() {
  const used = {};
  [getActiveUiSheet_(), getActiveDbSheet_()].forEach(sheet => {
    const lastRow = sheet.getLastRow();
    if (lastRow < CASE_DATA_START_ROW) return;
    sheet.getRange(CASE_DATA_START_ROW, CASE_COLS.CASE_NO, lastRow - CASE_DATA_START_ROW + 1, 1)
      .getValues()
      .forEach(rowValues => {
        const value = String(rowValues[0] == null ? '' : rowValues[0]).trim();
        if (value) used[value] = true;
      });
  });
  return used;
}

/**
 * 案件情報を読み取る（UID管理の正式な読み取り元）。
 * getMainSpreadsheet_() 経由で取得したシートを使い、
 * SpreadsheetApp.getActiveSheet() には一切依存しない。
 * まずメイン画面UIシートから探し、見つからなければ全案件DBシートを探す
 * （最終承認・案件中止済みの案件はUIシートから削除済みだが、全案件DBには残っているため。
 * サイドバー側はこれにより、終了済みの案件でも履歴閲覧ができる。書き込み系の関数
 * （setCaseFields_ 等）はUIシートのみを対象とするため、終了済み案件への誤操作は
 * computeButtonStates_ 側で全ボタンを非活性にすることで防ぐ）。
 */
function getCaseInfo_(caseNo) {
  const uiSheet = getActiveUiSheet_();
  let sheet = uiSheet;
  let row = findRowByCaseNo_(uiSheet, caseNo);
  if (!row) {
    const dbSheet = getActiveDbSheet_();
    row = findRowByCaseNo_(dbSheet, caseNo);
    sheet = dbSheet;
  }
  if (!row) {
    throw AppError_('CASE_NOT_FOUND', `案件番号「${caseNo}」が見つかりません。`);
  }
  const v = sheet.getRange(row, 1, 1, CASE_LAST_COL).getValues()[0];
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
    // 差し戻し済みフラグ（内部用）。空でなければ「差し戻し後、未再作成」の状態を表す。
    quoteRejectedAt: get(CASE_COLS.QUOTE_REJECTED_AT),
    invoiceRejectedAt: get(CASE_COLS.INVOICE_REJECTED_AT),
    // 再承認待ちフラグ（内部用）。空でなければ「差し戻し・再作成されて以降、まだ再承認
    // されていない」状態を表す。PDF出力ボタンの活性判定に使う（過去の承認記録は
    // quoteApprovedAt 等にそのまま残るため、履歴表示には影響しない）。
    quoteReapprovalPending: get(CASE_COLS.QUOTE_REAPPROVAL_PENDING),
    invoiceReapprovalPending: get(CASE_COLS.INVOICE_REAPPROVAL_PENDING),
    // 着手日時（書類ファイルを作成・再作成した日時）。作成中のまま放置されている
    // 書類の経過日数判定に使う（ReportService.gs）。
    quoteStartedAt: get(CASE_COLS.QUOTE_STARTED_AT),
    invoiceStartedAt: get(CASE_COLS.INVOICE_STARTED_AT),
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

/**
 * 案件中止: 社員は誰でも、案件が途中の状態からでも中止できる（承認ロール等の制限なし）。
 * ステータスを「中止」にし、表示用シート（メイン画面UI）からは削除する（全案件DBには残す）。
 * comment はサイドバーの確認モーダルで入力された中止理由（任意）。
 * 最終承認の削除処理（removeCaseFromUiAfterFinalApproval_）とは独立した処理として持つ。
 */
function cancelCaseForCase_(caseNo, comment) {
  // 見積書が未作成の場合は、中止の記録として先に見積書を作成しておく（作成のみ。
  // 完成＝承認依頼は行わない）。createDocumentForCase_ は自身で withLock_ を取得するため、
  // ロックの入れ子を避けて中止処理のロックを取る前に呼び出す。
  let autoCreatedQuote = false;
  if (!getCaseInfo_(caseNo).quoteLink) {
    createDocumentForCase_('quote', caseNo);
    autoCreatedQuote = true;
  }

  return withLock_('案件中止', () => {
    setCaseFields_(caseNo, { [CASE_COLS.STATUS]: STATUS.CANCELLED });

    // 全書類の案件フォルダを「中止案件」フォルダ配下へ移動する
    // （UI行を消す前に、案件名を含むフォルダ名を解決できる状態で行う）
    const movedLabels = moveCaseDocFoldersToCancelled_(getCaseInfo_(caseNo));

    const uiSheet = getActiveUiSheet_();
    syncCaseToDb_(caseNo); // 念のため最新状態をDBへバックアップ
    const row = findRowByCaseNo_(uiSheet, caseNo);
    if (row) {
      uiSheet.deleteRow(row);
    }

    const details = [comment || ''];
    if (autoCreatedQuote) details.push('見積書が未作成だったため自動作成しました');
    if (movedLabels.length) details.push(`中止案件フォルダへ移動: ${movedLabels.join('・')}`);
    appendOperationLog_(caseNo, '案件中止', details.filter(Boolean).join(' / '), false);

    return { status: STATUS.CANCELLED };
  }, caseNo);
}

/**
 * メイン画面UIシートの案件行を、案件番号の昇順へ並べ直す。
 *
 * 中止・最終承認を取り消した案件は末尾へ書き戻されるため、そのままだと
 * 番号順が崩れる。行の挿入や手動での並べ替えでも同様に崩れうる。
 *
 * 実際に並びが崩れているときだけ並べ替える。毎回無条件に並べ替えると、
 * 利用者が新しい案件を入力している最中に行が動いてしまうため。
 * 対象は採番済みの行が並ぶ範囲までで、その下にある入力途中の行は動かさない。
 *
 * @return {boolean} 並べ替えを実施した場合 true
 */
function sortUiSheetByCaseNoIfNeeded_() {
  const sheet = getActiveUiSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < CASE_UI_DATA_START_ROW) return false;

  const numRows = lastRow - CASE_UI_DATA_START_ROW + 1;
  const caseNos = sheet.getRange(CASE_UI_DATA_START_ROW, CASE_COLS.CASE_NO, numRows, 1)
    .getValues()
    .map(rowValues => String(rowValues[0] == null ? '' : rowValues[0]).trim());

  // 採番済みの行が現れる最後の位置を探し、そこまでを並べ替えの対象にする
  let lastNumberedIndex = -1;
  for (let i = numRows - 1; i >= 0; i--) {
    if (caseNos[i] !== '') { lastNumberedIndex = i; break; }
  }
  if (lastNumberedIndex < 1) return false; // 採番済みが1行以下なら並べ替える必要が無い

  const target = caseNos.slice(0, lastNumberedIndex + 1);
  if (isSortedByCaseNoAscending_(target)) return false;

  sheet.getRange(CASE_UI_DATA_START_ROW, 1, target.length, CASE_LAST_COL)
    .sort({ column: CASE_COLS.CASE_NO, ascending: true });
  SpreadsheetApp.flush();
  return true;
}

/**
 * 案件番号の並びが昇順になっているかを判定する。
 * 案件番号は「{期}-{3桁の連番}」でゼロ埋めされており、1枚のシートには
 * 同じ期の案件しか入らないため、文字列としての比較で正しく順序を判定できる。
 * 空欄（採番前の行）は末尾にあるのが正しい並びとみなす。
 */
function isSortedByCaseNoAscending_(caseNos) {
  for (let i = 1; i < caseNos.length; i++) {
    const previous = caseNos[i - 1];
    const current = caseNos[i];
    if (previous === '' && current !== '') return false;
    if (previous !== '' && current !== '' && previous > current) return false;
  }
  return true;
}
