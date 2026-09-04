/**
 * CaseRestoreService.gs
 * 「案件中止」「最終承認」を取り消して、案件を進行中の状態へ戻す。
 *
 * 中止・最終承認された案件はメイン画面UIシートから削除され、全案件DBにのみ残る。
 * 復元はその逆で、全案件DBの内容をメイン画面UIシートへ書き戻し、ステータスを
 * 中止・最終承認の直前の状態へ戻す。あわせて、中止時に「中止案件」フォルダへ
 * 移動した書類のフォルダも元の場所へ戻す。
 *
 * 操作は全案件DBシートで対象行を選択したうえで、サイドバーの「元に戻す」から行う
 * （中止・最終承認済みの案件はメイン画面UIシートに存在しないため）。
 */

/**
 * 中止または最終承認を取り消し、案件を進行中へ戻す。
 * @param {string} caseNo 案件番号
 * @param {string} comment 取り消し理由（任意）
 */
function restoreCaseForCase_(caseNo, comment) {
  return withLock_('案件の復元', () => {
    const email = getActiveUserEmail_();
    assertRole_(email, REQUIRED_ROLES.FINAL_APPROVE, '案件の復元');

    const caseInfo = getCaseInfo_(caseNo);
    const previousStatus = caseInfo.status;
    if (previousStatus !== STATUS.CANCELLED && previousStatus !== STATUS.FINAL_APPROVED) {
      throw AppError_('INVALID_STATE',
        `この案件は中止・最終承認済みではないため、元に戻す操作はできません（現在のステータス: ${previousStatus}）。`);
    }

    // 1. メイン画面UIシートへ行を書き戻す（setCaseFields_ はUIシートを対象とするため先に行う）
    restoreCaseRowToUi_(caseNo);

    // 2. 書類の進捗から、中止・最終承認の直前のステータスを復元する
    const restoredStatus = inferStatusFromDocumentProgress_(caseInfo);
    const fieldUpdates = { [CASE_COLS.STATUS]: restoredStatus };

    if (previousStatus === STATUS.FINAL_APPROVED) {
      // 最終承認は取り消されたので、その記録を消し、請求ステータスを請求済みへ戻す
      // （最終承認は請求済みの案件にのみ行えるため、直前は必ず「請求済み」だった）
      fieldUpdates[CASE_COLS.BILLING_STATUS] = BILLING_STATUS.BILLED;
      fieldUpdates[CASE_COLS.FINAL_APPROVER] = '';
      fieldUpdates[CASE_COLS.FINAL_APPROVED_AT] = '';
    }
    setCaseFields_(caseNo, fieldUpdates);

    // 3. 末尾へ書き戻したことで崩れた並びを、案件番号順へ戻す
    sortUiSheetByCaseNoIfNeeded_();

    // 4. 中止時に「中止案件」フォルダへ移した書類を元の場所へ戻す
    const movedLabels = previousStatus === STATUS.CANCELLED
      ? moveCaseDocFoldersFromCancelled_(caseInfo) : [];

    const details = [
      `${previousStatus} → ${restoredStatus}`,
      comment || '',
      movedLabels.length ? `元のフォルダへ移動: ${movedLabels.join('・')}` : '',
    ].filter(Boolean).join(' / ');
    appendOperationLog_(caseNo, previousStatus === STATUS.CANCELLED ? '案件中止の取り消し' : '最終承認の取り消し',
      details, false);

    return { status: restoredStatus };
  }, caseNo);
}

/**
 * 全案件DBの内容を、メイン画面UIシートの最終行へ書き戻す。
 * 既にUIシートに存在する場合は何もしない。
 */
function restoreCaseRowToUi_(caseNo) {
  const uiSheet = getActiveUiSheet_();
  if (findRowByCaseNo_(uiSheet, caseNo)) return;

  const dbSheet = getActiveDbSheet_();
  const dbRow = findRowByCaseNo_(dbSheet, caseNo);
  if (!dbRow) {
    throw AppError_('CASE_NOT_FOUND', `案件番号「${caseNo}」が全案件DBに見つかりません。`);
  }

  const values = dbSheet.getRange(dbRow, 1, 1, CASE_LAST_COL).getValues();
  const targetRow = Math.max(uiSheet.getLastRow() + 1, CASE_UI_DATA_START_ROW);
  uiSheet.getRange(targetRow, 1, 1, CASE_LAST_COL).setValues(values);
  SpreadsheetApp.flush();
}

/**
 * 書類の進捗から、中止・最終承認される直前のステータスを推定する。
 * 請求書 → 見積書 の順に、進んでいる段階を優先して判定する。
 * 差し戻し中・再承認待ちの書類は「作成中」として扱う。
 */
function inferStatusFromDocumentProgress_(caseInfo) {
  if (caseInfo.invoiceLink) {
    if (caseInfo.invoiceRejectedAt || caseInfo.invoiceReapprovalPending) return STATUS.INVOICE_IN_PROGRESS;
    if (caseInfo.invoiceApprovedAt) return STATUS.INVOICE_APPROVED;
    if (caseInfo.invoiceCreatedAt) return STATUS.INVOICE_DRAFTED;
    return STATUS.INVOICE_IN_PROGRESS;
  }
  if (caseInfo.quoteLink) {
    if (caseInfo.quoteRejectedAt || caseInfo.quoteReapprovalPending) return STATUS.QUOTE_IN_PROGRESS;
    if (caseInfo.quoteApprovedAt) return STATUS.QUOTE_APPROVED;
    if (caseInfo.quoteCreatedAt) return STATUS.QUOTE_DRAFTED;
    return STATUS.QUOTE_IN_PROGRESS;
  }
  return STATUS.CASE_REGISTERED;
}

/**
 * 「中止案件」フォルダへ移した案件フォルダを、書類種別ごとの元の場所へ戻す。
 * 請求書のみ、承認済みであれば「請求中案件」、それ以外は「未請求案件」へ戻す
 * （承認時に請求中案件へ移動する運用に合わせるため）。
 * @return {string[]} 戻した書類種別のラベル（操作ログ用）
 */
function moveCaseDocFoldersFromCancelled_(caseInfo) {
  const periodNumber = getCurrentPeriodNumber_();
  const targetName = caseFolderName_(caseInfo);
  const movedLabels = [];

  ['quote', 'invoice', 'delivery'].forEach(key => {
    const docType = DOC_TYPES[key];
    try {
      const cancelledFolder = getDocumentFolder_(docType.folderKind, periodNumber, SUBFOLDER.CANCELLED);
      const found = cancelledFolder.getFoldersByName(targetName);
      if (!found.hasNext()) return;

      const stage = (key === 'invoice' && caseInfo.invoiceApprovedAt) ? 'billed' : 'created';
      const destination = getDocumentFolder_(docType.folderKind, periodNumber, docType.folderForStage(stage));
      while (found.hasNext()) {
        moveFileOrFolder_(found.next(), destination);
        movedLabels.push(docType.label);
      }
    } catch (e) {
      // 1種別の移動に失敗しても復元自体は止めない（気付けるよう操作ログへ残す）
      console.warn(`${docType.label}フォルダを元の場所へ戻せませんでした: ${e}`);
      appendOperationLog_(caseInfo.caseNo, '案件復元（フォルダ移動）',
        `${docType.label}の移動に失敗: ${e && e.message ? e.message : e}`, true);
    }
  });

  return movedLabels;
}
