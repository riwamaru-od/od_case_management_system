/**
 * ApprovalService.gs
 * 見積書・請求書に共通する承認フロー（作成→完成→承認依頼→承認/差し戻し/再作成→出力）を
 * DOC_TYPES の設定を使って汎用的に実装する。納品書には承認ステップが無いため、
 * DeliveryService.gs 側で必要な部分のみ個別に実装している。
 */

/** 見積書 or 請求書 を新規作成する（テンプレートからコピーし、ヘッダー情報を転記） */
function createDocumentForCase_(docTypeKey, caseNo) {
  return withLock_(`${DOC_TYPES[docTypeKey].label}の作成`, () => {
    const docType = DOC_TYPES[docTypeKey];
    const caseInfo = getCaseInfo_(caseNo);

    const file = createLatestDocument_(docType, caseInfo, 'created');

    if (docTypeKey === 'quote') {
      fillQuoteDocument_(file, caseInfo);
    } else if (docTypeKey === 'invoice') {
      const quoteFileId = extractFileIdFromUrl_(caseInfo.quoteLink);
      fillInvoiceDocument_(file, caseInfo, quoteFileId);
    }

    const fieldUpdates = {
      [docType.col.link]: file.getUrl(),
      [CASE_COLS.STATUS]: docType.status.inProgress,
    };
    if (docType.col.rejectedAt) fieldUpdates[docType.col.rejectedAt] = '';
    setCaseFields_(caseNo, fieldUpdates);

    appendOperationLog_(caseNo, `${docType.label}作成`, `URL: ${file.getUrl()}`, false);

    return { url: file.getUrl() };
  }, caseNo);
}

/**
 * 「完成」ボタン: 必要事項の記載が終わったタイミングで押される。
 * 総務へ承認依頼を送信し、作成者・作成日時を記録して「作成済み」ステータスにする。
 * comment はダイアログで入力された承認依頼時コメント（テンプレートのN56/F56相当へ転記）。
 */
function completeDocumentForCase_(docTypeKey, caseNo, comment) {
  return withLock_(`${DOC_TYPES[docTypeKey].label}の完成`, () => {
    const docType = DOC_TYPES[docTypeKey];
    const caseInfo = getCaseInfo_(caseNo);
    const email = getActiveUserEmail_();
    const staff = findStaffByEmail_(email);
    const now = new Date();

    const fileId = extractFileIdFromUrl_(caseInfo[`${docTypeKey}Link`]);
    const file = DriveApp.getFileById(fileId);
    const sheet = getPrimarySheet_(file, docType);
    const cells = docType.cells();
    setCellValue_(sheet, cells.CREATOR_NAME, staff ? staff.name : email);
    setCellValue_(sheet, cells.CREATED_AT, formatDateTime_(now));
    if (comment) setCellValue_(sheet, cells.REQUEST_COMMENT, comment);

    const fieldUpdates = {
      [docType.col.creator]: staff ? staff.name : email,
      [docType.col.createdAt]: formatDateTime_(now),
      [CASE_COLS.STATUS]: docType.status.drafted,
    };
    if (docType.col.rejectedAt) fieldUpdates[docType.col.rejectedAt] = '';
    setCaseFields_(caseNo, fieldUpdates);

    const msg = buildApprovalRequestMessage_(docType.label, {
      caseNo: caseInfo.caseNo, caseName: caseInfo.caseName, clientName: caseInfo.clientName,
      requesterName: staff ? staff.name : email, docUrl: file.getUrl(),
    }, comment);
    notifyAdminDept_(msg.subject, msg.body, msg.htmlBody);

    appendOperationLog_(caseNo, `${docType.label}完成（承認依頼）`, comment || '', false);

    return { status: docType.status.drafted };
  }, caseNo);
}

/** 承認: 総務担当（該当ロール保持者）のみ実行可能 */
function approveDocumentForCase_(docTypeKey, caseNo, comment) {
  return withLock_(`${DOC_TYPES[docTypeKey].label}の承認`, () => {
    const docType = DOC_TYPES[docTypeKey];
    const email = getActiveUserEmail_();
    assertRole_(email, docType.approverRoles, `${docType.label}承認`);

    const caseInfo = getCaseInfo_(caseNo);
    const staff = findStaffByEmail_(email);
    const now = new Date();

    const fileId = extractFileIdFromUrl_(caseInfo[`${docTypeKey}Link`]);
    const file = DriveApp.getFileById(fileId);
    const sheet = getPrimarySheet_(file, docType);
    const cells = docType.cells();
    setCellValue_(sheet, cells.APPROVER_NAME, staff ? staff.name : email);
    setCellValue_(sheet, cells.APPROVED_AT, formatDateTime_(now));
    if (comment) setCellValue_(sheet, cells.APPROVE_COMMENT, comment);

    // 承認されたシートを保護（編集不可にする）
   try {
    protectSheet_(sheet);
    } catch (e) {
    console.warn(`シート保護(protectSheet_)の適用に失敗しました: ${e}`);
    }
    try {
    insertSealImage_(file, cells, docType);
    } catch (e) {
    console.warn(`社印画像の挿入に失敗しました: ${e}`);
    }


    setCaseFields_(caseNo, {
      [docType.col.approver]: staff ? staff.name : email,
      [docType.col.approvedAt]: formatDateTime_(now),
      [CASE_COLS.STATUS]: docType.status.approved,
    });

    if (docTypeKey === 'invoice') {
      // 請求書は承認済み以降、フォルダを「未請求案件」→「請求中案件」へ移動
      moveCaseDocFolderToStage_(docType, caseInfo, 'created', 'billed');
    }

    // 作成者へ通知
    const creatorStaff = getAllStaff_().find(s => s.name === caseInfo[`${docTypeKey}Creator`]);
    const notifyMsg = buildApprovedMessage_(docType.label, {
      caseNo: caseInfo.caseNo, caseName: caseInfo.caseName, docUrl: file.getUrl(),
    }, comment);
    if (creatorStaff) notifyStaff_(creatorStaff.email, notifyMsg.subject, notifyMsg.body, notifyMsg.htmlBody);

    appendOperationLog_(caseNo, `${docType.label}承認`, comment || '', false);

    return { status: docType.status.approved };
  }, caseNo);
}

/** 差し戻し: 総務担当が承認せず、作成中に戻す */
function rejectDocumentForCase_(docTypeKey, caseNo, comment) {
  return withLock_(`${DOC_TYPES[docTypeKey].label}の差し戻し`, () => {
    const docType = DOC_TYPES[docTypeKey];
    const email = getActiveUserEmail_();
    assertRole_(email, docType.approverRoles, `${docType.label}承認`);

    const caseInfo = getCaseInfo_(caseNo);
    const fileId = extractFileIdFromUrl_(caseInfo[`${docTypeKey}Link`]);
    const file = DriveApp.getFileById(fileId);
    const sheet = getPrimarySheet_(file, docType);
    if (comment) {
      setCellValue_(sheet, docType.cells().APPROVE_COMMENT, `[差し戻し] ${comment}`);
    }

    // 差し戻された書類のシートを保護（編集不可にする）。承認時と同様、
    // 処理全体を失敗させないよう警告ログに留める。
    try {
      protectSheet_(sheet);
    } catch (e) {
      console.warn(`シート保護(protectSheet_)の適用に失敗しました: ${e}`);
    }

    const fieldUpdates = { [CASE_COLS.STATUS]: docType.status.inProgress };
    if (docType.col.rejectedAt) fieldUpdates[docType.col.rejectedAt] = formatDateTime_(new Date());
    setCaseFields_(caseNo, fieldUpdates);

    appendOperationLog_(caseNo, `${docType.label}差し戻し`, comment || '', false);

    return { status: docType.status.inProgress };
  }, caseNo);
}

/** 再作成: 同一ファイル内で旧版シートをロック・退避し、新シートを「最新」として作り直す */
function recreateDocumentForCase_(docTypeKey, caseNo) {
  return withLock_(`${DOC_TYPES[docTypeKey].label}の再作成`, () => {
    const docType = DOC_TYPES[docTypeKey];
    const email = getActiveUserEmail_();
    assertRole_(email, docType.approverRoles, `${docType.label}承認`);

    const caseInfo = getCaseInfo_(caseNo);
    const file = recreateLatestDocument_(docType, caseInfo);

    if (docTypeKey === 'quote') {
      fillQuoteDocument_(file, caseInfo);
    } else if (docTypeKey === 'invoice') {
      const quoteFileId = extractFileIdFromUrl_(caseInfo.quoteLink);
      fillInvoiceDocument_(file, caseInfo, quoteFileId);
    }

    const fieldUpdates = {
      [docType.col.link]: file.getUrl(),
      [CASE_COLS.STATUS]: docType.status.inProgress,
    };
    if (docType.col.rejectedAt) fieldUpdates[docType.col.rejectedAt] = '';
    setCaseFields_(caseNo, fieldUpdates);

    appendOperationLog_(caseNo, `${docType.label}再作成`, `URL: ${file.getUrl()}`, false);

    return { url: file.getUrl(), status: docType.status.inProgress };
  }, caseNo);
}

/** 印刷（PDFプレビューを新規タブで開く用のURLを返す。印刷者/印刷日時をテンプレートへ記録） */
function printDocumentForCase_(docTypeKey, caseNo) {
  return withLock_(`${DOC_TYPES[docTypeKey].label}の印刷`, () => {
    const docType = DOC_TYPES[docTypeKey];
    const caseInfo = getCaseInfo_(caseNo);
    const email = getActiveUserEmail_();
    const staff = findStaffByEmail_(email);
    const now = new Date();

    const fileId = extractFileIdFromUrl_(caseInfo[`${docTypeKey}Link`]);
    const file = DriveApp.getFileById(fileId);
    const sheet = getPrimarySheet_(file, docType);
    const cells = docType.cells();
    setCellValue_(sheet, cells.PRINTED_BY, staff ? staff.name : email);
    setCellValue_(sheet, cells.PRINTED_AT, formatDateTime_(now));

    if (docType.col.outputBy) {
      setCaseFields_(caseNo, {
        [docType.col.outputBy]: staff ? staff.name : email,
        [docType.col.outputAt]: formatDateTime_(now),
      });
    }

    if (docTypeKey === 'invoice' || docTypeKey === 'delivery') {
      markBillingCompletedIfApplicable_(caseNo);
    }

    appendOperationLog_(caseNo, `${docType.label}印刷`, '', false);

    return { printUrl: getPrintPreviewUrl_(fileId, docType) };
  }, caseNo);
}

/** PDF出力: PDFを保存フォルダへ書き出し、リンクをシートへ記録する */
function exportDocumentPdfForCase_(docTypeKey, caseNo) {
  return withLock_(`${DOC_TYPES[docTypeKey].label}のPDF出力`, () => {
    const docType = DOC_TYPES[docTypeKey];
    const caseInfo = getCaseInfo_(caseNo);
    const email = getActiveUserEmail_();
    const staff = findStaffByEmail_(email);
    const now = new Date();

    const fileId = extractFileIdFromUrl_(caseInfo[`${docTypeKey}Link`]);
    const file = DriveApp.getFileById(fileId);
    const sheet = getPrimarySheet_(file, docType);
    const cells = docType.cells();

    const stage = docTypeKey === 'invoice' ? 'billed' : 'created';
    const folder = getCaseDocFolder_(docType, caseInfo, stage);
    const pdfFile = exportFileToPdf_(fileId, folder, `${docType.label}_${caseInfo.caseNo}`, docType);

    setCellValue_(sheet, cells.PDF_OUTPUT_BY, staff ? staff.name : email);
    setCellValue_(sheet, cells.PDF_OUTPUT_AT, formatDateTime_(now));

    const fieldUpdates = { [docType.col.outputLink]: pdfFile.getUrl() };
    if (docType.col.outputBy) {
      fieldUpdates[docType.col.outputBy] = staff ? staff.name : email;
      fieldUpdates[docType.col.outputAt] = formatDateTime_(now);
    }
    setCaseFields_(caseNo, fieldUpdates);

    if (docTypeKey === 'invoice' || docTypeKey === 'delivery') {
      markBillingCompletedIfApplicable_(caseNo);
    }

    appendOperationLog_(caseNo, `${docType.label}PDF出力`, `URL: ${pdfFile.getUrl()}`, false);

    return { pdfUrl: pdfFile.getUrl() };
  }, caseNo);
}

/**
 * 請求書・納品書の「印刷」または「PDF出力」のいずれかが実行された時点で、
 * 請求ステータスを「請求済み」に変更し、最終承認ボタンを有効化する
 * （サイドバー側はステータス文字列を見て活性/非活性を判断するため、ここでは
 *  ステータス更新のみ行う）。
 */
function markBillingCompletedIfApplicable_(caseNo) {
  const caseInfo = getCaseInfo_(caseNo);
  if (caseInfo.billingStatus === BILLING_STATUS.NOT_BILLED) {
    setCaseFields_(caseNo, { [CASE_COLS.BILLING_STATUS]: BILLING_STATUS.BILLED });
  }
}

/** GoogleドライブファイルのURLからファイルIDを取り出す */
function extractFileIdFromUrl_(url) {
  if (!url) {
    throw AppError_('DOC_NOT_FOUND', '対象の書類がまだ作成されていません。');
  }
  const match = String(url).match(/[-\w]{25,}/);
  if (!match) {
    throw AppError_('INVALID_URL', `書類のURLからファイルIDを取得できませんでした: ${url}`);
  }
  return match[0];
}
