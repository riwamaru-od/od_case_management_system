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

    // 二重作成の防止: 既にこの書類のファイルがある場合は作り直さず、既存のURLを返す。
    // （同じ作成リクエストが稀に二重で届き、同じ案件のファイルが2つできる事象への対策。
    //  getCaseInfo_ も含めて withLock_ の中で判定しているため、ほぼ同時に2件届いた場合も
    //  1件目の書き込み後に2件目がここで弾かれる）
    if (caseInfo[`${docTypeKey}Link`]) {
      appendOperationLog_(caseNo, `${docType.label}作成`, '既に作成済みのため作成をスキップしました（二重実行の防止）', false);
      return { url: caseInfo[`${docTypeKey}Link`] };
    }

    // 差し戻し・再作成されて再承認待ちの見積書から、請求書を作らせない
    // （サイドバー側でもボタンを非活性にしているが、権限・整合性の判定はサーバー側でも行う）
    if (docTypeKey === 'invoice' && caseInfo.quoteReapprovalPending) {
      throw AppError_('INVALID_STATE', '見積書が再作成・差し戻し中です。見積書が再度承認されてから請求書を作成してください。');
    }

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
    if (docType.col.reapprovalPending) fieldUpdates[docType.col.reapprovalPending] = '';
    if (docType.col.startedAt) fieldUpdates[docType.col.startedAt] = formatDateTime_(new Date());
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
function completeDocumentForCase_(docTypeKey, caseNo, comment, approverEmail) {
  return withLock_(`${DOC_TYPES[docTypeKey].label}の完成`, () => {
    const docType = DOC_TYPES[docTypeKey];
    const caseInfo = getCaseInfo_(caseNo);

    // 承認者の指定は任意。指定された場合のみ、その人が承認権限を持つか検証する
    const designatedApprover = approverEmail ? findStaffByEmail_(approverEmail) : null;
    if (approverEmail && !designatedApprover) {
      throw AppError_('APPROVER_NOT_FOUND', `指定された承認者（${approverEmail}）が社員DBに見つかりません。`);
    }
    if (designatedApprover && !hasAnyRole_(designatedApprover.email, docType.approverRoles)) {
      throw AppError_('APPROVER_NO_ROLE', `${designatedApprover.name}さんは${docType.label}の承認権限を持っていません。`);
    }

    // 差し戻し後、再作成せずに同じ書類のまま承認依頼を再送することを防ぐ
    // （差し戻された旧シートは protectSheet_ でロック済みのため、再作成せずに
    //  書き込もうとするとどのみち失敗するが、ここで明示的に弾いてユーザーへ案内する）
    if (docType.col.rejectedAt && caseInfo[`${docTypeKey}RejectedAt`]) {
      throw AppError_('INVALID_STATE', `差し戻し後は「再作成」を行ってから${docType.label}の完成（承認依頼）を行ってください。`);
    }

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
    notifyApprovalRequest_(msg, designatedApprover ? designatedApprover.email : '');

    const logDetail = [comment || '', designatedApprover ? `承認者指定: ${designatedApprover.name}` : '']
      .filter(Boolean).join(' / ');
    appendOperationLog_(caseNo, `${docType.label}完成（承認依頼）`, logDetail, false);

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

    // 承認されたシートを保護（編集不可にする）。
    // 見積書のみ、承認後も編集を続ける必要がある作業用エリアは保護対象から除く。
    try {
      protectSheet_(sheet, docType.unprotectedRangesAfterApproval);
    } catch (e) {
      console.warn(`シート保護(protectSheet_)の適用に失敗しました: ${e}`);
    }
    try {
      insertSealImage_(file, cells, docType);
    } catch (e) {
      // 承認処理自体は成立させるが、押印漏れに気付けるよう操作ログにエラーとして残す
      console.warn(`社印画像の挿入に失敗しました: ${e}`);
      appendOperationLog_(caseNo, `${docType.label}承認（社印）`, `社印画像の挿入に失敗: ${e && e.message ? e.message : e}`, true);
    }


    const approverName = staff ? staff.name : email;
    const approvalUpdates = {
      [docType.col.approver]: approverName,
      [docType.col.approvedAt]: formatDateTime_(now),
      [CASE_COLS.STATUS]: docType.status.approved,
    };
    // 再承認が完了したのでフラグを解除する（＝PDF出力ボタンが再び活性になる）
    if (docType.col.reapprovalPending) approvalUpdates[docType.col.reapprovalPending] = '';
    setCaseFields_(caseNo, approvalUpdates);

    if (docTypeKey === 'invoice') {
      // 請求書は承認済み以降、フォルダを「未請求案件」→「請求中案件」へ移動
      moveCaseDocFolderToStage_(docType, caseInfo, 'created', 'billed');
    }

    // このファイルへの書き込みが全て完了した後で、ファイル編集権限を降格する
    // （降格を先に行うと、後続の書き込み・フォルダ移動が実行者自身の権限不足で失敗しうるため）
    try {
      restrictFileEditAccessToAdminOnly_(file);
    } catch (e) {
      console.warn(`ファイル編集権限の降格に失敗しました: ${e}`);
    }

    // 作成者へ通知
    const creatorStaff = getAllStaff_().find(s => s.name === caseInfo[`${docTypeKey}Creator`]);
    const notifyMsg = buildApprovedMessage_(docType.label, {
      caseNo: caseInfo.caseNo, caseName: caseInfo.caseName, docUrl: file.getUrl(),
      approverName: approverName,
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
    try {
      restrictFileEditAccessToAdminOnly_(file);
    } catch (e) {
      console.warn(`ファイル編集権限の降格に失敗しました: ${e}`);
    }

    const fieldUpdates = { [CASE_COLS.STATUS]: docType.status.inProgress };
    if (docType.col.rejectedAt) fieldUpdates[docType.col.rejectedAt] = formatDateTime_(new Date());
    // 差し戻された版は承認されていないため、再承認されるまでPDF出力を禁止する
    // （過去に一度承認された書類を再作成→差し戻し、という流れでも確実に止める）
    if (docType.col.reapprovalPending) fieldUpdates[docType.col.reapprovalPending] = formatDateTime_(new Date());
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

    // 二重実行の防止: 同じ再作成リクエストが二重で届くと、1ファイル内に同じ版の
    // シートが2枚積み上がってしまう。着手日時（分単位）が今と同じであれば、
    // 直前の実行と同一の要求とみなして何もしない。
    // getCaseInfo_ も含めて withLock_ の中で判定しているため、ほぼ同時に届いた場合も
    // 1件目の書き込み後に2件目がここで弾かれる。
    if (docType.col.startedAt && caseInfo[`${docTypeKey}StartedAt`] === formatDateTime_(new Date())) {
      appendOperationLog_(caseNo, `${docType.label}再作成`, '直前の再作成と同一の要求のためスキップしました（二重実行の防止）', false);
      return { url: caseInfo[`${docTypeKey}Link`], status: docType.status.inProgress };
    }

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
    // 再作成した版はまだ承認されていないため、再承認されるまでPDF出力を禁止する
    if (docType.col.reapprovalPending) fieldUpdates[docType.col.reapprovalPending] = formatDateTime_(new Date());
    if (docType.col.startedAt) fieldUpdates[docType.col.startedAt] = formatDateTime_(new Date());
    setCaseFields_(caseNo, fieldUpdates);

    appendOperationLog_(caseNo, `${docType.label}再作成`, `URL: ${file.getUrl()}`, false);

    return { url: file.getUrl(), status: docType.status.inProgress };
  }, caseNo);
}

/**
 * PDF出力: PDFを保存フォルダへ書き出し、リンクをシートへ記録する。
 * 出力操作はこの「PDF出力」に統一されている（旧「印刷」機能は廃止）。
 * 挙動: ドライブへPDFを保存 → 呼び出し元（クライアント側）が返却されたURLを新規タブで開く。
 *
 * 順序に注意: 出力者・出力日時をシートへ書き込んでから（flushで確定させてから）
 * PDFを書き出す。逆順にすると、出力したPDFの操作履歴欄が空のままになってしまう。
 */
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

    // 先に出力ログをシートへ書き込み、確定させてからPDF化する
    setCellValue_(sheet, cells.PDF_OUTPUT_BY, staff ? staff.name : email);
    setCellValue_(sheet, cells.PDF_OUTPUT_AT, formatDateTime_(now));
    SpreadsheetApp.flush();

    const stage = docTypeKey === 'invoice' ? 'billed' : 'created';
    const folder = getCaseDocFolder_(docType, caseInfo, stage);
    const pdfFile = exportFileToPdf_(fileId, folder, `${docType.label}_${caseInfo.caseNo}`, docType);

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
 * 請求書・納品書の「PDF出力」が実行された時点で、
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
