/**
 * SidebarController.gs
 * サイドバーHTML（Sidebar.html / JavaScript.html）から google.script.run 経由で
 * 呼び出される、唯一の入り口。ここに並ぶ関数だけがクライアントに公開される想定。
 *
 * 設計方針（確定仕様9章）:
 *  - サイドバーは単一HTMLで、状態に応じてボタンの活性/非活性のみ切り替える
 *  - 承認者向け／一般担当者向けで画面は分けず、ボタンの活性制御のみで出し分ける
 *  - 表示内容: 選択中の案件の概要、現在のステータス、直近の操作履歴
 *
 * 権限について:
 *  - ここで computeButtonStates_ により事前にボタンを非活性にするが、
 *    実際の権限チェックは必ず各 *Service.gs 側（assertRole_）でも行う。
 *    （フロントの活性制御はUX目的であり、権限の最終判定はサーバー側で行う）
 */

/** HTMLテンプレート内で他のHTMLファイルを読み込むための共通ヘルパー */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * サイドバー初期表示・定期更新時に呼ばれる。
 * 「現在選択中の案件」は、メイン画面UIシート上でユーザーが選択しているセルの行から
 * 案件番号列を読み取って特定する（対象シート以外を選択中の場合は caseNo=null を返す）。
 */
function api_getSidebarData(preferredCaseNo) {
  const email = getActiveUserEmail_();
  const staff = findStaffByEmail_(email);
  let selection = getSelectedCaseNo_();

  if (!selection && preferredCaseNo) {
    try {
      if (getCaseInfo_(preferredCaseNo)) {
        selection = String(preferredCaseNo);
      }
    } catch (e) {}
  }

  const base = {
    currentUser: { email, name: staff ? staff.name : email, roles: staff ? staff.roles : [] },
    caseNo: null,
    caseInfo: null,
    buttons: null,
    history: [],
  };

  if (!selection) return base;

  let caseInfo;
  try {
    caseInfo = getCaseInfo_(selection);
  } catch (e) {
    return base; // 選択行がまだ案件番号未採番など
  }

  base.caseNo = selection;
  base.caseInfo = caseInfo;
  base.buttons = computeButtonStates_(caseInfo, email);
  base.history = buildHistoryList_(caseInfo);
  return base;
}

/**
 * アクティブなUIシート、または全案件DBシート上で、現在選択されているセルの行から
 * 案件番号を取得する（全案件DBとメイン画面UIは同一列構成 - Constants.js参照のため
 * CASE_COLS.CASE_NO の列番号をそのまま使い回せる）。
 */
function getSelectedCaseNo_() {
  try {
    const ss = getMainSpreadsheet_();
    const activeUiSheet = getActiveUiSheet_();
    const activeDbSheet = getActiveDbSheet_();
    const activeSheet = ss.getActiveSheet();
    // UIシート・全案件DBシート以外が選択されている場合はnullを返す
    const isUiOrDbSheet = !!activeSheet && (
      activeSheet.getSheetId() === activeUiSheet.getSheetId()
      || activeSheet.getSheetId() === activeDbSheet.getSheetId()
    );
    if (!isUiOrDbSheet) return null;

    const row = ss.getActiveRange() && ss.getActiveRange().getRow();
    if (!row || row < CASE_DATA_START_ROW) return null;

    const caseNo = activeSheet.getRange(row, CASE_COLS.CASE_NO).getValue();
    return caseNo ? String(caseNo).trim() : null;
  } catch (e) {
    return null;
  }
}

function isQuoteDraftedStatus_(status) {
  return status === STATUS.QUOTE_DRAFTED || status === '見積書作成済み' || status === '見積書承認待ち';
}

function isInvoiceDraftedStatus_(status) {
  return status === STATUS.INVOICE_DRAFTED || status === '請求書作成済み' || status === '請求書承認待ち';
}

/**
 * ステータス・請求ステータス・ユーザーのロールから、各ボタンの活性/非活性を判定する。
 * 各値は { enabled, reason } の形。reason は非活性時にツールチップとして表示する。
 */
function computeButtonStates_(caseInfo, email) {
  const s = caseInfo.status;
  const bs = caseInfo.billingStatus;
  const canApproveQuote = isQuoteApprover_(email);
  const canApproveInvoice = isInvoiceApprover_(email);
  const canFinalApprove = isFinalApprover_(email);

  const isQuoteDrafted = isQuoteDraftedStatus_(s);
  const isInvoiceDrafted = isInvoiceDraftedStatus_(s);

  const b = (enabled, reason) => ({ enabled: !!enabled, reason: enabled ? '' : (reason || '') });

  return {
    createQuote: b(!caseInfo.quoteLink, '既に見積書が作成されています'),
    completeQuote: b(s === STATUS.QUOTE_IN_PROGRESS && !!caseInfo.quoteLink, '見積書作成中の案件のみ操作できます'),
    approveQuote: b(isQuoteDrafted && canApproveQuote, !isQuoteDrafted ? '承認待ちの状態ではありません' : '見積書承認の権限がありません'),
    rejectQuote: b(isQuoteDrafted && canApproveQuote, !isQuoteDrafted ? '承認待ちの状態ではありません' : '見積書承認の権限がありません'),
    // 再作成は「差し戻し後」または「承認後」のみ活性化する（承認待ち＝作成直後の作成中と区別するため
    // quoteRejectedAt フラグを利用する。Constants.js/CASE_COLS.QUOTE_REJECTED_AT 参照）
    recreateQuote: b((!!caseInfo.quoteRejectedAt || s === STATUS.QUOTE_APPROVED) && canApproveQuote,
      (!caseInfo.quoteRejectedAt && s !== STATUS.QUOTE_APPROVED) ? '差し戻し後または承認後のみ操作できます' : '見積書承認の権限がありません'),

    printQuote: b(s === STATUS.QUOTE_APPROVED || !!caseInfo.quoteApprovedAt, '見積書承認済み以降のみ操作できます'),
    exportQuotePdf: b(s === STATUS.QUOTE_APPROVED || !!caseInfo.quoteApprovedAt, '見積書承認済み以降のみ操作できます'),
    createInvoice: b((s === STATUS.QUOTE_APPROVED || !!caseInfo.quoteApprovedAt) && !caseInfo.invoiceLink, !caseInfo.quoteApprovedAt ? '見積書承認済み以降のみ操作できます' : '既に請求書が作成されています'),

    completeInvoice: b(s === STATUS.INVOICE_IN_PROGRESS && !!caseInfo.invoiceLink, '請求書作成中の案件のみ操作できます'),
    approveInvoice: b(isInvoiceDrafted && canApproveInvoice, !isInvoiceDrafted ? '承認待ちの状態ではありません' : '請求書承認の権限がありません'),
    rejectInvoice: b(isInvoiceDrafted && canApproveInvoice, !isInvoiceDrafted ? '承認待ちの状態ではありません' : '請求書承認の権限がありません'),
    recreateInvoice: b((!!caseInfo.invoiceRejectedAt || s === STATUS.INVOICE_APPROVED) && canApproveInvoice,
      (!caseInfo.invoiceRejectedAt && s !== STATUS.INVOICE_APPROVED) ? '差し戻し後または承認後のみ操作できます' : '請求書承認の権限がありません'),

    printInvoice: b(!!caseInfo.invoiceApprovedAt, '請求書承認済み以降のみ操作できます'),
    exportInvoicePdf: b(!!caseInfo.invoiceApprovedAt, '請求書承認済み以降のみ操作できます'),
    createDelivery: b(!!caseInfo.invoiceApprovedAt && !caseInfo.deliveryLink, !caseInfo.invoiceApprovedAt ? '請求書承認済み以降のみ操作できます' : '既に納品書が作成されています'),

    printDelivery: b(!!caseInfo.deliveryLink, '納品書がまだ作成されていません'),
    exportDeliveryPdf: b(!!caseInfo.deliveryLink, '納品書がまだ作成されていません'),

    finalApprove: b(bs === BILLING_STATUS.BILLED && canFinalApprove, bs !== BILLING_STATUS.BILLED ? '請求書・納品書の印刷／PDF出力が完了していません' : '最終承認の権限がありません'),
  };
}

/** 記録済みの作成者・承認者・出力者情報を時系列に並べた「直近の操作履歴」を組み立てる */
function buildHistoryList_(caseInfo) {
  const entries = [
    ['見積書 作成', caseInfo.quoteCreator, caseInfo.quoteCreatedAt],
    ['見積書 承認', caseInfo.quoteApprover, caseInfo.quoteApprovedAt],
    ['見積書 出力', caseInfo.quoteOutputBy, caseInfo.quoteOutputAt],
    ['請求書 作成', caseInfo.invoiceCreator, caseInfo.invoiceCreatedAt],
    ['請求書 承認', caseInfo.invoiceApprover, caseInfo.invoiceApprovedAt],
    ['請求書 出力', caseInfo.invoiceOutputBy, caseInfo.invoiceOutputAt],
    ['納品書 作成', caseInfo.deliveryCreator, caseInfo.deliveryCreatedAt],
    ['最終承認', caseInfo.finalApprover, caseInfo.finalApprovedAt],
  ];
  return entries
    .filter(([, who, when]) => who && when)
    .map(([label, who, when]) => ({
    label: String(label),
    who: String(who),
    when: when instanceof Date ? formatDateTime_(when) : String(when),
    }))
    .reverse(); // 記録順（≒時系列）の末尾＝直近を先頭に
}

// ------------------------------------------------------------------
// アクション用の薄いラッパー。例外は AppError_ のまま投げ、
// クライアント側では withFailureHandler(error => alert(error.message)) で受ける。
// ------------------------------------------------------------------

function api_createQuote(caseNo) { return createDocumentForCase_('quote', caseNo); }
function api_completeQuote(caseNo, comment) { return completeDocumentForCase_('quote', caseNo, comment); }
function api_approveQuote(caseNo, comment) { return approveDocumentForCase_('quote', caseNo, comment); }
function api_rejectQuote(caseNo, comment) { return rejectDocumentForCase_('quote', caseNo, comment); }
function api_recreateQuote(caseNo) { return recreateDocumentForCase_('quote', caseNo); }
function api_printQuote(caseNo) { return printDocumentForCase_('quote', caseNo); }
function api_exportQuotePdf(caseNo) { return exportDocumentPdfForCase_('quote', caseNo); }

function api_createInvoice(caseNo) { return createDocumentForCase_('invoice', caseNo); }
function api_completeInvoice(caseNo, comment) { return completeDocumentForCase_('invoice', caseNo, comment); }
function api_approveInvoice(caseNo, comment) { return approveDocumentForCase_('invoice', caseNo, comment); }
function api_rejectInvoice(caseNo, comment) { return rejectDocumentForCase_('invoice', caseNo, comment); }
function api_recreateInvoice(caseNo) { return recreateDocumentForCase_('invoice', caseNo); }
function api_printInvoice(caseNo) { return printDocumentForCase_('invoice', caseNo); }
function api_exportInvoicePdf(caseNo) { return exportDocumentPdfForCase_('invoice', caseNo); }

function api_createDelivery(caseNo) { return createDeliveryForCase_(caseNo); }
function api_printDelivery(caseNo) { return printDocumentForCase_('delivery', caseNo); }
function api_exportDeliveryPdf(caseNo) { return exportDocumentPdfForCase_('delivery', caseNo); }

function api_finalApprove(caseNo, comment) { return finalApproveForCase_(caseNo, comment); }
