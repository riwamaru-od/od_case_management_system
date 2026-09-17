/**
 * DeliveryService.gs
 * 納品書は承認フローを持たない（確定仕様どおり）。
 * 請求書をベースに作成し、作成者/作成日時・社印のみ記録する。PDF出力は
 * ApprovalService.gs の共通関数（exportDocumentPdfForCase_）を
 * docTypeKey='delivery' で呼び出せば良いように設計している。
 *
 * 見積書・請求書が作り直されると、今ある納品書は内容が古くなるため
 * 「無効化」される（invalidateDownstreamDocuments_）。納品書には承認フローが無く
 * 「再作成」ボタンも持たないため、無効化された場合は同じ「納品書作成」の操作で
 * 作り直す（見積書・請求書の再作成と同じく、同一ファイル内に新しいシートを積む）。
 */

function createDeliveryForCase_(caseNo) {
  return withLock_('納品書の作成', () => {
    const docType = DOC_TYPES.delivery;
    const caseInfo = getCaseInfo_(caseNo);
    const email = getActiveUserEmail_();
    const staff = findStaffByEmail_(email);
    const now = new Date();

    // 既に納品書がある場合、無効化されていれば作り直し、そうでなければ何もしない
    // （二重作成の防止。createDocumentForCase_ と同じ考え方 - ApprovalService.gs のコメント参照）
    const isRecreate = !!caseInfo.deliveryLink && !!caseInfo.deliveryInvalidatedAt;
    if (caseInfo.deliveryLink && !isRecreate) {
      appendOperationLog_(caseNo, '納品書作成', '既に作成済みのため作成をスキップしました（二重実行の防止）', false);
      return { url: caseInfo.deliveryLink };
    }

    // 差し戻し・再作成されて再承認待ちの請求書から、納品書を作らせない
    if (caseInfo.invoiceReapprovalPending) {
      throw AppError_('INVALID_STATE', '請求書が再作成・差し戻し中です。請求書が再度承認されてから納品書を作成してください。');
    }

    const invoiceFileId = extractFileIdFromUrl_(caseInfo.invoiceLink);

    // 作り直しの場合は、古い納品書のPDFを削除してから旧版シートを退避する
    const trashedPdfCount = isRecreate ? trashCaseDocPdfs_(docType, caseInfo) : 0;
    const file = isRecreate
      ? recreateLatestDocument_(docType, caseInfo)
      : createLatestDocument_(docType, caseInfo, 'created');
    fillDeliveryDocument_(file, caseInfo, invoiceFileId);

    const sheet = getPrimarySheet_(file, docType);
    const cells = docType.cells();
    setCellValue_(sheet, cells.CREATOR_NAME, staff ? staff.name : email);
    setCellValue_(sheet, cells.CREATED_AT, formatDateTime_(now));

    // 納品書には承認フローが無いため、作成時に社印を押す（見積書・請求書は承認時）
    try {
      insertSealImage_(file, cells, docType);
    } catch (e) {
      // 作成処理自体は成立させるが、押印漏れに気付けるよう操作ログにエラーとして残す
      console.warn(`社印画像の挿入に失敗しました: ${e}`);
      appendOperationLog_(caseNo, '納品書作成（社印）', `社印画像の挿入に失敗: ${e && e.message ? e.message : e}`, true);
    }

    const fieldUpdates = {
      [docType.col.link]: file.getUrl(),
      [docType.col.creator]: staff ? staff.name : email,
      [docType.col.createdAt]: formatDateTime_(now),
      // 新しい版ができたので無効化を解除する（作り直しでなければ元から空欄）
      [docType.col.invalidatedAt]: '',
    };
    // 削除した古いPDFへのリンクが残らないようにする（出力者・出力日時は履歴として残す）
    if (isRecreate) fieldUpdates[docType.col.outputLink] = '';
    setCaseFields_(caseNo, fieldUpdates);

    const logDetail = [`URL: ${file.getUrl()}`, trashedPdfCount ? `古いPDF${trashedPdfCount}件を削除` : '']
      .filter(Boolean).join(' / ');
    appendOperationLog_(caseNo, isRecreate ? '納品書再作成' : '納品書作成', logDetail, false);

    return { url: file.getUrl() };
  }, caseNo);
}
