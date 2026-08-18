/**
 * DeliveryService.gs
 * 納品書は承認フローを持たない（確定仕様どおり）。
 * 請求書をベースに作成し、作成者/作成日時・社印のみ記録する。PDF出力は
 * ApprovalService.gs の共通関数（exportDocumentPdfForCase_）を
 * docTypeKey='delivery' で呼び出せば良いように設計している。
 */

function createDeliveryForCase_(caseNo) {
  return withLock_('納品書の作成', () => {
    const docType = DOC_TYPES.delivery;
    const caseInfo = getCaseInfo_(caseNo);
    const email = getActiveUserEmail_();
    const staff = findStaffByEmail_(email);
    const now = new Date();

    const invoiceFileId = extractFileIdFromUrl_(caseInfo.invoiceLink);
    const file = createLatestDocument_(docType, caseInfo, 'created');
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

    setCaseFields_(caseNo, {
      [docType.col.link]: file.getUrl(),
      [docType.col.creator]: staff ? staff.name : email,
      [docType.col.createdAt]: formatDateTime_(now),
    });

    appendOperationLog_(caseNo, '納品書作成', `URL: ${file.getUrl()}`, false);

    return { url: file.getUrl() };
  }, caseNo);
}
