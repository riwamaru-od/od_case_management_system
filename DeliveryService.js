/**
 * DeliveryService.gs
 * 納品書は承認フローを持たない（確定仕様どおり）。
 * 請求書をベースに作成し、作成者/作成日時のみ記録する。印刷・PDF出力は
 * ApprovalService.gs の共通関数（printDocumentForCase_ / exportDocumentPdfForCase_）を
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

    const sheet = getPrimarySheet_(file);
    const cells = docType.cells();
    setCellValue_(sheet, cells.CREATOR_NAME, staff ? staff.name : email);
    setCellValue_(sheet, cells.CREATED_AT, formatDateTime_(now));

    setCaseFields_(caseNo, {
      [docType.col.link]: file.getUrl(),
      [docType.col.creator]: staff ? staff.name : email,
      [docType.col.createdAt]: formatDateTime_(now),
    });

    return { url: file.getUrl() };
  });
}
