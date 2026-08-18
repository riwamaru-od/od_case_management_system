/**
 * ClientService.gs
 * 取引先DB（Googleフォームの新規取引先登録フォームから自動転記される想定）の参照。
 */

function getAllClients_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('ALL_CLIENTS');
  if (cached) return JSON.parse(cached);

  const sheet = getClientDbSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, CLIENT_COLS.NOTES).getValues();
  const clients = values
    .filter(row => row[CLIENT_COLS.COMPANY_NAME - 1])
    .map(row => ({
      odStaff: row[CLIENT_COLS.OD_STAFF - 1],
      companyName: row[CLIENT_COLS.COMPANY_NAME - 1],
      contactName: row[CLIENT_COLS.CONTACT_NAME - 1],
      department: row[CLIENT_COLS.DEPARTMENT - 1],
      postalCode: row[CLIENT_COLS.POSTAL_CODE - 1],
      address1: row[CLIENT_COLS.ADDRESS1 - 1],
      address2: row[CLIENT_COLS.ADDRESS2 - 1],
      contactPhone: row[CLIENT_COLS.CONTACT_PHONE - 1],
      contactEmail: row[CLIENT_COLS.CONTACT_EMAIL - 1],
      invoiceDeliveryMethod: row[CLIENT_COLS.INVOICE_DELIVERY_METHOD - 1],
      invoiceFormatSpec: row[CLIENT_COLS.INVOICE_FORMAT_SPEC - 1],
      qualifiedInvoiceNo: row[CLIENT_COLS.QUALIFIED_INVOICE_NO - 1],
      invoiceCutoffDay: row[CLIENT_COLS.INVOICE_CUTOFF_DAY - 1],
      paymentMonth: row[CLIENT_COLS.PAYMENT_MONTH - 1],
      paymentDay: row[CLIENT_COLS.PAYMENT_DAY - 1],
      notes: row[CLIENT_COLS.NOTES - 1],
    }));

  cache.put('ALL_CLIENTS', JSON.stringify(clients), 300);
  return clients;
}

/** 取引先社名（プルダウンで選択されたものと完全一致）から取引先情報を取得 */
function getClientByName_(companyName) {
  const client = getAllClients_().find(c => c.companyName === companyName);
  if (!client) {
    throw AppError_('CLIENT_NOT_FOUND', `取引先DBに「${companyName}」が見つかりません。`);
  }
  return client;
}

/** 案件一覧のプルダウン用に、取引先社名の一覧を返す */
function listClientNames_() {
  return getAllClients_().map(c => c.companyName);
}
