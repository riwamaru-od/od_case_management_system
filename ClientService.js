/**
 * ClientService.gs
 * 取引先DB（Googleフォームの新規取引先登録フォームから自動転記される想定）の参照。
 */

const CLIENT_CACHE_KEY = 'ALL_CLIENTS';

function getAllClients_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CLIENT_CACHE_KEY);
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

  cache.put(CLIENT_CACHE_KEY, JSON.stringify(clients), 300);
  return clients;
}

/**
 * 取引先DBのキャッシュを破棄する。
 * 取引先が追加された直後にプルダウンの選択肢を作り直す場合など、
 * キャッシュの有効期間（300秒）を待たずに最新を読み直したいときに使う。
 */
function clearClientCache_() {
  try {
    CacheService.getScriptCache().remove(CLIENT_CACHE_KEY);
  } catch (e) {
    console.warn(`取引先DBのキャッシュ破棄に失敗しました: ${e}`);
  }
}

/**
 * 案件シートのB列（取引先名）に表示する名前を組み立てる。
 * 取引先DBには同じ社名で担当者だけが異なる行が複数あるため、社名だけでは行を特定できない。
 * そこで「社名（担当者名）」を表示名として扱う（担当者名が空欄の取引先は社名のみ）。
 */
function buildClientOptionLabel_(client) {
  const companyName = String(client.companyName == null ? '' : client.companyName).trim();
  const contactName = String(client.contactName == null ? '' : client.contactName).trim();
  if (!contactName) return companyName;
  return `${companyName}${CLIENT_OPTION_LABEL_OPEN}${contactName}${CLIENT_OPTION_LABEL_CLOSE}`;
}

/**
 * 案件シートのB列（取引先名）のプルダウン用に、表示名の一覧を返す。
 * 社名・担当者名までまったく同じ行が重複登録されていた場合は1件にまとめる
 * （データ入力規則の選択肢に同じ文字列が並ぶのを避けるため）。
 */
function listClientOptionLabels_() {
  const seen = {};
  return getAllClients_()
    .map(buildClientOptionLabel_)
    .filter(label => {
      if (!label || seen[label]) return false;
      seen[label] = true;
      return true;
    });
}

/**
 * 案件シートのB列に入っている取引先名から、取引先DBの行を取得する。
 *
 * 通常は「社名（担当者名）」の表示名（buildClientOptionLabel_）と完全一致で引く。
 * プルダウンを導入する前に登録された案件は社名だけが入っているため、その場合は
 * 社名で引き当てる。ただし社名が重複していてどの担当者か決められないときは、
 * 誤った取引先の情報を書類へ転記しないよう、選び直しを促すエラーにする。
 */
function getClientByName_(clientName) {
  const name = String(clientName == null ? '' : clientName).trim();
  if (!name) {
    throw AppError_('CLIENT_NOT_FOUND', '案件シートの取引先名が未入力です。');
  }

  const clients = getAllClients_();
  const byLabel = clients.find(c => buildClientOptionLabel_(c) === name);
  if (byLabel) return byLabel;

  const byCompanyName = clients.filter(c => String(c.companyName == null ? '' : c.companyName).trim() === name);
  if (byCompanyName.length === 1) return byCompanyName[0];
  if (byCompanyName.length > 1) {
    throw AppError_('CLIENT_AMBIGUOUS',
      `取引先DBに「${name}」が担当者違いで${byCompanyName.length}件登録されているため、取引先を特定できません。`
      + '案件シートの取引先名を、プルダウンから「社名（担当者名）」の形式で選び直してください。');
  }
  throw AppError_('CLIENT_NOT_FOUND', `取引先DBに「${name}」が見つかりません。`);
}
