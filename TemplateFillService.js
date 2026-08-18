/**
 * TemplateFillService.gs
 * 生成した見積書／請求書／納品書ファイルへ、取引先情報・案件情報・本文を書き込む。
 * セル番地は Constants.gs の *_TEMPLATE_CELLS を参照する（変更時はそちらだけ直せばよい）。
 *
 * 前提: 再作成（recreate）のたびに同一ファイル内へシートが複製されていくため、
 * 「現行の最新版シート」は 1 枚目固定ではなく、シート名（`{書類種別}_最新` = LATEST_SUFFIX）で
 * 判定する（DocumentService.gs 参照）。docType 未指定時のみ、後方互換として1枚目を返す。
 */

function getPrimarySheet_(file, docType) {
  const ss = SpreadsheetApp.openById(file.getId());
  if (docType) {
    const latest = ss.getSheetByName(`${docType.label}${LATEST_SUFFIX}`);
    if (latest) return latest;
  }
  return ss.getSheets()[0];
}

function setCellValue_(sheet, a1, value) {
  sheet.getRange(a1).setValue(value);
}

/** 見積書・請求書・納品書に共通する宛先ヘッダー情報（取引先DB由来）を書き込む */
function fillCommonHeaderCells_(sheet, cells, client, caseInfo) {
  setCellValue_(sheet, cells.POSTAL_CODE, client.postalCode);
  setCellValue_(sheet, cells.ADDRESS1, client.address1);
  setCellValue_(sheet, cells.ADDRESS2, client.address2);
  setCellValue_(sheet, cells.COMPANY_NAME, client.companyName);
  setCellValue_(sheet, cells.DEPARTMENT, client.department);
  setCellValue_(sheet, cells.CONTACT_NAME, `${client.contactName}様`);
  setCellValue_(sheet, cells.CASE_NO, caseInfo.caseNo);
  setCellValue_(sheet, cells.SUBJECT, caseInfo.caseName);
}

/** 承認担当社員の氏名・メールアドレスを書き込む（G12/G13相当） */
function fillStaffCells_(sheet, cells, staff) {
  setCellValue_(sheet, cells.STAFF_NAME, staff.name);
  setCellValue_(sheet, cells.STAFF_EMAIL, staff.email);
}

/** この期・この書類種別の通し番号を発行してG6相当に書き込む（再作成時も新規採番） */
function fillSerialNumber_(sheet, cells, docTypeKey) {
  const period = getCurrentPeriodNumber_();
  const serial = nextSequence_(seqKey_(`${docTypeKey.toUpperCase()}_SERIAL`, period));
  setCellValue_(sheet, cells.SERIAL_NO, serial);
}

/** 見積書ファイルへヘッダー一式を書き込む */
function fillQuoteDocument_(file, caseInfo) {
  const sheet = getPrimarySheet_(file, DOC_TYPES.quote);
  const cells = QUOTE_TEMPLATE_CELLS;
  const client = getClientByName_(caseInfo.clientName);
  const staff = findStaffByName_(caseInfo.staffInCharge);

  fillCommonHeaderCells_(sheet, cells, client, caseInfo);
  fillStaffCells_(sheet, cells, staff);
  fillSerialNumber_(sheet, cells, 'quote');

  setCellValue_(sheet, cells.INVOICE_CUTOFF_DAY, client.invoiceCutoffDay);
  setCellValue_(sheet, cells.PAYMENT_MONTH_DAY_TEXT, `${client.paymentMonth}${client.paymentDay}`);
  setCellValue_(sheet, cells.INVOICE_DELIVERY_METHOD, client.invoiceDeliveryMethod);
  setCellValue_(sheet, cells.INVOICE_FORMAT_SPEC, client.invoiceFormatSpec);
  setCellValue_(sheet, cells.CLIENT_CONTACT_EMAIL, client.contactEmail);
  setCellValue_(sheet, cells.CLIENT_CONTACT_PHONE, client.contactPhone);
}

/** 請求書ファイルへヘッダー一式を書き込み、見積書から本文範囲を転記する */
function fillInvoiceDocument_(file, caseInfo, quoteFileId) {
  const sheet = getPrimarySheet_(file, DOC_TYPES.invoice);
  const cells = INVOICE_TEMPLATE_CELLS;
  const client = getClientByName_(caseInfo.clientName);
  const staff = findStaffByName_(caseInfo.staffInCharge);

  fillCommonHeaderCells_(sheet, cells, client, caseInfo);
  fillStaffCells_(sheet, cells, staff);
  fillSerialNumber_(sheet, cells, 'invoice');

  copyRanges_(quoteFileId, DOC_TYPES.quote, sheet, QUOTE_TEMPLATE_CELLS.BODY_COPY_RANGE_FOR_INVOICE);
}

/** 納品書ファイルへヘッダー一式を書き込み、請求書から本文範囲を転記する */
function fillDeliveryDocument_(file, caseInfo, invoiceFileId) {
  const sheet = getPrimarySheet_(file, DOC_TYPES.delivery);
  const cells = DELIVERY_TEMPLATE_CELLS;
  const client = getClientByName_(caseInfo.clientName);
  const staff = findStaffByName_(caseInfo.staffInCharge);

  fillCommonHeaderCells_(sheet, cells, client, caseInfo);
  fillStaffCells_(sheet, cells, staff);
  fillSerialNumber_(sheet, cells, 'delivery');

  copyRanges_(invoiceFileId, DOC_TYPES.invoice, sheet, INVOICE_TEMPLATE_CELLS.BODY_COPY_RANGE_FOR_DELIVERY);
}

/**
 * ソースファイルの指定範囲（複数可）を、書式・値とも含めて対象シートの同一番地へ転記する。
 * （Range.copyTo は別スプレッドシート間で使用できないため、値・書式を個別にコピーする）
 * sourceDocType はソース側で「現行の最新版シート」を特定するために使う（getPrimarySheet_参照）。
 * 例: copyRanges_(quoteFileId, DOC_TYPES.quote, targetSheet, ['B22:E47','E49','F50','F53'])
 */
function copyRanges_(sourceFileId, sourceDocType, targetSheet, rangeList) {
  const sourceSheet = getPrimarySheet_(DriveApp.getFileById(sourceFileId), sourceDocType);
  rangeList.forEach(a1 => {
    const sourceRange = sourceSheet.getRange(a1);
    const targetRange = targetSheet.getRange(a1);
    copyRangeAcrossSpreadsheets_(sourceRange, targetRange);
  });
  SpreadsheetApp.flush();
}

function copyRangeAcrossSpreadsheets_(sourceRange, targetRange) {
  const values = sourceRange.getValues();
  const formulas = sourceRange.getFormulas();
  const combined = values.map((row, rIdx) =>
    row.map((val, cIdx) => {
      const f = formulas[rIdx][cIdx];
      return (f && String(f).indexOf('=') === 0) ? f : val;
    })
  );
  targetRange.setValues(combined);
  targetRange.setBackgrounds(sourceRange.getBackgrounds());
  targetRange.setFontColors(sourceRange.getFontColors());
  targetRange.setFontFamilies(sourceRange.getFontFamilies());
  targetRange.setFontSizes(sourceRange.getFontSizes());
  targetRange.setFontStyles(sourceRange.getFontStyles());
  targetRange.setFontWeights(sourceRange.getFontWeights());
  targetRange.setHorizontalAlignments(sourceRange.getHorizontalAlignments());
  targetRange.setVerticalAlignments(sourceRange.getVerticalAlignments());
  targetRange.setNumberFormats(sourceRange.getNumberFormats());
  targetRange.setWrapStrategies(sourceRange.getWrapStrategies());
}

/** 承認後、社印画像を指定範囲へ挿入する（G8:G13相当） */
function insertSealImage_(file, cells, docType) {
  const sealUrl = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.COMPANY_SEAL_IMAGE_URL);
  if (!sealUrl) {
    console.warn('COMPANY_SEAL_IMAGE_URL が未設定のため、社印画像の挿入をスキップしました。');
    return;
  }
  const sheet = getPrimarySheet_(file, docType);
  const range = sheet.getRange(cells.SEAL_IMAGE_RANGE);
  sheet.insertImage(sealUrl, range.getColumn(), range.getRow());
}

/**
 * 承認済み・差し戻し済みのシートを保護し、自分を含め誰も編集できない状態にする（シート保護機能）。
 * 注意: Googleスプレッドシートのシート保護は、原則「ファイルのオーナー」であれば
 * 保護の設定変更・解除ができてしまう仕様のため、これだけでは実効性が無い。
 * ファイルオーナー自体を管理用アカウントへ統一する transferFileOwnerToAdminAccount_ と
 * 組み合わせて初めて「本当に誰も編集できない」状態になる。
 */
function protectSheet_(sheet) {
  const protection = sheet.protect().setDescription('承認済み/差し戻しにつき編集不可');
  const editors = protection.getEditors();
  editors.forEach(editor => {
    try {
      protection.removeEditor(editor);
    } catch (e) {
      // オーナー自身は編集者リストから外せない仕様のため失敗しうる（想定内）
      console.warn(`保護編集者(${editor.getEmail()})の削除に失敗しました: ${e}`);
    }
  });
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}

/**
 * 書類ファイルのオーナーを管理用アカウント（PROP_KEYS.ADMIN_TRIGGER_ACCOUNT_EMAIL）へ移譲する。
 * protectSheet_ の実効性を担保するため、書類ファイル生成時に実行者個人ではなく
 * 管理用アカウントへオーナーを統一しておく。
 * 同一 Google Workspace ドメイン内であれば即時反映されるが、個人Gmail等ドメインをまたぐ場合は
 * 招待制になり、相手側の承諾操作が完了するまで反映されない点に注意。
 */
function transferFileOwnerToAdminAccount_(file) {
  const adminEmail = getAdminTriggerAccountEmail_();
  file.setOwner(adminEmail);
}
