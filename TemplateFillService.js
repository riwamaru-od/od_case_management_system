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

/** 指定ファイル（見積書/請求書）の現行「最新」シートから、通し番号（SERIAL_NO）の現在値を読み取る */
function readDocumentSerialNo_(fileId, docType) {
  const sheet = getPrimarySheet_(DriveApp.getFileById(fileId), docType);
  return sheet.getRange(docType.cells().SERIAL_NO).getValue();
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

  const quoteSerialNo = readDocumentSerialNo_(quoteFileId, DOC_TYPES.quote);
  setCellValue_(sheet, cells.QUOTE_SERIAL_REF, `見積書No.${quoteSerialNo}`);

  copyRanges_(quoteFileId, DOC_TYPES.quote, sheet, QUOTE_TEMPLATE_CELLS.BODY_COPY_RANGE_FOR_INVOICE);

  // 転記が終わった直後の状態（＝見積書と完全に一致している状態）を基準として、
  // 以降に請求書側だけが編集された場合に黄色くなるよう差分ハイライトを設定する
  try {
    applyQuoteInvoiceDiffHighlight_(sheet, quoteFileId);
  } catch (e) {
    // ハイライトは補助機能のため、失敗しても請求書の作成自体は成立させる
    console.warn(`見積書との差分ハイライトの設定に失敗しました: ${e}`);
  }
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

  const quoteFileId = extractFileIdFromUrl_(caseInfo.quoteLink);
  const quoteSerialNo = readDocumentSerialNo_(quoteFileId, DOC_TYPES.quote);
  const invoiceSerialNo = readDocumentSerialNo_(invoiceFileId, DOC_TYPES.invoice);
  setCellValue_(sheet, cells.QUOTE_SERIAL_REF, `見積書No.${quoteSerialNo}`);
  setCellValue_(sheet, cells.INVOICE_SERIAL_REF, `請求書No.${invoiceSerialNo}`);

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

/**
 * 社印画像を指定範囲へ挿入する（G8:G13相当。見積書・請求書は承認時、納品書は作成時に呼ぶ）。
 * サイズは Constants.js の SEAL_IMAGE_WIDTH_PX / SEAL_IMAGE_HEIGHT_PX で明示的に指定する
 * （画像素材そのものの解像度に依存させないため）。社印サイズを変更したい場合はその2つの値を編集する。
 *
 * COMPANY_SEAL_IMAGE_URL には、Googleドライブの通常の共有リンク
 * （例: https://drive.google.com/file/d/xxxxx/view?usp=sharing）をそのまま設定してよい。
 * このURLはHTMLのプレビュー画面を返すため画像として直接フェッチすることはできない。
 * そのため URL からファイルIDを取り出し、DriveApp 経由で画像データそのものを取得して挿入する
 * （スクリプトの閲覧権限があれば取得できるため、ファイルを一般公開する必要も無い）。
 */
function insertSealImage_(file, cells, docType) {
  const source = getSealImageSource_();
  if (!source) return; // 未設定（警告ログ済み）

  const sheet = getPrimarySheet_(file, docType);
  const range = sheet.getRange(cells.SEAL_IMAGE_RANGE);
  removeSealImages_(sheet, range); // 再承認時などに二重で押印されないようにする
  const image = insertSealBlobWithFallback_(sheet, source, range);
  image.setWidth(SEAL_IMAGE_WIDTH_PX);
  image.setHeight(SEAL_IMAGE_HEIGHT_PX);
  SpreadsheetApp.flush();
  return image;
}

/**
 * 社印画像をシートへ挿入する。Sheet.insertImage には「2MB以下・100万画素以下」という
 * 上限があり、高解像度の社印画像はそのままでは挿入できないため、失敗した場合は
 * Driveが生成する縮小版（サムネイル）で再挑戦する。
 * 表示サイズは SEAL_IMAGE_WIDTH_PX 角（既定70px）と小さいため、縮小版でも実用上の
 * 画質劣化は問題にならない。
 */
function insertSealBlobWithFallback_(sheet, source, range) {
  const col = range.getColumn();
  const row = range.getRow();
  try {
    return sheet.insertImage(source.blob, col, row, SEAL_IMAGE_OFFSET_X_PX, SEAL_IMAGE_OFFSET_Y_PX);
  } catch (e) {
    console.warn(`社印画像の直接挿入に失敗したため、縮小版で再試行します: ${e}`);
    const thumbnail = getSealThumbnailBlob_(source.file);
    if (!thumbnail) {
      throw AppError_('SEAL_TOO_LARGE',
        `社印画像「${source.file.getName()}」は大きすぎて挿入できません`
        + `（上限: 2MB・100万画素 / 現在: ${Math.round(source.blob.getBytes().length / 1024)}KB）。`
        + '縮小版の取得にも失敗したため、社印画像そのものを小さいPNG（例: 300x300px程度）に差し替えてください。');
    }
    return sheet.insertImage(thumbnail, col, row, SEAL_IMAGE_OFFSET_X_PX, SEAL_IMAGE_OFFSET_Y_PX);
  }
}

/**
 * 社印の基準セルに貼られている既存の画像を取り除く。
 * 再承認時の二重押印や、再作成で複製されたシートに旧版の社印が残るのを防ぐ。
 */
function removeSealImages_(sheet, range) {
  const col = range.getColumn();
  const row = range.getRow();
  sheet.getImages().forEach(img => {
    const anchor = img.getAnchorCell();
    if (anchor.getColumn() === col && anchor.getRow() === row) img.remove();
  });
}

/** insertImage の上限を超える社印画像のための、Drive生成の縮小版を取得する */
function getSealThumbnailBlob_(sealFile) {
  // Drive API のサムネイル（取得サイズを指定できるため画質を確保しやすい）
  try {
    const token = ScriptApp.getOAuthToken();
    const metaRes = UrlFetchApp.fetch(
      `https://www.googleapis.com/drive/v3/files/${sealFile.getId()}?fields=thumbnailLink`,
      { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    if (metaRes.getResponseCode() === 200) {
      const link = JSON.parse(metaRes.getContentText()).thumbnailLink;
      if (link) {
        const sizedLink = link.replace(/=s\d+(-c)?$/, '') + `=s${SEAL_THUMBNAIL_MAX_PX}`;
        const imgRes = UrlFetchApp.fetch(sizedLink, { muteHttpExceptions: true });
        if (imgRes.getResponseCode() === 200) return imgRes.getBlob();
      }
    }
  } catch (e) {
    console.warn(`Driveサムネイルの取得に失敗しました: ${e}`);
  }
  // フォールバック: DriveApp の簡易サムネイル（小さめ・サイズ指定不可）
  try {
    return sealFile.getThumbnail();
  } catch (e) {
    console.warn(`簡易サムネイルの取得にも失敗しました: ${e}`);
    return null;
  }
}

/**
 * スクリプトプロパティの設定から社印画像のファイルとBlobを取得する。
 * 未設定の場合は null を返す（＝押印をスキップ）。設定されているが取得できない場合は
 * 原因が分かるメッセージで例外を投げる（呼び出し元が操作ログへ記録する）。
 */
function getSealImageSource_() {
  const sealProp = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.COMPANY_SEAL_IMAGE_URL);
  if (!sealProp) {
    console.warn('COMPANY_SEAL_IMAGE_URL が未設定のため、社印画像の挿入をスキップしました。');
    return null;
  }

  let sealFile;
  try {
    sealFile = DriveApp.getFileById(extractFileIdFromUrl_(sealProp));
  } catch (e) {
    throw AppError_('SEAL_NOT_ACCESSIBLE',
      `社印画像ファイルを取得できません（COMPANY_SEAL_IMAGE_URL: ${sealProp}）。`
      + `URLが正しいか、実行者がそのファイルを閲覧できるか確認してください。詳細: ${e && e.message ? e.message : e}`);
  }

  const blob = sealFile.getBlob();
  const mimeType = blob.getContentType();
  if (!mimeType || mimeType.indexOf('image/') !== 0) {
    throw AppError_('SEAL_NOT_IMAGE',
      `社印に指定されたファイル「${sealFile.getName()}」は画像ではありません（種類: ${mimeType}）。`
      + 'PNGやJPEGなどの画像ファイルを指定してください（Googleスライド・図形描画などは不可）。');
  }
  return { file: sealFile, blob: blob };
}

/**
 * メニュー「社印設定を確認する」から呼ばれる診断用の関数。
 * 社印画像が正しく設定・取得できるかをその場で確認し、結果をダイアログで表示する。
 */
function checkSealImageSetting() {
  const ui = SpreadsheetApp.getUi();
  const sealProp = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.COMPANY_SEAL_IMAGE_URL);
  if (!sealProp) {
    ui.alert('社印設定の確認\n\nスクリプトプロパティ「COMPANY_SEAL_IMAGE_URL」が未設定です。\n'
      + 'ファイル > プロジェクトの設定 > スクリプト プロパティ から、社印画像のGoogleドライブURLを設定してください。');
    return;
  }
  try {
    const source = getSealImageSource_();
    const sizeKb = Math.round(source.blob.getBytes().length / 1024);

    // 実際に一時シートへ挿入してみて、上限（2MB・100万画素）に掛からないか検証する
    const tempSs = SpreadsheetApp.create(`_tmp_seal_check_${Utilities.getUuid()}`);
    let insertedVia;
    try {
      const tempSheet = tempSs.getSheets()[0];
      try {
        tempSheet.insertImage(source.blob, 1, 1);
        insertedVia = '元の画像をそのまま挿入できます。';
      } catch (e) {
        const thumbnail = getSealThumbnailBlob_(source.file);
        if (!thumbnail) throw e;
        tempSheet.insertImage(thumbnail, 1, 1);
        insertedVia = '元の画像は大きすぎるため（上限: 2MB・100万画素）、'
          + 'Driveが生成する縮小版を自動で使用します。表示サイズが小さいため画質に影響はありません。';
      }
    } finally {
      DriveApp.getFileById(tempSs.getId()).setTrashed(true);
    }

    ui.alert([
      '社印設定の確認: 正常',
      '',
      `ファイル名: ${source.file.getName()}`,
      `種類: ${source.blob.getContentType()}`,
      `ファイルサイズ: ${sizeKb} KB`,
      `貼り付けサイズ: ${SEAL_IMAGE_WIDTH_PX} x ${SEAL_IMAGE_HEIGHT_PX} px`,
      '',
      insertedVia,
      '承認（納品書は作成）時に押印されます。',
    ].join('\n'));
  } catch (e) {
    ui.alert(`社印設定の確認: エラー\n\n${e && e.message ? e.message : e}`);
  }
}

/**
 * 承認済み・差し戻し済みのシートを保護し、オーナー以外は誰も編集できない状態にする。
 * 注意: Googleスプレッドシートのシート保護は、「ファイルのオーナー」であれば
 * 保護の設定変更・解除ができてしまう仕様（APIでも回避不可）のため、これ単体では
 * オーナーに対する実効性が無い。書類を必ず管理用アカウントで作成し、その1アカウント
 * 以外は編集できない状態にすることで担保する（verifyAdminAccountExecution_ 参照）。
 */
/**
 * シート全体を保護（編集不可に）する。
 * @param {Sheet} sheet 対象シート
 * @param {string[]} [unprotectedA1Ranges] 保護の対象外にする範囲（A1形式）。
 *   例: 見積書は承認後も作業用エリア（J21:O54）だけは編集できるようにする。
 */
function protectSheet_(sheet, unprotectedA1Ranges) {
  const protection = sheet.protect().setDescription('承認済み/差し戻しにつき編集不可');
  if (unprotectedA1Ranges && unprotectedA1Ranges.length) {
    protection.setUnprotectedRanges(unprotectedA1Ranges.map(a1 => sheet.getRange(a1)));
  }
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
 * 書類ファイルの作成が、管理用アカウントによって実行されているかを検証する。
 *
 * Google Workspace を利用していない環境では、DriveApp の setOwner() によるオーナー移譲は
 * 使えない（ドメインをまたぐ移譲は「招待制」となり、相手が承諾するまで反映されないため、
 * 移譲したつもりでオーナーが変わっていない、という危険な状態になりうる）。
 * そこで本システムは「書類を作成する操作は管理用アカウントで行う」という運用を前提とし、
 * ファイルは最初から管理用アカウントの所有物として作られるようにしている。
 *
 * この関数は、その前提が守られているかを実行時に検証する。管理用アカウント以外が
 * 書類を作成した場合、そのファイルのオーナーは実行者本人になり、承認後も本人であれば
 * シート保護を解除して編集できてしまうため、操作ログにエラーとして記録して気付けるようにする。
 * （処理自体は続行する。書類作成そのものを止めてしまうと業務が滞るため）
 */
function verifyAdminAccountExecution_(caseNo, actionLabel) {
  let adminEmail;
  try {
    adminEmail = getAdminTriggerAccountEmail_();
  } catch (e) {
    console.warn('ADMIN_TRIGGER_ACCOUNT_EMAIL が未設定のため、実行アカウントの検証をスキップしました。');
    return true;
  }

  let executor = '';
  try {
    executor = Session.getActiveUser().getEmail() || '';
  } catch (e) {
    executor = '';
  }
  if (!executor) {
    console.warn('実行ユーザーのメールアドレスを取得できないため、実行アカウントの検証をスキップしました。');
    return true;
  }

  if (executor.trim().toLowerCase() === String(adminEmail).trim().toLowerCase()) return true;

  const message = `管理用アカウント（${adminEmail}）以外（${executor}）が書類を作成したため、`
    + 'このファイルのオーナーは実行者本人になります。承認後も本人であればシート保護を解除して'
    + '編集できてしまうため、書類の作成は管理用アカウントで行ってください。';
  console.warn(message);
  try {
    appendOperationLog_(caseNo || '', `${actionLabel}（実行アカウント）`, message, true);
  } catch (e) {
    console.error(`実行アカウント警告の記録に失敗しました: ${e}`);
  }
  return false;
}

/**
 * ファイルレベルの編集権限を、オーナー（管理用アカウント）以外は閲覧のみへ引き下げる。
 * protectSheet_（シート保護）は「ファイルのオーナー」の編集権限までは剥奪できない
 * （Googleスプレッドシートの仕様上、いかなるAPIを使っても回避不可能な制約）ため、
 * これは「オーナー以外の全員が直接編集できない」状態を保証するための追加の防御層。
 * 管理用アカウント自身の編集権限を制限する方法は存在しないため、運用上は
 * ADMIN_TRIGGER_ACCOUNT_EMAIL のログイン情報を日常的な編集作業に使わないことが必須になる。
 */
function restrictFileEditAccessToAdminOnly_(file) {
  const owner = file.getOwner();
  const ownerEmail = owner ? owner.getEmail() : null;
  file.getEditors().forEach(editor => {
    if (ownerEmail && editor.getEmail() === ownerEmail) return;
    try {
      file.removeEditor(editor);
      file.addViewer(editor);
    } catch (e) {
      console.warn(`ファイル編集権限の降格に失敗しました(${editor.getEmail()}): ${e}`);
    }
  });
}
